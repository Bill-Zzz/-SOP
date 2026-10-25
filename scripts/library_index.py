#!/usr/bin/env python3
"""给本地研报库建索引（增量），供 library_search.py 检索。

研报库可以按机构、日期或公司分文件夹，也可以混在一起；本脚本会从路径、文件名和正文前几页识别：
机构、报告日期、标题、股票代码、评级、目标价、语言、页数、是否需要 OCR。
索引默认写在当前目录的 sva_index/library.db（SQLite + FTS5 全文检索），全文缓存在 sva_index/text/，不会写入研报库本身。

用法：
  python3 library_index.py ~/研报库                       # 建立/更新索引（只处理新增或改动的文件）
  python3 library_index.py ~/研报库 ~/另一个库 --db my.db  # 多个根目录
  python3 library_index.py ~/研报库 --ocr --lang chi_sim+eng   # 对扫描件做 OCR（较慢）
  python3 library_index.py ~/研报库 --rebuild             # 全部重建
  python3 library_index.py --stats                        # 查看索引概况
"""
import argparse
import datetime as dt
import hashlib
import json
import os
import re
import sqlite3
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import pdf_text  # noqa: E402

BROKERS = {
    "高盛": ["高盛", "goldman sachs", "goldman", "gs research"],
    "摩根士丹利": ["摩根士丹利", "大摩", "morgan stanley"],
    "摩根大通": ["摩根大通", "小摩", "j.p. morgan", "jp morgan", "jpmorgan", "j.p.morgan", "jpm"],
    "伯恩斯坦": ["伯恩斯坦", "bernstein"],
    "汇丰": ["汇丰", "hsbc"],
    "花旗": ["花旗", "citi", "citigroup"],
    "美银": ["美银", "美国银行", "bofa", "bank of america", "merrill"],
    "瑞银": ["瑞银", "ubs"],
    "巴克莱": ["巴克莱", "barclays"],
    "德银": ["德银", "德意志银行", "deutsche bank"],
    "野村": ["野村", "nomura"],
    "麦格理": ["麦格理", "macquarie"],
    "杰富瑞": ["杰富瑞", "jefferies"],
    "富国": ["富国银行", "wells fargo"],
    "瑞穗": ["瑞穗", "mizuho"],
    "傲基": ["evercore"],
    "晨星": ["晨星", "morningstar"],
    "中金": ["中金", "cicc"],
    "中信证券": ["中信证券", "citic securities"],
    "中信建投": ["中信建投"],
    "华泰": ["华泰"],
    "国泰君安": ["国泰君安", "国泰海通"],
    "海通": ["海通"],
    "招商证券": ["招商证券"],
    "广发": ["广发"],
    "申万宏源": ["申万", "申万宏源"],
    "天风": ["天风"],
    "兴业证券": ["兴业证券"],
    "东吴": ["东吴"],
    "浙商": ["浙商"],
    "开源": ["开源证券"],
    "国盛": ["国盛"],
    "长江证券": ["长江证券"],
    "方正": ["方正证券"],
    "民生": ["民生证券"],
}
RATING_WORDS = [
    ("Overweight", "增持"), ("Outperform", "跑赢"), ("Buy", "买入"), ("Accumulate", "增持"),
    ("Neutral", "中性"), ("Hold", "持有"), ("Market-Perform", "中性"), ("Equal-weight", "中性"),
    ("Underweight", "减持"), ("Underperform", "跑输"), ("Reduce", "减持"), ("Sell", "卖出"),
]
CN_RATINGS = ["买入", "增持", "推荐", "强烈推荐", "优于大市", "跑赢行业", "中性", "持有", "减持", "卖出", "跑输"]

TICKER_PATTERNS = [
    r"\b([A-Z]{2,5})\.(?:US|O|N|OQ)\b",            # CRWV.US / CRWV.O
    r"\b([A-Z]{2,5}) US\b",                          # CRWV US
    r"\((?:NASDAQ|NYSE|Nasdaq|NYSE American)?:?\s*([A-Z]{2,5})\)",  # (CRWV) / (NASDAQ: CRWV)
    r"\b(\d{6}\.(?:SH|SZ|BJ))\b",                    # A 股
    r"\b(\d{4,5}\.HK)\b",                            # 港股
]
TICKER_STOP = {"US", "USD", "EPS", "EV", "PE", "GAAP", "FY", "Q", "AI", "GPU", "CEO", "CFO", "NTM", "LTM", "EBIT",
               "EBITDA", "FCF", "ARR", "RPO", "TAM", "IRR", "WACC", "ROE", "ROIC", "DCF", "TP", "PT", "NA", "NM",
               "YOY", "QOQ", "CAGR", "IPO", "ETF", "SOFR", "BPS", "MW", "GW", "OK", "SEC", "FINRA", "LLC", "INC",
               "PLC", "ADR", "THE", "AND", "FOR", "BUY", "SELL", "HOLD", "CHINA", "USA", "UK", "EUR", "GBP", "HKD",
               "RMB", "CNY", "JPY", "OEM", "SPV", "DDTL", "SKU", "API", "CPU", "TPU", "HPC", "ESG", "AOI", "SBC", "MSCI", "SG", "AB", "SA", "NV", "AG", "SE"}

MONTHS = {m: i for i, m in enumerate(
    ["january", "february", "march", "april", "may", "june", "july", "august", "september", "october",
     "november", "december"], 1)}


def norm_date(y, m, d):
    try:
        return dt.date(int(y), int(m), int(d)).isoformat()
    except Exception:
        return None


def date_from_path(rel):
    """从文件名/路径识别日期。优先级：文件名完整日期 > 文件名 YYMMDD > 路径完整日期；只有年月的文件夹留作兜底。"""
    s = rel.replace("\\", "/")
    base = os.path.basename(s)
    full = r"(20\d{2})[-_.年]?(\d{1,2})[-_.月]?(\d{1,2})日?(?![0-9a-zA-Z])"
    for hay in (base, s):
        for m in re.finditer(r"(?<!\d)" + full, hay):
            if len(m.group(0)) >= 8 or re.search(r"[-_.年]", m.group(0)):
                d = norm_date(*m.groups())
                if d:
                    return d
        if hay is base:
            m = re.search(r"(?<!\d)(2\d)(\d{2})(\d{2})(?!\d)", base)   # 文件名中的 YYMMDD，如 -261015
            if m:
                d = norm_date("20" + m.group(1), m.group(2), m.group(3))
                if d:
                    return d
    return None


def month_from_path(rel):
    m = re.search(r"(?<!\d)(20\d{2})[-_.年/](\d{1,2})(?:月)?(?:/|$)", rel.replace("\\", "/"))
    return norm_date(m.group(1), m.group(2), 1) if m else None


def date_from_text(t):
    t = t[:6000]
    m = re.search(r"(20\d{2})年(\d{1,2})月(\d{1,2})日", t)
    if m:
        return norm_date(*m.groups())
    m = re.search(r"\b(\d{1,2})\s+(January|February|March|April|May|June|July|August|September|October|November|"
                  r"December)\s+(20\d{2})\b", t, re.I)
    if m:
        return norm_date(m.group(3), MONTHS[m.group(2).lower()], m.group(1))
    m = re.search(r"\b(January|February|March|April|May|June|July|August|September|October|November|December)"
                  r"\s+(\d{1,2}),\s+(20\d{2})\b", t, re.I)
    if m:
        return norm_date(m.group(3), MONTHS[m.group(1).lower()], m.group(2))
    m = re.search(r"\b(20\d{2})-(\d{2})-(\d{2})\b", t)
    if m:
        return norm_date(*m.groups())
    return None


def broker_of(rel, head):
    hay_path = " " + re.sub(r"[_\-/\\]+", " ", rel.lower()) + " "
    for canon, aliases in BROKERS.items():
        if any((" " + a + " ") in hay_path or (len(a) > 3 and a in hay_path) or
               ("\u4e00" <= a[0] <= "\u9fff" and a in hay_path) for a in aliases):
            return canon
    h = head[:4000].lower()
    best, pos = None, 10 ** 9
    for canon, aliases in BROKERS.items():
        for a in aliases:
            i = h.find(a)
            if 0 <= i < pos:
                best, pos = canon, i
    return best


def tickers_of(head, rel=""):
    found = []
    hay = re.sub(r"_+", " ", os.path.basename(rel)) + "\n" + head[:8000]
    for pat in TICKER_PATTERNS:
        for m in re.finditer(pat, hay):
            t = m.group(1).upper()
            if t not in TICKER_STOP and t not in found:
                found.append(t)
    return found[:8]


def rating_of(head):
    h = head[:6000]
    words = "|".join(re.escape(en) for en, _ in RATING_WORDS)
    cn_of = {en.lower(): cn for en, cn in RATING_WORDS}
    for pat in [r"(?:Rating|评级)[\s\S]{0,250}?\b(%s)\b" % words,          # Rating ... Underperform（可跨行）
                r"Upside[^\n]{0,40}?\b(%s)\b" % words,                     # 高盛封面：... Upside: 72.7%  BUY
                r"\b(%s)\b\s*(?:Rating|rated)" % words]:
        m = re.search(pat, h, re.I)
        if m:
            return cn_of[m.group(1).lower()]
    for cn in CN_RATINGS:
        if re.search(r"(?:评级|投资评级)[：:\s]*%s" % cn, h):
            return cn
    for en, cn in RATING_WORDS:          # 退而求其次：首页独立成行的评级词
        if re.search(r"(?m)^\s*%s\s*$" % re.escape(en), h, re.I):
            return cn
    return None


def target_of(head):
    h = head[:6000]
    pats = [r"Price Target[^\n$¥HK]{0,30}(?:\(\w+-\d+\))?[:：]?\s*(?:US)?\$\s?([\d,]+(?:\.\d+)?)",
            r"(?:目标价|目标价格)[：:\s]*(?:US)?\$?\s?([\d,]+(?:\.\d+)?)",
            r"\bPT\b[^\n]{0,15}\$\s?([\d,]+(?:\.\d+)?)",
            r"Price Target[\s\S]{0,250}?([\d,]+\.\d{2})\s*(?:USD|HKD|CNY|RMB)",
            r"\bTP\s*(?:USD|HKD|CNY|RMB)\s?([\d,]+(?:\.\d+)?)",
            r"Target Price[^\n]{0,30}?([\d,]+(?:\.\d+)?)"]
    for p in pats:
        m = re.search(p, h, re.I)
        if m:
            try:
                return float(m.group(1).replace(",", ""))
            except ValueError:
                pass
    return None


def title_from_filename(rel):
    stem = os.path.splitext(os.path.basename(rel))[0]
    stem = re.sub(r"^[0-9a-f]{6,}-\d+_?", "", stem)            # 去掉"哈希-数字"前缀
    stem = re.sub(r"^[0-9a-f]{8}[-_]", "", stem)                 # 去掉纯哈希前缀
    stem = re.sub(r"[_]+", " ", stem)
    stem = re.sub(r"-?\d{6}(?:\s*\d)?$", "", stem).strip(" -")  # 去掉末尾 YYMMDD
    return stem[:160]


def descriptive(t):
    return len(re.findall(r"[A-Za-z]{3,}", t)) >= 4 or len(re.findall(r"[\u4e00-\u9fff]", t)) >= 8


def title_of(head, rel, scanned=False):
    ft = title_from_filename(rel)
    if scanned or descriptive(ft):          # 文件名本身是描述性标题时优先用文件名
        return ft
    skip = re.compile(r"^(equity research|research|quick take|company update|\d+\s+\w+\s+20\d{2}.*|page \d+.*|"
                      r"rating|price target.*|underperform|outperform|market-perform|neutral|buy|sell|"
                      r".*@.*|\+?\d[\d\s()-]{6,}|.*\d{1,2}:\d{2}\s*[AP]M.*|.*(?:price target|upside:|目标价|market cap).*)$", re.I)
    segs = []
    for line in head.splitlines()[:40]:
        for seg in re.split(r"\s{3,}", line.strip()):
            seg = seg.strip()
            if 12 <= len(seg) <= 160 and not skip.match(seg) and re.search(r"[A-Za-z\u4e00-\u9fff]{3}", seg):
                if not re.fullmatch(r"[A-Za-z0-9]{20,}", seg):          # 跳过水印/乱码串
                    segs.append(seg)
    if not segs:
        return title_from_filename(rel)
    scored = [(len(x) + (40 if ":" in x or "：" in x else 0) - 3 * i, x) for i, x in enumerate(segs[:15])]
    return max(scored)[1][:160]


def lang_of(text):
    s = text[:5000]
    if not s:
        return None
    cjk = sum(1 for ch in s if "一" <= ch <= "鿿")
    return "zh" if cjk > len(s) * 0.15 else "en"


def sha1_of(path):
    h = hashlib.sha1()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


SCHEMA = """
create table if not exists docs(
  id integer primary key, path text unique, root text, rel text, size integer, mtime real, sha1 text,
  broker text, date text, title text, tickers text, rating text, target real, lang text, pages integer,
  backend text, needs_ocr integer, ocr_done integer default 0, indexed_at text);
create virtual table if not exists docs_fts using fts5(title, head, body, tokenize='unicode61');
"""


def connect(db):
    os.makedirs(os.path.dirname(os.path.abspath(db)), exist_ok=True)
    con = sqlite3.connect(db)
    con.executescript(SCHEMA)
    return con


def cjk_spaced(text):
    """FTS5 unicode61 不分中文词：在中文字符之间加空格，使逐字检索可用。"""
    return re.sub(r"([一-鿿])", r" \1 ", text)


def index_file(con, root, path, args):
    rel = os.path.relpath(path, root)
    st = os.stat(path)
    row = con.execute("select id,size,mtime from docs where path=?", (path,)).fetchone()
    if row and row[1] == st.st_size and abs(row[2] - st.st_mtime) < 1e-6 and not args.rebuild \
            and not (args.ocr and con.execute("select needs_ocr and not ocr_done from docs where id=?",
                                              (row[0],)).fetchone()[0]):
        return "skip"
    pages = pdf_text.page_count(path)
    body, backend = pdf_text.extract_text(path, args.max_pages)
    needs_ocr = int(pdf_text.looks_scanned(body, min(pages or 1, args.max_pages or 10 ** 6)))
    ocr_done = 0
    if needs_ocr and args.ocr:
        try:
            body, backend = pdf_text.ocr_pdf(path, args.lang, pages=list(range(1, min(pages or 1,
                                                                                     args.max_pages) + 1)))
            ocr_done = 1
        except Exception as e:
            sys.stderr.write(f"  OCR 失败：{rel}：{e}\n")
    sha1 = sha1_of(path)
    cache = os.path.join(os.path.dirname(os.path.abspath(args.db)), "text")
    os.makedirs(cache, exist_ok=True)
    with open(os.path.join(cache, sha1 + ".txt"), "w", encoding="utf8") as fh:   # 全文缓存，供 prepare_reading 复用
        fh.write(body)
    head = body[:8000]
    meta = dict(
        broker=broker_of(rel, head),
        date=date_from_path(rel) or date_from_text(head) or month_from_path(rel),
        title=title_of(head, rel, bool(needs_ocr and not ocr_done)),
        tickers=",".join(tickers_of(head, rel)),
        rating=rating_of(head),
        target=target_of(head),
        lang=lang_of(body),
    )
    vals = (path, root, rel, st.st_size, st.st_mtime, sha1, meta["broker"], meta["date"], meta["title"],
            meta["tickers"], meta["rating"], meta["target"], meta["lang"], pages, backend, needs_ocr, ocr_done,
            dt.datetime.now().isoformat(timespec="seconds"))
    if row:
        con.execute("""update docs set path=?,root=?,rel=?,size=?,mtime=?,sha1=?,broker=?,date=?,title=?,tickers=?,
                       rating=?,target=?,lang=?,pages=?,backend=?,needs_ocr=?,ocr_done=?,indexed_at=? where id=?""",
                    vals + (row[0],))
        did = row[0]
        con.execute("delete from docs_fts where rowid=?", (did,))
    else:
        cur = con.execute("""insert into docs(path,root,rel,size,mtime,sha1,broker,date,title,tickers,rating,target,
                             lang,pages,backend,needs_ocr,ocr_done,indexed_at) values(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                          vals)
        did = cur.lastrowid
    con.execute("insert into docs_fts(rowid,title,head,body) values(?,?,?,?)",
                (did, cjk_spaced(meta["title"] + " " + rel), cjk_spaced(head), cjk_spaced(body)))
    return "update" if row else "new"


def stats(con):
    n = con.execute("select count(*) from docs").fetchone()[0]
    print(f"索引文档数：{n}")
    for label, q in [("需要 OCR（未完成）", "select count(*) from docs where needs_ocr=1 and ocr_done=0"),
                     ("未识别机构", "select count(*) from docs where broker is null"),
                     ("未识别日期", "select count(*) from docs where date is null"),
                     ("未识别代码", "select count(*) from docs where tickers=''")]:
        print(f"  {label}：{con.execute(q).fetchone()[0]}")
    print("按机构：", dict(con.execute("select coalesce(broker,'未知'),count(*) from docs group by 1 order by 2 desc")))
    print("日期范围：", con.execute("select min(date),max(date) from docs").fetchone())


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("roots", nargs="*")
    ap.add_argument("--db", default="sva_index/library.db")
    ap.add_argument("--max-pages", type=int, default=80, help="每份最多抽取的页数（默认 80）")
    ap.add_argument("--ocr", action="store_true", help="对扫描件做 OCR")
    ap.add_argument("--lang", default="eng", help="OCR 语言，中文研报用 chi_sim+eng")
    ap.add_argument("--rebuild", action="store_true")
    ap.add_argument("--stats", action="store_true")
    args = ap.parse_args()
    con = connect(args.db)
    if args.stats or not args.roots:
        stats(con)
        return
    print("可用后端：", {k: v for k, v in pdf_text.backends().items() if v})
    counts = {"new": 0, "update": 0, "skip": 0, "error": 0}
    seen = set()
    for root in args.roots:
        root = os.path.abspath(os.path.expanduser(root))
        for dp, dn, fn in os.walk(root):
            dn[:] = [d for d in dn if not d.startswith(".") and d != "sva_index"]
            for f in fn:
                if not f.lower().endswith(".pdf"):
                    continue
                p = os.path.join(dp, f)
                seen.add(p)
                try:
                    r = index_file(con, root, p, args)
                    counts[r] += 1
                    if r != "skip":
                        print(f"  [{r}] {os.path.relpath(p, root)}")
                except Exception as e:
                    counts["error"] += 1
                    sys.stderr.write(f"  [error] {p}: {e}\n")
            con.commit()
        gone = [r for r in con.execute("select id,path from docs where root=?", (root,)) if r[1] not in seen]
        for did, _ in gone:
            con.execute("delete from docs where id=?", (did,))
            con.execute("delete from docs_fts where rowid=?", (did,))
        if gone:
            print(f"  移除已不存在的文件索引：{len(gone)}")
    con.commit()
    print("完成：", json.dumps(counts, ensure_ascii=False))
    stats(con)


if __name__ == "__main__":
    main()
