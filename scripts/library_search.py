#!/usr/bin/env python3
"""在研报库索引中检索相关研报，输出排序结果和阅读清单（manifest）。

用法示例：
  python3 library_search.py --ticker CRWV --name CoreWeave
  python3 library_search.py --ticker CRWV --name CoreWeave --kw DDTL --kw 残值 --since 2026-05-01
  python3 library_search.py --name 宁德时代 --ticker 300750.SZ --broker 中金 --broker 高盛
  python3 library_search.py --kw "GPU 租赁" --kw neocloud --limit 50 --out manifest.json

打分：代码命中（元数据）+20/个；公司名或别名出现在标题 +8；
      每个检索词在全文中的出现次数 2.5×ln(1+次数)（每词上限 8）+ BM25；近一年 +2。
至少命中一个代码、名称或关键词才会入选。同一文件（SHA1 相同）只保留一份。
"""
import argparse
import datetime as dt
import json
import math
import os
import re
import sqlite3
import sys


def cjk_spaced(text):
    return re.sub(r"([一-鿿])", r" \1 ", text)


def fts_phrase(term):
    t = cjk_spaced(term).strip()
    t = re.sub(r"\s+", " ", t).replace('"', "")
    return f'"{t}"'


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--db", default="sva_index/library.db")
    ap.add_argument("--ticker", action="append", default=[], help="股票代码，可多次")
    ap.add_argument("--name", action="append", default=[], help="公司名或别名，可多次")
    ap.add_argument("--kw", action="append", default=[], help="主题关键词，可多次")
    ap.add_argument("--broker", action="append", default=[], help="只看这些机构")
    ap.add_argument("--since")
    ap.add_argument("--until")
    ap.add_argument("--limit", type=int, default=30)
    ap.add_argument("--out", help="把结果写成阅读清单 JSON")
    ap.add_argument("--snippet", action="store_true", help="显示命中上下文")
    a = ap.parse_args()
    if not os.path.exists(a.db):
        sys.exit(f"找不到索引 {a.db}，请先运行 library_index.py")
    con = sqlite3.connect(a.db)
    con.row_factory = sqlite3.Row

    terms = [t for t in a.ticker + a.name + a.kw if t.strip()]
    fts = {}
    if terms:
        q = " OR ".join(fts_phrase(t) for t in terms)
        sql = ("select rowid, bm25(docs_fts, 5.0, 2.0, 1.0) as s, "
               "snippet(docs_fts, 2, '[', ']', ' … ', 12) as snip from docs_fts where docs_fts match ?")
        for r in con.execute(sql, (q,)):
            fts[r["rowid"]] = (-r["s"], r["snip"])

    today = dt.date.today()
    rows = []
    for d in con.execute("select * from docs"):
        if a.broker and (d["broker"] or "") not in a.broker:
            continue
        if a.since and (not d["date"] or d["date"] < a.since):
            continue
        if a.until and (not d["date"] or d["date"] > a.until):
            continue
        tk = set((d["tickers"] or "").split(",")) - {""}
        title = (d["title"] or "") + " " + (d["rel"] or "")
        score, why = 0.0, []
        hit_t = [t for t in a.ticker if t.upper() in tk]
        if hit_t:
            score += 20 * len(hit_t); why.append("代码:" + "/".join(hit_t))
        hit_n = [n for n in a.name + a.ticker if n.lower() in title.lower()]
        if hit_n:
            score += 8; why.append("标题:" + "/".join(hit_n))
        f = fts.get(d["id"])
        if f:
            body = (con.execute("select body from docs_fts where rowid=?", (d["id"],)).fetchone() or [""])[0]
            low = body.lower()
            freq = 0.0
            for t in terms:
                n = low.count(cjk_spaced(t).strip().lower()) if re.search(r"[\u4e00-\u9fff]", t) else \
                    len(re.findall(r"\b%s\b" % re.escape(t.lower()), low))
                freq += min(8.0, 2.5 * math.log1p(n))
            score += freq + max(f[0], 0); why.append(f"全文{freq:.1f}")
        if terms and not (hit_t or hit_n or f):
            continue
        if d["date"]:
            try:
                if (today - dt.date.fromisoformat(d["date"])).days <= 365:
                    score += 2
            except ValueError:
                pass
        rows.append((score, d, why, f[1] if f else ""))

    rows.sort(key=lambda x: (x[0], x[1]["date"] or ""), reverse=True)   # 分数高在前，同分新报告在前
    seen, out = set(), []
    for score, d, why, snip in rows:
        if d["sha1"] in seen:
            continue
        seen.add(d["sha1"])
        out.append(dict(score=round(score, 1), date=d["date"], broker=d["broker"], title=d["title"],
                        tickers=d["tickers"], rating=d["rating"], target=d["target"], pages=d["pages"],
                        lang=d["lang"], needs_ocr=bool(d["needs_ocr"] and not d["ocr_done"]), path=d["path"], sha1=d["sha1"],
                        why="；".join(why), snippet=snip))
        if len(out) >= a.limit:
            break

    print(f"命中 {len(out)} 份（检索词：{terms or '无'}）")
    print(f"{'分数':>5}  {'日期':<10}  {'机构':<6}  {'评级':<4}  {'目标价':>7}  {'页':>3}  OCR  标题")
    for r in out:
        tp = f"{r['target']:.0f}" if r["target"] else ""
        print(f"{r['score']:>5}  {r['date'] or '':<10}  {(r['broker'] or '?'):<6}  {(r['rating'] or ''):<4}  "
              f"{tp:>7}  {r['pages'] or '':>3}  {'需' if r['needs_ocr'] else '  '}   {r['title'][:70]}")
        if a.snippet and r["snippet"]:
            print("        …", re.sub(r"\s+", " ", r["snippet"])[:160])
    if a.out:
        manifest = dict(db=os.path.abspath(a.db), created=dt.datetime.now().isoformat(timespec="seconds"),
                        query=dict(ticker=a.ticker, name=a.name, kw=a.kw, broker=a.broker, since=a.since,
                                   until=a.until), docs=out)
        with open(a.out, "w", encoding="utf8") as fh:
            json.dump(manifest, fh, ensure_ascii=False, indent=2)
        print(f"阅读清单已写入 {a.out}；下一步：python3 prepare_reading.py {a.out} <输出目录>")


if __name__ == "__main__":
    main()
