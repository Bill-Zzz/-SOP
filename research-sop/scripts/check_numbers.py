#!/usr/bin/env python3
"""数字比对：报告正文里的每个数字，能不能在历史数据表、图表数据表或数字台账里找到。

用法：
  python3 check_numbers.py 报告.md [--xlsx 历史数据.xlsx ...] [--ledger 数字台账.csv ...] [--csv 其他数据.csv ...]
                           [--ignore 正则 ...]

做什么：
  1. 从 .md 正文抽数字（含表格与图注）：带小数的数、带数据单位的数（%、百分点、pp、bp、倍、x、亿、万、美元、元、股等）、
     不带单位但 ≥ 100 的整数。不抽：标题行、图片路径、来源行与"数据与可编辑图表"行、代码块、HTML 注释、链接地址、
     "附：口径与说明"及其后各节；日期（2025-03-31、2025 年 3 月、3 月 31 日）、年份（1900–2100 不带数据单位的整数）、
     图号表号（图 3、表 2）、版本号（V1.0）、与字母连写的编号（FY2025、1Q26、2026H1、fig3）、
     带口径单位的数（3σ、250 个交易日、第 2.1 节）、不带数据单位的 100 以下整数（3 家、5 年）。
  2. 候选值：--xlsx 里全部数值单元格（读缓存值；公式单元格没有缓存值的报个数，先用 scripts/lo_recalc.sh 重算）、
     --ledger 台账"数值"列、--csv 里全部数值。每个候选值同时按 ×1、×100、÷100、×1 万、÷1 万、÷10 万、÷1 亿 换算，
     覆盖小数与百分比、百万与亿、元与万元或亿元。
  3. 判定：正文数字按它自己的小数位数，与某个候选值（取绝对值）四舍五入后相等即命中；负数、括号负数按绝对值比。

输出：数字总数、命中数（台账 / 数据表 / 其他 csv）、未命中清单（所在章节 ｜ 前后文 ｜ 数字）。
未命中的：正文写错了就改正文；是计算所得、口径常数或表外来源的数，就补一行台账（出处类型写"计算""口径"或来源编号）。
全部命中退出码 0，有未命中 1，参数错误 2。

局限：只证明"这个数在数据里出现过"，不证明出现在对的行、对的年份；精度低的数（如 12%）可能碰巧撞上别的单元格。
核心数字对不对，由 research-sop 工单 B 抽查对回原始出处。
依赖：openpyxl（只有给了 --xlsx 才需要）；其余标准库。
"""
import argparse
import csv
import math
import os
import re
import sys

DATA_UNITS = (
    "个百分点", "百分点", "亿美元", "万美元", "亿元", "万元", "亿港元", "亿股", "万股", "美元", "港元", "元",
    "亿", "万", "%", "％", "pp", "bp", "倍", "x", "X", "股", "名员工",
)
SKIP_UNITS = (
    "σ", "个交易日", "交易日", "个季度", "季度", "个月", "月", "年", "周", "日", "天", "小时",
    "节", "章", "页", "版", "条", "项", "份", "篇", "张", "号",
)
FACTORS = (1, 100, 0.01, 1e4, 1e-4, 1e-5, 1e-8)
DEFAULT_IGNORE = (r"S&P\s?500", r"标普\s?500", r"纳斯达克\s?100", r"Nasdaq[-\s]?100", r"罗素\s?2000", r"Russell\s?2000",
                  r"道琼斯\s?30")

NUM_RE = re.compile(r"\d{1,3}(?:,\d{3})+(?:\.\d+)?|\d+(?:\.\d+)?")
DATE_RES = (
    re.compile(r"\d{4}[-/.]\d{1,2}[-/.]\d{1,2}"),
    re.compile(r"\d{4}[-/]\d{1,2}(?!\d)"),
    re.compile(r"\d{4}\s?年\s?\d{1,2}\s?月(?:\s?\d{1,2}\s?日)?"),
    re.compile(r"\d{1,2}\s?月\s?\d{1,2}\s?日"),
    re.compile(r"\d{8}"),
)
LINK_RE = re.compile(r"\]\([^)]*\)")
CODE_RE = re.compile(r"`[^`]*`")
COMMENT_RE = re.compile(r"<!--.*?-->", re.S)
PREFIX_SKIP = ("图", "表", "附", "第", "V", "v", "版本", "fig")
SOURCE_LINE = re.compile(r"^\s*(?:[-*]\s*)?(?:资料来源|来源|数据与可编辑图表)\s*[：:]")
HEADING = re.compile(r"^\s*#{1,6}\s+(.*)$")


def to_float(s):
    try:
        return float(str(s).replace(",", "").replace("，", "").strip())
    except ValueError:
        return None


def keys_for(value):
    """一个候选值按各换算系数、0–4 位小数四舍五入后的字符串键。"""
    out = set()
    for f in FACTORS:
        w = abs(value * f)
        if w == 0 or w > 1e13 or w < 1e-6:
            continue
        for d in range(5):
            out.add((d, f"{w:.{d}f}"))
            half_up = math.floor(w * 10 ** d + 0.5) / 10 ** d
            out.add((d, f"{half_up:.{d}f}"))
    return out


def add_value(index, value, tag):
    if value is None or isinstance(value, bool):
        return
    try:
        v = float(value)
    except (TypeError, ValueError):
        return
    if math.isnan(v) or math.isinf(v):
        return
    for k in keys_for(v):
        index.setdefault(k, tag)


def load_ledger(path, index):
    n = 0
    with open(path, encoding="utf-8-sig", newline="") as fh:
        rows = list(csv.reader(fh))
    if not rows:
        return 0
    head = [h.strip() for h in rows[0]]
    col = head.index("数值") if "数值" in head else None
    for row in rows[1:]:
        cells = [row[col]] if col is not None and col < len(row) else row
        for cell in cells:
            for m in NUM_RE.finditer(cell or ""):
                add_value(index, to_float(m.group()), "台账")
                n += 1
    return n


def load_csv(path, index):
    n = 0
    with open(path, encoding="utf-8-sig", newline="") as fh:
        for row in csv.reader(fh):
            for cell in row:
                v = to_float(cell)
                if v is not None:
                    add_value(index, v, "csv")
                    n += 1
    return n


def load_xlsx(path, index):
    try:
        from openpyxl import load_workbook
    except ImportError:
        sys.exit("缺少 openpyxl：pip install openpyxl --break-system-packages")
    wb_v = load_workbook(path, data_only=True, read_only=True)
    wb_f = load_workbook(path, data_only=False, read_only=True)
    n, missing_cache = 0, 0
    for ws_v, ws_f in zip(wb_v.worksheets, wb_f.worksheets):
        for row_v, row_f in zip(ws_v.iter_rows(values_only=True), ws_f.iter_rows(values_only=True)):
            for val, raw in zip(row_v, row_f):
                if isinstance(val, (int, float)) and not isinstance(val, bool):
                    add_value(index, val, "数据表")
                    n += 1
                elif val is None and isinstance(raw, str) and raw.startswith("="):
                    missing_cache += 1
    return n, missing_cache


def body_lines(text):
    """逐行给出 (章节, 行文本)；去掉不该抽数的行和片段。"""
    text = COMMENT_RE.sub("", text)
    section, in_code, stop = "（标题前）", False, False
    for line in text.splitlines():
        if line.strip().startswith("```"):
            in_code = not in_code
            continue
        if in_code:
            continue
        m = HEADING.match(line)
        if m:
            title = m.group(1).strip()
            if re.match(r"^附\s*[：:]\s*口径", title):
                stop = True                                           # 外壳规定它是最后一节
            section = title
            continue
        if stop or SOURCE_LINE.match(line):
            continue
        img = re.match(r"^\s*!\[(.*?)\]\(", line)
        if img:
            line = img.group(1)                                       # 图注里的结论句也要对
        line = LINK_RE.sub("]", line)
        line = CODE_RE.sub(" ", line)
        line = re.sub(r"^\s*(?:[-*+]|\d+[.)、])\s+", "", line)          # 列表符号与列表序号
        line = re.sub(r"^\s*\|?\s*[-:| ]+\|?\s*$", "", line)            # 表格分隔行
        line = re.sub(r"^\s*图\s*\d+\s*[　 ]", "", line)                 # 图注开头的图号
        yield section, line


def extract(line, ignore_res):
    masked = line
    for rx in list(DATE_RES) + ignore_res:
        masked = rx.sub(lambda m: " " * len(m.group()), masked)
    out = []
    for m in NUM_RE.finditer(masked):
        s, e = m.start(), m.end()
        tok = m.group()
        before = masked[max(0, s - 4):s]
        after = masked[e:e + 6].lstrip(" ")
        prev = masked[s - 1] if s > 0 else ""
        nxt = masked[e] if e < len(masked) else ""
        if re.match(r"[A-Za-z]", prev) or (re.match(r"[A-Za-z]", nxt) and not re.match(r"(x|X|pp|bp)(?![A-Za-z])", masked[e:e + 3])):
            continue                                                  # FY2025、1Q26、2026H1、fig3、V1.0
        if prev == "." or (nxt == "." and e + 1 < len(masked) and masked[e + 1].isdigit()):
            continue
        if any(before.rstrip().endswith(p) for p in PREFIX_SKIP):
            continue                                                  # 图 3、表 2、第 3、附一
        unit = next((u for u in DATA_UNITS if after.startswith(u)), "")
        skip_unit = next((u for u in SKIP_UNITS if after.startswith(u)), "")
        if not unit and skip_unit:
            continue
        value = to_float(tok)
        decimals = len(tok.split(".")[1]) if "." in tok else 0
        is_int = decimals == 0
        if is_int and "," not in tok and 1900 <= value <= 2100 and not unit:
            continue                                                  # 年份
        if is_int and value < 100 and not unit:
            continue                                                  # 3 家、5 年、序号
        ctx = line[max(0, s - 14):min(len(line), e + 10)].strip()
        out.append((tok, value, decimals, ctx))
    return out


def main():
    ap = argparse.ArgumentParser(description="报告正文数字与数据表、数字台账比对")
    ap.add_argument("md")
    ap.add_argument("--xlsx", action="append", default=[], help="历史数据表或图表数据表（可多次）")
    ap.add_argument("--ledger", action="append", default=[], help="数字台账.csv（可多次）")
    ap.add_argument("--csv", action="append", default=[], help="其他数值 csv（可多次）")
    ap.add_argument("--ignore", action="append", default=[], help="额外忽略的正则（可多次），如指数名")
    a = ap.parse_args()

    for p in [a.md] + a.xlsx + a.ledger + a.csv:
        if not os.path.exists(p):
            print(f"找不到文件：{p}", file=sys.stderr)
            sys.exit(2)
    if not (a.xlsx or a.ledger or a.csv):
        print("至少给一个 --xlsx、--ledger 或 --csv", file=sys.stderr)
        sys.exit(2)

    index, notes = {}, []
    for p in a.ledger:                                              # 台账先载入，命中优先记"台账"
        notes.append(f"台账 {os.path.basename(p)}：{load_ledger(p, index)} 个值")
    for p in a.xlsx:
        n, miss = load_xlsx(p, index)
        notes.append(f"数据表 {os.path.basename(p)}：{n} 个数值单元格" + (f"；{miss} 个公式没有缓存值（先重算）" if miss else ""))
    for p in a.csv:
        notes.append(f"csv {os.path.basename(p)}：{load_csv(p, index)} 个值")

    ignore_res = [re.compile(x) for x in DEFAULT_IGNORE + tuple(a.ignore)]
    with open(a.md, encoding="utf-8") as fh:
        text = fh.read()

    total, hits, misses = 0, {"台账": 0, "数据表": 0, "csv": 0}, []
    for section, line in body_lines(text):
        for tok, value, d, ctx in extract(line, ignore_res):
            total += 1
            key = (d, f"{abs(value):.{d}f}")
            tag = index.get(key) if d <= 4 else None
            if tag:
                hits[tag] += 1
            else:
                misses.append((section, ctx, tok))

    print("候选值：" + "；".join(notes))
    print(f"数字比对：正文数字 {total} 个，命中 {total - len(misses)} 个"
          f"（台账 {hits['台账']}、数据表 {hits['数据表']}、csv {hits['csv']}），未命中 {len(misses)} 个")
    if misses:
        print("未命中（正文写错就改正文；计算所得、口径常数、表外来源的数补一行台账）：")
        for section, ctx, tok in misses:
            print(f"  「{section}」｜{ctx}｜{tok}")
    sys.exit(1 if misses else 0)


if __name__ == "__main__":
    main()
