#!/usr/bin/env python3
"""科目映射优先于默认候选（0.23.2）。

起因：VRT 的科目映射把"资本开支"指向富途"资本性支出"，但这个名字本来就在默认候选里、排在
"固定资产交易净额"后面，旧的 cands() 只在映射名不在默认候选里时才前置，映射被静默忽略，
取到了处置固定资产所得（正数），资本开支与自由现金流全错，--check 用同一套候选所以也查不出。

用法：python3 tests/test_mapping_priority.py <临时目录>
"""
import csv
import glob
import os
import subprocess
import sys

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
SCRIPTS = os.path.join(HERE, "..", "research-sop", "scripts")
sys.path.insert(0, SCRIPTS)
import build_history_xlsx as bh  # noqa: E402

tmp = sys.argv[1]
root, out = os.path.join(tmp, "futu_map"), os.path.join(tmp, "hist_map")

# 1) 候选顺序
ex = bh.Export.__new__(bh.Export)
ex.mapping = {}
default = list(bh.CORE_BY_KEY["capex"][3])
assert ex.cands("capex") == default, "没有映射时默认候选顺序不变"
ex.mapping = {"capex": "资本性支出"}
assert ex.cands("capex") == ["资本性支出"] + [x for x in default if x != "资本性支出"], ex.cands("capex")
ex.mapping = {"capex": "购建长期资产"}
assert ex.cands("capex") == ["购建长期资产"] + default, ex.cands("capex")

# 2) 端到端：两个科目都有数且不同，映射指向"资本性支出"
subprocess.run([sys.executable, os.path.join(HERE, "make_history_fixture.py"), root],
               check=True, stdout=subprocess.DEVNULL)
long_csv = os.path.join(root, "ACME", "ACME_富途三表_长表.csv")
with open(long_csv, encoding="utf-8-sig") as f:
    rows = list(csv.DictReader(f))
extra, want = [], {}
for r in rows:
    if r["科目"] == "固定资产交易净额":
        true_capex = -abs(float(r["数值"])) * 2
        e = dict(r, 科目="资本性支出", 数值=str(true_capex))
        extra.append(e)
        r["数值"] = str(abs(float(r["数值"])) * 0.1)  # 处置所得，正数
        if r["报告类型"] == "7":
            want[int(r["财年"])] = true_capex
assert want, "样例长表里应有年度的固定资产交易净额"
with open(long_csv, "w", encoding="utf-8-sig", newline="") as f:
    w = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
    w.writeheader()
    w.writerows(rows + extra)
os.makedirs(os.path.join(root, "ACME", "底稿CSV"), exist_ok=True)
cn = next(k for k, v in bh.MAP_NAME.items() if v == "capex")
with open(os.path.join(root, "ACME", "底稿CSV", "ACME_科目映射.csv"), "w", encoding="utf-8-sig", newline="") as f:
    w = csv.writer(f)
    w.writerow(["中文名称", "公司原文名称"])
    w.writerow([cn, "资本性支出"])

res = subprocess.run([sys.executable, os.path.join(SCRIPTS, "build_history_xlsx.py"), "ACME", "--root", root,
                      "--out", out, "--ai", "测试", "--check"], capture_output=True, text=True)
assert "结果：12/12 通过" in res.stdout, res.stdout[-2000:] + res.stderr[-2000:]

wb = openpyxl.load_workbook(glob.glob(out + "/*_ACME_历史数据_v1_测试.xlsx")[0])
ws = wb["年度三表"]
hdr = next(r for r in ws.iter_rows(min_row=1, max_row=8) if sum(isinstance(c.value, int) for c in r) >= 3)
ycol = {c.value: c.column for c in hdr if isinstance(c.value, int)}
row = next(r for r in ws.iter_rows(min_row=hdr[0].row + 1) if str(r[2].value or "").strip() == cn)
assert "资本性支出" in str(row[1].value), f"原始科目名应是映射科目：{row[1].value}"
for y, v in want.items():
    g = ws.cell(row[0].row, ycol[y]).value
    assert isinstance(g, (int, float)) and g < 0, f"{y} 资本开支应为负数（按映射取资本性支出），实际 {g}"
    assert abs(g * 1e8 - v) <= abs(v) * 1e-6, f"{y} 资本开支应取映射科目：期望 {v}，实际 {g * 1e8}"
note = " ".join(str(c.value) for r in wb["年度衍生指标"].iter_rows(min_row=1, max_row=3) for c in r if c.value)
assert "按科目映射取富途'资本性支出'" in note, note[:300]
print("科目映射优先：通过")
