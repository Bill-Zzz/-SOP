#!/usr/bin/env python3
"""生成历史数据 Excel 脚本的最小输入：一家虚构公司 ACME 的富途长表与拆股调整日线。

用法：python3 tests/make_history_fixture.py <futu_export 根目录>
生成 <根目录>/ACME/ACME_富途三表_长表.csv、ACME_日线_库内拆股调整.csv。
数据满足：毛利 = 营收 − 成本；资产 = 负债 + 权益；四个单季之和 = 全年；
FY2023 有一天 +12% 的跳涨（股价事件回归）。
"""
import csv
import datetime as dt
import os
import sys

root = sys.argv[1] if len(sys.argv) > 1 else "tests/fixtures/futu_export"
d = os.path.join(root, "ACME")
os.makedirs(d, exist_ok=True)
SHARES = 1_000_000_000  # 10 亿股


def quarter_ends(fy):
    return [f"{fy}-03-31", f"{fy}-06-30", f"{fy}-09-30", f"{fy}-12-31"]


def is_items(rev):
    cost = round(rev * 0.4)
    op = round(rev * 0.25)
    pti = round(op * 1.02)
    tax = round(pti * 0.2)
    ni = pti - tax
    eps = round(ni / SHARES, 2)
    return {"营业总收入": rev, "总收入": rev, "营业总成本": cost, "毛利": rev - cost, "营业利润": op,
            "税前利润": pti, "所得税": tax, "归属于母公司股东净利润": ni, "净利润": ni,
            "归属于普通股股东净利润": ni, "基本每股收益": eps, "稀释每股收益": round(ni / (SHARES * 1.01), 2)}


def cf_items(rev):
    return {"经营活动现金流量净额": round(rev * 0.3), "固定资产交易净额": -round(rev * 0.1),
            "折旧摊销及损耗": round(rev * 0.05), "普通股发行/回购的净额": -round(rev * 0.02),
            "自由现金流": round(rev * 0.2)}


def bs_items(scale):
    ta = round(10e9 * scale)
    tl = round(ta * 0.45)
    eq = ta - tl
    return {"资产合计": ta, "负债合计": tl, "股东权益合计": eq, "归属于母公司股东权益合计": eq,
            "-现金和现金等价物": round(ta * 0.1), "现金及现金等价物和短期投资": round(ta * 0.2),
            "-应收账款净额": round(ta * 0.08), "存货": round(ta * 0.05), "-应付账款": round(ta * 0.06),
            "-长期借款": round(ta * 0.15)}


rows = []
for fy in range(2019, 2026):
    annual = 4e9 * (1.15 ** (fy - 2019))
    qrev = [round(annual * w) for w in (0.23, 0.24, 0.25, 0.28)]
    nq = 2 if fy == 2025 else 4  # 最新财年只到 Q2
    for i in range(nq):
        per, end = f"{fy}/Q{i + 1}", quarter_ends(fy)[i]
        for st, items in (("利润表", is_items(qrev[i])), ("现金流量表", cf_items(qrev[i])),
                          ("资产负债表", bs_items(1.1 ** (fy - 2019) * (1 + i * 0.02)))):
            for k, v in items.items():
                rows.append([st, per, end, fy, i + 1, "USD", k, float(v), "是"])
    if nq == 4:
        rev = sum(qrev)
        is_y = {k: sum(is_items(q)[k] for q in qrev) for k in is_items(1)}
        is_y["基本每股收益"] = round(is_y["归属于普通股股东净利润"] / SHARES, 2)
        is_y["稀释每股收益"] = round(is_y["归属于普通股股东净利润"] / (SHARES * 1.01), 2)
        cf_y = {k: sum(cf_items(q)[k] for q in qrev) for k in cf_items(1)}
        for st, items in (("利润表", is_y), ("现金流量表", cf_y), ("资产负债表", bs_items(1.1 ** (fy - 2019) * 1.06))):
            for k, v in items.items():
                rows.append([st, f"{fy}/FY", f"{fy}-12-31", fy, 7, "USD", k, float(v), "是"])
        rows.append(["关键指标", f"{fy}/FY", f"{fy}-12-31", fy, 7, "USD", "毛利率", 60.0, "是"])

with open(os.path.join(d, "ACME_富途三表_长表.csv"), "w", encoding="utf-8-sig", newline="") as f:
    w = csv.writer(f)
    w.writerow(["报表", "期间", "期末日", "财年", "报告类型", "币种", "科目", "数值", "日期已校正"])
    w.writerows(rows)

with open(os.path.join(d, "ACME_日线_库内拆股调整.csv"), "w", encoding="utf-8-sig", newline="") as f:
    w = csv.writer(f)
    w.writerow(["time_key", "open", "high", "low", "close", "adj_close", "volume", "turnover", "来源"])
    day, px = dt.date(2018, 6, 1), 20.0
    while day <= dt.date(2025, 7, 31):
        if day.weekday() < 5:
            px *= 1.12 if day == dt.date(2023, 3, 15) else 1.0004
            w.writerow([day.isoformat(), px, px, px, px, px, 1e6, "", "yhh"])
        day += dt.timedelta(days=1)
print("已生成", d)
