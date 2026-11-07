#!/usr/bin/env python3
"""check_numbers.py 的测试样例：在给定目录生成 hist.xlsx（虚构 ACME 三年数据）与 数字台账.csv。
用法：python3 tests/make_numbers_fixture.py <输出目录>"""
import os
import sys

from openpyxl import Workbook

out = sys.argv[1]
os.makedirs(out, exist_ok=True)
wb = Workbook()
ws = wb.active
ws.title = "年度衍生指标"
ws.append(["年度", "营收（亿美元）", "毛利率", "营业利润率"])
for row in [(2019, 40.0, 0.6, 0.12), (2020, 46.0, 0.6, 0.135), (2021, 52.9, 0.6, 0.14)]:
    ws.append(row)
ws["F2"] = "=B2+B3"            # 没有缓存值的公式：脚本应报"先重算"
wb.save(os.path.join(out, "hist.xlsx"))

with open(os.path.join(out, "数字台账.csv"), "w", encoding="utf-8") as fh:
    fh.write("关键数字,数值,单位,位置,出处类型,出处,页码或单元格,备注,数据截至,核心,复核状态\n")
    fh.write("三年合计营收,138.9,亿美元,二、2.1,计算,B2+B3+B4,,,FY2021,否,未复核\n")
    fh.write("收购交易价值,3.1,亿美元,二、2.1,S01,10-K,p45,,FY2021,是,未复核\n")
    fh.write("续约率门槛,90%,%,一、1.1,口径,,,,,否,未复核\n")
    fh.write("营收增速,15%,%,二、2.1,计算,,,,,否,未复核\n")
    fh.write("同行营收 ACME,80.5,亿美元,二、2.2,S02,年报,p3,,FY2021,否,未复核\n")
    fh.write("同行营收 对手甲,120.0,亿美元,二、2.2,S03,年报,p5,,FY2021,否,未复核\n")
