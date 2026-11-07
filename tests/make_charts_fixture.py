#!/usr/bin/env python3
"""check_charts.py 的测试样例：在给定目录生成 charts_good.xlsx 与 charts_bad.xlsx。
用法：python3 tests/make_charts_fixture.py <输出目录>"""
import datetime as dt
import os
import sys

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference

CM_PER_PX = 2.54 / 96
out = sys.argv[1]
os.makedirs(out, exist_ok=True)


def line(ws, col, rows, w_px, h_px, title, cat_col=1, marker=None):
    ch = LineChart()
    ch.title = title
    ch.add_data(Reference(ws, min_col=col, min_row=1, max_row=rows + 1), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=cat_col, min_row=2, max_row=rows + 1))
    ch.width, ch.height = w_px * CM_PER_PX, h_px * CM_PER_PX
    if marker:
        ch.series[0].marker.symbol = marker
    return ch


# 合规：月末 336 点独占一行 1180 px、横轴每 2 年标一次；下一行两张短序列各 600 px 并排
wb = Workbook()
ws = wb.active
ws.title = "图表计算"
ws.append(["期间", "月末指数", "年份", "年度占比"])
d = dt.date(1998, 12, 31)
for i in range(336):
    y, m = 1999 + i // 12, i % 12 + 1
    label = str(y) if m == 1 and y % 2 == 0 else None
    ws.append([label, 100 + i * 0.5, 2014 + i if i < 12 else None, 10 + i % 5 if i < 12 else None])
book = wb.create_sheet("图表数据")
book.add_chart(line(ws, 2, 336, 1180, 540, "XLF 与 SPY 月末总回报指数"), "A1")
bar = BarChart()
bar.title = "金融板块占比"
bar.add_data(Reference(ws, min_col=4, min_row=1, max_row=13), titles_from_data=True)
bar.set_categories(Reference(ws, min_col=3, min_row=2, max_row=13))
bar.width, bar.height = 600 * CM_PER_PX, 384 * CM_PER_PX
book.add_chart(bar, "A30")
book.add_chart(line(ws, 4, 12, 600, 384, "年度占比折线", cat_col=3), "K30")
wb.save(os.path.join(out, "charts_good.xlsx"))

# 不合规：一行三张；日数据 2000 点塞进 590 px；每天都有横轴标签；折线带圆点
wb = Workbook()
ws = wb.active
ws.title = "图表计算"
ws.append(["日期", "日收盘"])
d = dt.date(2018, 1, 1)
for i in range(2000):
    ws.append([(d + dt.timedelta(days=i)).isoformat(), 100 + (i % 37)])
book = wb.create_sheet("图表数据")
book.add_chart(line(ws, 2, 2000, 590, 384, "日收盘", marker="circle"), "A1")
book.add_chart(line(ws, 2, 12, 590, 384, "短序列一"), "K1")
book.add_chart(line(ws, 2, 12, 590, 384, "图 3　短序列二"), "U1")
wb.save(os.path.join(out, "charts_bad.xlsx"))
