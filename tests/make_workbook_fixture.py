#!/usr/bin/env python3
"""生成 check_delivery.py 第 13 项（图表工作簿结构，格式规范第 13.1 节）的测试工作簿。

用法：python3 tests/make_workbook_fixture.py <输出根目录>
在 <输出根目录>/good、bad、bad_ref 下各写一份 ACME_图表数据_20261105_V1.0.xlsx：
  good     合规：图表总览在第一张、2 张原生图表单列纵向排布（A3、A24）、只引用"图表计算"；语义页"历年营收与业务拆分"带本页图；
           "营收数据来源"；"图表计算"隐藏
  bad      不合规：第一张是"说明"；"图表总览"只有图号索引、没有图表；"图1数据""Chart 2 Data""Sheet1"编号或默认页名；
           没有来源页，也没有图表计算页
  bad_ref  总览图直接引用语义页而不是"图表计算"，其余合规
  bad_layout 总览两张图并排（A3、J3），行区间相交，其余合规
"""
import os
import sys

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference

NAME = "ACME_图表数据_20261105_V1.0.xlsx"
YEARS = [str(y) for y in range(2019, 2026)]
REV = [10.0, 12.5, 15.1, 17.8, 16.2, 19.4, 23.0]


def data_sheet(ws, title):
    ws.append(["口径：亿美元，日历年；来源 S01"])
    ws.append(["年份", title])
    for y, v in zip(YEARS, REV):
        ws.append([y, v])
    ws.freeze_panes = "A3"


def chart_on(ws, src, anchor, kind=BarChart):
    ch = kind()
    ch.title = "营收"
    ch.add_data(Reference(src, min_col=2, min_row=2, max_row=2 + len(YEARS)), titles_from_data=True)
    ch.set_categories(Reference(src, min_col=1, min_row=3, max_row=2 + len(YEARS)))
    ws.add_chart(ch, anchor)


def source_sheet(ws):
    ws.append(["来源 ID", "文件", "链接", "页码或字段"])
    ws.append(["S01", "ACME 10-K 2025", "https://example.com/10k", "第 45 页 合并利润表"])


def good(path, overview_ref_semantic=False, side_by_side=False):
    wb = Workbook()
    ov = wb.active
    ov.title = "图表总览"
    ov.append(["起始年份", 2019, "结束年份", 2025])
    sem = wb.create_sheet("历年营收与业务拆分")
    data_sheet(sem, "营收")
    chart_on(sem, sem, "D3")
    source_sheet(wb.create_sheet("营收数据来源"))
    calc = wb.create_sheet("图表计算")
    data_sheet(calc, "营收")
    calc.sheet_state = "hidden"
    src = sem if overview_ref_semantic else calc
    chart_on(ov, src, "A3")
    chart_on(ov, src, "J3" if side_by_side else "A24", LineChart)   # 单列纵向（0.21.0 起并排的图算重叠）
    wb.save(path)


def bad(path):
    wb = Workbook()
    wb.active.title = "说明"
    ov = wb.create_sheet("图表总览")
    ov.append(["图号", "内容", "数据页"])
    ov.append(["图 1", "营收", "图1数据"])
    d1 = wb.create_sheet("图1数据")
    data_sheet(d1, "营收")
    chart_on(d1, d1, "D3")
    data_sheet(wb.create_sheet("Chart 2 Data"), "营收")
    wb.create_sheet("Sheet1")
    wb.save(path)


def main():
    root = sys.argv[1]
    for sub in ("good", "bad", "bad_ref", "bad_layout"):
        os.makedirs(os.path.join(root, sub), exist_ok=True)
    good(os.path.join(root, "good", NAME))
    bad(os.path.join(root, "bad", NAME))
    good(os.path.join(root, "bad_ref", NAME), overview_ref_semantic=True)
    good(os.path.join(root, "bad_layout", NAME), side_by_side=True)
    print(f"图表工作簿样例：{root}/good、bad、bad_ref、bad_layout")


if __name__ == "__main__":
    main()
