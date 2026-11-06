#!/usr/bin/env python3
"""check_delivery.py 第 13 项的版面检查（转换后修补清单 S5、S6）与 build_history_xlsx.py 单列排布的单元测试。

用法：python3 tests/test_chart_layout.py <临时目录>     （由 tests/run_tests.sh 第 7 节调用）
测试数据全部用 openpyxl 现场生成，不依赖外部文件：
  1  两张图并排（A5、H5）→ 行区间相交，不通过，打印相交的两张图与行区间
  2  两张图单列纵向（A5、A26，720 × 360 px）→ 通过
  3  上一张底边正好贴着下一张顶边 → 不算相交
  4  twoCellAnchor（from/to）上下错开但区间部分重合 → 不通过；absoluteAnchor 也能读
  5  1000 行空行（模拟 Google 表格默认行数）→ 不通过；内容 + 50 行以内 → 通过；只有样式的空单元格不算内容
  6  build_history_xlsx.chart_anchors / trim_target：锚点同列、行距 ⌈360 / 21⌉ + 3 = 21 行，图宽高 720 × 360 px
"""
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", "research-sop", "scripts"))

from openpyxl import Workbook  # noqa: E402
from openpyxl.chart import BarChart, Reference  # noqa: E402
from openpyxl.drawing.spreadsheet_drawing import AbsoluteAnchor, AnchorMarker, TwoCellAnchor  # noqa: E402
from openpyxl.drawing.xdr import XDRPoint2D, XDRPositiveSize2D  # noqa: E402
from openpyxl.styles import PatternFill  # noqa: E402

import build_history_xlsx as B  # noqa: E402
import check_delivery as C  # noqa: E402

EMU_PX = 9525


def chart(ws, title="营收"):
    ch = BarChart()
    ch.title = title
    ch.add_data(Reference(ws, min_col=2, min_row=1, max_row=4), titles_from_data=True)
    ch.width, ch.height = B.CHART_W_CM, B.CHART_H_CM
    return ch


def book(path, anchors, extra=None):
    wb = Workbook()
    ws = wb.active
    ws.title = "图表总览"
    for r, v in enumerate([("年份", "营收"), (2023, 1.0), (2024, 2.0), (2025, 3.0)], start=1):
        ws.cell(r, 1, v[0])
        ws.cell(r, 2, v[1])
    for i, a in enumerate(anchors, start=1):
        ws.add_chart(chart(ws, f"图表{i}"), a)
    if extra:
        extra(wb, ws)
    wb.save(path)
    with open(path, "rb") as f:
        return C.layout_problems(C.parse_workbook(f.read()))


def main():
    tmp = sys.argv[1] if len(sys.argv) > 1 else "/tmp"
    os.makedirs(tmp, exist_ok=True)
    p = lambda n: os.path.join(tmp, n + ".xlsx")

    # 1 并排
    bad = book(p("side"), ["A5", "H5"])
    assert len(bad) == 1, bad
    assert "第 1 张图表“图表1”（第 5–22 行）与第 2 张图表“图表2”（第 5–22 行）行区间相交" in bad[0], bad

    # 2 单列纵向（与 build_history_xlsx 的规则一致）
    an = B.chart_anchors(3)
    assert an == ["A5", "A26", "A47"], an
    assert book(p("column"), an) == []

    # 3 贴边不算相交：360 px 图在默认 20 px 行（openpyxl 写出的默认行高 15 pt）上正好 18 行，A5 底边 = A23 顶边
    assert book(p("touch"), ["A5", "A23"]) == []
    assert book(p("touch_bad"), ["A5", "A22"]) != []

    # 4 twoCellAnchor 与 absoluteAnchor
    def two_cell(wb, ws):
        ch = chart(ws, "两格锚")
        ch.anchor = TwoCellAnchor(_from=AnchorMarker(col=0, row=20), to=AnchorMarker(col=6, row=30))
        ws.add_chart(ch)
        ch2 = chart(ws, "绝对锚")
        ch2.anchor = AbsoluteAnchor(pos=XDRPoint2D(0, 100 * 20 * EMU_PX),
                                    ext=XDRPositiveSize2D(720 * EMU_PX, 360 * EMU_PX))
        ws.add_chart(ch2)
    bad = book(p("twocell"), ["A5"], two_cell)   # A5 占第 5–22 行，两格锚占第 21–30 行
    assert len(bad) == 1 and "图表“图表1”（第 5–22 行）" in bad[0] and "图表“两格锚”（第 21–30 行）" in bad[0] \
        and "绝对锚" not in bad[0], bad

    # 5 多余空行：1000 行（Google 默认）不通过；内容 + 50 行以内通过；只有样式的空单元格不算内容
    def rows_1000(wb, ws):
        s = wb.create_sheet("来源")
        s["A1"] = "来源 ID"
        for r in range(1, 1001):
            s.row_dimensions[r].height = 15.75
        s["C900"].fill = PatternFill("solid", fgColor="F7F9FB")   # 只有样式
    bad = book(p("rows1000"), B.chart_anchors(2), rows_1000)
    assert len(bad) == 1 and "“来源”：1000 行，内容与图只到第 1 行" in bad[0], bad

    def rows_ok(wb, ws):
        s = wb.create_sheet("来源")
        s["A1"] = "来源 ID"
        for r in range(1, 52):
            s.row_dimensions[r].height = 15.75
        for r in range(1, ws.max_row + 60):   # 总览：图 2 底边第 43 行，+ 50 = 93 行以内
            ws.row_dimensions[r].height = 15
    assert book(p("rows_ok"), B.chart_anchors(2), rows_ok) == []

    def rows_chart(wb, ws):   # 总览只有图、没有更多内容：按图底边算
        for r in range(1, 43 + 51 + 1):
            ws.row_dimensions[r].height = 15
    bad = book(p("rows_chart"), B.chart_anchors(2), rows_chart)
    assert len(bad) == 1 and "“图表总览”：94 行，内容与图只到第 43 行" in bad[0], bad

    # 6 build_history_xlsx 的常量与裁剪目标
    assert B.CHART_GAP_ROWS == 21 and B.GSHEET_ROW_PX == 21
    assert abs(B.CHART_W_CM - 19.05) < 1e-9 and abs(B.CHART_H_CM - 9.525) < 1e-9
    wb = Workbook()
    ws = wb.active
    ws["D2"] = 1
    for col, px in (("A", 260), ("B", 100), ("C", 100), ("D", 100)):
        ws.column_dimensions[col].width = B.px_to_width(px)
    for a in B.chart_anchors(2):
        ws.add_chart(chart(ws), a)
    rows, cols = B.trim_target(ws)
    assert (rows, cols) == (26 + 18 - 1 + 5, 6 + 2), (rows, cols)   # 图 2 底边第 43 行；图右边落在 F 列
    print("图表版面：并排相交报出两张图与行区间；单列纵向、贴边不报；twoCellAnchor / absoluteAnchor 可读；"
          "1000 空行不通过、内容 + 50 以内通过、样式空格不算内容；build_history_xlsx 锚点 A5/A26/A47、"
          "行距 21、720×360 px（19.05×9.525 cm）、裁剪目标 48 行×8 列")


if __name__ == "__main__":
    main()
