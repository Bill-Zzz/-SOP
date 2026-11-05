#!/usr/bin/env python3
"""生成"转换实测"用的两份小样，用来上传到 Google Drive、转换为 Google 文档 / 表格后核对格式保留了什么。

用法：python3 tests/make_format_sample.py <输出目录> [--date 2026-11-04]
产出（文件名按存储规范）：
  ACME_格式小样报告_YYYYMMDD.docx  tests/fixtures/md2docx_sample.md 经 md2docx.py 转换，图片换成图表总览第 1 图的 PNG 快照
                                  （缩小、减色到十几 KB）。含页眉页脚、目录、单位表头与合计行的表格、图片、注释、附录。
  ACME_格式小样模型_YYYYMMDD.xlsx  用 build_history_xlsx.py 的格式函数（format_sheet、finish_chart 等）搭的四张小表 + 一张图：
                                  主标题、分区标题、列标题、隔行、合计行、蓝 / 黑 / 绿三种字体颜色、金额 / 百分比 / 倍数 / 股价 /
                                  年度 2024A 数字格式、负数与零值、长文本列、组合图（柱 + 副轴折线，文本横轴）。
两份都压到 9 KB 上下：Google Drive MCP 的 create_file 只收 base64 内容，要经 AI 自己的输出转写，
实测 1.2 万字符能成功，2.1 万字符转写出错（丢了约 4 KB）。
数据用 tests/make_history_fixture.py 的虚构公司 ACME。
"""
import argparse
import os
import shutil
import subprocess
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, os.path.join(ROOT, "research-sop", "scripts"))
import build_history_xlsx as B  # noqa: E402
import md2docx  # noqa: E402
import openpyxl  # noqa: E402
from openpyxl.chart import BarChart, LineChart, Reference  # noqa: E402


# 上传用的精简报告：覆盖 Part A 全部可程序化的项目（页眉页脚、摘要、目录、一二三级标题、单位表头、合计行、表题、
# 图片与图注、注释、附录），篇幅压到最小——MCP 上传按 base64 走，2 万字符以上实测会转写出错。
SAMPLE_MD = """# ACME 公司档案（格式小样）

## 一页摘要

ACME 在 2019 年从硬件转向订阅，营收从 40 亿美元增长到 52.9 亿美元 [S01 p3]。这一段故意写长一些，让它在版心里折成两到三行，用来量正文行距：正文 10.5 pt、行距 1.0 倍、段后 6 pt，转换为 Google 文档后应保持不变。

## 一、财务全史

### 1.1　全史总表

表 1 全史总表：营收与毛利率

| 年度 | 营收（亿美元） | 毛利率（%） | 当年大事 |
|---|---|---|---|
| 2020 | 46.0 | 60.0 | 订阅收入首次超过一半 |
| 2021 | 52.9 | 60.0 | 收购一家数据公司，交易价值 3.1 亿美元 |
| 合计 | 98.9 | — | — |

表注：毛利率 = 毛利 ÷ 营收。

![图 1 营收与同比](fig/图01_营收与增速.png)

数据与可编辑图表：ACME_格式小样模型_20261104 图表总览第 1 图

#### 1.1.1　口径

口径说明：财年即日历年。

## 二、未来：利好与利空

最大的问题是订阅续约率能否维持。

## 附一、本报告的取舍

只用来检查附录另起一页。
"""


def small_png(src, dst, width=260, colors=4):
    from PIL import Image
    im = Image.open(src).convert("RGB")
    h = round(im.height * width / im.width)
    im = im.resize((width, h), Image.LANCZOS).quantize(colors=colors, method=Image.Quantize.MEDIANCUT)
    im.save(dst, optimize=True, dpi=(150, 150))


def make_docx(out_dir, date, tmp):
    futu = os.path.join(tmp, "futu")
    subprocess.run([sys.executable, os.path.join(HERE, "make_history_fixture.py"), futu], check=True,
                   capture_output=True)
    B.main(["ACME", "--root", futu, "--out", os.path.join(tmp, "x"), "--png", os.path.join(tmp, "png")])
    md_dir = os.path.join(tmp, "md")
    os.makedirs(os.path.join(md_dir, "fig"), exist_ok=True)
    with open(os.path.join(md_dir, "sample.md"), "w", encoding="utf-8") as f:
        f.write(SAMPLE_MD)
    small_png(os.path.join(tmp, "png", "图01_营收与增速.png"), os.path.join(md_dir, "fig", "图01_营收与增速.png"))
    out = os.path.join(out_dir, f"ACME_格式小样报告_{date.replace('-', '')}.docx")
    rc = md2docx.main([os.path.join(md_dir, "sample.md"), "--out", out, "--date", date])
    assert rc == 0, "md2docx 格式检查未通过"
    return out, futu


def strip_parts(path):
    """上传版去掉主题与文档属性（与单元格、图表格式无关：字体与颜色都显式写在样式里），只为缩小 base64。"""
    import re
    import zipfile
    drop = {"xl/theme/theme1.xml", "docProps/app.xml", "docProps/core.xml"}
    with zipfile.ZipFile(path) as z:
        items = [(i.filename, z.read(i.filename)) for i in z.infolist() if i.filename not in drop]
    fix = {"[Content_Types].xml": r'<Override PartName="/(xl/theme/theme1|docProps/app|docProps/core)\.xml"[^>]*/>',
           "_rels/.rels": r'<Relationship [^>]*Target="/?docProps/(app|core)\.xml"[^>]*/>',
           "xl/_rels/workbook.xml.rels": r'<Relationship [^>]*Target="/?(xl/)?theme/theme1\.xml"[^>]*/>'}
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as z:
        for name, data in items:
            if name in fix:
                data = re.sub(fix[name], "", data.decode("utf-8")).encode("utf-8")
            z.writestr(name, data)


def make_xlsx(out_dir, date, futu):
    ex = B.Export("ACME", futu)
    years = ex.years[-3:]
    g = lambda key, y: ex.core(key, ex.aper[y])[0]
    wb = openpyxl.Workbook()
    B.set_default_font(wb)
    wb.remove(wb.active)
    ov, t3, mv, calc = (wb.create_sheet(n) for n in ("图表总览", "年度三表", "市值与估值", "图表计算"))
    col = {y: openpyxl.utils.get_column_letter(5 + i) for i, y in enumerate(years)}

    # 年度三表：常数（蓝）、本表公式（黑）、分区标题、合计行、负数括号、零值短横线
    B.title(t3, "ACME 年度三表（格式小样）", "金额：亿美元；每股收益：美元/股；空白 = 无数，不等于 0。")
    B.hdr(t3, 4, ["区块", "原始科目名", "标准化中文名", "单位"] + years)
    for y in years:
        t3[f"{col[y]}4"].number_format = B.F_YEAR
    B.section(t3, "A5", "核心科目（标准化）")
    rows = [("rev", "营收", "利润表·营业总收入", "亿美元"), ("cogs", "营业成本", "利润表·营业总成本", "亿美元"),
            ("gp", "毛利", "营收 − 营业成本（本表公式）", "亿美元"), ("ni", "净利润（归母）", "利润表·归母净利润", "亿美元"),
            ("ocf", "经营现金流", "现金流量表", "亿美元"), ("capex", "资本开支", "现金流量表（负数）", "亿美元"),
            ("div", "分红", "现金流量表（无，记 0）", "亿美元"), ("eps_b", "基本每股收益", "利润表", "美元/股")]
    r = 6
    rr = {}
    for key, name, src, unit in rows:
        t3[f"A{r}"], t3[f"B{r}"], t3[f"C{r}"], t3[f"D{r}"] = "核心", src, name, unit
        for y in years:
            c = t3[f"{col[y]}{r}"]
            if key == "gp":
                c.value = f"={col[y]}{rr['rev']}-{col[y]}{rr['cogs']}"
            elif key == "div":
                c.value = 0
            else:
                v = g(key, y)
                c.value = v if key == "eps_b" else v / 1e8
            c.number_format = B.F_EPS if key == "eps_b" else B.F_AMT
        rr[key] = r
        r += 1
    r += 1
    B.section(t3, f"A{r}", "分部营收（合计行加粗、加上边框）")
    r += 1
    seg0 = r
    for k, (name, w) in enumerate((("分部甲", 0.6), ("分部乙", 0.4))):
        t3[f"A{r}"], t3[f"C{r}"], t3[f"D{r}"] = "分部", name, "亿美元"
        for y in years:
            t3[f"{col[y]}{r}"] = round(g("rev", y) / 1e8 * w, 2)
            t3[f"{col[y]}{r}"].number_format = B.F_AMT
        r += 1
    t3[f"A{r}"], t3[f"C{r}"], t3[f"D{r}"] = "合计", "分部合计", "亿美元"
    for y in years:
        t3[f"{col[y]}{r}"] = f"=SUM({col[y]}{seg0}:{col[y]}{r - 1})"
        t3[f"{col[y]}{r}"].number_format = B.F_AMT

    # 市值与估值：日期与股价常数（蓝）、跨表引用（绿）、本表公式（黑）、倍数、百分比、长文本列
    B.title(mv, "ACME 市值与估值（格式小样）", "市值 = 收盘价 × 股数（亿股），亿美元；股数口径列是长文本列（120 px，换行）。")
    heads = ["年度", "财年期末日", "收盘价（美元）", "股数（亿股）", "股数口径", "市值（亿美元）", "营收（亿美元）", "P/S",
             "FCF 收益率"]
    B.hdr(mv, 4, heads)
    import datetime as dt
    for i, y in enumerate(years):
        r = 5 + i
        mv[f"A{r}"] = y
        mv[f"A{r}"].number_format = B.F_YEAR
        mv[f"B{r}"] = dt.date.fromisoformat(ex.meta[ex.aper[y]]["end"])
        mv[f"B{r}"].number_format = B.F_DATE
        mv[f"C{r}"] = round(20 * 1.15 ** i, 2)
        mv[f"C{r}"].number_format = B.F_PX
        mv[f"D{r}"] = f"='年度三表'!{col[y]}{rr['ni']}/'年度三表'!{col[y]}{rr['eps_b']}"
        mv[f"D{r}"].number_format = B.F_SH
        mv[f"E{r}"] = "基本加权股数（推算）" if y < years[-1] else "期末股数（资产负债表·普通股股数）"
        mv[f"F{r}"] = f"=C{r}*D{r}"
        mv[f"G{r}"] = f"='年度三表'!{col[y]}{rr['rev']}"
        mv[f"H{r}"] = f"=F{r}/G{r}"
        mv[f"I{r}"] = f"=('年度三表'!{col[y]}{rr['ocf']}+'年度三表'!{col[y]}{rr['capex']})/F{r}-0.05"
        for k, fmt in (("F", B.F_AMT), ("G", B.F_AMT), ("H", B.F_X), ("I", B.F_PCT)):
            mv[f"{k}{r}"].number_format = fmt

    # 图表计算：文本横轴标签列、跨表引用
    B.title(calc, "图表计算（格式小样）", "全部是公式；年度标签列是图表的文本横轴。")
    B.section(calc, "A3", "年度块")
    B.hdr(calc, 4, ["年度", "营收", "营收同比", "年度标签"])
    for i, y in enumerate(years):
        r = 5 + i
        calc[f"A{r}"] = f"='图表总览'!$B$2+{i}"
        calc[f"A{r}"].number_format = B.F_YEAR
        calc[f"B{r}"] = f"='年度三表'!{col[y]}{rr['rev']}"
        calc[f"C{r}"] = f"=IF(ROW()=5,NA(),B{r}/B{r - 1}-1)" if i else "=NA()"
        calc[f"D{r}"] = f'=A{r}&"A"'
        calc[f"B{r}"].number_format = B.F_AMT
        calc[f"C{r}"].number_format = B.F_PCT
    rN = 4 + len(years)

    # 图表总览：输入（蓝）、辅助说明（灰）、组合图 + 三系列折线
    B.title(ov, "ACME 图表总览（格式小样）")
    ov["A2"], ov["B2"] = "起始年份", years[0]
    ov["B2"].number_format = B.F_INT
    B.mark_note(ov, "A3", "组合图（柱 + 副轴折线）；横轴是文本标签 2021A……，数据区在'图表计算'。")
    cats = Reference(calc, min_col=4, min_row=5, max_row=rN)
    bar = BarChart()
    bar.type = "col"
    bar.add_data(Reference(calc, min_col=2, min_row=4, max_row=rN), titles_from_data=True)
    B.set_text_categories(bar, cats)
    ln = LineChart()
    ln.add_data(Reference(calc, min_col=3, min_row=4, max_row=rN), titles_from_data=True)
    B.set_text_categories(ln, cats)
    ln.y_axis.axId = 200
    ln.y_axis.crosses = "max"
    ln.y_axis.number_format = "0%"
    ln.y_axis.delete = False
    bar += ln
    B.finish_chart(bar, "1 营收（亿美元）与同比", "亿美元")
    ov.add_chart(bar, "A5")
    for ws in wb.worksheets:
        B.format_sheet(ws)
    wb.calculation.fullCalcOnLoad = True
    out = os.path.join(out_dir, f"ACME_格式小样模型_{date.replace('-', '')}.xlsx")
    wb.save(out)
    strip_parts(out)
    res = B.format_checks(out, openpyxl.load_workbook(out), B.Evaluator(openpyxl.load_workbook(out)))
    for name, ok, detail in res:
        print(f"[{'通过' if ok else '不通过'}] {name}\n        {detail}")
    assert all(ok for _, ok, _ in res), "小样格式检查未通过"
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("out")
    ap.add_argument("--date", default="2026-11-04")
    a = ap.parse_args()
    os.makedirs(a.out, exist_ok=True)
    with tempfile.TemporaryDirectory() as tmp:
        d, futu = make_docx(a.out, a.date, tmp)
        x = make_xlsx(a.out, a.date, futu)
    for p in (d, x):
        print(f"{p}：{os.path.getsize(p) / 1024:.1f} KB")


if __name__ == "__main__":
    main()
