#!/usr/bin/env bash
# 自检：
#   1) CRWV 复现引擎与 Excel 对账
#   2) 研报库索引 / 检索 / 阅读准备（含日期误判等回归测试）
#   3) xlsx 修改流程：保留图表、公式、隐藏表；强制重算标记
#   4) LibreOffice 重算（未安装时明确报告为"跳过"）
#   5) 历史数据 Excel 脚本：用虚构公司 ACME 的最小长表生成工作簿并跑 --check（含格式规范 Part B 第 9–12 项）；
#      --png 导出图表快照
#   6) md2docx.py：示例 .md（摘要、目录、编号问题、单位表头、合计行、现场生成的占位 PNG、注释、附录）转 .docx，
#      检查输出（含标题编号与格式规范 Part A 格式检查）并读回核对
# 用法：bash tests/run_tests.sh   （在仓库根目录运行）
set -euo pipefail
cd "$(dirname "$0")/.."
TMP="$(mktemp -d)"; trap 'rm -rf "$TMP"' EXIT
SKIPPED=()

echo "== 0. 可用工具 =="
python3 scripts/pdf_text.py check

echo "== 1. 复现引擎对账 =="
python3 examples/crwv/engine_crwv.py check | tee "$TMP/check.txt"
grep -q "复现通过" "$TMP/check.txt"

echo "== 2. 研报库索引与检索 =="
python3 - <<'EOF' || { echo "（无可用的 PDF 文本后端，请 pip install pypdf）"; exit 1; }
import sys; sys.path.insert(0, "scripts"); import pdf_text
b = pdf_text.backends(); assert b["pymupdf"] or b["pdftotext"] or b["pypdf"]
EOF
python3 tests/make_library.py "$TMP/lib" >/dev/null
python3 scripts/library_index.py "$TMP/lib" --db "$TMP/idx/library.db" >/dev/null
python3 scripts/library_search.py --db "$TMP/idx/library.db" --ticker ACME --name "Acme Compute" --kw "residual value" \
  --out "$TMP/manifest.json" > "$TMP/search.txt"
cat "$TMP/search.txt"
python3 - "$TMP/idx/library.db" "$TMP/manifest.json" <<'EOF'
import json, sqlite3, sys
con = sqlite3.connect(sys.argv[1])
rows = {r[0].split("/")[-1]: r[1:] for r in con.execute(
    "select rel, broker, date, tickers, rating, target, title from docs")}
gs = rows["6fd4a26e-GS_note-261021.pdf"]
assert gs[0] == "高盛", gs
assert gs[1] == "2026-10-21", f"日期误判（回归：文件夹年月被当成日期）：{gs[1]}"
assert "ACME" in gs[2] and gs[3] == "买入" and gs[4] == 106.0, gs
assert "funding gap" in gs[5], f"标题识别错误（回归：哈希文件名 / 目标价行被当成标题）：{gs[5]}"
bern = rows["acme_quick_take.pdf"]
assert bern[0] == "伯恩斯坦" and bern[1] == "2026-10-15" and bern[3] == "跑输" and bern[4] == 74.0, bern
hsbc = rows["2026-09-30_HSBC_sector.pdf"]
assert hsbc[0] == "汇丰" and hsbc[1] == "2026-09-30", hsbc
jpm = rows["other_company.pdf"]
assert jpm[0] == "摩根大通" and jpm[3] == "中性" and jpm[4] == 50.0, jpm
m = json.load(open(sys.argv[2], encoding="utf8"))
names = [d["path"].split("/")[-1] for d in m["docs"]]
assert "other_company.pdf" not in names, "无关公司不应入选"
assert names[0] in ("6fd4a26e-GS_note-261021.pdf", "acme_quick_take.pdf"), f"主角报告应排在前面：{names}"
assert names[-1] == "2026-09-30_HSBC_sector.pdf", f"顺带提及的行业报告应排在最后：{names}"
print("索引与检索：元数据识别、排序、过滤均正确")
EOF
python3 scripts/library_index.py "$TMP/lib" --db "$TMP/idx/library.db" | grep -q '"skip": 4' && echo "增量索引：未改动文件全部跳过"
python3 scripts/prepare_reading.py "$TMP/manifest.json" "$TMP/reading" >/dev/null
test -f "$TMP/reading/reading_pack.md" && test -f "$TMP/reading/sources.csv"
grep -q "===== 第 1 页 =====" "$TMP"/reading/text/S01_*.txt
grep -q "<!-- CARDS -->" templates/reading_pack.md
echo "阅读准备：全文（含页码标记）、阅读包骨架、来源清单已生成"

echo "== 3. xlsx 修改流程 =="
python3 tests/make_fixture.py "$TMP/sample.xlsx" >/dev/null
cat > "$TMP/edits.json" <<'EOF'
{"假设": {"B1": 5.5, "H1": {"text": "新出处说明：机构 日期 附表 1"}}}
EOF
python3 scripts/xlsx_patch.py "$TMP/sample.xlsx" "$TMP/patched.xlsx" "$TMP/edits.json" --full-calc-on-load
python3 - "$TMP/sample.xlsx" "$TMP/patched.xlsx" <<'EOF'
import sys, zipfile, openpyxl
a, b = sys.argv[1:]
charts = lambda p: [n for n in zipfile.ZipFile(p).namelist() if "/charts/" in n]
assert charts(a) == charts(b) and charts(b), "图表丢失"
wb = openpyxl.load_workbook(b)
assert wb["假设"]["B1"].value == 5.5 and wb["假设"]["H1"].value.startswith("新出处")
assert wb["假设"]["B3"].value == "=B2-B1", "公式被改动"
assert wb["隐藏计算"].sheet_state == "hidden"
assert 'fullCalcOnLoad="1"' in zipfile.ZipFile(b).read("xl/workbook.xml").decode(), "未写入强制重算标记"
print("修改后：图表、公式、隐藏表均保留，已标记打开时强制重算")
EOF

echo "== 4. LibreOffice 重算 =="
if command -v soffice >/dev/null || command -v libreoffice >/dev/null; then
  bash scripts/lo_recalc.sh "$TMP/patched.xlsx" "$TMP/out" >/dev/null
  python3 - "$TMP/out/patched.xlsx" <<'EOF'
import sys, zipfile, openpyxl
p = sys.argv[1]
v = openpyxl.load_workbook(p, data_only=True)
assert abs(v["假设"]["B3"].value - 29.5) < 1e-9, v["假设"]["B3"].value
assert abs(v["隐藏计算"]["A1"].value - 59.0) < 1e-9
assert [n for n in zipfile.ZipFile(p).namelist() if "/charts/" in n], "重算后图表丢失"
print("重算后：净债务 29.5、隐藏表 59.0，图表保留")
EOF
else
  SKIPPED+=("LibreOffice 重算测试（未安装 LibreOffice）")
fi

echo "== 5. 历史数据 Excel 脚本 =="
python3 tests/make_history_fixture.py "$TMP/futu" >/dev/null
python3 research-sop/scripts/build_history_xlsx.py ACME --root "$TMP/futu" --out "$TMP/hist" --ai 测试 --check \
  | tee "$TMP/hist.txt"
grep -q "结果：12/12 通过" "$TMP/hist.txt"
grep -q "股数口径：基本加权股数 5 年、期末股数 1 年" "$TMP/hist.txt"
python3 - "$TMP/hist" <<'EOF'
import glob, sys, zipfile, openpyxl
sys.path.insert(0, "research-sop/scripts")
p = glob.glob(sys.argv[1] + "/*_ACME_历史数据_v1_测试.xlsx")
assert len(p) == 1, p
wb = openpyxl.load_workbook(p[0])
for s in ["图表总览", "说明", "历年营收与业务拆分", "年度三表", "年度衍生指标", "近12季度明细", "市值与营收对比",
          "市值与估值", "股价事件", "IPO指标", "所得税与现金税", "重大融资时间线", "重要收购", "指引与实际",
          "缺口", "来源", "图表计算"]:
    assert s in wb.sheetnames, f"缺工作表 {s}"
q = [c.value for c in wb["近12季度明细"][3][2:14]]
assert q[-1] == "2Q25A" and q[0] == "3Q22A", q
assert str(wb["年度衍生指标"]["E5"].value).startswith("=IF(COUNT("), "衍生指标应为公式"
ev = [r for r in wb["股价事件"].iter_rows(min_row=5, values_only=True) if r[1] == "单日"]
assert len(ev) == 1 and str(ev[0][0])[:10] == "2023-03-15", ev
assert len([n for n in zipfile.ZipFile(p[0]).namelist() if n.startswith("xl/charts/chart")]) == 13
vs = wb["市值与估值"]
h = {c.value: c.column for c in vs[4] if c.value}
row = {vs.cell(r, 1).value: r for r in range(5, vs.max_row + 1) if isinstance(vs.cell(r, 1).value, int)}
basis = {y: vs.cell(r, h["股数口径"]).value for y, r in row.items()}
assert basis[2024] == "期末股数（资产负债表·普通股股数）", basis
assert basis[2023] == basis[2022] == "基本加权股数（推算）", basis
assert abs(vs.cell(row[2024], h["股数（亿股）"]).value - 10.2) < 1e-9, "股数单位应为亿股"
assert "未采用" in str(vs.cell(row[2023], h["备注"]).value), "量级不对的期末股数应注明未采用"
mc = vs.cell(row[2024], h["市值（亿美元）"]).value
assert mc.endswith("*" + vs.cell(row[2024], h["股数（亿股）"]).coordinate + ',""))'), f"市值 = 价 × 亿股，不再 ÷ 100：{mc}"
# 格式规范 Part B（抽查；--check 第 9–12 项逐格比过）
B = __import__("build_history_xlsx")
nf = wb._named_styles["Normal"].font
assert (nf.name, nf.sz) == ("Arial", 10), nf
for ws in wb.worksheets:
    assert ws.sheet_view.showGridLines is False and ws.freeze_panes == "B2", ws.title
    assert ws.sheet_properties.tabColor.rgb.endswith("17365D"), ws.title
    assert B.width_to_px(ws.column_dimensions["A"].width) == 260, ws.title
    assert round(ws.row_dimensions[1].height / 0.75) == 28, ws.title
    assert ws["A1"].fill.fgColor.rgb.endswith("17365D") and ws["A1"].font.color.rgb.endswith("FFFFFF"), ws.title
    assert not [r for r in ws.merged_cells.ranges if r.min_row != 1], ws.title
hc = vs.cell(4, h["年度"])
assert hc.fill.fgColor.rgb.endswith("1F4E78") and hc.font.color.rgb.endswith("FFFFFF") and hc.font.bold
r = row[2024]
assert vs.cell(r, h["财年期末日"]).font.color.rgb.endswith("0000FF"), "手输常数应为蓝色"
assert vs.cell(r, h["市值（亿美元）"]).font.color.rgb.endswith("000000"), "本表公式应为黑色"
assert vs.cell(r, h["总债务（亿美元）"]).font.color.rgb.endswith("008000"), "跨表引用应为绿色"
assert vs.cell(r, 1).number_format == '0"A"' and vs.cell(r, 1).value == 2024, "年度显示 2024A、数值不变"
assert vs.cell(r, h["市值（亿美元）"]).number_format == '#,##0.0;(#,##0.0);"-"'
assert vs.cell(r, h["收盘价（美元）"]).number_format == '$#,##0.00;($#,##0.00);"-"'
assert vs.cell(r, h["P/E"]).number_format == '0.0"x";(0.0"x");"-"'
assert wb["年度三表"]["A5"].fill.fgColor.rgb.endswith("D9EAF7"), "分区标题底色"
band = {vs.cell(rr, 2).fill.fgColor.rgb[-6:] for rr in row.values() if vs.cell(rr, 2).fill.fill_type == "solid"}
assert band == {"F7F9FB"}, band
xml = zipfile.ZipFile(p[0]).read("xl/charts/chart1.xml").decode()
assert "17365D" in xml and "5B9BD5" in xml and "D9D9D9" in xml and 'sz="1000"' in xml and "strRef" in xml
print("历史数据 Excel：17 张表、13 张图、近 12 季（2Q25A）、股价事件、公式、亿股与市值、Part B 格式抽查均正确")
EOF
if python3 -c "import matplotlib" 2>/dev/null; then
  python3 research-sop/scripts/build_history_xlsx.py ACME --root "$TMP/futu" --check-only "$TMP"/hist/*_ACME_历史数据_v1_测试.xlsx \
    --png "$TMP/png" | tee "$TMP/png.txt"
  grep -q "结果：13/13 通过" "$TMP/png.txt"
  grep -q "跳过：图 2 无数据" "$TMP/png.txt"
  python3 - "$TMP/png" <<'EOF'
import os, sys
fs = sorted(os.listdir(sys.argv[1]))
assert len(fs) == 10, fs
assert fs[0] == "图01_营收与增速.png" and "图11_股价与事件.png" in fs, fs
assert not any(f.startswith(("图02_", "图12_", "图13_")) for f in fs), "第二层未填的图应跳过"
assert all(os.path.getsize(os.path.join(sys.argv[1], f)) > 5000 for f in fs)
print("--png：10 张有数据的图已导出，第二层未填的 3 张跳过，编号与图表总览一致")
EOF
else
  SKIPPED+=("--png 导出测试（未安装 matplotlib）")
fi

echo "== 6. md2docx.py =="
python3 -c "import docx" 2>/dev/null || pip install -q python-docx --break-system-packages
mkdir -p "$TMP/md/fig"
cp tests/fixtures/md2docx_sample.md "$TMP/md/"
python3 -c "from PIL import Image; Image.new('RGB', (1600, 800), (78, 121, 167)).save('$TMP/md/fig/图01_营收与增速.png', dpi=(200, 200))"
python3 research-sop/scripts/md2docx.py "$TMP/md/md2docx_sample.md" --out "$TMP/md/out.docx" --date 2026-11-04 \
  | tee "$TMP/md.txt"
grep -q "标题 9（报告标题 1 / 一级 5 / 二级 2 / 三级 1）；表格 2；图片 1" "$TMP/md.txt"
grep -q "页眉：ACME 公司档案（测试样例） ⇥ ACME · 2026-11-04" "$TMP/md.txt"
grep -q "标题编号（只检查不改写）：正文一级 2、二级 2、三级 1；问题 2 处" "$TMP/md.txt"
grep -q "「1.2 同行比较」：编号与标题之间缺全角空格" "$TMP/md.txt"
grep -q "「三、未来：利好与利空」：跳号或错号，应为「二、」" "$TMP/md.txt"
grep -q "格式检查（格式规范 Part A，读回 .docx 逐项比）：10/10 通过" "$TMP/md.txt"
if python3 research-sop/scripts/md2docx.py "$TMP/md/md2docx_sample.md" --check-only --img-root "$TMP/nowhere" \
    >/dev/null 2>"$TMP/md_err.txt"; then
  echo "缺图时应以非零码退出"; exit 1
fi
grep -q "fig/图01_营收与增速.png" "$TMP/md_err.txt"
python3 - "$TMP/md/out.docx" <<'EOF'
import sys
import docx
from docx.oxml.ns import qn
import os
d = docx.Document(sys.argv[1])
assert os.path.getsize(sys.argv[1]) < 30 * 1024, "瘦身后应在 30 KB 以内（MCP 上传按 base64 走）"
st = lambda name: [p for p in d.paragraphs if p.style.name == name]
assert [p.text for p in st("Title")] == ["ACME 公司档案（测试样例）"]
assert [p.text for p in st("独立标题")] == ["一页摘要", "目录"]
assert len(st("Heading 1")) == 4 and len(st("Heading 2")) == 2 and len(st("Heading 3")) == 1
assert [p.text for p in d.paragraphs if p.paragraph_format.page_break_before] == \
    ["目录", "一、财务全史", "三、未来：利好与利空", "附一、本报告的取舍", "附二、来源"]
sec = d.sections[0]
assert (sec.page_width.twips, sec.page_height.twips) == (11906, 16838) and sec.left_margin.twips == 1134
assert "ACME · 2026-11-04" in sec.header.paragraphs[0].text and "NUMPAGES" in sec.footer._element.xml
assert 'TOC \\o "1-2"' in d.element.body.xml
assert len(d.tables) == 2 and d.tables[0].style.name == "Table Grid"
t = d.tables[0]
assert len(t.rows) == 5 and len(t.columns) == 5
hc = t.rows[0].cells[1]
assert hc._tc.tcPr.find(qn("w:shd")).get(qn("w:fill")) == "17365D"
runs = [r for r in hc.paragraphs[0].runs if r.text.strip()]
assert [r.text for r in runs] == ["营收", "（亿美元）"] and runs[1].font.size.pt == 9 and runs[1].bold is False
assert t.rows[1].cells[1].paragraphs[0].alignment == 1 and t.rows[1].cells[3].paragraphs[0].alignment == 0, "数值居中、文本左对齐"
assert all(r.bold for r in t.rows[4].cells[0].paragraphs[0].runs) and t.rows[4].cells[0]._tc.tcPr.find(qn("w:tcBorders")) is not None
assert sum(int(g.get(qn("w:w"))) for g in t._tbl.tblGrid.findall(qn("w:gridCol"))) == 9638
widths = [int(g.get(qn("w:w"))) for g in t._tbl.tblGrid.findall(qn("w:gridCol"))]
assert widths[3] == max(widths), f"叙述列应吃剩余宽度：{widths}"
assert len(d.inline_shapes) == 1 and d.inline_shapes[0].width <= 15 * 360000
normal = d.styles["Normal"]
rf = normal.element.rPr.rFonts
assert rf.get(qn("w:ascii")) == "Arial" and rf.get(qn("w:eastAsia")) == "Noto Sans SC", "西文 Arial、中文 Noto Sans SC 分设"
assert normal.font.size.pt == 10.5 and normal.paragraph_format.line_spacing == 1.25
assert [d.styles[f"Heading {k}"].font.size.pt for k in (1, 2, 3)] == [18, 14, 12]
notes = [p for p in d.paragraphs if p.style.name == "注释小字"]
assert {p.text.split("：")[0] for p in notes} >= {"表注", "数据与可编辑图表", "来源", "口径说明"}, [p.text for p in notes]
assert d.styles["注释小字"].font.size.pt == 9 and str(d.styles["注释小字"].font.color.rgb) == "666666"
assert [p.text for p in st("表题")] == ["表 1 全史总表：营收与毛利率"]
body = "\n".join(p.text for p in d.paragraphs)
assert "[R01234 p5]" in body and "**" not in body
print("md2docx：A4 页面、页眉页脚、摘要→目录→正文分页、标题层级与字号、单位表头、合计行、列宽、注释、表题、缺图报错均正确")
EOF

echo
if [ ${#SKIPPED[@]} -gt 0 ]; then
  echo "已通过，但以下测试被跳过："
  for s in "${SKIPPED[@]}"; do echo "  - $s"; done
else
  echo "全部通过（无跳过）"
fi
