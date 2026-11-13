#!/usr/bin/env bash
# 回归测试：
#   1) CRWV 复现引擎与 Excel 对账
#   2) 研报库索引 / 检索 / 阅读准备（含日期误判等回归测试）
#   3) xlsx 修改流程：保留图表、公式、隐藏表；强制重算标记
#   4) LibreOffice 重算（未安装时明确报告为"跳过"）
#   5) 历史数据 Excel 脚本：用虚构公司 ACME 的最小长表生成工作簿并跑 --check（含格式规范 Part B 第 9–12 项）；
#      --png 导出图表快照
#   6) md2docx.py：合规示例 .md（报告外壳：标题 → 一页摘要 → 目录 → 一、结论与概述 → 各章 → 附：口径与说明；单位表头、
#      合计行、现场生成的占位 PNG、注释）转 .docx，检查输出（含标题编号、报告外壳、出处标注、图下来源行、图号与文件名、
#      格式规范 Part A 格式检查：中文、英文与数字统一 Noto Sans SC、行距 1.0、表格 9.5 pt、目录「　…　」）并读回核对；
#      缺一页摘要报出；不合规样例（卷首元信息、第一章不是结论与概述、正文带 [R23475 p5]、来源行带编号、
#      取舍 / 来源 / 修订说明在正文）逐项报出
#      图表优先（格式规范第 7 节 V01–V05）：md2docx 软门槛提示数字墙、数值大表、没有图的正文章且不影响退出码；规则在规范、
#      审稿清单 ⑩⑪ 里
#   7) check_delivery.py：用 rclone lsjson 样例 JSON（tests/fixtures/check_delivery_*.json）离线核对 Drive 布局、
#      命名、格式，并核对本地 成品/、留痕/复用、复核；第 12 项 Google 文档目录二级条目带页码用导出文本样例
#      （tests/fixtures/toc_good、toc_missing）与人工确认样例（toc_evidence_*.txt）测硬门槛；第 13 项图表工作簿结构
#      （格式规范第 13.1 节）用 tests/make_workbook_fixture.py 现场生成的合规 / 不合规 xlsx 测硬门槛；不联网
#   8) research-prompt：SKILL.md 有 frontmatter name、任务书模板在；任务书模板库指向 Drive <研究库>/90_任务书模板库/、
#      先查 Drive，本地不留模板库副本；SKILL.md 与任务书模板不硬编码文件 ID、不含题型内容；
#      任务书第七节有“要读的角度”，research-sop SKILL.md 有 D1 文献检索、D2 角度清单；
#      题型模板（公司档案、行业、主题）齐，research-sop 无“待迁入”标注（0.20.0 解耦）
#   9) check_numbers.py：正文数字对数据表与台账（tests/make_numbers_fixture.py 现场生成）；check_charts.py 图表防乱
#      （tests/make_charts_fixture.py 现场生成）；research-delivery、工单 A/B、
#      AGENTS.md 指向齐，research-sop 不含格式与上传细节（0.22.0 机械活拆分）
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
assert (nf.name, nf.sz) == ("Noto Sans SC", 10), nf
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
python3 - "$TMP"/hist/*_ACME_历史数据_v1_测试.xlsx <<'EOF'
import sys; sys.path.insert(0, "research-sop/scripts")
import check_delivery
bad, note = check_delivery.workbook_problems(check_delivery.parse_workbook(open(sys.argv[1], "rb").read()))
assert not bad, bad
print("历史数据 Excel 符合格式规范第 13.1 节图表工作簿结构（语义页名、图表总览看板、图表计算、来源页）：" + note)
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
  python3 - "$TMP"/hist/*_ACME_历史数据_v1_测试.xlsx "$TMP/png_t" <<'EOF'
import sys
sys.path.insert(0, "research-sop/scripts")
import matplotlib
matplotlib.use("Agg")
from matplotlib.axes import Axes
seen, orig = [], Axes.set_title
def rec(self, label, *a, **k):
    seen.append(label)
    return orig(self, label, *a, **k)
Axes.set_title = rec
import build_history_xlsx as B
B.export_png(sys.argv[1], sys.argv[2], "ACME")
assert len(seen) == 10, seen
bad = [t for t in seen if t.lstrip().startswith("图") or t.lstrip()[:1].isdigit()]
assert not bad, f"图内标题不应写图号（图 N 只在文档图注里）：{bad}"
print("--png：图内标题不写图号，10 张均只写标题")
EOF
else
  SKIPPED+=("--png 导出测试（未安装 matplotlib）")
fi

echo "== 6. md2docx.py =="
python3 -c "import docx" 2>/dev/null || pip install -q python-docx --break-system-packages
mkdir -p "$TMP/md/fig"
cp tests/fixtures/md2docx_sample.md "$TMP/md/"
python3 -c "from PIL import Image; Image.new('RGB', (1600, 800), (78, 121, 167)).save('$TMP/md/fig/fig1_营收与增速.png', dpi=(200, 200))"
python3 research-sop/scripts/md2docx.py "$TMP/md/md2docx_sample.md" --out "$TMP/md/out.docx" --date 2026-11-04 \
  | tee "$TMP/md.txt"
for s in "标题 11（报告标题 1 / 一级 5 / 二级 4 / 三级 1）；表格 2；图片 1" "页眉：ACME 公司档案（测试样例） ⇥ ACME · 2026-11-04" \
    "标题编号（只检查不改写）：正文一级 3、二级 4、三级 1；问题 0 处" "格式检查（格式规范 Part A，读回 .docx 逐项比）：11/11 通过" \
    "图下来源行（只检查不改写）：带研报编号或页码 0 处" "报告外壳（只检查不改写）：问题 0 处" \
    "出处标注与过程用语（只检查不改写）：0 处" "图号与文件名（只检查不改写）：图片 1 张；问题 0 处"; do
  grep -qF -- "$s" "$TMP/md.txt" || { echo "合规样例缺少预期输出：$s"; exit 1; }
done
python3 - "$TMP/md" <<'EOF'
import sys
d = sys.argv[1]
s = open(d + "/md2docx_sample.md", encoding="utf8").read()
num = s.replace("### 2.2　同行比较", "### 2.2 同行比较").replace("## 三、未来", "## 四、未来")
open(d + "/num_bad.md", "w", encoding="utf8").write(num)
i, j = s.index("### 1.2　概述"), s.index("## 二、财务全史")
open(d + "/no_overview.md", "w", encoding="utf8").write(s[:i] + s[j:])
i, j = s.index("## 一页摘要"), s.index("## 一、结论与概述")
open(d + "/no_summary.md", "w", encoding="utf8").write(s[:i] + s[j:])
EOF
python3 research-sop/scripts/md2docx.py "$TMP/md/num_bad.md" --check-only --img-root "$TMP/md" > "$TMP/num_bad.txt"
grep -q "标题编号（只检查不改写）：正文一级 3、二级 4、三级 1；问题 2 处" "$TMP/num_bad.txt"
grep -q "「2.2 同行比较」：编号与标题之间缺全角空格" "$TMP/num_bad.txt"
grep -q "「四、未来：利好与利空」：跳号或错号，应为「三、」" "$TMP/num_bad.txt"
python3 research-sop/scripts/md2docx.py "$TMP/md/no_overview.md" --check-only --img-root "$TMP/md" > "$TMP/no_ov.txt"
grep -q "报告外壳（只检查不改写）：问题 1 处" "$TMP/no_ov.txt"
grep -qF "'一、结论与概述'下缺'概述'" "$TMP/no_ov.txt"
python3 research-sop/scripts/md2docx.py "$TMP/md/no_summary.md" --check-only --img-root "$TMP/md" > "$TMP/no_sum.txt"
grep -q "报告外壳（只检查不改写）：问题 1 处" "$TMP/no_sum.txt"
grep -qF "标题后第一个一级标题应为'一页摘要'，现为'一、结论与概述'" "$TMP/no_sum.txt"
echo "合规样例：外壳、出处标注、图下来源行、编号、图号均 0 处；编号问题副本报 2 处；缺'概述'、缺一页摘要报出"
python3 - "$TMP/md" <<'EOF'
import sys
from PIL import Image
d = sys.argv[1]
for n in (1, 2, 3):
    Image.new("RGB", (800, 400), (78, 121, 167)).save(f"{d}/fig/fig{n}_样例.png")
def doc(imgs, ref=""):
    body = "".join(f"![图 {a}　结论 {a}](fig/fig{f}_样例.png)\n\n数据与可编辑图表：样例 图表总览第 {f} 图\n\n" for a, f in imgs)
    return (f"# 图号样例\n\n## 一、结论与概述\n\n### 1.1　结论\n\n结论。{ref}\n\n### 1.2　概述\n\n概述。\n\n"
            f"## 二、正文\n\n{body}## 附：口径与说明\n\n说明。\n")
open(d + "/fig_order_bad.md", "w", encoding="utf8").write(doc([(2, 2), (1, 1), (3, 3)]))
open(d + "/fig_order_good.md", "w", encoding="utf8").write(doc([(1, 1), (2, 2), (3, 3)], "见图 3。"))
open(d + "/fig_file_bad.md", "w", encoding="utf8").write(doc([(1, 1), (2, 1)], "见图 4。"))
EOF
python3 research-sop/scripts/md2docx.py "$TMP/md/fig_order_bad.md" --check-only --img-root "$TMP/md" > "$TMP/fig_bad.txt"
grep -q "图号与文件名（只检查不改写）：图片 3 张；问题 2 处" "$TMP/fig_bad.txt"
grep -q "第 1 张图标为'图 2'：图号应按出现顺序编，这里应为图 1" "$TMP/fig_bad.txt"
grep -q "第 2 张图标为'图 1'：图号应按出现顺序编，这里应为图 2" "$TMP/fig_bad.txt"
python3 research-sop/scripts/md2docx.py "$TMP/md/fig_order_good.md" --check-only --img-root "$TMP/md" > "$TMP/fig_good.txt"
grep -q "图号与文件名（只检查不改写）：图片 3 张；问题 0 处" "$TMP/fig_good.txt"
python3 research-sop/scripts/md2docx.py "$TMP/md/fig_file_bad.md" --check-only --img-root "$TMP/md" > "$TMP/fig_file.txt"
grep -q "图注'图 2'与文件名 fig1（fig1_样例.png）不一致" "$TMP/fig_file.txt"
grep -q "正文引用了'图 4'，但全文只有 2 张图" "$TMP/fig_file.txt"
echo "图号与文件名：顺序 2,1,3 报错、1,2,3 通过、图 2 配 fig1 报错、引用超出图总数报错"
cp tests/fixtures/md2docx_bad.md "$TMP/md/"
python3 research-sop/scripts/md2docx.py "$TMP/md/md2docx_bad.md" --check-only --img-root "$TMP/md" > "$TMP/bad.txt"
for s in "报告外壳（只检查不改写）：问题 6 处" \
    "标题与第一个一级标题之间有 1 段（首段「2026-11-04 · 测试 · 版本 V2.0 · 数据」），标题后应直接进'一页摘要'；元信息放'附：口径与说明'" \
    "一页摘要与目录之后的第一章应为'一、结论与概述'，现为'一、财务全史'" \
    "'附一、本报告的取舍'作为节标题出现在正文：取舍进底稿" "'附二、来源'作为节标题出现在正文：来源清单进底稿" \
    "'修订说明'作为节标题出现在正文：修订说明进底稿 复核/修订说明_Vx.md" "最后一节应为'附：口径与说明'，现为'修订说明'" \
    "出处标注与过程用语（只检查不改写）：3 处" "「一页摘要」1 处：[R23475 p5]" "「一、财务全史」2 处：台账、对账" \
    "图下来源行（只检查不改写）：带研报编号或页码 1 处" "fig1_营收与增速.png 的来源行带研报编号或页码：「来源：WSTS，R23475 p5。」"; do
  grep -qF -- "$s" "$TMP/bad.txt" || { echo "不合规样例缺少预期输出：$s"; exit 1; }
done
python3 research-sop/scripts/md2docx.py research-sop/templates/report_shell.md --check-only --img-root "$TMP/nowhere" > "$TMP/shell.txt"
for s in "报告外壳（只检查不改写）：问题 0 处" "出处标注与过程用语（只检查不改写）：0 处" "标题编号（只检查不改写）：正文一级 1、二级 2、三级 0；问题 0 处"; do
  grep -qF -- "$s" "$TMP/shell.txt" || { echo "外壳模板 report_shell.md 缺少预期输出：$s"; exit 1; }
done
echo "报告外壳与出处标注：不合规样例报出卷首元信息、第一章不是结论与概述、取舍 / 来源 / 修订说明在正文、末节不是口径与说明、[R23475 p5] 与过程用语、来源行带编号；外壳模板 0 处"
if python3 research-sop/scripts/md2docx.py "$TMP/md/md2docx_sample.md" --check-only --img-root "$TMP/nowhere" \
    >/dev/null 2>"$TMP/md_err.txt"; then
  echo "缺图时应以非零码退出"; exit 1
fi
grep -q "fig/fig1_营收与增速.png" "$TMP/md_err.txt"
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
assert len(st("Heading 1")) == 4 and len(st("Heading 2")) == 4 and len(st("Heading 3")) == 1
assert [p.text for p in d.paragraphs if p.paragraph_format.page_break_before] == \
    ["目录", "一、结论与概述", "二、财务全史", "三、未来：利好与利空", "附：口径与说明"]
texts = [p.text for p in d.paragraphs]
assert texts.index("ACME 公司档案（测试样例）") < texts.index("一页摘要") < texts.index("目录") < texts.index("一、结论与概述"), \
    "标题→一页摘要→目录→第一章"
sec = d.sections[0]
assert (sec.page_width.twips, sec.page_height.twips) == (11906, 16838) and sec.left_margin.twips == 1134
assert "ACME · 2026-11-04" in sec.header.paragraphs[0].text and "NUMPAGES" in sec.footer._element.xml
assert 'TOC \\o "1-2"' in d.element.body.xml and '\\n 1-1 \\p "\u3000\u2026\u3000"' in d.element.body.xml, "目录：一级无页码、分隔符「　…　」"
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
assert rf.get(qn("w:ascii")) == "Noto Sans SC" and rf.get(qn("w:eastAsia")) == "Noto Sans SC", "中文、英文与数字统一 Noto Sans SC（0.17.0）"
assert normal.font.size.pt == 10.5 and normal.paragraph_format.line_spacing == 1.0
assert d.styles["表格表头"].font.size.pt == 9.5 and d.styles["表格正文"].font.size.pt == 9.5, "表格 9.5 pt"
assert d.styles["toc 2"].paragraph_format.line_spacing == 1.25, "目录行距 1.25"
assert [d.styles[f"Heading {k}"].font.size.pt for k in (1, 2, 3)] == [18, 14, 12]
notes = [p for p in d.paragraphs if p.style.name == "注释小字"]
assert {p.text.split("：")[0] for p in notes} >= {"表注", "数据与可编辑图表", "来源", "口径说明"}, [p.text for p in notes]
assert d.styles["注释小字"].font.size.pt == 9 and str(d.styles["注释小字"].font.color.rgb) == "666666"
assert [p.text for p in st("表题")] == ["表 1 全史总表：营收与毛利率"]
body = "\n".join(p.text for p in d.paragraphs)
assert "**" not in body and "[R" not in body
print("md2docx：A4 页面、页眉页脚、标题→一页摘要→目录→一、结论与概述→各章→附：口径与说明分页、字体对换与行距、表格 9.5 pt、目录分隔符、标题层级与字号、单位表头、合计行、列宽、注释、表题、缺图报错均正确")
EOF

# 图表优先（格式规范第 7 节 V01–V05）：md2docx 软门槛提示 + 规则落在正式文件里
cat > "$TMP/chartfirst.md" <<'EOF2'
# 某周期

## 一页摘要

结论一段。

## 一、结论与概述

### 1.1　结论

判断。

### 1.2　概述

概述。

## 二、历史

1985 年 12 月到 1986 年 6 月跌 17%，1989 年 3 月到 1990 年 7 月跌 21%，1996 年 1 月到 1998 年 6 月跌 25%，2000 年 9 月到 2001 年 12 月跌 45%。

| 年份 | 收入（亿美元） | 资本开支（亿美元） | 库存（天） | 价格（%） |
|---|---|---|---|---|
| 2016 | 1.0 | 2.0 | 3.0 | 4.0 |
| 2017 | 1.1 | 2.1 | 3.1 | 4.1 |
| 2018 | 1.2 | 2.2 | 3.2 | 4.2 |
| 2019 | 1.3 | 2.3 | 3.3 | 4.3 |
| 2020 | 1.4 | 2.4 | 3.4 | 4.4 |
| 2021 | 1.5 | 2.5 | 3.5 | 4.5 |
| 2022 | 1.6 | 2.6 | 3.6 | 4.6 |
| 2023 | 1.7 | 2.7 | 3.7 | 4.7 |
| 2024 | 1.8 | 2.8 | 3.8 | 4.8 |

## 附：口径与说明

口径。
EOF2
python3 research-sop/scripts/md2docx.py "$TMP/chartfirst.md" --check-only > "$TMP/cf.txt"   # 软门槛：有提示也退出码 0
grep -q "图表优先（只提示，不影响退出码；格式规范第 7 节 V02–V04）：数字密集段落 1 处、数值大表 1 张、没有图的正文章 1 个" "$TMP/cf.txt" \
  || { echo "图表优先提示不对"; cat "$TMP/cf.txt"; exit 1; }
grep -q "「二、历史」：没有图" "$TMP/cf.txt"
grep -q "「二、历史」的表：45 个纯数值单元格" "$TMP/cf.txt"
for f in research-sop/standards/格式规范.md research-sop/SKILL.md; do
  grep -q "图表优先" "$f" || { echo "$f 缺图表优先规则"; exit 1; }
done
grep -q "^- V05 \*\*硬门槛\*\*" research-sop/standards/格式规范.md
grep -q "^- V02 软门槛" research-sop/standards/格式规范.md
grep -q "^- ⑩ 历史按轮次组织的报告，每个重要轮次有一张股价走势 + 阶段图" research-sop/standards/审稿清单.md
grep -q "^- ⑪ \`md2docx.py\`\"图表优先\"列出的数字墙、数值大表、没有图的正文章，供写作者参考" research-sop/standards/审稿清单.md
if grep -n "审稿清单.md\` 九条" research-sop/SKILL.md research-prompt/templates/任务书模板.md; then
  echo "审稿清单已是十一条，不得再写九条"; exit 1
fi
grep -q '只拿到这一份文档的冷读者"来写：事实本身写清楚，读这一份就够。' research-sop/SKILL.md \
  || { echo "research-sop D3 缺写法硬要求 5（写给冷读者）"; exit 1; }
if grep -q "冷读者" research-sop/standards/格式规范.md && grep -q '只拿到这一份文档的冷读者"来写' research-sop/standards/格式规范.md; then
  echo "写法硬要求已移到 research-sop D3，格式规范只留指针"; exit 1
fi
grep -q "⑨ 写给冷读者" research-sop/standards/审稿清单.md
echo "图表优先：md2docx 报出数字墙、数值大表、没有图的正文章且不影响退出码；V01–V05 写进格式规范（V05 硬门槛、V02 软门槛），审稿清单 ⑩ ⑪ 为规则；写法硬要求 5 写给冷读者在 research-sop D3 与审稿清单 ⑨ 里"

echo "== 7. check_delivery.py（离线，lsjson 样例） =="
python3 research-sop/scripts/check_delivery.py --help | grep -q -- "--drive"
mkdir -p "$TMP/task/成品" "$TMP/task/留痕/ACME/复用" "$TMP/task/留痕/ACME/复核"
touch "$TMP/task/成品/ACME_库存周期研究_20261105_V1.0.md" "$TMP/task/留痕/ACME/复用/结论卡.md" \
  "$TMP/task/留痕/ACME/复用/来源清单.csv" "$TMP/task/留痕/ACME/复用/数字台账.csv"
python3 tests/make_workbook_fixture.py "$TMP/wb" >/dev/null
python3 research-sop/scripts/check_delivery.py --lsjson tests/fixtures/check_delivery_good.json --local "$TMP/task" \
  --toc-text tests/fixtures/toc_good --workbook-dir "$TMP/wb/good" | tee "$TMP/cd_good.txt"
grep -q "结果：16/16 通过" "$TMP/cd_good.txt"
grep -q "13. 通过    图表工作簿：语义页名、图表总览看板、图表计算、来源页（硬性）（1 份；ACME_图表数据_20261105_V1.0：工作表 4 张，总览原生图表 2 张，全簿 3 张）" "$TMP/cd_good.txt"
grep -q "12. 通过    Google 文档目录：二级条目必须有页码（硬性）（2 份）" "$TMP/cd_good.txt"
grep -q "Google 文档）（2 份）" "$TMP/cd_good.txt"
grep -q "PNG 只在 底稿/图片/ 下（2 张）" "$TMP/cd_good.txt"
grep -q "1 篇：ACME；另有 底稿说明、共用数据/" "$TMP/cd_good.txt"
rm -r "$TMP/task/留痕/ACME/复核" "$TMP/task/留痕/ACME/复用/数字台账.csv"
if python3 research-sop/scripts/check_delivery.py --lsjson tests/fixtures/check_delivery_bad.json --local "$TMP/task" \
    > "$TMP/cd_bad.txt"; then
  echo "不合规的交付应以非零码退出"; exit 1
fi
cat "$TMP/cd_bad.txt"
for s in "结果：3/16 通过" "13. 通过    图表工作簿：语义页名、图表总览看板、图表计算、来源页（硬性）（没有配套图表工作簿，不适用）" "      - 20261105_库存周期_v1_ClaudeCode" \
    "ACME_库存周期研究_20261105_v1：取不到文档文本，也没有 --toc-evidence 人工确认（缺证据不通过）" "      - ACME_库存周期研究_20261105_v1" \
    "ACME_图表数据_20261105_V1.0.xlsx：xlsx 原件" "      - 报告/" \
    "底稿/ACME/结论卡.md：结论卡应为 Google 文档，现为 md 原件" "底稿/ACME/数字台账.csv：数字台账应为 Google 表格，现为 csv 原件" \
    "      - 底稿/ACME/图02_分部.png" "底稿/图片/XYZ/：没有同名的 底稿/XYZ/" "底稿/图片/图03_现金流.png：应放进 底稿/图片/<篇>/" \
    "      - 底稿/ACME/raw/" "留痕/ACME/ 缺 复核/" "留痕/ACME/ 缺 数字台账" \
    "底稿/共用/ 缺 结论卡、来源清单、数字台账" "底稿/共用数据/价格序列.csv：csv 原件" \
    "底稿/ACME_底稿说明_补充.docx：应放进 底稿/<篇>/" "底稿/数字台账.xlsx：应放进 底稿/<篇>/"; do
  grep -qF -- "$s" "$TMP/cd_bad.txt" || { echo "缺少预期输出：$s"; exit 1; }
done
if grep -q "自检" "$TMP/cd_good.txt" "$TMP/cd_bad.txt"; then echo "check_delivery 不再要求或检查“自检”文件"; exit 1; fi
python3 - "$TMP/cd_proc.json" <<'PY2'
import json, sys
d = json.load(open("tests/fixtures/check_delivery_good.json", encoding="utf8"))
tpl = next(x for x in d if not x.get("IsDir"))
for path, is_dir in (("过程", True), ("过程/状态.md", False), ("过程/ACME", True), ("过程/ACME/fig1.png", False)):
    e = dict(tpl); e.update(Path=path, Name=path.split("/")[-1], IsDir=is_dir, MimeType="inode/directory" if is_dir else "text/plain", Size=10)
    d.append(e)
json.dump(d, open(sys.argv[1], "w", encoding="utf8"), ensure_ascii=False)
PY2
python3 research-sop/scripts/check_delivery.py --lsjson "$TMP/cd_proc.json" --local "$TMP/task" --toc-text tests/fixtures/toc_good --workbook-dir "$TMP/wb/good" > "$TMP/cd_proc.txt" 2>&1 || true
if grep -q "过程/" "$TMP/cd_proc.txt"; then echo "过程/ 应整个不查"; cat "$TMP/cd_proc.txt"; exit 1; fi
echo "check_delivery：过程/ 整个不查"
if grep -q "共用数据/ 缺\|ACME_底稿说明_20261105" "$TMP/cd_bad.txt"; then echo "共用数据/ 不按篇要求三类、第一份底稿说明应放行"; exit 1; fi
if grep -q "来源清单.xlsx" "$TMP/cd_bad.txt"; then echo "无元数据、Size -1 的 .xlsx 应认作 Google 表格"; exit 1; fi
if grep -q "10-K_2025.pdf\|IR页面存档" "$TMP/cd_good.txt"; then echo "来源/ 里的原件应原样放行"; exit 1; fi
echo "check_delivery：合规样例（含底稿说明、共用数据/、合规图表工作簿）16/16 通过；不合规样例逐项报出命名、原件、自创文件夹、底稿/ 根下散文件、缺件（结论卡、来源清单、数字台账三类，含 共用数据/ 之外的文件夹；不再要求自检）、格式（含 共用数据/）、PNG 位置、raw、目录缺证据、本地缺 复核/ 与数字台账，退出码 1"
# 第 12 项：Google 文档目录二级条目必须有页码（硬性；缺页码、无目录段、缺证据都不通过）
CD="python3 research-sop/scripts/check_delivery.py --lsjson tests/fixtures/check_delivery_good.json --workbook-dir $TMP/wb/good"
toc_fail() {  # 用法：toc_fail 期望输出 -- 额外参数…
  local want="$1"; shift 2
  if $CD "$@" > "$TMP/cd_toc.txt"; then echo "目录页码不合规应以非零码退出：$*"; exit 1; fi
  grep -q "12. 不通过  Google 文档目录：二级条目必须有页码（硬性）" "$TMP/cd_toc.txt" \
    || { echo "第 12 项应不通过：$*"; cat "$TMP/cd_toc.txt"; exit 1; }
  grep -qF -- "$want" "$TMP/cd_toc.txt" || { echo "缺少预期输出：$want"; cat "$TMP/cd_toc.txt"; exit 1; }
}
toc_fail "ACME_库存周期研究_20261105_V1.0：二级条目无页码 3/3 条，首条「1.1 结论 …」" -- --toc-text tests/fixtures/toc_missing
toc_fail "半导体_存储周期研究_20261105：找不到“目录”一行" -- --toc-text tests/fixtures/toc_missing
toc_fail "ACME_库存周期研究_20261105_V1.0：取不到文档文本，也没有 --toc-evidence 人工确认（缺证据不通过）" --
toc_fail "ACME_库存周期研究_20261105_V1.0：人工确认写的是「二级目录带页码：否」" -- --toc-evidence tests/fixtures/toc_evidence_no.txt
toc_fail "半导体_存储周期研究_20261105：取不到文档文本，--toc-evidence 里没有这份文档的确认行（缺证据不通过）" -- \
  --toc-evidence tests/fixtures/toc_evidence_no.txt
toc_fail "--toc-evidence 文件不存在" -- --toc-evidence "$TMP/没有这个文件.txt"
toc_fail "--report 指定的「不存在的报告」不在根目录的 Google 文档里" -- --toc-text tests/fixtures/toc_good --report 不存在的报告
$CD --toc-evidence tests/fixtures/toc_evidence_good.txt > "$TMP/cd_toc.txt"
grep -q "结果：13/13 通过" "$TMP/cd_toc.txt"
$CD --toc-text tests/fixtures/toc_missing --toc-evidence tests/fixtures/toc_evidence_good.txt > "$TMP/cd_toc.txt" \
  && { echo "取到文本时人工证据不能推翻文本结果"; exit 1; }
$CD --toc-text tests/fixtures/toc_missing --report 半导体_存储周期研究_20261105 > "$TMP/cd_toc.txt" \
  && { echo "--report 指定的无目录文档应不通过"; exit 1; }
$CD --toc-text tests/fixtures/toc_good --report ACME_库存周期研究_20261105_V1.0 > "$TMP/cd_toc.txt"
grep -q "Google 文档目录：二级条目必须有页码（硬性）（1 份）" "$TMP/cd_toc.txt"
echo "check_delivery 第 12 项：二级目录带页码通过；缺页码、无目录段、无文本又无证据、证据写否、证据缺行、证据文件不存在、--report 不存在都不通过且退出码 1；有文本时证据不能推翻；--report 只核对指定报告"
# 第 13 项：图表工作簿结构（硬性；格式规范第 13.1 节 WB01、WB02、WB05、WB06）
WB="python3 research-sop/scripts/check_delivery.py --lsjson tests/fixtures/check_delivery_good.json --toc-text tests/fixtures/toc_good"
wb_fail() {  # 用法：wb_fail 期望输出 -- 额外参数…
  local want="$1"; shift 2
  if $WB "$@" > "$TMP/cd_wb.txt"; then echo "图表工作簿不合规应以非零码退出：$*"; exit 1; fi
  grep -q "13. 不通过  图表工作簿：语义页名、图表总览看板、图表计算、来源页（硬性）" "$TMP/cd_wb.txt" \
    || { echo "第 13 项应不通过：$*"; cat "$TMP/cd_wb.txt"; exit 1; }
  grep -qF -- "$want" "$TMP/cd_wb.txt" || { echo "缺少预期输出：$want"; cat "$TMP/cd_wb.txt"; exit 1; }
}
for want in "ACME_图表数据_20261105_V1.0：编号或默认页名 3 张（图1数据、Chart 2 Data、Sheet1）" \
    "ACME_图表数据_20261105_V1.0：第一张可见工作表是“说明”，应是“图表总览”" \
    "ACME_图表数据_20261105_V1.0：“图表总览”上没有原生图表——总览是集中全部重要数据图的看板，只有图号索引不算" \
    "ACME_图表数据_20261105_V1.0：没有名含“计算”的工作表" "ACME_图表数据_20261105_V1.0：没有名含“来源”的来源页"; do
  wb_fail "$want" -- --workbook-dir "$TMP/wb/bad"
done
wb_fail "ACME_图表数据_20261105_V1.0：“图表总览”的图表引用了 历年营收与业务拆分——总览图的数据区放在名含“计算”的工作表" -- \
  --workbook-dir "$TMP/wb/bad_ref"
wb_fail "ACME_图表数据_20261105_V1.0：取不到工作簿（--drive 时用 rclone 导出 xlsx；离线用 --workbook-dir）" --
wb_fail "--workbook 指定的「不存在的表格」不在根目录的 Google 表格里" -- --workbook-dir "$TMP/wb/good" --workbook 不存在的表格
$WB --report 半导体_存储周期研究_20261105 > "$TMP/cd_wb.txt"   # 该研究对象没有配套表格：不适用，通过
grep -q "13. 通过    图表工作簿：语义页名、图表总览看板、图表计算、来源页（硬性）（没有配套图表工作簿，不适用）" "$TMP/cd_wb.txt"
$WB --report ACME_库存周期研究_20261105_V1.0 --workbook-dir "$TMP/wb/good" > "$TMP/cd_wb.txt"   # 同研究对象、同版本号的表格被挑中
grep -q "（1 份；ACME_图表数据_20261105_V1.0：" "$TMP/cd_wb.txt"
for f in research-sop/standards/格式规范.md research-sop/SKILL.md research-sop/templates/financial_history_spec.md; do
  grep -q "13.1" "$f" || { echo "$f 缺格式规范第 13.1 节图表工作簿结构的引用"; exit 1; }
done
grep -q "^### 13.1 研究报告图表工作簿结构$" research-sop/standards/格式规范.md
grep -q "^- WB02 \*\*硬门槛\*\*" research-sop/standards/格式规范.md
grep -q "^- WB04 软门槛" research-sop/standards/格式规范.md
grep -q "^- ⑫ 有配套图表工作簿时，结构符合格式规范第 13.1 节 WB01–WB06，报告每张数据图在总览上有对应原生图" research-sop/standards/审稿清单.md
grep -q "共 13 条" research-sop/standards/审稿清单.md
grep -q "^- ⑬ 有配套图表工作簿时，首页按 Excel 规范建" research-sop/standards/审稿清单.md
test -s research-sop/standards/Excel规范.md && python3 -c "import json,sys; json.load(open(sys.argv[1], encoding='utf8'))" research-sop/templates/图表总览模板.json
if grep -n "审稿清单.md\` 十一条\|十一条机械检查\|审稿清单.md\` 十二条\|十二条机械检查" research-sop/SKILL.md research-prompt/templates/*.md; then
  echo "审稿清单已是十三条，不得再写十一条或十二条"; exit 1
fi
echo "check_delivery 第 13 项：合规图表工作簿通过；编号页名、总览不在第一张、总览无原生图表、无图表计算页、无来源页、总览图引用语义页、取不到工作簿、--workbook 不存在都不通过且退出码 1；没有配套表格记不适用；--report 按研究对象与版本号挑表格；格式规范 13.1（WB02 硬、WB04 软）、审稿清单 ⑫ ⑬ 与计数同步"
# 取消 AI 自检（0.21.0）：只留脚本检查，留档不再有自检
if grep -rn "trail_selfcheck\|自检九问\|写进自检" research-sop; then echo "research-sop 不得再出现 trail_selfcheck、自检九问、写进自检"; exit 1; fi
if test -e research-sop/templates/trail_selfcheck.md; then echo "research-sop/templates/trail_selfcheck.md 已删除"; exit 1; fi
grep -q "^# 交付检查清单" research-sop/standards/审稿清单.md || { echo "审稿清单.md 须为交付检查清单"; exit 1; }
grep -q "只有 ④ 数字比对卡交付" research-sop/standards/审稿清单.md || { echo "审稿清单须写明只有数字比对卡交付"; exit 1; }
test -s gowork/协作与衔接.md || { echo "缺 gowork/协作与衔接.md"; exit 1; }
grep -q "^## 1. 不变的原则" gowork/协作与衔接.md || { echo "协作与衔接须有不变的原则"; exit 1; }
grep -q "协作与衔接.md" gowork/SKILL.md research-sop/SKILL.md research-delivery/SKILL.md research-prompt/SKILL.md || { echo "各 skill 须指向协作与衔接"; exit 1; }
grep -q "### 本次细则" research-prompt/templates/任务书模板.md || { echo "任务书模板须有本次细则"; exit 1; }
for t in 题型_公司档案 题型_行业 题型_主题; do grep -q "篇幅默认不设上限" "research-prompt/templates/$t.md" || { echo "$t 不应限页数"; exit 1; }; done
if grep -nE "Opus|Sonnet|Fable" research-sop/SKILL.md research-delivery/SKILL.md gowork/协作与衔接.md research-prompt/templates/任务书模板.md; then echo "规则不应绑定具体模型"; exit 1; fi
echo "协作与衔接：不变原则、本次细则、不限页数、规则不绑模型"
grep -q "^### D4 脚本检查" research-sop/SKILL.md || { echo "research-sop SKILL.md 须有 D4 脚本检查"; exit 1; }
echo "取消 AI 自检：research-sop 无 trail_selfcheck、自检九问、写进自检，自检模板已删除；审稿清单为交付硬规则清单，D4 为脚本检查"
# 全面精简（0.21.0）：对话版 PROMPT.md 已删，只能对话的 AI 直接粘贴 SKILL.md；主文件字节数上限
if test -e research-sop/PROMPT.md; then echo "research-sop/PROMPT.md 已删除"; exit 1; fi
n=$(wc -c < research-sop/SKILL.md); [ "$n" -le 20000 ] || { echo "research-sop/SKILL.md 应 ≤ 20000 字符（wc -c），现为 $n"; exit 1; }
n=$(wc -c < research-sop/standards/格式规范.md); [ "$n" -le 15000 ] || { echo "格式规范.md 应 ≤ 15000 字符（wc -c），现为 $n"; exit 1; }
echo "全面精简：research-sop/PROMPT.md 不存在；SKILL.md ≤ 20000、格式规范.md ≤ 15000（wc -c）"

echo "== 8. research-prompt =="
grep -q "^name: research-prompt$" research-prompt/SKILL.md
test -s research-prompt/templates/任务书模板.md || { echo "缺文件：research-prompt/templates/任务书模板.md"; exit 1; }
if test -e research-prompt/模板库; then echo "本地不应有 模板库/（任务书模板库以 Drive 为准）"; exit 1; fi
if test -e research-prompt/模板库_已弃用; then echo "本地 模板库_已弃用/ 已删除，任务书模板库只在 Drive"; exit 1; fi
grep -q "<研究库>/90_任务书模板库/" research-prompt/SKILL.md || { echo "SKILL.md 须指向 Drive <研究库>/90_任务书模板库/"; exit 1; }
grep -q "先查 Drive" research-prompt/SKILL.md || { echo "SKILL.md 须写明先查 Drive 当前内容"; exit 1; }
if grep -n "从 \`模板库/\`\|模板库/README.md\|模板库/参考骨架.md" research-prompt/SKILL.md research-prompt/templates/任务书模板.md; then
  echo "SKILL.md 与任务书模板不得再把本地 模板库/ 当在用的库"; exit 1
fi
if grep -nE "drive.google.com|[A-Za-z0-9_-]{25,}" research-prompt/SKILL.md research-prompt/templates/任务书模板.md; then
  echo "SKILL.md 与任务书模板不得硬编码 Drive 链接或文件 ID（只写 <研究库>/ 下的文件夹路径）"; exit 1
fi
if grep -n "半导体\|周期" research-prompt/SKILL.md research-prompt/templates/任务书模板.md; then
  echo "research-prompt 的 SKILL.md 与任务书模板不得写题型内容（周期内容只放模板库）"; exit 1
fi
if grep -rn "先问目的\|做什么决定" research-prompt/; then echo "research-prompt 0.4.0 起不问用途"; exit 1; fi
grep -q "不做交易判断" research-prompt/templates/任务书模板.md || { echo "任务书模板第五节须固定写不做交易判断"; exit 1; }
grep -q "要读的角度" research-prompt/SKILL.md || { echo "research-prompt SKILL.md 须要求任务书第七节写“要读的角度”"; exit 1; }
grep -q "要读的角度" research-prompt/templates/任务书模板.md || { echo "任务书模板第七节须有“要读的角度”"; exit 1; }
grep -q "文献检索" research-sop/SKILL.md || { echo "research-sop SKILL.md 须有 D1“文献检索”"; exit 1; }
grep -q "角度清单" research-sop/SKILL.md || { echo "research-sop SKILL.md 须有 D2“角度清单”"; exit 1; }
for t in 题型_公司档案 题型_行业 题型_主题; do
  test -s "research-prompt/templates/$t.md" || { echo "缺题型模板：research-prompt/templates/$t.md"; exit 1; }
  grep -q "读者要能回答的问题" "research-prompt/templates/$t.md" && grep -q "本题型加项" "research-prompt/templates/$t.md" \
    || { echo "$t.md 须有“读者要能回答的问题”与“本题型加项”"; exit 1; }
done
if grep -n "待迁入 research-prompt" research-sop/SKILL.md; then
  echo "research-sop 0.20.0 起类型层已迁入题型模板，不应再有“待迁入”标注"; exit 1
fi
grep -q "第五节：股价行为画像与 σ 门槛" research-prompt/templates/题型_公司档案.md || { echo "公司档案题型模板须有第五节 σ 门槛"; exit 1; }
if grep -n "^## 公司报告七节\|^### 第五节：股价行为画像" research-sop/SKILL.md; then echo "公司报告七节与 σ 门槛已迁到公司档案题型模板"; exit 1; fi
echo "题型模板：公司档案、行业、主题三份齐，research-sop 无“待迁入”标注，七节与 σ 门槛只在公司档案模板"
echo "research-prompt：frontmatter 与任务书模板齐；任务书第七节有“要读的角度”，research-sop 有文献检索与角度清单；模板库以 Drive <研究库>/90_任务书模板库/ 为准、先查 Drive，本地不留副本；SKILL.md 无文件 ID、无题型内容"

echo "== 9. check_numbers.py 与机械活拆分 =="
python3 tests/make_numbers_fixture.py "$TMP/num"
python3 research-sop/scripts/check_numbers.py tests/fixtures/md2docx_sample.md --xlsx "$TMP/num/hist.xlsx" --ledger "$TMP/num/数字台账.csv" > "$TMP/cn_good.txt" \
  || { cat "$TMP/cn_good.txt"; echo "合规样例应全部命中、退出码 0"; exit 1; }
grep -q "数字比对：正文数字 19 个，命中 19 个（台账 6、数据表 13、csv 0），未命中 0 个" "$TMP/cn_good.txt" || { cat "$TMP/cn_good.txt"; exit 1; }
grep -q "1 个公式没有缓存值（先重算）" "$TMP/cn_good.txt"
# 改错一个数、加入不该抽的数（年份、日期、图号、版本、FY、σ、交易日、指数名、小整数），只应报出改错的那个
cp tests/fixtures/md2docx_sample.md "$TMP/cn_bad.md"
sed -i 's/52.9 亿美元，毛利率/52.6 亿美元，毛利率/' "$TMP/cn_bad.md"
sed -i 's/^最大的问题是订阅续约率能否维持；判断见第一章。/最大的问题是订阅续约率能否维持；判断见第一章。FY2025 起按 2025-03-31 口径，单日 3σ、250 个交易日、2.5σ 门槛，对标 S\&P 500，见图 1 与表 1、第 2.1 节，V1.0 版，近 5 年 3 家对手。/' "$TMP/cn_bad.md"
if python3 research-sop/scripts/check_numbers.py "$TMP/cn_bad.md" --xlsx "$TMP/num/hist.xlsx" --ledger "$TMP/num/数字台账.csv" > "$TMP/cn_bad.txt"; then
  echo "改错的数字应报未命中、退出码 1"; exit 1
fi
grep -q "未命中 1 个" "$TMP/cn_bad.txt" || { cat "$TMP/cn_bad.txt"; echo "只应报出改错的 52.6"; exit 1; }
grep -q "「1.1　结论」｜.*｜52.6" "$TMP/cn_bad.txt"
set +e; python3 research-sop/scripts/check_numbers.py tests/fixtures/md2docx_sample.md > /dev/null 2>&1; rc=$?; set -e
[ "$rc" -eq 2 ] || { echo "没有候选值来源时应退出码 2"; exit 1; }
# check_charts.py：图表防乱（一行最多两张、宽度、点数、横轴标签、标记点、图号）
python3 tests/make_charts_fixture.py "$TMP/ch"
python3 research-sop/scripts/check_charts.py "$TMP/ch/charts_good.xlsx" > "$TMP/ch_good.txt" || { cat "$TMP/ch_good.txt"; echo "合规图表工作簿应通过"; exit 1; }
grep -q "结果：3/3 张图通过" "$TMP/ch_good.txt"
grep -q "1180×540 px ｜ 点数 336 ｜ 标签 14" "$TMP/ch_good.txt"
if python3 research-sop/scripts/check_charts.py "$TMP/ch/charts_bad.xlsx" > "$TMP/ch_bad.txt"; then echo "不合规图表工作簿应退出码 1"; exit 1; fi
for s in "结果：0/3 张图通过" "G1 这一行有 3 张图，最多 2 张" "G3 点数 2000 > 200，宽须 ≥ 1100 px 独占一行" "G4 点数 2000 > 图宽 ÷ 3" \
         "G5 横轴非空标签 2000 个 > 15" "G6 1 个折线系列带标记点" "G8 图内标题带图号"; do
  grep -qF -- "$s" "$TMP/ch_bad.txt" || { cat "$TMP/ch_bad.txt"; echo "check_charts 缺少预期输出：$s"; exit 1; }
done
set +e; python3 research-sop/scripts/check_charts.py > /dev/null 2>&1; rc=$?; set -e
[ "$rc" -eq 2 ] || { echo "check_charts 无参数应退出码 2"; exit 1; }
grep -q "check_charts.py" research-delivery/SKILL.md || { echo "research-delivery 须要求跑 check_charts.py"; exit 1; }
# 机械活拆分（0.22.0）：格式与交付独立成 research-delivery，取数与抽查是 research-sop 工单
test -s research-delivery/SKILL.md && grep -q "^name: research-delivery$" research-delivery/SKILL.md
test -s research-delivery/交接单模板.md
test -s research-sop/工单/工单A_取数建表.md && test -s research-sop/工单/工单B_数字抽查.md
grep -q "research-delivery/SKILL.md" AGENTS.md || { echo "AGENTS.md 须指向 research-delivery"; exit 1; }
if grep -n "research-sop/SKILL.md" research-delivery/SKILL.md; then echo "research-delivery 自成一体，不引用 research-sop/SKILL.md"; exit 1; fi
if grep -n "图表可读性\|转换后修补清单\|xlsx_patch\|rclone" research-sop/SKILL.md; then echo "格式、转换、上传细节已移到 research-delivery"; exit 1; fi
grep -q "check_numbers.py" research-sop/SKILL.md && grep -q "check_numbers.py" research-sop/工单/工单B_数字抽查.md
echo "check_charts：合规样例 3/3 通过（月末 336 点 1180 px、每 2 年一标签、两张并排 600 px）；一行三张、日数据塞半宽、标签过密、带圆点、图内图号逐项报出，退出码 1；无参数退出码 2"
echo "check_numbers：合规样例 19/19 命中（台账 6、数据表 13）、报公式缺缓存；改错的数报出且退出码 1，年份、日期、图号、版本、FY、σ、交易日、指数名、小整数不抽；无候选值退出码 2；research-delivery、交接单模板、工单 A/B 齐，AGENTS.md 指向 research-delivery，research-sop 不含格式与上传细节"
echo
if [ ${#SKIPPED[@]} -gt 0 ]; then
  echo "已通过，但以下测试被跳过："
  for s in "${SKIPPED[@]}"; do echo "  - $s"; done
else
  echo "全部通过（无跳过）"
fi
