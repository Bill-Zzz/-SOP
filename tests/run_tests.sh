#!/usr/bin/env bash
# 自检：
#   1) CRWV 复现引擎与 Excel 对账
#   2) 研报库索引 / 检索 / 阅读准备（含日期误判等回归测试）
#   3) xlsx 修改流程：保留图表、公式、隐藏表；强制重算标记
#   4) LibreOffice 重算（未安装时明确报告为"跳过"）
#   5) 历史数据 Excel 脚本：用虚构公司 ACME 的最小长表生成工作簿并跑 --check；--png 导出图表快照
#   6) md2docx.py：示例 .md（标题、表格、现场生成的占位 PNG、图注）转 .docx 后读回检查
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
grep -q "结果：8/8 通过" "$TMP/hist.txt"
grep -q "股数口径：基本加权股数 5 年、期末股数 1 年" "$TMP/hist.txt"
python3 - "$TMP/hist" <<'EOF'
import glob, sys, zipfile, openpyxl
p = glob.glob(sys.argv[1] + "/*_ACME_历史数据_v1_测试.xlsx")
assert len(p) == 1, p
wb = openpyxl.load_workbook(p[0])
for s in ["图表总览", "说明", "历年营收与业务拆分", "年度三表", "年度衍生指标", "近12季度明细", "市值与营收对比",
          "市值与估值", "股价事件", "IPO指标", "所得税与现金税", "重大融资时间线", "重要收购", "指引与实际",
          "缺口", "来源", "图表计算"]:
    assert s in wb.sheetnames, f"缺工作表 {s}"
q = [c.value for c in wb["近12季度明细"][3][2:14]]
assert q[-1] == "FY2025Q2" and q[0] == "FY2022Q3", q
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
assert abs(vs.cell(row[2024], h["股数（百万股）"]).value - 1020) < 1e-6
assert "未采用" in str(vs.cell(row[2023], h["备注"]).value), "量级不对的期末股数应注明未采用"
print("历史数据 Excel：17 张表、13 张图、近 12 季、股价事件、公式、市值股数口径均正确")
EOF
if python3 -c "import matplotlib" 2>/dev/null; then
  python3 research-sop/scripts/build_history_xlsx.py ACME --root "$TMP/futu" --check-only "$TMP"/hist/*_ACME_历史数据_v1_测试.xlsx \
    --png "$TMP/png" | tee "$TMP/png.txt"
  grep -q "结果：9/9 通过" "$TMP/png.txt"
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
python3 research-sop/scripts/md2docx.py "$TMP/md/md2docx_sample.md" --out "$TMP/md/out.docx" | tee "$TMP/md.txt"
grep -q "标题 6（一级 1 / 二级 3 / 三级 2）；表格 2；图片 1" "$TMP/md.txt"
if python3 research-sop/scripts/md2docx.py "$TMP/md/md2docx_sample.md" --check-only --img-root "$TMP/nowhere" \
    >/dev/null 2>"$TMP/md_err.txt"; then
  echo "缺图时应以非零码退出"; exit 1
fi
grep -q "fig/图01_营收与增速.png" "$TMP/md_err.txt"
python3 - "$TMP/md/out.docx" <<'EOF'
import sys
import docx
from docx.oxml.ns import qn
d = docx.Document(sys.argv[1])
heads = [p for p in d.paragraphs if p.style.name.startswith("Heading")]
assert len(heads) == 6, [p.text for p in heads]
assert len(d.tables) == 2 and d.tables[0].style.name == "Table Grid"
assert len(d.tables[0].rows) == 4 and len(d.tables[0].columns) == 4
assert all(r.bold for r in d.tables[0].rows[0].cells[0].paragraphs[0].runs), "表头应加粗"
assert len(d.inline_shapes) == 1 and d.inline_shapes[0].width <= 15 * 360000
normal = d.styles["Normal"]
rf = normal.element.rPr.rFonts
assert normal.font.name == "Noto Sans SC" and rf.get(qn("w:eastAsia")) == "Noto Sans SC", "正文字体应为 Noto Sans SC（含东亚）"
assert normal.font.size.pt == 10.5
notes = [p for p in d.paragraphs if p.style.name == "注释小字"]
assert {p.text.split("：")[0] for p in notes} >= {"表注", "数据与可编辑图表", "来源"}, [p.text for p in notes]
assert d.styles["注释小字"].font.size.pt == 9
body = "\n".join(p.text for p in d.paragraphs)
assert "[R01234 p5]" in body and "**" not in body
print("md2docx：6 个标题、2 张真表格（表头加粗）、1 张图（≤ 15 cm）、正文 Noto Sans SC 10.5 pt、注释 9 pt、缺图报错均正确")
EOF

echo
if [ ${#SKIPPED[@]} -gt 0 ]; then
  echo "已通过，但以下测试被跳过："
  for s in "${SKIPPED[@]}"; do echo "  - $s"; done
else
  echo "全部通过（无跳过）"
fi
