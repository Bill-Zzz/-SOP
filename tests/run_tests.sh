#!/usr/bin/env bash
# 自检：
#   1) CRWV 复现引擎与 Excel 对账
#   2) 研报库索引 / 检索 / 阅读准备（含日期误判等回归测试）
#   3) xlsx 修改流程：保留图表、公式、隐藏表；强制重算标记
#   4) LibreOffice 重算（未安装时明确报告为"跳过"）
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

echo
if [ ${#SKIPPED[@]} -gt 0 ]; then
  echo "已通过，但以下测试被跳过："
  for s in "${SKIPPED[@]}"; do echo "  - $s"; done
else
  echo "全部通过（无跳过）"
fi
