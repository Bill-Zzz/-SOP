#!/usr/bin/env bash
# 自检：1) CRWV 复现引擎与 Excel 对账；2) xlsx 修改流程保留图表、重算正确。
# 用法：bash tests/run_tests.sh   （在仓库根目录运行）
set -euo pipefail
cd "$(dirname "$0")/.."
TMP="$(mktemp -d)"; trap 'rm -rf "$TMP"' EXIT

echo "== 1. 复现引擎对账 =="
python3 examples/crwv/engine_crwv.py check | tee "$TMP/check.txt"
grep -q "复现通过" "$TMP/check.txt"

echo "== 2. xlsx 修改流程 =="
python3 tests/make_fixture.py "$TMP/sample.xlsx" >/dev/null
cat > "$TMP/edits.json" <<'EOF'
{"假设": {"B1": 5.5, "H1": {"text": "新出处说明：机构 日期 附表 1"}}}
EOF
python3 scripts/xlsx_patch.py "$TMP/sample.xlsx" "$TMP/patched.xlsx" "$TMP/edits.json"
python3 - "$TMP/sample.xlsx" "$TMP/patched.xlsx" <<'EOF'
import sys, zipfile, openpyxl
a, b = sys.argv[1:]
charts = lambda p: [n for n in zipfile.ZipFile(p).namelist() if "/charts/" in n]
assert charts(a) == charts(b) and charts(b), "图表丢失"
wb = openpyxl.load_workbook(b)
assert wb["假设"]["B1"].value == 5.5 and wb["假设"]["H1"].value.startswith("新出处")
assert wb["假设"]["B3"].value == "=B2-B1", "公式被改动"
assert wb["隐藏计算"].sheet_state == "hidden"
print("修改后：图表、公式、隐藏表均保留")
EOF

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
  python3 scripts/xlsx_diff.py "$TMP/sample.xlsx" "$TMP/out/patched.xlsx" --values 假设!B3
else
  echo "（未安装 LibreOffice，跳过重算测试）"
fi
echo "全部通过"
