#!/usr/bin/env bash
# 用 LibreOffice 无界面模式对 xlsx 做全量重算并另存（保留图表）。
# LibreOffice 默认打开 xlsx 时不重算，缓存值会过期；这里用临时配置强制重算。
# 用法：bash lo_recalc.sh <in.xlsx> <输出目录> [--pdf]   （--pdf 额外导出 PDF 用于目检图表）
# 依赖：LibreOffice（soffice）
set -euo pipefail
IN="${1:?xlsx}"; OUT="${2:?输出目录}"; PDF="${3:-}"
SOFFICE="$(command -v soffice || command -v libreoffice || true)"
[ -n "$SOFFICE" ] || { echo "未找到 soffice，请先安装 LibreOffice"; exit 1; }
PROF="$(mktemp -d)"; mkdir -p "$PROF/user"
cat > "$PROF/user/registrymodifications.xcu" <<'EOF'
<?xml version="1.0" encoding="UTF-8"?>
<oor:items xmlns:oor="http://openoffice.org/2001/registry" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="OOXMLRecalcMode" oor:op="fuse"><value>0</value></prop></item>
<item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="ODFRecalcMode" oor:op="fuse"><value>0</value></prop></item>
</oor:items>
EOF
mkdir -p "$OUT"
"$SOFFICE" -env:UserInstallation="file://$PROF" --headless --calc --convert-to xlsx --outdir "$OUT" "$IN" >/dev/null 2>&1
if [ "$PDF" = "--pdf" ]; then
  "$SOFFICE" -env:UserInstallation="file://$PROF" --headless --convert-to pdf --outdir "$OUT" "$OUT/$(basename "$IN")" >/dev/null 2>&1
fi
rm -rf "$PROF"
ls -la "$OUT"
