#!/usr/bin/env bash
# 批量抽取研报 PDF 文本；文本过少（扫描件/图片型）时自动 OCR。
# 用法：bash extract_pdfs.sh <pdf目录> <输出目录> [OCR语言，默认 eng，中文研报用 chi_sim+eng]
# 依赖：poppler-utils（pdftotext、pdftoppm、pdffonts）、tesseract
set -euo pipefail
IN="${1:?pdf目录}"; OUT="${2:?输出目录}"; LANGS="${3:-eng}"
MIN_BYTES=2000   # 低于此字节数视为图片型 PDF
mkdir -p "$OUT"
shopt -s nullglob
for f in "$IN"/*.pdf "$IN"/*.PDF; do
  name="$(basename "${f%.*}")"
  txt="$OUT/$name.txt"
  pdftotext -layout "$f" "$txt" 2>/dev/null || true
  size=$(wc -c < "$txt" 2>/dev/null || echo 0)
  pages=$(pdfinfo "$f" 2>/dev/null | awk '/^Pages:/{print $2}')
  if [ "${size:-0}" -lt "$MIN_BYTES" ]; then
    echo "[OCR] $name（$pages 页，文本仅 $size 字节）"
    tmp="$OUT/.ocr_$name"; mkdir -p "$tmp"
    pdftoppm -r 200 -png "$f" "$tmp/p"
    for img in "$tmp"/p-*.png; do
      tesseract "$img" "${img%.png}" -l "$LANGS" >/dev/null 2>&1 &
    done
    wait
    cat "$tmp"/p-*.txt > "$txt"
    echo "OCR" > "$OUT/$name.method"
  else
    echo "[TEXT] $name（$pages 页，$size 字节）"
    echo "TEXT" > "$OUT/$name.method"
  fi
done
echo "完成。注意：OCR 文本只用于定位，关键表格请用 render_page.sh 渲染原图核对。"
