#!/usr/bin/env bash
# 把 PDF 的指定页渲染成 PNG，用于人工/多模态模型核对表格原图。
# 用法：bash render_page.sh <file.pdf> <页码> [输出前缀，默认 page] [分辨率，默认 150]
set -euo pipefail
F="${1:?pdf}"; P="${2:?页码}"; PREFIX="${3:-page}"; R="${4:-150}"
pdftoppm -r "$R" -png -f "$P" -l "$P" "$F" "$PREFIX"
ls "$PREFIX"*.png
