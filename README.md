# sellside-valuation-audit

把券商研报转化为可验证的自建估值研究，并持续维护的一套方法和工具：

1. 复现估值模型（Python 逐项对账 Excel）
2. 逐条对照研报原文核对出处，并量化每处问题对每股价值的影响
3. 修订 Excel（保留图表和公式）和报告
4. 随新财报、新研报滚动更新

任何 AI 工具都能用：能跑代码的智能体按 `SKILL.md` 执行；只能对话的 AI 用 `PROMPT.md`。

## 目录
```
SKILL.md                 完整方法（阶段 0–8 + 滚动维护 + 常见陷阱）
PROMPT.md                精简版指令，粘贴到任意对话式 AI 的系统提示
AGENTS.md                给 Codex / Cursor 等编程智能体的入口说明
scripts/
  extract_pdfs.sh        批量抽取研报文本，扫描件自动 OCR
  render_page.sh         渲染 PDF 指定页为 PNG，核对表格原图
  dump_xlsx.py           导出 Excel 的公式和缓存值
  replicate_harness.py   独立复现骨架：对账、反推、弹性表、曲线
  xlsx_patch.py          直接改 XML 修改单元格（不丢图表）
  lo_recalc.sh           LibreOffice 强制全量重算
  xlsx_diff.py           核对修订前后只改了打算改的地方
templates/               出处台账、核查报告、修订记录、修订版报告大纲
examples/crwv/           CoreWeave 完整案例（复现引擎 + 流程记录）
tests/run_tests.sh       自检
```

## 安装依赖
```bash
# macOS
brew install poppler tesseract tesseract-lang
brew install --cask libreoffice
# Ubuntu / Debian
sudo apt-get install -y poppler-utils tesseract-ocr tesseract-ocr-chi-sim libreoffice-calc
# Python
pip install openpyxl numpy
# 自检
bash tests/run_tests.sh
```

## 在不同 AI 工具里使用
| 工具 | 用法 |
|---|---|
| Claude（应用 / Claude Code） | 把整个目录放进技能目录（如 `~/.claude/skills/sellside-valuation-audit/`），或在对话里让 Claude 读取 `SKILL.md`。`SKILL.md` 带有标准的 name/description 头部 |
| OpenAI Codex / Cursor / 其他编程智能体 | 克隆仓库到工作区，智能体会读取 `AGENTS.md`，再按 `SKILL.md` 执行 |
| ChatGPT 自定义 GPT / 项目 | 把 `PROMPT.md` 的正文粘贴进指令；`SKILL.md` 和 `templates/` 作为知识文件上传 |
| Gemini Gem、Kimi、豆包、DeepSeek 等 | 把 `PROMPT.md` 粘贴为系统提示或项目说明，再上传报告、模型和研报 |

## 典型用法
```text
请按 sellside-valuation-audit 的方法：
1) 解读附件的估值报告和 Excel 模型；
2) 对照 reports/ 目录里的研报原文核对出处，给出核查报告和出处台账；
3) 修订 Excel 和报告，列出留给我确认的判断。
```
滚动更新：
```text
<公司> 发布了 Q4 财报（附件）。请按 SKILL.md 的"滚动维护"更新估值，并追加修订记录。
```

## 版权与合规
- **不要把研报原文或抽取的文本提交到仓库。** `.gitignore` 默认排除 PDF、抽取的文本、xlsx 和 docx。
- 许多券商研报的使用条款限制转载，或限制作为 AI 系统的输入。使用前请确认你的授权范围。
- 本仓库的方法和输出仅用于研究框架演示，不构成投资建议。
