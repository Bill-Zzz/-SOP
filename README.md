# 研究 SOP

用 AI 做股票与行业研究的一套方法、模板和脚本。仓库里有五个 skill，可以单独用，也可以串起来用。

| skill | 位置 | 用来做什么 |
|---|---|---|
| sellside-valuation-audit | 根目录 `SKILL.md` | 估值：从本地研报库读研报，和你讨论方向后建模型、写报告；核查已有报告的出处；新研报入库后滚动更新 |
| research-sop | `research-sop/` | 研究：主题、行业、公司三层。全自动，不停下讨论；未来只给方向和量级，不给点预测 |
| research-delivery | `research-delivery/` | 格式与交付：把写好的 .md、图片、数据表做成 Google 文档与表格，上传、补目录页码、交付核对；只改格式不改内容，可交给便宜的 AI 或 Codex |
| research-prompt | `research-prompt/` | 立项：盘料、和你谈出角度，写一份任务书交给执行研究的 AI |
| gowork | `gowork/` | 开工规矩与协作：先复述确认再动手，规则绑角色不绑 AI，`协作与衔接.md` 写清 Drive 衔接、并发和谁都能接手 |

当前版本见 `VERSION`，变更记录见 `CHANGELOG.md`。

## 怎么选
| 想要 | 用 |
|---|---|
| 一份研究：一个主题、一个行业横向比较，或一次读懂一家公司 | 先用 research-prompt 写任务书，再交给 research-sop 执行 |
| 这家公司值多少钱，或核查券商、自建模型 | sellside-valuation-audit |
| 两者都要 | 先 research-sop，再把报告作为输入交给 sellside-valuation-audit |

## 目录
```
SKILL.md  METHODS.md  AGENTS.md   估值 skill：完整方法、估值方法规范、编程智能体入口
LESSONS.md  CHANGELOG.md  VERSION            复盘收件箱、版本记录、版本号
scripts/          估值 skill 的脚本：研报库索引与检索、PDF 抽取与 OCR、Excel 导出 / 修补 / 重算 / 对比、独立复现
templates/        估值 skill 的模板：阅读包、研究方向、模型规格、报告大纲、出处台账、修订记录、复盘
examples/crwv/    CoreWeave 案例
research-sop/
  SKILL.md              研究方法（只写判断与写作）
  工单/                 机械步骤：工单A_取数建表、工单B_数字抽查
  standards/            格式规范、Excel 规范（图表总览首页）、存储规范、审稿清单（交付硬规则）、转换后修补清单
  templates/            报告外壳 report_shell.md，台账、来源清单、结论卡等留档模板、图表总览模板
  scripts/              build_history_xlsx.py 历史数据 Excel；check_numbers.py 正文数字比对；check_charts.py 图表防乱；md2docx.py 报告转 .docx；check_delivery.py 交付检查
  prompts/              发给执行 AI 的一句话提示词模板
research-delivery/
  SKILL.md              格式与交付（工单 C）；交接单模板.md
research-prompt/
  SKILL.md              立项方法
  templates/            任务书模板（各题型共用部分）；题型模板 题型_公司档案 / 题型_行业 / 题型_主题（只写各类报告的差异）；通过的任务书入 Google Drive 任务书模板库
gowork/SKILL.md         开工规矩
tests/                  回归测试，GitHub Actions 每次推送自动运行
```

## 安装
```bash
git clone https://github.com/Bill-Zzz/-SOP.git sellside-valuation-audit
cd sellside-valuation-audit
pip install openpyxl numpy pypdf python-docx matplotlib pillow   # 必需
pip install pymupdf rapidocr-onnxruntime                         # 可选：没有管理员权限时替代 poppler / tesseract
# 有管理员权限时可装系统工具：
#   macOS:  brew install poppler tesseract tesseract-lang && brew install --cask libreoffice
#   Ubuntu: sudo apt-get install -y poppler-utils tesseract-ocr tesseract-ocr-chi-sim libreoffice-calc
python3 scripts/pdf_text.py check    # 查看可用工具
bash tests/run_tests.sh              # 回归测试，会列出跳过的项
```

装进 Claude Code，用软链接，之后 `git pull` 即更新：
```bash
mkdir -p ~/.claude/skills
ln -s "$(pwd)"                 ~/.claude/skills/sellside-valuation-audit
ln -s "$(pwd)/research-sop"    ~/.claude/skills/research-sop
ln -s "$(pwd)/research-delivery" ~/.claude/skills/research-delivery
ln -s "$(pwd)/research-prompt" ~/.claude/skills/research-prompt
ln -s "$(pwd)/gowork"          ~/.claude/skills/gowork
```

其他工具：Codex、Cursor 等克隆到工作区即可，会自动读 `AGENTS.md`；格式与交付这类机械活可以直接交给它们，见 `AGENTS.md` 末节。只能对话的 AI（ChatGPT、Gemini、Kimi 等）直接粘贴 `SKILL.md` 或 `research-sop/SKILL.md` 为系统提示。

## 怎么用
**立项并研究：**
```text
按 research-prompt 帮我给「AI 资本开支周期」写任务书。
```
任务书定稿后，发给执行的 AI 一句话即可，模板在 `research-sop/prompts/模板_简短提示词.md`：
```text
你有我 Google Drive 的权限，用 Drive 时按 Drive 里的规则做。任务：<任务一句话>。方法按 research-sop <版本>。
```

**估值新研究：**
```text
研报库在 ~/研报库。帮我研究 CoreWeave（CRWV），按 sellside-valuation-audit 的模式 A 做。
```
AI 先建索引、检索并读完研报，给出阅读包和三到五个研究方向，然后停下等你选，选定后再出模型和报告。

**核查与更新：**
```text
按模式 B 核查附件里的报告和模型，研报原文在 ~/研报库/CRWV。
研报库新增了几份 CRWV 的研报，Q3 财报也出了，按模式 C 更新。
```

**改进方法本身：** 每次运行后的复盘建议记在 `LESSONS.md`。你确认后再改 skill，同时更新 `VERSION`、`CHANGELOG.md`，并通过 `bash tests/run_tests.sh`。

## 版权与合规
- 研报 PDF、抽取的文本和索引 `sva_index/` 不入库，`.gitignore` 已默认排除。
- 许多券商研报的使用条款限制转载或限制作为 AI 输入，使用前请确认授权范围。
- 方法和输出仅用于研究，不构成投资建议。
