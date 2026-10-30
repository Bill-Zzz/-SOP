# sellside-valuation-audit

研报库驱动的股票估值研究方法与工具。三种模式：

- **A 新研究：** 从本地研报库检索相关研报并读完 → 和你讨论研究方向（在这里停下来等你决定）→ 生成新的估值模型（Excel + Python 对账）和研究报告 → 用核查流程检查自己的产出
- **B 核查：** 对照研报原文核对已有报告和模型的出处，量化每处问题的影响，修订 Excel（保留图表）和报告
- **C 滚动更新：** 新研报、新财报入库后增量更新

每次运行后都会复盘，把改进建议放进 `LESSONS.md`，经你确认后再升级方法本身（`VERSION` / `CHANGELOG.md`），并补上回归测试。

当前版本：见 `VERSION`。

## 目录
```
SKILL.md        完整方法（模式 A/B/C + 迭代机制 + 常见陷阱）
PROMPT.md       精简版，粘贴到任意对话式 AI
AGENTS.md       给 Codex / Cursor 等编程智能体的入口
METHODS.md      估值方法选择与新模型搭建规范
LESSONS.md      经验收件箱（待处理 / 已处理）
CHANGELOG.md    版本记录
scripts/
  library_index.py     本地研报库增量索引（机构 / 日期 / 标题 / 代码 / 评级 / 目标价 / 全文）
  library_search.py    检索并排序，输出阅读清单
  prepare_reading.py   抽取全文（含 OCR、页码标记），生成阅读包骨架
  pdf_text.py          PDF 文本 / 渲染 / OCR，自动选用可用工具（无需管理员权限也可用）
  extract_pdfs.sh      批量抽取文本（命令行版本）
  render_page.sh       渲染 PDF 页面
  dump_xlsx.py         导出 Excel 公式和缓存值
  replicate_harness.py 独立复现骨架：对账 / 反推 / 弹性 / 曲线
  xlsx_patch.py        修改已有 Excel 而不丢图表（可选打开时强制重算）
  lo_recalc.sh         LibreOffice 强制全量重算
  xlsx_diff.py         修订前后差异核对
templates/      阅读包、研究方向讨论、模型规格、报告大纲、出处台账、核查报告、修订记录、复盘
examples/crwv/  CoreWeave 案例
tests/          自检（GitHub Actions 自动运行）
company-dossier/        第二个 skill：公司档案研究（不建模、不停下、禁止点预测）
  SKILL.md              方法：范围 → 检索 → 阅读 → 主题/行业/公司三层写作 → 自检 → 交付 → 复盘
  PROMPT.md             精简版，粘贴到对话式 AI（必须联网）
  templates/            公司报告、行业报告、主题报告模板；股价事件台账、历史节点台账
```

## 两个 skill 怎么选
| 想要 | 用 |
|---|---|
| 一次读懂一家公司：历史、团队、股价事件、历史估值、未来利好利空、行业位置 | `company-dossier/`（全自动，不停下；未来只给方向和量级） |
| 这家公司值多少钱、核查券商或自建模型 | 根目录 `SKILL.md`（sellside-valuation-audit，A3 停下讨论方向） |
| 两者都要 | 先 company-dossier，再把档案报告作为输入交给 sellside-valuation-audit |

Claude Code 安装第二个 skill：
```bash
ln -s "$(pwd)/company-dossier" ~/.claude/skills/company-dossier
```

## 安装
```bash
git clone https://github.com/Bill-Zzz/-SOP.git sellside-valuation-audit
cd sellside-valuation-audit
pip install openpyxl numpy pypdf            # 必需
pip install pymupdf rapidocr-onnxruntime     # 可选：没有管理员权限时替代 poppler / tesseract
# 有管理员权限时也可以装系统工具：
#   macOS:  brew install poppler tesseract tesseract-lang && brew install --cask libreoffice
#   Ubuntu: sudo apt-get install -y poppler-utils tesseract-ocr tesseract-ocr-chi-sim libreoffice-calc
python3 scripts/pdf_text.py check            # 查看可用工具
bash tests/run_tests.sh                      # 自检（会说明跳过了哪些项）
```

### Claude Code：作为技能安装，并随仓库更新
```bash
mkdir -p ~/.claude/skills
ln -s "$(pwd)" ~/.claude/skills/sellside-valuation-audit   # 软链接，之后 git pull 即可更新
```

## 怎么用
**新研究：**
```text
研报库在 ~/研报库。帮我研究 CoreWeave（CRWV），按 sellside-valuation-audit 的模式 A 做。
```
AI 会先建索引并检索，读完后给你一份阅读包和 3–5 个研究方向，然后停下来等你选。你选定方向后，它再出模型和报告。

**核查：**
```text
按模式 B 核查附件里的报告和模型，研报原文在 ~/研报库/CRWV。
```

**滚动更新：**
```text
研报库新增了几份 CRWV 的研报，Q3 财报也出了，按模式 C 更新。
```

**迭代方法本身：**
```text
看一下 LESSONS.md 里待处理的建议，我同意第 1、3 条，帮我改进 skill 并发布新版本。
```

## 在其他 AI 工具中使用
| 工具 | 用法 |
|---|---|
| Claude Code | 如上，软链接到 `~/.claude/skills/` |
| Claude 应用 | 把技能保存到账户；关联电脑后，让 Claude 访问研报库所在的文件夹 |
| Codex / Cursor 等 | 克隆仓库到工作区，智能体会自动读取 `AGENTS.md` |
| ChatGPT / Gemini / Kimi 等对话式 AI | 把 `PROMPT.md` 粘贴为系统提示，再上传研报 |

## 版权与合规
- 研报 PDF、抽取的文本和索引（`sva_index/`）不要提交到仓库，`.gitignore` 已默认排除。
- 许多券商研报的使用条款限制转载，或限制作为 AI 输入，使用前请确认你的授权范围。
- 方法和输出仅用于研究，不构成投资建议。
