# CHANGELOG

版本号规则：大版本.小版本.修订。流程或产出结构变化升小版本；脚本修复、陷阱补充升修订号；不兼容的重构升大版本。

## 0.2.0 — 2026-10-26
### 新增
- **模式 A：从研报库出发做新研究。** 流程为检索 → 阅读 → 讨论检查点 → 方法选择 → 建模 → 写报告 → 自检 → 交付 → 复盘。
- `scripts/library_index.py`：本地研报库增量索引（SQLite + FTS5），自动识别机构、日期、标题、代码、评级、目标价、语言，以及是否需要 OCR。
- `scripts/library_search.py`：按代码、公司名、关键词、机构、日期检索，按相关度排序，并输出阅读清单。
- `scripts/prepare_reading.py`：根据阅读清单抽取全文（必要时 OCR），生成阅读包骨架和来源清单。
- `scripts/pdf_text.py`：统一的 PDF 文本抽取、渲染和 OCR，按可用工具自动回退（PyMuPDF → poppler → pypdf；tesseract → RapidOCR），没有管理员权限也能用。
- 模板：`reading_pack.md`、`research_directions.md`、`model_spec.md`、`retro.md`；新增 `METHODS.md`（估值方法选择与建模规范）。
- **迭代机制：** 每次运行后写复盘 → 修改建议进入 `LESSONS.md` → 用户确认后改 SKILL，并记入 `CHANGELOG.md`、更新 `VERSION`，同时加回归测试。
- `xlsx_patch.py` 新增 `--full-calc-on-load`，在无法重算的环境里，让 Excel 打开文件时强制重算。

### 修改
- 原流程改为"模式 B：核查现有报告"；新增"模式 C：滚动更新"。
- 自检会明确报告跳过了哪些测试，并新增研报库索引与检索的回归测试。

## 0.1.0 — 2026-10-25
- 初版：复现、核对出处、修订 Excel 与报告、滚动更新；CoreWeave 案例。
