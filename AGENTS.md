# AGENTS.md

这个仓库是一套"研报库驱动的估值研究"方法：从本地研报库检索并阅读研报，和用户讨论研究方向，再生成新的估值模型和研究报告；也可以核查已有报告、滚动更新估值。每次运行后复盘，持续迭代。

- 完整方法：`SKILL.md`（模式 A 新研究 / 模式 B 核查 / 模式 C 滚动更新 / 迭代本 skill）
- 估值方法与建模规范：`METHODS.md`
- 脚本：`scripts/`（每个脚本开头都有用法）；查看可用工具：`python3 scripts/pdf_text.py check`
- 模板：`templates/`；经验收件箱：`LESSONS.md`；版本：`VERSION`、`CHANGELOG.md`

硬性规则：
1. 模式 A 写完 `research_directions.md` 后必须停下，等用户选定方向，再建模和写报告。
2. 报告和模型里的每个数字都要有来源（编号 + 页码/附表）或标为本文假设；Excel 必须与 Python 引擎对账一致（误差小于 1e-6）。
3. 不要用 openpyxl 保存已有的 Excel（会丢图表），用 `xlsx_patch.py` → `lo_recalc.sh` → `xlsx_diff.py`。
4. 有争议的口径不擅自决定，写清两种结果的影响后交给用户。
5. 研报 PDF、抽取文本和索引（`sva_index/`）不要提交到仓库。
6. 每次运行结束都要写复盘，把修改建议追加到 `LESSONS.md`；经用户确认后才修改 SKILL，修改时同步更新 `VERSION`、`CHANGELOG.md`，并通过 `bash tests/run_tests.sh`。

## 研究报告的机械工单（research-sop 0.22.0 起）
接到下面三类活时，只读对应文件照做，不读 `research-sop/SKILL.md`；不改报告的数字和意思，内容问题写打回单：
- 格式、转换、上传、补目录页码、交付核对（交接单 `交接单.md`）：`research-delivery/SKILL.md`
- 取数建表：`research-sop/工单/工单A_取数建表.md`
- 数字抽查：`research-sop/工单/工单B_数字抽查.md`
