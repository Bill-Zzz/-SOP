# AGENTS.md

本目录是一套"券商研报 → 自建估值研究"的复现、核查、修订和滚动更新方法。

- 完整方法：见 `SKILL.md`，按阶段 0–8 执行；滚动更新见同文件"滚动维护"一节。
- 工具脚本：`scripts/`，每个脚本开头都有用法说明。
- 模板：`templates/`。
- 完整案例：`examples/crwv/`（CoreWeave，2026-10）。

硬性规则：
1. 复现（`replicate_harness.py check` 误差小于 1e-6）通过之前，不评论模型假设。
2. 不要用 openpyxl 保存修订后的 Excel（会丢图表）。改用 `xlsx_patch.py` → `lo_recalc.sh` → `xlsx_diff.py`。
3. 每条批评都要有原文位置（机构、日期、页码/附表）和每股影响。
4. 有争议的口径不擅自改输入，写清另一口径的影响后交给用户决定。
