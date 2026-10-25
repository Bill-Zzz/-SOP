#!/usr/bin/env python3
"""根据阅读清单准备阅读材料：全文抽取（扫描件自动 OCR）+ 阅读包骨架。

用法：
  python3 prepare_reading.py manifest.json reading/ [--lang chi_sim+eng] [--no-ocr]
产出：
  reading/text/SNN_<机构>_<日期>.txt    每份研报的全文（带"===== 第 N 页 ====="标记，便于标注页码）
  reading/reading_pack.md               阅读包骨架：每份研报一张摘要卡（元数据已预填）+ 对比表 + 分歧图
  reading/sources.csv                   来源清单（编号、机构、日期、标题、路径、抽取方式）
之后由 AI 阅读 text/ 下的全文，填写 reading_pack.md；关键表格用 pdf_text.py render 看原图。
"""
import csv
import json
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import pdf_text  # noqa: E402


def safe(s):
    return re.sub(r"[^\w一-鿿-]+", "_", s or "未知")[:30]


def main():
    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(1)
    manifest, out = sys.argv[1], sys.argv[2]
    lang = sys.argv[sys.argv.index("--lang") + 1] if "--lang" in sys.argv else "eng"
    no_ocr = "--no-ocr" in sys.argv
    m = json.load(open(manifest, encoding="utf8"))
    docs = m["docs"]
    cache_dir = os.path.join(os.path.dirname(m.get("db", "sva_index/library.db")), "text")
    os.makedirs(os.path.join(out, "text"), exist_ok=True)
    rows = []
    for i, d in enumerate(sorted(docs, key=lambda x: x.get("date") or ""), 1):
        sid = f"S{i:02d}"
        name = f"{sid}_{safe(d.get('broker'))}_{d.get('date') or '无日期'}.txt"
        pages = pdf_text.page_count(d["path"])
        cached = os.path.join(cache_dir, (d.get("sha1") or "-") + ".txt")
        if os.path.exists(cached) and not pdf_text.looks_scanned(open(cached, encoding="utf8").read(), pages):
            text, be = open(cached, encoding="utf8").read(), "索引缓存"
        else:
            text, be = pdf_text.extract_text(d["path"])
        if pdf_text.looks_scanned(text, pages) and not no_ocr:
            try:
                text, be = pdf_text.ocr_pdf(d["path"], lang)
                be += "(OCR)"
            except Exception as e:
                be += f"(OCR失败:{e})"
        parts = text.split("\f")
        marked = "".join(f"\n===== 第 {k} 页 =====\n{pg}" for k, pg in enumerate(parts, 1) if pg.strip())
        with open(os.path.join(out, "text", name), "w", encoding="utf8") as fh:
            fh.write(marked)
        rows.append(dict(id=sid, broker=d.get("broker") or "", date=d.get("date") or "", title=d.get("title") or "",
                         rating=d.get("rating") or "", target=d.get("target") or "", pages=pages or "",
                         method=be, text=f"text/{name}", path=d["path"]))
        print(f"{sid} [{be}] {d.get('broker')} {d.get('date')} {d.get('title', '')[:60]}")

    with open(os.path.join(out, "sources.csv"), "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.DictWriter(fh, fieldnames=list(rows[0].keys()) if rows else ["id"])
        w.writeheader()
        w.writerows(rows)

    tpl = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "templates", "reading_pack.md")
    head = open(tpl, encoding="utf8").read().split("<!-- CARDS -->")[0] if os.path.exists(tpl) else ""
    cards = []
    for r in rows:
        cards.append(
            f"### {r['id']} {r['broker']} {r['date']}：{r['title']}\n\n"
            f"- 评级 / 目标价：{r['rating'] or '待填'} / {r['target'] or '待填'}　页数：{r['pages']}　抽取：{r['method']}\n"
            f"- 全文：`{r['text']}`\n"
            f"- 核心观点（1–3 句）：\n- 估值方法与关键参数（附页码/附表）：\n- 关键数据点（附页码/附表）：\n"
            f"- 与其他机构的分歧：\n- 值得核实或存疑的地方：\n")
    with open(os.path.join(out, "reading_pack.md"), "w", encoding="utf8") as fh:
        fh.write(head + "\n".join(cards))
    print(f"完成：{len(rows)} 份。阅读包骨架：{os.path.join(out, 'reading_pack.md')}")


if __name__ == "__main__":
    main()
