#!/usr/bin/env python3
"""交付核对：上传到 Google Drive 的工作区任务文件夹是否按"工作区任务文件夹布局与交付清单"摆放、命名、转格式。

用法：
  python3 check_delivery.py --drive "gdrive:百万美金研究库/10_工作区/YYYYMMDD_任务名" [--local 本地任务文件夹]
  python3 check_delivery.py --lsjson 已保存的lsjson.json [--local 本地任务文件夹]     离线核对（测试用）

读 Drive：调用 `rclone lsjson -R -M <路径>`（-M 读出 Drive 的真实类型 content-type；rclone 版本不支持 -M 时自动去掉重试）。
rclone 把 Google 文档、表格显示成带 .docx / .xlsx 后缀、Size 为 -1 的文件：本脚本按 content-type 判断，没有元数据时
按"Size -1 且是 Office 类型"认作 Google 格式；命名检查用去掉导出后缀后的名字。

Drive 检查项（布局见 Drive 入口文件；命名按 standards/存储规范.md）：
  1  根目录至少有一份报告（Google 文档）
  2  根目录文件命名为 研究对象_研究主题_YYYYMMDD 或末尾加 _V1.0（正则 ^[^_]+_[^_]+_\\d{8}(_V\\d+\\.\\d+)?$）
  3  根目录文件只能是 Google 文档或 Google 表格
  4  根目录没有 png、docx、xlsx、md、csv 原件
  5  根目录只有 底稿/ 一个文件夹
  6  有 底稿/，其下至少一篇（底稿/<篇>/；图片/、共用数据/ 是特殊文件夹，不算篇）；底稿/ 根下只允许最多一份
     文件名含"底稿说明"的 Google 文档，其他文件须进 底稿/<篇>/
  7  每篇至少有 结论卡、自检、来源清单、数字台账 四类（按文件名含关键字判断；共用数据/ 不要求）
  8  底稿格式：结论卡、自检、复核记录、数据缺口是 Google 文档；来源清单、数字台账、研报摘要卡、历史节点、
     股价事件、台账、处理后数据是 Google 表格；来源/ 之外没有 md、csv、docx、xlsx、txt、json 原件（含 共用数据/）
  9  PNG 只在 底稿/图片/ 下
  10 底稿/图片/ 下按篇分子文件夹，名字与 底稿/<篇>/ 对应
  11 没有 raw 文件夹（raw 不上传，只在来源清单记路径）

本地检查项（给了 --local 才做；本地任务文件夹里 成品/ 可在根目录，也可在每篇子文件夹里）：
  L1 有 成品/，里面有报告（.md 或 .docx）
  L2 有 留痕/，每个 留痕/ 或 留痕/<篇>/ 下都有 复用/ 与 复核/
  L3 复用/ 里有 结论卡、来源清单、数字台账；复核/ 里有 自检

输出逐项"通过 / 不通过"及不通过的文件，最后一行"结果：x/y 通过"。全部通过退出码 0，有不通过 1，参数或 rclone 出错 2。
结果原样粘进自检的机器查部分。只用标准库；不改 Drive、不改本地文件。
"""
import argparse
import json
import os
import re
import subprocess
import sys

NAME_RE = re.compile(r"^[^_]+_[^_]+_\d{8}(_V\d+\.\d+)?$")
GDOC = "application/vnd.google-apps.document"
GSHEET = "application/vnd.google-apps.spreadsheet"
GOOGLE_PREFIX = "application/vnd.google-apps."
# rclone 导出 Google 格式时用的类型与后缀
EXPORT_MIME = {
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document": GDOC,
    "application/vnd.oasis.opendocument.text": GDOC,
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": GSHEET,
    "application/vnd.oasis.opendocument.spreadsheet": GSHEET,
}
EXPORT_EXT = (".docx", ".xlsx", ".pptx", ".odt", ".ods", ".odp", ".pdf", ".rtf", ".txt", ".csv", ".html", ".svg", ".json")
RAW_EXT = (".png", ".docx", ".xlsx", ".md", ".csv")          # 根目录不许出现的原件
TRAIL_RAW_EXT = (".md", ".csv", ".docx", ".xlsx", ".txt", ".json")  # 底稿/<篇>/（来源/ 之外）不许出现的中间件
DOC_KEYS = ("结论卡", "自检", "复核记录", "数据缺口")
SHEET_KEYS = ("来源清单", "数字台账", "研报摘要卡", "历史节点", "股价事件", "台账", "处理后数据")
REQUIRED = ("结论卡", "自检", "来源清单", "数字台账")
TRAIL, PICS, SOURCES, SHARED, NOTE = "底稿", "图片", "来源", "共用数据", "底稿说明"


class Item:
    """lsjson 的一行，换算成本脚本要用的字段。"""

    def __init__(self, d):
        self.path = d["Path"].strip("/")
        self.parts = self.path.split("/")
        self.is_dir = bool(d.get("IsDir"))
        name = d.get("Name") or self.parts[-1]
        mime = d.get("MimeType") or ""
        meta = (d.get("Metadata") or {}).get("content-type") or ""
        size = d.get("Size", 0)
        if meta.startswith(GOOGLE_PREFIX):
            self.kind = meta
        elif mime.startswith(GOOGLE_PREFIX):
            self.kind = mime
        elif not meta and size == -1 and mime in EXPORT_MIME:
            self.kind = EXPORT_MIME[mime]
        else:
            self.kind = ""
        self.google = self.kind.startswith(GOOGLE_PREFIX)
        base = name
        if self.google:
            for ext in EXPORT_EXT:
                if base.lower().endswith(ext):
                    base = base[: -len(ext)]
                    break
        self.name = base
        self.ext = "" if self.google else os.path.splitext(name)[1].lower()

    def label(self):
        if self.kind == GDOC:
            return "Google 文档"
        if self.kind == GSHEET:
            return "Google 表格"
        if self.google:
            return "其他 Google 格式（" + self.kind.rsplit(".", 1)[-1] + "）"
        return (self.ext.lstrip(".") or "无后缀") + " 原件"


def load_items(args):
    if args.lsjson:
        with open(args.lsjson, encoding="utf8") as f:
            data = json.load(f)
    else:
        base = ["rclone", "lsjson", "-R", args.drive]
        try:
            r = subprocess.run(base[:3] + ["-M"] + base[3:], capture_output=True, text=True)
            if r.returncode != 0:
                r = subprocess.run(base, capture_output=True, text=True)
        except FileNotFoundError:
            sys.exit("找不到 rclone（需要已配置 Google Drive 的 rclone）")
        if r.returncode != 0:
            print(r.stderr.strip(), file=sys.stderr)
            sys.exit(2)
        data = json.loads(r.stdout or "[]")
    return [Item(d) for d in data]


def trail_key(name):
    for k in DOC_KEYS:
        if k in name:
            return k, GDOC
    for k in SHEET_KEYS:
        if k in name:
            return k, GSHEET
    return None, None


def check_drive(items):
    res = []
    add = lambda name, bad, ok_note="": res.append((name, bad, ok_note))
    root_files = [i for i in items if len(i.parts) == 1 and not i.is_dir]
    root_dirs = [i for i in items if len(i.parts) == 1 and i.is_dir]

    docs = [i for i in root_files if i.kind == GDOC]
    add("根目录至少有一份报告（Google 文档）", [] if docs else ["根目录没有 Google 文档"], f"{len(docs)} 份")
    add("根目录文件命名：研究对象_研究主题_YYYYMMDD[_V1.0]",
        [f"{i.name}" for i in root_files if not NAME_RE.match(i.name)], f"{len(root_files)} 个文件")
    add("根目录文件只能是 Google 文档 / Google 表格",
        [f"{i.path}：{i.label()}" for i in root_files if i.kind not in (GDOC, GSHEET)])
    add("根目录没有 png、docx、xlsx、md、csv 原件",
        [i.path for i in root_files if not i.google and i.ext in RAW_EXT])
    add("根目录只有 底稿/ 一个文件夹", [i.path + "/" for i in root_dirs if i.name != TRAIL])

    trail_dir = [i for i in root_dirs if i.name == TRAIL]
    pieces = sorted({i.parts[1] for i in items if len(i.parts) >= 2 and i.parts[0] == TRAIL
                     and i.is_dir and len(i.parts) == 2 and i.parts[1] not in (PICS, SHARED)})
    bad = []
    if not trail_dir:
        bad.append("没有 底稿/")
    elif not pieces:
        bad.append("底稿/ 下没有篇的子文件夹")
    loose = [i for i in items if len(i.parts) == 2 and i.parts[0] == TRAIL and not i.is_dir]
    notes = [i for i in loose if NOTE in i.name and i.kind == GDOC][:1]  # 最多一份"底稿说明"Google 文档
    bad += [f"{i.path}：应放进 底稿/<篇>/（底稿/ 根下只放一份'底稿说明' Google 文档）" for i in loose if i not in notes]
    extra = [x for x in (("底稿说明" if notes else ""), ("共用数据/" if any(
        i.is_dir and i.parts == [TRAIL, SHARED] for i in items) else "")) if x]
    add("有 底稿/，按篇分子文件夹", bad, (f"{len(pieces)} 篇：{'、'.join(pieces)}" if pieces else "")
        + (f"；另有 {'、'.join(extra)}" if extra and pieces else ""))

    bad = []
    for p in pieces:
        files = [i for i in items if not i.is_dir and len(i.parts) == 3 and i.parts[:2] == [TRAIL, p]]
        miss = [k for k in REQUIRED if not any(k in i.name for i in files)]
        if miss:
            bad.append(f"底稿/{p}/ 缺 {'、'.join(miss)}")
    add("每篇有 结论卡、自检、来源清单、数字台账", bad)

    bad = []
    for i in items:
        if i.is_dir or len(i.parts) < 3 or i.parts[0] != TRAIL or i.parts[1] == PICS:
            continue
        if SOURCES in i.parts[2:-1]:
            continue  # 来源/ 原件原样
        k, want = trail_key(i.name)
        if k and i.kind != want:
            bad.append(f"{i.path}：{k}应为 {'Google 文档' if want == GDOC else 'Google 表格'}，现为 {i.label()}")
        elif not k and not i.google and i.ext in TRAIL_RAW_EXT:
            bad.append(f"{i.path}：{i.label()}，应转为 Google 文档 / 表格（中间件不上传）")
    add("底稿格式：文档类为 Google 文档、表格类为 Google 表格，来源/ 之外无原件", bad)

    pngs = [i for i in items if not i.is_dir and i.ext == ".png"]
    add("PNG 只在 底稿/图片/ 下", [i.path for i in pngs if i.parts[:2] != [TRAIL, PICS]], f"{len(pngs)} 张")

    bad = []
    for i in items:
        if i.parts[:2] != [TRAIL, PICS] or len(i.parts) < 3:
            continue
        if len(i.parts) == 3 and not i.is_dir:
            bad.append(f"{i.path}：应放进 底稿/图片/<篇>/")
        elif len(i.parts) == 3 and i.is_dir and i.parts[2] not in pieces:
            bad.append(f"{i.path}/：没有同名的 底稿/{i.parts[2]}/")
    add("底稿/图片/ 按篇分子文件夹、与篇对应", bad)

    add("没有 raw 文件夹（raw 不上传）", [i.path + "/" for i in items if i.is_dir and i.parts[-1].lower() == "raw"])
    return res


def _has(dirpath, key):
    for _, _, fs in os.walk(dirpath):
        if any(key in f for f in fs):
            return True
    return False


def check_local(root):
    res = []
    add = lambda name, bad, ok_note="": res.append((name, bad, ok_note))
    if not os.path.isdir(root):
        add("本地任务文件夹存在", [f"{root} 不存在"])
        return res
    found = {"成品": [], "留痕": []}
    for dp, dns, _ in os.walk(root):
        depth = os.path.relpath(dp, root).count(os.sep) + (0 if dp == root else 1)
        if depth > 2:
            dns[:] = []
            continue
        for d in dns:
            if d in found:
                found[d].append(os.path.join(dp, d))
    bad = []
    if not found["成品"]:
        bad.append("没有 成品/")
    for d in found["成品"]:
        fs = os.listdir(d)
        if not any(f.endswith((".md", ".docx")) for f in fs):
            bad.append(f"{os.path.relpath(d, root)}/ 里没有报告（.md 或 .docx）")
    add("本地有 成品/ 且有报告", bad, f"{len(found['成品'])} 处")

    units = []
    bad = []
    if not found["留痕"]:
        bad.append("没有 留痕/")
    for t in found["留痕"]:
        subs = [t] if {"复用", "复核"} & set(os.listdir(t)) else \
            [os.path.join(t, s) for s in sorted(os.listdir(t)) if os.path.isdir(os.path.join(t, s))]
        if not subs:
            bad.append(f"{os.path.relpath(t, root)}/ 是空的")
        for u in subs:
            units.append(u)
            miss = [s for s in ("复用", "复核") if not os.path.isdir(os.path.join(u, s))]
            if miss:
                bad.append(f"{os.path.relpath(u, root)}/ 缺 {'、'.join(m + '/' for m in miss)}")
    add("本地 留痕/ 有 复用/ 与 复核/", bad, f"{len(units)} 篇")

    bad = []
    for u in units:  # 缺 复用/ 或 复核/ 的已在上一项报出，这里只查在的那个
        rel = os.path.relpath(u, root)
        miss = []
        if os.path.isdir(os.path.join(u, "复用")):
            miss += [k for k in ("结论卡", "来源清单", "数字台账") if not _has(os.path.join(u, "复用"), k)]
        if os.path.isdir(os.path.join(u, "复核")):
            miss += [k for k in ("自检",) if not _has(os.path.join(u, "复核"), k)]
        if miss:
            bad.append(f"{rel}/ 缺 {'、'.join(miss)}")
    add("本地 复用/ 有结论卡、来源清单、数字台账，复核/ 有自检", bad)
    return res


def main(argv=None):
    ap = argparse.ArgumentParser(
        description="交付核对：Drive 工作区任务文件夹的布局、命名、格式（及本地 成品/、留痕/ 是否齐）。"
                    "全部通过退出码 0，有不通过 1，出错 2。结果粘进自检。")
    src = ap.add_mutually_exclusive_group(required=True)
    src.add_argument("--drive", help='rclone 路径，如 "gdrive:百万美金研究库/10_工作区/20261104_任务名"')
    src.add_argument("--lsjson", help="离线：读已保存的 rclone lsjson -R -M 输出（JSON 文件）")
    ap.add_argument("--local", help="本地任务文件夹（可选）：另查 成品/ 与 留痕/复用、复核 是否齐")
    args = ap.parse_args(argv)

    items = load_items(args)
    print(f"交付核对：{args.drive or args.lsjson}（{sum(not i.is_dir for i in items)} 个文件）")
    groups = [("Drive", check_drive(items))]
    if args.local:
        groups.append((f"本地 {args.local}", check_local(args.local)))
    n = ok = 0
    for title, res in groups:
        print(f"== {title} ==")
        for name, bad, note in res:
            n += 1
            if bad:
                print(f"{n:>2}. 不通过  {name}")
                for b in bad[:30]:
                    print(f"      - {b}")
                if len(bad) > 30:
                    print(f"      - ……另 {len(bad) - 30} 处")
            else:
                ok += 1
                print(f"{n:>2}. 通过    {name}" + (f"（{note}）" if note else ""))
    print(f"结果：{ok}/{n} 通过")
    return 0 if ok == n else 1


if __name__ == "__main__":
    sys.exit(main())
