#!/usr/bin/env python3
"""交付核对：上传到 Google Drive 的工作区任务文件夹是否按"工作区任务文件夹布局与交付清单"摆放、命名、转格式。

用法：
  python3 check_delivery.py --drive "gdrive:<研究库>/10_工作区/YYYYMMDD_任务名" [--local 本地任务文件夹]
  python3 check_delivery.py --lsjson 已保存的lsjson.json [--local 本地任务文件夹]     离线核对（测试用）
  第 12 项（目录页码）可另加：--report 报告名（可多次） --toc-text 导出文本目录 --toc-evidence 人工确认文件
  第 13 项（图表工作簿结构）可另加：--workbook 工作簿名（可多次） --workbook-dir 导出 xlsx 的目录（离线）

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
  12 Google 文档目录：二级条目必须有页码（硬性；格式规范第 5 节"一级无页码、二级 标题　…　页码"）。
     取文本：--drive 时对每份报告 `rclone cat "<路径>/<文档名>.txt" --drive-export-formats txt`；离线用 --toc-text。
     在文本里找"目录"一行，往下读目录条目，到正文重新出现第一个条目为止；形如"1.1 标题"的二级条目，
     行末须是页码（"标题　…　12""标题\\t12""标题 12"），任一条没有即不通过；找不到目录段、目录没有条目也不通过。
     取不到某份报告的文本时，必须由 --toc-evidence 给出人工确认（每行：文档名｜二级目录带页码：是｜YYYY-MM-DD｜确认人），
     缺文件、缺这一行、写"否"、格式不对都不通过——不默认通过。取到了文本就以文本为准，人工证据不能推翻文本结果。
     默认核对根目录全部 Google 文档；同一任务文件夹里有旧版本或非报告文档时，用 --report 指定本次交付的报告（可多次）。
  13 图表工作簿结构（硬性；格式规范第 13.1 节 WB01、WB02、WB05、WB06）：根目录名含"图表数据"或"历史数据"的
     Google 表格，导出为 xlsx（`rclone cat "<路径>/<表格名>.xlsx" --drive-export-formats xlsx`；离线用 --workbook-dir）
     后逐份查：工作表用语义名，没有"图1数据""Chart 2 Data""Sheet1"这类编号或默认页名；第一张可见工作表是"图表总览"，
     上面有原生图表（只有图号索引、没有图表不算看板），总览图表的数据引用指向名含"计算"的工作表；有名含"计算"的
     工作表（如"图表计算"）；有名含"来源"的来源页。取不到工作簿即不通过。
     另查版面（转换后修补清单 S5、S6）：每张工作表读 drawing 锚点（twoCellAnchor 的 from/to；oneCellAnchor 的 from + ext
     高，按该表行高换算，没写行高按 Google 默认 21 px），任意两张图（原生图表或图片）的行区间相交即不通过，打印相交的
     两张图与各自行区间；每张工作表的行数（<dimension> 末行与 <row> 的较大者）超过"最后有内容的行或最后图底边 + 50"
     即不通过（Google 表格默认每页 1000 行，转换后须裁到内容区 + 5 行）。
     核对哪几份：给了 --workbook 只查这几份；否则给了 --report 时，查与报告同一研究对象、版本号相同或不带版本号的
     图表数据 / 历史数据表格；都没给时查根目录全部这类表格。一份都没有时记"不适用"（报告没有配套图表工作簿）。
     来源 ID 能否追到、报告每张数据图在总览上有没有对应原生图，脚本查不了，由审稿清单 ⑫ 人工查。

本地检查项（给了 --local 才做；本地任务文件夹里 成品/ 可在根目录，也可在每篇子文件夹里）：
  L1 有 成品/，里面有报告（.md 或 .docx）
  L2 有 留痕/，每个 留痕/ 或 留痕/<篇>/ 下都有 复用/ 与 复核/
  L3 复用/ 里有 结论卡、来源清单、数字台账；复核/ 里有 自检

输出逐项"通过 / 不通过"及不通过的文件，最后一行"结果：x/y 通过"。全部通过退出码 0，有不通过 1，参数或 rclone 出错 2。
结果原样粘进自检的机器查部分。只用标准库；不改 Drive、不改本地文件。
"""
import argparse
import io
import json
import math
import os
import posixpath
import re
import subprocess
import sys
import zipfile
from xml.etree import ElementTree as ET

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
# 目录检查（第 12 项）
TOC_TITLE = "目录"
TOC_L1_RE = re.compile(r"^(?:[一二三四五六七八九十]+、|附[一二三四五六七八九十]*[：:、]|附录|一页摘要)")
TOC_L2_RE = re.compile(r"^\d+\.\d+[\s　]")
TOC_PAGE_RE = re.compile(r"(?:[\t　…]|\.{2,}|\s)\s*\d{1,4}$")
TOC_TAIL_RE = re.compile(r"[\s　….]*\d{0,4}$")
TOC_CLAIM = "二级目录带页码：是"
DATE_RE = re.compile(r"^\d{4}-\d{2}-\d{2}$")
# 图表工作簿检查（第 13 项）
WB_KEYS = ("图表数据", "历史数据")
WB_OVERVIEW, WB_CALC, WB_SOURCE = "图表总览", "计算", "来源"
WB_NUMBERED_RE = re.compile(r"^\s*(?:图表?|表|fig(?:ure)?\.?|chart|sheet|工作表)\s*[0-9０-９]+", re.I)
WB_VERSION_RE = re.compile(r"_V\d+\.\d+$")
WB_ROW_PT = 15.75        # 没写行高时按 Google 表格默认行高 21 px = 15.75 pt 换算图高
EMU_PER_PT = 12700
WB_EPS = 0.05            # 行区间相交的容差（行）：上一张底边正好贴着下一张顶边不算相交
WB_SPARE_ROWS = 50       # 行数超过"最后有内容的行或最后图底边 + 50"即不通过（修补清单 S6 要求裁到 + 5）


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


def _toc_key(line):
    """目录条目去掉行末分隔符与页码后的标题，用来认出正文重新出现的第一个条目。"""
    return TOC_TAIL_RE.sub("", line).strip()


def parse_toc(text):
    """返回 (目录条目列表, None)；找不到目录段时返回 (None, 原因)。"""
    lines = [ln.replace("﻿", "").strip() for ln in text.replace("\r", "\n").split("\n")]
    lines = [ln for ln in lines if ln]
    start = next((k for k, ln in enumerate(lines) if ln == TOC_TITLE), None)
    if start is None:
        return None, "找不到“目录”一行"
    entries = []
    for ln in lines[start + 1:]:
        is_entry = (TOC_L1_RE.match(ln) or TOC_L2_RE.match(ln) or "…" in ln or re.search(r"\t\d{1,4}$", ln))
        if not is_entry or (entries and _toc_key(ln) == _toc_key(entries[0])) or len(entries) >= 400:
            break  # 正文从第一个目录条目的标题重新开始，或已不像目录条目
        entries.append(ln)
    if not entries:
        return None, "“目录”下面没有条目"
    return entries, None


def toc_problems(text):
    """一份文档的目录问题列表；空列表 = 通过。"""
    entries, why = parse_toc(text)
    if why:
        return [why]
    l2 = [e for e in entries if TOC_L2_RE.match(e)]
    miss = [e for e in l2 if not TOC_PAGE_RE.search(e)]
    if miss:
        return [f"二级条目无页码 {len(miss)}/{len(l2)} 条，首条「{miss[0]}」——按转换后修补清单 D1 补页码后重跑"]
    return []


def load_toc_evidence(path):
    """人工确认证据：每行 文档名｜二级目录带页码：是｜YYYY-MM-DD｜确认人。返回 {文档名: 问题或 None}，文件问题放在 '' 键。"""
    if not path:
        return {}
    if not os.path.isfile(path):
        return {"": f"--toc-evidence 文件不存在：{path}"}
    out = {}
    with open(path, encoding="utf8") as f:
        for ln in f:
            ln = ln.strip()
            if not ln or ln.startswith("#"):
                continue
            parts = [p.strip() for p in re.split(r"[｜|]", ln)]
            if len(parts) < 4 or not parts[0]:
                out.setdefault(parts[0] if parts else "", "证据行格式不对（文档名｜二级目录带页码：是｜日期｜确认人）")
                continue
            name, claim, date, who = parts[:4]
            if claim.replace(" ", "").replace(":", "：") != TOC_CLAIM:
                out[name] = f"人工确认写的是「{claim}」，不是「{TOC_CLAIM}」"
            elif not DATE_RE.match(date):
                out[name] = f"人工确认的日期「{date}」不是 YYYY-MM-DD"
            elif not who:
                out[name] = "人工确认缺确认人"
            else:
                out[name] = None
    return out


def export_doc_text(drive, item):
    """用 rclone 把 Drive 上的 Google 文档导出成纯文本；失败返回 None。"""
    path = "/".join(item.parts[:-1] + [item.name + ".txt"])
    try:
        r = subprocess.run(["rclone", "cat", f"{drive.rstrip('/')}/{path}", "--drive-export-formats", "txt"],
                           capture_output=True, text=True)
    except FileNotFoundError:
        return None
    return r.stdout if r.returncode == 0 and r.stdout.strip() else None


def check_toc(items, args):
    """第 12 项：Google 文档目录二级条目必须有页码。缺文本又缺人工证据 = 不通过。"""
    docs = [i for i in items if len(i.parts) == 1 and not i.is_dir and i.kind == GDOC]
    bad = []
    if args.report:
        names = {d.name for d in docs}
        bad += [f"--report 指定的「{r}」不在根目录的 Google 文档里" for r in args.report if r not in names]
        docs = [d for d in docs if d.name in args.report]
    if not docs:
        bad.append("没有要核对目录的 Google 文档")
    evidence = load_toc_evidence(args.toc_evidence)
    if "" in evidence and evidence[""]:
        bad.append(evidence[""])
    for d in docs:
        text = None
        if args.toc_text:
            p = os.path.join(args.toc_text, d.name + ".txt")
            if os.path.isfile(p):
                with open(p, encoding="utf8") as f:
                    text = f.read()
        elif args.drive:
            text = export_doc_text(args.drive, d)
        if text is not None:
            bad += [f"{d.name}：{x}" for x in toc_problems(text)]
            continue
        if not args.toc_evidence:
            bad.append(f"{d.name}：取不到文档文本，也没有 --toc-evidence 人工确认（缺证据不通过）")
        elif d.name not in evidence:
            bad.append(f"{d.name}：取不到文档文本，--toc-evidence 里没有这份文档的确认行（缺证据不通过）")
        elif evidence[d.name]:
            bad.append(f"{d.name}：{evidence[d.name]}")
    return [("Google 文档目录：二级条目必须有页码（硬性）", bad, f"{len(docs)} 份")]


REL_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"


def _rels(z, part):
    """读 part 的关系文件，返回 {Id: (Type 末段, 目标部件路径)}。"""
    d, f = posixpath.split(part)
    path = posixpath.join(d, "_rels", f + ".rels")
    if path not in z.namelist():
        return {}
    out = {}
    for el in ET.fromstring(z.read(path)):
        tgt = el.get("Target", "")
        tgt = tgt[1:] if tgt.startswith("/") else posixpath.normpath(posixpath.join(d, tgt))
        out[el.get("Id")] = (el.get("Type", "").rsplit("/", 1)[-1], tgt)
    return out


def _ref_sheet(ref):
    """图表数据引用 'Sheet 名'!$A$1:$A$9 → Sheet 名。"""
    m = re.match(r"^\s*'((?:[^']|'')+)'!|^\s*([^!'\s]+)!", ref or "")
    if not m:
        return None
    return (m.group(1) or m.group(2)).replace("''", "'")


def _local(el):
    return el.tag.rsplit("}", 1)[-1]


def _child(el, name):
    return next((c for c in el if _local(c) == name), None)


def _int(el, name, default=0):
    c = _child(el, name) if el is not None else None
    try:
        return int(c.text) if c is not None and c.text else default
    except ValueError:
        return default


def _sheet_rows(z, part):
    """工作表 XML → (行高表 {行号(1 起): 磅}, 默认行高磅, 最大行号, 最后有内容的行号)。
    最大行号取 <dimension> 末行与 <row r> 的较大者（Google 导出会把默认 1000 行都写成空 <row>）；
    "有内容"= 单元格带值、公式或内联文本，只有样式的空单元格不算。"""
    heights, default, max_row, last = {}, WB_ROW_PT, 0, 0
    if part not in z.namelist():
        return heights, default, max_row, last
    for ev, el in ET.iterparse(io.BytesIO(z.read(part)), events=("end",)):
        tag = _local(el)
        if tag == "dimension":
            m = re.search(r"(\d+)$", el.get("ref", ""))
            if m:
                max_row = max(max_row, int(m.group(1)))
        elif tag == "sheetFormatPr":
            try:
                default = float(el.get("defaultRowHeight") or default)
            except ValueError:
                pass
        elif tag == "row":
            try:
                r = int(el.get("r", "0"))
            except ValueError:
                continue
            max_row = max(max_row, r)
            if el.get("ht"):
                try:
                    heights[r] = float(el.get("ht"))
                except ValueError:
                    pass
            if any(_local(c) == "c" and any(_local(x) in ("v", "f", "is") for x in c) for c in el):
                last = max(last, r)
            el.clear()
    return heights, default, max_row, last


def _emu_to_row(emu, heights, default):
    """从第 1 行顶边往下 emu 处落在第几行（0 起、带小数）。"""
    r, left = 0, emu
    while True:
        h = heights.get(r + 1, default) * EMU_PER_PT
        if h <= 0:
            r += 1
            continue
        if left < h or r > 100000:
            return r + left / h
        left -= h
        r += 1


def _row_to_emu(row, off, heights, default):
    return sum(heights.get(i + 1, default) * EMU_PER_PT for i in range(row)) + off


def _chart_title(z, chart):
    try:
        root = ET.fromstring(z.read(chart))
    except (KeyError, ET.ParseError):
        return ""
    t = next((e for e in root.iter() if _local(e) == "title"), None)
    if t is None:
        return ""
    return "".join(e.text or "" for e in t.iter() if _local(e) == "t").strip()


def _anchors(z, drawing, heights, default):
    """drawing XML → [{label, top, bottom}]：每张图（原生图表或图片）占的行区间（0 起、带小数，bottom 不含）。
    twoCellAnchor 用 from/to；oneCellAnchor 用 from + ext 的高按行高换算；absoluteAnchor 用 pos + ext。"""
    if drawing not in z.namelist():
        return []
    rels = _rels(z, drawing)
    out = []
    for a in ET.fromstring(z.read(drawing)):
        kind = _local(a)
        if kind not in ("twoCellAnchor", "oneCellAnchor", "absoluteAnchor"):
            continue
        label = "图片"
        for e in a.iter():
            if _local(e) == "chart" and e.get(REL_NS):
                tgt = rels.get(e.get(REL_NS), ("", ""))[1]
                ttl = _chart_title(z, tgt)
                label = "图表" + (f"“{ttl}”" if ttl else "")
                break
        else:
            if not any(_local(e) == "pic" for e in a.iter()):
                continue
        ext = _child(a, "ext")
        cy = int(ext.get("cy", "0")) if ext is not None else 0
        if kind == "twoCellAnchor":
            f, t = _child(a, "from"), _child(a, "to")
            top_emu = _row_to_emu(_int(f, "row"), _int(f, "rowOff"), heights, default)
            bot_emu = _row_to_emu(_int(t, "row"), _int(t, "rowOff"), heights, default)
        elif kind == "oneCellAnchor":
            f = _child(a, "from")
            top_emu = _row_to_emu(_int(f, "row"), _int(f, "rowOff"), heights, default)
            bot_emu = top_emu + cy
        else:
            pos = _child(a, "pos")
            top_emu = int(pos.get("y", "0")) if pos is not None else 0
            bot_emu = top_emu + cy
        out.append({"label": label, "top": _emu_to_row(top_emu, heights, default),
                    "bottom": _emu_to_row(bot_emu, heights, default)})
    for i, x in enumerate(out, start=1):
        x["label"] = f"第 {i} 张{x['label']}"
    return out


def parse_workbook(data):
    """xlsx 字节 → [{name, hidden, charts, refs, anchors, max_row, last_row}]，按工作簿里的顺序。
    只用标准库，读 Google 导出与 openpyxl 写出的 xlsx。"""
    z = zipfile.ZipFile(io.BytesIO(data))
    wb_rels = _rels(z, "xl/workbook.xml")
    sheets = []
    for el in ET.fromstring(z.read("xl/workbook.xml")).iter():
        if not el.tag.endswith("}sheet"):
            continue
        part = wb_rels.get(el.get(REL_NS), ("", ""))[1]
        heights, default, max_row, last = _sheet_rows(z, part)
        charts, refs, anchors = 0, set(), []
        for typ, drawing in _rels(z, part).values():
            if typ != "drawing":
                continue
            anchors += _anchors(z, drawing, heights, default)
            for typ2, chart in _rels(z, drawing).values():
                if typ2 != "chart" or chart not in z.namelist():
                    continue
                charts += 1
                for f in ET.fromstring(z.read(chart)).iter():
                    if f.tag.endswith("}f") and f.text:
                        sh = _ref_sheet(f.text)
                        if sh:
                            refs.add(sh)
        sheets.append({"name": el.get("name", ""), "hidden": el.get("state", "visible") != "visible",
                       "charts": charts, "refs": refs, "anchors": anchors, "max_row": max_row, "last_row": last})
    return sheets


def _rows_text(a):
    """行区间（0 起、bottom 不含）→ 1 起的"第 x–y 行"。"""
    top = int(a["top"]) + 1
    bot = max(top, math.ceil(a["bottom"] - 1e-6))
    return f"第 {top}–{bot} 行"


def layout_problems(sheets):
    """图表互不重叠、页签无多余空行（转换后修补清单 S5、S6）。"""
    bad = []
    for s in sheets:
        an = s.get("anchors", [])
        for i in range(len(an)):
            for j in range(i + 1, len(an)):
                a, b = an[i], an[j]
                if a["top"] < b["bottom"] - WB_EPS and b["top"] < a["bottom"] - WB_EPS:
                    bad.append(f"“{s['name']}”：{a['label']}（{_rows_text(a)}）与{b['label']}（{_rows_text(b)}）"
                               "行区间相交——图表单列纵向排布，按转换后修补清单 S5 重新摆")
        bottom = max([math.ceil(a["bottom"] - 1e-6) for a in an] + [0])
        used = max(s.get("last_row", 0), bottom)
        if s.get("max_row", 0) > used + WB_SPARE_ROWS:
            bad.append(f"“{s['name']}”：{s['max_row']} 行，内容与图只到第 {used} 行（超出 {WB_SPARE_ROWS} 行以上）"
                       "——按转换后修补清单 S6 删到内容区 + 5 行")
    return bad


def workbook_problems(sheets):
    """一份图表工作簿的问题列表与说明；问题为空 = 通过。"""
    bad = []
    names = [s["name"] for s in sheets]
    numbered = [n for n in names if WB_NUMBERED_RE.match(n)]
    if numbered:
        bad.append(f"编号或默认页名 {len(numbered)} 张（{'、'.join(numbered[:5])}{'……' if len(numbered) > 5 else ''}）"
                   "——按内容起语义名，如“历年营收与业务拆分”“××数据来源”")
    ov = next((s for s in sheets if s["name"] == WB_OVERVIEW), None)
    visible = [s for s in sheets if not s["hidden"]]
    if ov is None:
        bad.append(f"没有“{WB_OVERVIEW}”工作表")
    else:
        if not visible or visible[0] is not ov:
            bad.append(f"第一张可见工作表是“{visible[0]['name'] if visible else '无'}”，应是“{WB_OVERVIEW}”")
        if ov["charts"] == 0:
            bad.append(f"“{WB_OVERVIEW}”上没有原生图表——总览是集中全部重要数据图的看板，只有图号索引不算")
        off = sorted(r for r in ov["refs"] if WB_CALC not in r)
        if off:
            bad.append(f"“{WB_OVERVIEW}”的图表引用了 {'、'.join(off[:3])}——总览图的数据区放在名含“{WB_CALC}”的工作表")
    if not any(WB_CALC in n for n in names):
        bad.append(f"没有名含“{WB_CALC}”的工作表（如“图表计算”，放总览图的公式数据区，可隐藏）")
    if not any(WB_SOURCE in n for n in names):
        bad.append(f"没有名含“{WB_SOURCE}”的来源页（来源 ID → 文件、链接、页码或字段）")
    bad += layout_problems(sheets)
    total = sum(s["charts"] for s in sheets)
    note = f"工作表 {len(sheets)} 张，总览原生图表 {ov['charts'] if ov else 0} 张，全簿 {total} 张"
    return bad, note


def _wb_version(name):
    m = WB_VERSION_RE.search(name)
    return m.group(0) if m else ""


def select_workbooks(items, args):
    """第 13 项要查的根目录 Google 表格；返回 (表格列表, 问题列表)。"""
    sheets = [i for i in items if len(i.parts) == 1 and not i.is_dir and i.kind == GSHEET]
    if args.workbook:
        names = {s.name for s in sheets}
        bad = [f"--workbook 指定的「{w}」不在根目录的 Google 表格里" for w in args.workbook if w not in names]
        return [s for s in sheets if s.name in args.workbook], bad
    wbs = [s for s in sheets if any(k in s.name for k in WB_KEYS)]
    if args.report:
        keep = []
        for r in args.report:
            obj, ver = r.split("_", 1)[0], _wb_version(r)
            keep += [w for w in wbs if w.name.split("_", 1)[0] == obj and _wb_version(w.name) in (ver, "")
                     and w not in keep]
        wbs = keep
    return wbs, []


def export_workbook(drive, item):
    """用 rclone 把 Drive 上的 Google 表格导出成 xlsx 字节；失败返回 None。"""
    path = "/".join(item.parts[:-1] + [item.name + ".xlsx"])
    try:
        r = subprocess.run(["rclone", "cat", f"{drive.rstrip('/')}/{path}", "--drive-export-formats", "xlsx"],
                           capture_output=True)
    except FileNotFoundError:
        return None
    return r.stdout if r.returncode == 0 and r.stdout[:2] == b"PK" else None


def check_workbooks(items, args):
    """第 13 项：图表工作簿结构（语义页名、图表总览看板、图表计算、来源页）。取不到工作簿 = 不通过。"""
    wbs, bad = select_workbooks(items, args)
    notes = []
    for w in wbs:
        data = None
        if args.workbook_dir:
            p = os.path.join(args.workbook_dir, w.name + ".xlsx")
            if os.path.isfile(p):
                with open(p, "rb") as f:
                    data = f.read()
        elif args.drive:
            data = export_workbook(args.drive, w)
        if data is None:
            bad.append(f"{w.name}：取不到工作簿（--drive 时用 rclone 导出 xlsx；离线用 --workbook-dir）")
            continue
        try:
            probs, note = workbook_problems(parse_workbook(data))
        except (zipfile.BadZipFile, KeyError, ET.ParseError) as e:
            bad.append(f"{w.name}：读不出工作簿结构（{e}）")
            continue
        bad += [f"{w.name}：{x}" for x in probs]
        notes.append(f"{w.name}：{note}")
    ok_note = f"{len(wbs)} 份；" + "；".join(notes) if wbs else "没有配套图表工作簿，不适用"
    return [("图表工作簿：语义页名、图表总览看板、图表计算、来源页、图表不重叠、无多余空行（硬性）", bad, ok_note)]


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
    src.add_argument("--drive", help='rclone 路径，如 "gdrive:<研究库>/10_工作区/20261104_任务名"')
    src.add_argument("--lsjson", help="离线：读已保存的 rclone lsjson -R -M 输出（JSON 文件）")
    ap.add_argument("--local", help="本地任务文件夹（可选）：另查 成品/ 与 留痕/复用、复核 是否齐")
    ap.add_argument("--report", action="append", default=[],
                    help="第 12 项只核对这份报告（根目录 Google 文档名，不带后缀；可多次）；不给则核对根目录全部 Google 文档")
    ap.add_argument("--toc-text", help="第 12 项离线：读这个目录里的 <文档名>.txt（Google 文档导出的纯文本）")
    ap.add_argument("--toc-evidence",
                    help="第 12 项人工确认证据文件：每行 文档名｜二级目录带页码：是｜YYYY-MM-DD｜确认人；"
                         "只在取不到文档文本时采用，缺证据即不通过")
    ap.add_argument("--workbook", action="append", default=[],
                    help="第 13 项只核对这份图表工作簿（根目录 Google 表格名，不带后缀；可多次）；"
                         "不给时按 --report 的研究对象与版本号挑，都不给则查全部名含“图表数据”“历史数据”的表格")
    ap.add_argument("--workbook-dir", help="第 13 项离线：读这个目录里的 <表格名>.xlsx（Google 表格导出的 xlsx）")
    args = ap.parse_args(argv)

    items = load_items(args)
    print(f"交付核对：{args.drive or args.lsjson}（{sum(not i.is_dir for i in items)} 个文件）")
    groups = [("Drive", check_drive(items) + check_toc(items, args) + check_workbooks(items, args))]
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
