#!/usr/bin/env python3
"""图表防乱检查：工作簿里每张原生图的排布、宽度、点数、横轴标签、标记点、系列数。

用法：
  python3 check_charts.py 工作簿.xlsx [工作簿2.xlsx ...]
  Google 表格先导出 xlsx：rclone cat "<路径>/<表格名>.xlsx" --drive-export-formats xlsx > 本地.xlsx，
  或 Drive API files.export（mimeType application/vnd.openxmlformats-officedocument.spreadsheetml.sheet）。

逐张图检查（research-delivery C3"图表防乱"，硬性）：
  G1 一行最多两张图：上下范围重叠一半以上的图算同一行，同一行超过 2 张不通过。
  G2 图与图不重叠。
  G3 宽度：每张图宽 ≥ 560 px；点数 > 200 的图宽 ≥ 1100 px（独占一行）。
  G4 点数 ≤ 图宽 px ÷ 3：日数据先改周末、月末或季末，直到满足。
  G5 横轴标签：折线、面积图非空标签 ≤ 15 个，柱图 ≤ 20 个；长序列的标签列只在每 N 年首期写年份，其余留空。
  G6 折线无标记点：折线系列的 marker 是 none；系列没写 marker 时看折线图本身的 marker 开关。
  G7 系列数 ≤ 3；"图表总览"页的交互图 ≤ 10（Excel 规范 2.5）。
  G8 图内标题不带图号（"图 3 …"）；图号只在报告图注。
点数取系列引用范围里的非空数值；公式没有缓存值时按范围行数估，并在输出注明（Google 导出的 xlsx 有缓存值）。
宽高按 xlsx 里的锚点算：oneCellAnchor / absoluteAnchor 用 ext，twoCellAnchor 用列宽行高换算（列宽 px ≈ 宽度 × 7 + 5）。

输出：每张图一行（工作表 ｜ 第几张 ｜ 标题 ｜ 宽×高 ｜ 点数 ｜ 标签数 ｜ 系列数 ｜ 通过 / 问题），最后一行"结果：x/y 张图通过"。
全部通过退出码 0，有不通过 1，参数或文件错误 2。只读，不改文件。依赖 openpyxl（读单元格值）。
"""
import posixpath
import re
import sys
import zipfile
from xml.etree import ElementTree as ET

EMU_PER_PX = 9525
REL_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
MIN_W, WIDE_W, WIDE_POINTS, PX_PER_POINT = 560, 1100, 200, 3
MAX_LABELS_LINE, MAX_LABELS_BAR = 15, 20
MAX_SERIES, MAX_SERIES_OVERVIEW = 3, 10


def local(tag):
    return tag.rsplit("}", 1)[-1]


def rels(z, part):
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


def child(el, name):
    for c in el:
        if local(c.tag) == name:
            return c
    return None


def sheet_geometry(z, part):
    """列宽、行高（px）与默认值。"""
    root = ET.fromstring(z.read(part))
    cols, rows = {}, {}
    def_col_px, def_row_px = 64.0, 20.0
    for el in root.iter():
        t = local(el.tag)
        if t == "sheetFormatPr":
            if el.get("defaultColWidth"):
                def_col_px = float(el.get("defaultColWidth")) * 7 + 5
            if el.get("defaultRowHeight"):
                def_row_px = float(el.get("defaultRowHeight")) * 96 / 72
        elif t == "col" and el.get("width"):
            for c in range(int(el.get("min")), int(el.get("max")) + 1):
                cols[c - 1] = float(el.get("width")) * 7 + 5
        elif t == "row" and el.get("ht"):
            rows[int(el.get("r")) - 1] = float(el.get("ht")) * 96 / 72

    def x(col, off=0):
        return sum(cols.get(i, def_col_px) for i in range(col)) + off / EMU_PER_PX

    def y(row, off=0):
        return sum(rows.get(i, def_row_px) for i in range(row)) + off / EMU_PER_PX
    return x, y


def marker_pos(el):
    v = {local(c.tag): int(c.text or 0) for c in el}
    return v.get("col", 0), v.get("colOff", 0), v.get("row", 0), v.get("rowOff", 0)


def anchors(z, drawing, x, y):
    """[(chart 部件路径, left, top, width, height)]"""
    out = []
    rmap = rels(z, drawing)
    for a in ET.fromstring(z.read(drawing)):
        kind = local(a.tag)
        cid = None
        for el in a.iter():
            if local(el.tag) == "chart" and el.get(REL_NS):
                cid = el.get(REL_NS)
        if not cid or cid not in rmap:
            continue
        chart = rmap[cid][1]
        if kind == "oneCellAnchor":
            c, co, r, ro = marker_pos(child(a, "from"))
            ext = child(a, "ext")
            left, top = x(c, co), y(r, ro)
            w, h = int(ext.get("cx")) / EMU_PER_PX, int(ext.get("cy")) / EMU_PER_PX
        elif kind == "twoCellAnchor":
            c1, co1, r1, ro1 = marker_pos(child(a, "from"))
            c2, co2, r2, ro2 = marker_pos(child(a, "to"))
            left, top = x(c1, co1), y(r1, ro1)
            w, h = x(c2, co2) - left, y(r2, ro2) - top
        elif kind == "absoluteAnchor":
            pos, ext = child(a, "pos"), child(a, "ext")
            left, top = int(pos.get("x")) / EMU_PER_PX, int(pos.get("y")) / EMU_PER_PX
            w, h = int(ext.get("cx")) / EMU_PER_PX, int(ext.get("cy")) / EMU_PER_PX
        else:
            continue
        out.append((chart, left, top, w, h))
    return out


REF_RE = re.compile(r"^\s*(?:'((?:[^']|'')+)'|([^!'\s]+))!\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?\s*$")


def ref_values(wb, ref):
    """引用 → (值列表, 是否全是没有缓存的公式)。"""
    m = REF_RE.match(ref or "")
    if not m:
        return [], False
    sheet = (m.group(1) or m.group(2)).replace("''", "'")
    if sheet not in wb.sheetnames:
        return [], False
    a = f"{m.group(3)}{m.group(4)}"
    b = f"{m.group(5) or m.group(3)}{m.group(6) or m.group(4)}"
    vals = [c.value for row in wb[sheet][a:b] for c in row]
    return vals, bool(vals) and all(v is None for v in vals)


def first_text(el):
    if el is None:
        return ""
    return "".join(t.text or "" for t in el.iter() if local(t.tag) == "t").strip()


def parse_chart(z, part, wb):
    root = ET.fromstring(z.read(part))
    plot = next((e for e in root.iter() if local(e.tag) == "plotArea"), None)
    title_el = next((e for e in root.iter() if local(e.tag) == "title"), None)
    info = {"title": first_text(title_el), "kinds": set(), "series": 0, "points": 0, "labels": 0,
            "estimated": False, "markers": 0}
    if plot is None:
        return info
    for grp in plot:
        kind = local(grp.tag)
        if not kind.endswith("Chart"):
            continue
        info["kinds"].add(kind)
        for ser in grp:
            if local(ser.tag) != "ser":
                continue
            info["series"] += 1
            for sub in ser:
                name = local(sub.tag)
                f = next((e.text for e in sub.iter() if local(e.tag) == "f"), None)
                if name in ("val", "yVal") and f:
                    vals, no_cache = ref_values(wb, f)
                    n = len(vals) if no_cache else sum(isinstance(v, (int, float)) for v in vals)
                    info["estimated"] |= no_cache
                    info["points"] = max(info["points"], n)
                elif name in ("cat", "xVal") and f:
                    vals, _ = ref_values(wb, f)
                    info["labels"] = max(info["labels"], sum(v is not None and str(v).strip() != "" for v in vals))
            if kind == "lineChart":
                mk = child(ser, "marker")
                sym = child(mk, "symbol") if mk is not None else None
                grp_mk = child(grp, "marker")
                grp_on = grp_mk is not None and grp_mk.get("val", "1") in ("1", "true")
                if (sym is not None and sym.get("val") != "none") or (sym is None and grp_on):
                    info["markers"] += 1
    return info


def check(path):
    try:
        from openpyxl import load_workbook
    except ImportError:
        sys.exit("缺少 openpyxl：pip install openpyxl --break-system-packages")
    wb = load_workbook(path, data_only=True)
    z = zipfile.ZipFile(path)
    wb_rels = rels(z, "xl/workbook.xml")
    rows_out, ok_n, total = [], 0, 0
    for el in ET.fromstring(z.read("xl/workbook.xml")).iter():
        if local(el.tag) != "sheet":
            continue
        name = el.get("name", "")
        part = wb_rels.get(el.get(REL_NS), ("", ""))[1]
        if part not in z.namelist():
            continue
        x, y = sheet_geometry(z, part)
        charts = []
        for typ, drawing in rels(z, part).values():
            if typ == "drawing" and drawing in z.namelist():
                charts += anchors(z, drawing, x, y)
        charts.sort(key=lambda c: (round(c[2]), c[1]))
        for i, (cpart, left, top, w, h) in enumerate(charts, 1):
            total += 1
            info = parse_chart(z, cpart, wb)
            probs = []
            same_row = [c for c in charts
                        if min(top + h, c[2] + c[4]) - max(top, c[2]) > 0.5 * min(h, c[4])]
            if len(same_row) > 2:
                probs.append(f"G1 这一行有 {len(same_row)} 张图，最多 2 张")
            for j, c in enumerate(charts, 1):
                if j != i and (min(left + w, c[1] + c[3]) - max(left, c[1]) > 2) and (min(top + h, c[2] + c[4]) - max(top, c[2]) > 2):
                    probs.append(f"G2 与第 {j} 张图重叠")
            if w < MIN_W:
                probs.append(f"G3 宽 {w:.0f} px < {MIN_W}")
            if info["points"] > WIDE_POINTS and w < WIDE_W:
                probs.append(f"G3 点数 {info['points']} > {WIDE_POINTS}，宽须 ≥ {WIDE_W} px 独占一行（现 {w:.0f}）")
            if w > 0 and info["points"] > w / PX_PER_POINT:
                probs.append(f"G4 点数 {info['points']} > 图宽 ÷ {PX_PER_POINT} = {int(w / PX_PER_POINT)}，改用周末 / 月末 / 季末数据")
            is_bar = info["kinds"] and info["kinds"] <= {"barChart", "bar3DChart"}
            cap = MAX_LABELS_BAR if is_bar else MAX_LABELS_LINE
            if info["labels"] > cap:
                probs.append(f"G5 横轴非空标签 {info['labels']} 个 > {cap}，标签列只在每 N 年首期写年份")
            if info["markers"]:
                probs.append(f"G6 {info['markers']} 个折线系列带标记点")
            if re.match(r"^\s*(图|fig|chart)\s*\d", info["title"], re.I):
                probs.append("G8 图内标题带图号；图号只写在报告图注里，图内标题写内容与口径")
            cap_s = MAX_SERIES_OVERVIEW if name == "图表总览" else MAX_SERIES
            if info["series"] > cap_s:
                probs.append(f"G7 系列 {info['series']} 个 > {cap_s}")
            est = "（公式无缓存，按范围估）" if info["estimated"] else ""
            line = (f"{name} ｜ 第 {i} 张 ｜ {info['title'][:30] or '（无标题）'} ｜ {w:.0f}×{h:.0f} px ｜ 点数 {info['points']}{est}"
                    f" ｜ 标签 {info['labels']} ｜ 系列 {info['series']} ｜ " + ("通过" if not probs else "；".join(probs)))
            rows_out.append(line)
            ok_n += not probs
    return rows_out, ok_n, total


def main(argv=None):
    paths = (argv if argv is not None else sys.argv[1:])
    if not paths or any(p.startswith("-") for p in paths):
        print(__doc__)
        sys.exit(2)
    all_ok, all_total = 0, 0
    for p in paths:
        try:
            rows, ok_n, total = check(p)
        except (OSError, zipfile.BadZipFile, KeyError) as e:
            print(f"读不了 {p}：{e}", file=sys.stderr)
            sys.exit(2)
        print(f"== {p}：{total} 张图")
        for r in rows:
            print("  " + r)
        all_ok += ok_n
        all_total += total
    print(f"结果：{all_ok}/{all_total} 张图通过")
    sys.exit(0 if all_ok == all_total else 1)


if __name__ == "__main__":
    main()
