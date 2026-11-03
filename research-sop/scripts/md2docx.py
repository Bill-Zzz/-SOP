#!/usr/bin/env python3
"""把 research-sop 报告的 .md 转成 .docx（报告只写一遍：AI 只写 .md，.docx 由本脚本生成）。

用法：
  python3 md2docx.py 输入.md --out 输出.docx [--img-root 目录]
  python3 md2docx.py 输入.md --check-only [--img-root 目录]      只做检查，不输出 docx

支持：# ## ### 标题（→ Heading 1/2/3，更深的按 Heading 3）、普通段落、**粗体**、*斜体*、`代码`、[文字](链接)、
无序 / 有序列表（缩进两格以上为第二级）、管道表格（表头 + 对齐符行 → Word 真表格，样式 Table Grid，表头加粗，
单元格内 <br> 换行）、图片 ![图注](路径)（居中、宽度不超过 15 cm，图注另起一段）、> 引用段、``` 代码块。
水平线与 HTML 注释忽略；方括号来源标注如 [R01234 p5] 原样保留。

样式按 SKILL.md"字体与层级"：正文 Noto Sans SC 10.5 pt（东亚字体同为 Noto Sans SC）；图注、表注、来源 9 pt 深灰常规；
以"表注""来源""资料来源""数据与可编辑图表""图注"开头的段落自动套 9 pt 深灰样式。

检查（每次都做，结果打印到 stdout）：标题、表格、图片数量；表格行的列数与表头不一致的；图下是否注明
"数据与可编辑图表"。.md 里有以 ![ 开头、但文件不存在的图片时，列出缺的路径并以非零码退出（不输出 docx）。

只依赖 python-docx 与标准库（pip install python-docx --break-system-packages）；不调用 pandoc，不联网。
"""
import argparse
import os
import re
import sys

try:
    from docx import Document
    from docx.enum.style import WD_STYLE_TYPE
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.image.image import Image as DocxImage
    from docx.opc.constants import RELATIONSHIP_TYPE as RT
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Cm, Pt, RGBColor
except ImportError:  # pragma: no cover
    sys.exit("缺少 python-docx：pip install python-docx --break-system-packages")

FONT = "Noto Sans SC"
BODY_PT = 10.5
NOTE_PT = 9
NOTE_COLOR = RGBColor(0x59, 0x59, 0x59)
NOTE_STYLE = "注释小字"
MAX_IMG_CM = 15.0
NOTE_PREFIX = ("表注", "来源", "资料来源", "数据与可编辑图表", "图注")

RE_HEAD = re.compile(r"^(#{1,6})\s+(.*?)\s*#*\s*$")
RE_IMG = re.compile(r"^!\[(?P<alt>[^\]]*)\]\((?P<path><[^>]+>|[^)\s]+)(?:\s+\"[^\"]*\")?\)\s*$")
RE_UL = re.compile(r"^(?P<ind>\s*)[-*+]\s+(?P<txt>.*)$")
RE_OL = re.compile(r"^(?P<ind>\s*)\d+[.)]\s+(?P<txt>.*)$")
RE_HR = re.compile(r"^\s*([-*_])(\s*\1){2,}\s*$")
RE_SEP = re.compile(r"^\s*\|?\s*:?-{2,}:?\s*(\|\s*:?-{2,}:?\s*)*\|?\s*$")
RE_INLINE = re.compile(r"\*\*(?P<b>.+?)\*\*|(?<![\*\w])\*(?P<i>[^*\s][^*]*?)\*(?!\*)|`(?P<c>[^`]+)`"
                       r"|(?<!!)\[(?P<lt>[^\]]+)\]\((?P<lu>[^)\s]+)\)")
CJK = re.compile(r"[⺀-鿿＀-￯　-〿]")


# ---------------------------------------------------------------- 解析
def split_row(line):
    s = line.strip()
    if s.startswith("|"):
        s = s[1:]
    if s.endswith("|") and not s.endswith("\\|"):
        s = s[:-1]
    cells, cur, i = [], "", 0
    while i < len(s):
        if s[i] == "\\" and i + 1 < len(s) and s[i + 1] == "|":
            cur += "|"
            i += 2
            continue
        if s[i] == "|":
            cells.append(cur.strip())
            cur = ""
        else:
            cur += s[i]
        i += 1
    cells.append(cur.strip())
    return cells


def join_lines(a, b):
    if not a:
        return b
    if CJK.search(a[-1:]) or CJK.search(b[:1]):
        return a + b
    return a + " " + b


def parse(md):
    """返回块列表：(kind, data)。"""
    lines = md.splitlines()
    blocks, para, i = [], [], 0

    def flush():
        if para:
            text = ""
            for ln in para:
                text = join_lines(text, ln.strip())
            blocks.append(("p", text))
            para.clear()

    in_comment = False
    while i < len(lines):
        ln = lines[i]
        st = ln.strip()
        if in_comment:
            if "-->" in st:
                in_comment = False
            i += 1
            continue
        if st.startswith("<!--"):
            flush()
            if "-->" not in st:
                in_comment = True
            i += 1
            continue
        if st.startswith("```"):
            flush()
            code = []
            i += 1
            while i < len(lines) and not lines[i].strip().startswith("```"):
                code.append(lines[i])
                i += 1
            blocks.append(("code", code))
            i += 1
            continue
        if not st:
            flush()
            i += 1
            continue
        m = RE_HEAD.match(st)
        if m:
            flush()
            blocks.append(("h", (len(m.group(1)), m.group(2))))
            i += 1
            continue
        if RE_HR.match(st):
            flush()
            i += 1
            continue
        if st.startswith("|") and i + 1 < len(lines) and RE_SEP.match(lines[i + 1]):
            flush()
            head = split_row(st)
            aligns = []
            for c in split_row(lines[i + 1]):
                c = c.strip()
                aligns.append("center" if c.startswith(":") and c.endswith(":") else
                              "right" if c.endswith(":") else "left")
            rows, line_no = [], i + 1
            i += 2
            while i < len(lines) and lines[i].strip().startswith("|"):
                rows.append((i + 1, split_row(lines[i])))
                i += 1
            blocks.append(("table", {"head": head, "aligns": aligns, "rows": rows, "line": line_no}))
            continue
        m = RE_IMG.match(st)
        if m:
            flush()
            blocks.append(("img", {"alt": m.group("alt").strip(), "path": m.group("path").strip("<>"), "line": i + 1}))
            i += 1
            continue
        if st.startswith(">"):
            flush()
            q = []
            while i < len(lines) and lines[i].strip().startswith(">"):
                q.append(lines[i].strip()[1:].strip())
                i += 1
            text = ""
            for x in q:
                text = join_lines(text, x)
            blocks.append(("quote", text))
            continue
        m = RE_UL.match(ln) or RE_OL.match(ln)
        if m:
            flush()
            kind = "ul" if RE_UL.match(ln) else "ol"
            level = 2 if len(m.group("ind").expandtabs(4)) >= 2 else 1
            text = m.group("txt")
            i += 1
            # 列表项的续行（缩进、非新项、非空）
            while i < len(lines) and lines[i].strip() and lines[i].startswith(("  ", "\t")) \
                    and not (RE_UL.match(lines[i]) or RE_OL.match(lines[i])):
                text = join_lines(text, lines[i].strip())
                i += 1
            blocks.append((kind, (level, text)))
            continue
        para.append(ln)
        i += 1
    flush()
    return blocks


# ---------------------------------------------------------------- 样式
def set_font(style_or_run_font_owner, size=None, color=None, bold=None):
    f = style_or_run_font_owner.font
    f.name = FONT
    if size is not None:
        f.size = Pt(size)
    if color is not None:
        f.color.rgb = color
    if bold is not None:
        f.bold = bold
    rpr = style_or_run_font_owner.element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.insert(0, rfonts)
    for a in ("w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme"):
        if rfonts.get(qn(a)) is not None:
            del rfonts.attrib[qn(a)]
    for a in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rfonts.set(qn(a), FONT)


def setup_styles(doc):
    st = doc.styles
    normal = st["Normal"]
    set_font(normal, BODY_PT)
    normal.paragraph_format.space_before = Pt(0)
    normal.paragraph_format.space_after = Pt(6)
    normal.paragraph_format.line_spacing = 1.3
    for name, size in (("Heading 1", 16), ("Heading 2", 13), ("Heading 3", 11.5), ("Title", 20)):
        try:
            s = st[name]
        except KeyError:
            continue
        set_font(s, size, RGBColor(0x1F, 0x1F, 0x1F), True)
        s.paragraph_format.space_before = Pt(12 if name != "Heading 3" else 8)
        s.paragraph_format.space_after = Pt(6)
    for name in ("List Bullet", "List Bullet 2", "List Number", "List Number 2", "Quote"):
        try:
            set_font(st[name], BODY_PT)
        except KeyError:
            pass
    if NOTE_STYLE not in [s.name for s in st]:
        note = st.add_style(NOTE_STYLE, WD_STYLE_TYPE.PARAGRAPH)
        note.base_style = normal
    note = st[NOTE_STYLE]
    set_font(note, NOTE_PT, NOTE_COLOR, False)
    note.font.italic = False
    note.paragraph_format.space_before = Pt(2)
    note.paragraph_format.space_after = Pt(8)
    note.paragraph_format.line_spacing = 1.2


def add_hyperlink(par, text, url):
    part = par.part
    r_id = part.relate_to(url, RT.HYPERLINK, is_external=True)
    h = OxmlElement("w:hyperlink")
    h.set(qn("r:id"), r_id)
    r = OxmlElement("w:r")
    rpr = OxmlElement("w:rPr")
    c = OxmlElement("w:color")
    c.set(qn("w:val"), "2F5597")
    u = OxmlElement("w:u")
    u.set(qn("w:val"), "single")
    rpr.append(c)
    rpr.append(u)
    r.append(rpr)
    t = OxmlElement("w:t")
    t.text = text
    t.set(qn("xml:space"), "preserve")
    r.append(t)
    h.append(r)
    par._p.append(h)


def add_inline(par, text, bold=False):
    """把行内 Markdown 写进段落：**粗体**、*斜体*、`代码`、[文字](链接)；其余原样（含 [R01234 p5]）。"""
    text = text.replace("\\*", "\x00").replace("\\_", "_").replace("\\[", "[").replace("\\]", "]")
    parts = re.split(r"<br\s*/?>", text, flags=re.I)
    for k, seg in enumerate(parts):
        if k:
            par.add_run().add_break()
        pos = 0
        for m in RE_INLINE.finditer(seg):
            if m.start() > pos:
                r = par.add_run(seg[pos:m.start()].replace("\x00", "*"))
                r.bold = bold or None
            if m.group("b") is not None:
                inner = m.group("b")
                # 粗体里允许斜体
                im = re.fullmatch(r"\*(.+)\*", inner)
                r = par.add_run((im.group(1) if im else inner).replace("\x00", "*"))
                r.bold = True
                if im:
                    r.italic = True
            elif m.group("i") is not None:
                r = par.add_run(m.group("i").replace("\x00", "*"))
                r.italic = True
                r.bold = bold or None
            elif m.group("c") is not None:
                r = par.add_run(m.group("c"))
                r.font.name = "Consolas"
            else:
                add_hyperlink(par, m.group("lt"), m.group("lu"))
            pos = m.end()
        if pos < len(seg):
            r = par.add_run(seg[pos:].replace("\x00", "*"))
            r.bold = bold or None


def plain(text):
    return re.sub(r"\*\*|(?<!\w)\*|\*(?!\w)", "", text).strip()


def is_note(text):
    return plain(text).startswith(NOTE_PREFIX)


def resolve(path, img_root):
    if os.path.isabs(path):
        return path
    return os.path.normpath(os.path.join(img_root, path))


def image_width_cm(path):
    try:
        im = DocxImage.from_file(path)
        cm = im.width / 360000.0  # EMU → cm（按图片自带 dpi）
        return min(cm, MAX_IMG_CM) if cm > 0 else MAX_IMG_CM
    except Exception:  # noqa: BLE001
        return MAX_IMG_CM


# ---------------------------------------------------------------- 检查与转换
def check(blocks, img_root):
    heads = [b for b in blocks if b[0] == "h"]
    tables = [b for b in blocks if b[0] == "table"]
    imgs = [b for b in blocks if b[0] == "img"]
    missing = [b[1]["path"] for b in imgs if not os.path.isfile(resolve(b[1]["path"], img_root))]
    bad_rows = []
    for _, t in tables:
        n = len(t["head"])
        for line_no, row in t["rows"]:
            if len(row) != n:
                bad_rows.append(f"第 {line_no} 行：{len(row)} 列，表头 {n} 列")
    no_src = []
    for k, b in enumerate(blocks):
        if b[0] != "img":
            continue
        nxt = [x for x in blocks[k + 1:k + 4] if x[0] in ("p", "quote")]
        if not any("数据与可编辑图表" in x[1] or plain(x[1]).startswith(("来源", "资料来源")) for x in nxt):
            no_src.append(b[1]["path"])
    lv = {n: sum(1 for b in heads if min(b[1][0], 3) == n) for n in (1, 2, 3)}
    return {"heads": len(heads), "lv": lv, "tables": len(tables), "imgs": len(imgs), "missing": missing,
            "bad_rows": bad_rows, "no_src": no_src}


def convert(blocks, img_root, out):
    doc = Document()
    setup_styles(doc)
    prev = None
    for kind, data in blocks:
        if kind == "h":
            level, text = data
            p = doc.add_heading(level=min(level, 3))
            add_inline(p, text)
        elif kind == "p":
            style = NOTE_STYLE if is_note(data) else None
            p = doc.add_paragraph(style=style)
            add_inline(p, data)
        elif kind in ("ul", "ol"):
            level, text = data
            base = "List Bullet" if kind == "ul" else "List Number"
            p = doc.add_paragraph(style=base + (" 2" if level == 2 else ""))
            add_inline(p, text)
        elif kind == "quote":
            p = doc.add_paragraph(style=NOTE_STYLE if is_note(data) else "Quote")
            add_inline(p, data)
        elif kind == "code":
            for ln in data or [""]:
                p = doc.add_paragraph()
                r = p.add_run(ln)
                r.font.name = "Consolas"
                r.font.size = Pt(9)
                p.paragraph_format.space_after = Pt(0)
        elif kind == "img":
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            path = resolve(data["path"], img_root)
            p.add_run().add_picture(path, width=Cm(image_width_cm(path)))
            p.paragraph_format.keep_with_next = True
            if data["alt"]:
                c = doc.add_paragraph(style=NOTE_STYLE)
                c.alignment = WD_ALIGN_PARAGRAPH.CENTER
                add_inline(c, data["alt"])
        elif kind == "table":
            if prev == "table":  # 两张表紧挨着时 Word 会把它们并成一张，中间隔一个空段
                doc.add_paragraph()
            n = len(data["head"])
            t = doc.add_table(rows=1, cols=n)
            t.style = doc.styles["Table Grid"]
            t.alignment = WD_TABLE_ALIGNMENT.CENTER
            al = {"left": WD_ALIGN_PARAGRAPH.LEFT, "center": WD_ALIGN_PARAGRAPH.CENTER,
                  "right": WD_ALIGN_PARAGRAPH.RIGHT}
            aligns = (data["aligns"] + ["left"] * n)[:n]
            for j, txt in enumerate(data["head"][:n]):
                cp = t.rows[0].cells[j].paragraphs[0]
                cp.alignment = al[aligns[j]]
                add_inline(cp, txt, bold=True)
            for _, row in data["rows"]:
                if len(row) > n:  # 多出的格并入最后一列（检查里已报告）
                    row = row[:n - 1] + [" | ".join(row[n - 1:])]
                cells = t.add_row().cells
                for j, txt in enumerate(row[:n]):
                    cp = cells[j].paragraphs[0]
                    cp.alignment = al[aligns[j]]
                    add_inline(cp, txt)
            for row in t.rows:
                for c in row.cells:
                    for cp in c.paragraphs:
                        cp.paragraph_format.space_after = Pt(0)
                        cp.paragraph_format.line_spacing = 1.15
        prev = kind
    doc.save(out)


def main(argv=None):
    ap = argparse.ArgumentParser(description="报告 .md → .docx（research-sop 样式），附检查")
    ap.add_argument("md")
    ap.add_argument("--out", help="输出 .docx 路径")
    ap.add_argument("--img-root", help="图片相对路径的根目录（默认为 .md 所在目录）")
    ap.add_argument("--check-only", action="store_true", help="只做检查，不输出 docx")
    a = ap.parse_args(argv)
    if not a.check_only and not a.out:
        ap.error("缺 --out（或用 --check-only）")
    with open(a.md, encoding="utf-8-sig") as f:
        md = f.read()
    img_root = a.img_root or os.path.dirname(os.path.abspath(a.md))
    blocks = parse(md)
    r = check(blocks, img_root)
    print(f"== md2docx 检查：{os.path.basename(a.md)} ==")
    print(f"标题 {r['heads']}（一级 {r['lv'][1]} / 二级 {r['lv'][2]} / 三级 {r['lv'][3]}）；表格 {r['tables']}；图片 {r['imgs']}")
    print(f"表格列数与表头不一致：{'；'.join(r['bad_rows']) if r['bad_rows'] else '无'}")
    print(f"图下注明'数据与可编辑图表'或来源：{r['imgs'] - len(r['no_src'])}/{r['imgs']}"
          + (f"（未注明：{'、'.join(r['no_src'])}）" if r["no_src"] else ""))
    if r["missing"]:
        print(f"错误：{len(r['missing'])} 张图片文件不存在（图片根目录 {img_root}）：", file=sys.stderr)
        for p in r["missing"]:
            print(f"  - {p}", file=sys.stderr)
        return 1
    if a.check_only:
        print("只检查，未输出 docx")
        return 0
    os.makedirs(os.path.dirname(os.path.abspath(a.out)), exist_ok=True)
    convert(blocks, img_root, a.out)
    print(f"已输出 {a.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
