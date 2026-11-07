#!/usr/bin/env python3
"""把 research-sop 报告的 .md 转成 .docx（报告只写一遍：AI 只写 .md，.docx 由本脚本生成），版式按格式规范 Part A。

用法：
  python3 md2docx.py 输入.md --out 输出.docx [--img-root 目录] [--title 报告名] [--ticker 代码] [--date YYYY-MM-DD] [--no-toc]
  python3 md2docx.py 输入.md --check-only [--img-root 目录]      只做检查，不输出 docx

标题层级：.md 只有一个 # 且它是第一个标题时，# 是报告标题（Title 样式），##、###、#### 依次为一、二、三级标题
（Heading 1/2/3）；否则 #、##、### 直接为一、二、三级。"目录"与"一页摘要"（第一个正文章之前含"摘要"的一级标题）
用独立样式，不进目录、不计编号；"一、结论与概述"是正文第一章，照常编号；以"附"开头的是附录，照常进目录、另起一页，
按自己的编号体系。含"修订说明 / 修订记录"的仍另起一页，但"报告外壳"检查会报。

版式（standards/格式规范.md Part A）：
  页面    A4 竖版（全部页面，无横向页），四边 2 cm，单栏；页眉左报告名、右"代码 · 日期"（报告名取 --title 或报告标题；代码、日期取
          --ticker / --date，否则从报告标题、再从文件名里认），页脚居中"页码 / 总页数"（PAGE / NUMPAGES 域），
          页眉页脚 Noto Sans SC 8 pt。
  字体    中文、英文与数字统一 Noto Sans SC（格式规范第 3 节，research-sop 0.17.0；rFonts 的 eastAsia 与 ascii/hAnsi
          都写 Noto Sans SC，正文样式层统一配置，中文段另在 run 上显式写，防 Google 转换回退日文字体）；正文 10.5 pt，行距 1.0，段后 6 pt；一、二、三级标题 18 / 14 / 12 pt
          加粗黑色，段前段后 16/8、12/6、8/4 pt，行距 1.0。
  注释    以"表注""图注""来源""资料来源""数据与可编辑图表""口径说明""注："开头的段落与图片的图注：9 pt 灰色 #666666。
  顺序    报告标题 → 一页摘要 → 目录 → 一、结论与概述 → 各章 → 附：口径与说明；一页摘要紧接标题，目录、正文各一级
          板块、附录另起一页（段前分页）。
  目录    Word 目录域 TOC \\o "1-2" \\h \\z \\u \\n 1-1 \\p "　…　"，打开时更新：一级条目 10 pt 加粗、不显示页码；
          二级条目 10 pt 常规、左缩进 18 pt，"标题　…　页码"（页码灰色 #666666，不右对齐、不用点线）；行距 1.25，
          一级段前 10 pt、二级段后 5 pt。"研究正文 / 附录与资料"分组栏头与同名二级导航条目不由脚本生成（见修补清单 D1）。
  表格    真表格；表头底色 #17365D、白字 9.5 pt 加粗、水平垂直居中，"指标（单位）"把全角括号里的单位软换行到下一行、
          9 pt 常规；表体 9.5 pt、白底、段后 0、行距 1.0（说明文字同）、边框 #D9D9D9 0.5 pt、内边距 4 pt；数值、日期、
          编号、评级居中，文本左对齐，垂直顶端；首格为"合计 / 总计"的行加粗、加上边框；表宽 = 版心
          （表格缩进 4 pt = 单元格左内边距，外框正好落在版心左右边线上），
          列宽按最长实际值分配，叙述列吃剩余宽度；表题（表格前一段以"表 1"之类开头）10 pt 加粗；字号不缩小。
  体积    去掉模板里用不到的样式与部件（stylesWithEffects、主题、webSettings、字体表、缩略图、customXml），
          不含图片时约 8 KB。Google Drive MCP 的 create_file 只收 base64，内容要经 AI 自己的输出转写，
          实测 2 万字符的 base64 会转写出错；带多张图的报告应由用户上传或走 Drive 同步，见修补清单。

支持的 Markdown：标题、普通段落、**粗体**、*斜体*、`代码`、[文字](链接)、无序 / 有序列表（缩进两格以上为第二级）、
管道表格（单元格内 <br> 换行）、图片 ![图注](路径)（居中、宽度不超过 15 cm，图注另起一段）、> 引用段、``` 代码块。
水平线与 HTML 注释忽略；方括号来源标注如 [R01234 p5] 原样保留（但"出处标注"检查会报：正文应零标注）。

检查（每次都做，打印到 stdout）：标题、表格、图片数量；表格行的列数与表头不一致的；图下是否注明"数据与可编辑图表"；
图下来源行（只检查不改写：来源行或"数据与可编辑图表"行带 R / S 编号或页码的报出，来源行只写数据提供方与口径）；
标题编号（只检查不改写：一级"一、"连续、二级"1.1　"父级前缀与序号连续且后接全角空格、三级"1.1.1"，附录"附一、"连续；
报错号、跳号、重号、缺编号、缺全角空格）；表头单位（半角括号、数值列表头缺单位）；报告外壳（只检查不改写：报告标题与
"一页摘要"之间有段落、表格或列表；标题后第一个一级标题不是"一页摘要"，或一页摘要下没有内容、出现第二个摘要、"目录"
写在一页摘要之前；一页摘要之后的第一章不含"结论与概述"，或其下没有"结论""概述"两个子标题或以 **结论** / **概述**
开头的加粗段；最后一个一级标题不是"口径与说明"；"修订说明""本报告的取舍""来源"作为节标题——都报出）；出处标注与过程用语（只检查不改写：正文——不含图下"数据与可编辑图表"行与代码块——里的
"[R数字""[S数字""p数字]"和"研报库""台账""对账""取证"，报总数并按所在标题列出）；
图表优先（只提示，不影响退出码：一段或一条列表项里数字 ≥ 12 个、一张表里纯数值单元格 ≥ 40 个、正文章——不含"一、结论与概述"和附——没有图，
都列出，执行 AI 逐条改成图或删到关键节点；格式规范第 7 节 V02–V04）；图号与文件名（只检查不改写：图注"图 N"按出现顺序 1、2、3……严格递增，图片文件名里的 figN 与图号一致，
正文引用的"图 N"不超过图总数）。生成 .docx 后读回逐项核对格式
（页面、页眉页脚、字体字号行距、标题、注释、顺序与分页、目录样式、表格底色边框内边距对齐合计行列宽），打印"格式检查：x/y 通过"。
.md 里的图片文件不存在时列出并以非零码退出（不输出 docx）；格式检查有不通过项时退出码 1。

只依赖 python-docx 与标准库（pip install python-docx --break-system-packages）；不调用 pandoc，不联网。
"""
import argparse
import datetime as dt
import os
import re
import sys

try:
    from docx import Document
    from docx.enum.section import WD_ORIENT
    from docx.enum.style import WD_STYLE_TYPE
    from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
    from docx.image.image import Image as DocxImage
    from docx.opc.constants import RELATIONSHIP_TYPE as RT
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Cm, Pt, RGBColor, Twips
except ImportError:  # pragma: no cover
    sys.exit("缺少 python-docx：pip install python-docx --break-system-packages")

LATIN = "Noto Sans SC"   # 英文、拉丁字母与数字（格式规范 G02）
CJK_FONT = "Noto Sans SC" # 中文（格式规范 G01；0.17.0 起与英文数字统一，不再用 Arial：中英分设在 Google 文档里逐行行高不同）
BODY_PT = 10.5
TABLE_PT = 9.5           # 表头与表体（格式规范第 6 节）
NOTE_PT = 9
HF_PT = 8
CAPTION_PT = 10
BODY_LINE = 1.0          # 正文与表格文字单倍行距（格式规范 S01）
HEAD_LINE = 1.0
TEXT_CELL_LINE = 1.0     # 表体说明文字不再单设行距（V1.23）
TOC_LINE = 1.25          # 目录独立行距（S02）
TOC_SEP = "\u3000\u2026\u3000"   # 目录标题与页码之间的固定分隔符「　…　」（C06）
BLACK = RGBColor(0, 0, 0)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
NOTE_COLOR = RGBColor(0x66, 0x66, 0x66)
HEAD_FILL = "17365D"
BORDER = "D9D9D9"
BORDER_SZ = 4          # 1/8 pt：4 = 0.5 pt
TOTAL_BORDER = ("000000", 8)  # 合计行上边框：黑色 1 pt
CELL_MAR = 80          # twips：4 pt
PAGE_W, PAGE_H, MARGIN = 11906, 16838, 1134   # twips：A4、2 cm
TEXT_W = PAGE_W - 2 * MARGIN                  # 9638 twips = 17 cm
HEADS = {1: (18, 16, 8), 2: (14, 12, 6), 3: (12, 8, 4)}  # 字号、段前、段后（pt）
NOTE_STYLE = "注释小字"
CAPTION_STYLE = "表题"
ALONE_STYLE = "独立标题"
TH_STYLE = "表格表头"
TD_STYLE = "表格正文"
TDT_STYLE = "表格说明文字"
TOC_STYLES = ("toc 1", "toc 2")
MAX_IMG_CM = 15.0
NOTE_PREFIX = ("表注", "图注", "来源", "资料来源", "数据与可编辑图表", "口径说明", "注：", "注:")
TOTAL_PREFIX = ("合计", "总计")

RE_HEAD = re.compile(r"^(#{1,6})\s+(.*?)\s*#*\s*$")
RE_IMG = re.compile(r"^!\[(?P<alt>[^\]]*)\]\((?P<path><[^>]+>|[^)\s]+)(?:\s+\"[^\"]*\")?\)\s*$")
RE_UL = re.compile(r"^(?P<ind>\s*)[-*+]\s+(?P<txt>.*)$")
RE_OL = re.compile(r"^(?P<ind>\s*)\d+[.)]\s+(?P<txt>.*)$")
RE_HR = re.compile(r"^\s*([-*_])(\s*\1){2,}\s*$")
RE_SEP = re.compile(r"^\s*\|?\s*:?-{2,}:?\s*(\|\s*:?-{2,}:?\s*)*\|?\s*$")
RE_INLINE = re.compile(r"\*\*(?P<b>.+?)\*\*|(?<![\*\w])\*(?P<i>[^*\s][^*]*?)\*(?!\*)|`(?P<c>[^`]+)`"
                       r"|(?<!!)\[(?P<lt>[^\]]+)\]\((?P<lu>[^)\s]+)\)")
CJK = re.compile(r"[⺀-鿿＀-￯　-〿]")
RE_CJK_RUN = re.compile(r"([⺀-鿿豈-﫿︰-﹏＀-￯　-〿]+)")
RE_CAPTION = re.compile(r"^(表|Table)\s*[0-9A-Za-z一二三四五六七八九十]+[\s　：:．.、]")
RE_UNIT = re.compile(r"^(?P<ind>.*?\S)\s*(?P<unit>（[^（）]+）)\s*$")
RE_HALF_UNIT = re.compile(r"\([^()]*[一-鿿%$][^()]*\)\s*$")
CN_NUM = "零一二三四五六七八九十"
CN = "[一二三四五六七八九十]+"
RE_H1_NUM = re.compile(rf"^({CN})、")
RE_APPX = re.compile(rf"^附({CN})、")
RE_H2_NUM = re.compile(r"^(\d+)\.(\d+)(?![.\d])(.?)")
RE_H3_NUM = re.compile(r"^(\d+)\.(\d+)\.(\d+)(?![.\d])")
NUM = r"[-−+]?[$¥€£]?\(?[-−+]?\d[\d,]*(?:\.\d+)?\)?"
UNIT_WORDS = r"(?:%|％|x|X|倍|pp|bp|bps|个百分点|百分点|[万亿]?(?:美元|港元|人民币|元|股|台币|新台币|日元|欧元)|天|次|家|个|年|月|日|σ|小时|分钟)?"
RE_VALUE = re.compile(rf"""^(?:
    {NUM}\s*{UNIT_WORDS}(?:\s*/\s*{UNIT_WORDS})?
  | {NUM}\s*{UNIT_WORDS}\s*[-–~～至]\s*{NUM}\s*{UNIT_WORDS}
  | \d{{4}}(?:[-/.]\d{{1,2}}){{0,2}}
  | \d{{4}}年(?:\d{{1,2}}月(?:\d{{1,2}}日)?)?
  | FY\d{{2,4}}(?:Q[1-4])?[AE]? | [1-4]Q\d{{2}}[AE]? | \d{{4}}[AE] | Q[1-4] | H[12]
  | [RSFM]\d{{1,6}}(?:\s*p\.?\s*\d+)? | \[[RS]\d+[^\]]*\]
  | [-–—] | N/?A | n/?a | 不适用 | 待补 | 待核实
)$""", re.X)
VALUE_WORDS = {"高", "中", "低", "大", "小", "是", "否", "有", "无", "买入", "增持", "中性", "减持", "卖出", "跑赢",
               "跑输", "持有", "通过", "不通过", "偏乐观", "偏悲观", "一致", "有误", "查不到", "核心", "宏观", "财报", "点名"}


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


def plain(text):
    t = re.sub(r"\[([^\]]+)\]\([^)\s]+\)", r"\1", text)
    t = re.sub(r"<br\s*/?>", " ", t, flags=re.I)
    return re.sub(r"\*\*|(?<!\w)\*|\*(?!\w)|`", "", t).strip()


def is_note(text):
    return plain(text).startswith(NOTE_PREFIX)


def cn2int(s):
    if not s:
        return None
    if s == "十":
        return 10
    if "十" in s:
        a, _, b = s.partition("十")
        return (CN_NUM.index(a) if a else 1) * 10 + (CN_NUM.index(b) if b else 0)
    return CN_NUM.index(s) if len(s) == 1 else None


def int2cn(n):
    if n <= 10:
        return CN_NUM[n]
    if n < 20:
        return "十" + (CN_NUM[n - 10] if n > 10 else "")
    return CN_NUM[n // 10] + "十" + (CN_NUM[n % 10] if n % 10 else "")


def outline(blocks):
    """给每个标题定角色与层级。返回 (heads, title_mode)；heads 每项含 idx、md 层级、level（0 = 报告标题，1–3）、role、text。"""
    hs = [(k, b[1][0], plain(b[1][1])) for k, b in enumerate(blocks) if b[0] == "h"]
    title_mode = bool(hs) and hs[0][1] == 1 and sum(1 for h in hs if h[1] == 1) == 1
    out, seen_body = [], False
    for k, md_lv, text in hs:
        lv = (md_lv - 1) if title_mode else md_lv
        lv = min(max(lv, 0), 3)
        role = "title" if lv == 0 else "body"
        if lv == 1:
            if text.strip() in ("目录", "目 录"):
                role = "toc"
            elif "摘要" in text and not seen_body:
                role = "summary"
            elif text.startswith("附"):
                role = "appendix"
            elif "修订说明" in text or "修订记录" in text:
                role = "revision"
            else:
                role = "body"
            if role in ("body", "appendix", "revision"):
                seen_body = True
        out.append({"idx": k, "md": md_lv, "level": lv, "role": role, "text": text})
    # 二、三级跟随所属一级的角色
    cur = "summary"
    for h in out:
        if h["level"] == 1:
            cur = h["role"]
        elif h["level"] >= 2:
            h["role"] = "sub-" + cur
    return out, title_mode


def check_numbering(heads):
    """只检查不改写。返回问题列表与计数。"""
    issues = []
    n1 = n2 = n3 = 0
    exp1, exp_ap = 1, 1
    k1 = None   # 当前一级的数字
    exp2, k2, exp3 = 1, None, 1
    for h in heads:
        t, lv, role = h["text"], h["level"], h["role"]
        if lv == 1 and role == "body":
            n1 += 1
            m = RE_H1_NUM.match(t)
            if not m:
                issues.append(f"「{t}」：一级标题缺编号，应为「{int2cn(exp1)}、」")
                k1 = exp1
            else:
                k = cn2int(m.group(1))
                if k != exp1:
                    kind = "重号" if k is not None and k < exp1 else "跳号或错号"
                    issues.append(f"「{t}」：{kind}，应为「{int2cn(exp1)}、」，实为「{m.group(1)}、」")
                k1 = k if k is not None else exp1
            exp1 = k1 + 1
            exp2, k2 = 1, None
        elif lv == 1 and role == "appendix":
            k1 = None
            m = RE_APPX.match(t)
            if m:
                k = cn2int(m.group(1))
                if k != exp_ap:
                    issues.append(f"「{t}」：附录跳号或错号，应为「附{int2cn(exp_ap)}、」")
                exp_ap = (k or exp_ap) + 1
        elif lv == 1:
            k1 = None
        elif lv == 2 and role == "sub-body" and k1 is not None:
            n2 += 1
            m = RE_H2_NUM.match(t)
            if not m:
                issues.append(f"「{t}」：二级标题缺编号，应为「{k1}.{exp2}　」")
                k2 = exp2
            else:
                a, b, sep = int(m.group(1)), int(m.group(2)), m.group(3)
                if a != k1:
                    issues.append(f"「{t}」：父级前缀错，应为「{k1}.」")
                if b != exp2:
                    issues.append(f"「{t}」：{'重号' if b < exp2 else '跳号'}，应为「{k1}.{exp2}」")
                if sep != "　":
                    issues.append(f"「{t}」：编号与标题之间缺全角空格")
                k2 = b
            exp2 = k2 + 1
            exp3 = 1
        elif lv == 3 and role == "sub-body" and k1 is not None:
            n3 += 1
            m = RE_H3_NUM.match(t)
            if not m:
                issues.append(f"「{t}」：三级标题缺编号，应为「{k1}.{k2 or 1}.{exp3}」")
                exp3 += 1
            else:
                a, b, c = (int(x) for x in m.groups())
                if (a, b) != (k1, k2 or 1):
                    issues.append(f"「{t}」：父级前缀错，应为「{k1}.{k2 or 1}.」")
                if c != exp3:
                    issues.append(f"「{t}」：{'重号' if c < exp3 else '跳号'}，应为「{k1}.{k2 or 1}.{exp3}」")
                exp3 = c + 1
    return issues, (n1, n2, n3)


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


# ---------------------------------------------------------------- 表格：单元格分类与列宽
def is_value(text):
    t = plain(text).replace("　", " ").strip()
    if not t:
        return True
    return t in VALUE_WORDS or bool(RE_VALUE.match(t))


def text_width_pt(text, size, bold=False):
    w = 0.0
    for ch in text:
        if CJK.match(ch) or ord(ch) > 0x2E80:
            w += 1.0
        elif ch in " .,:;'|!il":
            w += 0.28
        elif ch.isdigit() or ch.isupper() or ch in "$%-+()#&":
            w += 0.62
        else:
            w += 0.52
    return w * size * (1.06 if bold else 1.0)


def split_head(text):
    """表头"指标（单位）" → (指标, "（单位）")；没有全角括号单位的返回 (原文, None)。"""
    t = plain(text)
    m = RE_UNIT.match(t)
    if not m:
        return t, None
    return m.group("ind"), m.group("unit")


def cell_need_pt(text, size, bold=False):
    lines = re.split(r"<br\s*/?>", text, flags=re.I)
    return max(text_width_pt(plain(x), size, bold) for x in lines) if lines else 0.0


def column_widths(head, rows):
    """按最长实际值分配列宽（twips），总宽 = 版心。返回 (widths, overflow)。"""
    n = len(head)
    pad = 2 * CELL_MAR / 20 + 3  # 左右内边距 + 余量（pt）
    need, is_text, longest = [], [], []
    for j in range(n):
        body = [r[j] for r in rows if j < len(r)]
        vals = [cell_need_pt(c, TABLE_PT) for c in body] or [0.0]
        ind, unit = split_head(head[j])
        txt = any(not is_value(c) for c in body)
        is_text.append(txt)
        longest.append(max(vals))
        # 完整数值单行显示；单位不拆行；指标名可以换行（至少容得下 3 个汉字）
        ind_need = min(text_width_pt(ind, TABLE_PT, True), max(max(vals), 3 * TABLE_PT * 1.06))
        need.append(max(max(vals), text_width_pt(unit or "", NOTE_PT), ind_need) + pad)
    total = TEXT_W / 20.0
    overflow = False
    narrative = [j for j in range(n) if is_text[j] and longest[j] > 6 * TABLE_PT]
    if sum(need) <= total:
        extra = total - sum(need)
        grow = narrative or list(range(n))
        base = sum(need[j] for j in grow)
        w = [need[j] + (extra * need[j] / base if j in grow else 0) for j in range(n)]
    else:
        data = [j for j in range(n) if not is_text[j]]
        texts = [j for j in range(n) if is_text[j]]
        rest = total - sum(need[j] for j in data)
        min_text = 4 * TABLE_PT + pad
        if texts and rest >= min_text * len(texts):
            base = sum(need[j] for j in texts)
            w = [need[j] if j in data else max(min_text, rest * need[j] / base) for j in range(n)]
            s = sum(w[j] for j in texts)
            w = [w[j] if j in data else w[j] * rest / s for j in range(n)]
        else:
            overflow = True
            w = [need[j] * total / sum(need) for j in range(n)]
    tw = [int(round(x * 20)) for x in w]
    tw[-1] += TEXT_W - sum(tw)
    return tw, overflow, is_text


# ---------------------------------------------------------------- 样式
def set_fonts(owner, size=None, color=None, bold=None, italic=None, latin=LATIN, east=CJK_FONT):
    f = owner.font
    f.name = latin
    if size is not None:
        f.size = Pt(size)
    if color is not None:
        f.color.rgb = color
        c = owner.element.rPr.find(qn("w:color"))
        for a in ("w:themeColor", "w:themeShade", "w:themeTint"):
            if c is not None and c.get(qn(a)) is not None:
                del c.attrib[qn(a)]
    if bold is not None:
        f.bold = bold
    if italic is not None:
        f.italic = italic
    rpr = owner.element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.insert(0, rfonts)
    for a in ("w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme"):
        if rfonts.get(qn(a)) is not None:
            del rfonts.attrib[qn(a)]
    for a in ("w:ascii", "w:hAnsi", "w:cs"):
        rfonts.set(qn(a), latin)
    rfonts.set(qn("w:eastAsia"), east)


def para_fmt(style, before=0, after=6, line=BODY_LINE, keep_next=None, align=None):
    pf = style.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.line_spacing = line
    if keep_next is not None:
        pf.keep_with_next = keep_next
    if align is not None:
        pf.alignment = align
    ppr = style.element.get_or_add_pPr()
    for tag in ("w:pBdr", "w:contextualSpacing"):
        e = ppr.find(qn(tag))
        if e is not None:
            ppr.remove(e)


def get_style(doc, name, base="Normal", kind=WD_STYLE_TYPE.PARAGRAPH):
    st = doc.styles
    try:
        return st[name]
    except KeyError:
        s = st.add_style(name, kind)
        if base:
            s.base_style = st[base]
        return s


def setup_styles(doc):
    st = doc.styles
    # 文档默认：字体、字号、语言（中文字形）
    rpr_def = st.element.find(qn("w:docDefaults")).find(qn("w:rPrDefault")).find(qn("w:rPr"))
    rf = rpr_def.find(qn("w:rFonts"))
    for a in list(rf.attrib):
        del rf.attrib[a]
    for a in ("w:ascii", "w:hAnsi", "w:cs"):
        rf.set(qn(a), LATIN)
    rf.set(qn("w:eastAsia"), CJK_FONT)
    for tag, val in (("w:sz", "21"), ("w:szCs", "21")):
        e = rpr_def.find(qn(tag))
        if e is None:
            e = OxmlElement(tag)
            rpr_def.append(e)
        e.set(qn("w:val"), val)
    lang = rpr_def.find(qn("w:lang"))
    if lang is None:
        lang = OxmlElement("w:lang")
        rpr_def.append(lang)
    lang.set(qn("w:val"), "en-US")
    lang.set(qn("w:eastAsia"), "zh-CN")

    normal = st["Normal"]
    set_fonts(normal, BODY_PT, BLACK)
    para_fmt(normal, 0, 6, BODY_LINE)
    for lv, (size, before, after) in HEADS.items():
        s = st[f"Heading {lv}"]
        set_fonts(s, size, BLACK, True, False)
        para_fmt(s, before, after, HEAD_LINE, keep_next=True)
    title = st["Title"]
    set_fonts(title, HEADS[1][0], BLACK, True, False)
    rpr = title.element.rPr
    for tag in ("w:spacing", "w:kern"):
        e = rpr.find(qn(tag))
        if e is not None:
            rpr.remove(e)
    para_fmt(title, 0, HEADS[1][2], HEAD_LINE, keep_next=True)
    alone = get_style(doc, ALONE_STYLE)
    set_fonts(alone, HEADS[1][0], BLACK, True, False)
    para_fmt(alone, HEADS[1][1], HEADS[1][2], HEAD_LINE, keep_next=True)
    for name in ("List Bullet", "List Bullet 2", "List Number", "List Number 2"):
        set_fonts(st[name], BODY_PT, BLACK)
        para_fmt(st[name], 0, 6, BODY_LINE)
    q = st["Quote"]
    set_fonts(q, BODY_PT, BLACK, None, False)
    para_fmt(q, 0, 6, BODY_LINE)
    q.paragraph_format.left_indent = Pt(18)
    note = get_style(doc, NOTE_STYLE)
    set_fonts(note, NOTE_PT, NOTE_COLOR, False, False)
    para_fmt(note, 0, 6, BODY_LINE)
    cap = get_style(doc, CAPTION_STYLE)
    set_fonts(cap, CAPTION_PT, BLACK, True, False)
    para_fmt(cap, 0, 4, BODY_LINE, keep_next=True)
    th = get_style(doc, TH_STYLE)
    set_fonts(th, TABLE_PT, WHITE, True, False)
    para_fmt(th, 0, 0, HEAD_LINE, align=WD_ALIGN_PARAGRAPH.CENTER)
    td = get_style(doc, TD_STYLE)
    set_fonts(td, TABLE_PT, BLACK, None, False)
    para_fmt(td, 0, 0, BODY_LINE)
    tdt = get_style(doc, TDT_STYLE, TD_STYLE)
    para_fmt(tdt, 0, 0, TEXT_CELL_LINE)
    for name in ("Header", "Footer"):
        s = st[name]
        set_fonts(s, HF_PT, BLACK, False, False)
        para_fmt(s, 0, 0, 1.0)
        ppr = s.element.get_or_add_pPr()
        tabs = ppr.find(qn("w:tabs"))
        if tabs is not None:
            ppr.remove(tabs)
    for k, name in enumerate(TOC_STYLES):
        s = get_style(doc, name)
        set_fonts(s, CAPTION_PT, BLACK, k == 0, False)
        para_fmt(s, 10 if k == 0 else 0, 0 if k == 0 else 5, TOC_LINE)
        s.paragraph_format.left_indent = Pt(0 if k == 0 else 18)
        ppr = s.element.get_or_add_pPr()  # 不用右对齐制表位与点线（C06、C07）
        tabs = ppr.find(qn("w:tabs"))
        if tabs is not None:
            ppr.remove(tabs)


# ---------------------------------------------------------------- 段内元素
def set_run_font(rpr_owner_el, name):
    rpr = rpr_owner_el.get_or_add_rPr() if hasattr(rpr_owner_el, "get_or_add_rPr") else rpr_owner_el
    rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts")
        rpr.insert(0, rf)
    for a in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rf.set(qn(a), name)


def script_runs(par, text, size=None):
    """按文字分段写 run：中文段（汉字、全角标点）显式写 CJK_FONT，西文与数字段继承样式（LATIN）；0.17.0 起两者都是 Noto Sans SC。
    Google 文档转换时只认 run 的西文字体、不认东亚字体，中文会回退到 MS PGothic 一类日文字体并丢掉粗体
    （2026-11-04 实测），所以中文段在 run 上再写一次同一字体。中英统一一个字体族，Google 里每行行高一致。"""
    out = []
    for piece in RE_CJK_RUN.split(text):
        if not piece:
            continue
        r = par.add_run(piece)
        if RE_CJK_RUN.fullmatch(piece):
            set_run_font(r._r, CJK_FONT)
        if size:
            r.font.size = Pt(size)
        out.append(r)
    return out


def code_runs(par, text, size=None):
    """代码段：西文用 Consolas，中文段同样显式中文字体 CJK_FONT。"""
    for piece in RE_CJK_RUN.split(text):
        if not piece:
            continue
        r = par.add_run(piece)
        if RE_CJK_RUN.fullmatch(piece):
            set_run_font(r._r, CJK_FONT)
        else:
            r.font.name = "Consolas"
        if size:
            r.font.size = Pt(size)


def add_hyperlink(par, text, url=None, anchor=None):
    h = OxmlElement("w:hyperlink")
    if url:
        h.set(qn("r:id"), par.part.relate_to(url, RT.HYPERLINK, is_external=True))
    if anchor:
        h.set(qn("w:anchor"), anchor)
        h.set(qn("w:history"), "1")
    for piece in RE_CJK_RUN.split(text):
        if not piece:
            continue
        r = OxmlElement("w:r")
        rpr = OxmlElement("w:rPr")
        if RE_CJK_RUN.fullmatch(piece):
            set_run_font(rpr, CJK_FONT)
        if url:
            c = OxmlElement("w:color")
            c.set(qn("w:val"), "2F5597")
            u = OxmlElement("w:u")
            u.set(qn("w:val"), "single")
            rpr.append(c)
            rpr.append(u)
        if len(rpr):
            r.append(rpr)
        t = OxmlElement("w:t")
        t.text = piece
        t.set(qn("xml:space"), "preserve")
        r.append(t)
        h.append(r)
    par._p.append(h)
    return h


def add_inline(par, text, bold=False, size=None):
    """把行内 Markdown 写进段落：**粗体**、*斜体*、`代码`、[文字](链接)；其余原样（含 [R01234 p5]）。"""
    text = text.replace("\\*", "\x00").replace("\\_", "_").replace("\\[", "[").replace("\\]", "]")
    parts = re.split(r"<br\s*/?>", text, flags=re.I)

    def run(t, b=None, i=None):
        for r in script_runs(par, t.replace("\x00", "*"), size):
            r.bold = b
            if i:
                r.italic = True

    for k, seg in enumerate(parts):
        if k:
            par.add_run().add_break()
        pos = 0
        for m in RE_INLINE.finditer(seg):
            if m.start() > pos:
                run(seg[pos:m.start()], bold or None)
            if m.group("b") is not None:
                inner = m.group("b")
                im = re.fullmatch(r"\*(.+)\*", inner)  # 粗体里允许斜体
                run(im.group(1) if im else inner, True, bool(im))
            elif m.group("i") is not None:
                run(m.group("i"), bold or None, True)
            elif m.group("c") is not None:
                code_runs(par, m.group("c"), size)
            else:
                add_hyperlink(par, m.group("lt"), url=m.group("lu"))
            pos = m.end()
        if pos < len(seg):
            run(seg[pos:], bold or None)


def fld_run(kind, text=None):
    r = OxmlElement("w:r")
    if kind == "instr":
        t = OxmlElement("w:instrText")
        t.set(qn("xml:space"), "preserve")
        t.text = f" {text} "
        r.append(t)
    elif kind == "text":
        t = OxmlElement("w:t")
        t.set(qn("xml:space"), "preserve")
        t.text = text
        r.append(t)
    else:
        fc = OxmlElement("w:fldChar")
        fc.set(qn("w:fldCharType"), kind)
        r.append(fc)
    return r


def add_field(p_el, instr, cached=""):
    for r in (fld_run("begin"), fld_run("instr", instr), fld_run("separate"), fld_run("text", cached), fld_run("end")):
        p_el.append(r)


def add_bookmark(par, name, bid):
    s = OxmlElement("w:bookmarkStart")
    s.set(qn("w:id"), str(bid))
    s.set(qn("w:name"), name)
    e = OxmlElement("w:bookmarkEnd")
    e.set(qn("w:id"), str(bid))
    ppr = par._p.pPr
    if ppr is not None:
        ppr.addnext(s)
    else:
        par._p.insert(0, s)
    par._p.append(e)


def tc_prop(cell, tag, attrs=None, children=None):
    tcpr = cell._tc.get_or_add_tcPr()
    e = tcpr.find(qn(tag))
    if e is None:
        e = OxmlElement(tag)
        # tcPr 子元素有固定顺序：tcW → gridSpan → … → tcBorders → shd → … → vAlign
        order = ["w:tcW", "w:gridSpan", "w:vMerge", "w:tcBorders", "w:shd", "w:noWrap", "w:tcMar", "w:vAlign"]
        later = [tcpr.find(qn(t)) for t in order[order.index(tag) + 1:]] if tag in order else []
        later = [x for x in later if x is not None]
        if later:
            later[0].addprevious(e)
        else:
            tcpr.append(e)
    for k, v in (attrs or {}).items():
        e.set(qn(k), v)
    for c in children or []:
        e.append(c)
    return e


def border_el(tag, color, sz):
    b = OxmlElement(tag)
    b.set(qn("w:val"), "single")
    b.set(qn("w:sz"), str(sz))
    b.set(qn("w:space"), "0")
    b.set(qn("w:color"), color)
    return b


def style_table(t, widths):
    tbl = t._tbl
    tblpr = tbl.tblPr
    for tag in ("w:tblW", "w:tblInd", "w:tblBorders", "w:tblLayout", "w:tblCellMar", "w:tblLook"):
        e = tblpr.find(qn(tag))
        if e is not None:
            tblpr.remove(e)

    def add(tag, **attrs):
        e = OxmlElement(tag)
        for k, v in attrs.items():
            e.set(qn(k), v)
        tblpr.append(e)
        return e

    add("w:tblW", **{"w:w": str(TEXT_W), "w:type": "dxa"})
    # 表格缩进 = 单元格左内边距：Word 2010 兼容模式与 Google 文档都把 tblInd 量到单元格文字，
    # 为 0 时外框会伸出版心 4 pt（2026-11-04 实测）；缩进 4 pt 后外框正好落在版心左右边线上
    add("w:tblInd", **{"w:w": str(CELL_MAR), "w:type": "dxa"})
    bd = add("w:tblBorders")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        bd.append(border_el(f"w:{side}", BORDER, BORDER_SZ))
    add("w:tblLayout", **{"w:type": "fixed"})
    mar = add("w:tblCellMar")
    for side in ("top", "left", "bottom", "right"):
        e = OxmlElement(f"w:{side}")
        e.set(qn("w:w"), str(CELL_MAR))
        e.set(qn("w:type"), "dxa")
        mar.append(e)
    add("w:tblLook", **{"w:val": "0000", "w:firstRow": "0", "w:lastRow": "0", "w:firstColumn": "0",
                        "w:lastColumn": "0", "w:noHBand": "1", "w:noVBand": "1"})
    grid = tbl.tblGrid
    for gc, w in zip(grid.findall(qn("w:gridCol")), widths):
        gc.set(qn("w:w"), str(w))
    for row in t.rows:
        for c, w in zip(row.cells, widths):
            tc_prop(c, "w:tcW", {"w:w": str(w), "w:type": "dxa"})


# ---------------------------------------------------------------- 元信息（页眉）
def guess_meta(title, md_path, a):
    stem = os.path.splitext(os.path.basename(md_path))[0]
    ticker = a.ticker
    if not ticker:
        m = re.search(r"(?<![A-Za-z])(HK\.\d{5}|\d{6}\.(?:SH|SZ)|[A-Z]{1,5}(?:\.[A-Z]{1,2})?)(?![A-Za-z])", title or "")
        if not m:
            m = re.match(r"^([A-Z]{1,5}(?:\.[A-Z]{1,2})?)_", stem)
        ticker = m.group(1) if m else ""
    date = a.date
    if not date:
        for s in (title or "", stem):
            m = re.search(r"(20\d{2})[-/.年]?(\d{1,2})[-/.月]?(\d{1,2})日?", s)
            if m:
                try:
                    date = dt.date(int(m.group(1)), int(m.group(2)), int(m.group(3))).isoformat()
                    break
                except ValueError:
                    pass
    date = date or dt.date.today().isoformat()
    return ticker, date


def setup_page(doc, report, ticker, date):
    sec = doc.sections[0]
    sec.orientation = WD_ORIENT.PORTRAIT
    sec.page_width, sec.page_height = Twips(PAGE_W), Twips(PAGE_H)
    for side in ("top_margin", "bottom_margin", "left_margin", "right_margin"):
        setattr(sec, side, Twips(MARGIN))
    sec.header_distance = sec.footer_distance = Cm(1)
    cols = sec._sectPr.find(qn("w:cols"))
    if cols is None:
        cols = OxmlElement("w:cols")
        sec._sectPr.append(cols)
    cols.set(qn("w:num"), "1")
    hp = sec.header.paragraphs[0]
    hp.style = doc.styles["Header"]
    hp.paragraph_format.tab_stops.add_tab_stop(Twips(TEXT_W), WD_TAB_ALIGNMENT.RIGHT)
    right = " · ".join(x for x in (ticker, date) if x)
    script_runs(hp, f"{report}\t{right}")
    fp = sec.footer.paragraphs[0]
    fp.style = doc.styles["Footer"]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_field(fp._p, "PAGE", "1")
    fp._p.append(fld_run("text", " / "))
    add_field(fp._p, "NUMPAGES", "1")


def add_toc(doc, entries):
    """目录：独立标题"目录"（段前分页）+ 目录域（TOC \\o "1-2" \\h \\z \\u \\n 1-1 \\p "　…　"，缓存结果为各条目，
    页码打开时更新）。一级条目只有标题链接、不显示页码；二级条目"标题　…　页码"，链接只在标题文字上，页码灰色。"""
    h = doc.add_paragraph(style=ALONE_STYLE)
    script_runs(h, "目录")
    h.paragraph_format.page_break_before = True
    body = doc.element.body
    sdt = OxmlElement("w:sdt")
    sdtpr = OxmlElement("w:sdtPr")
    dpo = OxmlElement("w:docPartObj")
    g = OxmlElement("w:docPartGallery")
    g.set(qn("w:val"), "Table of Contents")
    u = OxmlElement("w:docPartUnique")
    dpo.append(g)
    dpo.append(u)
    sdtpr.append(dpo)
    sdt.append(sdtpr)
    content = OxmlElement("w:sdtContent")
    sdt.append(content)
    paras = []
    for lv, text, anchor in entries or [(1, "（目录在 Word 打开时更新）", None)]:
        p = doc.add_paragraph(style=TOC_STYLES[min(lv, 2) - 1])
        paras.append(p)
        if anchor:
            add_hyperlink(p, text, anchor=anchor)
            if lv >= 2:
                for sep in script_runs(p, TOC_SEP):
                    sep.font.color.rgb = NOTE_COLOR
                for r in (fld_run("begin"), fld_run("instr", f"PAGEREF {anchor} \\h"), fld_run("separate"),
                          fld_run("text", ""), fld_run("end")):
                    rpr = OxmlElement("w:rPr")
                    c = OxmlElement("w:color")
                    c.set(qn("w:val"), "666666")
                    rpr.append(c)
                    r.insert(0, rpr)
                    p._p.append(r)
        else:
            script_runs(p, text)
    first, last = paras[0]._p, paras[-1]._p
    head = [fld_run("begin"), fld_run("instr", f'TOC \\o "1-2" \\h \\z \\u \\n 1-1 \\p "{TOC_SEP}"'), fld_run("separate")]
    for r in reversed(head):
        if first.pPr is not None:
            first.pPr.addnext(r)
        else:
            first.insert(0, r)
    last.append(fld_run("end"))
    for p in paras:
        body.remove(p._p)
        content.append(p._p)
    body.insert(list(body).index(body[-1]), sdt)  # 放在 sectPr 之前
    return h


# ---------------------------------------------------------------- 转换
def write_table(doc, data, caption_par):
    n = len(data["head"])
    rows = []
    for _, row in data["rows"]:
        if len(row) > n:  # 多出的格并入最后一列（检查里已报告）
            row = row[:n - 1] + [" | ".join(row[n - 1:])]
        rows.append((row + [""] * n)[:n])
    widths, overflow, is_text = column_widths(data["head"], rows)
    t = doc.add_table(rows=1, cols=n)
    t.style = doc.styles["Table Grid"]
    style_table(t, widths)
    t.alignment = WD_TABLE_ALIGNMENT.LEFT
    tr_pr = t.rows[0]._tr.get_or_add_trPr()
    tr_pr.append(OxmlElement("w:tblHeader"))
    for j, txt in enumerate(data["head"][:n]):
        cell = t.rows[0].cells[j]
        cp = cell.paragraphs[0]
        cp.style = doc.styles[TH_STYLE]
        ind, unit = split_head(txt)
        if unit:
            add_inline(cp, ind)
            cp.add_run().add_break()
            for r in script_runs(cp, unit, NOTE_PT):
                r.bold = False
        else:
            add_inline(cp, txt)
        tc_prop(cell, "w:shd", {"w:val": "clear", "w:color": "auto", "w:fill": HEAD_FILL})
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    for row in rows:
        total = plain(row[0]).startswith(TOTAL_PREFIX)
        cells = t.add_row().cells
        for j, txt in enumerate(row):
            cell = cells[j]
            cp = cell.paragraphs[0]
            val = is_value(txt)
            cp.style = doc.styles[TD_STYLE if val else TDT_STYLE]
            cp.alignment = WD_ALIGN_PARAGRAPH.CENTER if val else WD_ALIGN_PARAGRAPH.LEFT
            add_inline(cp, txt, bold=total)
            tc_prop(cell, "w:shd", {"w:val": "clear", "w:color": "auto", "w:fill": "FFFFFF"})
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
            if total:
                tc_prop(cell, "w:tcBorders", children=[border_el("w:top", *TOTAL_BORDER)])
    return t, overflow


def convert(blocks, heads, title_mode, img_root, out, meta, toc=True):
    doc = Document()
    setup_styles(doc)
    report, ticker, date = meta
    setup_page(doc, report, ticker, date)
    for p in list(doc.paragraphs):  # 模板自带的空段
        p._p.getparent().remove(p._p)
    hmap = {h["idx"]: h for h in heads}
    toc_entries, bid, warnings = [], 0, []
    for h in heads:
        if h["level"] in (1, 2) and h["role"] in ("body", "appendix", "revision", "sub-body", "sub-appendix",
                                                   "sub-revision"):
            bid += 1
            h["anchor"] = f"_Toc{bid:06d}"
            toc_entries.append((h["level"], h["text"], h["anchor"]))
    toc_done = not toc
    has_toc_head = any(h["role"] == "toc" for h in heads)
    prev, prev_par_kind = None, None
    first_after_title = True
    for k, (kind, data) in enumerate(blocks):
        if kind == "h":
            h = hmap[k]
            lv, role = h["level"], h["role"]
            if role == "toc":
                if not toc_done:
                    add_toc(doc, toc_entries)
                    toc_done = True
                prev, first_after_title = "toc", False
                continue
            if lv == 1 and role in ("body", "appendix", "revision") and not toc_done and not has_toc_head:
                add_toc(doc, toc_entries)
                toc_done = True
                first_after_title = False
            if lv == 0:
                p = doc.add_paragraph(style="Title")
            elif lv == 1 and role == "summary":
                p = doc.add_paragraph(style=ALONE_STYLE)
            else:
                p = doc.add_paragraph(style=f"Heading {lv}")
            add_inline(p, data[1])
            if lv == 1 and role != "summary" and not first_after_title:
                p.paragraph_format.page_break_before = True
            if h.get("anchor"):
                add_bookmark(p, h["anchor"], int(h["anchor"][4:]))
            first_after_title = lv == 0
            prev = kind
            continue
        first_after_title = False
        if kind == "p":
            nxt = blocks[k + 1][0] if k + 1 < len(blocks) else None
            if nxt == "table" and RE_CAPTION.match(plain(data)):
                p = doc.add_paragraph(style=CAPTION_STYLE)
            else:
                p = doc.add_paragraph(style=NOTE_STYLE if is_note(data) else None)
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
                code_runs(p, ln, 9)
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
            _, overflow = write_table(doc, data, None)
            if overflow:
                warnings.append(f"第 {data['line']} 行的表格：版心放不下全部数据列，数据可能断行，按竖版分表处理（格式规范 A03，不用横向页）")
        prev = kind
    if not toc_done:
        add_toc(doc, toc_entries)
    settings = doc.settings.element
    uf = OxmlElement("w:updateFields")
    uf.set(qn("w:val"), "true")
    settings.append(uf)
    slim(doc)
    doc.save(out)
    return warnings


def slim(doc):
    """去掉模板里用不到的部分：stylesWithEffects、customXml、主题、webSettings、字体表、扩展属性、缩略图、
    潜在样式表、未使用的样式和编辑器元数据（字体都已显式写在样式里，不靠主题）。"""
    part = doc.part
    for rid, rel in list(part.rels.items()):
        if rel.reltype.endswith(("/stylesWithEffects", "/customXml", "/theme", "/webSettings")):
            part.drop_rel(rid)
    for rid, rel in list(part.rels.items()):
        if rel.reltype.endswith("/fontTable"):
            part.drop_rel(rid)
    pkg = part.package
    for rid, rel in list(pkg.rels.items()):
        if rel.reltype.endswith(("/thumbnail", "/extended-properties", "/core-properties")):
            pkg.rels.pop(rid)
    st_el = doc.settings.element
    for tag in ("w:savePreviewPicture", "w:rsids", "m:mathPr", "w:shapeDefaults", "w:proofState", "w14:docId",
                "w:attachedTemplate", "w:doNotAutoCompressPictures", "w:clrSchemeMapping"):
        e = st_el.find(qn(tag))
        if e is not None:
            st_el.remove(e)
    tfl = st_el.find(qn("w:themeFontLang"))
    if tfl is not None:
        st_el.remove(tfl)
    styles = doc.styles.element
    ls = styles.find(qn("w:latentStyles"))
    if ls is not None:
        styles.remove(ls)
    for e in styles.findall(qn("w:style")):  # 去掉链接的字符样式与编辑器元数据（不影响版式）
        for tag in ("w:link", "w:rsid", "w:uiPriority", "w:qFormat", "w:unhideWhenUsed", "w:semiHidden"):
            for x in e.findall(qn(tag)):
                e.remove(x)
    used = set()
    trees = [part.element] + [p.part.element for sec in doc.sections for p in (sec.header, sec.footer)]
    for tree in trees:
        for tag in ("w:pStyle", "w:rStyle", "w:tblStyle"):
            for e in tree.iter(qn(tag)):
                used.add(e.get(qn("w:val")))
    by_id = {e.get(qn("w:styleId")): e for e in styles.findall(qn("w:style"))}
    keep = set(used) | {i for i, e in by_id.items() if e.get(qn("w:default")) in ("1", "true")}
    keep |= {"Header", "Footer", "Heading1", "Heading2", "Heading3", "Title", "Hyperlink"} & set(by_id)
    keep |= {i for i, e in by_id.items() if e.find(qn("w:name")).get(qn("w:val")) in TOC_STYLES + (NOTE_STYLE,)}
    stack = list(keep)
    while stack:
        e = by_id.get(stack.pop())
        if e is None:
            continue
        for tag in ("w:basedOn", "w:link", "w:next"):
            x = e.find(qn(tag))
            if x is not None and x.get(qn("w:val")) not in keep:
                keep.add(x.get(qn("w:val")))
                stack.append(x.get(qn("w:val")))
    for i, e in by_id.items():
        if i not in keep:
            styles.remove(e)
    if part.element.find(".//" + qn("w:numPr")) is None and styles.find(".//" + qn("w:numPr")) is None:
        for rid, rel in list(part.rels.items()):  # 没有列表时不带编号定义
            if rel.reltype.endswith("/numbering"):
                part.drop_rel(rid)


# ---------------------------------------------------------------- 检查
RE_META = re.compile(r"数据截至|不构成投资建议|^口径[:：]|^\S.*·.*·")
RE_HEAD_PREFIX = re.compile(rf"^(?:附{CN}?[、：:.．\s　]*|{CN}、|\d+(?:\.\d+)*[\s　]*)")
# 出处标注与过程用语：[R123 p5]、[S01]、p5]、研报库、台账、对账、取证（写法硬要求 1）
RE_CITE = re.compile(r"\[[RS]\d+[^\]\n]*\]|\[[RS]\d+|p\.?\s?\d+\]|研报库|台账|对账|取证")
# 图下来源行里的研报 / 公开资料编号与页码
RE_SRC_ID = re.compile(r"(?<![A-Za-z0-9])[RS]\d{2,}|(?<![A-Za-z0-9])p\.?\s?\d+(?!\d)|第\s*\d+\s*页")


def head_core(text):
    """去掉标题的编号前缀（一、 / 1.1　 / 附一、 / 附：），只留标题文字。"""
    return RE_HEAD_PREFIX.sub("", text, count=1).strip()


def check_shell(blocks, heads):
    """格式规范第 2 节报告外壳：标题 → 一页摘要 → 目录 → 一、结论与概述（下有结论、概述）→ 任务书定的各章 →
    附：口径与说明；目录由脚本插在一页摘要之后；取舍、来源、修订说明进底稿。只检查不改写。"""
    issues = []
    h1 = [h for h in heads if h["level"] == 1 and h["role"] != "toc"]
    title = next((h for h in heads if h["level"] == 0), None)
    if not h1:
        return ["没有一级标题：标题后应为'一页摘要'，再目录、'一、结论与概述'，最后一节为'附：口径与说明'"]
    first = h1[0]
    start = title["idx"] + 1 if title and title["idx"] < first["idx"] else 0
    front = [b for b in blocks[start:first["idx"]] if b[0] != "h"]
    if front:
        lead = plain(front[0][1]) if front[0][0] in ("p", "quote") else front[0][0]
        meta = any(b[0] in ("p", "quote") and RE_META.search(plain(b[1])) for b in front)
        issues.append(f"标题与第一个一级标题之间有 {len(front)} 段（首段「{lead[:30]}」），标题后应直接进'一页摘要'"
                      + ("；元信息放'附：口径与说明'" if meta else ""))
    summary = first if first["role"] == "summary" else None
    if summary is None:
        issues.append(f"标题后第一个一级标题应为'一页摘要'，现为'{first['text']}'")
    else:
        end = h1[1]["idx"] if len(h1) > 1 else len(blocks)
        if not any(b[0] != "h" for b in blocks[summary["idx"] + 1:end]):
            issues.append("'一页摘要'下没有内容：应写结论与主要内容的一页浓缩")
        toc_h = next((h for h in heads if h["role"] == "toc"), None)
        if toc_h and toc_h["idx"] < summary["idx"]:
            issues.append("'目录'写在一页摘要之前：目录应在一页摘要之后、'一、结论与概述'之前（不写时脚本自动插）")
    for h in h1[1:]:
        if h["role"] == "summary":
            issues.append(f"'{h['text']}'：一页摘要只能有一个，紧接标题")
    body = [h for h in h1 if h["role"] != "summary"]
    if not body:
        issues.append("一页摘要之后没有正文章：应为'一、结论与概述'及任务书定的各章")
    elif "结论与概述" not in body[0]["text"]:
        issues.append(f"一页摘要与目录之后的第一章应为'一、结论与概述'，现为'{body[0]['text']}'")
    else:
        chap = body[0]
        nxt = [h for h in h1 if h["idx"] > chap["idx"]]
        end = nxt[0]["idx"] if nxt else len(blocks)
        subs = [head_core(h["text"]) for h in heads if chap["idx"] < h["idx"] < end and h["level"] >= 2]
        bolds = [b[1] if b[0] in ("p", "quote") else b[1][1] for b in blocks[chap["idx"] + 1:end]
                 if b[0] in ("p", "quote", "ul", "ol")]
        for need in ("结论", "概述"):
            if not any(t.startswith(need) for t in subs) and \
                    not any(re.match(rf"\s*\*\*{need}", t) for t in bolds):
                issues.append(f"'{chap['text']}'下缺'{need}'：应有'{need}'子标题或以 **{need}** 开头的加粗段")
    last = h1[-1]
    if not any("口径与说明" in h["text"] for h in h1):
        issues.append("文末没有'附：口径与说明'一节（口径、数据截至、版本、免责声明放这里）")
    elif "口径与说明" not in last["text"]:
        issues.append(f"最后一节应为'附：口径与说明'，现为'{last['text']}'")
    for h in heads:
        if h["level"] < 1 or h["role"] in ("toc", "summary"):
            continue
        t, core = h["text"], head_core(h["text"])
        if "一页摘要" in t or (h["level"] == 1 and core.startswith("摘要")):
            issues.append(f"'{t}'作为节标题出现在正文：一页摘要只放在标题之后、目录之前")
        elif "修订说明" in t or "修订记录" in t:
            issues.append(f"'{t}'作为节标题出现在正文：修订说明进底稿 复核/修订说明_Vx.md")
        elif "本报告的取舍" in t or core.startswith("取舍"):
            issues.append(f"'{t}'作为节标题出现在正文：取舍进底稿")
        elif core.startswith("来源"):
            issues.append(f"'{t}'作为节标题出现在正文：来源清单进底稿")
    return issues


def check_citations(blocks, heads):
    """写法硬要求 1：正文零出处标注。查 [R… / [S… / p…] 与"研报库""台账""对账""取证"，
    不查图下"数据与可编辑图表"行和代码块。只检查不改写。返回 (总数, [(所在标题, [命中…])])。"""
    hmap = {h["idx"]: h["text"] for h in heads}
    where, hits, order = "（标题前）", {}, []
    for k, b in enumerate(blocks):
        if b[0] == "h":
            where = hmap.get(k, plain(b[1][1]))
            texts = [b[1][1]]
        elif b[0] in ("p", "quote"):
            texts = [] if plain(b[1]).startswith("数据与可编辑图表") else [b[1]]
        elif b[0] in ("ul", "ol"):
            texts = [b[1][1]]
        elif b[0] == "table":
            texts = list(b[1]["head"]) + [c for _, r in b[1]["rows"] for c in r]
        elif b[0] == "img":
            texts = [b[1]["alt"]]
        else:
            continue
        for t in texts:
            found = RE_CITE.findall(plain(t))
            if found:
                if where not in hits:
                    hits[where] = []
                    order.append(where)
                hits[where].extend(found)
    return sum(len(v) for v in hits.values()), [(w, hits[w]) for w in order]


# 图表优先（格式规范第 7 节 V02–V04，软门槛：只提示，不影响退出码）
RE_NUM_TOKEN = re.compile(r"(?<![A-Za-z])[-−+]?\d[\d,]*(?:\.\d+)?%?")
RE_NUM_CELL = re.compile(r"[-−+]?\(?[-−+]?\d[\d,]*(\.\d+)?\)?\s*(%|x|倍|亿|万|pt|个月|季|天|年)?")
NOTE_LEAD = ("表注", "图注", "来源", "资料来源", "数据与可编辑图表", "口径说明", "注：", "注:")
DENSE_NUMS = 12        # 一段（或一条列表项）里数字 ≥ 12 个 → 数字墙提示
BIG_TABLE_CELLS = 40   # 一张表里纯数值单元格 ≥ 40 个 → 数值大表提示


def check_chart_first(blocks, heads):
    """格式规范第 7 节"图表优先"：可比序列、走势、阶段、先后用图表达，正文与表格只留关键节点数字与口径。
    只提示不改写、不影响退出码：数字密集段落、数值大表、没有图的正文章。返回 dict。"""
    hmap = {h["idx"]: h for h in heads}
    where, dense, big, chap_imgs, cur = "（标题前）", [], [], {}, None
    for k, b in enumerate(blocks):
        if b[0] == "h":
            h = hmap.get(k)
            if h:
                where = h["text"]
                if h["level"] == 1:
                    cur = h if h["role"] == "body" and "结论与概述" not in h["text"] else None
                    if cur is not None:
                        chap_imgs[cur["text"]] = 0
            continue
        if b[0] == "img":
            if cur is not None:
                chap_imgs[cur["text"]] += 1
            continue
        if b[0] in ("p", "quote", "ul", "ol"):
            t = plain(b[1] if b[0] in ("p", "quote") else b[1][1])
            if t.startswith(NOTE_LEAD):
                continue
            n = len(RE_NUM_TOKEN.findall(t))
            if n >= DENSE_NUMS:
                dense.append(f"「{where}」：一段 {n} 个数字（「{t[:20]}…」）——能画成图的改成图，正文只留关键节点数字")
        elif b[0] == "table":
            cells = [plain(c) for _, r in b[1]["rows"] for c in r]
            nnum = sum(1 for c in cells if c and RE_NUM_CELL.fullmatch(c.replace(" ", "")))
            if nnum >= BIG_TABLE_CELLS:
                big.append(f"「{where}」的表：{nnum} 个纯数值单元格——序列改成图，表格只留口径与关键节点")
    nofig = [f"「{c}」：没有图" for c, n in chap_imgs.items() if n == 0]
    return {"dense": dense, "big": big, "nofig": nofig}

RE_FIG_NO = re.compile(r"^\s*图\s*(\d+)")
RE_FIG_FILE = re.compile(r"fig0*(\d+)", re.I)
RE_FIG_REF = re.compile(r"图\s?(\d+)")


def check_figures(blocks):
    """格式规范第 7 节：图号按正文出现顺序编（第一张是图 1），图片文件名 figN 与图号一致；正文引用的图 N 不超过图总数。
    只检查不改写。返回 (图片数, 问题列表)。"""
    issues, k = [], 0
    for b in blocks:
        if b[0] != "img":
            continue
        k += 1
        alt, path = plain(b[1]["alt"]), b[1]["path"]
        m = RE_FIG_NO.match(alt)
        fm = RE_FIG_FILE.search(os.path.basename(path))
        if not m:
            issues.append(f"第 {k} 张图（{os.path.basename(path)}）的图注没有写'图 {k}'")
        elif int(m.group(1)) != k:
            issues.append(f"第 {k} 张图标为'图 {m.group(1)}'：图号应按出现顺序编，这里应为图 {k}")
        n = int(m.group(1)) if m else k
        if not fm:
            issues.append(f"图 {n} 的文件名 {os.path.basename(path)} 没有 fig{n}：文件名 figN 应与图号一致")
        elif int(fm.group(1)) != n:
            issues.append(f"图注'图 {n}'与文件名 fig{fm.group(1)}（{os.path.basename(path)}）不一致")
    refs = set()
    for b in blocks:
        if b[0] in ("p", "quote"):
            texts = [b[1]]
        elif b[0] in ("ul", "ol"):
            texts = [b[1][1]]
        elif b[0] == "table":
            texts = list(b[1]["head"]) + [c for _, r in b[1]["rows"] for c in r]
        else:
            continue
        for t in texts:
            refs.update(int(x) for x in RE_FIG_REF.findall(plain(t)))
    for n in sorted(x for x in refs if x > k or x == 0):
        issues.append(f"正文引用了'图 {n}'，但全文只有 {k} 张图")
    return k, issues


def check(blocks, img_root):
    heads, title_mode = outline(blocks)
    tables = [b for b in blocks if b[0] == "table"]
    imgs = [b for b in blocks if b[0] == "img"]
    missing = [b[1]["path"] for b in imgs if not os.path.isfile(resolve(b[1]["path"], img_root))]
    bad_rows, unit_issues = [], []
    for ti, (_, t) in enumerate(tables, start=1):
        n = len(t["head"])
        for line_no, row in t["rows"]:
            if len(row) != n:
                bad_rows.append(f"第 {line_no} 行：{len(row)} 列，表头 {n} 列")
        for j, hd in enumerate(t["head"]):
            ph = plain(hd)
            body = [plain(r[j]) for _, r in t["rows"] if j < len(r) and plain(r[j])]
            if RE_HALF_UNIT.search(ph):
                unit_issues.append(f"表 {ti}「{ph}」：单位用了半角括号，应为全角（）")
            elif (body and all(re.fullmatch(r"[-−+]?\(?[-−+]?\d[\d,]*(\.\d+)?\)?", x) for x in body)
                  and split_head(hd)[1] is None and not re.search(r"年|日期|期间|编号|代码|序号|年度|季度", ph)):
                unit_issues.append(f"表 {ti}「{ph}」：数值列表头没有标单位")
    no_src, src_ids = [], []
    for k, b in enumerate(blocks):
        if b[0] != "img":
            continue
        nxt = [x for x in blocks[k + 1:k + 4] if x[0] in ("p", "quote")]
        srcs = [plain(x[1]) for x in nxt if "数据与可编辑图表" in x[1] or plain(x[1]).startswith(("来源", "资料来源"))]
        if not srcs:
            no_src.append(b[1]["path"])
        for t in srcs:
            if RE_SRC_ID.search(t):
                src_ids.append(f"{os.path.basename(b[1]['path'])} 的来源行带研报编号或页码：「{t[:40]}」，"
                               "来源行只写数据提供方与口径")
    lv = {n: sum(1 for h in heads if h["level"] == n) for n in (0, 1, 2, 3)}
    num_issues, ncount = check_numbering(heads)
    return {"heads": len(heads), "lv": lv, "tables": len(tables), "imgs": len(imgs), "missing": missing,
            "bad_rows": bad_rows, "no_src": no_src, "title_mode": title_mode, "outline": heads,
            "num_issues": num_issues, "ncount": ncount, "unit_issues": unit_issues,
            "shell_issues": check_shell(blocks, heads), "cite": check_citations(blocks, heads),
            "src_ids": src_ids, "fig_issues": check_figures(blocks)[1], "chart_first": check_chart_first(blocks, heads)}


def verify_docx(path, expect):
    """读回 .docx，逐项核对格式规范 Part A 里能程序化读取的项目。返回 [(项目, 通过, 说明)]。"""
    d = Document(path)
    res = []

    def add(name, ok, detail=""):
        res.append((name, bool(ok), detail))

    def rfonts(el):
        rf = el.find(qn("w:rPr") + "/" + qn("w:rFonts")) if el is not None else None
        return (rf.get(qn("w:ascii")), rf.get(qn("w:eastAsia"))) if rf is not None else (None, None)

    sec = d.sections[0]
    sp = sec._sectPr
    pg, mar = sp.find(qn("w:pgSz")), sp.find(qn("w:pgMar"))
    cols = sp.find(qn("w:cols"))
    ok = (pg.get(qn("w:w")), pg.get(qn("w:h"))) == (str(PAGE_W), str(PAGE_H)) and pg.get(qn("w:orient")) in (None, "portrait")
    ok &= all(mar.get(qn(f"w:{s}")) == str(MARGIN) for s in ("top", "bottom", "left", "right"))
    ok &= cols is None or cols.get(qn("w:num")) in (None, "1")
    add("A4 竖版、四边 2 cm、单栏", ok, f"纸张 {pg.get(qn('w:w'))}×{pg.get(qn('w:h'))} twips，边距 {mar.get(qn('w:top'))}")
    st = d.styles
    hdr_txt = "".join(p.text for p in sec.header.paragraphs)
    hdr_ok = expect["report"] in hdr_txt and "\t" in hdr_txt and expect["date"] in hdr_txt
    hdr_ok &= (not expect["ticker"]) or expect["ticker"] in hdr_txt
    fxml = sec.footer._element.xml
    ftr_ok = "PAGE" in fxml and "NUMPAGES" in fxml and sec.footer.paragraphs[0].alignment == WD_ALIGN_PARAGRAPH.CENTER
    hf_ok = all(st[n].font.size == Pt(HF_PT) and st[n].font.name == LATIN for n in ("Header", "Footer"))
    add(f"页眉左报告名、右代码与日期；页脚居中 页码 / 总页数；{LATIN} 8 pt", hdr_ok and ftr_ok and hf_ok,
        f"页眉「{hdr_txt.replace(chr(9), ' ⇥ ')}」")
    n = st["Normal"]
    a, e = rfonts(n.element)
    pf = n.paragraph_format
    add(f"正文中文、英文与数字统一 {LATIN}，{BODY_PT} pt、行距 {BODY_LINE}、段后 6 pt",
        (a, e) == (LATIN, CJK_FONT) and n.font.size == Pt(BODY_PT) and pf.line_spacing == BODY_LINE
        and pf.space_after == Pt(6), f"英文数字 {a}、中文 {e}、{n.font.size.pt if n.font.size else None} pt、行距 {pf.line_spacing}")
    bad = []
    for lv, (size, before, after) in HEADS.items():
        s = st[f"Heading {lv}"]
        f = s.paragraph_format
        c = s.element.rPr.find(qn("w:color"))
        if not (s.font.size == Pt(size) and s.font.bold and c is not None and c.get(qn("w:val")) == "000000"
                and c.get(qn("w:themeColor")) is None and f.space_before == Pt(before) and f.space_after == Pt(after)
                and f.line_spacing == HEAD_LINE and rfonts(s.element) == (LATIN, CJK_FONT)):
            bad.append(f"Heading {lv}")
    add("一二三级标题 18/14/12 pt 加粗黑色、段前段后 16/8 12/6 8/4、行距 1.0", not bad, f"不符 {bad or '无'}")
    ns = st[NOTE_STYLE]
    paras = d.paragraphs
    notes = [p for p in paras if p.style.name == NOTE_STYLE]
    ok = ns.font.size == Pt(NOTE_PT) and ns.font.color.rgb == NOTE_COLOR and len(notes) == expect["notes"]
    add("注释类段落 9 pt #666666", ok, f"注释段 {len(notes)}（应为 {expect['notes']}）")
    small = [r.text for p in paras for r in p.runs if r.font.size is not None and r.font.size < Pt(NOTE_PT)]
    for t in d.tables:
        for row in t.rows:
            for c in row.cells:
                for p in c.paragraphs:
                    small += [r.text for r in p.runs if r.font.size is not None and r.font.size < Pt(NOTE_PT)]
    add("没有小于 9 pt 的文字（字号不为塞下而缩小）", not small, f"{small[:3] or '无'}")
    cjk_bad, n_cjk = [], 0
    parts = [d.element.body] + [x._element for x in (sec.header, sec.footer)]
    for el in parts:
        for r in el.iter(qn("w:r")):
            t = "".join(x.text or "" for x in r.iter(qn("w:t")))
            if not t or not CJK.search(t):
                continue
            n_cjk += 1
            rf = r.find(qn("w:rPr") + "/" + qn("w:rFonts"))
            if rf is None or rf.get(qn("w:ascii")) != CJK_FONT or RE_CJK_RUN.sub("", t):
                cjk_bad.append(t[:8])
    add(f"中文段落单独成 run 并显式 {CJK_FONT}（Google 文档只认 run 字体）", not cjk_bad,
        f"中文 run {n_cjk}；不符 {cjk_bad[:4] or '无'}")
    body = d.element.body
    xml = body.xml
    has_toc = 'TOC \\o "1-2"' in xml
    pb = [p.text for p in paras if p.paragraph_format.page_break_before]
    order_ok = True
    if expect["summary"] and has_toc:
        texts = [p.text for p in paras]
        order_ok = texts.index(expect["summary"]) < texts.index("目录") < texts.index(expect["first_body"]) \
            if expect["first_body"] else True
    want_pb = ["目录"] + expect["h1_break"]
    add("顺序 标题→一页摘要→目录→正文；目录、各一级板块、附录另起一页", has_toc and order_ok and pb == want_pb,
        f"目录域{'有' if has_toc else '无'}；段前分页 {len(pb)} 处（应为 {len(want_pb)}）")
    t1, t2 = st[TOC_STYLES[0]], st[TOC_STYLES[1]]
    toc_ok = f'\\n 1-1 \\p "{TOC_SEP}"' in xml if has_toc else True
    toc_ok &= t1.font.size == Pt(CAPTION_PT) and bool(t1.font.bold) and t2.font.size == Pt(CAPTION_PT) and not t2.font.bold
    toc_ok &= t2.paragraph_format.left_indent == Pt(18) and t1.paragraph_format.space_before == Pt(10) \
        and t2.paragraph_format.space_after == Pt(5)
    toc_ok &= all(x.paragraph_format.line_spacing == TOC_LINE and x.element.pPr.find(qn("w:tabs")) is None for x in (t1, t2))
    add("目录：一级 10 pt 加粗无页码；二级 10 pt 常规缩进 18 pt，「标题　…　页码」不右对齐不用点线；行距 1.25", toc_ok,
        f"目录域开关{'含' if has_toc and TOC_SEP in xml else '缺'} \\n 1-1 与「　…　」")
    tb_bad = []
    for ti, t in enumerate(d.tables, start=1):
        tblpr = t._tbl.tblPr
        bd = tblpr.find(qn("w:tblBorders"))
        sides = [bd.find(qn(f"w:{s}")) for s in ("top", "left", "bottom", "right", "insideH", "insideV")] if bd is not None else []
        if not sides or any(x is None or x.get(qn("w:color")) != BORDER or x.get(qn("w:sz")) != str(BORDER_SZ) for x in sides):
            tb_bad.append(f"表{ti} 边框")
        mar = tblpr.find(qn("w:tblCellMar"))
        if mar is None or any(mar.find(qn(f"w:{s}")).get(qn("w:w")) != str(CELL_MAR) for s in ("top", "left", "bottom", "right")):
            tb_bad.append(f"表{ti} 内边距")
        widths = [int(g.get(qn("w:w"))) for g in t._tbl.tblGrid.findall(qn("w:gridCol"))]
        if sum(widths) != TEXT_W or tblpr.find(qn("w:tblW")).get(qn("w:w")) != str(TEXT_W):
            tb_bad.append(f"表{ti} 表宽 {sum(widths)}≠{TEXT_W}")
        ind = tblpr.find(qn("w:tblInd"))
        if ind is None or ind.get(qn("w:w")) != str(CELL_MAR):
            tb_bad.append(f"表{ti} 缩进（外框应对齐版心）")
        for ri, row in enumerate(t.rows):
            for c in row.cells:
                shd = c._tc.tcPr.find(qn("w:shd"))
                fill = shd.get(qn("w:fill")) if shd is not None else None
                p = c.paragraphs[0]
                if ri == 0:
                    runs = list(p.runs)
                    brk = next((k for k, r in enumerate(runs) if r.text == "\n"), None)
                    unit_runs = [r for r in runs[brk + 1:] if r.text] if brk is not None else []
                    head_ok = fill == HEAD_FILL and p.style.name == TH_STYLE and c.vertical_alignment == WD_CELL_VERTICAL_ALIGNMENT.CENTER
                    head_ok &= all(r.font.size == Pt(NOTE_PT) and r.bold is False for r in unit_runs)
                    if not head_ok:
                        tb_bad.append(f"表{ti} 表头「{c.text[:8]}」")
                else:
                    if fill != "FFFFFF" or p.style.name not in (TD_STYLE, TDT_STYLE) or \
                            c.vertical_alignment != WD_CELL_VERTICAL_ALIGNMENT.TOP:
                        tb_bad.append(f"表{ti} 第{ri}行「{c.text[:8]}」")
                    want = WD_ALIGN_PARAGRAPH.CENTER if is_value(c.text) else WD_ALIGN_PARAGRAPH.LEFT
                    if p.alignment != want:
                        tb_bad.append(f"表{ti} 第{ri}行「{c.text[:8]}」对齐")
            if ri and row.cells[0].text.strip().startswith(TOTAL_PREFIX):
                for c in row.cells:
                    top = c._tc.tcPr.find(qn("w:tcBorders"))
                    if top is None or top.find(qn("w:top")) is None or not all(r.bold for r in c.paragraphs[0].runs if r.text):
                        tb_bad.append(f"表{ti} 合计行")
                        break
    th, td = st[TH_STYLE], st[TD_STYLE]
    sty_ok = th.font.size == Pt(TABLE_PT) and th.font.bold and th.font.color.rgb == WHITE and td.font.size == Pt(TABLE_PT) \
        and td.paragraph_format.space_after == Pt(0) and st[TDT_STYLE].paragraph_format.line_spacing == TEXT_CELL_LINE
    add("表格：表头 #17365D 白字 9.5 pt 加粗居中、单位 9 pt 常规；表体 9.5 pt 白底段后 0 行距 1.0、边框 #D9D9D9 0.5 pt、"
        "内边距 4 pt、数值居中文本左对齐顶端、合计行加粗上边框、表宽 = 版心", not tb_bad and sty_ok,
        f"{len(d.tables)} 张表；不符 {tb_bad[:6] or '无'}")
    caps = [p for p in paras if p.style.name == CAPTION_STYLE]
    cap_ok = len(caps) == expect["captions"]
    if caps:
        cs = st[CAPTION_STYLE]
        cap_ok &= cs.font.size == Pt(CAPTION_PT) and bool(cs.font.bold) and all(
            p._p.getnext() is not None and p._p.getnext().tag == qn("w:tbl") for p in caps)
    add("表题在表上方 10 pt 加粗", cap_ok, f"表题 {len(caps)}（应为 {expect['captions']}）")
    return res


def main(argv=None):
    ap = argparse.ArgumentParser(description="报告 .md → .docx（research-sop 样式，格式规范 Part A），附检查")
    ap.add_argument("md")
    ap.add_argument("--out", help="输出 .docx 路径")
    ap.add_argument("--img-root", help="图片相对路径的根目录（默认为 .md 所在目录）")
    ap.add_argument("--check-only", action="store_true", help="只做检查，不输出 docx")
    ap.add_argument("--title", help="页眉左侧的报告名（默认取报告标题）")
    ap.add_argument("--ticker", help="页眉右侧的股票代码（默认从报告标题或文件名里认）")
    ap.add_argument("--date", help="页眉右侧的报告日期 YYYY-MM-DD（默认从报告标题或文件名里认，再不行用今天）")
    ap.add_argument("--no-toc", action="store_true", help="不插目录")
    a = ap.parse_args(argv)
    if not a.check_only and not a.out:
        ap.error("缺 --out（或用 --check-only）")
    with open(a.md, encoding="utf-8-sig") as f:
        md = f.read()
    img_root = a.img_root or os.path.dirname(os.path.abspath(a.md))
    blocks = parse(md)
    r = check(blocks, img_root)
    heads = r["outline"]
    title_txt = next((h["text"] for h in heads if h["level"] == 0), None) or \
        next((h["text"] for h in heads if h["level"] == 1), None) or os.path.splitext(os.path.basename(a.md))[0]
    report = a.title or title_txt
    ticker, date = guess_meta(title_txt, a.md, a)
    print(f"== md2docx 检查：{os.path.basename(a.md)} ==")
    print(f"标题 {r['heads']}（报告标题 {r['lv'][0]} / 一级 {r['lv'][1]} / 二级 {r['lv'][2]} / 三级 {r['lv'][3]}）；"
          f"表格 {r['tables']}；图片 {r['imgs']}")
    print(f"层级映射：{'# 为报告标题，##/###/#### 为一/二/三级' if r['title_mode'] else '#/##/### 为一/二/三级（没有单独的报告标题）'}；"
          f"页眉：{report} ⇥ {' · '.join(x for x in (ticker, date) if x)}")
    print(f"表格列数与表头不一致：{'；'.join(r['bad_rows']) if r['bad_rows'] else '无'}")
    print(f"图下注明'数据与可编辑图表'或来源：{r['imgs'] - len(r['no_src'])}/{r['imgs']}"
          + (f"（未注明：{'、'.join(r['no_src'])}）" if r["no_src"] else ""))
    print(f"图下来源行（只检查不改写）：带研报编号或页码 {len(r['src_ids'])} 处"
          + "".join(f"\n  - {x}" for x in r["src_ids"]))
    n1, n2, n3 = r["ncount"]
    print(f"标题编号（只检查不改写）：正文一级 {n1}、二级 {n2}、三级 {n3}；问题 {len(r['num_issues'])} 处")
    for x in r["num_issues"]:
        print(f"  - {x}")
    print(f"表头单位：问题 {len(r['unit_issues'])} 处" + "".join(f"\n  - {x}" for x in r["unit_issues"]))
    print(f"报告外壳（只检查不改写）：问题 {len(r['shell_issues'])} 处"
          + "".join(f"\n  - {x}" for x in r["shell_issues"]))
    n_cite, cite_at = r["cite"]
    print(f"出处标注与过程用语（只检查不改写）：{n_cite} 处"
          + "".join(f"\n  - 「{w}」{len(v)} 处：{'、'.join(v)}" for w, v in cite_at))
    print(f"图号与文件名（只检查不改写）：图片 {r['imgs']} 张；问题 {len(r['fig_issues'])} 处"
          + "".join(f"\n  - {x}" for x in r["fig_issues"]))
    cf = r["chart_first"]
    print(f"图表优先（只提示，不影响退出码；格式规范第 7 节 V02–V04）：数字密集段落 {len(cf['dense'])} 处、"
          f"数值大表 {len(cf['big'])} 张、没有图的正文章 {len(cf['nofig'])} 个"
          + "".join(f"\n  - {x}" for x in (cf["dense"] + cf["big"] + cf["nofig"])[:30])
          + (f"\n  - ……另 {len(cf['dense']) + len(cf['big']) + len(cf['nofig']) - 30} 处" if
             len(cf["dense"]) + len(cf["big"]) + len(cf["nofig"]) > 30 else ""))
    if r["missing"]:
        print(f"错误：{len(r['missing'])} 张图片文件不存在（图片根目录 {img_root}）：", file=sys.stderr)
        for p in r["missing"]:
            print(f"  - {p}", file=sys.stderr)
        return 1
    if a.check_only:
        print("只检查，未输出 docx")
        return 0
    os.makedirs(os.path.dirname(os.path.abspath(a.out)), exist_ok=True)
    warns = convert(blocks, heads, r["title_mode"], img_root, a.out, (report, ticker, date), toc=not a.no_toc)
    for w in warns:
        print(f"提醒：{w}")
    summary = next((h["text"] for h in heads if h["role"] == "summary"), None)
    h1b = [h["text"] for h in heads if h["level"] == 1 and h["role"] in ("body", "appendix", "revision")]
    notes = sum(1 for b in blocks if (b[0] in ("p", "quote") and is_note(b[1])) or (b[0] == "img" and b[1]["alt"]))
    caps = sum(1 for k, b in enumerate(blocks) if b[0] == "p" and k + 1 < len(blocks) and blocks[k + 1][0] == "table"
               and RE_CAPTION.match(plain(b[1])))
    expect = {"report": report, "ticker": ticker, "date": date, "notes": notes, "summary": summary,
              "first_body": h1b[0] if h1b else None, "h1_break": h1b, "captions": caps}
    if a.no_toc:
        expect["h1_break"] = h1b
    res = verify_docx(a.out, expect)
    if a.no_toc:  # 不插目录时，顺序项只看一级板块分页
        res = [(n, ok or n.startswith("顺序"), d) for n, ok, d in res]
    n_ok = sum(1 for _, ok, _ in res if ok)
    print(f"格式检查（格式规范 Part A，读回 .docx 逐项比）：{n_ok}/{len(res)} 通过")
    for name, ok, detail in res:
        print(f"  [{'通过' if ok else '不通过'}] {name}" + (f"：{detail}" if detail else ""))
    print(f"已输出 {a.out}（{os.path.getsize(a.out) / 1024:.1f} KB）")
    return 0 if n_ok == len(res) else 1


if __name__ == "__main__":
    sys.exit(main())
