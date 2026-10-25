#!/usr/bin/env python3
"""在不破坏图表的前提下修改 xlsx 的少量单元格（直接改 XML，不经过 openpyxl 保存）。

为什么：openpyxl 读写会丢失原有图表；本脚本只改目标单元格，其余字节原样保留。
改完后请运行 lo_recalc.sh 重算缓存值，再用 xlsx_diff.py 核对。

用法：
  python3 xlsx_patch.py in.xlsx out.xlsx edits.json

edits.json 格式（工作表用显示名称）：
{
  "假设": {
    "B10": 5.5,                         # 数值
    "H10": {"text": "高盛 2026-10-21 附表 1：..."}   # 文本（会作为新共享字符串追加）
  },
  "说明": {"B3": {"text": "模型日期：..."}}
}
限制：目标单元格必须已存在（有 <c r="..."> 节点）；不新增行列、不改公式。
"""
import html
import json
import re
import shutil
import sys
import tempfile
import zipfile


def sheet_map(z):
    wb = z.read("xl/workbook.xml").decode("utf8")
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf8")
    rid2target = {}
    for m in re.finditer(r"<Relationship\b[^>]*>", rels):
        tag = m.group(0)
        rid = re.search(r'Id="([^"]+)"', tag).group(1)
        tgt = re.search(r'Target="([^"]+)"', tag).group(1)
        rid2target[rid] = tgt
    out = {}
    for m in re.finditer(r"<sheet\b[^>]*>", wb):
        tag = m.group(0)
        name = html.unescape(re.search(r'name="([^"]+)"', tag).group(1))
        rid = re.search(r'r:id="([^"]+)"', tag).group(1)
        tgt = rid2target[rid].lstrip("/")
        if not tgt.startswith("xl/"):
            tgt = "xl/" + tgt
        out[name] = tgt
    return out


def main():
    if len(sys.argv) != 4:
        print(__doc__)
        sys.exit(1)
    src, dst, spec_path = sys.argv[1:]
    spec = json.load(open(spec_path, encoding="utf8"))
    zin = zipfile.ZipFile(src)
    files = {n: zin.read(n) for n in zin.namelist()}
    order = zin.namelist()
    smap = sheet_map(zin)

    ss_name = "xl/sharedStrings.xml"
    ss = files[ss_name].decode("utf8") if ss_name in files else None
    n0 = len(re.findall(r"<si>|<si ", ss)) if ss else 0
    new_strings = []

    for sheet, cells in spec.items():
        if sheet not in smap:
            sys.exit(f"找不到工作表：{sheet}；现有：{list(smap)}")
        part = smap[sheet]
        xml = files[part].decode("utf8")
        for ref, val in cells.items():
            m = re.search(r'<c r="%s"(\s[^>]*)?(/>|>(.*?)</c>)' % re.escape(ref), xml, flags=re.S)
            if not m:
                sys.exit(f"{sheet}!{ref} 不存在（本脚本不新增单元格）")
            attrs = m.group(1) or ""
            attrs = re.sub(r'\st="[^"]*"', "", attrs)
            if isinstance(val, dict) and "text" in val:
                if ss is None:  # 没有共享字符串表时写内联字符串
                    t = html.escape(val["text"], quote=False)
                    new_c = f'<c r="{ref}"{attrs} t="inlineStr"><is><t xml:space="preserve">{t}</t></is></c>'
                else:
                    idx = n0 + len(new_strings)
                    new_strings.append(val["text"])
                    new_c = f'<c r="{ref}"{attrs} t="s"><v>{idx}</v></c>'
            else:
                if isinstance(m.group(3), str) and "<f>" in m.group(3):
                    sys.exit(f"{sheet}!{ref} 是公式单元格，拒绝覆盖")
                new_c = f'<c r="{ref}"{attrs} t="n"><v>{val}</v></c>'
            xml = xml[: m.start()] + new_c + xml[m.end():]
            print(f"已改 {sheet}!{ref}")
        files[part] = xml.encode("utf8")

    if new_strings:
        add = "".join(
            '<si><t xml:space="preserve">%s</t></si>' % html.escape(t, quote=False) for t in new_strings
        )
        ss = ss.replace("</sst>", add + "</sst>")
        ss = re.sub(r'uniqueCount="\d+"', f'uniqueCount="{n0 + len(new_strings)}"', ss)
        files[ss_name] = ss.encode("utf8")

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx").name
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for n in order:
            zout.writestr(n, files[n])
    shutil.move(tmp, dst)
    print(f"已写出 {dst}。下一步：bash lo_recalc.sh {dst} <输出目录>")


if __name__ == "__main__":
    main()
