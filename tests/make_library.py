#!/usr/bin/env python3
"""生成一个迷你研报库（纯手写 PDF，不依赖任何库），用于测试索引与检索。

目录结构故意混用"按机构"和"按日期"两种组织方式，并包含容易误判的路径：
  <root>/高盛/2026-10/6fd4a26e-GS_note-261021.pdf   文件夹年月 + 哈希前缀（日期应为 2026-10-21，不是 10-06）
  <root>/Bernstein/acme_quick_take.pdf               日期只在正文里
  <root>/2026/2026-09-30_HSBC_sector.pdf             按日期组织，提到 ACME 但不是主角
  <root>/JPM/other_company.pdf                       无关公司
"""
import os
import sys


def make_pdf(path, lines):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    esc = lambda s: s.replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)")
    content = "BT /F1 11 Tf 50 780 Td 14 TL\n" + "\n".join(f"({esc(l)}) Tj T*" for l in lines) + "\nET"
    objs = [
        "<< /Type /Catalog /Pages 2 0 R >>",
        "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
        "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 842] /Contents 4 0 R "
        "/Resources << /Font << /F1 5 0 R >> >> >>",
        f"<< /Length {len(content.encode('latin-1'))} >>\nstream\n{content}\nendstream",
        "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
    ]
    out = b"%PDF-1.4\n"
    offsets = []
    for i, o in enumerate(objs, 1):
        offsets.append(len(out))
        out += f"{i} 0 obj\n{o}\nendobj\n".encode("latin-1")
    xref = len(out)
    out += f"xref\n0 {len(objs) + 1}\n0000000000 65535 f \n".encode()
    out += "".join(f"{o:010d} 00000 n \n" for o in offsets).encode()
    out += f"trailer\n<< /Size {len(objs) + 1} /Root 1 0 R >>\nstartxref\n{xref}\n%%EOF\n".encode()
    with open(path, "wb") as f:
        f.write(out)


def main(root):
    filler = ["Our analysis of unit economics, backlog conversion and financing costs follows."] * 12
    make_pdf(os.path.join(root, "高盛", "2026-10", "6fd4a26e-GS_note-261021.pdf"), [
        "Equity Research", "21 October 2026",
        "Acme Compute Inc. (ACME): How big is the funding gap?",
        "ACME 12m Price Target: $106.00   Price: $61.82   Upside: 71.5%   BUY",
        "Goldman Sachs", "Acme residual value and GPU useful life drive returns.",
    ] + filler)
    make_pdf(os.path.join(root, "Bernstein", "acme_quick_take.pdf"), [
        "QUICK TAKE", "15 October 2026", "Rating", "Underperform", "Price Target", "ACME   74.00 USD",
        "Acme Compute (ACME): Debt-maxxing... growth at any price?",
        "Bernstein Research. Acme residual value is uncertain; DDTL spreads widened.",
    ] + filler)
    make_pdf(os.path.join(root, "2026", "2026-09-30_HSBC_sector.pdf"), [
        "HSBC Global Research", "US Technology Q2 takeaways: enterprise software returns",
        "Coverage includes many names; Acme Compute (ACME US, Reduce, TP USD39.00) is one of them.",
    ] + filler)
    make_pdf(os.path.join(root, "JPM", "other_company.pdf"), [
        "J.P. Morgan", "11 August 2026", "Widget Corp (WDGT): steady demand", "Neutral", "Price Target (Dec-27): $50.00",
        "Widget demand remains steady.",
    ] + filler)
    print("迷你研报库已生成：", root)


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "tests/_library")
