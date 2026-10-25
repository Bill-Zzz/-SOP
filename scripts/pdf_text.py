#!/usr/bin/env python3
"""统一的 PDF 文本抽取 / 页面渲染 / OCR，按可用工具自动回退，无需管理员权限也能工作。

文本：PyMuPDF（pip install pymupdf）→ pdftotext（poppler）→ pypdf（pip install pypdf）
渲染：PyMuPDF → pdftoppm（poppler）
OCR ：tesseract（命令行）→ RapidOCR（pip install rapidocr-onnxruntime，支持中文）

命令行：
  python3 pdf_text.py text  file.pdf [--max-pages N]      # 打印文本
  python3 pdf_text.py render file.pdf 页码 out.png [--dpi 150]
  python3 pdf_text.py ocr   file.pdf [--lang chi_sim+eng] # 整份 OCR
  python3 pdf_text.py check                               # 列出可用的后端
作为模块：from pdf_text import extract_text, page_count, render_page, ocr_pdf, backends
"""
import os
import shutil
import subprocess
import sys
import tempfile


def _has_mod(name):
    try:
        __import__(name)
        return True
    except Exception:
        return False


def backends():
    return {
        "pymupdf": _has_mod("fitz"),
        "pdftotext": shutil.which("pdftotext") is not None,
        "pdftoppm": shutil.which("pdftoppm") is not None,
        "pypdf": _has_mod("pypdf"),
        "tesseract": shutil.which("tesseract") is not None,
        "rapidocr": _has_mod("rapidocr_onnxruntime"),
    }


def page_count(path):
    b = backends()
    try:
        if b["pymupdf"]:
            import fitz
            with fitz.open(path) as d:
                return d.page_count
        if b["pypdf"]:
            import pypdf
            return len(pypdf.PdfReader(path).pages)
        if shutil.which("pdfinfo"):
            out = subprocess.run(["pdfinfo", path], capture_output=True, text=True).stdout
            for line in out.splitlines():
                if line.startswith("Pages:"):
                    return int(line.split()[1])
    except Exception:
        pass
    return None


def extract_text(path, max_pages=None):
    """返回 (文本, 使用的后端)。max_pages 只取前 N 页（用于快速建索引）。"""
    b = backends()
    if b["pymupdf"]:
        try:
            import fitz
            with fitz.open(path) as d:
                n = d.page_count if not max_pages else min(max_pages, d.page_count)
                return "\n\f".join(d.load_page(i).get_text() for i in range(n)), "pymupdf"
        except Exception:
            pass
    if b["pdftotext"]:
        cmd = ["pdftotext", "-layout"]
        if max_pages:
            cmd += ["-l", str(max_pages)]
        r = subprocess.run(cmd + [path, "-"], capture_output=True, text=True, errors="ignore")
        if r.returncode == 0:
            return r.stdout, "pdftotext"
    if b["pypdf"]:
        try:
            import pypdf
            pages = pypdf.PdfReader(path).pages
            n = len(pages) if not max_pages else min(max_pages, len(pages))
            return "\n\f".join((pages[i].extract_text() or "") for i in range(n)), "pypdf"
        except Exception:
            pass
    return "", "none"


def render_page(path, page, out_png, dpi=150):
    """渲染第 page 页（从 1 开始）为 PNG，返回输出路径。"""
    b = backends()
    if b["pymupdf"]:
        import fitz
        with fitz.open(path) as d:
            d.load_page(page - 1).get_pixmap(dpi=dpi).save(out_png)
        return out_png
    if b["pdftoppm"]:
        prefix = out_png[:-4] if out_png.endswith(".png") else out_png
        subprocess.run(["pdftoppm", "-r", str(dpi), "-png", "-singlefile", "-f", str(page), "-l", str(page),
                        path, prefix], check=True)
        return prefix + ".png"
    raise RuntimeError("无法渲染：请 pip install pymupdf 或安装 poppler-utils")


def _ocr_image(png, lang):
    b = backends()
    if b["tesseract"]:
        r = subprocess.run(["tesseract", png, "-", "-l", lang], capture_output=True, text=True, errors="ignore")
        if r.returncode == 0:
            return r.stdout
    if b["rapidocr"]:
        from rapidocr_onnxruntime import RapidOCR
        eng = _ocr_image.__dict__.setdefault("eng", RapidOCR())
        res, _ = eng(png)
        return "\n".join(item[1] for item in (res or []))
    raise RuntimeError("无法 OCR：请安装 tesseract，或 pip install rapidocr-onnxruntime")


def ocr_pdf(path, lang="eng", dpi=200, pages=None):
    """整份或指定页 OCR，返回 (文本, 后端)。"""
    n = page_count(path) or 0
    idx = pages or list(range(1, n + 1))
    tmp = tempfile.mkdtemp(prefix="ocr_")
    out = []
    try:
        for p in idx:
            png = render_page(path, p, os.path.join(tmp, f"p{p}.png"), dpi=dpi)
            out.append(_ocr_image(png, lang))
    finally:
        shutil.rmtree(tmp, ignore_errors=True)  # 只删除本函数创建的临时目录
    be = "tesseract" if backends()["tesseract"] else "rapidocr"
    return "\n\f".join(out), be


def looks_scanned(text, pages):
    """文本过少则视为扫描件/图片型 PDF。"""
    pages = pages or 1
    return len(text.strip()) < 200 * pages * 0.3


def main():
    a = sys.argv[1:]
    if not a or a[0] == "check":
        for k, v in backends().items():
            print(f"{k:<10} {'可用' if v else '—'}")
        return
    cmd, path = a[0], a[1]
    if cmd == "text":
        mp = int(a[a.index("--max-pages") + 1]) if "--max-pages" in a else None
        t, be = extract_text(path, mp)
        sys.stderr.write(f"[后端 {be}]\n")
        print(t)
    elif cmd == "render":
        dpi = int(a[a.index("--dpi") + 1]) if "--dpi" in a else 150
        print(render_page(path, int(a[2]), a[3], dpi))
    elif cmd == "ocr":
        lang = a[a.index("--lang") + 1] if "--lang" in a else "eng"
        t, be = ocr_pdf(path, lang)
        sys.stderr.write(f"[后端 {be}]\n")
        print(t)
    else:
        print(__doc__)


if __name__ == "__main__":
    main()
