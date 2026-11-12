#!/usr/bin/env python3
"""
Render every slide of a .pptx to a JPEG image, for visual QA.

This machine has LibreOffice but no Poppler (pdftoppm), so this uses
PyMuPDF (pymupdf) to rasterize the intermediate PDF instead — no extra
system dependency beyond what's already installed for this skill.

Usage:
    python render_slides.py deck.pptx [out_dir] [--dpi 150]

Writes out_dir/slide-1.jpg, slide-2.jpg, ... (out_dir defaults to a
"<deck-stem>-slides" folder next to the pptx) and prints their absolute
paths, one per line, so they can be fed straight to an image-viewing tool.
"""
import argparse
import pathlib
import subprocess
import sys

def _find_soffice():
    import os
    import shutil
    found = shutil.which("soffice") or shutil.which("libreoffice")
    if found:
        return found
    for c in (r"C:\Program Files\LibreOffice\program\soffice.exe",
              r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
              "/Applications/LibreOffice.app/Contents/MacOS/soffice"):
        if os.path.exists(c):
            return c
    sys.exit("LibreOffice (soffice) not found. Install it or add it to PATH.")


SOFFICE = _find_soffice()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pptx", type=pathlib.Path)
    ap.add_argument("out_dir", nargs="?", type=pathlib.Path, default=None)
    ap.add_argument("--dpi", type=int, default=150)
    args = ap.parse_args()

    pptx_path = args.pptx.resolve()
    if not pptx_path.exists():
        sys.exit(f"not found: {pptx_path}")

    out_dir = (args.out_dir or pptx_path.with_name(pptx_path.stem + "-slides")).resolve()
    out_dir.mkdir(parents=True, exist_ok=True)

    # 1. pptx -> pdf via headless LibreOffice, into out_dir
    # encoding="utf-8", errors="replace": soffice's own stderr chatter is UTF-8,
    # but subprocess defaults to the console's codepage (cp932 on ja-JP Windows)
    # and throws UnicodeDecodeError partway through capture if we let it guess.
    result = subprocess.run(
        [SOFFICE, "--headless", "--norestore", "--convert-to", "pdf",
         "--outdir", str(out_dir), str(pptx_path)],
        capture_output=True, text=True, encoding="utf-8", errors="replace",
    )
    if result.returncode != 0:
        sys.exit(f"soffice failed:\n{result.stdout}\n{result.stderr}")

    pdf_path = out_dir / (pptx_path.stem + ".pdf")
    if not pdf_path.exists():
        sys.exit(f"expected PDF not found: {pdf_path}\nsoffice said:\n{result.stdout}")

    # 2. pdf -> one jpg per page via PyMuPDF
    import pymupdf  # imported here so --help works even before pymupdf is installed

    doc = pymupdf.open(pdf_path)
    paths = []
    for i, page in enumerate(doc, start=1):
        pix = page.get_pixmap(dpi=args.dpi)
        img_path = out_dir / f"slide-{i}.jpg"
        pix.save(img_path)
        paths.append(img_path)

    for p in paths:
        print(p)


if __name__ == "__main__":
    main()
