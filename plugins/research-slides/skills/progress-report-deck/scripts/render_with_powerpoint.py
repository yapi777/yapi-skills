#!/usr/bin/env python3
"""PowerPoint 本体で pptx を PDF 化し、ページ画像（PNG）に落とす。

LibreOffice 版（render_slides.py）より遅いが、**本番と同じ見え方** が確認できる。
とくに次のケースは LibreOffice では正しく確認できないので、必ずこちらを使う:
  - OMML 数式（LibreOffice は mc:Fallback のプレーンテキストを表示してしまう）
  - フォント埋め込み・自動採番・テーマ色の細かい差

使い方:
    python render_with_powerpoint.py <pptx> <outdir> [--dpi 110] [--pages 1,7,13]

注意: Box / OneDrive 同期フォルダ上のファイルを直接開くと COM が E_FAIL を返すこと
があるため、必ずローカル一時フォルダへコピーしてから開く（本スクリプトは自動で行う）。
"""
import argparse
import os
import shutil
import subprocess
import sys
import tempfile
import uuid

PS_TEMPLATE = r"""
$ErrorActionPreference = "Stop"
$src = "{src}"
$out = "{out}"
if (Test-Path $out) {{ Remove-Item $out -Force }}
$ppt = New-Object -ComObject PowerPoint.Application
try {{
  $pres = $ppt.Presentations.Open($src, $true, $false, $false)
  $pres.SaveAs($out, 32)
  $pres.Close()
  Write-Output "EXPORTED"
}} finally {{
  $ppt.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
}}
"""


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pptx")
    ap.add_argument("outdir")
    ap.add_argument("--dpi", type=int, default=110)
    ap.add_argument("--pages", default="", help="1始まりのページ番号をカンマ区切りで（既定は全ページ）")
    args = ap.parse_args()

    src = os.path.abspath(args.pptx)
    if not os.path.exists(src):
        sys.exit(f"not found: {src}")

    tmpdir = os.path.join(tempfile.gettempdir(), "pptrender-" + uuid.uuid4().hex[:8])
    os.makedirs(tmpdir, exist_ok=True)
    local_pptx = os.path.join(tmpdir, "deck.pptx")
    local_pdf = os.path.join(tmpdir, "deck.pdf")
    shutil.copyfile(src, local_pptx)

    ps = PS_TEMPLATE.format(src=local_pptx, out=local_pdf)
    proc = subprocess.run(
        ["powershell", "-NoProfile", "-NonInteractive", "-Command", ps],
        capture_output=True, text=True,
    )
    if "EXPORTED" not in proc.stdout:
        shutil.rmtree(tmpdir, ignore_errors=True)
        sys.exit(f"PowerPoint export failed:\n{proc.stdout}\n{proc.stderr}")

    import pymupdf  # noqa: E402

    os.makedirs(args.outdir, exist_ok=True)
    doc = pymupdf.open(local_pdf)
    wanted = None
    if args.pages.strip():
        wanted = {int(p) for p in args.pages.split(",") if p.strip()}

    written = []
    for i in range(doc.page_count):
        if wanted and (i + 1) not in wanted:
            continue
        path = os.path.join(args.outdir, f"slide-{i + 1}.png")
        doc[i].get_pixmap(dpi=args.dpi).save(path)
        written.append(path)
    doc.close()
    shutil.rmtree(tmpdir, ignore_errors=True)

    for p in written:
        print(p)


if __name__ == "__main__":
    main()
