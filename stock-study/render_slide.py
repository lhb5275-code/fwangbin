#!/usr/bin/env python3
"""지난 덱의 슬라이드 한 장을 고화질 이미지로 저장 (F/U 덱의 '지난 분석 회고' 슬라이드용).

    python render_slide.py output/NXT_기업분석.pptx companies/NXT/img/prev_cover.png --slide 1
    python render_slide.py output/NXT_기업분석.pptx companies/NXT/img/prev_price.png --slide 2 --dpi 220

LibreOffice(soffice)와 pdftoppm이 필요하다. 미리보기 글꼴은 설치된 글꼴로 대체될 수 있다.
"""
from __future__ import annotations

import argparse
import shutil
import subprocess
import tempfile
from pathlib import Path


def render(pptx: Path, out: Path, slide: int = 1, dpi: int = 200):
    with tempfile.TemporaryDirectory() as td:
        tmp = Path(td)
        src = tmp / "deck.pptx"
        shutil.copy(pptx, src)
        profile = (tmp / "profile").as_uri()
        subprocess.run(["soffice", f"-env:UserInstallation={profile}", "--headless", "--convert-to", "pdf",
                        "--outdir", str(tmp), str(src)], check=True, capture_output=True,
                       env={"SAL_USE_VCLPLUGIN": "svp", "PATH": "/usr/bin:/bin:/usr/local/bin", "HOME": td})
        subprocess.run(["pdftoppm", "-png", "-r", str(dpi), "-f", str(slide), "-l", str(slide), "-singlefile",
                        str(tmp / "deck.pdf"), str(tmp / "page")], check=True)
        out.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy(tmp / "page.png", out)
    print(f"저장: {out}")


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("pptx", type=Path)
    ap.add_argument("out", type=Path)
    ap.add_argument("--slide", type=int, default=1)
    ap.add_argument("--dpi", type=int, default=200)
    a = ap.parse_args()
    render(a.pptx, a.out, a.slide, a.dpi)


if __name__ == "__main__":
    main()
