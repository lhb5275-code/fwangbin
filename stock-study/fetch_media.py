#!/usr/bin/env python3
"""슬라이드용 이미지 내려받기 (위키미디어 공용·위키백과·일반 URL).

    # 위키미디어 공용 파일 (SVG는 PNG로 변환해 받는다)
    python fetch_media.py commons "Alphabet Inc Logo 2015.svg" companies/GOOGL/img/logo.png --width 800
    # 위키백과 문서의 대표 이미지
    python fetch_media.py wiki "Sundar Pichai" companies/GOOGL/img/ceo.jpg
    python fetch_media.py wiki "알파벳 (기업)" companies/GOOGL/img/hq.jpg --lang ko
    # 일반 URL (IR 자료 이미지 등)
    python fetch_media.py url https://example.com/a.png companies/GOOGL/img/a.png
    # PDF(주주서한·IR 자료)에 들어 있는 사진을 원본 해상도 그대로 추출
    python fetch_media.py pdf letter.pdf companies/NXT/img/letter

이미지는 기본 2560px 폭의 고화질로 받는다 (원본이 더 작으면 원본 크기).

받은 파일의 출처는 같은 폴더의 CREDITS.md에 한 줄씩 기록된다.
"""
from __future__ import annotations

import argparse
import json
import os
import urllib.parse
import urllib.request
from pathlib import Path

UA = os.environ.get("MEDIA_USER_AGENT", "fwangbin-stock-study/1.0 (https://github.com/lhb5275-code/fwangbin)")


def get(url: str, tries: int = 5) -> bytes:
    import time
    import urllib.error
    for i in range(tries):
        try:
            req = urllib.request.Request(url, headers={"User-Agent": UA})
            with urllib.request.urlopen(req, timeout=60) as r:
                return r.read()
        except urllib.error.HTTPError as e:
            if e.code != 429 or i == tries - 1:   # 위키미디어 요청 한도 초과면 잠시 쉬었다 재시도
                raise
            time.sleep(5 * (i + 1))


def credit(out: Path, source: str, note: str = ""):
    f = out.parent / "CREDITS.md"
    line = f"- `{out.name}` ← {source}" + (f" ({note})" if note else "") + "\n"
    old = f.read_text(encoding="utf-8") if f.exists() else "# 이미지 출처\n\n"
    if line not in old:
        f.write_text(old + line, encoding="utf-8")


def commons(name: str, out: Path, width: int):
    name = name.removeprefix("File:").removeprefix("파일:")
    url = ("https://commons.wikimedia.org/w/index.php?title=Special:FilePath/"
           + urllib.parse.quote(name) + f"&width={width}")
    out.write_bytes(get(url))
    credit(out, "https://commons.wikimedia.org/wiki/File:" + urllib.parse.quote(name.replace(" ", "_")))


def wiki(title: str, out: Path, width: int, lang: str):
    api = (f"https://{lang}.wikipedia.org/w/api.php?action=query&format=json&prop=pageimages&redirects=1"
           f"&piprop=thumbnail|name&pithumbsize={width}&titles=" + urllib.parse.quote(title))
    pages = json.loads(get(api))["query"]["pages"]
    page = next(iter(pages.values()))
    thumb = page.get("thumbnail", {}).get("source")
    if not thumb:
        raise SystemExit(f"'{title}' 문서에 대표 이미지가 없습니다.")
    out.write_bytes(get(thumb))
    fname = page.get("pageimage", "")
    credit(out, f"https://{lang}.wikipedia.org/wiki/{urllib.parse.quote(title.replace(' ', '_'))}",
           f"파일: {fname}" if fname else "")


def extract_pdf(pdf: Path, outdir: Path, min_px: int = 400):
    """PDF에 들어 있는 사진을 원본 해상도로 꺼낸다.

    JPEG로 들어 있는 사진은 원래 바이트 그대로(.jpg), 나머지는 무손실 PNG로 저장한다.
    마스크(smask)와 작은 아이콘은 버린다.
    """
    import subprocess
    outdir.mkdir(parents=True, exist_ok=True)
    listing = subprocess.run(["pdfimages", "-list", str(pdf)], capture_output=True, text=True, check=True).stdout
    keep = {}
    for line in listing.splitlines()[2:]:
        f = line.split()
        if len(f) > 4 and f[2] == "image" and min(int(f[3]), int(f[4])) >= min_px:
            keep[int(f[1])] = (int(f[0]), int(f[3]), int(f[4]))
    prefix = outdir / "p"
    subprocess.run(["pdfimages", "-png", "-j", "-p", str(pdf), str(prefix)], check=True)
    for f in sorted(outdir.glob("p-*")):
        num = int(f.stem.split("-")[-1])
        if num not in keep:
            f.unlink()
            continue
        page, w, h = keep[num]
        print(f"저장: {f}  ({page}쪽, {w}x{h})")
    credit(outdir / "p-*", pdf.name, "PDF 내장 이미지, 원본 해상도")


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("kind", choices=["commons", "wiki", "url", "pdf"])
    ap.add_argument("name", help="파일 이름 / 문서 제목 / URL / PDF 경로")
    ap.add_argument("out", type=Path)
    ap.add_argument("--width", type=int, default=2560, help="받을 폭(px), 기본 2560 (고화질, 위키미디어 표준 썸네일 크기)")
    ap.add_argument("--lang", default="en")
    a = ap.parse_args()
    if a.kind == "pdf":
        return extract_pdf(Path(a.name), a.out)
    a.out.parent.mkdir(parents=True, exist_ok=True)
    if a.kind == "commons":
        commons(a.name, a.out, a.width)
    elif a.kind == "wiki":
        wiki(a.name, a.out, a.width, a.lang)
    else:
        a.out.write_bytes(get(a.name))
        credit(a.out, a.name)
    print(f"저장: {a.out} ({a.out.stat().st_size:,} bytes)")


if __name__ == "__main__":
    main()
