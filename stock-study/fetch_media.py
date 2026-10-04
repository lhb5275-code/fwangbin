#!/usr/bin/env python3
"""슬라이드용 이미지 내려받기 (위키미디어 공용·위키백과·일반 URL).

    # 위키미디어 공용 파일 (SVG는 PNG로 변환해 받는다)
    python fetch_media.py commons "Alphabet Inc Logo 2015.svg" companies/GOOGL/img/logo.png --width 800
    # 위키백과 문서의 대표 이미지
    python fetch_media.py wiki "Sundar Pichai" companies/GOOGL/img/ceo.jpg
    python fetch_media.py wiki "알파벳 (기업)" companies/GOOGL/img/hq.jpg --lang ko
    # 일반 URL (IR 자료 이미지 등)
    python fetch_media.py url https://example.com/a.png companies/GOOGL/img/a.png

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


def get(url: str) -> bytes:
    req = urllib.request.Request(url, headers={"User-Agent": UA})
    with urllib.request.urlopen(req, timeout=60) as r:
        return r.read()


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


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("kind", choices=["commons", "wiki", "url"])
    ap.add_argument("name", help="파일 이름 / 문서 제목 / URL")
    ap.add_argument("out", type=Path)
    ap.add_argument("--width", type=int, default=1200)
    ap.add_argument("--lang", default="en")
    a = ap.parse_args()
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
