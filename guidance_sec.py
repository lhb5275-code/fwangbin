"""SEC 8-K 실적 보도자료(Item 2.02, Exhibit 99)에서 다음 분기 매출 가이던스를 추출하는 도구.

흐름:
    1) data.sec.gov/submissions 로 회사의 8-K 목록(Item 2.02 = 실적 발표)을 가져온다.
    2) 직전 분기(Q-1) 실적 발표문 = 최근 분기(Q0)의 가이던스가 들어 있는 문서를 고른다.
    3) Exhibit 99.x 본문에서 '매출 + 전망 표현 + 금액'이 함께 있는 문장/표 행을 찾아 금액 범위를 파싱한다.
    4) 분기 규모와 맞지 않는 값(연간 가이던스, 부문 매출 등)은 Q0 실제 매출과의 비율로 걸러낸다.

추출은 규칙 기반이라 오류가 있을 수 있으므로, 결과에 근거 문장(snippet)을 함께 남긴다.
"""

import html
import re
import time

import requests

SEC_UA = {"User-Agent": "sp500-last5q research admin@example.com"}
SUBMISSIONS = "https://data.sec.gov/submissions/CIK{cik:010d}.json"
INDEX_HTM = "https://www.sec.gov/Archives/edgar/data/{cik}/{acc}/{acc_dash}-index.htm"
DOC_URL = "https://www.sec.gov/Archives/edgar/data/{cik}/{acc}/{name}"
TICKERS_JSON = "https://www.sec.gov/files/company_tickers.json"

_session = requests.Session()
_session.headers.update(SEC_UA)


def sec_get(url, tries=3):
    for a in range(tries):
        try:
            r = _session.get(url, timeout=60)
            if r.status_code == 200:
                time.sleep(0.12)  # SEC 권장 10 req/s 이하
                return r
            if r.status_code == 404:
                return None
        except requests.RequestException:
            pass
        time.sleep(1.5 * (a + 1))
    return None


def cik_map():
    r = sec_get(TICKERS_JSON)
    return {v["ticker"].upper().replace(".", "-"): int(v["cik_str"]) for v in r.json().values()}


def earnings_8ks(cik):
    """[(filing_date, accession, primary_doc)] - Item 2.02 가 포함된 8-K (최신순)."""
    r = sec_get(SUBMISSIONS.format(cik=cik))
    if r is None:
        return []
    rec = r.json().get("filings", {}).get("recent", {})
    out = []
    for form, date, acc, items, doc in zip(
        rec.get("form", []), rec.get("filingDate", []), rec.get("accessionNumber", []),
        rec.get("items", []), rec.get("primaryDocument", []),
    ):
        if form in ("8-K", "8-K/A") and "2.02" in (items or ""):
            out.append((date, acc, doc))
    return out


def html_to_text(raw):
    t = re.sub(r"(?is)<(script|style).*?</\1>", " ", raw)
    t = re.sub(r"(?i)</t[dh]\s*>", " | ", t)
    t = re.sub(r"(?i)<br\s*/?>|</p\s*>|</div\s*>|</tr\s*>|</li\s*>|</h\d\s*>", "\n", t)
    t = re.sub(r"<[^>]+>", " ", t)
    t = html.unescape(t).replace("\xa0", " ")
    t = re.sub(r"[ \t\r\f\v]+", " ", t)
    t = re.sub(r"( \| )+", " | ", t)
    t = re.sub(r"\n\s*\n+", "\n", t)
    return t


TEXT_CACHE_DIR = None  # 지정하면 Exhibit 본문을 {dir}/{cik}_{acc}.json 으로 저장·재사용


def exhibit_texts(cik, acc):
    """8-K 의 Exhibit 99.x 문서 본문 텍스트 목록 [(문서명, 텍스트)]."""
    import json
    import os

    cpath = os.path.join(TEXT_CACHE_DIR, f"{cik}_{acc}.json") if TEXT_CACHE_DIR else None
    if cpath and os.path.exists(cpath):
        return [tuple(x) for x in json.load(open(cpath, encoding="utf-8"))]
    acc_nd = acc.replace("-", "")
    r = sec_get(INDEX_HTM.format(cik=cik, acc=acc_nd, acc_dash=acc))
    if r is None:
        return []
    names = []
    # 제출 색인표의 각 행: <a href=".../문서명">문서명</a> ... <td>EX-99.1</td>
    for row in re.findall(r"(?is)<tr[^>]*>(.*?)</tr>", r.text):
        m = re.search(r'(?i)href="[^"]*/([^"/]+\.(?:htm|html|txt))"', row)
        if m and re.search(r"(?i)>\s*EX-99[\.\d]*\s*<", row):
            names.append(m.group(1))
    texts = []
    for n in names[:4]:
        d = sec_get(DOC_URL.format(cik=cik, acc=acc_nd, name=n))
        if d is not None:
            texts.append((n, html_to_text(d.text)))
    if cpath:
        os.makedirs(TEXT_CACHE_DIR, exist_ok=True)
        json.dump(texts, open(cpath, "w", encoding="utf-8"), ensure_ascii=False)
    return texts


# ---------------------------------------------------------------- 가이던스 파싱
_NUM = r"\d{1,3}(?:,\d{3})*(?:\.\d+)?|\d+(?:\.\d+)?"
_UNIT = r"(?:(billion|million|thousand|bn|mn|B|M)\b)"
_MONEY = rf"\$\s?({_NUM})\s*{_UNIT}?"
_REV_WORDS = r"(?i)\b(total\s+)?(net\s+)?(revenues?|sales)\b"
# 앞으로의 값을 말하는 표현 (문장 안에서 이 위치 이후 금액만 가이던스로 본다)
_FWD = re.compile(r"(?i)\b(expects?|expected|expecting|anticipates?|anticipated|forecast(?:s|ing)?|projects?|projected|"
                  r"guid(?:e|es|ed|ing|ance)|outlook|to be (?:in the range|between|approximately)|in the range of|"
                  r"plus or minus)\b")
_QTR_WORDS = re.compile(r"(?i)\b(quarter|q[1-4]|three months)\b")
_ANNUAL = re.compile(r"(?i)\b(full[- ]year|fiscal year|annual|full fiscal|calendar year|fy\s?\d{2,4}|"
                     r"(?:for|in|of) (?:fiscal )?20\d\d)\b")
# 과거 실적을 말하는 표현
_PAST = re.compile(r"(?i)\b(versus|compared (to|with)|prior quarter|same (period|quarter) (of )?last year|a year ago|"
                   r"year[- ]ago|was|were|increased|decreased|grew|declined|record|reported|delivered|achieved|"
                   r"representing|up \d+(\.\d+)?\s?(%|percent)|down \d+(\.\d+)?\s?(%|percent)|higher than|lower than|"
                   r"year[- ]over[- ]year|up from|raised|raising|reaffirm\w*|maintain\w*)\b")
# 매출 앞에 오면 '전체 매출'로 보는 단어 (그 외 명사가 오면 부문·제품 매출로 간주)
_REV_PREFIX_OK = re.compile(r"(?i)^(total|net|consolidated|gaap|quarter|quarterly|q[1-4]|first|second|third|fourth|"
                            r"fiscal|and|of|the|our|its|in|for|to|expects?|expected|anticipates?|a|an|with|be|is|are|"
                            r"company|group|\d{4}|(first|second|third|fourth)-quarter|•|▪|l|-|–|\(|revenue|sales)$")
_GROWTH = re.compile(r"(?i)(?:grow\w*|increase\w*|up)\s+(?:by\s+)?(?:approximately\s+|about\s+|in the range of\s+|between\s+)?"
                     r"(-?\d+(?:\.\d+)?)\s?%\s*(?:to|-|–|and)\s*(-?\d+(?:\.\d+)?)\s?%")

_UNIT_MULT = {"billion": 1e9, "bn": 1e9, "b": 1e9, "million": 1e6, "mn": 1e6, "m": 1e6, "thousand": 1e3}


def _to_usd(num, unit, default_unit=None):
    v = float(num.replace(",", ""))
    u = (unit or default_unit or "").lower()
    return v * _UNIT_MULT.get(u, 1.0)


def _parse_range(s, default_unit=None):
    """문장에서 (low, high, kind) 를 파싱. 없으면 None. default_unit: 표 머리의 '(In millions)' 등."""
    m = re.search(rf"{_MONEY}\s*(?:,\s*)?(?:±|\+/-|\+/−|plus or minus)\s*(\$\s?({_NUM})\s*{_UNIT}?|({_NUM})\s?%)", s, re.I)
    if m:
        center = _to_usd(m.group(1), m.group(2), default_unit)
        if m.group(6):
            pct = float(m.group(6)) / 100
            return center * (1 - pct), center * (1 + pct), "±%"
        delta = _to_usd(m.group(4), m.group(5) or m.group(2), default_unit)
        return center - delta, center + delta, "±$"
    m = re.search(rf"\$\s?({_NUM})\s*{_UNIT}?\s*(?:to|-|–|—|and)\s*\$?\s?({_NUM})\s*{_UNIT}?", s, re.I)
    if m:
        u2 = m.group(4) or m.group(2)
        lo, hi = _to_usd(m.group(1), m.group(2) or u2, default_unit), _to_usd(m.group(3), u2, default_unit)
        if lo <= hi:
            return lo, hi, "range"
    m = re.search(rf"(?:approximately|about|around|roughly|of|be|to)\s+{_MONEY}", s, re.I)
    if m and (m.group(2) or default_unit):
        v = _to_usd(m.group(1), m.group(2), default_unit)
        return v, v, "point"
    return None


def _is_heading(u):
    return (len(u) < 70 and "$" not in u and re.search(r"(?i)\b(outlook|guidance|looking ahead|financial targets?)\b", u)
            and not re.search(r"(?i)\b(raises?|maintains?|reaffirms?|updates?|announces?|reports?|provides?)\b", u))


def candidate_spans(text):
    """가이던스 후보 [(문장, Outlook 구간 여부, 기본 단위)].

    - 표 셀 구분자(|)는 공백으로 바꾸고, '매출' 라벨만 있는 줄은 다음 줄(값)과 합친다.
    - 'Outlook/Guidance' 소제목 이후 15개 단위는 전망 표현 없이도 후보로 보며, '(In millions)' 표기를 기본 단위로 쓴다.
    """
    raw = [u.strip() for u in re.split(r"(?<=[.;])\s+(?=[A-Z])|\n", text)]
    units = [re.sub(r"\s*\|\s*", " ", u).strip() for u in raw]
    units = [u for u in units if u]
    spans, in_outlook, unit = [], 0, None
    for i, u in enumerate(units):
        if _is_heading(u):
            in_outlook, unit = 15, None
        if in_outlook:
            mu = re.search(r"(?i)in (millions|billions|thousands)", u)
            if mu:
                unit = {"millions": "million", "billions": "billion", "thousands": "thousand"}[mu.group(1).lower()]
        cand = u
        if re.search(_REV_WORDS, u) and "$" not in u and len(u) < 80:
            for j in (1, 2):  # 라벨 줄 + 값 줄
                if i + j < len(units):
                    cand = cand + " " + units[i + j]
                    if "$" in units[i + j]:
                        break
        if re.search(_REV_WORDS, cand) and "$" in cand and (in_outlook > 0 or _FWD.search(cand)):
            spans.append((cand[:600], in_outlook > 0, unit if in_outlook else None))
        if in_outlook:
            in_outlook -= 1
    return spans


def _is_segment(s):
    """'Product revenue', 'Electrification: Revenue' 처럼 매출 앞에 부문명이 붙으면 True."""
    m = re.search(r"(?i)\b(revenues?|sales)\b", s)
    if not m:
        return False
    before = s[: m.start()].rstrip()
    if re.search(r"[A-Za-z][\w&' -]{2,40}\s*:\s*$", before) and not re.search(r"(?i)(total|net)\s*:\s*$", before):
        return True
    words = before.split()
    return bool(words) and not _REV_PREFIX_OK.match(words[-1].strip(",;:"))


def extract_revenue_guidance(texts, actual_rev, prior_year_rev=None):
    """후보 중 Q0 실제 매출 규모와 맞는 분기 전체 매출 가이던스를 고른다.

    반환 dict(low, high, mid, kind, snippet, doc, flags, score) 또는 None.
    """
    best, best_score = None, -1e9
    for doc, text in texts:
        for s, in_sec, dunit in candidate_spans(text):
            if _ANNUAL.search(s) and not _QTR_WORDS.search(s):
                continue  # 연간 가이던스
            if _is_segment(s):
                continue  # 부문·제품 매출
            fwd = _FWD.search(s)
            past = _PAST.search(s)
            body = s
            if past:
                if not fwd:
                    continue  # 과거 실적 문장
                body = s[fwd.start():]  # 실적+가이던스가 섞인 문장: 전망 표현 뒤만 사용
            elif not fwd and not in_sec:
                continue
            rng = _parse_range(body, dunit)
            kind_extra = ""
            if rng is None and prior_year_rev and _GROWTH.search(body) and _QTR_WORDS.search(s):
                g = _GROWTH.search(body)
                g1, g2 = sorted((float(g.group(1)), float(g.group(2))))
                rng = (prior_year_rev * (1 + g1 / 100), prior_year_rev * (1 + g2 / 100), "growth%")
                kind_extra = "전년 동기 매출 × 가이던스 성장률"
            if rng is None:
                continue
            lo, hi, kind = rng
            mid = (lo + hi) / 2
            if not actual_rev or mid <= 0:
                continue
            ratio = actual_rev / mid
            if not 0.6 <= ratio <= 1.6:  # 연간·부문·단위 오류 배제
                continue
            score = 0.0
            score += 3 if _QTR_WORDS.search(s) else 0
            score += 2 if in_sec else 0
            score += 2 if kind in ("±%", "±$", "range") else 0
            score += 1 if re.search(r"(?i)\btotal (net )?revenues?\b", s) else 0
            score -= 2 if past else 0
            score -= abs(ratio - 1)  # 동점이면 실제 매출에 더 가까운(보수적) 후보
            if score > best_score:
                best_score = score
                best = {
                    "low": lo, "high": hi, "mid": mid, "kind": kind, "snippet": s, "doc": doc,
                    "flags": ";".join(f for f in (kind_extra, "분기 표현 없음" if not _QTR_WORDS.search(s) else "",
                                                  "실적·전망 혼합 문장" if past else "") if f),
                    "score": round(score, 2),
                }
    return best
