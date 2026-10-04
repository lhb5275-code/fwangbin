#!/usr/bin/env python3
"""기업분석용 기초 데이터 수집기.

SEC EDGAR(XBRL companyfacts, 제출 목록)와 Yahoo Finance 주가를 받아
companies/<TICKER>/data/ 아래에 data.json, summary.md, 원본 JSON을 저장한다.

    python fetch_data.py GOOGL
    python fetch_data.py NVDA --out companies/NVDA/data

summary.md에는 슬라이드에 바로 옮겨 쓸 수 있도록 '억 달러' 단위로 정리한
연간·분기 실적, 재무상태, 현금흐름, 주가·밸류에이션 지표가 들어간다.
애널리스트 의견·목표주가, IR 자료, 어닝콜 내용은 이 스크립트가 아니라
웹 검색으로 따로 조사한다 (SKILL.md 참고).
"""
from __future__ import annotations

import argparse
import datetime as dt
import json
import os
import sys
import time
import urllib.request
from collections import defaultdict
from pathlib import Path

# SEC는 연락처(이메일) 형식의 User-Agent를 요구한다. 실제 연락처로 바꾸려면 SEC_USER_AGENT 환경변수를 지정.
UA = os.environ.get("SEC_USER_AGENT", "fwangbin-stock-study research@example.com")
# Yahoo는 일반 브라우저 UA를 오히려 429로 막는 경우가 있어 연구용 UA를 그대로 쓴다.
YAHOO_UA = os.environ.get("YAHOO_USER_AGENT", "fwangbin-stock-study research@example.com")

# 지표별 XBRL 태그 후보 (앞쪽 우선, 가장 최근 값이 있는 태그를 쓴다)
FLOW = {
    "revenue": ["RevenueFromContractWithCustomerExcludingAssessedTax", "Revenues", "SalesRevenueNet",
                "RevenueFromContractWithCustomerIncludingAssessedTax"],
    "cost_of_revenue": ["CostOfRevenue", "CostOfGoodsAndServicesSold"],
    "gross_profit": ["GrossProfit"],
    "rnd": ["ResearchAndDevelopmentExpense"],
    "operating_income": ["OperatingIncomeLoss"],
    "net_income": ["NetIncomeLoss", "ProfitLoss"],
    "cfo": ["NetCashProvidedByUsedInOperatingActivities"],
    "cfi": ["NetCashProvidedByUsedInInvestingActivities"],
    "cff": ["NetCashProvidedByUsedInFinancingActivities"],
    "capex": ["PaymentsToAcquirePropertyPlantAndEquipment", "PaymentsToAcquireProductiveAssets"],
    "acquisitions": ["PaymentsToAcquireBusinessesNetOfCashAcquired"],
    "buyback": ["PaymentsForRepurchaseOfCommonStock"],
    "dividends_paid": ["PaymentsOfDividends", "PaymentsOfDividendsCommonStock"],
    "debt_issued": ["ProceedsFromIssuanceOfLongTermDebt", "ProceedsFromIssuanceOfDebt"],
}
PER_SHARE = {
    "eps_diluted": ["EarningsPerShareDiluted"],
    "eps_basic": ["EarningsPerShareBasic"],
    "dps": ["CommonStockDividendsPerShareDeclared", "CommonStockDividendsPerShareCashPaid"],
}
SHARES = {"diluted_shares": ["WeightedAverageNumberOfDilutedSharesOutstanding"]}
INSTANT = {
    "assets": ["Assets"],
    "current_assets": ["AssetsCurrent"],
    "liabilities": ["Liabilities"],
    "current_liabilities": ["LiabilitiesCurrent"],
    "equity": ["StockholdersEquity", "StockholdersEquityIncludingPortionAttributableToNoncontrollingInterest"],
    "liab_and_equity": ["LiabilitiesAndStockholdersEquity"],
    "cash": ["CashAndCashEquivalentsAtCarryingValue"],
    "marketable_securities": ["MarketableSecuritiesCurrent", "AvailableForSaleSecuritiesDebtSecuritiesCurrent"],
    "ppe": ["PropertyPlantAndEquipmentNet"],
    "long_term_debt": ["LongTermDebtNoncurrent", "LongTermDebt"],
}
LABELS = {
    "revenue": "매출", "cost_of_revenue": "매출원가", "gross_profit": "매출총이익", "rnd": "연구개발비",
    "operating_income": "영업이익", "net_income": "당기순이익", "cfo": "영업활동 현금흐름",
    "cfi": "투자활동 현금흐름", "cff": "재무활동 현금흐름", "capex": "CapEx(유형자산 취득)",
    "acquisitions": "인수합병", "buyback": "자사주 매입", "dividends_paid": "배당금 지급",
    "debt_issued": "장기부채 발행", "eps_diluted": "희석 EPS", "eps_basic": "기본 EPS", "dps": "주당배당금",
    "diluted_shares": "희석 주식수", "assets": "총자산", "current_assets": "유동자산", "liabilities": "총부채",
    "current_liabilities": "유동부채", "equity": "총자본", "cash": "현금및현금성자산",
    "marketable_securities": "단기 유가증권", "ppe": "유형자산", "long_term_debt": "장기부채",
    "fcf": "잉여현금흐름(FCF)",
}


def get(url: str, ua: str = UA, tries: int = 4) -> bytes:
    last = None
    for i in range(tries):
        try:
            req = urllib.request.Request(url, headers={"User-Agent": ua, "Accept-Encoding": "identity"})
            with urllib.request.urlopen(req, timeout=40) as r:
                return r.read()
        except Exception as e:  # 네트워크 오류·429는 잠시 후 재시도
            last = e
            time.sleep(2 ** (i + 1))
    raise RuntimeError(f"다운로드 실패: {url} ({last})")


def jget(url: str, ua: str = UA):
    return json.loads(get(url, ua))


def d(s: str) -> dt.date:
    return dt.date.fromisoformat(s)


# ---------------------------------------------------------------- SEC
def lookup_cik(ticker: str) -> tuple[str, str]:
    data = jget("https://www.sec.gov/files/company_tickers.json")
    t = ticker.upper().replace(".", "-")
    for row in data.values():
        if row["ticker"].upper() == t:
            return f"{int(row['cik_str']):010d}", row["title"]
    raise SystemExit(f"SEC에서 티커 {ticker}를 찾지 못했습니다 (미국 상장 기업만 지원).")


def fy_end_month(fiscal_year_end: str) -> int:
    # "1231", "0126" 같은 값. 52/53주 회계연도는 월말 근처 날짜라 10일 당겨서 월을 정한다.
    m, dd = int(fiscal_year_end[:2]), int(fiscal_year_end[2:])
    return (dt.date(2001, m, min(dd, 28)) - dt.timedelta(days=10)).month


def fiscal_label(end: dt.date, fy_month: int) -> tuple[int, int]:
    e = (end - dt.timedelta(days=10)).month
    q = ((e - fy_month) % 12) // 3
    q = 4 if q == 0 else q
    fy = (end - dt.timedelta(days=10)).year + (1 if e > fy_month else 0)
    if fy_month == 12:
        fy = (end - dt.timedelta(days=10)).year
    return fy, q


def pick_tag(facts: dict, tags: list[str], unit: str):
    """후보 태그 중 가장 최근 값이 있는 태그를 주 태그로 고르고,
    주 태그에 없는 기간은 다른 후보 태그 값으로 채운다 (회사가 태그를 바꾼 경우 대비)."""
    found = []
    for t in tags:
        node = facts.get("us-gaap", {}).get(t) or facts.get("dei", {}).get(t)
        if not node or unit not in node.get("units", {}):
            continue
        rows = node["units"][unit]
        found.append((max(r["end"] for r in rows), t, rows))
    if not found:
        return None
    found.sort(key=lambda x: x[0], reverse=True)
    main_tag = found[0][1]
    merged, seen = [], set()
    for _, _, rows in found:
        keys = {(r.get("start"), r["end"]) for r in rows}
        merged += [r for r in rows if (r.get("start"), r["end"]) not in seen]
        seen |= keys
    return main_tag, merged


def dedupe(rows: list[dict], key) -> dict:
    """같은 기간의 값이 여러 공시에 있으면 가장 최근 공시 값을 쓴다."""
    out = {}
    for r in sorted(rows, key=lambda r: r["filed"]):
        out[key(r)] = r
    return out


def flow_series(rows: list[dict], fy_month: int, additive: bool = True):
    """연간/분기 시계열. 10-Q의 누적(YTD) 값에서 분기 값을 역산한다."""
    dur = defaultdict(dict)  # start -> {end: val}
    direct_q = {}
    annual = {}
    for (s, e), r in dedupe(rows, lambda r: (r.get("start"), r["end"])).items():
        if not s:
            continue
        days = (d(e) - d(s)).days
        dur[s][e] = r["val"]
        if 80 <= days <= 100:
            direct_q[e] = r["val"]
        elif 350 <= days <= 380:
            annual[e] = r["val"]
    quarters = dict(direct_q)
    if additive:
        for s, ends in dur.items():
            seq = sorted(ends.items())
            prev_e, prev_v = None, None
            for e, v in seq:
                days = (d(e) - d(s)).days
                if prev_e is not None and e not in quarters:
                    gap = (d(e) - d(prev_e)).days
                    if 80 <= gap <= 100 and days <= 380:
                        quarters[e] = v - prev_v
                prev_e, prev_v = e, v
    annual_l = {}
    for e, v in annual.items():
        fy, _ = fiscal_label(d(e), fy_month)
        annual_l[fy] = {"end": e, "value": v}
    q_l = {}
    for e, v in quarters.items():
        fy, q = fiscal_label(d(e), fy_month)
        q_l[f"FY{fy}Q{q}"] = {"end": e, "value": v, "fy": fy, "q": q}
    return annual_l, q_l, dur


def instant_series(rows: list[dict], fy_month: int):
    out = {}
    for e, r in dedupe(rows, lambda r: r["end"]).items():
        fy, q = fiscal_label(d(e), fy_month)
        out[f"FY{fy}Q{q}"] = {"end": e, "value": r["val"], "fy": fy, "q": q}
    return out


def ttm(dur: dict, annual: dict, latest_end: str) -> float | None:
    """TTM = 직전 연간 + 올해 누적 − 작년 같은 기간 누적."""
    for s, ends in dur.items():
        if latest_end in ends:
            days = (d(latest_end) - d(s)).days
            if 350 <= days <= 380:
                return ends[latest_end]
            if days > 380:
                continue
            ytd = ends[latest_end]
            prev_end = None
            for fy, a in annual.items():
                if abs((d(a["end"]) - d(s)).days) <= 7:
                    prev_end = a
            if not prev_end:
                return None
            # 작년 같은 누적 기간
            for s2, ends2 in dur.items():
                for e2, v2 in ends2.items():
                    if abs((d(e2) - (d(latest_end) - dt.timedelta(days=364))).days) <= 7 and \
                            abs(((d(e2) - d(s2)).days) - days) <= 7:
                        return prev_end["value"] + ytd - v2
    return None


def shares_outstanding(facts: dict):
    """발행주식수: 표지의 dei 값(클래스별 합산) → 없으면 재무상태표의 CommonStockSharesOutstanding."""
    node = facts.get("dei", {}).get("EntityCommonStockSharesOutstanding")
    rows = node["units"].get("shares", []) if node else []
    if rows:
        last_filed = max(r["filed"] for r in rows)
        latest = [r for r in rows if r["filed"] == last_filed]
        return sum(r["val"] for r in latest), latest[0]["end"]
    node = facts.get("us-gaap", {}).get("CommonStockSharesOutstanding")
    rows = node["units"].get("shares", []) if node else []
    if rows:
        r = max(rows, key=lambda r: (r["end"], r["filed"]))
        return r["val"], r["end"]
    return None, None


def recent_filings(sub: dict, cik: str, n: int = 12):
    rec = sub["filings"]["recent"]
    out = []
    for i, form in enumerate(rec["form"]):
        if form not in ("10-K", "10-Q", "8-K", "20-F", "6-K", "DEF 14A", "13F-HR"):
            continue
        acc = rec["accessionNumber"][i].replace("-", "")
        url = f"https://www.sec.gov/Archives/edgar/data/{int(cik)}/{acc}/{rec['primaryDocument'][i]}"
        out.append({"form": form, "filed": rec["filingDate"][i], "period": rec["reportDate"][i],
                    "items": rec.get("items", [""] * len(rec["form"]))[i], "url": url})
        if len(out) >= n:
            break
    return out


# ---------------------------------------------------------------- 주가
def yahoo_chart(ticker: str, rng: str, interval: str):
    for host in ("query1", "query2"):
        try:
            j = jget(f"https://{host}.finance.yahoo.com/v8/finance/chart/{ticker}?range={rng}&interval={interval}",
                     YAHOO_UA)
            res = j["chart"]["result"][0]
            ts = res.get("timestamp") or []
            closes = res["indicators"]["quote"][0]["close"]
            pts = [(dt.datetime.utcfromtimestamp(t).date().isoformat(), round(c, 2))
                   for t, c in zip(ts, closes) if c is not None]
            return res["meta"], pts
        except Exception as e:  # 다른 호스트로 재시도
            err = e
    print(f"경고: Yahoo 주가를 받지 못했습니다 ({err})", file=sys.stderr)
    return None, []


# ---------------------------------------------------------------- 서식
def eok(v: float | None, signed: bool = False) -> str:
    """달러 → '$1,099억' 표기 (1억 달러 단위)."""
    if v is None:
        return "-"
    x = v / 1e8
    s = f"{abs(x):,.0f}" if abs(x) >= 10 else f"{abs(x):,.1f}"
    sign = "-" if v < 0 else ("+" if signed else "")
    return f"{sign}${s}억"


def pct(a: float | None, b: float | None) -> str:
    if a is None or b in (None, 0):
        return "-"
    return f"{(a / b - 1) * 100:+.0f}%"


def build(ticker: str, out: Path):
    out.mkdir(parents=True, exist_ok=True)
    cik, title = lookup_cik(ticker)
    print(f"[SEC] {ticker} → CIK {cik} ({title})")
    sub = jget(f"https://data.sec.gov/submissions/CIK{cik}.json")
    facts_all = jget(f"https://data.sec.gov/api/xbrl/companyfacts/CIK{cik}.json")
    (out / "companyfacts.json").write_text(json.dumps(facts_all), encoding="utf-8")
    facts = facts_all["facts"]
    fym = fy_end_month(sub.get("fiscalYearEnd") or "1231")

    data = {"ticker": ticker.upper(), "cik": cik, "name": sub.get("name") or title,
            "sic": sub.get("sicDescription"), "exchanges": sub.get("exchanges"),
            "fiscal_year_end": sub.get("fiscalYearEnd"), "generated": dt.date.today().isoformat(),
            "filings": recent_filings(sub, cik), "annual": {}, "quarterly": {}, "tags": {}}

    durs = {}
    for group, unit, additive in ((FLOW, "USD", True), (PER_SHARE, "USD/shares", False),
                                  (SHARES, "shares", False)):
        for key, tags in group.items():
            picked = pick_tag(facts, tags, unit)
            if not picked:
                continue
            tag, rows = picked
            a, q, dur = flow_series(rows, fym, additive)
            data["tags"][key] = tag
            data["annual"][key] = a
            data["quarterly"][key] = q
            durs[key] = (dur, a)
    for key, tags in INSTANT.items():
        picked = pick_tag(facts, tags, "USD")
        if not picked:
            continue
        tag, rows = picked
        data["tags"][key] = tag
        data["quarterly"][key] = instant_series(rows, fym)
    # 총부채가 없으면 (부채와자본총계 − 자본)으로 계산
    if "liabilities" not in data["quarterly"] and "liab_and_equity" in data["quarterly"]:
        le, eq = data["quarterly"]["liab_and_equity"], data["quarterly"].get("equity", {})
        data["quarterly"]["liabilities"] = {k: {**v, "value": v["value"] - eq[k]["value"]}
                                            for k, v in le.items() if k in eq}
    # 매출총이익 태그가 없으면 (매출 − 매출원가)
    for scope in ("annual", "quarterly"):
        rev, cor = data[scope].get("revenue", {}), data[scope].get("cost_of_revenue", {})
        gp = data[scope].setdefault("gross_profit", {})
        for k, v in rev.items():
            if k not in gp and k in cor:
                gp[k] = {**v, "value": v["value"] - cor[k]["value"]}
    # 4분기 EPS는 10-K에 따로 없으므로 (연간 EPS − 1~3분기 EPS)로 근사하고 approx 표시
    eq = data["quarterly"].get("eps_diluted", {})
    for fy, a in data["annual"].get("eps_diluted", {}).items():
        k4 = f"FY{fy}Q4"
        parts = [eq.get(f"FY{fy}Q{i}") for i in (1, 2, 3)]
        if k4 not in eq and all(parts):
            eq[k4] = {"end": a["end"], "value": round(a["value"] - sum(p["value"] for p in parts), 2),
                      "fy": fy, "q": 4, "approx": True}
    # FCF
    for scope in ("annual", "quarterly"):
        cfo, cap = data[scope].get("cfo", {}), data[scope].get("capex", {})
        data[scope]["fcf"] = {k: {**v, "value": v["value"] - cap[k]["value"]} for k, v in cfo.items() if k in cap}

    so, so_date = shares_outstanding(facts)
    data["shares_outstanding"] = so
    data["shares_outstanding_date"] = so_date

    # 주가
    meta, daily = yahoo_chart(ticker, "1y", "1d")
    _, weekly = yahoo_chart(ticker, "5y", "1wk")
    price = {}
    if meta:
        price = {"price": meta.get("regularMarketPrice"), "currency": meta.get("currency"),
                 "as_of": daily[-1][0] if daily else None, "high_52w": meta.get("fiftyTwoWeekHigh"),
                 "low_52w": meta.get("fiftyTwoWeekLow"), "long_name": meta.get("longName"),
                 "exchange": meta.get("fullExchangeName")}
        if daily:
            price["change_1y"] = daily[-1][1] - daily[0][1]
            price["change_1y_pct"] = (daily[-1][1] / daily[0][1] - 1) * 100
    data["price"] = price
    data["price_daily_1y"] = daily
    data["price_weekly_5y"] = weekly

    # 밸류에이션
    q_rev = data["quarterly"].get("revenue", {})
    latest_key = max(q_rev, key=lambda k: q_rev[k]["end"]) if q_rev else None
    latest_end = q_rev[latest_key]["end"] if latest_key else None
    val = {}
    if latest_end:
        for k in ("revenue", "net_income", "eps_diluted", "cfo", "capex", "dps"):
            if k in durs:
                val[f"ttm_{k}"] = ttm(durs[k][0], durs[k][1], latest_end)
    p = price.get("price")
    if p and so:
        val["market_cap"] = p * so
    if p and val.get("ttm_eps_diluted"):
        val["trailing_pe"] = p / val["ttm_eps_diluted"]
    if p and val.get("ttm_dps"):
        val["dividend_yield_pct"] = val["ttm_dps"] / p * 100
    if val.get("market_cap") and val.get("ttm_revenue"):
        val["ps"] = val["market_cap"] / val["ttm_revenue"]
    data["valuation"] = val
    data["latest_quarter"] = latest_key

    (out / "data.json").write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")
    (out / "summary.md").write_text(summary(data), encoding="utf-8")
    print(f"[저장] {out / 'data.json'}\n[저장] {out / 'summary.md'}")


def summary(data: dict) -> str:
    A, Q = data["annual"], data["quarterly"]
    L = [f"# {data['name']} ({data['ticker']}) 기초 데이터", "",
         f"- 생성일: {data['generated']} · SEC CIK {data['cik']} · 업종: {data.get('sic')}",
         f"- 회계연도 말: {data.get('fiscal_year_end')} (MMDD)", ""]
    pr = data.get("price") or {}
    v = data.get("valuation") or {}
    if pr:
        L += ["## 주가·밸류에이션", "",
              f"- 주가: ${pr.get('price'):,.2f} ({pr.get('as_of')} 종가 기준, {pr.get('exchange')})",
              f"- 52주 범위: ${pr.get('low_52w'):,.2f} ~ ${pr.get('high_52w'):,.2f}"]
        if pr.get("change_1y_pct") is not None:
            L.append(f"- 1년 수익률: {pr['change_1y_pct']:+.1f}%")
        if v.get("market_cap"):
            L.append(f"- 시가총액: {eok(v['market_cap'])} (≈ ${v['market_cap'] / 1e12:,.2f}조) "
                     f"· 발행주식수 {data['shares_outstanding'] / 1e8:,.1f}억주 ({data['shares_outstanding_date']} 기준)")
        if v.get("trailing_pe"):
            L.append(f"- Trailing P/E: {v['trailing_pe']:.1f}배 (TTM 희석 EPS ${v['ttm_eps_diluted']:.2f})")
        if v.get("ps"):
            L.append(f"- P/S: {v['ps']:.1f}배 (TTM 매출 {eok(v['ttm_revenue'])})")
        if v.get("dividend_yield_pct"):
            L.append(f"- 배당수익률: {v['dividend_yield_pct']:.2f}% (TTM 주당배당 ${v['ttm_dps']:.2f})")
        L += ["- Forward P/E, PEG, 목표주가·투자의견은 웹 검색으로 보완", ""]

    years = sorted({y for k in ("revenue", "net_income") for y in A.get(k, {})})[-5:]
    if years:
        L += ["## 연간 실적 (최근 5개 회계연도)", "", "| 항목 | " + " | ".join(f"FY{y}" for y in years) + " |",
              "|---|" + "---|" * len(years)]
        for k in ("revenue", "gross_profit", "operating_income", "net_income", "eps_diluted",
                  "diluted_shares", "rnd", "cfo", "capex", "fcf", "buyback", "dps"):
            if k not in A:
                continue
            cells = []
            for y in years:
                x = A[k].get(y, {}).get("value") if isinstance(A[k].get(y), dict) else None
                if x is None:
                    cells.append("-")
                elif k in ("eps_diluted", "dps"):
                    cells.append(f"${x:.2f}")
                elif k == "diluted_shares":
                    cells.append(f"{x / 1e8:,.1f}억주")
                else:
                    cells.append(eok(x))
            L.append(f"| {LABELS[k]} | " + " | ".join(cells) + " |")
        if "revenue" in A and "operating_income" in A:
            L.append("| 영업이익률 | " + " | ".join(
                f"{A['operating_income'][y]['value'] / A['revenue'][y]['value'] * 100:.1f}%"
                if y in A["operating_income"] and y in A["revenue"] else "-" for y in years) + " |")
        if "revenue" in A and "capex" in A:
            L.append("| 매출 대비 CapEx | " + " | ".join(
                f"{A['capex'][y]['value'] / A['revenue'][y]['value'] * 100:.1f}%"
                if y in A["capex"] and y in A["revenue"] else "-" for y in years) + " |")
        L.append("")

    lk = data.get("latest_quarter")
    if lk:
        fy, q = int(lk[2:6]), int(lk[-1])
        prev_y = f"FY{fy - 1}Q{q}"
        prev_q = f"FY{fy}Q{q - 1}" if q > 1 else f"FY{fy - 1}Q4"

        def qv(k, key):
            return Q.get(k, {}).get(key, {}).get("value")

        L += [f"## 최근 분기 {lk} (기간 말 {Q['revenue'][lk]['end']})", "", "### 손익계산서 (YoY)", ""]
        for k in ("revenue", "cost_of_revenue", "gross_profit", "operating_income", "rnd", "net_income"):
            if qv(k, lk) is not None:
                L.append(f"- {LABELS[k]}: {eok(qv(k, lk))} (YoY {pct(qv(k, lk), qv(k, prev_y))})")
        if qv("operating_income", lk) and qv("revenue", lk):
            om = qv("operating_income", lk) / qv("revenue", lk) * 100
            om0 = (qv("operating_income", prev_y) / qv("revenue", prev_y) * 100) if qv("revenue", prev_y) else None
            L.append(f"- 영업이익률: {om:.1f}%" + (f" (YoY {om - om0:+.1f}%p)" if om0 else ""))
        if qv("eps_diluted", lk) is not None:
            L.append(f"- 희석 EPS: ${qv('eps_diluted', lk):.2f} (YoY {pct(qv('eps_diluted', lk), qv('eps_diluted', prev_y))})")
        L += ["", "### 재무상태표 (QoQ)", ""]
        for k in ("assets", "current_assets", "cash", "ppe", "liabilities", "current_liabilities",
                  "long_term_debt", "equity"):
            if qv(k, lk) is not None:
                L.append(f"- {LABELS[k]}: {eok(qv(k, lk))} (QoQ {pct(qv(k, lk), qv(k, prev_q))})")
        if qv("current_assets", lk) and qv("current_liabilities", lk):
            L.append(f"- 유동비율 = {qv('current_assets', lk) / qv('current_liabilities', lk) * 100:.0f}%")
        if qv("liabilities", lk) and qv("equity", lk):
            L.append(f"- 부채비율 = {qv('liabilities', lk) / qv('equity', lk) * 100:.0f}%")
        L += ["", "### 현금흐름표 (분기, YoY)", ""]
        for k in ("cfo", "cfi", "capex", "acquisitions", "cff", "debt_issued", "buyback", "dividends_paid", "fcf"):
            if qv(k, lk) is not None:
                L.append(f"- {LABELS[k]}: {eok(qv(k, lk), signed=k in ('cfo', 'cfi', 'cff', 'fcf'))} "
                         f"(YoY {pct(qv(k, lk), qv(k, prev_y))})")
        L.append("")
        qs = sorted(Q.get("revenue", {}), key=lambda k: Q["revenue"][k]["end"])[-8:]
        L += ["### 최근 8개 분기 추이", "", "| 항목 | " + " | ".join(qs) + " |", "|---|" + "---|" * len(qs)]
        for k in ("revenue", "operating_income", "net_income", "eps_diluted", "capex", "fcf"):
            if k in Q:
                L.append(f"| {LABELS[k]} | " + " | ".join(
                    (("≈" if Q[k][x].get("approx") else "") + f"${Q[k][x]['value']:.2f}"
                     if k == "eps_diluted" else eok(Q[k][x]["value"]))
                    if x in Q[k] else "-" for x in qs) + " |")
        L.append("")

    L += ["## 최근 공시", ""]
    for f in data["filings"]:
        L.append(f"- {f['filed']} {f['form']} (기간 {f['period']}) {f['items'] or ''} — {f['url']}")
    L += ["", "## 사용한 XBRL 태그", ""] + [f"- {k}: {t}" for k, t in data["tags"].items()]
    return "\n".join(L) + "\n"


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("ticker")
    ap.add_argument("--out", type=Path)
    a = ap.parse_args()
    here = Path(__file__).resolve().parent
    build(a.ticker, a.out or here / "companies" / a.ticker.upper() / "data")


if __name__ == "__main__":
    main()
