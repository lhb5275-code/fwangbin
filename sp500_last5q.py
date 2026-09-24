#!/usr/bin/env python3
"""S&P 500 / Russell 1000 구성 종목의 최근 5개 분기 매출/영업이익을 수집해 Excel로 정리한다.

대상 지수 (--universe, 기본 sp500):
    sp500       : Wikipedia "List of S&P 500 companies" (접속 불가 시 GitHub datasets/s-and-p-500-companies)
    russell1000 : iShares Russell 1000 ETF(IWB) 보유종목 CSV (지수 추종 ETF의 주식 보유분)

데이터 소스 (--source, 기본 auto):
    yahoo : Yahoo Finance via yfinance (quarterly_income_stmt의 Total Revenue / Operating Income)
    sec   : SEC EDGAR XBRL companyfacts API (미국 정부 공개 데이터, API 키 불필요)
            매출 = Revenues 계열 us-gaap 태그, 영업이익 = OperatingIncomeLoss
    auto  : Yahoo Finance 접속이 되면 yahoo, 안 되면 sec 로 자동 전환
종목 리스트: Wikipedia "List of S&P 500 companies" → 접속 불가 시 GitHub datasets/s-and-p-500-companies

사용법:
    python sp500_last5q.py                 # 전체 S&P 500 수집 (캐시가 있으면 이어서)
    python sp500_last5q.py --tickers AAPL MSFT NVDA JPM   # 일부 종목만 (테스트용)
    python sp500_last5q.py --universe russell1000          # Russell 1000
    python sp500_last5q.py --as-of 2025-09-24              # 과거 시점 기준 (SEC, 그날까지 제출된 값만 사용)
    python sp500_last5q.py --source sec    # SEC EDGAR 강제 사용
    python sp500_last5q.py --fresh         # 캐시 무시하고 처음부터

SEC는 연락처가 포함된 User-Agent를 요구한다:
    SEC_USER_AGENT="홍길동 you@example.com" python sp500_last5q.py --source sec

중간 결과는 {universe}_cache_{source}_YYYYMMDD.csv 에 종목별로 한 줄씩 추가된다.
중단 후 다시 실행하면 이미 성공(ok/partial)한 종목은 건너뛰고, 실패한 종목만 다시 시도한다.
"""

import argparse
import csv
import os
import random
import sys
import time
import unicodedata
from datetime import datetime, timezone, timedelta
from io import StringIO

import numpy as np
import pandas as pd
import requests
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

WIKI_URL = "https://en.wikipedia.org/wiki/List_of_S%26P_500_companies"
GITHUB_SP500_CSV = "https://raw.githubusercontent.com/datasets/s-and-p-500-companies/main/data/constituents.csv"
YAHOO_PROBE_URL = "https://query2.finance.yahoo.com/v8/finance/chart/AAPL?range=1d&interval=1d"
SEC_FACTS_URL = "https://data.sec.gov/api/xbrl/companyfacts/CIK{cik:010d}.json"
SEC_TICKERS_URL = "https://www.sec.gov/files/company_tickers.json"
SEC_DEFAULT_UA = "sp500-last5q research script (set SEC_USER_AGENT to 'Name email')"
# 매출 태그 우선순위 (같은 최신 분기를 가진 태그가 여럿이면 앞쪽 태그 사용)
SEC_REV_TAGS = [
    "Revenues",
    "RevenueFromContractWithCustomerExcludingAssessedTax",
    "RevenueFromContractWithCustomerIncludingAssessedTax",
    "SalesRevenueNet",
    "RevenuesNetOfInterestExpense",
    "RegulatedAndUnregulatedOperatingRevenue",  # 전력·가스 유틸리티
]
# 은행: 총수익 태그가 없으면 순이자이익 + 비이자이익 (은행 영업수익의 표준 정의)
SEC_BANK_REV_PARTS = ("InterestIncomeExpenseNet", "NoninterestIncome")
# 지주회사 재편 등으로 CIK가 바뀐 경우, 새 CIK에 과거 자료가 없으면 이전 CIK로 조회
SEC_PREDECESSOR_CIK = {"XOM": 34088}
SEC_STALE_DAYS = 200  # 회사의 최신 보고 분기보다 이만큼 이상 오래된 태그는 사용 중단된 것으로 보고 제외
SEC_OI_TAG = "OperatingIncomeLoss"
SEC_FORMS = {"10-Q", "10-K", "10-Q/A", "10-K/A", "10-QT", "10-KT"}
BROWSER_UA = {"User-Agent": "Mozilla/5.0 (sp500-last5q script)"}
SOURCE_DESC = {
    "yahoo": "Yahoo Finance via yfinance (yf.Ticker(t).quarterly_income_stmt: Total Revenue / Operating Income)",
    "sec": "SEC EDGAR XBRL companyfacts API (data.sec.gov) - 매출: us-gaap Revenues 계열 태그, "
    "영업이익: us-gaap OperatingIncomeLoss, 10-Q/10-K 제출값",
}
N_Q = 5
MAX_RETRIES = 3
Q_LABELS = ["Q0", "Q-1", "Q-2", "Q-3", "Q-4"]
VERIFY_TICKERS = ["AAPL", "MSFT", "NVDA", "JPM"]
KST = timezone(timedelta(hours=9))
QOQ_HOT = 0.10  # 최근 QoQ 매출 성장률 기준
BUSINESS_KO_CSV = os.path.join(os.path.dirname(os.path.abspath(__file__)), "business_ko.csv")
IWB_HOLDINGS_CSV = "https://www.ishares.com/us/products/239707/ishares-russell-1000-etf/latest-holdings.csv"
UNIVERSE_LABEL = {"sp500": "S&P 500", "russell1000": "Russell 1000"}
TIER_LABELS = {0: "QoQ 가속 + 최근 QoQ≥10%", 1: "QoQ 가속", 2: "최근 QoQ≥10%"}
TIER_FILLS = {0: "C6EFCE", 1: "FFF2CC", 2: "FFF2CC"}

CACHE_FIELDS = (
    ["ticker", "name", "status", "error", "note"]
    + [f"date_{i}" for i in range(N_Q)]
    + [f"rev_{i}" for i in range(N_Q)]
    + [f"oi_{i}" for i in range(N_Q)]
)


# --------------------------------------------------------------------------- 1. 종목 리스트
def _normalize_universe(df):
    out = pd.DataFrame(
        {
            "name": df["Security"].astype(str).str.strip(),
            "ticker": df["Symbol"].astype(str).str.strip().str.replace(".", "-", regex=False),
            "cik": pd.to_numeric(df["CIK"], errors="coerce") if "CIK" in df else np.nan,
        }
    )
    return out.drop_duplicates("ticker").reset_index(drop=True)


def get_sp500_list():
    """(기업명, Yahoo 티커, CIK) 리스트와 출처 설명을 반환한다. Wikipedia → GitHub 순으로 시도."""
    try:
        resp = requests.get(WIKI_URL, headers=BROWSER_UA, timeout=30)
        resp.raise_for_status()
        df = pd.read_html(StringIO(resp.text), attrs={"id": "constituents"})[0]
        return _normalize_universe(df), WIKI_URL
    except Exception as e:  # noqa: BLE001
        print(f"  Wikipedia 접속 실패 ({type(e).__name__}) → GitHub 공개 데이터셋 사용")
    resp = requests.get(GITHUB_SP500_CSV, timeout=30)
    resp.raise_for_status()
    df = pd.read_csv(StringIO(resp.text))
    return _normalize_universe(df), f"{GITHUB_SP500_CSV} (Wikipedia 표 미러, Wikipedia 접속 불가로 대체)"


def get_russell1000_list():
    """iShares Russell 1000 ETF(IWB) 보유 주식 목록 → (기업명, Yahoo 티커, CIK) 리스트와 출처 설명."""
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko)"}
    resp = requests.get(IWB_HOLDINGS_CSV, headers=headers, timeout=60)
    resp.raise_for_status()
    lines = resp.text.splitlines()
    as_of = next((ln.split(",", 1)[1].strip('"') for ln in lines[:10] if ln.startswith("Fund Holdings as of")), "?")
    start = next(i for i, ln in enumerate(lines) if ln.startswith("Ticker,"))
    df = pd.read_csv(StringIO("\n".join(lines[start:])))
    df = df[df["Asset Class"] == "Equity"]
    tickers = df["Ticker"].astype(str).str.strip().str.replace(r"[ .]", "-", regex=True)
    names = df["Name"].astype(str).str.replace(r"\s+CLASS [A-Z]$", "", regex=True).str.strip()
    out = pd.DataFrame({"name": names.values, "ticker": tickers.values, "cik": np.nan})
    # S&P 500 편입 종목은 Wikipedia 표기의 기업명을 사용 (iShares 명칭은 대문자 약칭)
    try:
        sp, _ = get_sp500_list()
        sp_names = dict(zip(sp["ticker"], sp["name"]))
        sp_cik = dict(zip(sp["ticker"], sp["cik"]))
        out["name"] = [sp_names.get(t, n) for t, n in zip(out["ticker"], out["name"])]
        out["cik"] = [sp_cik.get(t, np.nan) for t in out["ticker"]]
    except Exception:  # noqa: BLE001
        pass
    out = out.drop_duplicates("ticker").reset_index(drop=True)
    return out, f"{IWB_HOLDINGS_CSV} (iShares Russell 1000 ETF 보유종목, {as_of} 기준)"


def dedupe_share_classes(universe, cache):
    """같은 회사의 복수 주식 클래스(GOOGL/GOOG 등)는 재무 데이터가 동일하므로 하나만 남긴다.

    유니버스 순서(IWB는 비중 순)상 먼저 나온 티커를 남기고, 제외된 티커는 [(제외, 남긴 티커)]로 반환.
    """
    seen, keep, dropped = {}, [], []
    for i, t in enumerate(universe["ticker"]):
        r = cache.get(t) or {}
        rev = tuple(_num(r.get(f"rev_{k}")) for k in range(N_Q))
        oi = tuple(_num(r.get(f"oi_{k}")) for k in range(N_Q))
        if r.get("status") in ("ok", "partial") and not all(np.isnan(v) for v in rev):
            key = (tuple(str(r.get(f"date_{k}")) for k in range(N_Q)), tuple(np.nan_to_num(rev, nan=-1.0)), tuple(np.nan_to_num(oi, nan=-1.0)))
            if key in seen:
                dropped.append((t, seen[key]))
                continue
            seen[key] = t
        keep.append(i)
    return universe.iloc[keep].reset_index(drop=True), dropped


def get_sec_cik_map(user_agent):
    resp = requests.get(SEC_TICKERS_URL, headers={"User-Agent": user_agent}, timeout=30)
    resp.raise_for_status()
    return {v["ticker"].upper().replace(".", "-"): (int(v["cik_str"]), v["title"]) for v in resp.json().values()}


def yahoo_reachable():
    try:
        return requests.get(YAHOO_PROBE_URL, headers=BROWSER_UA, timeout=10).status_code == 200
    except Exception:  # noqa: BLE001
        return False


# --------------------------------------------------------------------------- 2. 데이터 수집
def _row(stmt, label):
    return stmt.loc[label] if label in stmt.index else None


class NoDataError(Exception):
    """재시도해도 의미 없는 '데이터 없음' 오류."""


def fetch_yahoo(ticker, cik=None, **_):
    """yfinance에서 최근 5개 분기 매출/영업이익을 가져온다.

    반환: dict(dates, rev, oi, note) - 각 길이 N_Q, 없는 값은 NaN/None.
    데이터가 비어 있으면 예외를 던져 재시도 대상으로 만든다.
    """
    import yfinance as yf

    stmt = yf.Ticker(ticker).quarterly_income_stmt
    if stmt is None or stmt.empty:
        raise RuntimeError("quarterly_income_stmt 가 비어 있음")

    rev_row = _row(stmt, "Total Revenue")
    oi_row = _row(stmt, "Operating Income")
    # Yahoo가 아직 채우지 않은 분기 열(매출·영업이익 모두 비어 있음)은 분기로 세지 않는다
    all_cols = sorted(stmt.columns, key=pd.Timestamp, reverse=True)
    skipped = [
        c
        for c in all_cols
        if (rev_row is None or pd.isna(rev_row[c])) and (oi_row is None or pd.isna(oi_row[c]))
    ]
    cols = [c for c in all_cols if c not in skipped][:N_Q]
    if not cols:
        raise RuntimeError("Total Revenue / Operating Income 값이 없음")
    newer_skipped = [c for c in skipped if pd.Timestamp(c) > pd.Timestamp(cols[0])]
    note = (
        "Yahoo에 값 없는 최신 분기 열 제외: " + ", ".join(pd.Timestamp(c).strftime("%Y-%m-%d") for c in newer_skipped)
        if newer_skipped
        else ""
    )

    dates, rev, oi = [], [], []
    for i in range(N_Q):
        if i < len(cols):
            c = cols[i]
            dates.append(pd.Timestamp(c).strftime("%Y-%m-%d"))
            rev.append(float(rev_row[c]) if rev_row is not None and pd.notna(rev_row[c]) else np.nan)
            oi.append(float(oi_row[c]) if oi_row is not None and pd.notna(oi_row[c]) else np.nan)
        else:
            dates.append(None)
            rev.append(np.nan)
            oi.append(np.nan)
    return {"dates": dates, "rev": rev, "oi": oi, "note": note}


def _sec_periods(tag_facts, cutoff=None):
    """{(start, end): value} - 10-Q/10-K의 USD 기간값. 같은 기간이 여러 번 나오면 가장 최근 제출값."""
    best = {}
    for f in tag_facts.get("units", {}).get("USD", []):
        if f.get("form") not in SEC_FORMS or "start" not in f:
            continue
        if cutoff and f.get("filed", "9999-99-99") > cutoff:  # 기준일 이후 제출된 값은 당시 알 수 없었음
            continue
        key = (pd.Timestamp(f["start"]), pd.Timestamp(f["end"]))
        filed = f.get("filed", "")
        if key not in best or filed > best[key][1]:
            best[key] = (float(f["val"]), filed)
    return {k: v[0] for k, v in best.items()}


def sec_quarterly(tag_facts, cutoff=None):
    """{분기 종료일: (3개월 값, 산출여부)}.

    10-Q에 3개월 값이 직접 있으면 그 값을 쓴다. 없으면(대표적으로 4분기는 10-K에 연간값만 있음)
    같은 회계연도 시작일의 누적(YTD) 값 차이로 구한다: 예) Q4 = 12개월 - 9개월.
    """
    periods = _sec_periods(tag_facts, cutoff)
    q = {}
    for (st, en), v in sorted(periods.items()):
        if 80 <= (en - st).days <= 100:
            q.setdefault(en, (v, False))
    by_start = {}
    for (st, en), v in periods.items():
        by_start.setdefault(st, []).append((en, v))
    for st, lst in by_start.items():
        lst.sort()
        for (e1, v1), (e2, v2) in zip(lst, lst[1:]):
            if e2 not in q and 80 <= (e2 - e1).days <= 100 and (e2 - st).days <= 380:
                q[e2] = (v2 - v1, True)
    # 직전 누적값이 없으면: 누적값 - 같은 기간 안의 나머지 분기 합 (예: 연간 - (Q1+Q2+Q3))
    for (st, en), v in sorted(periods.items(), key=lambda kv: kv[0][1]):
        span = (en - st).days
        if en in q or not 170 <= span <= 380:
            continue
        need = round(span / 91) - 1
        inner = [val for e, (val, _) in q.items() if st < e < en and (e - st).days >= 80]
        if len(inner) == need:
            q[en] = (v - sum(inner), True)
    return q


def parse_sec_facts(facts, cutoff=None):
    """companyfacts JSON → dict(dates, rev, oi, note)."""
    gaap = facts.get("facts", {}).get("us-gaap", {})
    ends = [
        pd.Timestamp(f["end"])
        for v in gaap.values()
        for f in v.get("units", {}).get("USD", [])
        if f.get("form") in SEC_FORMS and "start" in f and (not cutoff or f.get("filed", "9999") <= cutoff)
    ]
    if not ends:
        raise NoDataError("기준일 이전 SEC 제출 자료 없음 (이후 상장·분사 또는 CIK 변경)" if cutoff else "SEC 재무 자료 없음")
    latest = max(ends)

    def fresh(series):
        return series if series and (latest - max(series)).days <= SEC_STALE_DAYS else {}

    rev_tag, rev_q = None, {}
    best_key = None
    for prio, tag in enumerate(SEC_REV_TAGS):
        if tag not in gaap:
            continue
        series = fresh(sec_quarterly(gaap[tag], cutoff))
        if not series:
            continue
        key = (max(series), -prio)
        if best_key is None or key > best_key:
            best_key, rev_tag, rev_q = key, tag, series
    bank_rev = False
    if not rev_q and all(t in gaap for t in SEC_BANK_REV_PARTS):
        nii, nonii = (fresh(sec_quarterly(gaap[t], cutoff)) for t in SEC_BANK_REV_PARTS)
        rev_q = {d: (nii[d][0] + nonii[d][0], nii[d][1] or nonii[d][1]) for d in set(nii) & set(nonii)}
        if rev_q:
            rev_tag, bank_rev = "순이자이익+비이자이익(은행)", True
    oi_q = fresh(sec_quarterly(gaap[SEC_OI_TAG], cutoff)) if SEC_OI_TAG in gaap else {}
    if not rev_q and not oi_q:
        raise NoDataError("SEC XBRL에 매출/영업이익 태그 없음 (회사 고유 태그 사용 등)")

    grid = sorted(set(rev_q) | set(oi_q), reverse=True)[:N_Q]
    dates, rev, oi, derived = [], [], [], []
    for k in range(N_Q):
        if k >= len(grid):
            dates.append(None)
            rev.append(np.nan)
            oi.append(np.nan)
            continue
        d = grid[k]
        dates.append(d.strftime("%Y-%m-%d"))
        for series, out, label in ((rev_q, rev, "매출"), (oi_q, oi, "영업이익")):
            v, is_derived = series.get(d, (np.nan, False))
            out.append(v)
            if is_derived:
                derived.append(f"{label} {Q_LABELS[k]}")
    note = f"매출태그={rev_tag or '없음(회사 고유 태그 등)'}"
    if bank_rev:
        note += "; 은행 매출은 두 공시 항목의 합"
    if derived:
        note += "; 누적값 차감으로 산출: " + ", ".join(derived)
    return {"dates": dates, "rev": rev, "oi": oi, "note": note}


def fetch_sec(ticker, cik=None, user_agent=SEC_DEFAULT_UA, cutoff=None):
    if cik is None or (isinstance(cik, float) and np.isnan(cik)):
        raise NoDataError("CIK 없음")
    ciks = [int(cik)] + ([SEC_PREDECESSOR_CIK[ticker]] if ticker in SEC_PREDECESSOR_CIK else [])
    err = None
    for c in ciks:
        resp = requests.get(SEC_FACTS_URL.format(cik=c), headers={"User-Agent": user_agent}, timeout=60)
        if resp.status_code == 404:
            err = NoDataError(f"SEC companyfacts 없음 (CIK {c})")
            continue
        resp.raise_for_status()
        try:
            data = parse_sec_facts(resp.json(), cutoff)
        except NoDataError as e:
            err = e
            continue
        if c != int(cik):
            data["note"] += f"; 이전 CIK {c} 자료 사용"
        return data
    raise err


FETCHERS = {"yahoo": fetch_yahoo, "sec": fetch_sec}


def polite_sleep():
    time.sleep(random.uniform(0.5, 1.0))


def fetch_with_retry(fetcher, ticker, **kw):
    last_err = None
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            return fetcher(ticker, **kw), None
        except NoDataError as e:
            polite_sleep()
            return None, f"NoData: {e}"
        except Exception as e:  # noqa: BLE001 - 네트워크/파싱 오류 모두 재시도
            last_err = f"{type(e).__name__}: {e}"
            print(f"    [{ticker}] 시도 {attempt}/{MAX_RETRIES} 실패 - {last_err}")
            if attempt < MAX_RETRIES:
                time.sleep(attempt * random.uniform(1.0, 2.0))
        finally:
            polite_sleep()
    return None, last_err


def load_cache(path):
    if not os.path.exists(path):
        return {}
    df = pd.read_csv(path, dtype={"ticker": str})
    # 같은 티커가 여러 번 기록되었으면(재시도) 마지막 기록을 사용
    df = df.drop_duplicates("ticker", keep="last")
    return {r["ticker"]: r for r in df.to_dict("records")}


def append_cache(path, rec):
    new = not os.path.exists(path)
    with open(path, "a", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=CACHE_FIELDS)
        if new:
            w.writeheader()
        w.writerow(rec)


def collect(universe, cache_path, fetcher, **fetch_kw):
    cache = load_cache(cache_path)
    total = len(universe)
    for i, (name, ticker, cik) in enumerate(zip(universe["name"], universe["ticker"], universe["cik"]), 1):
        cached = cache.get(ticker)
        if cached is not None and cached["status"] in ("ok", "partial"):
            continue
        print(f"[{i}/{total}] {ticker} ({name})")
        data, err = fetch_with_retry(fetcher, ticker, cik=cik, **fetch_kw)
        rec = {"ticker": ticker, "name": name}
        if data is None:
            rec.update(status="failed", error=err, note="")
        else:
            complete = all(d is not None for d in data["dates"]) and not any(
                np.isnan(v) for v in data["rev"] + data["oi"]
            )
            rec.update(status="ok" if complete else "partial", error="", note=data["note"])
            for k in range(N_Q):
                rec[f"date_{k}"] = data["dates"][k]
                rec[f"rev_{k}"] = data["rev"][k]
                rec[f"oi_{k}"] = data["oi"][k]
        append_cache(cache_path, rec)
        cache[ticker] = rec
    return cache


# --------------------------------------------------------------------------- 3. 계산
def _num(v):
    try:
        v = float(v)
    except (TypeError, ValueError):
        return np.nan
    return v


def ratio_growth(cur, base):
    cur, base = _num(cur), _num(base)
    if np.isnan(cur) or np.isnan(base) or base == 0:
        return np.nan
    return cur / base - 1


def op_income_yoy(q0, q4):
    """영업이익 YoY. 기준(Q-4)이 0 이하이면 N/M, 부호 전환은 텍스트로 표시."""
    q0, q4 = _num(q0), _num(q4)
    if np.isnan(q0) or np.isnan(q4):
        return np.nan
    if q4 <= 0:
        return "흑자전환" if q0 > 0 else "N/M"
    if q0 < 0:
        return "적자전환"
    return q0 / q4 - 1


def build_table(universe, cache):
    rows = []
    for name, ticker in zip(universe["name"], universe["ticker"]):
        r = cache.get(ticker) or {}
        rev = [_num(r.get(f"rev_{k}")) for k in range(N_Q)]
        oi = [_num(r.get(f"oi_{k}")) for k in range(N_Q)]
        dates = [r.get(f"date_{k}") if isinstance(r.get(f"date_{k}"), str) else None for k in range(N_Q)]
        row = {
            "name": name,
            "ticker": ticker,
            "status": r.get("status", "failed"),
            "error": r.get("error") if isinstance(r.get("error"), str) else "",
            "note": r.get("note") if isinstance(r.get("note"), str) else "",
            "dates": dates,
            "rev": rev,
            "oi": oi,
            "rev_yoy": ratio_growth(rev[0], rev[4]),
            "oi_yoy": op_income_yoy(oi[0], oi[4]),
            "qoq": [ratio_growth(rev[k], rev[k + 1]) for k in range(4)],
        }
        row["tier"] = qoq_tier(row["qoq"])
        rows.append(row)
    rows.sort(key=sort_key)
    return rows


def qoq_tier(qoq):
    """0: QoQ 가속 & 최근 QoQ≥10%, 1: QoQ 가속만, 2: 최근 QoQ≥10%만, 3: 해당 없음.

    QoQ 가속 = 매출 QoQ 성장률이 4개 구간 연속 상승: QoQ(Q0) > QoQ(Q-1) > QoQ(Q-2) > QoQ(Q-3)
    """
    accel = not any(np.isnan(v) for v in qoq) and qoq[0] > qoq[1] > qoq[2] > qoq[3]
    hot = not np.isnan(qoq[0]) and qoq[0] >= QOQ_HOT
    if accel and hot:
        return 0
    if accel:
        return 1
    if hot:
        return 2
    return 3


def sort_key(x):
    """선정 기업(0~2)은 등급 → 최근 QoQ 내림차순, 나머지는 매출 YoY 내림차순(NaN 맨 아래)."""
    if x["tier"] < 3:
        return (0, min(x["tier"], 1), -x["qoq"][0])
    yoy = x["rev_yoy"]
    return (1, int(np.isnan(yoy)), -yoy if not np.isnan(yoy) else 0)


def load_business_ko(path=BUSINESS_KO_CSV):
    if not os.path.exists(path):
        return {}
    df = pd.read_csv(path, dtype=str, encoding="utf-8-sig").dropna()
    return dict(zip(df["ticker"], df["business_ko"]))


def yahoo_industry(ticker):
    """한글 설명이 없는 종목용 대체값: Yahoo 섹터/업종."""
    try:
        import yfinance as yf

        info = yf.Ticker(ticker).info
        parts = [p for p in (info.get("sector"), info.get("industry")) if p]
        return f"(Yahoo 업종) {' / '.join(parts)}" if parts else ""
    except Exception:  # noqa: BLE001
        return ""
    finally:
        polite_sleep()


def attach_business(rows):
    ko = load_business_ko()
    for r in rows:
        if r["tier"] < 3:
            r["business"] = ko.get(r["ticker"]) or yahoo_industry(r["ticker"])
        else:
            r["business"] = ""


# --------------------------------------------------------------------------- 4. Excel
HEADERS = (
    ["기업명", "티커", "최근 분기 종료일"]
    + [f"매출 {q} (백만$)" for q in Q_LABELS]
    + [f"영업이익 {q} (백만$)" for q in Q_LABELS]
    + ["매출성장률(YoY)", "영업이익성장률(YoY)"]
    + [f"매출 QoQ({q})" for q in Q_LABELS[:4]]
    + ["선정 기준", "주요 사업"]
)
FMT_MUSD = "#,##0"
FMT_PCT = "0.0%"


def _display_width(s):
    return sum(2 if unicodedata.east_asian_width(ch) in ("W", "F") else 1 for ch in str(s))


def _cell_text(v, fmt):
    if v is None:
        return ""
    if isinstance(v, float):
        if fmt == FMT_PCT:
            return f"{v * 100:.1f}%"
        if fmt == FMT_MUSD:
            return f"{v:,.0f}"
    return str(v)


def _blank(v):
    return None if isinstance(v, float) and np.isnan(v) else v


def _musd(v):
    return None if np.isnan(v) else v / 1e6


def autofit(ws, formats, min_w=6, max_w=60):
    widths = {}
    for row in ws.iter_rows():
        for c in row:
            w = _display_width(_cell_text(c.value, formats.get(c.column)))
            widths[c.column] = max(widths.get(c.column, 0), w)
    for col, w in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = min(max(w + 2, min_w), max_w)


def write_excel(
    rows, universe_size, failed, missing, collected_at, path, source, universe_src, dup_classes=(), as_of=None
):
    wb = Workbook()
    ws = wb.active
    ws.title = "실적"
    ws.append(HEADERS)
    hdr_fill = PatternFill("solid", fgColor="DDEBF7")
    for c in ws[1]:
        c.font = Font(bold=True)
        c.fill = hdr_fill
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=False)

    for r in rows:
        ws.append(
            [r["name"], r["ticker"], r["dates"][0]]
            + [_musd(v) for v in r["rev"]]
            + [_musd(v) for v in r["oi"]]
            + [_blank(r["rev_yoy"]), _blank(r["oi_yoy"])]
            + [_blank(v) for v in r["qoq"]]
            + [TIER_LABELS.get(r["tier"], ""), r["business"]]
        )
        if r["tier"] < 3:
            fill = PatternFill("solid", fgColor=TIER_FILLS[r["tier"]])
            for c in ws[ws.max_row]:
                c.fill = fill

    formats = {}
    for col in range(4, 14):
        formats[col] = FMT_MUSD
    for col in range(14, 20):
        formats[col] = FMT_PCT
    for row in ws.iter_rows(min_row=2):
        for c in row:
            fmt = formats.get(c.column)
            if fmt and isinstance(c.value, (int, float)):
                c.number_format = fmt
            elif c.column == 15 and isinstance(c.value, str):
                c.alignment = Alignment(horizontal="right")

    ws.freeze_panes = "C2"  # 헤더 행 + 기업명/티커 열 고정
    ws.auto_filter.ref = ws.dimensions
    autofit(ws, formats, max_w=80)

    # ---- 메모 시트
    memo = wb.create_sheet("메모")
    bold = Font(bold=True)
    ok_count = sum(1 for r in rows if r["status"] in ("ok", "partial"))
    memo.append(["데이터 출처", SOURCE_DESC[source]])
    if as_of:
        memo.append(
            [
                "데이터 기준일",
                f"{as_of} 시점 (point-in-time): 이날까지 SEC에 제출된 10-Q/10-K 값만 사용. "
                "이후 정정·재작성된 수치와 이후 발표된 분기는 반영하지 않음. 종목 구성은 수집일 현재 S&P 500",
            ]
        )
    memo.append(["종목 리스트 출처", universe_src])
    memo.append(["수집 일시", collected_at])
    memo.append(["수집 성공 / 전체", f"{ok_count} / {universe_size}"])
    memo.append(["단위", "매출·영업이익: 백만 달러(USD mn), 성장률: %"])
    memo.append(
        [
            "영업이익 YoY 규칙",
            "Q-4 영업이익 ≤ 0 → Q0 > 0 이면 '흑자전환', 아니면 'N/M' / Q-4 > 0 이고 Q0 < 0 → '적자전환'",
        ]
    )
    n_tier = {t: sum(1 for r in rows if r["tier"] == t) for t in range(3)}
    memo.append(
        [
            "행 정렬 기준",
            f"① 초록: QoQ 가속(QoQ(Q0)>QoQ(Q-1)>QoQ(Q-2)>QoQ(Q-3)) 이면서 최근 QoQ≥10% ({n_tier[0]}개) → "
            f"② 노랑: 둘 중 하나만 해당 (가속 {n_tier[1]}개, QoQ≥10% {n_tier[2]}개) — ①②는 최근 QoQ 내림차순 → "
            "③ 나머지: 매출 YoY 내림차순(계산 불가는 맨 아래). "
            "주요 사업은 ①② 기업에 기재 (business_ko.csv, 없으면 Yahoo 업종)",
        ]
    )
    for r in range(1, memo.max_row + 1):
        memo.cell(r, 1).font = bold

    memo.append([])
    memo.append([f"수집 실패 티커 ({len(failed)}개)"])
    memo.cell(memo.max_row, 1).font = bold
    memo.append(["티커", "기업명", "오류"])
    for c in memo[memo.max_row]:
        c.font = bold
    for t, n, e in failed:
        memo.append([t, n, e])
    if not failed:
        memo.append(["(없음)"])

    if dup_classes:
        memo.append([])
        memo.append([f"동일 기업 중복 주식 클래스 제외 ({len(dup_classes)}개)"])
        memo.cell(memo.max_row, 1).font = bold
        memo.append(["제외 티커", "남긴 티커", "사유"])
        for c in memo[memo.max_row]:
            c.font = bold
        for d, k in dup_classes:
            memo.append([d, k, "재무 데이터 동일 (같은 회사의 다른 주식 클래스)"])

    memo.append([])
    memo.append([f"데이터 누락 티커 ({len(missing)}개)"])
    memo.cell(memo.max_row, 1).font = bold
    memo.append(["티커", "기업명", "누락 내용"])
    for c in memo[memo.max_row]:
        c.font = bold
    for t, n, e in missing:
        memo.append([t, n, e])
    if not missing:
        memo.append(["(없음)"])

    memo.append([])
    memo.append(["기업별 실제 분기 종료일"])
    memo.cell(memo.max_row, 1).font = bold
    memo.append(["티커", "기업명"] + [f"{q} 종료일" for q in Q_LABELS] + ["비고"])
    for c in memo[memo.max_row]:
        c.font = bold
    for r in sorted(rows, key=lambda x: x["ticker"]):
        memo.append([r["ticker"], r["name"]] + [d or "" for d in r["dates"]] + [r["note"]])
    if source == "sec":
        memo.insert_rows(7)
        memo["A7"] = "SEC 데이터 참고"
        memo["A7"].font = bold
        memo["B7"] = (
            "4분기 등 3개월 값이 직접 공시되지 않은 분기는 같은 회계연도 누적값 차이(예: 연간-9개월)로 구함(비고란 표시). "
            "Yahoo의 Total Revenue와 정의가 다를 수 있음(특히 금융사). 분기 종료일은 회사 회계기준일 그대로."
        )
    autofit(memo, {}, max_w=80)

    wb.save(path)


def describe_missing(r):
    if r["status"] == "failed":
        return None
    issues = []
    n_dates = sum(1 for d in r["dates"] if d)
    if n_dates < N_Q:
        issues.append(f"분기 {n_dates}개만 제공")
    miss_rev = [Q_LABELS[k] for k in range(n_dates) if np.isnan(r["rev"][k])]
    miss_oi = [Q_LABELS[k] for k in range(n_dates) if np.isnan(r["oi"][k])]
    if miss_rev:
        issues.append("매출 없음: " + ",".join(miss_rev))
    if miss_oi:
        issues.append("영업이익 없음: " + ",".join(miss_oi))
    if r["note"].startswith("Yahoo에 값 없는"):
        issues.append(r["note"])
    if r["dates"][0] and r["dates"][4]:
        gap = (pd.Timestamp(r["dates"][0]) - pd.Timestamp(r["dates"][4])).days
        if not 330 <= gap <= 400:
            issues.append(f"Q0~Q-4 간격 {gap}일(1년 아님, YoY 해석 주의)")
    return "; ".join(issues) or None


# --------------------------------------------------------------------------- 5. 검증 출력
def _fmt_m(v):
    return "" if np.isnan(v) else f"{v / 1e6:,.0f}"


def _fmt_p(v):
    if isinstance(v, str):
        return v
    return "" if np.isnan(v) else f"{v * 100:.1f}%"


def print_verification(rows, source):
    by_t = {r["ticker"]: r for r in rows}
    where = {
        "yahoo": "Yahoo Finance > Financials > Income Statement > Quarterly",
        "sec": "SEC EDGAR 10-Q/10-K 손익계산서 (Yahoo Finance와도 대조 가능)",
    }[source]
    print("\n" + "=" * 78)
    print(f"검증용 값 (단위: 백만$) - {where} 와 대조")
    print("=" * 78)
    for t in VERIFY_TICKERS:
        r = by_t.get(t)
        if r is None:
            print(f"\n{t}: 대상 목록에 없음")
            continue
        print(f"\n{t} - {r['name']}  (status: {r['status']})" + (f"  [{r['note']}]" if r["note"] else ""))
        print(f"  {'분기':<5}{'종료일':<12}{'매출':>14}{'영업이익':>14}")
        for k in range(N_Q):
            print(f"  {Q_LABELS[k]:<6}{r['dates'][k] or '-':<14}{_fmt_m(r['rev'][k]):>14}{_fmt_m(r['oi'][k]):>16}")
        print(f"  매출 YoY: {_fmt_p(r['rev_yoy'])}   영업이익 YoY: {_fmt_p(r['oi_yoy'])}")
        print("  매출 QoQ: " + ", ".join(f"{Q_LABELS[k]} {_fmt_p(r['qoq'][k])}" for k in range(4)))


# --------------------------------------------------------------------------- main
def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--tickers", nargs="+", help="전체 대신 이 티커들만 수집 (Yahoo 형식, 예: BRK-B)")
    ap.add_argument("--fresh", action="store_true", help="기존 캐시 CSV를 지우고 처음부터 수집")
    ap.add_argument("--outdir", default=".", help="출력/캐시 디렉터리 (기본: 현재 디렉터리)")
    ap.add_argument("--source", choices=["auto", "yahoo", "sec"], default="auto", help="데이터 소스 (기본 auto)")
    ap.add_argument(
        "--as-of",
        help="과거 기준일(YYYY-MM-DD). 이 날까지 SEC에 제출된 10-Q/10-K 값만으로 '당시 최근 5개 분기'를 구성 (SEC 소스 강제)",
    )
    ap.add_argument("--universe", choices=list(UNIVERSE_LABEL), default="sp500", help="대상 지수 (기본 sp500)")
    ap.add_argument(
        "--sec-user-agent",
        default=os.environ.get("SEC_USER_AGENT", SEC_DEFAULT_UA),
        help="SEC 요청용 User-Agent ('이름 이메일' 형식 권장, 환경변수 SEC_USER_AGENT)",
    )
    args = ap.parse_args()

    source = args.source
    if args.as_of:
        args.as_of = pd.Timestamp(args.as_of).strftime("%Y-%m-%d")
        if source == "yahoo":
            ap.error("--as-of 는 제출일 기록이 있는 SEC 소스만 지원합니다")
        source = "sec"
        print(f"기준일 {args.as_of}: SEC EDGAR에서 그날까지 제출된 값만 사용")
    elif source == "auto":
        print("Yahoo Finance 접속 확인 중...")
        source = "yahoo" if yahoo_reachable() else "sec"
        print(f"  → 데이터 소스: {source}" + ("" if source == "yahoo" else " (Yahoo 접속 불가, SEC EDGAR로 대체)"))
    fetch_kw = {}
    if source == "sec":
        fetch_kw["user_agent"] = args.sec_user_agent
        if args.as_of:
            fetch_kw["cutoff"] = args.as_of
        if args.sec_user_agent == SEC_DEFAULT_UA:
            print("  ! SEC는 연락처가 담긴 User-Agent를 요구합니다. 403이 나오면 SEC_USER_AGENT='이름 이메일' 을 설정하세요.")

    today = datetime.now(KST).strftime("%Y%m%d")
    os.makedirs(args.outdir, exist_ok=True)
    prefix = args.universe + (f"_asof{args.as_of.replace('-', '')}" if args.as_of else "")
    cache_path = os.path.join(args.outdir, f"{prefix}_cache_{source}_{today}.csv")
    out_path = os.path.join(args.outdir, f"{prefix}_last5q_{today}.xlsx")
    if args.fresh and os.path.exists(cache_path):
        os.remove(cache_path)

    print(f"{UNIVERSE_LABEL[args.universe]} 종목 리스트 가져오는 중...")
    universe, universe_src = get_sp500_list() if args.universe == "sp500" else get_russell1000_list()
    if source == "sec" and universe["cik"].isna().any():
        try:
            cmap = get_sec_cik_map(args.sec_user_agent)
            universe["cik"] = [c if pd.notna(c) else cmap.get(t, (np.nan,))[0] for t, c in zip(universe["ticker"], universe["cik"])]
        except Exception as e:  # noqa: BLE001
            print(f"  SEC 티커→CIK 매핑 실패: {e}")
    if args.tickers:
        wanted = [t.upper().replace(".", "-") for t in args.tickers]
        sub = universe[universe["ticker"].isin(wanted)]
        extra = [t for t in wanted if t not in set(sub["ticker"])]
        extra_df = pd.DataFrame({"name": extra, "ticker": extra, "cik": np.nan})
        if extra and source == "sec":
            try:
                cmap = get_sec_cik_map(args.sec_user_agent)
                extra_df["cik"] = [cmap.get(t, (np.nan,))[0] for t in extra]
                extra_df["name"] = [cmap.get(t, (None, t))[1] for t in extra]
            except Exception as e:  # noqa: BLE001
                print(f"  SEC 티커→CIK 매핑 실패: {e}")
        universe = pd.concat([sub, extra_df], ignore_index=True)
    print(f"대상 종목: {len(universe)}개  |  캐시: {cache_path}")

    cache = collect(universe, cache_path, FETCHERS[source], **fetch_kw)
    collected_at = datetime.now(KST).strftime("%Y-%m-%d %H:%M:%S KST")
    universe, dup_classes = dedupe_share_classes(universe, cache)
    if dup_classes:
        print(f"동일 기업 중복 주식 클래스 {len(dup_classes)}개 제외: " + ", ".join(f"{d}(={k})" for d, k in dup_classes))

    rows = build_table(universe, cache)
    attach_business(rows)
    failed = [(r["ticker"], r["name"], r["error"]) for r in rows if r["status"] == "failed"]
    missing = [(r["ticker"], r["name"], m) for r in rows if (m := describe_missing(r))]
    write_excel(
        rows, len(universe), failed, missing, collected_at, out_path, source, universe_src, dup_classes, args.as_of
    )

    print_verification(rows, source)

    ok = len(universe) - len(failed)
    print("\n" + "=" * 78)
    print(f"데이터 소스: {SOURCE_DESC[source]}")
    print(f"수집 성공: {ok} / {len(universe)} 종목  (데이터 일부 누락: {len(missing)}개)")
    print(f"최종 실패 티커 ({len(failed)}개): {', '.join(t for t, _, _ in failed) or '없음'}")
    print(f"Excel 저장: {out_path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
