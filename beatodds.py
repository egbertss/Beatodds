# BeatOdds. https://github.com/egbertss/beatodds
# Author: Jacobus Egberts Sonatori (github.com/egbertss). Idea, design, code and analysis.
# Started as a course project, Master in Finance, International University of Monaco, 2026.
# Run in Google Colab. First: !pip install -q yfinance openpyxl scikit-learn
# Keys: GEMINI_API_KEY and SEC_CONTACT below, or as Colab secrets of the same names. Never commit a filled-in key.
# STEP 1 asks for saved RawData zips, one file per dialog (Cancel when done, or at once to start from zero).
# Cell 1: paste this script with STEP = 1 and run. Downloads all raw data and gives GroupA2_LSEG_request.xlsx. Fill it in Excel with LSEG Workspace, save.
# Cell 2: STEP = 2 then run(). Upload the filled LSEG file. Builds GroupA2_P3_Workbook.xlsx.
# ask("..."): optional portfolio agent, needs a Gemini key (GEMINI_API_KEY) or Colab's built-in Gemini.
# Raw_ sheets hold downloaded data and Python outputs (Raw_Models, Raw_Bootstrap). Every other cell is a formula.
# Part 3 adds: Crowd_Adds sheet (odds drift, crowd vs StarMine and analysts, crowd-warning screen) and Business sheet.
# Generative AI: Claude helped draft and debug this script. The author chose the questions, rules and checks.

import hashlib
import io
import json
import os
import re
import sys
import time
import warnings
import zipfile
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, timedelta, timezone

import numpy as np
import pandas as pd
import requests

warnings.filterwarnings("ignore", category=FutureWarning)

# =====================================================================================================================
# GEMINI API KEY: free at aistudio.google.com > Get API key. Set it as a Colab secret named GEMINI_API_KEY, or paste it between the quotes.
GEMINI_API_KEY = "" or os.environ.get("GEMINI_API_KEY", "")
# SEC CONTACT: your name and e-mail (the SEC asks for one to read EDGAR). Colab secret SEC_CONTACT, or paste it between the quotes.
SEC_CONTACT = "" or os.environ.get("SEC_CONTACT", "")
# =====================================================================================================================

# Settings
GROUP_NUMBER = "A2"
HIGH_P = 0.80                    # confident crowd. Fixed in Part 1, do not change after seeing new data
FALL = -0.05                     # a fall: abnormal return of -5% or worse
BIG_FALL = -0.10
PUT_OTM = 0.05                   # protective put struck 5% below the last close before the report
DEFAULT_PUT_COST = 0.02          # put price used on the Hedge sheet until snapshots exist (share of the share price)
FREEZE_DATE = "2026-10-08"       # Part 1 data run. Later reports are the out-of-sample test
CUTOFF_HOUR_UTC = 10             # odds read at 10:00 UTC on the report day
LEAD_DAYS = 7
MIN_PRIOR_EVENTS = 20
MIN_HISTORY = 4
MIN_PAST_MOVES = 3
RUNUP_DAYS = 10
VOL_DAYS = 20
MARKET = "SPY"
K_MIN_PRIOR_EVENTS = 10

SNAPSHOT_ONLY = False            # True: only save today's snapshot (daily run). False: full download and workbook
TAKE_SNAPSHOT = True
USE_KALSHI = True
USE_EDGAR = True                 # release times from SEC EDGAR (8-K item 2.02), to check the odds were read before the release
SEC_USER_AGENT = SEC_CONTACT or os.environ.get("SEC_USER_AGENT", "")   # from SEC_CONTACT at the top. Empty = EDGAR is skipped
USE_LSEG = True                  # option benchmark from LSEG Workspace implied volatility (fast, needs Workspace in Excel)
STEP = 1                         # 1: download all raw data and give the LSEG Excel file. 2: upload the filled LSEG file and build the workbook "score": fill the Scorecard from two saved workbooks
LSEG_INTERACTIVE = False         # (old one-cell pause, replaced by STEP 1 and STEP 2)
LSEG_ONLY = False
LSEG_ALL_REPORTS = True          # True: the request lists every settled report. False: only reports without implied volatility yet
LSEG_FIELD = "TR.30DayAtTheMoneyImpliedVolatilityIndexForPutOptions"   # check it in Build Formula, then change it in the request file
LSEG_CODE_TYPE = "IV RIC"        # tested on Apple: the field returns history on LSEG's volatility codes (AAPLATMIV.U), not on AAPL.O
LSEG_PROBE_FIELDS = ["TR.30DayAtTheMoneyImpliedVolatilityIndexForPutOptions", "TR.30DayAtTheMoneyImpliedVolatilityIndexForCallOptions",
                     "TR.60DayAtTheMoneyImpliedVolatilityIndexForPutOptions", "TR.90DayAtTheMoneyImpliedVolatilityIndexForPutOptions",
                     "TR.ImpliedVolatility", "TR.PriceClose", "TR.CLOSEPRICE", "TR.Volatility30D"]   # tried side by side on the Setup sheet
IV_DAYS = 30                     # maturity of the implied volatility used for the event move (days)
LSEG_EPS_FIELD = "TR.EPSMean"    # I/B/E/S consensus EPS, read as it stood on a past date (analyst revisions for every past report)
LSEG_EPS_PARAMS = "Period=FQ1"   # the quarter about to be reported
REV_DAYS = 30                    # revision window: consensus the day before the report vs this many days earlier
LSEG_SMART_FIELD = "TR.EPSSmartEst"   # LSEG StarMine SmartEstimate (the professional tool that predicts beats and misses)
TRIM = 0.50                      # backtest: share of the position sold at the close before the report and bought back after
TCOST = 0.001                    # backtest: round-trip trading cost per unit traded (0.10%)
DRIFT = 0.05                     # Part 3: the crowd lost confidence when p fell by at least this over the week
CROWD_WARN_WEIGHT = 1.0          # Part 3: crowd-warning screen = past |move| x (1 + weight) when the crowd warns. Fixed before the test
Q_LOW = 0.10                     # the AI downside model predicts this quantile of the abnormal return (1-in-10 worst case)
USE_GENAI = True                 # generative AI step: Colab's built-in Gemini (google.colab.ai), or a free Gemini API key below
GENAI_PARSE_N = 25               # Polymarket contracts read by the generative AI to check our contract parser
GENAI_KALSHI_N = 25              # Kalshi KPI contracts read by the generative AI (free-text titles)
GENAI_BATCH = 25                 # earnings releases read per generative AI call (25: a run from zero fits in one free daily quota)
PEER_DAYS = 30                   # peers: same-sector companies that reported in the previous PEER_DAYS days
LSEG_CALL_FIELD = "TR.30DayAtTheMoneyImpliedVolatilityIndexForCallOptions"   # put minus call implied volatility = skew
GENAI_NOTES_N = 15               # upcoming reports that get a client note written by the generative AI
STRESS_N = 40                    # Polymarket contracts in the stress test (contract text only, no slug, plus four format changes)
GENAI_NEW_N = 400                # each run, Gemini reads up to this many contracts it has not read before (open or closed), newest first
PARSE_BATCH = 20                 # contracts per Gemini call in the parser check
STRESS_NEW_N = 15                # each run, new contracts added to the stress test (each read in five text versions)
KMAP_BATCH = 25                  # Kalshi KPI series titles mapped to a stock ticker per generative AI call
HAND_MINUTES = 2.0               # minutes an analyst needs to map one contract by hand (assumption, blue cell on AI_Checks)
GUARD_TOL = 0.05                 # guardrail: a number in an AI text matches the data within this, or as the data rounded
SIZE_CAP = 1.0                   # sized trims: the most of one position sold before a report (1.0 = all of it)
KALSHI_FEAT_DAYS = 3             # Kalshi KPI contracts closing this many days around a Polymarket report feed its model input
AGENT_MAX_STEPS = 6              # tool calls the agent may make before it must answer
AGENT_DEMO = ("Client portfolio: AAPL 8%, AMZN 6%, PLTR 5%, DASH 4%, PINS 3%, KO 5%. Value EUR 2 million. "
              "Which holdings should we act on before their earnings reports?")   # demo brief for the workbook ("" = none)
USE_MASSIVE = False              # free alternative: past option prices from Massive (formerly Polygon.io). Slow: about 140 reports an hour
MASSIVE_API_KEY = ""   # free key from massive.com, Options Basic plan
OPTION_HISTORY_MINUTES = 240     # minutes per run spent downloading past option prices (free plan: 5 requests a minute)
OPTION_HISTORY_ATM = False       # False: only the protective put (2 requests per report). True: also the ATM put for the implied move
CONF_LEVELS = (0.90, 0.95, 0.99)
MIN_WEEKS_CLUSTER = 20           # fewer report weeks than this: standard errors by report instead of by week
GAP_THRESHOLDS = (0.70, 0.75, 0.80, 0.85, 0.90)
SNAPSHOT_DAYS = 21
START_DATE = "2025-01-01"
HIST_START = "2024-01-01"        # price history start, so past earnings moves and volatility can be measured
ODDS_WINDOW_DAYS = 10
MIN_VOLUME = 1000
K_MIN_VOLUME = 200
BOOT = 2000
SEED = 7
THREADS = 8
K_THREADS = 3                    # Kalshi limits request rates. Fewer threads = complete downloads
DATA_DIR = "BeatOdds_data"
OUTPUT_FILE = f"Group{GROUP_NUMBER}_P3_Workbook.xlsx"
RAW_ZIP = f"Group{GROUP_NUMBER}_P3_RawData.zip"
GAMMA = "https://gamma-api.polymarket.com"
CLOB = "https://clob.polymarket.com"
KALSHI = "https://api.elections.kalshi.com/trade-api/v2"
K_OVERRIDES = {"KXWALMARTA": "WMT", "KXUBERTRIPSA": "UBER", "KXCOINBASE": "COIN", "KXKLARNAA": "KLAR", "KXMETA": "META",
               "KXDIS": "DIS", "KXDISA": "DIS", "KXTSM": "TSM", "KXSPCXA": None, "GOOGLESHARE": None, "OAIAGI": None}

F_MARKETS, F_ODDS, F_PRICES, F_EPS = "polymarket_markets.csv", "polymarket_odds.csv", "prices_yfinance.csv", "eps_yfinance.csv"
F_SNAP, F_KM, F_KODDS, F_LOG = "snapshots.csv", "kalshi_kpi_markets.csv", "kalshi_odds.csv", "download_log.txt"
F_EDGAR = "edgar_releases.csv"
F_OPTHIST = "options_history.csv"
F_LSEG = "lseg_iv.csv"
F_AI_PARSE, F_AI_LIVE = "genai_parse_check.csv", "ai_live.csv"
F_PRIOR, F_AI_REL, F_SECTOR = "prior_releases.csv", "genai_releases.csv", "sectors_yfinance.csv"
F_KALL, F_KMAP, F_STRESS = "kalshi_kpi_all.csv", "genai_kalshi_map.csv", "genai_stress.csv"
F_FLAGS, F_SCORE, F_GUARD, F_AGENT = "flags_log.csv", "scorecard.csv", "genai_guardrail.csv", "agent_briefs.csv"
F_HASH, F_HEDGE = "flags_hash.csv", "hedge_stats.json"
BRIEF_FILE = f"Group{GROUP_NUMBER}_P3_ClientBrief.md"
LSEG_REQUEST = f"Group{GROUP_NUMBER}_LSEG_request.xlsx"
MASSIVE = "https://api.massive.com"
OH_COLS = ["slug", "ticker", "report_date", "entry_date", "exit_date", "spot", "spot_after", "expiry", "atm_strike", "atm_ticker",
           "atm_entry", "atm_entry_date", "put_strike", "put_ticker", "put_entry", "put_entry_date", "put_exit", "put_exit_date", "status"]
SNAP_COLS = ["snap_utc", "slug", "ticker", "report_date", "p_gamma", "p_mid", "volume", "spot", "expiry", "atm_strike",
             "implied_move", "put95_strike", "put95_cost", "put90_strike", "put90_cost", "eps_now", "eps_7d_ago", "eps_30d_ago",
             "eps_90d_ago", "rev_up_7d", "rev_down_7d", "rev_up_30d", "rev_down_30d", "n_analysts", "eps_low", "eps_high", "eps_avg", "put95_bid", "put95_ask",
             "put95_last"]

_S = requests.Session()
_S.headers.update({"User-Agent": "IUM-BeatOdds-student-project/2.0"})


def _p(name):
    return os.path.join(DATA_DIR, name)


def _get(url, params=None, tries=4, pause=0.0):
    for i in range(tries):
        try:
            if pause:
                time.sleep(pause)
            r = _S.get(url, params=params, timeout=30)
            if r.status_code == 429:
                time.sleep(2 * (i + 1) ** 2); continue
            r.raise_for_status()
            return r.json()
        except Exception:
            if i == tries - 1:
                raise
            time.sleep(1 + 2 * i)


def _jl(x):
    return json.loads(x) if isinstance(x, str) else (x or [])


def _f(x):
    try:
        v = float(x)
        return v if np.isfinite(v) else None
    except Exception:
        return None


def _ts(s):
    try:
        return int(pd.Timestamp(s).timestamp())
    except Exception:
        return None


def yf_symbol(t):
    return str(t).replace(".", "-")


def cut_ts(rdate):
    d = pd.Timestamp(rdate)
    return int(datetime(d.year, d.month, d.day, CUTOFF_HOUR_UTC, tzinfo=timezone.utc).timestamp())


def _quiet_yf():
    import logging
    logging.getLogger("yfinance").setLevel(logging.CRITICAL)
    import yfinance as yf
    return yf


# 1. Polymarket earnings contracts (same parsing as Part 1)
def fetch_events(closed):
    out, off = [], 0
    while True:
        page = _get(GAMMA + "/events", {"tag_slug": "earnings", "closed": str(closed).lower(), "limit": 100, "offset": off}) or []
        out += page
        if len(page) < 100:
            break
        off += 100
        time.sleep(0.2)
    return out


def _parse_market(ev, m):
    q = (m.get("question") or ev.get("title") or "")
    if "beat" not in q.lower():
        return None
    slug = m.get("slug") or ev.get("slug") or ""
    tk = re.search(r"\(([A-Z][A-Z.\-]{0,6})\)", q) or re.search(r"\(([A-Z][A-Z.\-]{0,6})\)", ev.get("title") or "")
    ticker = tk.group(1) if tk else None
    if not ticker:
        s0 = re.match(r"^([a-z]{1,6})-quarterly-earnings", slug)
        ticker = s0.group(1).upper() if s0 else None
    d = (re.search(r"(\d{2})-(\d{2})-(\d{4})", slug), re.search(r"(\d{4})-(\d{2})-(\d{2})", slug))
    if d[0]:
        rdate = pd.Timestamp(f"{d[0].group(3)}-{d[0].group(1)}-{d[0].group(2)}")
    elif d[1]:
        rdate = pd.Timestamp(d[1].group(0))
    else:
        e = m.get("endDate") or ev.get("endDate")
        rdate = pd.Timestamp(str(e)[:10]) if e else None
    th = re.search(r"(neg-?|m)?(\d+)pt(\d+)$", slug)
    if th:
        thr = float(f"{th.group(2)}.{th.group(3)}") * (-1 if th.group(1) else 1)
    else:
        tt = re.search(r"\$\s?(-?\d+(?:\.\d+)?)", q + " " + (ev.get("description") or ""))
        thr = float(tt.group(1)) if tt else None
    outs, prices = [o.lower() for o in _jl(m.get("outcomes"))], _jl(m.get("outcomePrices"))
    toks = _jl(m.get("clobTokenIds"))
    if "yes" not in outs or not toks:
        return None
    iy = outs.index("yes")
    p_last = _f(prices[iy]) if len(prices) > iy else None
    closed = bool(m.get("closed"))
    resolved = closed and p_last in (0.0, 1.0)
    return dict(slug=slug, question=q, ticker=ticker, report_date=rdate, threshold=thr,
                eps_type="non-GAAP" if "nongaap" in slug else ("GAAP" if "gaap" in slug else "not stated"),
                token=toks[iy], volume=float(m.get("volume") or 0), created=str(m.get("startDate") or ev.get("startDate") or "")[:19],
                closed=closed, resolved=resolved, outcome=(int(p_last) if resolved else None), p_now=p_last,
                rules=" ".join(str(m.get("description") or ev.get("description") or "").split())[:900])


def fetch_markets():
    rows = []
    for closed in (True, False):
        evs = fetch_events(closed)
        print(f"   {'closed' if closed else 'open'} earnings events: {len(evs)}", flush=True)
        for ev in evs:
            for m in ev.get("markets", []) or []:
                x = _parse_market(ev, m)
                if x:
                    rows.append(x)
    return pd.DataFrame(rows).drop_duplicates("slug")


def fetch_odds(token, rdate):
    cut = cut_ts(rdate)
    try:
        js = _get(CLOB + "/prices-history", {"market": token, "startTs": cut - ODDS_WINDOW_DAYS * 86400, "endTs": cut + 3600,
                                             "fidelity": 60}) or {}
        return [(int(p["t"]), float(p["p"])) for p in js.get("history", []) or [] if int(p["t"]) <= cut]
    except Exception:
        return []


def clob_mid(token):
    try:
        return _f((_get(CLOB + "/midpoint", {"token_id": token}, tries=2) or {}).get("mid"))
    except Exception:
        return None


# 2. Kalshi KPI contracts (Part 1 logic, with slower requests and a merge with earlier runs)
def k_series():
    out = {}
    for params in ({"tags": "KPIs"}, {"category": "Financials"}, {"category": "Companies"}):
        try:
            js = _get(KALSHI + "/series", params, pause=0.2) or {}
        except Exception:
            continue
        for s in js.get("series") or []:
            tags = [str(t).lower() for t in (s.get("tags") or [])]
            if "kpis" in tags or "kpi" in str(s.get("title", "")).lower():
                out[s["ticker"]] = s.get("title", "")
    return out


def k_markets(series):
    rows = []
    for url, extra in ((KALSHI + "/historical/markets", {}), (KALSHI + "/markets", {"status": "settled"})):
        cursor = None
        for _ in range(50):
            params = {"series_ticker": series, "limit": 1000, **extra}
            if cursor:
                params["cursor"] = cursor
            try:
                js = _get(url, params, tries=5, pause=0.25) or {}
            except Exception:
                break
            rows += js.get("markets") or []
            cursor = js.get("cursor")
            if not cursor or not js.get("markets"):
                break
    return rows


def _k_parse(series, m):
    res = str(m.get("result") or "").lower()
    if res not in ("yes", "no"):
        return None
    fs, cap = _f(m.get("floor_strike")), m.get("cap_strike")
    st = str(m.get("strike_type") or "").lower()
    sub = str(m.get("yes_sub_title") or m.get("subtitle") or "").lower()
    title = str(m.get("title") or "")
    above = st in ("greater", "greater_or_equal") or (st in ("structured", "custom", "") and
                                                      (sub.startswith("above") or sub.startswith("at least") or " above " in title.lower()))
    cts = _ts(m.get("close_time"))
    if fs is None or cap is not None or not above or not cts:
        return None
    return dict(mticker=m.get("ticker"), series=series, event=m.get("event_ticker"), title=title, strike=fs,
                actual=_f(m.get("expiration_value")), y=1 if res == "yes" else 0,
                volume=_f(m.get("volume_fp")) or _f(m.get("volume")) or 0.0, close_ts=cts)


def k_candidates(series, extra=None):
    if series in K_OVERRIDES:
        return [K_OVERRIDES[series]] if K_OVERRIDES[series] else []
    s = series[2:] if series.startswith("KX") else series
    c = [s] + ([s[:-1]] if s.endswith("A") and len(s) > 2 else [])
    c = [x for x in c if 1 <= len(x) <= 5]
    a = (extra or {}).get(series)
    return c + [a] if a and a not in c else c


def _k_price(c):
    def g(d, *keys):
        for k in keys:
            v = _f((d or {}).get(k))
            if v is not None:
                return v / 100 if v > 1.0001 else v
        return None
    p = g(c.get("price"), "close_dollars", "close", "previous_dollars", "previous")
    if p is None:
        b, a = g(c.get("yes_bid"), "close_dollars", "close"), g(c.get("yes_ask"), "close_dollars", "close")
        p = (a + b) / 2 if a is not None and b is not None and a - b <= 0.2 else None
    return p


def k_odds(row):
    t, cut = row["mticker"], int(row["cut"])
    p = {"start_ts": cut - ODDS_WINDOW_DAYS * 86400, "end_ts": cut + 3600, "period_interval": 60}
    for url, extra, key in ((f"{KALSHI}/historical/markets/{t}/candlesticks", {}, None),
                            (f"{KALSHI}/markets/candlesticks", {"market_tickers": t}, "markets"),
                            (f"{KALSHI}/series/{row['series']}/markets/{t}/candlesticks", {}, None)):
        try:
            js = _get(url, {**p, **extra}, tries=3, pause=0.25) or {}
        except Exception:
            continue
        cs = (js.get("markets") or [{}])[0].get("candlesticks", []) if key else js.get("candlesticks", [])
        pts = [(int(c["end_period_ts"]), _k_price(c)) for c in cs or [] if c.get("end_period_ts")]
        pts = [(a, b) for a, b in pts if b is not None and a <= cut]
        if pts:
            return sorted(pts)
    return []


def k_resolve(km, est_by, extra=None):
    pick = {}
    for s, g in km.groupby("series"):
        best, score = None, 0
        days = [pd.Timestamp(c - 5 * 3600, unit="s").normalize() for c in g["close_ts"]]
        for c in k_candidates(s, extra):
            ds = [x["date"] for x in est_by.get(c, [])]
            n = sum(any(abs((d - e).days) <= 2 for e in ds) for d in days)
            if n > score:
                best, score = c, n
        pick[s] = best
    out = []
    for r in km.to_dict("records"):
        t = pick.get(r["series"])
        if not t:
            continue
        d0 = pd.Timestamp(r["close_ts"] - 5 * 3600, unit="s").normalize()
        near = [x["date"] for x in est_by.get(t, []) if abs((x["date"] - d0).days) <= 2]
        rd = min(near, key=lambda e: abs((e - d0).days)) if near else d0
        cut = min(cut_ts(rd), r["close_ts"] - 3600)
        out.append({**r, "ticker": t, "report_date": rd, "cut": cut, "date_matched": bool(near)})
    return pd.DataFrame(out)


# 3. Yahoo Finance
def fetch_prices(tickers, start):
    yf = _quiet_yf()
    syms = sorted({yf_symbol(t) for t in tickers} | {MARKET})
    data = yf.download(syms, start=str(pd.Timestamp(start).date()), auto_adjust=True, progress=False, group_by="column", threads=True)
    px = data["Close"] if isinstance(data.columns, pd.MultiIndex) else data[["Close"]].rename(columns={"Close": syms[0]})
    px.index = pd.to_datetime(px.index).tz_localize(None).normalize()
    px.index.name = "Date"
    return px.dropna(how="all")


def fetch_estimates(ticker):
    yf = _quiet_yf()
    try:
        e = yf.Ticker(yf_symbol(ticker)).get_earnings_dates(limit=24)
        if e is None or e.empty:
            return []
        e = e.reset_index()
        dcol = e.columns[0]
        out = []
        for _, r in e.iterrows():
            est, rep = r.get("EPS Estimate"), r.get("Reported EPS")
            if pd.notna(est) and pd.notna(rep):
                out.append(dict(ticker=ticker, date=pd.Timestamp(r[dcol]).tz_localize(None).normalize(),
                                estimate=float(est), reported=float(rep),
                                surprise_pct=float(r["Surprise(%)"]) if pd.notna(r.get("Surprise(%)")) else None))
        return out
    except Exception:
        return []


def _row(df, key):
    try:
        if df is None or len(df) == 0 or key not in df.index:
            return {}
        return {str(k): _f(v) for k, v in df.loc[key].items()}
    except Exception:
        return {}


def _pick(d, *names):
    low = {k.lower(): v for k, v in d.items()}
    for n in names:
        if n.lower() in low and low[n.lower()] is not None:
            return low[n.lower()]
    return None


def snapshot_ticker(ticker, rdate):
    yf = _quiet_yf()
    out = {}
    t = None
    try:
        t = yf.Ticker(yf_symbol(ticker))
        h = t.history(period="1mo", auto_adjust=True)["Close"].dropna()
        spot = float(h.iloc[-1]) if len(h) else None
        out["spot"] = spot
        exps = [e for e in (t.options or []) if pd.Timestamp(e) > pd.Timestamp(rdate)]
        if spot and exps:
            ch = t.option_chain(exps[0])
            out["expiry"] = exps[0]

            def mid(df, k):
                x = df[df["strike"] == k]
                if x.empty:
                    return None, None, None
                b, a = (_f(x.iloc[0][c]) or 0.0 for c in ("bid", "ask"))
                if a > 0 and b > 0 and a >= b and (a - b) <= 0.5 * (a + b):
                    return (a + b) / 2, b, a
                return None, None, None

            both = sorted(set(ch.calls["strike"]) & set(ch.puts["strike"]), key=lambda k: abs(k - spot))
            if both:
                c, p = mid(ch.calls, both[0])[0], mid(ch.puts, both[0])[0]
                out["atm_strike"] = float(both[0])
                out["implied_move"] = (c + p) / spot if c and p else None
            below = [k for k in ch.puts["strike"] if k <= spot]
            for lab, mny in (("put95", 1 - PUT_OTM), ("put90", 0.90)):
                ks = sorted(below, key=lambda k: abs(k - mny * spot))
                if ks:
                    pm, pb, pa = mid(ch.puts, ks[0])
                    out[f"{lab}_strike"] = float(ks[0])
                    out[f"{lab}_cost"] = pm / spot if pm else None
                    if lab == "put95":
                        out["put95_bid"], out["put95_ask"] = pb, pa
                        xl = ch.puts[ch.puts["strike"] == ks[0]]
                        if len(xl):
                            lp = _f(xl.iloc[0].get("lastPrice"))
                            out["put95_last"] = lp / spot if lp else None
    except Exception:
        pass
    if t is not None:
        try:
            tr = _row(t.eps_trend, "0q")
            out.update(eps_now=_pick(tr, "current"), eps_7d_ago=_pick(tr, "7daysAgo"), eps_30d_ago=_pick(tr, "30daysAgo"),
                       eps_90d_ago=_pick(tr, "90daysAgo"))
        except Exception:
            pass
        try:
            rv = _row(t.eps_revisions, "0q")
            out.update(rev_up_7d=_pick(rv, "upLast7days"), rev_down_7d=_pick(rv, "downLast7days", "downLast7Days"),
                       rev_up_30d=_pick(rv, "upLast30days"), rev_down_30d=_pick(rv, "downLast30days", "downLast30Days"))
        except Exception:
            pass
        try:
            es = _row(t.earnings_estimate, "0q")
            out.update(n_analysts=_pick(es, "numberOfAnalysts"), eps_low=_pick(es, "low"), eps_high=_pick(es, "high"),
                       eps_avg=_pick(es, "avg"))
        except Exception:
            pass
    return out


# 4. Downloads
def take_snapshot(mk):
    now = pd.Timestamp.now(tz="UTC").tz_localize(None)
    m = mk.copy()
    m["report_date"] = pd.to_datetime(m["report_date"])
    up = m[(~m["closed"].astype(bool)) & m["ticker"].notna() & (m["report_date"] >= now.normalize())
           & (m["report_date"] <= now.normalize() + pd.Timedelta(days=SNAPSHOT_DAYS))]
    print(f"   reports in the next {SNAPSHOT_DAYS} days: {len(up)}", flush=True)
    if up.empty:
        return
    with ThreadPoolExecutor(THREADS) as ex:
        mids = list(ex.map(clob_mid, up["token"]))
        snaps = list(ex.map(lambda a: snapshot_ticker(*a), zip(up["ticker"], up["report_date"])))
    rows = []
    for (_, r), mid, s in zip(up.iterrows(), mids, snaps):
        rows.append(dict(snap_utc=now.strftime("%Y-%m-%d %H:%M"), slug=r["slug"], ticker=r["ticker"],
                         report_date=r["report_date"].strftime("%Y-%m-%d"), p_gamma=r["p_now"], p_mid=mid, volume=r["volume"], **s))
    new = pd.DataFrame(rows).reindex(columns=SNAP_COLS)
    old = pd.read_csv(_p(F_SNAP)) if os.path.exists(_p(F_SNAP)) else pd.DataFrame(columns=SNAP_COLS)
    allx = pd.concat([old, new], ignore_index=True)
    allx.to_csv(_p(F_SNAP), index=False)
    print(f"   saved {len(new)} rows ({new['implied_move'].notna().sum()} with option prices, "
          f"{new['eps_30d_ago'].notna().sum()} with analyst revisions). Snapshots so far: {len(allx)}", flush=True)


def download_all(mk):
    use = mk[mk["resolved"].astype(bool) & mk["ticker"].notna() & (mk["volume"] >= MIN_VOLUME)
             & (pd.to_datetime(mk["report_date"]) >= pd.Timestamp(START_DATE))].copy()
    use["report_date"] = pd.to_datetime(use["report_date"])
    old = pd.read_csv(_p(F_ODDS)) if os.path.exists(_p(F_ODDS)) else pd.DataFrame(columns=["slug", "ts", "p"])
    todo = use[~use["slug"].isin(set(old["slug"]))]
    print(f"   settled contracts: {len(use)}, odds saved earlier: {len(use) - len(todo)}, downloading: {len(todo)}", flush=True)
    with ThreadPoolExecutor(THREADS) as ex:
        got = list(ex.map(lambda a: fetch_odds(*a), zip(todo["token"], todo["report_date"])))
    new = pd.DataFrame([(s, t, p) for s, pts in zip(todo["slug"], got) for t, p in pts], columns=["slug", "ts", "p"])
    pd.concat([old, new], ignore_index=True).drop_duplicates(["slug", "ts"]).to_csv(_p(F_ODDS), index=False)

    km = pd.DataFrame()
    if USE_KALSHI:
        print("3/6 Kalshi KPI contracts (slow on purpose: Kalshi limits request rates)", flush=True)
        try:
            ser = k_series()
            with ThreadPoolExecutor(K_THREADS) as ex:
                raw = list(ex.map(k_markets, list(ser)))
            rows = [x for s, ms in zip(ser, raw) for m in ms for x in [_k_parse(s, m)] if x]
            km = pd.DataFrame(rows)
            print(f"   KPI series: {len(ser)}, settled 'above X' contracts: {len(km)}", flush=True)
            if len(km):
                km.assign(series_title=km["series"].map(ser)).to_csv(_p(F_KALL), index=False)
        except Exception as e:
            print(f"   Kalshi download failed, keeping earlier Kalshi data: {e}", flush=True)
    kai = {}
    try:
        kai = kalshi_ai_map(km)
    except Exception as e:
        print(f"   Generative AI Kalshi mapping failed, rules only: {e}", flush=True)
    kcand = sorted({c for s in (km["series"].unique() if len(km) else []) for c in k_candidates(s, kai)})

    tick = sorted(set(use["ticker"]) | set(mk.loc[~mk["closed"].astype(bool) & mk["ticker"].notna(), "ticker"]))
    print(f"4/6 Yahoo Finance: EPS history for {len(tick) + len(kcand)} tickers", flush=True)
    with ThreadPoolExecutor(THREADS) as ex:
        est_all = dict(zip(tick + kcand, ex.map(fetch_estimates, tick + kcand)))
    e = pd.DataFrame([x for v in est_all.values() for x in v])
    if os.path.exists(_p(F_EPS)):
        o = pd.read_csv(_p(F_EPS), parse_dates=["date"])
        e = pd.concat([o[~o["ticker"].isin(set(e["ticker"]) if len(e) else set())], e], ignore_index=True)
    e["date"] = pd.to_datetime(e["date"]).dt.strftime("%Y-%m-%d")
    e.drop_duplicates(["ticker", "date"], keep="last").to_csv(_p(F_EPS), index=False)

    if len(km):
        km = km.drop_duplicates("mticker")
        km = km[(km["close_ts"] >= _ts(START_DATE)) & (km["volume"] >= K_MIN_VOLUME)]
        km = km.sort_values("volume", ascending=False).drop_duplicates("event")
        try:
            kalshi_map_count(km, est_all, kai)
        except Exception as e:
            print(f"   Kalshi mapping count failed: {e}", flush=True)
        km = k_resolve(km, est_all, kai)
        if len(km):
            km["report_date"] = pd.to_datetime(km["report_date"]).dt.strftime("%Y-%m-%d")
    if os.path.exists(_p(F_KM)):
        ok = pd.read_csv(_p(F_KM))
        km = pd.concat([ok, km], ignore_index=True).drop_duplicates("mticker", keep="last") if len(km) else ok
    if len(km):
        km = km.sort_values(["report_date", "mticker"]).reset_index(drop=True)
        km.to_csv(_p(F_KM), index=False)
        kold = pd.read_csv(_p(F_KODDS)) if os.path.exists(_p(F_KODDS)) else pd.DataFrame(columns=["market", "ts", "p"])
        ktodo = km[~km["mticker"].isin(set(kold["market"]))]
        with ThreadPoolExecutor(K_THREADS) as ex:
            kp = list(ex.map(k_odds, ktodo.to_dict("records")))
        knew = pd.DataFrame([(m, t, p) for m, pts in zip(ktodo["mticker"], kp) for t, p in pts], columns=["market", "ts", "p"])
        pd.concat([kold, knew], ignore_index=True).drop_duplicates(["market", "ts"]).to_csv(_p(F_KODDS), index=False)
        print(f"   Kalshi KPI events matched to a stock: {len(km)} (odds downloaded for {len(ktodo)} new ones)", flush=True)

    ktick = sorted(set(km["ticker"])) if len(km) else []
    print(f"5/6 Yahoo Finance: daily prices for {len(set(tick) | set(ktick))} tickers from {HIST_START}", flush=True)
    px = fetch_prices(sorted(set(tick) | set(ktick)), HIST_START)
    px.to_csv(_p(F_PRICES))


def fetch_edgar(ev):
    if not USE_EDGAR:
        return
    if not SEC_USER_AGENT.strip():
        print("   EDGAR skipped: set SEC_USER_AGENT (your name and email) to check release times.", flush=True)
        return
    old = pd.read_csv(_p(F_EDGAR)) if os.path.exists(_p(F_EDGAR)) else pd.DataFrame(columns=["ticker", "report_date"])
    have = set(zip(old["ticker"].astype(str), old["report_date"].astype(str)))
    need = ev[["ticker", "report_date"]].drop_duplicates()
    need = need[[(t, d.strftime("%Y-%m-%d")) not in have for t, d in zip(need["ticker"], need["report_date"])]]
    if need.empty:
        return
    print(f"   EDGAR release times for {len(need)} reports", flush=True)
    sec = requests.Session(); sec.headers.update({"User-Agent": SEC_USER_AGENT.strip()})

    def sget(url):
        for i in range(4):
            time.sleep(0.12)
            r = sec.get(url, timeout=30)
            if r.status_code == 429:
                time.sleep(2 * (i + 1)); continue
            r.raise_for_status()
            return r.json()
        return {}
    try:
        tk = sget("https://www.sec.gov/files/company_tickers.json")
    except Exception as e:
        print(f"   EDGAR not reachable, skipped: {e}", flush=True)
        return
    cik = {str(v["ticker"]).upper(): int(v["cik_str"]) for v in tk.values()}
    today = pd.Timestamp.now().normalize()
    rows = []
    for t, g in need.groupby("ticker"):
        c = cik.get(str(t).upper()) or cik.get(str(t).upper().replace(".", "-"))
        fl = pd.DataFrame()
        if c:
            try:
                js = sget(f"https://data.sec.gov/submissions/CIK{c:010d}.json")
                fl = pd.DataFrame(js["filings"]["recent"])
                first = g["report_date"].min() - pd.Timedelta(days=5)
                if len(fl) and pd.to_datetime(fl["filingDate"]).min() > first:
                    for f in js["filings"].get("files", []):
                        if pd.Timestamp(f.get("filingTo", "1900-01-01")) >= first:
                            fl = pd.concat([fl, pd.DataFrame(sget("https://data.sec.gov/submissions/" + f["name"]))], ignore_index=True)
                fl = fl[(fl["form"] == "8-K") & fl["items"].astype(str).str.contains("2.02", regex=False)].copy()
                fl["fd"] = pd.to_datetime(fl["filingDate"])
            except Exception:
                fl = pd.DataFrame()
        for rd in g["report_date"]:
            x = fl[(fl["fd"] >= rd - pd.Timedelta(days=3)) & (fl["fd"] <= rd + pd.Timedelta(days=1))] if len(fl) else fl
            if len(x):
                x = x.assign(gap=(x["fd"] - rd).abs()).sort_values("gap").iloc[0]
                rows.append(dict(ticker=t, report_date=rd.strftime("%Y-%m-%d"), cik=c, accession=x["accessionNumber"],
                                 filing_date=x["filingDate"], acceptance=x["acceptanceDateTime"]))
            elif rd < today - pd.Timedelta(days=10):
                rows.append(dict(ticker=t, report_date=rd.strftime("%Y-%m-%d"), cik=c, accession=None, filing_date=None, acceptance=None))
    new = pd.DataFrame(rows)
    pd.concat([old, new], ignore_index=True).drop_duplicates(["ticker", "report_date"], keep="last").to_csv(_p(F_EDGAR), index=False)
    print(f"   EDGAR: {new['acceptance'].notna().sum() if len(new) else 0} release times found", flush=True)


REL_KEYS = re.compile(r"guidance|outlook|expect|forecast|reaffirm|raise|rais|lower|reduc|full[- ]year|fiscal (?:year|20)|withdr|"
                      r"headwind|uncertain|challeng|tariff|demand|margin", re.I)
KW_UP = re.compile(r"\b(rais(?:e|es|ed|ing)|increas(?:e|es|ed|ing)|record|strong(?:er)?|above)\b", re.I)
KW_DOWN = re.compile(r"\b(lower(?:s|ed|ing)?|reduc(?:e|es|ed|ing)|cut|weak(?:er|ness)?|below|headwinds?|declin(?:e|es|ed|ing))\b", re.I)


def _sec_session():
    sec = requests.Session(); sec.headers.update({"User-Agent": SEC_USER_AGENT.strip()})
    return sec


def _sget(sec, url, as_json=True):
    for i in range(4):
        time.sleep(0.12)
        r = sec.get(url, timeout=30)
        if r.status_code == 429:
            time.sleep(2 * (i + 1)); continue
        if r.status_code == 404:
            return None
        r.raise_for_status()
        return r.json() if as_json else r.text
    return None


def _release_excerpt(html):
    import html as _h
    t = re.sub(r"(?is)<(script|style).*?</\1>", " ", html)
    t = _h.unescape(re.sub(r"(?s)<[^>]+>", " ", t))
    t = re.sub(r"\s+", " ", t).strip()
    sents = re.split(r"(?<=[.!?])\s+", t)
    head, keep, n = " ".join(sents[:6])[:900], [], 0
    for x in sents[6:]:
        if REL_KEYS.search(x) and 20 < len(x) < 600:
            keep.append(x); n += len(x)
            if n > 2600:
                break
    return head + " ... " + " ".join(keep), t


COVER_RE = r"(?i)securities and exchange commission|form 8-k|current report pursuant|commission file number"


def fetch_prior_releases(ev, snap=None):
    if not USE_EDGAR or not SEC_USER_AGENT.strip():
        return
    need = ev[["ticker", "report_date"]].copy()
    if snap is not None and len(snap):
        up = snap[["ticker", "report_date"]].copy(); up["report_date"] = pd.to_datetime(up["report_date"])
        need = pd.concat([need, up])
    need = need.drop_duplicates()
    old = pd.read_csv(_p(F_PRIOR)) if os.path.exists(_p(F_PRIOR)) else pd.DataFrame(columns=["ticker", "report_date"])
    if len(old) and "excerpt" in old:
        cover = old["excerpt"].astype(str).str[:400].str.contains(COVER_RE, regex=True)
        if cover.any():
            bad_acc = set(old.loc[cover, "accession"].dropna())
            old = old[~cover]
            old.to_csv(_p(F_PRIOR), index=False)
            if os.path.exists(_p(F_AI_REL)):
                ar_ = pd.read_csv(_p(F_AI_REL))
                ar_[~ar_["accession"].isin(bad_acc)].to_csv(_p(F_AI_REL), index=False)
            print(f"   EDGAR: {int(cover.sum())} earlier readings were the 8-K cover page, not the press release: read again", flush=True)
    have = set(zip(old["ticker"].astype(str), old["report_date"].astype(str)))
    need = need[[(t, d.strftime("%Y-%m-%d")) not in have for t, d in zip(need["ticker"], need["report_date"])]]
    if need.empty:
        return
    print(f"   EDGAR: last quarter's earnings release for {len(need)} reports", flush=True)
    sec = _sec_session()
    try:
        tk = _sget(sec, "https://www.sec.gov/files/company_tickers.json")
    except Exception as e:
        print(f"   EDGAR not reachable, skipped: {e}", flush=True)
        return
    cik = {str(v["ticker"]).upper(): int(v["cik_str"]) for v in tk.values()}
    rows, docs = [], {}
    for t, g in need.groupby("ticker"):
        c = cik.get(str(t).upper()) or cik.get(str(t).upper().replace(".", "-"))
        fl = pd.DataFrame()
        if c:
            try:
                js = _sget(sec, f"https://data.sec.gov/submissions/CIK{c:010d}.json")
                fl = pd.DataFrame(js["filings"]["recent"])
                first = g["report_date"].min() - pd.Timedelta(days=140)
                if len(fl) and pd.to_datetime(fl["filingDate"]).min() > first:
                    for f in js["filings"].get("files", []):
                        if pd.Timestamp(f.get("filingTo", "1900-01-01")) >= first:
                            fl = pd.concat([fl, pd.DataFrame(_sget(sec, "https://data.sec.gov/submissions/" + f["name"]))], ignore_index=True)
                fl = fl[(fl["form"] == "8-K") & fl["items"].astype(str).str.contains("2.02", regex=False)].copy()
                fl["fd"] = pd.to_datetime(fl["filingDate"])
            except Exception:
                fl = pd.DataFrame()
        for rd in g["report_date"]:
            x = fl[(fl["fd"] >= rd - pd.Timedelta(days=130)) & (fl["fd"] <= rd - pd.Timedelta(days=20))] if len(fl) else fl
            row = dict(ticker=t, report_date=rd.strftime("%Y-%m-%d"), cik=c, accession=None, filing_date=None, document=None,
                       excerpt=None, kw_up=None, kw_down=None)
            if len(x):
                x = x.sort_values("fd").iloc[-1]
                acc = x["accessionNumber"]
                row.update(accession=acc, filing_date=x["filingDate"])
                if acc not in docs:
                    docs[acc] = (None, None, None)
                    try:
                        idx = _sget(sec, f"https://www.sec.gov/Archives/edgar/data/{c}/{acc.replace('-', '')}/index.json")
                        its = [i for i in (idx or {}).get("directory", {}).get("item", [])
                               if str(i.get("name", "")).lower().endswith((".htm", ".html", ".txt")) and "index" not in str(i.get("name", "")).lower()]
                        prim = str(x.get("primaryDocument") or "")
                        ex = [i["name"] for i in its if re.search(r"ex(hibit)?[-_ .]?99|99[-_.]?0?1|press|release|earnings", i["name"], re.I)
                              and i["name"] != prim]
                        rest = sorted([i for i in its if i["name"] != prim], key=lambda i: -(_f(i.get("size")) or 0))
                        name = sorted(ex)[0] if ex else (rest[0]["name"] if rest else prim)
                        html = _sget(sec, f"https://www.sec.gov/Archives/edgar/data/{c}/{acc.replace('-', '')}/{name}", as_json=False)
                        if html:
                            exc, full = _release_excerpt(html)
                            docs[acc] = (name, exc, (len(KW_UP.findall(full)), len(KW_DOWN.findall(full))))
                    except Exception:
                        pass
                name, exc, kw = docs[acc]
                row.update(document=name, excerpt=exc, kw_up=kw[0] if kw else None, kw_down=kw[1] if kw else None)
            rows.append(row)
    new = pd.DataFrame(rows)
    pd.concat([old, new], ignore_index=True).drop_duplicates(["ticker", "report_date"], keep="last").to_csv(_p(F_PRIOR), index=False)
    print(f"   EDGAR: {new['excerpt'].notna().sum()} previous releases read", flush=True)


def score_releases():
    if not os.path.exists(_p(F_PRIOR)):
        return
    pr = pd.read_csv(_p(F_PRIOR)).dropna(subset=["accession", "excerpt"]).drop_duplicates("accession")
    old = pd.read_csv(_p(F_AI_REL)) if os.path.exists(_p(F_AI_REL)) else pd.DataFrame(columns=["accession"])
    todo = pr[~pr["accession"].isin(set(old["accession"]))]
    if todo.empty:
        return
    if genai("Reply with the word ok.") is None:
        print(f"   Generative AI not available: {len(todo)} earnings releases not read yet (they will be read on the next run)", flush=True)
        return
    print(f"   Generative AI reading {len(todo)} earnings releases ({GENAI_BATCH} per call)", flush=True)
    out, fails = [], 0
    recs = todo.to_dict("records")
    for k in range(0, len(recs), GENAI_BATCH):
        chunk = recs[k:k + GENAI_BATCH]
        body = "\n\n".join(f"ID {i}: {r_['ticker']}, filed {r_['filing_date']}\n{str(r_['excerpt'])[:3200]}" for i, r_ in enumerate(chunk))
        prompt = ("You read company earnings press releases for a portfolio manager. For each release below, reply with a JSON list, "
                  "one object per ID, keys: id (number), guidance (one of: raised, maintained, lowered, withdrawn, not given; about the "
                  "company's own outlook for future periods), tone (integer from -2 very negative to 2 very positive), risks (integer "
                  "0 to 3: how many material headwinds or risks management flags), summary (at most 15 words). JSON only.\n\n" + body)
        t = genai(prompt) or ""
        mm = re.search(r"\[.*\]", t, re.S)
        try:
            js = json.loads(mm.group(0)) if mm else []
        except Exception:
            js = []
        got = {int(j.get("id", -1)): j for j in js if isinstance(j, dict) and str(j.get("id", "")).lstrip("-").isdigit()}
        for i, r_ in enumerate(chunk):
            j = got.get(i)
            if j is None:
                fails += 1
                continue
            g = str(j.get("guidance", "")).lower().strip()
            out.append(dict(accession=r_["accession"], guidance=g if g in ("raised", "maintained", "lowered", "withdrawn", "not given") else "not given",
                            tone=_f(j.get("tone")), risks=_f(j.get("risks")), summary=" ".join(str(j.get("summary", "")).split())[:200],
                            model=GEMINI_MODEL if GEMINI_API_KEY else "google.colab.ai (Gemini)"))
        if fails > 3 * GENAI_BATCH and not out:
            print("   Generative AI calls are failing (daily limit?). Stopping here, the rest is read on the next run.", flush=True)
            break
        if (k // GENAI_BATCH) % 25 == 24:
            pd.concat([old, pd.DataFrame(out)], ignore_index=True).to_csv(_p(F_AI_REL), index=False)
            print(f"   ... {len(out)} releases read", flush=True)
    pd.concat([old, pd.DataFrame(out)], ignore_index=True).drop_duplicates("accession", keep="last").to_csv(_p(F_AI_REL), index=False)
    print(f"   Generative AI read {len(out)} releases ({fails} without an answer, retried next run)", flush=True)


def fetch_sectors(tickers):
    old = pd.read_csv(_p(F_SECTOR)) if os.path.exists(_p(F_SECTOR)) else pd.DataFrame(columns=["ticker", "sector", "industry"])
    todo = sorted(set(map(str, tickers)) - set(old["ticker"].astype(str)))
    if not todo:
        return
    print(f"   Yahoo Finance: sector for {len(todo)} companies", flush=True)
    yf = _quiet_yf()

    def one(t):
        try:
            inf = yf.Ticker(yf_symbol(t)).info or {}
            return dict(ticker=t, sector=inf.get("sector"), industry=inf.get("industry"))
        except Exception:
            return dict(ticker=t, sector=None, industry=None)
    with ThreadPoolExecutor(max_workers=6) as pool:
        new = list(pool.map(one, todo))
    pd.concat([old, pd.DataFrame(new)], ignore_index=True).drop_duplicates("ticker", keep="last").to_csv(_p(F_SECTOR), index=False)


def load_releases():
    cols = ["ticker", "report_date", "accession", "filing_date", "kw_up", "kw_down", "guidance", "tone", "risks", "summary"]
    if not os.path.exists(_p(F_PRIOR)):
        return pd.DataFrame(columns=cols)
    pr = pd.read_csv(_p(F_PRIOR))
    ai = pd.read_csv(_p(F_AI_REL)) if os.path.exists(_p(F_AI_REL)) else pd.DataFrame(columns=["accession", "guidance", "tone", "risks", "summary"])
    return pr.merge(ai.drop_duplicates("accession"), on="accession", how="left").reindex(columns=cols)


GUIDE_SCORE = {"raised": 1.0, "maintained": 0.0, "not given": 0.0, "lowered": -1.0, "withdrawn": -1.0}


def add_text_features(df, rel):
    r = rel.copy()
    r["report_date"] = pd.to_datetime(r["report_date"])
    r["guid_score"] = r["guidance"].map(GUIDE_SCORE)
    r["kw_net"] = (r["kw_up"] - r["kw_down"]) / (r["kw_up"] + r["kw_down"]).where(lambda x: x > 0)
    cols = ["guid_score", "tone", "risks", "kw_net", "guidance", "summary"]
    df = df.drop(columns=[c for c in cols if c in df]).merge(r[["ticker", "report_date"] + cols].drop_duplicates(["ticker", "report_date"]),
                                                             on=["ticker", "report_date"], how="left")
    return df


def signed_history(est, px):
    dates, lpx = px.index, np.log(px)
    rows = []
    for t, d in zip(est["ticker"], est["date"]):
        sym = yf_symbol(t)
        a = dates.searchsorted(d - pd.Timedelta(days=1), side="right") - 1
        b = dates.searchsorted(d, side="right")
        if sym in px.columns and a >= 0 and b < len(dates):
            rows.append((t, d, (lpx[sym].iloc[b] - lpx[sym].iloc[a]) - (lpx[MARKET].iloc[b] - lpx[MARKET].iloc[a])))
    return pd.DataFrame(rows, columns=["ticker", "date", "move"]).dropna().drop_duplicates(["ticker", "date"])


def add_peer_features(df, est, px):
    sec_ = pd.read_csv(_p(F_SECTOR)) if os.path.exists(_p(F_SECTOR)) else pd.DataFrame(columns=["ticker", "sector"])
    smap = dict(zip(sec_["ticker"].astype(str), sec_["sector"]))
    h = signed_history(est, px)
    h["sector"] = h["ticker"].astype(str).map(smap)
    by = {k: g for k, g in h.dropna(subset=["sector"]).groupby("sector")}
    pm, pf, pn = [], [], []
    for t, d in zip(df["ticker"], pd.to_datetime(df["report_date"])):
        g = by.get(smap.get(str(t)))
        x = g[(g["ticker"] != t) & (g["date"] < d - pd.Timedelta(days=1)) & (g["date"] >= d - pd.Timedelta(days=PEER_DAYS))] if g is not None else []
        n_ = len(x)
        pn.append(n_); pm.append(x["move"].mean() if n_ >= 3 else np.nan); pf.append((x["move"] <= FALL).mean() if n_ >= 3 else np.nan)
    return df.assign(peer_n=pn, peer_mean=pm, peer_fall=pf)


def _mget(url, params):
    for i in range(4):
        time.sleep(12.5)
        r = requests.get(url, params={**params, "apiKey": MASSIVE_API_KEY.strip()}, timeout=30)
        if r.status_code == 429:
            time.sleep(60); continue
        if r.status_code in (401, 403):
            raise PermissionError(f"Massive refused the request ({r.status_code}): {r.text[:200]}")
        r.raise_for_status()
        return r.json()
    return {}


def _bars(opt_ticker, d_from, d_to):
    js = _mget(f"{MASSIVE}/v2/aggs/ticker/{opt_ticker}/range/1/day/{d_from:%Y-%m-%d}/{d_to:%Y-%m-%d}",
               {"adjusted": "true", "sort": "asc", "limit": 50})
    return [(pd.Timestamp(b["t"], unit="ms").normalize(), float(b["c"])) for b in js.get("results") or []]


def option_history_one(rec, px):
    sym, rd, dates = yf_symbol(rec["ticker"]), pd.Timestamp(rec["report_date"]), px.index
    i0 = dates.searchsorted(rd - pd.Timedelta(days=1), side="right") - 1
    i1 = dates.searchsorted(rd, side="right")
    base = dict(slug=rec["slug"], ticker=rec["ticker"], report_date=rd.strftime("%Y-%m-%d"))
    if i0 < 0 or i1 >= len(dates) or sym not in px.columns or not np.isfinite(px[sym].iloc[i0]):
        return {**base, "status": "no share prices"}
    d0, d1 = dates[i0], dates[i1]
    spot, spot1 = float(px[sym].iloc[i0]), float(px[sym].iloc[i1])
    base.update(entry_date=d0.strftime("%Y-%m-%d"), exit_date=d1.strftime("%Y-%m-%d"), spot=spot, spot_after=spot1)
    q = {"underlying_ticker": str(rec["ticker"]).upper(), "contract_type": "put", "expiration_date.gte": d1.strftime("%Y-%m-%d"),
         "expiration_date.lte": (d1 + pd.Timedelta(days=21)).strftime("%Y-%m-%d"), "strike_price.gte": round(0.85 * spot, 2),
         "strike_price.lte": round(1.05 * spot, 2), "as_of": d0.strftime("%Y-%m-%d"), "limit": 1000}
    res = (_mget(f"{MASSIVE}/v3/reference/options/contracts", {**q, "expired": "true"}).get("results") or [])
    if not res:
        res = (_mget(f"{MASSIVE}/v3/reference/options/contracts", {**q, "expired": "false"}).get("results") or [])
    if not res:
        return {**base, "status": "no contracts"}
    c = pd.DataFrame(res)
    c = c[c["expiration_date"] == c["expiration_date"].min()]
    base["expiry"] = c["expiration_date"].iloc[0]
    atm = c.iloc[(c["strike_price"] - spot).abs().argsort()].iloc[0]
    below = c[c["strike_price"] <= spot]
    p95 = below.iloc[(below["strike_price"] - (1 - PUT_OTM) * spot).abs().argsort()].iloc[0] if len(below) else None
    for lab, row in ((("atm", atm),) if OPTION_HISTORY_ATM else ()) + (("put", p95),):
        if row is None:
            continue
        base[f"{lab}_strike"], base[f"{lab}_ticker"] = float(row["strike_price"]), row["ticker"]
        bars = _bars(row["ticker"], d0 - pd.Timedelta(days=7), d1)
        before = [(d, c_) for d, c_ in bars if d <= d0]
        if before:
            base[f"{lab}_entry_date"], base[f"{lab}_entry"] = before[-1][0].strftime("%Y-%m-%d"), before[-1][1]
        if lab == "put":
            after = [(d, c_) for d, c_ in bars if d == d1]
            if after:
                base["put_exit_date"], base["put_exit"] = d1.strftime("%Y-%m-%d"), after[-1][1]
    base["status"] = "ok"
    return base


def fetch_option_history(ev, px):
    if not USE_MASSIVE or not MASSIVE_API_KEY.strip() or OPTION_HISTORY_MINUTES <= 0:
        if USE_MASSIVE:
            print("   Option history skipped: set MASSIVE_API_KEY (free at massive.com) to download past option prices.", flush=True)
        return
    old = pd.read_csv(_p(F_OPTHIST)) if os.path.exists(_p(F_OPTHIST)) else pd.DataFrame(columns=OH_COLS)
    limit_date = pd.Timestamp.now().normalize() - pd.Timedelta(days=2 * 365 - 10)
    todo = ev[~ev["slug"].isin(set(old["slug"])) & (ev["report_date"] >= limit_date) & ev["ar"].notna()]
    todo = todo.sample(frac=1, random_state=SEED)
    print(f"   option history: {len(old)} reports saved, {len(todo)} to go, up to {OPTION_HISTORY_MINUTES} minutes this run "
          f"(about {140 if not OPTION_HISTORY_ATM else 95} reports an hour)", flush=True)
    t_end, done = time.time() + 60 * OPTION_HISTORY_MINUTES, 0
    for rec in todo.to_dict("records"):
        if time.time() > t_end:
            break
        try:
            row = option_history_one(rec, px)
        except PermissionError as e:
            print(f"   {e}", flush=True); break
        except Exception as e:
            row = dict(slug=rec["slug"], ticker=rec["ticker"], report_date=pd.Timestamp(rec["report_date"]).strftime("%Y-%m-%d"),
                       status=f"error: {str(e)[:60]}")
        pd.DataFrame([row]).reindex(columns=OH_COLS).to_csv(_p(F_OPTHIST), mode="a", header=not os.path.exists(_p(F_OPTHIST)), index=False)
        done += 1
    print(f"   option history: {done} reports added this run", flush=True)


LSEG_SUFFIX = {"Nasdaq": ".O", "NYSE": ".N", "NYSE American": ".A", "NYSE Arca": ".P", "CBOE": ".Z", "BATS": ".Z"}


def sec_exchanges():
    if not SEC_USER_AGENT.strip():
        return {}
    try:
        r = requests.get("https://www.sec.gov/files/company_tickers_exchange.json", headers={"User-Agent": SEC_USER_AGENT.strip()}, timeout=30)
        js = r.json(); f = js["fields"]
        return {str(row[f.index("ticker")]).upper(): row[f.index("exchange")] for row in js["data"]}
    except Exception:
        return {}


def to_ric(ticker, exch):
    base, _, cls = str(ticker).upper().partition(".")
    return base + cls.lower() + LSEG_SUFFIX.get(exch, ".N")


def lseg_collect():
    path = _p(LSEG_REQUEST)
    if not os.path.exists(path):
        return
    try:
        d = pd.read_excel(path, sheet_name="Request", header=4)
    except Exception as e:
        print(f"   LSEG request file could not be read: {e}", flush=True)
        return
    d = d.rename(columns={"IV before (entry)": "iv_entry", "IV after (exit)": "iv_exit", "RIC": "ric", "Code 1": "ric", "Contract": "slug",
                          "Ticker": "ticker", "Entry date": "entry_date", "Exit date": "exit_date", "EPS before (entry)": "eps_entry",
                          "EPS earlier": "eps_earlier", "Earlier date": "earlier_date", "IV earlier": "iv_earlier",
                          "SmartEstimate before (entry)": "smart_est", "IV call before (entry)": "iv_call"})
    d = d.loc[:, ~d.columns.duplicated()]
    for c in ("iv_entry", "iv_exit", "eps_entry", "eps_earlier", "iv_earlier", "smart_est", "iv_call"):
        d[c] = pd.to_numeric(d[c], errors="coerce") if c in d.columns else np.nan
    for c, raw in (("iv_entry", ("LSEG: before, code 1", "LSEG: before, code 2")), ("iv_exit", ("LSEG: after, code 1", "LSEG: after, code 2")),
                   ("eps_entry", ("EPS: before, code 1", "EPS: before, code 2")), ("eps_earlier", ("EPS: earlier, code 1", "EPS: earlier, code 2")),
                   ("iv_earlier", ("LSEG: IV earlier",)), ("smart_est", ("SmartEst: code 1", "SmartEst: code 2")),
                   ("iv_call", ("LSEG: IV call",))):
        for rc in raw:
            if rc in d.columns:
                d[c] = d[c].fillna(pd.to_numeric(d[rc], errors="coerce"))
    for c in LSEG_COLS:
        if c not in d.columns:
            d[c] = None
    got = d[(d["iv_entry"].notna() & d["iv_exit"].notna()) | (d["eps_entry"].notna() & d["eps_earlier"].notna())
            | d["smart_est"].notna() | d["iv_call"].notna()][LSEG_COLS]
    if got.empty:
        print("   LSEG request file found but no values yet: open it in Excel with Workspace, wait for the values, save.", flush=True)
        return
    got = got.drop_duplicates("slug", keep="last").set_index("slug")
    if os.path.exists(_p(F_LSEG)):
        old = pd.read_csv(_p(F_LSEG)).reindex(columns=LSEG_COLS).drop_duplicates("slug", keep="last").set_index("slug")
        got = got.combine_first(old)
    got.reset_index().reindex(columns=LSEG_COLS).to_csv(_p(F_LSEG), index=False)
    print(f"   LSEG: read from {LSEG_REQUEST}: {(got['iv_entry'].notna() & got['iv_exit'].notna()).sum()} reports with implied volatility, "
          f"{(got['eps_entry'].notna() & got['eps_earlier'].notna()).sum()} with consensus EPS, {got['smart_est'].notna().sum()} with StarMine", flush=True)


def lseg_request(ev, px):
    if not USE_LSEG:
        return 0
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill
    up = os.path.join(os.getcwd(), LSEG_REQUEST)
    if os.path.exists(up) and os.path.abspath(up) != os.path.abspath(_p(LSEG_REQUEST)):
        import shutil
        shutil.copy(up, _p(LSEG_REQUEST))
    lseg_collect()
    have = set(pd.read_csv(_p(F_LSEG))["slug"]) if os.path.exists(_p(F_LSEG)) else set()
    need = ev[ev["ar"].notna()] if LSEG_ALL_REPORTS else ev[~ev["slug"].isin(have) & ev["ar"].notna()]
    exch = sec_exchanges()
    dates = px.index
    rows = []
    for rec in need.to_dict("records"):
        rd = pd.Timestamp(rec["report_date"])
        i0 = dates.searchsorted(rd - pd.Timedelta(days=1), side="right") - 1
        i1 = dates.searchsorted(rd, side="right")
        ie = dates.searchsorted(rd - pd.Timedelta(days=1 + REV_DAYS), side="right") - 1
        if i0 >= 0 and i1 < len(dates) and ie >= 0:
            rows.append((rec["slug"], str(rec["ticker"]).upper(), dates[i0], dates[i1], dates[ie]))
    if not rows:
        return 0
    BLUE, BOLD, NOTE = Font(color="0000FF", bold=True), Font(bold=True), Font(italic=True, color="666666")
    YEL = PatternFill("solid", fgColor="FFFF99")
    wb = Workbook()
    rq = wb.active; rq.title = "Request"
    st = wb.create_sheet("Setup")
    st["A1"] = "LSEG Workspace request (BeatOdds). Open in Excel with the Workspace add-in signed in."; st["A1"].font = BOLD
    st["A2"] = "Field (implied volatility, exact name from Build Formula > Time Series)"
    st["B2"] = LSEG_FIELD; st["B2"].font = BLUE; st["B2"].fill = YEL
    st["A3"] = "Instrument: RIC (share codes .O and .N) or IV RIC (LSEG implied volatility codes, e.g. AAPLATMIV.U)"
    st["B3"] = LSEG_CODE_TYPE; st["B3"].font = BLUE; st["B3"].fill = YEL
    st["A4"] = "Consensus EPS field and its parameters (analyst revisions, read on share codes .O and .N)"
    st["B4"] = LSEG_EPS_FIELD; st["B4"].font = BLUE; st["B4"].fill = YEL
    st["C4"] = LSEG_EPS_PARAMS; st["C4"].font = BLUE; st["C4"].fill = YEL
    st["A5"] = "Test: Apple's implied volatility on one date. This cell should show a number:"; st["A5"].font = BOLD
    st["A7"] = "Test: Apple's consensus EPS on 1 Oct 2026 (should show a number, about 1.7 to 1.8):"
    st["B7"] = '=_xll.RDP.Data("AAPL.O",$B$4,$C$4&" SDate=2026-10-01 EDate=2026-10-01 Frq=D")'
    st["A6"] = '=IF($B$3="IV RIC","AAPLATMIV.U","AAPL.O")'
    st["B6"] = '=_xll.RDP.Data(A6,$B$2,"SDate=2026-10-01 EDate=2026-10-01 Frq=D")'
    st["D3"] = "StarMine SmartEstimate field"; st["D4"] = LSEG_SMART_FIELD; st["D4"].font = BLUE; st["D4"].fill = YEL
    st["E3"] = "Call implied volatility field (for the skew)"; st["E4"] = LSEG_CALL_FIELD; st["E4"].font = BLUE; st["E4"].fill = YEL
    st["C6"] = '=_xll.RDP.Data(A6,$E$4,"SDate=2026-10-01 EDate=2026-10-01 Frq=D")'
    st["D6"] = "<- call implied volatility for Apple on the same date (should be a number close to B6)"
    st["A8"] = "Test: Apple's StarMine SmartEstimate on 1 Oct 2026 (should show a number close to B7):"; st["A8"].font = BOLD
    st["B8"] = '=_xll.RDP.Data("AAPL.O",$D$4,$C$4&" SDate=2026-10-01 EDate=2026-10-01 Frq=D")'
    tips = ["Tested on Apple: TR.30DayAtTheMoneyImpliedVolatilityIndexForPutOptions on AAPLATMIV.U returns 26.21 on 1 Oct 2026. "
            "On the share code AAPL.O it returns NULL. So B3 is set to IV RIC.",
            "If B6 shows a number: the field works. Go to the Request sheet and wait until columns L and M fill in, then save.",
            "If B6 shows #NAME?: Workspace is not signed in or the add-in is not loaded. Sign in, then press Ctrl+Alt+F9.",
            "If B6 shows an error message or stays empty: the field name is wrong. In Workspace: Build Formula > Time Series, "
            "instrument AAPL.O, search 'implied volatility', copy the exact field name of the 30-day at-the-money implied "
            "volatility into B2. Every formula reads B2.",
            "If AAPL.O never returns the field: type IV RIC in B3.",
            "Columns H to K on Request are the raw LSEG answers (Nasdaq code and NYSE code). L and M keep whichever is a number.",
            "Columns P to S ask for the consensus EPS (B4) on the share codes: the day before the report and 30 days earlier. "
            "T and U keep the number. If B7 shows no number, try Period=FQ0 in C4 or another EPS field from Build Formula.",
            "Column W asks for the implied volatility on the earlier date (X keeps it). Columns Y and Z ask for the StarMine "
            "SmartEstimate (D4) the day before the report, AA keeps it. If B8 shows no number, search 'SmartEstimate' in Build Formula."]
    for i, t in enumerate(tips):
        st.cell(row=12 + i, column=1, value=t).font = NOTE
    st["A19"] = ("FIELD FINDER. Each cell asks LSEG for Apple on 1 Oct 2026 (or the latest value). Find a cell with a number, "
                 "copy its field name (column A) into B2 and set B3 to the code type of that column. Then press Ctrl+Alt+F9.")
    st["A19"].font = BOLD
    probe_codes = [("AAPL.O", "RIC", True), ("AAPLATMIV.U", "IV RIC", True), ("AAPL.O", "RIC", False), ("AAPLATMIV.U", "IV RIC", False)]
    for j, (code, kind, dated) in enumerate(probe_codes):
        st.cell(row=20, column=2 + j, value=code).font = BOLD
        st.cell(row=21, column=2 + j, value=f"{kind}, {'1 Oct 2026' if dated else 'latest value'}").font = NOTE
    for i, fld_ in enumerate(LSEG_PROBE_FIELDS):
        rr = 22 + i
        st.cell(row=rr, column=1, value=fld_)
        for j, (code, kind, dated) in enumerate(probe_codes):
            cl = chr(ord("B") + j)
            par_ = ',"SDate=2026-10-01 EDate=2026-10-01 Frq=D"' if dated else ""
            st.cell(row=rr, column=2 + j, value=f'=_xll.RDP.Data({cl}$20,$A{rr}{par_})')
    for c_ in "CDE":
        st.column_dimensions[c_].width = 22
    st.column_dimensions["A"].width = 100; st.column_dimensions["B"].width = 58
    rq["A1"] = ("LSEG request: one row per report. Columns H to K ask LSEG for implied volatility (code 1 and code 2, the day "
                "before and the day after). Columns L and M keep the number. Check the Setup sheet first.")
    rq["A1"].font = BOLD
    rq["A2"] = "Fields and instrument type: Setup sheet. When L, M, T, U, X, AA and AC show numbers, save the file and upload it to Colab."
    rq["A4"] = "Empty or error in H to K is normal for the code that does not exist (each share trades on Nasdaq or NYSE, not both)."
    rq["A4"].font = NOTE
    hdr = ["Contract", "Ticker", "Code 1", "Code 2", "IV RIC", "Entry date", "Exit date", "LSEG: before, code 1",
           "LSEG: after, code 1", "LSEG: before, code 2", "LSEG: after, code 2", "IV before (entry)", "IV after (exit)",
           "Share code 1", "Share code 2", "EPS: before, code 1", "EPS: earlier, code 1", "EPS: before, code 2", "EPS: earlier, code 2",
           "EPS before (entry)", "EPS earlier", "Earlier date", "LSEG: IV earlier", "IV earlier", "SmartEst: code 1", "SmartEst: code 2",
           "SmartEstimate before (entry)", "LSEG: IV call", "IV call before (entry)"]
    for j, h in enumerate(hdr):
        rq.cell(row=5, column=1 + j, value=h).font = BOLD
    for k, (slug, t, d0, d1, de) in enumerate(rows):
        r = 6 + k
        base, _, cls = t.partition(".")
        ivr = base + cls + "ATMIV.U"
        first = to_ric(t, exch.get(t) or "Nasdaq")
        second = to_ric(t, "NYSE") if first.endswith(".O") else to_ric(t, "Nasdaq")
        rq.cell(row=r, column=1, value=slug); rq.cell(row=r, column=2, value=t)
        rq.cell(row=r, column=3, value=f'=IF(Setup!$B$3="IV RIC","{ivr}","{first}")')
        rq.cell(row=r, column=4, value=f'=IF(Setup!$B$3="IV RIC","","{second}")')
        rq.cell(row=r, column=5, value=ivr)
        rq.cell(row=r, column=6, value=d0.strftime("%Y-%m-%d")); rq.cell(row=r, column=7, value=d1.strftime("%Y-%m-%d"))
        for c, code, dcol in ((8, "C", "F"), (9, "C", "G"), (10, "D", "F"), (11, "D", "G")):
            rq.cell(row=r, column=c, value=f'=_xll.RDP.Data({code}{r},Setup!$B$2,"SDate="&{dcol}{r}&" EDate="&{dcol}{r}&" Frq=D")')
        rq.cell(row=r, column=12, value=f'=IFERROR(1*H{r},IFERROR(1*J{r},""))')
        rq.cell(row=r, column=13, value=f'=IFERROR(1*I{r},IFERROR(1*K{r},""))')
        rq.cell(row=r, column=14, value=first); rq.cell(row=r, column=15, value=second)
        rq.cell(row=r, column=22, value=de.strftime("%Y-%m-%d"))
        for c, code, dcol in ((16, "N", "F"), (17, "N", "V"), (18, "O", "F"), (19, "O", "V")):
            rq.cell(row=r, column=c, value=f'=_xll.RDP.Data({code}{r},Setup!$B$4,Setup!$C$4&" SDate="&{dcol}{r}&" EDate="&{dcol}{r}&" Frq=D")')
        rq.cell(row=r, column=20, value=f'=IFERROR(1*P{r},IFERROR(1*R{r},""))')
        rq.cell(row=r, column=21, value=f'=IFERROR(1*Q{r},IFERROR(1*S{r},""))')
        rq.cell(row=r, column=23, value=f'=_xll.RDP.Data(C{r},Setup!$B$2,"SDate="&V{r}&" EDate="&V{r}&" Frq=D")')
        rq.cell(row=r, column=24, value=f'=IFERROR(1*W{r},"")')
        for c, code in ((25, "N"), (26, "O")):
            rq.cell(row=r, column=c, value=f'=_xll.RDP.Data({code}{r},Setup!$D$4,Setup!$C$4&" SDate="&F{r}&" EDate="&F{r}&" Frq=D")')
        rq.cell(row=r, column=27, value=f'=IFERROR(1*Y{r},IFERROR(1*Z{r},""))')
        rq.cell(row=r, column=28, value=f'=_xll.RDP.Data(C{r},Setup!$E$4,"SDate="&F{r}&" EDate="&F{r}&" Frq=D")')
        rq.cell(row=r, column=29, value=f'=IFERROR(1*AB{r},"")')
    from openpyxl.utils import get_column_letter as L
    for j_, w in enumerate((48, 8, 10, 10, 14, 11, 11, 13, 13, 13, 13, 14, 14, 10, 10, 12, 12, 12, 12, 12, 12, 11, 12, 12, 12, 12, 14, 12, 14)):
        rq.column_dimensions[L(1 + j_)].width = w
    rq.freeze_panes = "C6"
    wb.save(_p(LSEG_REQUEST))
    print(f"   LSEG: {len(have)} reports already have implied volatility. Request file for {len(rows)} reports: {_p(LSEG_REQUEST)}", flush=True)
    return len(rows)


def lseg_interactive():
    import sys
    msg = ("\n" + "=" * 78 + "\nLSEG STEP. The request file is downloading now.\n"
           "  1. Open GroupA2_LSEG_request.xlsx in Excel with LSEG Workspace signed in.\n"
           "  2. Look at the Setup sheet: cell B6 should show a number (Apple's implied volatility). If not, follow the tips there.\n"
           "  3. Wait until columns L and M of the Request sheet show numbers (a few minutes), then save the file.\n"
           "  4. Upload the saved file with the button below. Click Cancel to skip LSEG for this run.\n" + "=" * 78)
    try:
        from google.colab import files
        files.download(_p(LSEG_REQUEST))
        print(msg, flush=True)
        up = files.upload()
        if not up:
            print("   LSEG skipped for this run.", flush=True)
            return
        name, data = next(iter(up.items()))
        with open(_p(LSEG_REQUEST), "wb") as fh:
            fh.write(data)
        print(f"   received {name}", flush=True)
    except ImportError:
        if not sys.stdin or not sys.stdin.isatty():
            return
        print(msg.replace("downloading now", f"saved as {_p(LSEG_REQUEST)}"), flush=True)
        path = input("Path to the filled file (Enter to skip): ").strip()
        if not path:
            return
        import shutil
        shutil.copy(path, _p(LSEG_REQUEST))
    lseg_collect()


LSEG_COLS = ["slug", "ticker", "ric", "entry_date", "exit_date", "iv_entry", "iv_exit", "earlier_date", "eps_entry", "eps_earlier",
             "iv_earlier", "smart_est", "iv_call"]


def load_lseg():
    cols = LSEG_COLS
    if not os.path.exists(_p(F_LSEG)):
        return pd.DataFrame(columns=cols)
    return pd.read_csv(_p(F_LSEG)).reindex(columns=cols).drop_duplicates("slug", keep="last").reset_index(drop=True)


def load_opthist():
    if not os.path.exists(_p(F_OPTHIST)):
        return pd.DataFrame(columns=OH_COLS)
    oh = pd.read_csv(_p(F_OPTHIST)).reindex(columns=OH_COLS).drop_duplicates("slug", keep="last")
    return oh[oh["status"] == "ok"].sort_values(["report_date", "ticker"]).reset_index(drop=True)


def load_edgar():
    if not os.path.exists(_p(F_EDGAR)):
        return pd.DataFrame(columns=["ticker", "report_date", "cik", "accession", "filing_date", "acceptance"])
    ed = pd.read_csv(_p(F_EDGAR))
    return ed[ed["acceptance"].notna()].sort_values(["report_date", "ticker"]).reset_index(drop=True)


# 5. Python copy of the Events sheet: used for row pointers, models and bootstrap draws
def load_data():
    mk = pd.read_csv(_p(F_MARKETS))
    mk["report_date"] = pd.to_datetime(mk["report_date"])
    for c in ("closed", "resolved"):
        mk[c] = mk[c].astype(str).str.lower().eq("true")
    od = pd.read_csv(_p(F_ODDS))
    px = pd.read_csv(_p(F_PRICES), index_col=0, parse_dates=True).sort_index()
    px.index = pd.to_datetime(px.index).normalize()
    est = pd.read_csv(_p(F_EPS), parse_dates=["date"]).sort_values(["ticker", "date"]).reset_index(drop=True)
    snap = pd.read_csv(_p(F_SNAP)) if os.path.exists(_p(F_SNAP)) else pd.DataFrame(columns=SNAP_COLS)
    snap = snap.reindex(columns=SNAP_COLS)
    snap["snap_ts"] = pd.to_datetime(snap["snap_utc"])
    snap = snap.sort_values(["snap_ts", "slug"]).reset_index(drop=True)
    km = pd.read_csv(_p(F_KM)) if os.path.exists(_p(F_KM)) else pd.DataFrame()
    ko = pd.read_csv(_p(F_KODDS)) if os.path.exists(_p(F_KODDS)) else pd.DataFrame(columns=["market", "ts", "p"])
    return mk, od, px, est, snap, km, ko


def build_events(mk, od, px, est):
    m = mk[mk["resolved"] & mk["ticker"].notna() & (mk["volume"] >= MIN_VOLUME) & (mk["report_date"] >= pd.Timestamp(START_DATE))]
    m = m[m["slug"].isin(set(od["slug"])) & m["ticker"].map(yf_symbol).isin(set(px.columns))]
    m = m.sort_values("volume", ascending=False).drop_duplicates(["ticker", "report_date"])
    od = od.sort_values(["slug", "ts"])
    odds = {s: (g["ts"].to_numpy(), g["p"].to_numpy()) for s, g in od.groupby("slug")}

    def p_at(slug, ts):
        t, p = odds[slug]
        i = np.searchsorted(t, ts, side="right") - 1
        return float(p[i]) if i >= 0 else np.nan

    dates, lpx = px.index, np.log(px)
    dlog = lpx.diff()
    ab = dlog.sub(dlog[MARKET], axis=0)
    est = est.copy()
    est["beat"] = (est["reported"] > est["estimate"]).astype(float)

    def report_move(sym, d):
        a = dates.searchsorted(d - pd.Timedelta(days=1), side="right") - 1
        b = dates.searchsorted(d, side="right")
        if a < 0 or b >= len(dates):
            return np.nan
        return (lpx[sym].iloc[b] - lpx[sym].iloc[a]) - (lpx[MARKET].iloc[b] - lpx[MARKET].iloc[a])

    est["abs_move"] = [abs(report_move(yf_symbol(t), d)) if yf_symbol(t) in px.columns else np.nan
                       for t, d in zip(est["ticker"], est["date"])]
    eg = {t: g for t, g in est.groupby("ticker")}
    rows = []
    for x in m.itertuples():
        sym, rd, c = yf_symbol(x.ticker), x.report_date, cut_ts(x.report_date)
        r = dict(slug=x.slug, ticker=x.ticker, report_date=rd, eps_type=x.eps_type, volume=x.volume, y=int(x.outcome),
                 p=p_at(x.slug, c), p_7d=p_at(x.slug, c - LEAD_DAYS * 86400))
        tt, pp = odds[x.slug]
        w7 = pp[(tt <= c) & (tt > c - 7 * 86400)]
        r["odds_vol"] = float(np.std(np.diff(w7))) if len(w7) >= 12 else np.nan
        r["odds_range"] = float(w7.max() - w7.min()) if len(w7) >= 12 else np.nan
        r["late24"] = r["p"] - p_at(x.slug, c - 86400)
        i0 = dates.searchsorted(rd - pd.Timedelta(days=1), side="right") - 1
        i1 = dates.searchsorted(rd, side="right")
        ok = i0 >= 0 and i1 < len(dates)
        r["ar"] = report_move(sym, rd)
        r["raw_ret"] = px[sym].iloc[i1] / px[sym].iloc[i0] - 1 if ok else np.nan
        r["runup10"] = ((lpx[sym].iloc[i0] - lpx[sym].iloc[i0 - RUNUP_DAYS]) - (lpx[MARKET].iloc[i0] - lpx[MARKET].iloc[i0 - RUNUP_DAYS])
                        if ok and i0 >= RUNUP_DAYS else np.nan)
        w = ab[sym].iloc[max(0, i0 - VOL_DAYS + 1):i0 + 1].dropna() if ok else pd.Series(dtype=float)
        r["vol20"] = float(w.std()) if len(w) >= VOL_DAYS * 0.75 else np.nan
        g = eg.get(x.ticker)
        past = g[g["date"] < rd - pd.Timedelta(days=3)] if g is not None else None
        r["company_rate"] = past["beat"].mean() if past is not None and len(past) >= MIN_HISTORY else np.nan
        mv = past["abs_move"].dropna() if past is not None else []
        r["past_abs_move"] = mv.mean() if len(mv) >= MIN_PAST_MOVES else np.nan
        rows.append(r)
    ev = pd.DataFrame(rows).sort_values(["report_date", "ticker"]).reset_index(drop=True)
    ev["dp7"] = ev["p"] - ev["p_7d"]
    ev["confident"] = (ev["p"] >= HIGH_P).astype(float).where(ev["p"].notna())
    ev["sample"] = np.where(ev["report_date"] > pd.Timestamp(FREEZE_DATE), "New", "Part 1")
    ev["week"] = ev["report_date"].dt.to_period("W").astype(str)
    ev["month"] = ev["report_date"].dt.to_period("M").astype(str)
    ev["fall5"] = (ev["ar"] <= FALL).astype(float).where(ev["ar"].notna())
    ev["big10"] = (ev["ar"] <= BIG_FALL).astype(float).where(ev["ar"].notna())
    ev["abs_ar"] = ev["ar"].abs()
    return ev


def snapshot_pointer(ev, snap):
    ptr = []
    by = {k: g for k, g in snap.groupby("slug")}
    for r in ev.itertuples():
        g = by.get(r.slug)
        lim = pd.Timestamp(r.report_date) + pd.Timedelta(hours=CUTOFF_HOUR_UTC)
        g = g[g["snap_ts"] <= lim] if g is not None else None
        ptr.append(int(g.index[-1]) + 1 if g is not None and len(g) else None)
    return ptr


# 6. Python-only steps: walk-forward models and bootstrap draws
def walk_forward(ev, features, target, kind, model, min_train=300, flags=None, aux=None):
    from sklearn.ensemble import HistGradientBoostingClassifier, HistGradientBoostingRegressor
    from sklearn.impute import SimpleImputer
    from sklearn.linear_model import LinearRegression, LogisticRegression
    from sklearn.pipeline import make_pipeline
    from sklearn.preprocessing import StandardScaler
    out = pd.Series(np.nan, index=ev.index)
    d = ev[ev[target].notna() & ev["p"].notna()]
    for mth in sorted(d["month"].unique()):
        tr, te = d[d["month"] < mth], ev[(ev["month"] == mth) & ev["p"].notna()]
        if len(tr) < min_train or te.empty or (kind == "clf" and tr[target].nunique() < 2):
            continue
        if model == "hist":
            out.loc[te.index] = np.quantile(tr[target], Q_LOW)
            continue
        if model == "scaled":
            f0 = features[0]
            ok = tr[f0] > 0
            out.loc[te.index] = np.quantile(tr.loc[ok, target] / tr.loc[ok, f0], Q_LOW) * te[f0]
            continue
        base_tr = base_te = None
        if model == "resid":
            tr, te = tr[tr["past_abs_move"] > 0], te[te["past_abs_move"] > 0]
            if len(tr) < min_train or te.empty:
                continue
            base_tr, base_te = tr["past_abs_move"].to_numpy(float), te["past_abs_move"].to_numpy(float)
        early = tr.sort_values("report_date", kind="stable").iloc[:int(len(tr) * 0.8)]
        fs_ = [f for f in features if early[f].nunique() > 1]
        X, Xt, yv = tr[fs_].to_numpy(float), te[fs_].to_numpy(float), tr[target].to_numpy(float)
        if base_tr is not None:
            yv = yv / base_tr
        if kind == "q10":
            est = HistGradientBoostingRegressor(loss="quantile", quantile=Q_LOW, max_depth=3, learning_rate=0.05, max_iter=200,
                                                min_samples_leaf=30, random_state=SEED)
        elif model == "linear":
            est = make_pipeline(SimpleImputer(strategy="median", add_indicator=True), StandardScaler(),
                                LogisticRegression(max_iter=2000) if kind == "clf" else LinearRegression())
        elif kind == "clf":
            est = HistGradientBoostingClassifier(max_depth=3, learning_rate=0.05, max_iter=200, min_samples_leaf=30, random_state=SEED)
        else:
            est = HistGradientBoostingRegressor(max_depth=3, learning_rate=0.05, max_iter=200, min_samples_leaf=30,
                                                loss="absolute_error", random_state=SEED)
        shift = 0.0
        if kind == "q10":
            from sklearn.base import clone
            k = int(len(tr) * 0.8)
            order = np.argsort(tr["report_date"].to_numpy(), kind="stable")
            e_c = clone(est).fit(X[order[:k]], yv[order[:k]])
            shift = float(np.quantile(yv[order[k:]] - e_c.predict(X[order[k:]]), Q_LOW))
        est.fit(X, yv)
        if kind == "clf":
            pred, fit = est.predict_proba(Xt)[:, 1], est.predict_proba(X)[:, 1]
        else:
            pred, fit = est.predict(Xt) + shift, est.predict(X) + shift
            if base_tr is not None:
                pred, fit = pred * base_te, fit * base_tr
        out.loc[te.index] = pred
        share = (tr["p"] >= HIGH_P).mean()
        if flags is not None:
            if kind == "clf":
                flags.loc[te.index] = (pred >= np.quantile(fit, 1 - share)).astype(float)
            else:
                flags.loc[te.index] = (pred <= np.quantile(fit, share)).astype(float)
        if aux is not None:
            aux[mth] = dict(share=share, med=float(np.median(fit)))
    return out


MODEL_SPECS = [
    ("beat_crowd_mgr_lin", "P(beat): crowd + manager's data, logistic", "y", "clf", "linear", "crowd+mgr"),
    ("beat_crowd_mgr_gb", "P(beat): crowd + manager's data, boosted trees", "y", "clf", "gb", "crowd+mgr"),
    ("beat_mgr_lin", "P(beat): manager's data only, logistic", "y", "clf", "linear", "mgr"),
    ("fall_crowd_lin", "P(fall 5%+): crowd p alone, logistic (benchmark)", "fall5", "clf", "linear", "crowd"),
    ("fall_crowd_mgr_lin", "P(fall 5%+): crowd + manager's data, logistic", "fall5", "clf", "linear", "crowd+mgr"),
    ("move_mgr_lin", "|Move|: manager's data only, linear", "abs_ar", "reg", "linear", "mgr"),
    ("move_crowd_mgr_gb", "|Move|: crowd + manager's data, boosted trees", "abs_ar", "reg", "gb", "crowd+mgr"),
    ("beat_sm_lin", "P(beat): StarMine predicted surprise + analyst revision, logistic", "y", "clf", "linear", "starmine"),
    ("beat_crowd_sm_lin", "P(beat): crowd + StarMine, logistic", "y", "clf", "linear", "crowd+sm"),
    ("q10_hist", "10% worst case: 10th percentile of all earlier reports", "ar", "q10", "hist", "none"),
    ("q10_past_scaled", "10% worst case: spreadsheet, k x company's average past |move| (k from earlier months)", "ar", "q10", "scaled", "pastmove"),
    ("q10_crowd_mgr_gb", "10% worst case: AI downside model (crowd + manager's data + LSEG), boosted trees", "ar", "q10", "gb", "crowd+mgr"),
    ("q10_notext_gb", "10% worst case: the same AI model without the release read by generative AI (ablation)", "ar", "q10", "gb", "crowd+mgr-text"),
    ("beat_crowd_text_lin", "P(beat): crowd + last quarter's release read by generative AI, logistic", "y", "clf", "linear", "crowd+text"),
    ("beat_crowd_kw_lin", "P(beat): crowd + keyword count of the same release (the spreadsheet way), logistic", "y", "clf", "linear", "crowd+kw"),
    ("fall10_gb", "P(fall 10%+): crowd + manager's data, boosted trees (pre-registered 1, trained on the decision)", "big10", "clf", "gb", "crowd+mgr"),
    ("q10_resid_gb", "10% worst case: AI learns k for each report on top of the spreadsheet (pre-registered 2)", "ar", "q10", "resid", "crowd+mgr"),
    ("q10_kalshi_gb", "10% worst case: AI downside model + Kalshi KPI odds (reports with a Kalshi KPI contract only)", "ar", "q10", "gb", "crowd+mgr+kalshi"),
]
BENCH = {"beat_crowd_sm_lin": "beat_sm_lin", "q10_crowd_mgr_gb": "q10_hist", "q10_past_scaled": "q10_hist",
         "q10_resid_gb": "q10_past_scaled", "q10_kalshi_gb": "q10_crowd_mgr_gb"}
POLICY_COLS = {"q10_crowd_mgr_gb": "ai_trim", "fall10_gb": "fall10_trim", "q10_resid_gb": "resid_trim"}
NEW_POLICIES = [("fall10_trim", "Fall-classifier trims (1/0)", "Outcome: fall-classifier policy", False),
                ("resid_trim", "Residual-AI trims (1/0)", "Outcome: residual-AI policy", False),
                ("size_t", "Sized trim (share of the position)", "Outcome: sized AI trims", True)]
BT_SUB = ["Outcome: hold", "Outcome: rule trims", "Outcome: AI trims", "Outcome: past-move rule", "Outcome: implied-move rule", "Outcome: crowd-warning rule",
          "Outcome: fall-classifier policy", "Outcome: residual-AI policy", "Outcome: sized AI trims"]
BT_SETS = ["no 3 past reports", "new reports (holdout)"]
CA_COLS = ["Crowd lost confidence (1/0)", "StarMine expects a miss, confident crowd (1/0)", "Analysts cut estimates, confident crowd (1/0)",
           "Crowd warns (1/0)", "AR | crowd lost confidence", "AR | crowd steady or more sure", "AR | confident, StarMine expects a miss",
           "AR | confident, StarMine agrees", "AR | confident, analysts cut", "AR | confident, analysts up or flat",
           "Crowd-warning rule trims (1/0)", "Outcome: crowd-warning rule"]
SHEET_RULES = [("past_trim", "past_abs_move", "Spreadsheet: trim the biggest past moves (company's average past |move|)"),
               ("iv_trim", "iv_event_pre", "Spreadsheet: trim the biggest options-implied moves (LSEG, before the report)"),
               ("el_trim", "exp_loss", "Spreadsheet: trim the biggest expected loss = crowd P(miss) x average past |move|")]
LSEG_FEATS = ["rev_lseg", "iv_b", "iv_ratio", "iv_event_pre", "smart_surp", "skew"]
NEW_FEATS = ["odds_vol", "odds_range", "late24", "peer_mean", "peer_fall"]
TEXT_FEATS = ["guid_score", "tone", "risks"]


def run_models(ev):
    ev = ev.copy()
    ev["gaap"] = (ev["eps_type"] == "GAAP").astype(float)
    ev["log_volume"] = np.log1p(ev["volume"])
    ev["unc"] = ev["p"] * (1 - ev["p"])
    have = [f for f in LSEG_FEATS + NEW_FEATS + TEXT_FEATS + ["kw_net"]
            if f in ev and ev[f].notna().sum() >= 300 and ev[f].nunique() > 1]
    txt = [f for f in TEXT_FEATS if f in have]
    feats = {"mgr": ["company_rate", "runup10", "vol20", "past_abs_move", "gaap", "log_volume"] + [f for f in have if f != "kw_net"],
             "crowd": ["p", "unc"], "none": [], "pastmove": ["past_abs_move"], "starmine": [f for f in ("smart_surp", "rev_lseg") if f in have]}
    feats["crowd+mgr"] = ["p", "dp7", "unc"] + feats["mgr"]
    feats["crowd+sm"] = ["p", "unc"] + feats["starmine"]
    feats["crowd+mgr-text"] = [f for f in feats["crowd+mgr"] if f not in TEXT_FEATS]
    feats["crowd+text"] = ["p", "unc"] + txt
    feats["crowd+kw"] = ["p", "unc"] + [f for f in ("kw_net",) if f in have]
    feats["crowd+mgr+kalshi"] = feats["crowd+mgr"] + ["k_kpi_p", "k_kpi_n"]
    n_k = int(ev["k_kpi_p"].notna().sum()) if "k_kpi_p" in ev else 0
    out = pd.DataFrame({"slug": ev["slug"]})
    aux_main = {}
    for col, _, tgt, kind, model, fs in MODEL_SPECS:
        if (fs in ("starmine", "crowd+sm") and "smart_surp" not in have) or (fs == "crowd+text" and not txt) or \
                (fs == "crowd+kw" and "kw_net" not in have) or (fs == "crowd+mgr-text" and not txt) or \
                (fs == "crowd+mgr+kalshi" and n_k < 30):
            out[col] = np.nan
            if col in POLICY_COLS:
                out[POLICY_COLS[col]] = np.nan
            print(f"   {col}: not enough data yet" + (f" ({n_k} reports with a Kalshi KPI contract)" if fs == "crowd+mgr+kalshi" else ""), flush=True)
            continue
        fl = pd.Series(np.nan, index=ev.index) if col in POLICY_COLS else None
        ax = aux_main if col == "q10_crowd_mgr_gb" else None
        out[col] = walk_forward(ev, feats[fs], tgt, kind, model, flags=fl, aux=ax)
        if fl is not None:
            out[POLICY_COLS[col]] = fl
        if fs == "crowd+mgr+kalshi":
            out[col] = out[col].where(ev["k_kpi_p"].notna())
        print(f"   {col}: {out[col].notna().sum()} out-of-sample forecasts", flush=True)
    miss = out["resid_trim"].isna() & ~(ev["past_abs_move"] > 0) & out["ai_trim"].notna()
    out.loc[miss, "resid_trim"] = out.loc[miss, "ai_trim"]
    out["size_t"] = size_policy(ev, out["q10_crowd_mgr_gb"], aux_main)
    ev = ev.assign(exp_loss=(1 - ev["p"]) * ev["past_abs_move"])
    for k, feat, _ in SHEET_RULES:
        out[k] = sheet_rule_flags(ev, feat)
    warn = ((ev["dp7"] <= -DRIFT) | ((ev["p"] >= HIGH_P) & (ev.get("smart_surp", pd.Series(np.nan, index=ev.index)) < 0))).astype(float)
    out["cw_trim"] = sheet_rule_flags(ev.assign(crowd_past=ev["past_abs_move"] * (1 + CROWD_WARN_WEIGHT * warn)), "crowd_past")
    print(f"   Crowd-warning screen trims {out['cw_trim'][out['ai_trim'].notna()].mean():.0%} (crowd warns on {warn.mean():.0%} of reports)", flush=True)
    print("   Spreadsheet rules: " + ", ".join(f"{k} trims {out[k][out['ai_trim'].notna()].mean():.0%}" for k, *_ in SHEET_RULES), flush=True)
    print(f"   AI policy: {int(out['ai_trim'].notna().sum())} reports in the backtest, trims {out['ai_trim'].mean():.0%}. "
          f"Pre-registered: fall classifier trims {out['fall10_trim'].mean():.0%}, residual AI {out['resid_trim'].mean():.0%}, "
          f"sized trims sell {out['size_t'].mean():.0%} of a position on average (50% rule: {TRIM * out['ai_trim'].mean():.0%})", flush=True)
    return out


def _alloc(s, budget, cap, w=None):
    s = np.asarray(s, float)
    w = np.ones(len(s)) if w is None else np.asarray(w, float)
    if (s * w).sum() <= 0 or budget <= 0:
        return np.zeros(len(s))
    if (cap * w * (s > 0)).sum() <= budget:
        return np.where(s > 0, cap, 0.0)
    lo, hi = 0.0, 1.0
    while (w * np.minimum(cap, hi * s)).sum() < budget:
        hi *= 2
    for _ in range(60):
        mid = (lo + hi) / 2
        lo, hi = (mid, hi) if (w * np.minimum(cap, mid * s)).sum() < budget else (lo, mid)
    return np.minimum(cap, hi * s)


def size_policy(ev, pred, aux):
    t = pd.Series(np.nan, index=ev.index)
    for m, a in aux.items():
        idx = ev.index[(ev["month"] == m) & pred.notna()]
        if len(idx):
            s = np.maximum(0.0, a["med"] - pred.loc[idx].to_numpy(float))
            t.loc[idx] = _alloc(s, TRIM * a["share"] * len(idx), SIZE_CAP)
    return t


def add_kalshi_features(ev, km, ko):
    ev = ev.drop(columns=[c for c in ("k_kpi_p", "k_kpi_n") if c in ev])
    if not len(km) or not len(ko) or "cut" not in km:
        return ev.assign(k_kpi_p=np.nan, k_kpi_n=0.0)
    k = km.copy()
    k["report_date"] = pd.to_datetime(k["report_date"])
    oz = {m: (g["ts"].to_numpy(), g["p"].to_numpy()) for m, g in ko.sort_values("ts").groupby("market")}

    def p_at(m, c):
        if m not in oz:
            return np.nan
        t, p = oz[m]
        i = np.searchsorted(t, c, side="right") - 1
        return float(p[i]) if i >= 0 else np.nan
    k["kp"] = [p_at(m, c) for m, c in zip(k["mticker"], k["cut"])]
    by = {t: g for t, g in k.dropna(subset=["kp"]).groupby("ticker")}
    kp, kn = [], []
    for t, d in zip(ev["ticker"], pd.to_datetime(ev["report_date"])):
        g = by.get(t)
        x = g[((g["report_date"] - d).abs() <= pd.Timedelta(days=KALSHI_FEAT_DAYS)) & (g["cut"] <= cut_ts(d))] if g is not None else []
        kp.append(float(x["kp"].mean()) if len(x) else np.nan)
        kn.append(float(len(x)))
    print(f"   Kalshi KPI odds as a model input: {int(np.isfinite(kp).sum())} Polymarket reports have one", flush=True)
    return ev.assign(k_kpi_p=kp, k_kpi_n=kn)


def sheet_rule_flags(ev, feat):
    d = ev[ev["p"].notna()]
    flag = pd.Series(np.nan, index=ev.index)
    for m in sorted(d["month"].unique()):
        tr = d[(d["month"] < m) & d[feat].notna()]
        if len(tr) < 300:
            continue
        thr = np.quantile(tr[feat], 1 - (tr["p"] >= HIGH_P).mean())
        te = d[d["month"] == m]
        flag.loc[te.index] = (te[feat] >= thr).astype(float)
    return flag


def add_lseg_features(ev, ls):
    l = ls.copy()
    for c in ("iv_entry", "iv_exit", "iv_earlier", "eps_entry", "eps_earlier", "smart_est"):
        l[c] = pd.to_numeric(l[c], errors="coerce")
    dec = lambda x: x.where(~(x > 3), x / 100)
    l["iv_b"], l["iv_a"], l["iv_e"] = dec(l["iv_entry"]), dec(l["iv_exit"]), dec(l["iv_earlier"])
    l["rev_lseg"] = (l["eps_entry"] - l["eps_earlier"]) / l["eps_earlier"].abs()
    l["iv_ratio"] = l["iv_b"] / l["iv_e"]
    l["iv_event_pre"] = np.sqrt(np.maximum(0, l["iv_b"] ** 2 - l["iv_e"] ** 2) * IV_DAYS / 365)
    l["smart_surp"] = (l["smart_est"] - l["eps_entry"]) / l["eps_entry"].abs()
    l["iv_call"] = pd.to_numeric(l.get("iv_call"), errors="coerce") if "iv_call" in l else np.nan
    l["skew"] = l["iv_b"] - dec(l["iv_call"]) if "iv_call" in l else np.nan
    l["lseg_put_ok"] = ((l["iv_b"] > 0) & (l["iv_a"] > 0)).astype(float)
    cols = ["rev_lseg", "iv_b", "iv_a", "iv_ratio", "iv_event_pre", "smart_surp", "skew", "lseg_put_ok"]
    l[cols] = l[cols].replace([np.inf, -np.inf], np.nan)
    ev = ev.drop(columns=[c for c in cols if c in ev]).merge(l[["slug"] + cols], on="slug", how="left")
    ev["lseg_put_ok"] = ev["lseg_put_ok"].fillna(0)
    return ev


def add_backtest(ev, models):
    ev = ev.copy()
    ev["ai_trim"] = models["ai_trim"].to_numpy()
    bt = ev["ar"].notna() & ev["p"].notna() & ev["ai_trim"].notna() & (ev["lseg_put_ok"] == 1) & ev["raw_ret"].notna()
    rule = (ev["p"] >= HIGH_P).astype(float)
    ev["out_hold"] = ev["ar"].where(bt)
    flags = [("rule", rule), ("ai", ev["ai_trim"])] + [(k, models[k].to_numpy()) for k, *_ in SHEET_RULES] + \
            [(k, models[k].to_numpy()) for k in ("fall10_trim", "resid_trim")]
    for k, f in flags:
        f = pd.Series(f, index=ev.index).fillna(0)
        ev[f"out_{k}"] = (ev["ar"] * (1 - TRIM * f) - TCOST * TRIM * f).where(bt)
    t = pd.Series(models["size_t"].to_numpy(), index=ev.index).fillna(0)
    ev["out_size"] = (ev["ar"] * (1 - t) - TCOST * t).where(bt)
    f = pd.Series(models["cw_trim"].to_numpy(), index=ev.index).fillna(0)
    ev["out_cw"] = (ev["ar"] * (1 - TRIM * f) - TCOST * TRIM * f).where(bt)
    ev["hist_ok"] = ev["past_abs_move"].notna()
    return ev


GEMINI_MODEL = "gemini-2.5-flash"
GEMINI_MODELS = ["gemini-2.5-flash", "gemini-flash-latest", "gemini-2.5-flash-lite", "gemini-flash-lite-latest", "gemini-2.0-flash",
                 "gemini-2.0-flash-lite"]
_GEM_OK = {}
LIVE_FEATS = ["p", "unc", "company_rate", "runup10", "vol20", "past_abs_move", "peer_mean", "peer_fall", "guid_score", "tone", "risks"]


_GEM_DEAD = set()


def _gem_post(body, timeout=90):
    global GEMINI_MODEL
    order = ([_GEM_OK["m"]] if "m" in _GEM_OK else []) + [m for m in GEMINI_MODELS if m != _GEM_OK.get("m")]
    for m in order:
        if m in _GEM_DEAD:
            continue
        for i in range(4):
            try:
                r = _S.post(f"https://generativelanguage.googleapis.com/v1beta/models/{m}:generateContent",
                            headers={"x-goog-api-key": GEMINI_API_KEY.strip(), "Content-Type": "application/json"}, json=body, timeout=timeout)
            except Exception:
                time.sleep(5); continue
            if r.status_code == 429:
                if "perday" in r.text.lower().replace(" ", "") or i == 3:
                    _GEM_DEAD.add(m)
                    print(f"   Gemini: daily limit reached on {m}, moving to the next model", flush=True)
                    break
                time.sleep(20 * (i + 1)); continue
            if r.status_code in (400, 404):
                t_ = r.text.lower()
                if r.status_code == 404 or "not found" in t_ or "not supported" in t_ or "is not available" in t_:
                    _GEM_DEAD.add(m)
                    break
                print(f"   Gemini API refused the request (400): {r.text[:200]}", flush=True)
                return None
            if r.status_code >= 500:
                time.sleep(5); continue
            if r.status_code != 200:
                print(f"   Gemini API refused the request ({r.status_code}): {r.text[:200]}", flush=True)
                return None
            if _GEM_OK.get("m") != m:
                _GEM_OK["m"] = m; GEMINI_MODEL = m
                print(f"   Gemini API: using {m}", flush=True)
            try:
                return r.json()
            except Exception:
                return None
    return None


def genai(prompt):
    if not USE_GENAI:
        return None
    if GEMINI_API_KEY.strip():
        js = _gem_post({"contents": [{"parts": [{"text": prompt}]}]})
        try:
            return js["candidates"][0]["content"]["parts"][0]["text"]
        except Exception:
            return None
    try:
        from google.colab import ai
        return str(ai.generate_text(prompt))
    except Exception:
        return None


def _company_history(est, px):
    dates, lpx = px.index, np.log(px)
    est = est.copy()
    est["beat"] = (est["reported"] > est["estimate"]).astype(float)

    def mv(sym, d):
        a = dates.searchsorted(d - pd.Timedelta(days=1), side="right") - 1
        b = dates.searchsorted(d, side="right")
        if a < 0 or b >= len(dates) or sym not in px.columns:
            return np.nan
        return abs((lpx[sym].iloc[b] - lpx[sym].iloc[a]) - (lpx[MARKET].iloc[b] - lpx[MARKET].iloc[a]))
    est["abs_move"] = [mv(yf_symbol(t), d) for t, d in zip(est["ticker"], est["date"])]
    return {t: g for t, g in est.groupby("ticker")}


def live_ai(ev, snap, px, est, mk):
    from sklearn.ensemble import HistGradientBoostingRegressor
    today = pd.Timestamp.now(tz="UTC").tz_localize(None).normalize()
    if not len(snap):
        return
    lv = snap[pd.to_datetime(snap["report_date"]) >= today]
    if lv.empty:
        print("   Live AI: no upcoming reports in the snapshots", flush=True)
        return
    lv = lv.sort_values("snap_ts").drop_duplicates("slug", keep="last")
    lv = lv.sort_values("volume", ascending=False).drop_duplicates(["ticker", "report_date"]).copy()
    eg = _company_history(est, px)
    dates, lpx = px.index, np.log(px)
    ab = lpx.diff().sub(lpx.diff()[MARKET], axis=0)
    rows = []
    for x in lv.itertuples():
        sym, rd = yf_symbol(x.ticker), pd.Timestamp(x.report_date)
        p = x.p_mid if pd.notna(x.p_mid) else x.p_gamma
        r = dict(slug=x.slug, ticker=x.ticker, report_date=rd, snap_utc=x.snap_utc, p=p, unc=p * (1 - p) if pd.notna(p) else np.nan,
                 log_volume=np.log1p(x.volume) if pd.notna(x.volume) else np.nan, put95_cost=x.put95_cost,
                 put95_last=getattr(x, "put95_last", np.nan))
        i0 = min(dates.searchsorted(rd - pd.Timedelta(days=1), side="right") - 1, len(dates) - 1)
        ok = sym in px.columns and i0 >= RUNUP_DAYS
        r["runup10"] = ((lpx[sym].iloc[i0] - lpx[sym].iloc[i0 - RUNUP_DAYS]) - (lpx[MARKET].iloc[i0] - lpx[MARKET].iloc[i0 - RUNUP_DAYS])
                        if ok else np.nan)
        w = ab[sym].iloc[max(0, i0 - VOL_DAYS + 1):i0 + 1].dropna() if ok else pd.Series(dtype=float)
        r["vol20"] = float(w.std()) if len(w) >= VOL_DAYS * 0.75 else np.nan
        g = eg.get(x.ticker)
        past = g[g["date"] < rd - pd.Timedelta(days=3)] if g is not None else None
        r["company_rate"] = past["beat"].mean() if past is not None and len(past) >= MIN_HISTORY else np.nan
        mv = past["abs_move"].dropna() if past is not None else []
        r["past_abs_move"] = mv.mean() if len(mv) >= MIN_PAST_MOVES else np.nan
        rows.append(r)
    lv = pd.DataFrame(rows)
    lv = add_peer_features(add_text_features(lv, load_releases()), est, px)
    ev = ev.copy()
    ev["unc"], ev["log_volume"] = ev["p"] * (1 - ev["p"]), np.log1p(ev["volume"])
    for f in LIVE_FEATS:
        if f not in ev:
            ev[f] = np.nan
    tr = ev[ev["ar"].notna() & ev["p"].notna()]
    kw = dict(max_depth=3, learning_rate=0.05, max_iter=200, min_samples_leaf=30, random_state=SEED)
    mq = HistGradientBoostingRegressor(loss="quantile", quantile=Q_LOW, **kw).fit(tr[LIVE_FEATS].to_numpy(float), tr["ar"])
    mm = HistGradientBoostingRegressor(loss="absolute_error", **kw).fit(tr[LIVE_FEATS].to_numpy(float), tr["abs_ar"])
    lv["ai_q10"] = mq.predict(lv[LIVE_FEATS].to_numpy(float))
    med = tr[LIVE_FEATS].median()
    names = {"p": "the crowd's probability of a beat", "unc": "how unsure the crowd is", "company_rate": "the company's record of beats",
             "runup10": "the share's run-up over the last 10 days", "vol20": "recent volatility",
             "past_abs_move": "the size of its past earnings moves", "peer_mean": "how same-sector peers reacted this season",
             "peer_fall": "the share of peers that fell 5% or more", "guid_score": "last quarter's guidance (1 raised, -1 cut)",
             "tone": "the tone of last quarter's release (-2 to 2)", "risks": "the risks management flagged last quarter (0 to 3)"}
    Xl = lv[LIVE_FEATS].to_numpy(float)
    eff = {}
    for j, f in enumerate(LIVE_FEATS):
        X2 = Xl.copy(); X2[:, j] = med[f]
        eff[f] = lv["ai_q10"].to_numpy() - mq.predict(X2)
    effd = pd.DataFrame(eff, index=lv.index)
    fmt = {"p": "{:.0%}", "unc": "{:.2f}", "company_rate": "{:.0%}", "runup10": "{:+.1%}", "vol20": "{:.1%}", "past_abs_move": "{:.1%}",
           "peer_mean": "{:+.1%}", "peer_fall": "{:.0%}", "guid_score": "{:+.0f}", "tone": "{:+.0f}", "risks": "{:.0f}"}
    lv["drivers"] = [
        "; ".join(f"{names[f]} ({fmt[f].format(lv.at[i, f])} against a typical {fmt[f].format(med[f])}, {effd.at[i, f]:+.1%} on the worst case)"
                  for f in effd.loc[i].sort_values().index[:2] if pd.notna(lv.at[i, f]) and effd.at[i, f] < -0.002) or "no single input stands out"
        for i in lv.index]
    lv["ai_move"] = mm.predict(lv[LIVE_FEATS].to_numpy(float))
    fit_q = mq.predict(tr[LIVE_FEATS].to_numpy(float))
    lv["ai_cutoff"] = np.quantile(fit_q, (tr["p"] >= HIGH_P).mean())
    lv["ai_action"] = np.where(lv["ai_q10"] <= lv["ai_cutoff"], f"trim {TRIM:.0%}", "hold")
    lv["rule_action"] = np.where(lv["p"] >= HIGH_P, f"trim {TRIM:.0%}", "hold")
    lv = lv.sort_values(["report_date", "ticker"]).reset_index(drop=True)
    old = pd.read_csv(_p(F_AI_LIVE)) if os.path.exists(_p(F_AI_LIVE)) else pd.DataFrame(columns=["slug", "note"])
    notes = dict(zip(old["slug"], old.get("note", pd.Series(dtype=str)))) if len(old) else {}
    lv["note"] = lv["slug"].map(notes).astype(object)
    ofacts, odrafts = {}, {}
    for _, o_ in (old[old["note"].notna()].iterrows() if "note" in old else []):
        try:
            ofacts[o_["slug"]] = o_["facts"] if isinstance(o_.get("facts"), str) else note_facts(o_)
        except Exception:
            pass
        odrafts[o_["slug"]] = o_.get("note_drafts", 1)
    lv["facts"], lv["note_drafts"] = lv["slug"].map(ofacts).astype(object), lv["slug"].map(odrafts)
    soon = lv["report_date"] <= today + pd.Timedelta(days=7)
    todo = lv[lv["note"].isna() & soon].sort_values("ai_q10").head(GENAI_NOTES_N)
    made = 0
    for i, x in todo.iterrows():
        facts = note_facts(x)
        t = genai(note_prompt(facts))
        if t:
            t, d_ = " ".join(str(t).split()), 1
            g = guard_text(t, facts, x.ai_action)
            if g["unsupported"] or g["action_ok"] is False:
                t2 = genai(note_prompt(facts, g))
                if t2:
                    t, d_ = " ".join(str(t2).split()), 2
            lv.loc[i, "note"], lv.loc[i, "note_drafts"], lv.loc[i, "facts"] = t, d_, facts
            made += 1
    lv["note_model"] = np.where(lv["note"].notna(), GEMINI_MODEL if GEMINI_API_KEY else "google.colab.ai (Gemini)", "")
    lv.to_csv(_p(F_AI_LIVE), index=False)
    log_flags(lv)
    print(f"   Live AI: {len(lv)} upcoming reports scored, {(lv['ai_action'] != 'hold').sum()} flagged to trim, "
          f"{made} new client notes written by generative AI" + ("" if made or not len(todo) else " (no generative AI available)"), flush=True)


def genai_parse_check(mk, km=None):
    old = pd.read_csv(_p(F_AI_PARSE)) if os.path.exists(_p(F_AI_PARSE)) else pd.DataFrame(columns=["slug", "source"])
    if len(old) and "ai_ticker" in old:
        old = old[old["ai_ticker"].notna() | old.get("ai_threshold", pd.Series(np.nan, index=old.index)).notna()]
    done = set(old["slug"].astype(str))
    pm = mk[mk["ticker"].notna() & ~mk["slug"].astype(str).isin(done)]
    if not len(old):
        pm = pm[pm["resolved"]]
        smp = pm.sample(min(GENAI_PARSE_N, len(pm)), random_state=SEED)
    else:
        smp = pm.sort_values(["report_date", "created"], ascending=False).head(GENAI_NEW_N)
    kdone = set(old.loc[old.get("source", pd.Series(dtype=str)) == "Kalshi", "slug"].astype(str)) if len(old) else set()
    ktodo = km[~km["mticker"].astype(str).isin(kdone)] if km is not None and len(km) and "mticker" in km else pd.DataFrame()
    if not len(smp) and not len(ktodo):
        print(f"   Parser check: all {len(old)} contracts already read by generative AI, nothing new", flush=True)
        return
    if genai("Reply with the word ok.") is None:
        print("   Generative AI not available (Colab AI or GEMINI_API_KEY): parser check skipped, new contracts are read on the next run", flush=True)
        return
    print(f"   Parser check: generative AI reads {len(smp)} Polymarket contracts not read before ({PARSE_BATCH} per call)", flush=True)
    rows, texts = [], []
    for x in smp.itertuples():
        rules = str(getattr(x, "rules", "") or "")
        rules = "" if rules == "nan" else rules
        for params in ({"slug": x.slug, "closed": "true"}, {"slug": x.slug}):
            if rules:
                break
            js = _get(f"{GAMMA}/markets", params=params)
            m = js[0] if isinstance(js, list) and js else {}
            rules = " ".join(str(m.get("description") or "").split())
        if rules:
            texts.append((x.slug, f"Question: {x.question}\nRules: {rules[:1500]}"))
    for k in range(0, len(texts), PARSE_BATCH):
        chunk = texts[k:k + PARSE_BATCH]
        body = "\n\n".join(f"ID {i}:\n{tx}" for i, (_, tx) in enumerate(chunk))
        js = _genai_json("Each item is a prediction market contract about a company's quarterly earnings. For each ID reply with a JSON "
                         "list of objects, keys: id (number), ticker (stock symbol), report_date (YYYY-MM-DD, the day the earnings are "
                         "released), eps_type (GAAP, non-GAAP or not stated), threshold (the EPS number to beat, as a number). "
                         "JSON only.\n\n" + body) or []
        got = {int(j.get("id", -1)): j for j in js if isinstance(j, dict) and str(j.get("id", "")).lstrip("-").isdigit()}
        for i, (slug, _) in enumerate(chunk):
            j = got.get(i, {})
            rows.append(dict(slug=slug, ai_ticker=j.get("ticker"), ai_report_date=j.get("report_date"), ai_eps_type=j.get("eps_type"),
                             ai_threshold=_f(j.get("threshold")), model=GEMINI_MODEL if GEMINI_API_KEY else "google.colab.ai (Gemini)",
                             source="Polymarket"))
    if len(ktodo):
        ks = ktodo.sample(min(GENAI_KALSHI_N, len(ktodo)), random_state=SEED) if not kdone else ktodo.tail(min(GENAI_NEW_N, len(ktodo)))
        for x in ks.itertuples():
            prompt = ("Read this prediction market contract about a company's quarterly results. Reply with JSON only, keys: ticker "
                      "(the company's main US stock symbol), metric (the business measure, a few words), threshold (the number in the "
                      "contract written out in full as a plain number, e.g. 42.5 million = 42500000, 4% = 4), quarter (e.g. Q1 2026).\n\n"
                      f"Contract: {x.title}")
            t = genai(prompt) or ""
            mm = re.search(r"\{.*\}", t, re.S)
            try:
                j = json.loads(mm.group(0)) if mm else {}
            except Exception:
                j = {}
            rows.append(dict(slug=x.mticker, ai_ticker=j.get("ticker"), ai_report_date=None, ai_eps_type=j.get("metric"),
                             ai_threshold=_f(j.get("threshold")), model=GEMINI_MODEL if GEMINI_API_KEY else "google.colab.ai (Gemini)",
                             source="Kalshi"))
    new = pd.DataFrame(rows)
    if len(new):
        new = new[new["ai_ticker"].notna() | new["ai_threshold"].notna()]
    pd.concat([old, new], ignore_index=True).drop_duplicates("slug", keep="last").to_csv(_p(F_AI_PARSE), index=False)
    print(f"   Parser check: {len(new)} new contracts read, {len(old) + len(new)} in total", flush=True)
    print(f"   Generative AI read {len(rows)} contracts for the parser check", flush=True)


MONTHS = ("january", "february", "march", "april", "may", "june", "july", "august", "september", "october", "november", "december")
FLAG_COLS = ["slug", "ticker", "report_date", "snap_utc", "p", "ai_q10", "ai_move", "ai_cutoff", "ai_action", "rule_action",
             "put95_cost", "put95_last"]
SCORE_COLS = ["slug", "ticker", "report_date", "flag_utc", "release_utc", "before_release", "p", "ai_q10", "ai_cutoff", "ai_action",
              "rule_action", "put_price", "ar", "raw_ret", "y"]


def _genai_json(prompt, kind="list"):
    t = genai(prompt)
    if not t:
        return None
    m = re.search(r"\[.*\]" if kind == "list" else r"\{.*\}", t, re.S)
    try:
        return json.loads(m.group(0)) if m else None
    except Exception:
        return None


# ---------- 1. Live scorecard: every flag the product showed, scored once the report is out ----------
def log_flags(lv=None):
    if lv is None:
        if not os.path.exists(_p(F_AI_LIVE)):
            return
        lv = pd.read_csv(_p(F_AI_LIVE))
    if not len(lv):
        return
    lv = lv.copy()
    if "snap_utc" not in lv or lv["snap_utc"].isna().any():
        sn = pd.read_csv(_p(F_SNAP)) if os.path.exists(_p(F_SNAP)) else pd.DataFrame(columns=["slug", "snap_utc"])
        last = sn.sort_values("snap_utc").drop_duplicates("slug", keep="last").set_index("slug")["snap_utc"]
        have = lv["snap_utc"] if "snap_utc" in lv else pd.Series(np.nan, index=lv.index, dtype=object)
        lv["snap_utc"] = have.fillna(lv["slug"].map(last))
    if "ai_action" not in lv:
        lv["ai_action"] = np.where(lv["ai_q10"] <= lv["ai_cutoff"], f"trim {TRIM:.0%}", "hold")
    lv["report_date"] = pd.to_datetime(lv["report_date"]).dt.strftime("%Y-%m-%d")
    new = lv.reindex(columns=FLAG_COLS).dropna(subset=["snap_utc"])
    old = pd.read_csv(_p(F_FLAGS)) if os.path.exists(_p(F_FLAGS)) else pd.DataFrame(columns=FLAG_COLS)
    allx = pd.concat([old, new], ignore_index=True).drop_duplicates(["slug", "snap_utc"], keep="first")
    allx.sort_values(["report_date", "ticker", "snap_utc"]).to_csv(_p(F_FLAGS), index=False)
    hash_flags(allx[allx["snap_utc"].isin(set(new["snap_utc"]))])


def _release_utc(acc):
    try:
        g = pd.Timestamp(str(acc)[:19].replace("T", " "))
    except Exception:
        return None
    if pd.isna(g):
        return None
    y = g.year
    d0 = pd.Timestamp(y, 3, 8); d0 += pd.Timedelta(days=(6 - d0.weekday()) % 7)
    d1 = pd.Timestamp(y, 11, 1); d1 += pd.Timedelta(days=(6 - d1.weekday()) % 7)
    off = pd.Timedelta(hours=4 if d0 <= g < d1 else 5)
    for ny in (g - off, g - 2 * off):
        h = ny.hour + ny.minute / 60
        if 6 <= h < 9.5 or 16 <= h < 17.6:
            return ny + off
    return None


def score_flags(px, mk):
    if not os.path.exists(_p(F_FLAGS)):
        return
    fl = pd.read_csv(_p(F_FLAGS))
    fl["report_date"] = pd.to_datetime(fl["report_date"])
    fl["snap_ts"] = pd.to_datetime(fl["snap_utc"])
    today = pd.Timestamp.now(tz="UTC").tz_localize(None).normalize()
    due = fl[fl["report_date"] < today]
    if due.empty:
        print("   Live scorecard: no flagged report has passed yet", flush=True)
        return
    try:
        fetch_edgar(due[["ticker", "report_date"]].drop_duplicates())
    except Exception as e:
        print(f"   Live scorecard: EDGAR release times not updated ({e})", flush=True)
    ed = pd.read_csv(_p(F_EDGAR)) if os.path.exists(_p(F_EDGAR)) else pd.DataFrame(columns=["ticker", "report_date", "acceptance"])
    rel = {(str(t), str(d)[:10]): _release_utc(a) for t, d, a in zip(ed["ticker"], ed["report_date"], ed["acceptance"]) if pd.notna(a)}
    out = dict(zip(mk.loc[mk["resolved"], "slug"], mk.loc[mk["resolved"], "outcome"])) if "resolved" in mk else {}
    dates, lpx = px.index, np.log(px)
    rows = []
    for slug, g in due.groupby("slug"):
        rd, tk = g["report_date"].iloc[0], str(g["ticker"].iloc[0])
        ru = rel.get((tk, rd.strftime("%Y-%m-%d")))
        limit = ru if ru is not None else rd + pd.Timedelta(hours=CUTOFF_HOUR_UTC)
        ok = g[g["snap_ts"] < limit]
        x = (ok if len(ok) else g).sort_values("snap_ts").iloc[-1]
        sym, ar, raw = yf_symbol(tk), np.nan, np.nan
        a = dates.searchsorted(rd - pd.Timedelta(days=1), side="right") - 1
        b = dates.searchsorted(rd, side="right")
        if sym in px.columns and a >= 0 and b < len(dates):
            ar = (lpx[sym].iloc[b] - lpx[sym].iloc[a]) - (lpx[MARKET].iloc[b] - lpx[MARKET].iloc[a])
            raw = px[sym].iloc[b] / px[sym].iloc[a] - 1
        pp = x["put95_cost"] if pd.notna(x["put95_cost"]) else x["put95_last"]
        rows.append(dict(slug=slug, ticker=tk, report_date=rd.strftime("%Y-%m-%d"), flag_utc=x["snap_utc"],
                         release_utc=ru.strftime("%Y-%m-%d %H:%M") if ru is not None else None, before_release=int(len(ok) > 0),
                         p=x["p"], ai_q10=x["ai_q10"], ai_cutoff=x["ai_cutoff"], ai_action=x["ai_action"], rule_action=x["rule_action"],
                         put_price=pp, ar=ar if np.isfinite(ar) else np.nan, raw_ret=raw if np.isfinite(raw) else np.nan,
                         y=out.get(slug)))
    sc = pd.DataFrame(rows, columns=SCORE_COLS).sort_values(["report_date", "ticker"])
    sc.to_csv(_p(F_SCORE), index=False)
    s = sc[(sc["before_release"] == 1) & sc["ar"].notna()]
    print(f"   Live scorecard: {len(sc)} flagged reports passed, {int(sc['before_release'].sum())} flagged before the release, "
          f"{len(s)} scored so far (AI trimmed {int((s['ai_action'] != 'hold').sum())})", flush=True)


# ---------- 2. Stress test: our rule-based parser against generative AI, contract text only ----------
NAME_DROP = {"inc", "incorporated", "corp", "corporation", "co", "company", "ltd", "limited", "plc", "holdings", "holding", "group",
             "the", "sa", "nv", "ag", "se", "lp", "llc", "class", "a", "b", "com", "international", "intl"}


def _norm_name(s):
    w = re.sub(r"[^a-z0-9 ]", " ", str(s).lower().replace("&", " and ")).split()
    return " ".join(x for x in w if x not in NAME_DROP)


def _sec_names():
    if not SEC_USER_AGENT.strip():
        return {}
    try:
        js = _sget(_sec_session(), "https://www.sec.gov/files/company_tickers.json") or {}
    except Exception:
        return {}
    out = {}
    for v in js.values():
        out.setdefault(_norm_name(v["title"]), str(v["ticker"]).upper())
    return out


def rule_parse_text(text, names=None):
    tk = re.search(r"\(([A-Z][A-Z.\-]{0,6})\)", text)
    ticker = tk.group(1) if tk else None
    if not ticker and names:
        mq = re.search(r"Will (.+?) (?:beat|report|post|exceed)", text)
        ticker = names.get(_norm_name(mq.group(1))) if mq else None
    date = None
    md = re.search(r"\b(" + "|".join(MONTHS) + r")\s+(\d{1,2}),?\s+(\d{4})", text, re.I)
    if md:
        date = f"{md.group(3)}-{MONTHS.index(md.group(1).lower()) + 1:02d}-{int(md.group(2)):02d}"
    else:
        mi = re.search(r"\b(\d{4})-(\d{2})-(\d{2})\b", text)
        date = mi.group(0) if mi else None
    th = re.search(r"(-)?\$\s?(-?\d+(?:\.\d+)?)", text)
    thr = (-1 if th.group(1) else 1) * float(th.group(2)) if th else None
    eps = "non-GAAP" if re.search(r"non[- ]?gaap", text, re.I) else ("GAAP" if re.search(r"\bgaap\b", text, re.I) else "not stated")
    return dict(ticker=ticker, date=date, eps=eps, thr=thr)


def _rw_noticker(t):
    return re.sub(r"\s*\([A-Z][A-Z.\-]{0,6}\)", "", t)


def _rw_date(t):
    t = re.sub(r"\b(" + "|".join(MONTHS) + r")\s+(\d{1,2}),?\s+(\d{4})",
               lambda m: f"{int(m.group(2))}.{MONTHS.index(m.group(1).lower()) + 1}.{m.group(3)}", t, flags=re.I)
    return re.sub(r"\b(\d{4})-(\d{2})-(\d{2})\b", lambda m: f"{int(m.group(3))}.{int(m.group(2))}.{m.group(1)}", t)


def _rw_cents(t):
    return re.sub(r"(-)?\$\s?(\d+(?:\.\d+)?)", lambda m: f"{'-' if m.group(1) else ''}{round(float(m.group(2)) * 100):d} cents", t)


def _rw_adjusted(t):
    return re.sub(r"non[- ]?gaap", "adjusted", t, flags=re.I)


STRESS_VARIANTS = [("as listed", []), ("no ticker in the text", [_rw_noticker]), ("dates as day.month.year", [_rw_date]),
                   ("EPS written in cents", [_rw_cents]), ("'adjusted' instead of non-GAAP", [_rw_adjusted]),
                   ("all four changes", [_rw_noticker, _rw_date, _rw_cents, _rw_adjusted])]


def _contract_text(x):
    rules = str(getattr(x, "rules", "") or "")
    rules = "" if rules == "nan" else rules
    for params in ({"slug": x.slug, "closed": "true"}, {"slug": x.slug}):
        if rules:
            break
        try:
            js = _get(f"{GAMMA}/markets", params=params)
        except Exception:
            js = []
        m = js[0] if isinstance(js, list) and js else {}
        rules = " ".join(str(m.get("description") or "").split())
    return f"Question: {x.question}\nRules: {rules[:1500]}" if rules else None


def stress_test(mk):
    old = pd.read_csv(_p(F_STRESS)) if os.path.exists(_p(F_STRESS)) else pd.DataFrame(columns=["slug", "variant"])
    done = set(old["slug"].astype(str))
    good = mk[mk["resolved"] & mk["ticker"].notna() & mk["threshold"].notna() & (mk["eps_type"] != "not stated")]
    good = good[~good["slug"].astype(str).isin(done)]
    if (old["variant"] == "as listed").sum() < min(STRESS_N, 10):
        smp = good.sample(min(STRESS_N, len(good)), random_state=SEED + 1)
    else:
        smp = good.sort_values("report_date", ascending=False).head(STRESS_NEW_N)
    gap = mk[mk["resolved"] & (mk["ticker"].isna() | (mk["eps_type"] == "not stated")) & ~mk["slug"].astype(str).isin(done)]
    if not len(smp) and not len(gap):
        print(f"   Stress test: no new contracts ({old['slug'].nunique()} already tested)", flush=True)
        return
    if genai("Reply with the word ok.") is None:
        print("   Generative AI not available: stress test not extended, new contracts are added on the next run", flush=True)
        return
    names = _sec_names()
    items = []
    for x in smp.itertuples():
        txt = _contract_text(x)
        if txt:
            for lab, fns in STRESS_VARIANTS:
                t = txt
                for f_ in fns:
                    t = f_(t)
                items.append(dict(slug=x.slug, variant=lab, text=t, truth_ticker=x.ticker, truth_date=pd.Timestamp(x.report_date).strftime("%Y-%m-%d"),
                                  truth_eps=x.eps_type, truth_thr=x.threshold))
    for x in gap.itertuples():
        txt = _contract_text(x)
        if txt:
            items.append(dict(slug=x.slug, variant="rules could not parse it fully", text=txt, truth_ticker=None, truth_date=None,
                              truth_eps=None, truth_thr=None))
    print(f"   Stress test: {len(items)} contract texts read by our rules and by generative AI", flush=True)
    for it in items:
        r_ = rule_parse_text(it["text"], names)
        it.update(rule_ticker=r_["ticker"], rule_date=r_["date"], rule_eps=r_["eps"], rule_thr=r_["thr"])
    for k in range(0, len(items), 10):
        chunk = items[k:k + 10]
        body = "\n\n".join(f"ID {i}:\n{it['text']}" for i, it in enumerate(chunk))
        t0 = time.time()
        js = _genai_json("Each item is a prediction market contract about a company's quarterly earnings. For each ID reply with a JSON "
                         "list of objects, keys: id (number), ticker (main US stock symbol), report_date (YYYY-MM-DD, the day the earnings "
                         "are released), eps_type (GAAP, non-GAAP or not stated), threshold (the EPS number to beat, in US dollars, as a "
                         "number). JSON only.\n\n" + body) or []
        sec_ = (time.time() - t0) / len(chunk)
        got = {int(j.get("id", -1)): j for j in js if isinstance(j, dict) and str(j.get("id", "")).lstrip("-").isdigit()}
        for i, it in enumerate(chunk):
            j = got.get(i, {})
            it.update(ai_ticker=j.get("ticker"), ai_date=j.get("report_date"), ai_eps=j.get("eps_type"), ai_thr=_f(j.get("threshold")),
                      seconds=sec_, model=GEMINI_MODEL if GEMINI_API_KEY else "google.colab.ai (Gemini)")
    new = pd.DataFrame(items)
    if len(new):
        ans = new.groupby("slug")["ai_ticker"].transform(lambda s: s.notna().any()) | new.groupby("slug")["ai_thr"].transform(lambda s: s.notna().any())
        new = new[ans]
    pd.concat([old, new], ignore_index=True).drop_duplicates(["slug", "variant"], keep="last").to_csv(_p(F_STRESS), index=False)
    print(f"   Stress test: {new['slug'].nunique() if len(new) else 0} new contracts added, {len(old) + len(new)} texts in total", flush=True)


# ---------- 3. Kalshi KPI contracts mapped to a stock by generative AI ----------
def kalshi_ai_map(km):
    if not len(km) or "series" not in km:
        return {}
    old = pd.read_csv(_p(F_KMAP)) if os.path.exists(_p(F_KMAP)) else pd.DataFrame(columns=["series", "answered"])
    done = set(old.loc[pd.to_numeric(old.get("answered"), errors="coerce").fillna(0) == 1, "series"])
    titles = km.sort_values("volume", ascending=False).drop_duplicates("series").set_index("series")["title"]
    todo = [s for s in titles.index if s not in done and s not in K_OVERRIDES]
    if todo and genai("Reply with the word ok.") is not None:
        print(f"   Generative AI mapping {len(todo)} Kalshi KPI series to a stock ticker", flush=True)
        rows = []
        for k in range(0, len(todo), KMAP_BATCH):
            chunk = todo[k:k + KMAP_BATCH]
            body = "\n".join(f"ID {i}: {titles[s]}" for i, s in enumerate(chunk))
            t0 = time.time()
            js = _genai_json("Each line is a prediction market contract about numbers a company reports. For each ID reply with a JSON "
                             "list of objects, keys: id (number), company (the company's name), ticker (its main US stock symbol as on "
                             "Yahoo Finance, null if the company is private or not listed in the US). JSON only.\n\n" + body) or []
            sec_ = (time.time() - t0) / len(chunk)
            got = {int(j.get("id", -1)): j for j in js if isinstance(j, dict) and str(j.get("id", "")).lstrip("-").isdigit()}
            for i, s in enumerate(chunk):
                j = got.get(i, {})
                tk = str(j.get("ticker") or "").upper().strip().replace("-", ".")
                rows.append(dict(series=s, title=titles[s], ai_company=j.get("company"),
                                 ai_ticker=tk if re.fullmatch(r"[A-Z]{1,5}(\.[A-Z])?", tk) else None, answered=int(i in got),
                                 seconds=sec_, model=GEMINI_MODEL if GEMINI_API_KEY else "google.colab.ai (Gemini)"))
        old = pd.concat([old[~old["series"].isin([r["series"] for r in rows])], pd.DataFrame(rows)], ignore_index=True)
        old.to_csv(_p(F_KMAP), index=False)
    elif todo:
        print(f"   Generative AI not available: {len(todo)} Kalshi series not mapped yet (rules only)", flush=True)
    if "ai_ticker" not in old:
        return {}
    return {s: t for s, t in zip(old["series"], old["ai_ticker"]) if isinstance(t, str) and t}


def kalshi_map_count(km, est_all, kai):
    if not len(km):
        return
    a, b = k_resolve(km, est_all), k_resolve(km, est_all, kai)
    cnt = lambda d: d.groupby("series").size() if len(d) else pd.Series(dtype=float)
    pick = lambda d: d.drop_duplicates("series").set_index("series")["ticker"] if len(d) else pd.Series(dtype=object)
    base = km.sort_values("volume", ascending=False).drop_duplicates("series")[["series", "title"]]
    old = pd.read_csv(_p(F_KMAP)) if os.path.exists(_p(F_KMAP)) else pd.DataFrame(columns=["series"])
    keep = [c for c in ("series", "ai_company", "ai_ticker", "answered", "seconds", "model") if c in old]
    mp = base.merge(old[keep], on="series", how="left") if len(old) else base
    mp["rule_ticker"], mp["final_ticker"] = mp["series"].map(pick(a)), mp["series"].map(pick(b))
    mp["events"] = mp["series"].map(km.groupby("series").size()).fillna(0)
    mp["events_rules"], mp["events_with_ai"] = mp["series"].map(cnt(a)).fillna(0), mp["series"].map(cnt(b)).fillna(0)
    mp.to_csv(_p(F_KMAP), index=False)
    print(f"   Kalshi KPI events linked to a company report: rules {int(mp['events_rules'].sum())}, rules + generative AI "
          f"{int(mp['events_with_ai'].sum())} (of {int(mp['events'].sum())})", flush=True)


# ---------- 4. Guardrail: every number in an AI text must be in the facts it was given ----------
_UNITS = {w: i for i, w in enumerate("zero one two three four five six seven eight nine ten eleven twelve thirteen fourteen "
                                         "fifteen sixteen seventeen eighteen nineteen".split())}
_TENS = {w: 10 * (i + 2) for i, w in enumerate("twenty thirty forty fifty sixty seventy eighty ninety".split())}
_SCALE = {"hundred": 100, "thousand": 1e3, "million": 1e6, "billion": 1e9}
_ORD = re.compile(r"(first|second|third|fourth|fifth|sixth|seventh|eighth|ninth|tenth|eleventh|twelfth|teenth|tieth)\b", re.I)
_NW = "|".join(list(_UNITS) + list(_TENS) + list(_SCALE))
_WORDNUM = re.compile(rf"\b(?:{_NW})(?:[\s-]+(?:{_NW}|and(?=[\s-]+(?:{_NW}))|point))*\b", re.I)
_MON = (r"(?:january|february|march|april|may|june|july|august|september|october|november|december|jan|feb|mar|apr|jun|jul|"
        r"aug|sept|sep|oct|nov|dec)\b\.?")
_DATES = re.compile(rf"\b\d{{1,2}}(?:st|nd|rd|th)?\s+(?:of\s+)?{_MON},?(?:\s+\d{{4}})?|\b{_MON}\s+\d{{1,2}}(?!\d)(?:st|nd|rd|th)?,?(?:\s+\d{{4}})?"
                    r"|\b\d{4}-\d{2}-\d{2}\b|\b\d{1,2}[/.]\d{1,2}[/.]\d{2,4}\b", re.I)


def _w2n(s):
    w = [x for x in re.split(r"[\s-]+", s.lower()) if x and x != "and"]
    dec = ""
    if "point" in w:
        i = w.index("point")
        dec = "".join(str(_UNITS[x]) for x in w[i + 1:] if x in _UNITS and _UNITS[x] < 10)
        w = w[:i]
    tot, cur = 0, 0
    for x in w:
        if x in _UNITS:
            cur += _UNITS[x]
        elif x in _TENS:
            cur += _TENS[x]
        elif x == "hundred":
            cur = (cur or 1) * 100
        elif x in _SCALE:
            tot += (cur or 1) * _SCALE[x]; cur = 0
    v = tot + cur
    return f"{v:g}.{dec}" if dec else f"{v:g}"


def _numbers(text):
    t = str(text).replace("−", "-")

    def sub(m):
        nxt = t[m.end():m.end() + 14].strip(" -")
        if _ORD.match(nxt):
            return " "
        v = _w2n(m.group(0))
        if "." not in v and float(v) <= 10 and not re.match(r"(%|per ?cent|point)", nxt, re.I):
            return " "
        return f" {v} "
    t = _WORDNUM.sub(sub, t)
    t = _DATES.sub(" ", t)
    t = re.sub(r"\b1[- ]in[- ]10\b|\b1 in 10\b", " ", t, flags=re.I)
    t = re.sub(r"\b[QH][1-4]\b|\bFY\s?\d{2,4}\b|\b\d{1,2}-[KQ]\b|\b10-Q\b", " ", t)
    out = []
    for m in re.finditer(r"(?<![\w.])-?\d{1,3}(?:,\d{3})+(?:\.\d+)?|(?<![\w.])-?\d+(?:\.\d+)?", t):
        v = float(m.group(0).replace(",", ""))
        if v.is_integer() and 1990 <= v <= 2100:
            continue
        out.append(abs(v))
    return out


def guard_text(text, facts, action=None):
    fn = _numbers(facts)
    got = _numbers(text)
    bad = [v for v in got if not any(abs(v - f) <= GUARD_TOL + 1e-9 or abs(v - round(f)) < 1e-9 or abs(v - (100 - f)) <= GUARD_TOL + 1e-9
                                     for f in fn)]
    a_ok = None
    if isinstance(action, str) and action:
        lo = str(text).lower()
        says_trim = bool(re.search(r"(suggest|recommend|advis|propos)\w*\s+(that\s+)?(you\s+)?(to\s+)?(trim|reduc|sell|cut)|\btrim(ming)?\b", lo))
        not_flag = bool(re.search(r"(does not|doesn't|did not|not) flag|not flagged", lo))
        a_ok = (says_trim and not not_flag) if action.startswith("trim") else \
            bool(re.search(r"\bhold|no action|not flag|keep", lo)) and not re.search(r"(suggest|recommend)\w*\s+(that\s+)?(you\s+)?(trim|reduc|sell)", lo)
    return dict(n=len(got), unsupported=bad, action_ok=a_ok)


def note_facts(x):
    g = lambda k: x.get(k) if hasattr(x, "get") else getattr(x, k, None)
    pc, pl = _f(g("put95_cost")), _f(g("put95_last"))
    put_txt = (f"A 5% out-of-the-money put costs {pc:.1%} of the share price" if pc is not None else
               (f"A 5% out-of-the-money put last traded at {pl:.1%} of the share price (market closed)" if pl is not None else
                "No put price yet"))
    s = (f"{g('ticker')} reports earnings on {pd.Timestamp(g('report_date')):%d %B %Y}. Prediction markets give a {g('p'):.0%} chance "
         f"that it beats the earnings forecast. Our model's 1-in-10 worst case for the move against the market is {g('ai_q10'):.1%}. "
         f"Its expected size of the move is {g('ai_move'):.1%}. {put_txt}. The suggested action is: {g('ai_action')}. "
         f"What drives the worst case most: {g('drivers')}.")
    gd = g("guidance")
    if isinstance(gd, str) and gd:
        s += f" Last quarter's earnings release, as read by our AI: guidance {gd}. {g('summary')}."
    return s


def note_prompt(facts, g=None):
    fix = ""
    if g and g["unsupported"]:
        fix += f" A first draft used numbers that are not in the facts ({', '.join(f'{v:g}' for v in g['unsupported'])}). Do not."
    if g and g.get("action_ok") is False:
        fix += " A first draft did not state the suggested action. State it."
    return ("You are an assistant to a private banker. Write 3 short plain sentences for a client who owns this share. Facts: " + facts +
            " Explain in plain words why the model flags or does not flag this report, using the drivers, and state the suggested "
            "action. Use only these facts and copy their numbers exactly, written as digits. No advice beyond the suggested action. "
            "British spelling. No em dashes." + fix)


def guard_notes():
    rows = []
    if os.path.exists(_p(F_AI_LIVE)):
        al = pd.read_csv(_p(F_AI_LIVE))
        for _, x in al[al["note"].notna()].iterrows() if "note" in al else []:
            facts = x["facts"] if isinstance(x.get("facts"), str) else note_facts(x)
            g = guard_text(x["note"], facts, x.get("ai_action"))
            rows.append(dict(kind="client note", key=x["slug"], ticker=x["ticker"], text=x["note"], facts=facts, n_numbers=g["n"],
                             n_unsupported=len(g["unsupported"]), unsupported="; ".join(f"{v:g}" for v in g["unsupported"]),
                             action_ok=None if g["action_ok"] is None else int(g["action_ok"]), drafts=x.get("note_drafts", 1)))
    if os.path.exists(_p(F_AGENT)):
        for _, x in pd.read_csv(_p(F_AGENT)).iterrows():
            g = guard_text(x["brief"], x["facts"])
            rows.append(dict(kind="agent brief", key=x["asked_utc"], ticker=x["holdings"], text=x["brief"], facts=x["facts"],
                             n_numbers=g["n"], n_unsupported=len(g["unsupported"]), unsupported="; ".join(f"{v:g}" for v in g["unsupported"]),
                             action_ok=None, drafts=x.get("drafts", 1)))
    new = pd.DataFrame(rows)
    old = pd.read_csv(_p(F_GUARD)) if os.path.exists(_p(F_GUARD)) else pd.DataFrame()
    allx = pd.concat([old, new], ignore_index=True)
    if len(allx):
        allx = allx.drop_duplicates(["kind", "key", "text"], keep="last")
        allx.to_csv(_p(F_GUARD), index=False)
        n_, u_ = allx["n_numbers"].sum(), allx["n_unsupported"].sum()
        print(f"   Guardrail: {len(allx)} AI texts, {int(n_)} numbers checked, {int(u_)} not in the data "
              f"({u_ / max(n_, 1):.1%}), {int((allx['n_unsupported'] > 0).sum())} texts with at least one", flush=True)


# ---------- 5. Portfolio agent ----------
def _ask_plain(text, show=True):
    if not os.path.exists(_p(F_AI_LIVE)):
        print("Run STEP 2 first: the agent reads today's model output.", flush=True)
        return None
    al = pd.read_csv(_p(F_AI_LIVE))
    al["report_date"] = pd.to_datetime(al["report_date"])
    js = _genai_json("Read this message from a private banker. Reply with JSON only, keys: holdings (list of objects with keys ticker "
                     "(US stock symbol) and weight (share of the portfolio in percent, null if not given)), value (portfolio value as a "
                     "number, null if not given), currency (null if not given), question (the question asked, null if none).\n\n" + text,
                     kind="dict")
    if not isinstance(js, dict):
        hold = [dict(ticker=t, weight=_f(w) if w else None) for t, w in re.findall(r"\b([A-Z]{1,5}(?:\.[A-Z])?)\b\s*(\d+(?:\.\d+)?)?\s*%?", text)
                if t not in ("EUR", "USD", "GBP", "CHF", "I", "A")]
        js = dict(holdings=hold, value=None, currency=None, question=text)
    hold = pd.DataFrame(js.get("holdings") or [], columns=["ticker", "weight"])
    hold["ticker"] = hold["ticker"].astype(str).str.upper().str.strip().str.replace("-", ".", regex=False)
    hold["weight"] = pd.to_numeric(hold["weight"], errors="coerce")
    hold = hold.drop_duplicates("ticker")
    val, cur = _f(js.get("value")), str(js.get("currency") or "")
    nxt = al.sort_values("report_date").drop_duplicates("ticker")
    lines, fr = [], []
    for h in hold.itertuples():
        m = nxt[nxt["ticker"].astype(str).str.upper() == h.ticker]
        w = f"{h.weight:g}% of the portfolio" if pd.notna(h.weight) else "weight not given"
        if m.empty:
            lines.append(f"{h.ticker} ({w}): no earnings report in the next {SNAPSHOT_DAYS} days in our data.")
            fr.append(dict(ticker=h.ticker, weight=h.weight, ai_q10=np.nan, ai_action="no report"))
            continue
        x = m.iloc[0]
        lines.append(f"{h.ticker} ({w}): " + note_facts(x))
        fr.append(dict(ticker=h.ticker, weight=h.weight, ai_q10=x["ai_q10"], ai_action=x["ai_action"]))
    fr = pd.DataFrame(fr, columns=["ticker", "weight", "ai_q10", "ai_action"])
    port = ""
    rep = fr[fr["ai_q10"].notna() & fr["weight"].notna()]
    if len(rep):
        flag = rep[rep["ai_action"].astype(str).str.startswith("trim")]
        loss, cut = (rep["weight"] * rep["ai_q10"]).sum() / 100, (flag["weight"] * flag["ai_q10"]).sum() / 100 * TRIM
        port = (f" Portfolio: {len(rep)} of {len(fr)} holdings report in the next {SNAPSHOT_DAYS} days and {len(flag)} are flagged to "
                f"trim ({flag['weight'].sum():g}% of the portfolio). If every reporting holding had its 1-in-10 worst case, the "
                f"portfolio would lose {abs(loss):.2%}. Trimming the flagged holdings by {TRIM:.0%} brings that to {abs(loss - cut):.2%}.")
        if val:
            port += f" On a portfolio of {val:,.0f} {cur} that is {abs(loss) * val:,.0f} against {abs(loss - cut) * val:,.0f}."
    facts = " ".join(lines) + port
    q = js.get("question") or "What should the client know before these reports?"
    prompt = ("You are an assistant to a private banker. Write a client brief of at most 250 words: one short paragraph per holding "
              "that reports, one line for the holdings with no report, and a closing paragraph that answers the question. Question: "
              f"{q}\nFacts: {facts}\nUse only these facts and copy their numbers exactly, written as digits. State each suggested action. "
              "If the facts do not answer the question, say so. No advice beyond the suggested actions. British spelling. No em dashes.")
    brief, drafts = genai(prompt), 1
    g = guard_text(brief or "", facts)
    if brief and g["unsupported"]:
        b2 = genai(prompt + f" A first draft used numbers that are not in the facts ({', '.join(f'{v:g}' for v in g['unsupported'])}). Do not.")
        if b2:
            brief, drafts = b2, 2
            g = guard_text(brief, facts)
    if not brief:
        brief = "Generative AI not available. The facts the brief would be written from: " + facts
        g = guard_text(brief, facts)
    status = ("checked: every number is in the model output" if not g["unsupported"] else
              f"warning: {len(g['unsupported'])} numbers not in the model output ({', '.join(f'{v:g}' for v in g['unsupported'])})")
    rec = dict(asked_utc=pd.Timestamp.now(tz="UTC").strftime("%Y-%m-%d %H:%M:%S"), request=text,
               holdings="; ".join(f"{t}:{'' if pd.isna(w) else f'{w:g}'}" for t, w in zip(fr["ticker"], fr["weight"])),
               facts=facts, brief=str(brief).strip(), n_numbers=g["n"], n_unsupported=len(g["unsupported"]), status=status, drafts=drafts,
               model=GEMINI_MODEL if GEMINI_API_KEY else "google.colab.ai (Gemini)")
    old = pd.read_csv(_p(F_AGENT)) if os.path.exists(_p(F_AGENT)) else pd.DataFrame()
    pd.concat([old, pd.DataFrame([rec])], ignore_index=True).to_csv(_p(F_AGENT), index=False)
    with open(BRIEF_FILE, "w") as fh:
        fh.write(f"# BeatOdds client brief\n\n*{rec['asked_utc']} UTC. Request: {text}*\n\n{rec['brief']}\n\n---\n\nGuardrail: {status}.\n\n"
                 f"Model output used (facts):\n\n{facts}\n")
    if show:
        print(rec["brief"] + f"\n\n[Guardrail: {status}]", flush=True)
        try:
            from google.colab import files
            files.download(BRIEF_FILE)
        except Exception:
            pass
    return rec["brief"]


# ---------- 6. Verifiable track record: a hash of each day's flags ----------
HASH_COLS = ["snap_utc", "flags", "sha256", "logged_utc"]


def _flags_text(g):
    g = g.reindex(columns=FLAG_COLS).copy()
    for c in FLAG_COLS:
        if c in ("p", "ai_q10", "ai_move", "ai_cutoff", "put95_cost", "put95_last"):
            g[c] = [("" if not np.isfinite(v) else f"{v:.6g}") for v in pd.to_numeric(g[c], errors="coerce").astype(float)]
        else:
            g[c] = ["" if pd.isna(v) else str(v) for v in g[c]]
    return g.sort_values("slug").to_csv(index=False, lineterminator="\n")


def hash_flags(new):
    old = pd.read_csv(_p(F_HASH)) if os.path.exists(_p(F_HASH)) else pd.DataFrame(columns=HASH_COLS)
    rows, done = [], set(old["snap_utc"].astype(str))
    for s, g in new.groupby("snap_utc"):
        if str(s) in done:
            continue
        h = hashlib.sha256(_flags_text(g).encode()).hexdigest()
        rows.append(dict(snap_utc=s, flags=len(g), sha256=h, logged_utc=pd.Timestamp.now(tz="UTC").strftime("%Y-%m-%d %H:%M")))
        print(f"   Flags of {s} UTC ({len(g)} reports): SHA-256 {h}. Post this hash to prove later that the flags were not changed.", flush=True)
    if rows:
        pd.concat([old, pd.DataFrame(rows)], ignore_index=True).to_csv(_p(F_HASH), index=False)


def verify_flags():
    if not os.path.exists(_p(F_HASH)) or not os.path.exists(_p(F_FLAGS)):
        return pd.DataFrame(columns=HASH_COLS + ["ok"])
    h, fl = pd.read_csv(_p(F_HASH)), pd.read_csv(_p(F_FLAGS))
    by = {str(s): g for s, g in fl.groupby("snap_utc")}
    h["ok"] = [int(str(s) in by and hashlib.sha256(_flags_text(by[str(s)]).encode()).hexdigest() == x) for s, x in zip(h["snap_utc"], h["sha256"])]
    return h[HASH_COLS + ["ok"]]


# ---------- 7. Tool-using portfolio agent ----------
def save_hedge_stats(ev):
    d = ev[ev["raw_ret"].notna() & ev["p"].notna()].copy()
    d["pay"] = np.maximum(0, (1 - PUT_OTM) - (1 + d["raw_ret"]))
    st = {g: {("miss" if y == 0 else "beat"): float(d.loc[(d["confident"] == c) & (d["y"] == y), "pay"].mean())
              for y in (0, 1)} for g, c in (("confident", 1), ("less sure", 0))}
    with open(_p(F_HEDGE), "w") as fh:
        json.dump(st, fh)


def _pc(x, nd=1, sign=False):
    return "n/a" if x is None or not np.isfinite(x) else (f"{x:+.{nd}%}" if sign else f"{x:.{nd}%}")


def _live_table():
    al = pd.read_csv(_p(F_AI_LIVE))
    al["report_date"] = pd.to_datetime(al["report_date"])
    al["tk"] = al["ticker"].astype(str).str.upper()
    return al.sort_values("report_date").drop_duplicates("tk").set_index("tk")


def _holdings(arg):
    h = pd.DataFrame(arg or [], columns=["ticker", "weight"])
    h["ticker"] = h["ticker"].astype(str).str.upper().str.strip().str.replace("-", ".", regex=False)
    h["weight"] = pd.to_numeric(h["weight"], errors="coerce")
    return h.drop_duplicates("ticker")


def tool_lookup(tickers):
    al = _live_table()
    out = []
    for t in [str(x).upper().strip().replace("-", ".") for x in tickers or []]:
        if t not in al.index:
            out.append(dict(ticker=t, status=f"no earnings report in the next {SNAPSHOT_DAYS} days in our data"))
            continue
        x = al.loc[t]
        pc, pl = _f(x.get("put95_cost")), _f(x.get("put95_last"))
        out.append(dict(ticker=t, report_date=x["report_date"].strftime("%d %B %Y"), crowd_chance_of_beat=_pc(x["p"], 0),
                        ai_worst_case_1_in_10=_pc(x["ai_q10"], 1, True), ai_expected_size_of_move=_pc(x["ai_move"]),
                        put_5pct_otm_price=_pc(pc if pc is not None else pl, 1) + ("" if pc is not None else " (last trade)"),
                        ai_action=x["ai_action"], crowd_rule_action=x.get("rule_action"), drivers=x.get("drivers"),
                        last_quarter_guidance=x.get("guidance") if isinstance(x.get("guidance"), str) else "not read"))
    return out


def _worst(hold, trims=None):
    al = _live_table()
    tr = {str(k).upper(): v for k, v in (trims or {}).items()}
    rows, tot0, tot1 = [], 0.0, 0.0
    for h in hold.itertuples():
        if h.ticker not in al.index or not np.isfinite(h.weight):
            continue
        q = float(al.loc[h.ticker, "ai_q10"])
        c0 = h.weight / 100 * q
        c1 = c0 * (1 - tr.get(h.ticker, 0.0))
        tot0, tot1 = tot0 + c0, tot1 + c1
        rows.append(dict(ticker=h.ticker, weight=f"{h.weight:g}%", worst_case=_pc(q, 1, True), share_sold=_pc(tr.get(h.ticker, 0.0), 0),
                         portfolio_loss_before=_pc(c0, 2, True), portfolio_loss_after=_pc(c1, 2, True)))
    return rows, tot0, tot1


def tool_what_if(holdings, trims=None):
    hold = _holdings(holdings)
    tr = {str(x.get("ticker", "")).upper(): (_f(x.get("share")) or 0.0) / 100 for x in (trims or [])}
    rows, t0, t1 = _worst(hold, tr)
    return dict(holdings=rows, portfolio_1_in_10_loss_without_trims=_pc(t0, 2, True), portfolio_1_in_10_loss_with_trims=_pc(t1, 2, True),
                note="Sum of each reporting holding's 1-in-10 worst case, weighted. Holdings with no report count as 0.")


def tool_size_trims(holdings, budget_pct=None, do_not_sell=None, max_trim_pct=None):
    hold = _holdings(holdings)
    al = _live_table()
    rep = hold[hold["ticker"].isin(al.index) & hold["weight"].notna()].copy()
    if rep.empty:
        return dict(status="none of these holdings reports in the coming weeks")
    rep["q"] = [float(al.loc[t, "ai_q10"]) for t in rep["ticker"]]
    rep["flag"] = [str(al.loc[t, "ai_action"]).startswith("trim") for t in rep["ticker"]]
    ban = {str(x).upper() for x in (do_not_sell or [])}
    med = float(al["ai_q10"].median())
    s = np.where(rep["ticker"].isin(ban), 0.0, np.maximum(0.0, med - rep["q"].to_numpy()))
    budget = _f(budget_pct)
    budget = budget / 100 if budget is not None else TRIM * float((rep["weight"] * rep["flag"]).sum()) / 100
    cap = (_f(max_trim_pct) or SIZE_CAP * 100) / 100
    t = _alloc(s, budget, cap, rep["weight"].to_numpy() / 100)
    trims = {tk: v for tk, v in zip(rep["ticker"], t)}
    rows, t0, t1 = _worst(hold, trims)
    return dict(proposed=[dict(ticker=tk, sell=_pc(v, 0) + " of the holding") for tk, v in trims.items()],
                total_sold=_pc(float((rep["weight"] / 100 * t).sum()), 2) + " of the portfolio", not_sold=sorted(ban),
                portfolio_1_in_10_loss_without_trims=_pc(t0, 2, True), portfolio_1_in_10_loss_with_trims=_pc(t1, 2, True),
                rule="Budget = what the 50% rule would sell. It goes to the holdings whose worst case is furthest below this week's "
                     "typical worst case, in proportion.")


def tool_put_value(ticker):
    al = _live_table()
    t = str(ticker).upper()
    if t not in al.index:
        return dict(ticker=t, status="no report in the coming weeks")
    x = al.loc[t]
    pc, pl = _f(x.get("put95_cost")), _f(x.get("put95_last"))
    price = pc if pc is not None else pl
    st = json.load(open(_p(F_HEDGE))) if os.path.exists(_p(F_HEDGE)) else None
    if price is None or st is None:
        return dict(ticker=t, status="no put price or no history yet")
    g = st["confident" if x["p"] >= HIGH_P else "less sure"]
    exp = (1 - x["p"]) * g["miss"] + x["p"] * g["beat"]
    return dict(ticker=t, put_price=_pc(price, 2), expected_payoff_from_past_reports=_pc(exp, 2), value=_pc(exp - price, 2, True),
                reading="put looks cheap" if exp > price else "put is dear, trimming is cheaper")


def tool_track_record(ticker=None):
    if not os.path.exists(_p(F_SCORE)):
        return dict(status="no flagged report has been scored yet")
    sc = pd.read_csv(_p(F_SCORE))
    sc = sc[(sc["before_release"] == 1) & sc["ar"].notna()]
    if ticker:
        sc = sc[sc["ticker"].astype(str).str.upper() == str(ticker).upper()]
    if sc.empty:
        return dict(status="no scored flag for this request yet")
    tr = sc["ai_action"].astype(str).str.startswith("trim")
    return dict(scored=len(sc), ai_said_trim=int(tr.sum()), average_move_when_ai_said_trim=_pc(sc.loc[tr, "ar"].mean(), 1, True),
                average_move_when_ai_said_hold=_pc(sc.loc[~tr, "ar"].mean(), 1, True),
                fell_below_ai_worst_case=_pc(float((sc["ar"] < sc["ai_q10"]).mean()), 0),
                last=[dict(ticker=r_.ticker, report_date=r_.report_date, ai_action=r_.ai_action, move=_pc(r_.ar, 1, True))
                      for r_ in sc.tail(5).itertuples()])


_HOLD = {"type": "ARRAY", "items": {"type": "OBJECT", "properties": {"ticker": {"type": "STRING"}, "weight": {"type": "NUMBER",
         "description": "percent of the portfolio"}}}}
AGENT_TOOLS = [
    dict(name="lookup_holdings", description="This week's model output for each ticker: report date, crowd chance of a beat, AI 1-in-10 "
         "worst case, expected size of the move, put price, AI action, crowd-rule action, drivers, last quarter's guidance.",
         parameters={"type": "OBJECT", "properties": {"tickers": {"type": "ARRAY", "items": {"type": "STRING"}}}, "required": ["tickers"]}),
    dict(name="what_if", description="Portfolio loss if every reporting holding had its 1-in-10 worst case, before and after the trims given.",
         parameters={"type": "OBJECT", "properties": {"holdings": _HOLD, "trims": {"type": "ARRAY", "items": {"type": "OBJECT", "properties": {
             "ticker": {"type": "STRING"}, "share": {"type": "NUMBER", "description": "percent of the holding sold"}}}}}, "required": ["holdings"]}),
    dict(name="size_trims", description="Proposes how much of each reporting holding to sell before its report, within a budget, "
         "respecting holdings the client will not sell and a cap per holding.",
         parameters={"type": "OBJECT", "properties": {"holdings": _HOLD, "budget_pct": {"type": "NUMBER", "description":
             "percent of the portfolio to sell in total; leave out for the default (what the 50% rule would sell)"},
             "do_not_sell": {"type": "ARRAY", "items": {"type": "STRING"}}, "max_trim_pct": {"type": "NUMBER",
             "description": "most of one holding to sell, percent"}}, "required": ["holdings"]}),
    dict(name="put_value", description="Price of a 5% out-of-the-money put before the report against its expected payoff from past reports.",
         parameters={"type": "OBJECT", "properties": {"ticker": {"type": "STRING"}}, "required": ["ticker"]}),
    dict(name="track_record", description="How earlier flags did after their reports (scorecard), for one ticker or for all.",
         parameters={"type": "OBJECT", "properties": {"ticker": {"type": "STRING"}}}),
]
TOOL_FN = {"lookup_holdings": tool_lookup, "what_if": tool_what_if, "size_trims": tool_size_trims, "put_value": tool_put_value,
           "track_record": tool_track_record}
AGENT_SYSTEM = ("You are BeatOdds, an assistant to a private banker. Use the tools: always call lookup_holdings for the holdings in "
                "the message before you answer, and call the other tools when the question needs them. Use only numbers the tools "
                "return, copied exactly and written as digits. State the suggested action for each holding. If the tools cannot answer, "
                "say so. No advice beyond the tools' actions. At most 250 words, plain British English, no em dashes.")


def genai_chat(contents, tools=None, system=None):
    if not USE_GENAI or not GEMINI_API_KEY.strip():
        return None
    body = {"contents": contents}
    if tools:
        body["tools"] = [{"functionDeclarations": tools}]
    if system:
        body["systemInstruction"] = {"parts": [{"text": system}]}
    js = _gem_post(body, timeout=120)
    try:
        return js["candidates"][0]["content"]
    except Exception:
        return None


def _run_tool(name, args):
    try:
        return TOOL_FN[name](**(args or {})) if name in TOOL_FN else dict(error=f"unknown tool {name}")
    except Exception as e:
        return dict(error=str(e)[:200])


def ask(text, show=True):
    if not os.path.exists(_p(F_AI_LIVE)):
        print("Run STEP 2 first: the agent reads today's model output.", flush=True)
        return None
    if not GEMINI_API_KEY.strip() or len(_GEM_DEAD) >= len(GEMINI_MODELS):
        return _ask_plain(text, show)
    contents, log, used, brief = [{"role": "user", "parts": [{"text": text}]}], [], [], ""
    for _ in range(AGENT_MAX_STEPS):
        c = genai_chat(contents, AGENT_TOOLS, AGENT_SYSTEM)
        if c is None:
            return _ask_plain(text, show)
        contents.append(c)
        calls = [p_["functionCall"] for p_ in c.get("parts", []) if "functionCall" in p_]
        if not calls:
            brief = " ".join(p_.get("text", "") for p_ in c.get("parts", []) if "text" in p_).strip()
            break
        resp = []
        for fc in calls:
            res = _run_tool(fc["name"], fc.get("args"))
            used.append(fc["name"])
            log.append(f"{fc['name']}({json.dumps(fc.get('args') or {})}) -> {json.dumps(res)}")
            resp.append({"functionResponse": {"name": fc["name"], "response": {"result": res}}})
        contents.append({"role": "user", "parts": resp})
    if not brief:
        return _ask_plain(text, show)
    facts = f"Client message: {text}\n" + "\n".join(log)
    g, drafts = guard_text(brief, facts), 1
    if g["unsupported"]:
        contents.append({"role": "user", "parts": [{"text": "These numbers are not in the tool results: " +
                                                     ", ".join(f"{v:g}" for v in g["unsupported"]) +
                                                     ". Rewrite the answer using only numbers from the tool results."}]})
        c = genai_chat(contents, None, AGENT_SYSTEM)
        b2 = " ".join(p_.get("text", "") for p_ in (c or {}).get("parts", []) if "text" in p_).strip()
        if b2:
            brief, drafts = b2, 2
            g = guard_text(brief, facts)
    status = ("checked: every number is in the tool results" if not g["unsupported"] else
              f"warning: {len(g['unsupported'])} numbers not in the tool results ({', '.join(f'{v:g}' for v in g['unsupported'])})")
    looked = [x.get("ticker") for fc in log if fc.startswith("lookup_holdings") for x in json.loads(fc.split(" -> ", 1)[1])]
    rec = dict(asked_utc=pd.Timestamp.now(tz="UTC").strftime("%Y-%m-%d %H:%M:%S"), request=text,
               holdings="; ".join(f"{t}:" for t in looked), facts=facts, brief=brief, n_numbers=g["n"],
               n_unsupported=len(g["unsupported"]), status=status + f" | tools: {', '.join(used)}", drafts=drafts, model=GEMINI_MODEL)
    old = pd.read_csv(_p(F_AGENT)) if os.path.exists(_p(F_AGENT)) else pd.DataFrame()
    pd.concat([old, pd.DataFrame([rec])], ignore_index=True).to_csv(_p(F_AGENT), index=False)
    with open(BRIEF_FILE, "w") as fh:
        fh.write(f"# BeatOdds client brief\n\n*{rec['asked_utc']} UTC. Request: {text}*\n\n{brief}\n\n---\n\nGuardrail: {status}. "
                 f"Tools used: {', '.join(used)}.\n\nTool results (facts):\n\n{facts}\n")
    if show:
        print(brief + f"\n\n[Guardrail: {status}. Tools used: {', '.join(used)}]", flush=True)
        try:
            from google.colab import files
            files.download(BRIEF_FILE)
        except Exception:
            pass
    return brief


# ---------- Workbook sheets for the additions ----------
def add_extra_sheets(wb, A):
    from openpyxl.styles import Alignment, Font, PatternFill
    BOLD = Font(name="Arial", bold=True); NOTE = Font(name="Arial", italic=True, color="666666")
    TITLE = Font(name="Arial", bold=True, size=14, color="1F3A5F"); SEC = Font(name="Arial", bold=True, size=11, color="1F3A5F")
    HEAD = PatternFill("solid", fgColor="DDE4EE"); BLUE = Font(name="Arial", color="0000FF")
    PCT, PC1 = "+0.0%;-0.0%", "0.0%"

    def sheet(name, after):
        return wb.create_sheet(name, index=wb.sheetnames.index(after) + 1) if after else wb.create_sheet(name)

    def head(ws, r, vals, h=45):
        for j, v in enumerate(vals):
            c = ws.cell(row=r, column=1 + j, value=v); c.font = BOLD; c.fill = HEAD
            c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r].height = h

    def put(ws, r, c, v, nf=None, font=None):
        x = ws.cell(row=r, column=c, value=v)
        if nf:
            x.number_format = nf
        if font:
            x.font = font
        return x

    def title(ws, t, sub):
        ws["A1"] = t; ws["A1"].font = TITLE
        ws["A2"] = sub; ws["A2"].alignment = Alignment(wrap_text=True, vertical="top"); ws.merge_cells("A2:L2"); ws.row_dimensions[2].height = 60

    def raw(ws, note, cols, df, nfs=None):
        ws["A1"] = note; ws["A1"].font = BOLD
        head(ws, 4, cols)
        for i, rec in enumerate(df.itertuples(index=False)):
            for j, v in enumerate(rec):
                if isinstance(v, (float, np.floating)) and not np.isfinite(v):
                    v = None
                elif isinstance(v, np.generic):
                    v = v.item()
                put(ws, 5 + i, 1 + j, v, (nfs or {}).get(j))
        ws.freeze_panes = "A5"
        return 4 + max(len(df), 1)

    rd_csv = lambda f, cols: pd.read_csv(_p(f)).reindex(columns=cols) if os.path.exists(_p(f)) else pd.DataFrame(columns=cols)

    # ---------- Scorecard ----------
    sc = rd_csv(F_SCORE, SCORE_COLS)
    RSC = sheet("Raw_Scorecard", None)
    E = raw(RSC, "Live scorecard written by the script: each flag shown on the Live sheet (time = snapshot used), the EDGAR release "
                 "time, and what happened. Columns A to O as saved. Columns P to Y are formulas.",
            ["Contract", "Ticker", "Report date", "Flag made (UTC)", "Release (UTC, EDGAR)", "Flag made before the release (1/0)",
             "Crowd p", "AI 10% worst case", "AI cut-off", "AI action", "Rule action", "Put price / share price", "Abnormal return",
             "Share return", "Outcome (1 = beat)"], sc, {6: "0%", 7: PCT, 8: PCT, 11: "0.00%", 12: PCT, 13: PCT})
    for j, h in enumerate(["AI trims (1/0)", "Rule trims (1/0)", "Scored (1/0)", "Fell 5% or more (1/0)", "Below the AI worst case (1/0)",
                           "Outcome: hold", "Outcome: rule trims", "Outcome: AI trims", "Put payoff / share price", "Put P&L / share price"]):
        c = RSC.cell(row=4, column=16 + j, value=h); c.font = BOLD; c.fill = HEAD; c.alignment = Alignment(wrap_text=True, vertical="top")
    for i in range(len(sc)):
        q = 5 + i
        put(RSC, q, 16, f'=IF(AND(ISNUMBER(H{q}),ISNUMBER(I{q})),IF(H{q}<=I{q},1,0),"")')
        put(RSC, q, 17, f'=IF(ISNUMBER(G{q}),IF(G{q}>={A["HP"]},1,0),"")')
        put(RSC, q, 18, f'=IF(AND(F{q}=1,ISNUMBER(M{q})),1,0)')
        put(RSC, q, 19, f'=IF(R{q}=1,IF(M{q}<={A["FALL"]},1,0),"")')
        put(RSC, q, 20, f'=IF(AND(R{q}=1,ISNUMBER(H{q})),IF(M{q}<H{q},1,0),"")')
        put(RSC, q, 21, f'=IF(R{q}=1,M{q},"")', PCT)
        put(RSC, q, 22, f'=IF(R{q}=1,M{q}*(1-{A["TRIM"]}*N(Q{q}))-{A["TCOST"]}*{A["TRIM"]}*N(Q{q}),"")', PCT)
        put(RSC, q, 23, f'=IF(R{q}=1,M{q}*(1-{A["TRIM"]}*N(P{q}))-{A["TCOST"]}*{A["TRIM"]}*N(P{q}),"")', PCT)
        put(RSC, q, 24, f'=IF(AND(R{q}=1,ISNUMBER(N{q})),MAX(0,(1-{A["OTM"]})-(1+N{q})),"")', "0.00%")
        put(RSC, q, 25, f'=IF(AND(ISNUMBER(X{q}),ISNUMBER(L{q})),X{q}-L{q},"")', "+0.00%;-0.00%")
    SR = lambda c: f"Raw_Scorecard!${c}$5:${c}${E}"
    hs = verify_flags()
    RFH = sheet("Raw_FlagHash", None)
    HE = raw(RFH, "SHA-256 hash of each day's flags (flags_log.csv rows of that snapshot, sorted by contract), written when the flags "
                  "were made. Column E: the script recomputed the hash from the saved flags on this run (1 = unchanged).",
             ["Snapshot (UTC)", "Flags", "SHA-256", "Logged (UTC)", "Still matches (1/0)"], hs)
    HR = lambda c: f"Raw_FlagHash!${c}$5:${c}${HE}"
    SCS = sheet("Scorecard", "Live")
    title(SCS, "Scorecard: our dated flags against what happened",
          "Every flag on the Live sheet is saved with the time it was made. A flag counts only if it was made before the company's "
          "release (EDGAR time, or 10:00 UTC on the report day if the time is unknown). Outcome = abnormal return over the same window "
          "as Events. Few reports so far: read this as a track record that grows each week, not as a test.")
    r = 4
    head(SCS, r, ["", "Value"]); r += 1
    sc0 = r
    for lab, v, nf in (("Flagged reports whose report day has passed", f"=COUNTA({SR('A')})", "0"),
                       ("Flags made before the release", f"=COUNTIF({SR('F')},1)", "0"),
                       ("Flags made after the release (not counted)", f"=COUNTIF({SR('F')},0)", "0"),
                       ("Scored: flag before the release and the share price after the report available", f"=SUM({SR('R')})", "0"),
                       ("Share of scored reports that fell below the AI 1-in-10 worst case (target 10%)", f'=IFERROR(AVERAGE({SR("T")}),"n/a")', PC1),
                       ("Days of flags with a SHA-256 hash (Raw_FlagHash)", f"=COUNTA({HR('A')})", "0"),
                       ("Of which the saved flags still match the hash", f"=SUM({HR('E')})", "0")):
        put(SCS, r, 1, lab); put(SCS, r, 2, v, nf); r += 1
    r += 1
    put(SCS, r, 1, "What happened, by what the product said", font=SEC); r += 1
    head(SCS, r, ["Group", "Reports", "Average abnormal return", "Share that fell 5% or more", "Share that beat"]); r += 1
    for lab, col, val in (("AI said trim", "P", 1), ("AI said hold", "P", 0), ("Crowd rule said trim", "Q", 1), ("Crowd rule said hold", "Q", 0)):
        put(SCS, r, 1, lab)
        put(SCS, r, 2, f"=COUNTIFS({SR('R')},1,{SR(col)},{val})", "0")
        put(SCS, r, 3, f'=IFERROR(AVERAGEIFS({SR("M")},{SR("R")},1,{SR(col)},{val}),"n/a")', PCT)
        put(SCS, r, 4, f'=IFERROR(AVERAGEIFS({SR("S")},{SR("R")},1,{SR(col)},{val}),"n/a")', PC1)
        put(SCS, r, 5, f'=IFERROR(AVERAGEIFS({SR("O")},{SR("R")},1,{SR(col)},{val},{SR("O")},">=0"),"n/a")', PC1)
        r += 1
    r += 1
    head(SCS, r, ["Policy on the scored reports (one equal position each)", "Average outcome", "Average of the worst 10%", "Worst"]); r += 1
    for lab, col in (("Hold everything", "U"), ("Crowd rule: trim when confident", "V"), ("AI: trim the worst predicted downside", "W")):
        put(SCS, r, 1, lab)
        put(SCS, r, 2, f'=IFERROR(AVERAGE({SR(col)}),"n/a")', PCT)
        put(SCS, r, 3, f'=IFERROR(AVERAGEIF({SR(col)},"<="&PERCENTILE({SR(col)},0.1)),"n/a")', PCT)
        put(SCS, r, 4, f'=IFERROR(MIN({SR(col)}),"n/a")', PCT)
        r += 1
    r += 1
    put(SCS, r, 1, "What the backtest result is worth in euros (Backtest sheet, out of sample)", font=SEC); r += 1
    put(SCS, r, 1, "Portfolio value (EUR)"); put(SCS, r, 2, 10_000_000, "#,##0", BLUE); eu0 = r; r += 1
    put(SCS, r, 1, "Equal-weight holdings, each reporting once a season"); put(SCS, r, 2, 20, "0", BLUE); r += 1
    put(SCS, r, 1, "Position per holding (EUR)"); put(SCS, r, 2, f"=B{eu0}/B{eu0 + 1}", "#,##0"); r += 1
    put(SCS, r, 1, "Reports in the worst 10% per season (holdings x 10%)"); put(SCS, r, 2, f"=B{eu0 + 1}*0.1", "0.0"); r += 1
    head(SCS, r, ["Policy", "Average of the worst 10% (Backtest)", "Loss per position in a bad report (EUR)",
                  "Saved per bad report vs holding (EUR)", "Saved per season vs holding (EUR)", "Saved per season, % of the portfolio"]); r += 1
    h_ = r
    for lab, br in (("Hold everything", 5), ("Crowd rule: trim when confident", 6), ("Smarter spreadsheet: trim the biggest past moves", 10),
                    ("Smarter spreadsheet: trim the biggest implied moves (LSEG)", 11), ("AI: trim the worst predicted downside", 7)):
        put(SCS, r, 1, lab); put(SCS, r, 2, f"=Backtest!E{br}", PCT)
        put(SCS, r, 3, f'=IFERROR(B{r}*$B${eu0 + 2},"n/a")', "#,##0")
        put(SCS, r, 4, f'=IFERROR(C{r}-$C${h_},"n/a")', "#,##0")
        put(SCS, r, 5, f'=IFERROR(D{r}*$B${eu0 + 3},"n/a")', "#,##0")
        put(SCS, r, 6, f'=IFERROR(E{r}/$B${eu0},"n/a")', "0.00%")
        r += 1
    put(SCS, r + 1, 1, "Illustration only: it applies the backtest averages to one equal-weight portfolio. Returns are abnormal "
                       "returns (against the market), before tax, trimming costs included (Assumptions).", font=NOTE)
    SCS.column_dimensions["A"].width = 70
    for c_ in "BCDEF":
        SCS.column_dimensions[c_].width = 18
    SCS.sheet_properties.tabColor = "C00000"

    # ---------- AI_Checks ----------
    st_cols = ["slug", "variant", "text", "truth_ticker", "truth_date", "truth_eps", "truth_thr", "rule_ticker", "rule_date", "rule_eps",
               "rule_thr", "ai_ticker", "ai_date", "ai_eps", "ai_thr", "seconds"]
    st = rd_csv(F_STRESS, st_cols)
    st["text"] = st["text"].astype(str).str.slice(0, 400)
    RST = sheet("Raw_Stress", None)
    SE = raw(RST, "Stress test written by the script: Polymarket contract text (question and rules, no slug), as listed and with four "
                  "format changes (our own rewrites of the real text), read by our rules and by generative AI. Truth = the full parser "
                  "with the slug. Columns A to P as saved. Columns Q to Z are formulas.",
             ["Contract", "Text version", "Text read (first 400 characters)", "Truth: ticker", "Truth: report date", "Truth: EPS type",
              "Truth: threshold", "Rules: ticker", "Rules: report date", "Rules: EPS type", "Rules: threshold", "AI: ticker",
              "AI: report date", "AI: EPS type", "AI: threshold", "AI seconds per contract"], st)
    tick = lambda a, b, q: f'=IF({a}{q}="","",IF(UPPER(SUBSTITUTE({b}{q},"-","."))=UPPER(SUBSTITUTE({a}{q},"-",".")),1,0))'
    datef = lambda a, b, q: f'=IF({a}{q}="","",IF(LEFT({b}{q}&"",10)=LEFT({a}{q}&"",10),1,0))'
    epsf = lambda a, b, q: f'=IF({a}{q}="","",IF(LOWER(SUBSTITUTE({b}{q},"-",""))=LOWER(SUBSTITUTE({a}{q},"-","")),1,0))'
    thrf = lambda a, b, q: f'=IF(ISNUMBER({a}{q}),IF(ISNUMBER({b}{q}),IF(ABS({b}{q}-{a}{q})<=0.005,1,0),0),"")'
    hd = ["Rules: ticker right", "Rules: date right", "Rules: EPS type right", "Rules: threshold right",
          "AI: ticker right", "AI: date right", "AI: EPS type right", "AI: threshold right", "Rules: all four right", "AI: all four right"]
    for j, h in enumerate(hd):
        c = RST.cell(row=4, column=17 + j, value=h); c.font = BOLD; c.fill = HEAD; c.alignment = Alignment(wrap_text=True, vertical="top")
    for i in range(len(st)):
        q = 5 + i
        for j, (fn, a, b) in enumerate(((tick, "D", "H"), (datef, "E", "I"), (epsf, "F", "J"), (thrf, "G", "K"),
                                        (tick, "D", "L"), (datef, "E", "M"), (epsf, "F", "N"), (thrf, "G", "O"))):
            put(RST, q, 17 + j, fn(a, b, q))
        put(RST, q, 25, f'=IF(D{q}="","",IF(SUM(Q{q}:T{q})=4,1,0))')
        put(RST, q, 26, f'=IF(D{q}="","",IF(SUM(U{q}:X{q})=4,1,0))')
    KM_ = ["series", "title", "ai_company", "ai_ticker", "rule_ticker", "final_ticker", "events", "events_rules", "events_with_ai",
           "answered", "seconds"]
    kmap = rd_csv(F_KMAP, KM_)
    RKP = sheet("Raw_K_Map", None)
    KE = raw(RKP, "Kalshi KPI series mapped to a stock: the rules (series code) and generative AI (contract title). A ticker counts "
                  "only if the company's earnings dates match the contract dates. Columns A to K as saved. Columns L and M are formulas.",
             ["Series", "Title (most traded contract)", "AI: company", "AI: ticker", "Rules: ticker linked", "Ticker linked with AI",
              "KPI events (after volume filter)", "Events linked by rules", "Events linked by rules + AI", "AI answered (1/0)",
              "AI seconds per series"], kmap)
    for j, h in enumerate(["Events linked only thanks to AI", "AI agrees with the rules (1/0)"]):
        c = RKP.cell(row=4, column=12 + j, value=h); c.font = BOLD; c.fill = HEAD
    for i in range(len(kmap)):
        q = 5 + i
        put(RKP, q, 12, f"=MAX(0,N(I{q})-N(H{q}))", "0")
        put(RKP, q, 13, f'=IF(AND(D{q}<>"",E{q}<>""),IF(UPPER(D{q})=UPPER(E{q}),1,0),"")')
    gd_cols = ["kind", "key", "ticker", "text", "facts", "n_numbers", "n_unsupported", "unsupported", "action_ok", "drafts"]
    gd = rd_csv(F_GUARD, gd_cols)
    RGD = sheet("Raw_Guard", None)
    GE = raw(RGD, "Guardrail written by the script: each AI client note and agent brief, the facts it was written from, how many numbers "
                  "it contains and which of them are not in the facts. Columns A to J as saved. Column K is a formula.",
             ["Kind", "Key", "Ticker(s)", "Text", "Facts given to the AI", "Numbers in the text", "Numbers not in the facts",
              "Which", "Suggested action stated correctly (1/0)", "Drafts"], gd)
    RGD.cell(row=4, column=11, value="All numbers supported (1/0)").font = BOLD
    for i in range(len(gd)):
        put(RGD, 5 + i, 11, f"=IF(G{5 + i}=0,1,0)")
    AC = sheet("AI_Checks", "GenAI")
    title(AC, "AI_Checks: can the AI be trusted with the jobs rules cannot do?",
          "A) Stress test: contract text only, as listed and with four format changes a new venue could bring. B) Kalshi KPI contracts "
          "linked to a company by generative AI where our rules fail. C) Guardrail: every number in every AI text checked against the "
          "data it was given. D) Time and cost. All cells are formulas on Raw_Stress, Raw_K_Map and Raw_Guard.")
    STR = lambda c: f"Raw_Stress!${c}$5:${c}${SE}"
    r = 4
    put(AC, r, 1, "A) Stress test: share of fields read right", font=SEC); r += 1
    head(AC, r, ["Text version", "Contracts", "Rules: ticker", "Rules: date", "Rules: EPS type", "Rules: threshold", "Rules: all four",
                 "AI: ticker", "AI: date", "AI: EPS type", "AI: threshold", "AI: all four"]); r += 1
    ac_a = {}
    for lab, _ in STRESS_VARIANTS:
        ac_a[lab] = r
        put(AC, r, 1, lab); put(AC, r, 2, f'=COUNTIF({STR("B")},A{r})', "0")
        for j, c_ in enumerate("QRSTYUVWXZ"):
            put(AC, r, 3 + j, f'=IFERROR(AVERAGEIFS({STR(c_)},{STR("B")},$A{r}),"n/a")', "0%")
        r += 1
    put(AC, r, 1, "Contracts our rules cannot fully parse today (no ticker or EPS type not stated)")
    put(AC, r, 2, f'=COUNTIF({STR("B")},"rules could not parse it fully")', "0")
    put(AC, r, 3, "AI gave a ticker:"); put(AC, r, 4, f'=COUNTIFS({STR("B")},"rules could not parse it fully",{STR("L")},"?*")', "0")
    put(AC, r, 5, "AI gave an EPS type:")
    put(AC, r, 6, f'=COUNTIFS({STR("B")},"rules could not parse it fully",{STR("N")},"*GAAP*")', "0"); r += 2
    KR = lambda c: f"Raw_K_Map!${c}$5:${c}${KE}"
    put(AC, r, 1, "B) Kalshi KPI contracts linked to a company report", font=SEC); r += 1
    head(AC, r, ["", "Value"]); r += 1
    b0 = r
    for lab, v, nf in (("KPI series", f"=COUNTA({KR('A')})", "0"), ("KPI events (after the volume filter)", f"=SUM({KR('G')})", "0"),
                       ("Events linked by our rules", f"=SUM({KR('H')})", "0"), ("Events linked by rules + generative AI", f"=SUM({KR('I')})", "0"),
                       ("Increase in usable Kalshi events", f'=IFERROR(B{b0 + 3}/B{b0 + 2}-1,"n/a")', PCT),
                       ("Series the AI linked where the rules could not", f'=COUNTIFS({KR("E")},"",{KR("F")},"?*")', "0"),
                       ("Where both linked a ticker: AI agrees with the rules", f'=IFERROR(AVERAGE({KR("M")}),"n/a")', PC1)):
        put(AC, r, 1, lab); put(AC, r, 2, v, nf); r += 1
    r += 1
    GR = lambda c: f"Raw_Guard!${c}$5:${c}${GE}"
    put(AC, r, 1, "C) Guardrail: are the numbers in the AI texts in the data?", font=SEC); r += 1
    head(AC, r, ["Kind", "Texts", "Numbers checked", "Numbers not in the data", "Share not in the data (hallucination rate)",
                 "Texts with at least one", "Action stated correctly"]); r += 1
    c0 = r
    for kind in ("client note", "agent brief"):
        put(AC, r, 1, kind); put(AC, r, 2, f'=COUNTIF({GR("A")},A{r})', "0")
        put(AC, r, 3, f'=SUMIFS({GR("F")},{GR("A")},A{r})', "0"); put(AC, r, 4, f'=SUMIFS({GR("G")},{GR("A")},A{r})', "0")
        put(AC, r, 5, f'=IFERROR(D{r}/C{r},"n/a")', PC1)
        put(AC, r, 6, f'=COUNTIFS({GR("A")},A{r},{GR("G")},">0")', "0")
        put(AC, r, 7, f'=IFERROR(AVERAGEIFS({GR("I")},{GR("A")},A{r}),"n/a")', PC1)
        r += 1
    put(AC, r, 1, "Notes written before the guardrail existed have 1 draft. New notes get one redraft when a number is not in the data.",
        font=NOTE); r += 2
    put(AC, r, 1, "D) Time and cost: reading contracts by hand against generative AI", font=SEC); r += 1
    head(AC, r, ["", "Value"]); r += 1
    d0 = r
    put(AC, r, 1, "Minutes for an analyst to map one contract by hand (assumption)"); put(AC, r, 2, HAND_MINUTES, "0.0", BLUE); r += 1
    put(AC, r, 1, "Analyst cost per hour, EUR (assumption)"); put(AC, r, 2, 60, "0", BLUE); r += 1
    put(AC, r, 1, "Contracts read so far: Polymarket contracts + Kalshi KPI events")
    put(AC, r, 2, f"=COUNTA(Raw_Markets!$A$5:$A${wb['Raw_Markets'].max_row})+B{b0 + 1}", "#,##0"); r += 1
    put(AC, r, 1, "By hand: hours"); put(AC, r, 2, f"=B{d0 + 2}*B{d0}/60", "#,##0"); r += 1
    put(AC, r, 1, "By hand: cost, EUR"); put(AC, r, 2, f"=B{d0 + 3}*B{d0 + 1}", "#,##0"); r += 1
    put(AC, r, 1, "Generative AI: seconds per contract (measured)")
    put(AC, r, 2, f'=IFERROR(AVERAGE({STR("P")},{KR("K")}),"n/a")', "0.0"); r += 1
    put(AC, r, 1, "Generative AI: hours"); put(AC, r, 2, f'=IFERROR(B{d0 + 2}*B{d0 + 5}/3600,"n/a")', "0.0"); r += 1
    put(AC, r, 1, "Generative AI: cost, EUR (free Gemini tier)"); put(AC, r, 2, 0, "0", BLUE); r += 1
    AC.column_dimensions["A"].width = 62
    for c_ in "BCDEFGHIJKL":
        AC.column_dimensions[c_].width = 12
    AC.sheet_properties.tabColor = "1F3A5F"

    # ---------- Agent ----------
    ag_cols = ["asked_utc", "request", "holdings", "facts", "brief", "n_numbers", "n_unsupported", "status", "drafts", "model"]
    ag = rd_csv(F_AGENT, ag_cols)
    RAG = sheet("Raw_Agent", None)
    AE = raw(RAG, "Portfolio agent briefs written by the script (ask() in Colab): the request, the holdings it read, the model output "
                  "it looked up (facts), the brief, and the guardrail result.",
             ["Asked (UTC)", "Request", "Holdings read (ticker:weight %)", "Facts given to the AI", "Brief", "Numbers in the brief",
              "Numbers not in the facts", "Guardrail", "Drafts", "Model"], ag)
    AG = sheet("Agent", "AI_Checks")
    title(AG, "Agent: a client portfolio in, a checked client brief out",
          "In Colab: ask(\"I hold AAPL 10%, PLTR 5%. Anything to do before their reports?\"). The agent reads the holdings with "
          "generative AI, looks each one up in today's model output, writes the brief from those numbers only and checks every number. "
          "The table below repeats the lookup with formulas for the last brief: change a blue ticker or weight to try another portfolio.")
    last = f"MAX(1,COUNTA(Raw_Agent!$A$5:$A${AE}))"
    put(AG, 4, 1, "Last request"); put(AG, 4, 2, f'=IFERROR(INDEX(Raw_Agent!$B$5:$B${AE},{last}),"no brief yet")')
    put(AG, 5, 1, "Guardrail"); put(AG, 5, 2, f'=IFERROR(INDEX(Raw_Agent!$H$5:$H${AE},{last}),"")')
    put(AG, 6, 1, "Brief"); put(AG, 6, 2, f'=IFERROR(INDEX(Raw_Agent!$E$5:$E${AE},{last}),"")')
    AG.merge_cells("B6:H6"); AG["B6"].alignment = Alignment(wrap_text=True, vertical="top"); AG.row_dimensions[6].height = 260
    head(AG, 8, ["Ticker", "Weight (%)", "Report date", "Crowd p", "AI 10% worst case", "AI action", "Weight x worst case (% of portfolio)"])
    hl = [h.split(":") for h in str(ag["holdings"].iloc[-1]).split("; ")] if len(ag) and isinstance(ag["holdings"].iloc[-1], str) else []
    AL = lambda c: f"Raw_AI_Live!${c}$5:${c}$2000"
    for i in range(max(len(hl), 8)):
        q = 9 + i
        t_, w_ = (hl[i] + [""])[:2] if i < len(hl) else ("", "")
        put(AG, q, 1, t_ or None, None, BLUE); put(AG, q, 2, _f(w_), "0.0", BLUE)
        m = f"MATCH(A{q},{AL('B')},0)"
        put(AG, q, 3, f'=IF(A{q}="","",IFERROR(INDEX({AL("C")},{m}),"no report"))', "yyyy-mm-dd")
        put(AG, q, 4, f'=IF(A{q}="","",IFERROR(INDEX({AL("D")},{m}),""))', "0%")
        put(AG, q, 5, f'=IF(A{q}="","",IFERROR(INDEX({AL("E")},{m}),""))', PCT)
        put(AG, q, 6, f'=IF(ISNUMBER(E{q}),IF(E{q}<=INDEX({AL("G")},{m}),"trim {TRIM:.0%}","hold"),"")')
        put(AG, q, 7, f'=IF(AND(ISNUMBER(B{q}),ISNUMBER(E{q})),B{q}/100*E{q},"")', "+0.00%;-0.00%")
    qe = 9 + max(len(hl), 8)
    put(AG, qe, 1, "If every reporting holding had its 1-in-10 worst case", font=BOLD); put(AG, qe, 7, f"=SUM(G9:G{qe - 1})", "+0.00%;-0.00%")
    put(AG, qe + 1, 1, "After trimming the flagged holdings by the Assumptions share", font=BOLD)
    put(AG, qe + 1, 7, f'=G{qe}-{A["TRIM"]}*SUMIF(F9:F{qe - 1},"trim*",G9:G{qe - 1})', "+0.00%;-0.00%")
    AG.column_dimensions["A"].width = 26; AG.column_dimensions["B"].width = 14
    for c_ in "CDEFGH":
        AG.column_dimensions[c_].width = 16
    AG.sheet_properties.tabColor = "2E7D32"

    # ---------- Summary: section 8 ----------
    S = wb["Summary"]
    r = S.max_row + 2
    S.cell(row=r, column=1, value="8. New in this version: track record, trust checks, the agent").font = SEC; r += 1
    a1, a2 = ac_a["as listed"], ac_a["all four changes"]
    for lab, v, nf, src in (
            ("Live flags made before the release / scored so far", f'=Scorecard!B{sc0 + 1}&" / "&Scorecard!B{sc0 + 3}', None, "Scorecard"),
            ("Scored reports below the AI 1-in-10 worst case (target 10%)", f"=Scorecard!B{sc0 + 4}", PC1, "Scorecard"),
            ("Euro value: AI saved per season vs holding (EUR 10m, 20 holdings)", f"=Scorecard!E{h_ + 4}", "#,##0", "Scorecard"),
            ("Contract text as listed: all four fields right, rules / AI", f'=TEXT(AI_Checks!G{a1},"0%")&" / "&TEXT(AI_Checks!L{a1},"0%")', None, "AI_Checks A"),
            ("After four format changes: all four fields right, rules / AI", f'=TEXT(AI_Checks!G{a2},"0%")&" / "&TEXT(AI_Checks!L{a2},"0%")', None, "AI_Checks A"),
            ("Kalshi KPI events usable: rules / rules + AI", f'=AI_Checks!B{b0 + 2}&" / "&AI_Checks!B{b0 + 3}', None, "AI_Checks B"),
            ("Guardrail: numbers in AI client notes not in the data", f"=AI_Checks!E{c0}", PC1, "AI_Checks C"),
            ("Guardrail: numbers in agent briefs not in the data", f"=AI_Checks!E{c0 + 1}", PC1, "AI_Checks C")):
        S.cell(row=r, column=1, value=lab)
        c = S.cell(row=r, column=3, value=v)
        if nf:
            c.number_format = nf
        S.cell(row=r, column=9, value=src); r += 1
    for nm in ("Raw_Scorecard", "Raw_FlagHash", "Raw_Stress", "Raw_K_Map", "Raw_Guard", "Raw_Agent"):
        wb[nm].sheet_properties.tabColor = "A6A6A6"


def _pctl_tail(a, q=10):
    a = np.asarray(a, float); a = a[np.isfinite(a)]
    if len(a) == 0:
        return np.nan
    thr = np.percentile(a, q)
    return a[a <= thr].mean()


BOOT_STATS = [("miss_gap", "Average move on a miss: confident minus less sure"),
              ("mean_gap", "Average move: confident minus less sure"),
              ("fall_gap", "Share falling 5% or more: confident minus less sure"),
              ("tail_gap", "Average of the worst 10%: confident minus less sure"),
              ("bt_tail_ai_rule", "Backtest: average of the worst 10%, AI policy minus spreadsheet rule"),
              ("bt_mean_ai_rule", "Backtest: average outcome, AI policy minus spreadsheet rule"),
              ("bt_tail_ai_hold", "Backtest: average of the worst 10%, AI policy minus holding everything"),
              ("bt_tail_ai_past", "Backtest: average of the worst 10%, AI policy minus trimming the biggest past moves"),
              ("bt_tail_ai_iv", "Backtest: average of the worst 10%, AI policy minus trimming the biggest implied moves"),
              ("bt_tail_ai_el", "Backtest: average of the worst 10%, AI policy minus trimming the biggest expected loss"),
              ("bt_tail_fall10_past", "Backtest: worst 10%, pre-registered 1 (fall classifier) minus trimming the biggest past moves"),
              ("bt_tail_resid_past", "Backtest: worst 10%, pre-registered 2 (AI learns k) minus trimming the biggest past moves"),
              ("bt_tail_size_past", "Backtest: worst 10%, pre-registered 3 (sized trims) minus trimming the biggest past moves"),
              ("bt_tail_ai_past_nohist", "Backtest, reports without 3 past reports: worst 10%, AI minus the past-move rule"),
              ("bt_tail_cw_past", "Backtest: worst 10%, crowd-warning screen minus trimming the biggest past moves"),
              ("bt_ca_drift_tail", "Crowd lost confidence over the week: worst 10% minus the other reports"),
              ("bt_ca_drift_abs", "Crowd lost confidence over the week: average |move| minus the other reports"),
              ("bt_ca_sm_tail", "Confident crowd, StarMine expects a miss: worst 10% minus StarMine agrees"),
              ("bt_ca_rev_tail", "Confident crowd, analysts cut estimates: worst 10% minus estimates up or flat")]


def _stats(d):
    c, l = d[d["confident"] == 1], d[d["confident"] == 0]
    return dict(miss_gap=c.loc[c["y"] == 0, "ar"].mean() - l.loc[l["y"] == 0, "ar"].mean(),
                mean_gap=c["ar"].mean() - l["ar"].mean(),
                fall_gap=c["fall5"].mean() - l["fall5"].mean(),
                tail_gap=_pctl_tail(c["ar"]) - _pctl_tail(l["ar"]),
                bt_tail_ai_rule=_pctl_tail(d["out_ai"]) - _pctl_tail(d["out_rule"]) if "out_ai" in d else np.nan,
                bt_mean_ai_rule=d["out_ai"].mean() - d["out_rule"].mean() if "out_ai" in d else np.nan,
                bt_tail_ai_hold=_pctl_tail(d["out_ai"]) - _pctl_tail(d["out_hold"]) if "out_ai" in d else np.nan,
                **{f"bt_tail_ai_{k[:-5]}": (_pctl_tail(d["out_ai"]) - _pctl_tail(d[f"out_{k}"]) if f"out_{k}" in d else np.nan)
                   for k, *_ in SHEET_RULES},
                **{f"bt_tail_{k}_past": (_pctl_tail(d[f"out_{o}"]) - _pctl_tail(d["out_past_trim"]) if f"out_{o}" in d else np.nan)
                   for k, o in (("fall10", "fall10_trim"), ("resid", "resid_trim"), ("size", "size"))},
                bt_tail_ai_past_nohist=(_pctl_tail(d.loc[~d["hist_ok"], "out_ai"]) - _pctl_tail(d.loc[~d["hist_ok"], "out_past_trim"])
                                        if "hist_ok" in d else np.nan),
                bt_tail_cw_past=(_pctl_tail(d["out_cw"]) - _pctl_tail(d["out_past_trim"]) if "out_cw" in d else np.nan),
                **_crowd_stats(d))


def _crowd_stats(d):
    nan = pd.Series(np.nan, index=d.index)
    dp, sm, rv = d.get("dp7", nan), d.get("smart_surp", nan), d.get("rev_lseg", nan)
    lost, rest = d.loc[dp <= -DRIFT, "ar"], d.loc[dp > -DRIFT, "ar"]
    c = d["confident"] == 1
    return dict(bt_ca_drift_tail=_pctl_tail(lost) - _pctl_tail(rest),
                bt_ca_drift_abs=lost.abs().mean() - rest.abs().mean(),
                bt_ca_sm_tail=_pctl_tail(d.loc[c & (sm < 0), "ar"]) - _pctl_tail(d.loc[c & (sm >= 0), "ar"]),
                bt_ca_rev_tail=_pctl_tail(d.loc[c & (rv < 0), "ar"]) - _pctl_tail(d.loc[c & (rv >= 0), "ar"]))


def run_bootstrap(ev):
    rng = np.random.default_rng(SEED)
    cols = {}
    for smp in ("All", "New"):
        d = ev[ev["ar"].notna() & ev["p"].notna()]
        if smp == "New":
            d = d[d["sample"] == "New"]
        weeks = d["week"].unique()
        enough = len(d) >= 30 and len(weeks) >= 5 and min(((d["y"] == 0) & (d["confident"] == g)).sum() for g in (0, 1)) >= 10
        groups = {w: g for w, g in d.groupby("week")}
        draws = []
        for _ in range(BOOT if enough else 0):
            pick = rng.choice(weeks, size=len(weeks), replace=True)
            draws.append(_stats(pd.concat([groups[w] for w in pick], ignore_index=True)))
        for k, _ in BOOT_STATS:
            cols[f"{smp}: {k}"] = [x[k] for x in draws]
    n = max([len(v) for v in cols.values()] + [0])
    return pd.DataFrame({k: v + [np.nan] * (n - len(v)) for k, v in cols.items()})


# 7. Workbook: Raw_ sheets hold data, every other cell is a formula
def bs_put(S, K, v, T):
    d1 = f"((LN({S}/{K})+{v}^2*{T}/2)/({v}*SQRT({T})))"
    return f"{K}*NORMSDIST(-({d1}-{v}*SQRT({T})))-{S}*NORMSDIST(-{d1})"


def build(mk, od, px, est, snap, km, ko, ev, models, boot, download_utc, ed=None, oh=None, ls=None, rel=None):
    from openpyxl import Workbook
    from openpyxl.chart import Reference, ScatterChart, Series
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter as L
    from openpyxl.worksheet.formula import ArrayFormula

    BOLD = Font(name="Arial", bold=True); NOTE = Font(name="Arial", italic=True, color="666666")
    TITLE = Font(name="Arial", bold=True, size=14, color="1F3A5F"); SEC = Font(name="Arial", bold=True, size=11, color="1F3A5F")
    HEAD = PatternFill("solid", fgColor="DDE4EE"); BLUE = Font(name="Arial", color="0000FF")
    PCT, PC1, PC2 = "+0.0%;-0.0%", "0.0%", "0.00%"

    def head(ws, r, vals, c0=1, h=45):
        for j, v in enumerate(vals):
            c = ws.cell(row=r, column=c0 + j, value=v); c.font = BOLD; c.fill = HEAD
            c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r].height = h

    def title(ws, t, sub, span="A2:J2", h=45):
        ws["A1"] = t; ws["A1"].font = TITLE
        ws["A2"] = sub; ws["A2"].alignment = Alignment(wrap_text=True, vertical="top"); ws.merge_cells(span); ws.row_dimensions[2].height = h

    def put(ws, r, c, v, nf=None, font=None):
        x = ws.cell(row=r, column=c, value=v)
        if nf:
            x.number_format = nf
        if font:
            x.font = font
        return x

    def clean(v):
        if v is None:
            return None
        if isinstance(v, pd.Timestamp):
            return v.to_pydatetime()
        if isinstance(v, (float, np.floating)):
            return float(v) if np.isfinite(v) else None
        if isinstance(v, np.generic):
            return v.item()
        if isinstance(v, str) and v.lower() == "nan":
            return None
        return v

    wb = Workbook(); wb.remove(wb.active)
    names = ["Summary", "Results", "Tests", "Tail", "Hedge", "Backtest", "Live", "Models", "GenAI", "Bootstrap", "Kalshi_Results",
             "Events", "Test_Calc", "Weeks", "Model_Calc", "Kalshi_Events", "Assumptions", "Raw_Markets", "Raw_Odds", "Raw_Prices",
             "Raw_Estimates", "Raw_Snapshots", "Raw_EDGAR", "Raw_OptHist", "Raw_LSEG", "Raw_Models", "Raw_Bootstrap", "Raw_AI_Live",
             "Raw_GenAI", "Raw_K_Markets", "Raw_K_Odds", "Raw_Releases", "Raw_Features"]
    W = {n: wb.create_sheet(n) for n in names}
    SUM_, RES, TST, TAIL, HED, BTS, LIVE, MOD, GAI, BT, KR, EVT, TC, WK, MC, KE, ASM = (W[n] for n in names[:17])
    RM, RO, RP, RE, RS, RED, ROH, RLS, RMOD, RB, RAL, RGA, RKM, RKO, RREL, RFE = (W[n] for n in names[17:])
    if rel is None:
        rel = pd.DataFrame(columns=["ticker", "report_date", "accession", "filing_date", "kw_up", "kw_down", "guidance", "tone", "risks", "summary"])
    al = pd.read_csv(_p(F_AI_LIVE)) if os.path.exists(_p(F_AI_LIVE)) else pd.DataFrame()
    ga = pd.read_csv(_p(F_AI_PARSE)) if os.path.exists(_p(F_AI_PARSE)) else pd.DataFrame()
    if ls is None:
        ls = pd.DataFrame(columns=["slug", "ticker", "ric", "entry_date", "exit_date", "iv_entry", "iv_exit"])
    if oh is None:
        oh = pd.DataFrame(columns=OH_COLS)
    if ed is None:
        ed = pd.DataFrame(columns=["ticker", "report_date", "cik", "accession", "filing_date", "acceptance"])

    # ---------- Assumptions: blue inputs that the formulas use ----------
    title(ASM, "Assumptions: inputs (blue), sources and method",
          "Every formula in the workbook takes its thresholds from the blue cells below. Change a blue cell and the formula "
          "sheets update. The Python steps (Raw_Models, Raw_Bootstrap, row pointers) used the values shown when the script ran.", "A2:C2")
    head(ASM, 3, ["Input", "Value", "Used on"])
    inputs = [("HP", "Confident crowd: p at least", HIGH_P, "0.00", "Results, Tail, Hedge, Live, Events"),
              ("FALL", "A fall: abnormal return at or below", FALL, "0%", "Tail, Events"),
              ("BIG", "A large fall: abnormal return at or below", BIG_FALL, "0%", "Tail, Events"),
              ("OTM", "Protective put: strike this far below the last close before the report", PUT_OTM, "0%", "Events, Hedge"),
              ("COST", "Put price used when there are no snapshot prices (share of the share price)", DEFAULT_PUT_COST, "0.00%", "Hedge"),
              ("FRZ", "Part 1 data run: reports after this date are new (out of sample)", pd.Timestamp(FREEZE_DATE).to_pydatetime(), "yyyy-mm-dd", "Events"),
              ("CUT", "Odds read at (UTC hour, report day)", CUTOFF_HOUR_UTC, "0", "Events, Kalshi_Events"),
              ("LEAD", "Second reading of the odds, days earlier", LEAD_DAYS, "0", "Events"),
              ("MINB", "Earlier events needed for the running base rate", MIN_PRIOR_EVENTS, "0", "Events"),
              ("MINH", "Past reports needed for a company's beat rate", MIN_HISTORY, "0", "Events"),
              ("MINM", "Past reports needed for a company's average move", MIN_PAST_MOVES, "0", "Events, Live"),
              ("RUN", "Trading days for the run-up before the report", RUNUP_DAYS, "0", "Events"),
              ("VOL", "Trading days for the volatility before the report", VOL_DAYS, "0", "Events"),
              ("MKT", "Market index for abnormal returns (column name on Raw_Prices)", MARKET, None, "Events, Raw_Estimates"),
              ("MKC", "Column of the market index on Raw_Prices (formula)", None, "0", "Events, Raw_Estimates, Kalshi_Events"),
              ("KMINB", "Earlier Kalshi events needed for the Kalshi base rate", K_MIN_PRIOR_EVENTS, "0", "Kalshi_Events"),
              ("CL1", "Confidence level 1", CONF_LEVELS[0], "0%", "Tests, Summary"),
              ("CL2", "Confidence level 2", CONF_LEVELS[1], "0%", "Tests, Summary"),
              ("CL3", "Confidence level 3", CONF_LEVELS[2], "0%", "Tests, Summary"),
              ("MINW", "Report weeks needed to cluster standard errors by week", MIN_WEEKS_CLUSTER, "0", "Tests"),
              ("IVD", "Maturity of the LSEG implied volatility (days)", IV_DAYS, "0", "Raw_LSEG"),
              ("TRIM", "Backtest: share of the position sold before the report when a policy trims", TRIM, "0%", "Events, Backtest"),
              ("TCOST", "Backtest: round-trip trading cost per unit traded", TCOST, "0.00%", "Events, Backtest"),
              ("DRIFT", "Crowd lost confidence: p fell over the week by at least (Part 3)", DRIFT, "0.00", "Events, Crowd_Adds"),
              ("QLOW", "AI downside model: quantile predicted (1-in-10 worst case). Change only with a rerun", Q_LOW, "0.00", "Model_Calc"),
              ("AM0", "Normal release window before the open: from (New York time)", 6 / 24, "hh:mm", "Raw_EDGAR"),
              ("AM1", "Normal release window before the open: to", 9.5 / 24, "hh:mm", "Raw_EDGAR"),
              ("PM0", "Normal release window after the close: from", 16 / 24, "hh:mm", "Raw_EDGAR"),
              ("PM1", "Normal release window after the close: to", 17.6 / 24, "hh:mm", "Raw_EDGAR")] + [
              (f"TH{i + 1}", f"Threshold sweep {i + 1}: crowd p at least", th, "0.00", "Tests") for i, th in enumerate(GAP_THRESHOLDS)]
    A = {}
    for i, (key, lab, v, nf, used) in enumerate(inputs):
        r = 4 + i
        put(ASM, r, 1, lab); put(ASM, r, 3, used)
        A[key] = f"Assumptions!$B${r}"
        if key != "MKC":
            put(ASM, r, 2, v, nf, BLUE)
    PE = 4 + max(len(px), 1)
    LP = L(1 + len(px.columns))
    put(ASM, 4 + [k for k, *_ in inputs].index("MKC"), 2, f"=MATCH({A['MKT']},Raw_Prices!$A$4:${LP}$4,0)", "0")
    r = 4 + len(inputs) + 1
    head(ASM, r, ["Source", "What", "Downloaded (UTC)"]); r += 1
    log = open(_p(F_LOG)).read().strip().splitlines()[-1] if os.path.exists(_p(F_LOG)) else download_utc
    for a, b in [("Polymarket Gamma API", f"{GAMMA}/events?tag_slug=earnings: contracts, outcomes, current prices"),
                 ("Polymarket CLOB API", f"{CLOB}/prices-history (hourly YES prices before each report), /midpoint (snapshots)"),
                 ("Kalshi Trade API v2", f"{KALSHI}: KPI series, settled markets, hourly candlesticks"),
                 ("LSEG Workspace (Excel add-in, RDP.Data)", f"Implied volatility the day before and after each report, field {LSEG_FIELD}. Consensus EPS ({LSEG_EPS_FIELD}, "
                                                             f"{LSEG_EPS_PARAMS}) the day before and {REV_DAYS} days earlier"),
                 ("SEC EDGAR", "data.sec.gov/submissions: 8-K item 2.02 acceptance times (release time check)"),
                 ("Massive, formerly Polygon.io (optional)", "Free Options Basic plan: past option prices. Off in this run "
                                                             "(USE_MASSIVE). Same use as LSEG, slower: 5 requests a minute"),
                 ("Yahoo Finance via yfinance", "Daily adjusted close, EPS estimates and reported EPS. Snapshots: option chains, "
                                                "eps_trend, eps_revisions, earnings_estimate")]:
        put(ASM, r, 1, a); put(ASM, r, 2, b); put(ASM, r, 3, log); r += 1
    r += 1
    for n_ in ["Abnormal return = stock log return minus market log return, from the close before the report day to the close after it.",
               "Put payoff = max(0, (1 - strike gap) - S1/S0), measured at the close after the report. It ignores the option's remaining "
               "time value, so it understates what the put is worth after a fall.",
               "Snapshot used for a report = the last daily snapshot taken before the odds cut-off on the report day (row set by the script).",
               "Raw_Models: out-of-sample forecasts. Each month is forecast by models trained only on earlier months (at least 300 reports).",
               "Raw_Bootstrap: each draw resamples whole report weeks, so reports in the same earnings week stay together.",
               "Kalshi: settled 'above X' KPI contracts, one per KPI event (the most traded strike), merged across runs.",
               "Results describe one sample and are associations, not causes."]:
        put(ASM, r, 1, "- " + n_); r += 1
    ASM.column_dimensions["A"].width = 70; ASM.column_dimensions["B"].width = 75; ASM.column_dimensions["C"].width = 32

    # ---------- Raw sheets ----------
    RM["A1"] = f"Polymarket earnings contracts as downloaded: {GAMMA}/events?tag_slug=earnings. Downloaded {download_utc} UTC."
    RM["A1"].font = BOLD
    mcols = [("slug", "Contract slug"), ("question", "Question"), ("ticker", "Ticker"), ("report_date", "Report date"),
             ("threshold", "EPS threshold"), ("eps_type", "EPS type"), ("volume", "Volume (USD)"), ("created", "Listed (UTC)"),
             ("outcome", "Outcome (1 = beat, 0 = miss)"), ("p_now", "Last price"), ("token", "CLOB token (YES)"), ("sym", "Yahoo symbol")]
    head(RM, 4, [h for _, h in mcols])
    mk = mk.reset_index(drop=True).copy()
    mk["sym"] = mk["ticker"].map(lambda t: yf_symbol(t) if isinstance(t, str) else None)
    for i, rec in enumerate(mk.to_dict("records")):
        for j, (k, _) in enumerate(mcols):
            put(RM, 5 + i, 1 + j, clean(rec.get(k)), "yyyy-mm-dd" if k == "report_date" else None)
    mrow = {s: 5 + i for i, s in enumerate(mk["slug"])}

    RO["A1"] = f"Hourly YES prices as downloaded: {CLOB}/prices-history, last {ODDS_WINDOW_DAYS} days before the cut-off."; RO["A1"].font = BOLD
    head(RO, 4, ["Contract slug", "Time (unix, UTC)", "Time (UTC)", "YES price"])
    od = od[od["slug"].isin(set(ev["slug"]))].sort_values(["slug", "ts"])
    span, rr = {}, 5
    for slug, g in od.groupby("slug", sort=False):
        a = rr
        for t, p in zip(g["ts"], g["p"]):
            RO.cell(row=rr, column=1, value=slug); RO.cell(row=rr, column=2, value=int(t))
            RO.cell(row=rr, column=3, value=datetime.fromtimestamp(int(t), timezone.utc).replace(tzinfo=None)).number_format = "yyyy-mm-dd hh:mm"
            RO.cell(row=rr, column=4, value=float(p)); rr += 1
        span[slug] = (a, rr - 1)

    RP["A1"] = f"Daily adjusted close as downloaded: Yahoo Finance via yfinance, from {HIST_START}. Downloaded {download_utc} UTC."; RP["A1"].font = BOLD
    head(RP, 4, ["Date"] + list(px.columns), h=20)
    for i, (d, row) in enumerate(px.iterrows()):
        RP.cell(row=5 + i, column=1, value=d.to_pydatetime()).number_format = "yyyy-mm-dd"
        for j, v in enumerate(row.to_numpy()):
            if np.isfinite(v):
                RP.cell(row=5 + i, column=2 + j, value=float(v))
    DATES = f"Raw_Prices!$A$5:$A${PE}"
    PX = f"Raw_Prices!$A$5:${LP}${PE}"
    HDR = f"Raw_Prices!$A$4:${LP}$4"
    MC_ = A["MKC"]

    RE["A1"] = f"EPS estimates and reported EPS as downloaded: Yahoo Finance (get_earnings_dates). Columns F to I are formulas."; RE["A1"].font = BOLD
    head(RE, 4, ["Ticker", "Report date", "EPS estimate", "Reported EPS", "Surprise (%) as published", "Beat (1/0)",
                 "Yahoo symbol", "Price column on Raw_Prices", "|Abnormal move| on that report"])
    EE = 4 + max(len(est), 1)
    for i, rec in enumerate(est.to_dict("records")):
        q = 5 + i
        put(RE, q, 1, rec["ticker"]); put(RE, q, 2, clean(rec["date"]), "yyyy-mm-dd")
        put(RE, q, 3, clean(rec["estimate"])); put(RE, q, 4, clean(rec["reported"])); put(RE, q, 5, clean(rec["surprise_pct"]))
        put(RE, q, 6, f"=IF(D{q}>C{q},1,0)")
        put(RE, q, 7, yf_symbol(rec["ticker"]))
        put(RE, q, 8, f'=IFERROR(MATCH(G{q},{HDR},0),"")')
        a_, b_ = f"MATCH(B{q}-1,{DATES},1)", f"MATCH(B{q},{DATES},1)+1"
        put(RE, q, 9, f'=IFERROR(ABS(LN(INDEX({PX},{b_},H{q})/INDEX({PX},{a_},H{q}))-LN(INDEX({PX},{b_},{MC_})/INDEX({PX},{a_},{MC_}))),"")', PC1)
    RE_ = lambda c: f"Raw_Estimates!${c}$5:${c}${EE}"

    RS["A1"] = "Daily snapshots as saved by the script: Polymarket price, Yahoo option chain (first expiry after the report) and analyst estimates."
    RS["A1"].font = BOLD
    head(RS, 4, SNAP_COLS)
    SE = 4 + max(len(snap), 1)
    for i, rec in enumerate(snap[SNAP_COLS].to_dict("records")):
        for j, k in enumerate(SNAP_COLS):
            v = clean(rec[k])
            if k == "report_date" and v:
                v = pd.Timestamp(v).to_pydatetime()
            put(RS, 5 + i, 1 + j, v, "yyyy-mm-dd" if k == "report_date" else None)
    SCOL = {k: L(1 + j) for j, k in enumerate(SNAP_COLS)}
    SN = lambda k: f"Raw_Snapshots!${SCOL[k]}$5:${SCOL[k]}${SE}"

    RED["A1"] = ("SEC EDGAR: 8-K with item 2.02 (results of operations) filed around each report. Columns A to F as downloaded from "
                 "data.sec.gov/submissions. Columns G to L are formulas.")
    RED["A1"].font = BOLD
    RED["A2"] = ("EDGAR's acceptanceDateTime is labelled UTC, but for some filings it is shifted by the New York offset (Apple, Netflix, "
                 "JPMorgan, P&G) and for others it is not (Coca-Cola, Tesla). Each time is read both ways (columns I and J). The reading "
                 "that falls in a normal release window (Assumptions) is kept. If neither does, the release time is left empty.")
    RED["A2"].alignment = Alignment(wrap_text=True); RED.merge_cells("A2:L2"); RED.row_dimensions[2].height = 45
    head(RED, 4, ["Ticker", "Report date", "CIK", "Accession", "Filing date", "acceptanceDateTime as published",
                  "Shown time", "New York offset (hours, summer 4 / winter 5)", "New York time if the shown time is UTC",
                  "New York time if the shown time is shifted", "First reading in a release window (1/0)",
                  "Second reading in a release window (1/0)", "Release (New York)", "Release (UTC)", "Session", "Release (unix)"], h=60)
    EDE = 4 + max(len(ed), 1)
    for i, rec in enumerate(ed.to_dict("records")):
        q = 5 + i
        put(RED, q, 1, rec["ticker"]); put(RED, q, 2, pd.Timestamp(rec["report_date"]).to_pydatetime(), "yyyy-mm-dd")
        put(RED, q, 3, clean(rec["cik"])); put(RED, q, 4, rec["accession"]); put(RED, q, 5, rec["filing_date"]); put(RED, q, 6, rec["acceptance"])
        put(RED, q, 7, f"=DATE(VALUE(LEFT(F{q},4)),VALUE(MID(F{q},6,2)),VALUE(MID(F{q},9,2)))"
                       f"+TIME(VALUE(MID(F{q},12,2)),VALUE(MID(F{q},15,2)),VALUE(MID(F{q},18,2)))", "yyyy-mm-dd hh:mm")
        dst0 = f"DATE(YEAR(G{q}),3,8)+MOD(8-WEEKDAY(DATE(YEAR(G{q}),3,8)),7)"
        dst1 = f"DATE(YEAR(G{q}),11,1)+MOD(8-WEEKDAY(DATE(YEAR(G{q}),11,1)),7)"
        put(RED, q, 8, f"=IF(AND(G{q}>={dst0},G{q}<{dst1}),4,5)")
        put(RED, q, 9, f"=G{q}-H{q}/24", "yyyy-mm-dd hh:mm"); put(RED, q, 10, f"=G{q}-2*H{q}/24", "yyyy-mm-dd hh:mm")
        win = lambda c: (f'IF(OR(AND(MOD({c}{q},1)>={A["AM0"]},MOD({c}{q},1)<{A["AM1"]}),'
                         f'AND(MOD({c}{q},1)>={A["PM0"]},MOD({c}{q},1)<{A["PM1"]})),1,0)')
        put(RED, q, 11, "=" + win("I")); put(RED, q, 12, "=" + win("J"))
        put(RED, q, 13, f'=IF(K{q}=1,I{q},IF(L{q}=1,J{q},""))', "yyyy-mm-dd hh:mm")
        put(RED, q, 14, f'=IF(ISNUMBER(M{q}),M{q}+H{q}/24,"")', "yyyy-mm-dd hh:mm")
        put(RED, q, 15, f'=IF(ISNUMBER(M{q}),IF(MOD(M{q},1)<{A["AM1"]},"before the open","after the close"),"unclear")')
        put(RED, q, 16, f'=IF(ISNUMBER(N{q}),(N{q}-DATE(1970,1,1))*86400,"")', "0")
    edrow = {(str(t), pd.Timestamp(d).strftime("%Y-%m-%d")): i + 1 for i, (t, d) in enumerate(zip(ed["ticker"], ed["report_date"]))}
    ED_ = lambda c: f"Raw_EDGAR!${c}$5:${c}${EDE}"

    ROH["A1"] = ("Past option prices from Massive (formerly Polygon.io), free Options Basic plan: daily closing prices of the put 5% below "
                 "and the at-the-money put on the first expiry after the report. Columns A to S as downloaded. Columns T to Y are formulas.")
    ROH["A1"].font = BOLD
    ROH["A2"] = ("Entry = close the day before the report day, exit = close the day after (same window as the abnormal return). "
                 "Closing trade prices, not bid/ask. A price is used only if the option traded on the entry day.")
    ROH["A2"].alignment = Alignment(wrap_text=True); ROH.merge_cells("A2:L2"); ROH.row_dimensions[2].height = 35
    head(ROH, 4, OH_COLS + ["Entry prices current (1/0)", "Put price / share price", "Put value after / share price (intrinsic if no trade)",
                            "Hedge P&L / share price", "Implied move (2 x ATM put / share price)", "Exit source"], h=60)
    OHE = 4 + max(len(oh), 1)
    OHC = {k: L(1 + j) for j, k in enumerate(OH_COLS)}
    for i, rec in enumerate(oh.to_dict("records")):
        q = 5 + i
        for j, k in enumerate(OH_COLS):
            v = clean(rec[k])
            if k in ("report_date", "entry_date", "exit_date", "expiry", "atm_entry_date", "put_entry_date", "put_exit_date") and v:
                v = pd.Timestamp(v).to_pydatetime()
            put(ROH, q, 1 + j, v, "yyyy-mm-dd" if "date" in k or k == "expiry" else None)
        g = lambda k: f"{OHC[k]}{q}"
        put(ROH, q, 20, f'=IF(AND(ISNUMBER({g("put_entry")}),{g("put_entry_date")}={g("entry_date")}),1,0)')
        put(ROH, q, 21, f'=IF(T{q}=1,{g("put_entry")}/{g("spot")},"")', PC2)
        put(ROH, q, 22, f'=IF(T{q}=1,IF(ISNUMBER({g("put_exit")}),{g("put_exit")},MAX(0,{g("put_strike")}-{g("spot_after")}))/{g("spot")},"")', PC2)
        put(ROH, q, 23, f'=IF(T{q}=1,V{q}-U{q},"")', "+0.00%;-0.00%")
        put(ROH, q, 24, f'=IF(AND(ISNUMBER({g("atm_entry")}),{g("atm_entry_date")}={g("entry_date")}),2*{g("atm_entry")}/{g("spot")},"")', PC1)
        put(ROH, q, 25, f'=IF(T{q}=1,IF(ISNUMBER({g("put_exit")}),"traded","intrinsic"),"")')
    ohrow = {s_: i + 1 for i, s_ in enumerate(oh["slug"])}
    OH_ = lambda c: f"Raw_OptHist!${c}$5:${c}${OHE}"

    RLS["A1"] = ("LSEG Workspace: implied volatility at the close before the report day (entry) and the close after it (exit), as "
                 "returned by RDP.Data in the request file, and the I/B/E/S consensus EPS the day before the report and about "
                 f"{REV_DAYS} days earlier, the implied volatility on that earlier date and the StarMine SmartEstimate. Columns A to G, "
                 "K to M and O to P as retrieved. Columns H to J, N and Q to S are formulas.")
    RLS["A1"].font = BOLD
    RLS["A2"] = ("Implied event move = square root of (IV before squared - IV after squared) x maturity / 365. The drop in implied "
                 "volatility after the report measures the move the options market priced for the report itself.")
    RLS["A2"].alignment = Alignment(wrap_text=True); RLS.merge_cells("A2:J2"); RLS.row_dimensions[2].height = 35
    head(RLS, 4, ["Contract", "Ticker", "RIC", "Entry date", "Exit date", "IV before (as returned)", "IV after (as returned)",
                  "IV before (decimal)", "IV after (decimal)", "Implied event move", "Earlier date", "Consensus EPS before",
                  "Consensus EPS earlier", "Analyst revision (EPS before / earlier - 1)", "IV earlier (as returned)",
                  "StarMine SmartEstimate before", "IV earlier (decimal)", "Implied move before the report (no look-ahead)",
                  "StarMine predicted surprise (SmartEstimate / consensus - 1)", "Call IV before (as returned)", "Call IV before (decimal)",
                  "Skew: put minus call at-the-money implied volatility"], h=45)
    LSE = 4 + max(len(ls), 1)
    for i, rec in enumerate(ls.to_dict("records")):
        q = 5 + i
        for j, k in enumerate(("slug", "ticker", "ric", "entry_date", "exit_date", "iv_entry", "iv_exit")):
            v = clean(rec[k])
            if k in ("entry_date", "exit_date") and v:
                v = pd.Timestamp(v).to_pydatetime()
            put(RLS, q, 1 + j, v, "yyyy-mm-dd" if "date" in k else None)
        put(RLS, q, 8, f'=IF(ISNUMBER(F{q}),IF(F{q}>3,F{q}/100,F{q}),"")', PC1)
        put(RLS, q, 9, f'=IF(ISNUMBER(G{q}),IF(G{q}>3,G{q}/100,G{q}),"")', PC1)
        put(RLS, q, 10, f'=IF(AND(ISNUMBER(H{q}),ISNUMBER(I{q})),SQRT(MAX(0,H{q}^2-I{q}^2)*{A["IVD"]}/365),"")', PC1)
        v = clean(rec.get("earlier_date"))
        put(RLS, q, 11, pd.Timestamp(v).to_pydatetime() if v else None, "yyyy-mm-dd")
        put(RLS, q, 12, clean(rec.get("eps_entry")), "0.00"); put(RLS, q, 13, clean(rec.get("eps_earlier")), "0.00")
        put(RLS, q, 14, f'=IF(AND(ISNUMBER(L{q}),ISNUMBER(M{q})),IF(M{q}<>0,(L{q}-M{q})/ABS(M{q}),""),"")', "+0.0%;-0.0%")
        put(RLS, q, 15, clean(rec.get("iv_earlier"))); put(RLS, q, 16, clean(rec.get("smart_est")), "0.00")
        put(RLS, q, 17, f'=IF(ISNUMBER(O{q}),IF(O{q}>3,O{q}/100,O{q}),"")', PC1)
        put(RLS, q, 18, f'=IF(AND(ISNUMBER(H{q}),ISNUMBER(Q{q})),SQRT(MAX(0,H{q}^2-Q{q}^2)*{A["IVD"]}/365),"")', PC1)
        put(RLS, q, 19, f'=IF(AND(ISNUMBER(P{q}),ISNUMBER(L{q})),IF(L{q}<>0,(P{q}-L{q})/ABS(L{q}),""),"")', "+0.0%;-0.0%")
        put(RLS, q, 20, clean(rec.get("iv_call")))
        put(RLS, q, 21, f'=IF(ISNUMBER(T{q}),IF(T{q}>3,T{q}/100,T{q}),"")', PC1)
        put(RLS, q, 22, f'=IF(AND(ISNUMBER(H{q}),ISNUMBER(U{q})),H{q}-U{q},"")', "+0.0%;-0.0%")
    lsrow = {s_: i + 1 for i, s_ in enumerate(ls["slug"])}

    RREL["A1"] = ("Last quarter's earnings release for each report: the SEC 8-K (item 2.02) filed 20 to 130 days before, its press-release "
                  "exhibit, a keyword count (the spreadsheet way) and the generative AI reading (guidance, tone, risks, one line). "
                  "Columns A to K as retrieved. Column L is a formula."); RREL["A1"].font = BOLD
    head(RREL, 4, ["Key (ticker|report date)", "Ticker", "Report date", "Release accession (EDGAR)", "Release filed", "Keywords up",
                   "Keywords down", "Guidance (generative AI)", "Tone -2 to 2 (generative AI)", "Risks flagged 0 to 3 (generative AI)",
                   "Release in one line (generative AI)", "Keyword balance (up - down) / (up + down)"], h=45)
    rl = rel.dropna(subset=["accession"]).reset_index(drop=True) if len(rel) else rel
    for i, rec in enumerate(rl.to_dict("records")):
        q = 5 + i
        put(RREL, q, 1, f"{rec['ticker']}|{str(rec['report_date'])[:10]}"); put(RREL, q, 2, rec["ticker"])
        put(RREL, q, 3, pd.Timestamp(rec["report_date"]).to_pydatetime(), "yyyy-mm-dd"); put(RREL, q, 4, clean(rec["accession"]))
        put(RREL, q, 5, clean(rec["filing_date"]))
        for j, k in enumerate(("kw_up", "kw_down", "guidance", "tone", "risks", "summary")):
            put(RREL, q, 6 + j, clean(rec[k]))
        put(RREL, q, 12, f'=IF(AND(ISNUMBER(F{q}),ISNUMBER(G{q})),IF(F{q}+G{q}>0,(F{q}-G{q})/(F{q}+G{q}),""),"")', "+0.00;-0.00")
    REE = 4 + max(len(rl), 1)
    RL_ = lambda c_: f"Raw_Releases!${c_}$5:${c_}${REE}"

    RFE["A1"] = ("Model inputs computed by the script (same row order as Events): how the odds moved in the last week before the cut-off, "
                 f"and same-sector peers that reported in the previous {PEER_DAYS} days (Yahoo sector). Inputs to Raw_Models."); RFE["A1"].font = BOLD
    fcols = [("slug", "Contract slug"), ("odds_vol", "Odds: SD of hourly changes, last 7 days"), ("odds_range", "Odds: high - low, last 7 days"),
             ("late24", "Odds: change in the last 24 hours"), ("peer_n", "Peers that reported"), ("peer_mean", "Peers: average abnormal move"),
             ("peer_fall", f"Peers: share falling {abs(FALL):.0%} or more"),
             ("k_kpi_p", "Kalshi KPI contracts: average chance above the level"), ("k_kpi_n", "Kalshi KPI contracts around the report")]
    head(RFE, 4, [h for _, h in fcols], h=45)
    for i, rec in enumerate(ev.reindex(columns=[k for k, _ in fcols]).to_dict("records")):
        for j, (k, _) in enumerate(fcols):
            put(RFE, 5 + i, 1 + j, clean(rec[k]), None if k in ("slug", "peer_n") else "0.000")
    LS_ = lambda c: f"Raw_LSEG!${c}$5:${c}${LSE}"

    # ---------- Events ----------
    title(EVT, "Events: one row per settled report. Every cell is a formula on the Raw_ sheets, except the row pointers set by the script.",
          "p = YES price at the cut-off hour on the report day. Reaction = close before the report day to close after it. "
          "Abnormal = stock minus market log return. Pointers: Raw_Markets row, first and last Raw_Odds row, Raw_Snapshots row.", "A2:T2")
    ec = ["Raw_Markets row", "Ticker", "Report date", "EPS type", "Volume (USD)", "Outcome y (1 = beat)", "Odds first row", "Odds last row",
          "Cut-off (unix)", "p: P(beat) at cut-off", "p 24 hours earlier", "p days earlier (Assumptions)", "Sample", "Confident (1/0)",
          "Price column", "Row: close before", "Row: close after", "S0: close before", "S1: close after", "Market before", "Market after",
          "Stock log return", "Market log return", "Abnormal return", "Share return (simple)", "Put payoff (share of S0)",
          "Run-up before (abnormal)", "Volatility before (SD of daily abnormal returns)", "Company past beat rate",
          "Company's average past |move|", "Analyst surprise (%), capped", "Running base rate", "Brier: crowd", "Brier: base rate",
          "Brier: company record", "All three forecasts (1/0)", "Change in p over the week", "Fall (1/0)", "Large fall (1/0)",
          "Snapshot row", "Snapshot time (UTC)", "Option-implied move (snapshot)", "Put price / share price (snapshot)",
          "EPS estimate change, 30 days (snapshot)", "Net revisions up / analysts, 30 days (snapshot)", "Estimate range / estimate (snapshot)",
          "AR if confident", "AR if less sure", "AR if confident, new", "AR if less sure, new", "AR if confident, Part 1",
          "AR if less sure, Part 1", "Share return if confident", "AR: confident miss", "AR: other miss", "AR: confident miss, new",
          "AR: other miss, new", "AR: confident miss, Part 1", "AR: other miss, Part 1", "Week (Monday)", "EDGAR row",
          "Release (New York time)", "Release session", "Odds read before the release (1/0)", "Hedge P&L: payoff - snapshot put price",
          "Option-implied move - realised |share return|", "Option history row", "Put price / share price (history)",
          "Put value after / share price (history)", "Hedge P&L with actual option prices (history)", "Implied move (history)",
          "Implied move (history) - realised |share return|", "LSEG row", "IV before (LSEG)", "IV after (LSEG)",
          "Implied event move (LSEG)", "Implied event move (LSEG) - realised |share return|", "Analyst revision (LSEG)",
          "Revision up (1/0)", "Put price / share price (LSEG IV)", "Put value after / share price (LSEG IV)",
          "Hedge P&L with LSEG-priced put", "Implied move before the report (LSEG)", "StarMine predicted surprise (LSEG)",
          "AI 10% worst case (Raw_Models)", "In backtest (1/0)", "Rule trims (1/0)", "AI trims (1/0)", "Outcome: hold",
          "Outcome: rule trims", "Outcome: AI trims", "Outcome: rule buys the put", "Outcome: AI buys the put",
          "Past-move rule trims (1/0)", "Implied-move rule trims (1/0)", "Expected-loss rule trims (1/0)", "Outcome: past-move rule",
          "Outcome: implied-move rule", "Outcome: expected-loss rule", "Last release: guidance (generative AI)",
          "Last release: tone (generative AI)", "Last release: keyword balance", "Skew (LSEG put - call IV)"] + \
         [x for _, lf, lo, _s in NEW_POLICIES for x in (lf, lo)] + CA_COLS + ["Has 3 past reports (1/0)"] + \
         [f"{oc} | {sl}" for oc in BT_SUB for sl in BT_SETS]
    head(EVT, 4, ec, h=60)
    C = {h: L(1 + j) for j, h in enumerate(ec)}
    n = len(ev)
    E1 = 4 + max(n, 1)
    sp_ptr = snapshot_pointer(ev, snap)
    for i, rec in enumerate(ev.to_dict("records")):
        r = 5 + i
        raw = mrow[rec["slug"]]
        a, b = span[rec["slug"]]
        O = lambda h: f"{C[h]}{r}"
        f = {
            "Raw_Markets row": raw, "Ticker": f"=Raw_Markets!C{raw}", "Report date": f"=Raw_Markets!D{raw}",
            "EPS type": f"=Raw_Markets!F{raw}", "Volume (USD)": f"=Raw_Markets!G{raw}", "Outcome y (1 = beat)": f"=Raw_Markets!I{raw}",
            "Odds first row": a, "Odds last row": b,
            "Cut-off (unix)": f"=({O('Report date')}-DATE(1970,1,1))*86400+{A['CUT']}*3600",
            "p: P(beat) at cut-off": f'=IFERROR(INDEX(Raw_Odds!$D${a}:$D${b},MATCH({O("Cut-off (unix)")},Raw_Odds!$B${a}:$B${b},1)),"")',
            "p 24 hours earlier": f'=IFERROR(INDEX(Raw_Odds!$D${a}:$D${b},MATCH({O("Cut-off (unix)")}-86400,Raw_Odds!$B${a}:$B${b},1)),"")',
            "p days earlier (Assumptions)": f'=IFERROR(INDEX(Raw_Odds!$D${a}:$D${b},MATCH({O("Cut-off (unix)")}-{A["LEAD"]}*86400,Raw_Odds!$B${a}:$B${b},1)),"")',
            "Sample": f'=IF({O("Report date")}>{A["FRZ"]},"New","Part 1")',
            "Confident (1/0)": f'=IF(ISNUMBER({O("p: P(beat) at cut-off")}),IF({O("p: P(beat) at cut-off")}>={A["HP"]},1,0),"")',
            "Price column": f'=IFERROR(MATCH(Raw_Markets!L{raw},{HDR},0),"")',
            "Row: close before": f'=IFERROR(MATCH({O("Report date")}-1,{DATES},1),"")',
            "Row: close after": f'=IFERROR(MATCH({O("Report date")},{DATES},1)+1,"")',
        }
        q0, q1, pc = O("Row: close before"), O("Row: close after"), O("Price column")
        pos = lambda rowx, colx: f'=IFERROR(IF(INDEX({PX},{rowx},{colx})>0,INDEX({PX},{rowx},{colx}),""),"")'
        f["S0: close before"] = pos(q0, pc); f["S1: close after"] = pos(q1, pc)
        f["Market before"] = pos(q0, MC_); f["Market after"] = pos(q1, MC_)
        s0, s1, m0, m1 = O("S0: close before"), O("S1: close after"), O("Market before"), O("Market after")
        f["Stock log return"] = f'=IFERROR(LN({s1}/{s0}),"")'
        f["Market log return"] = f'=IFERROR(LN({m1}/{m0}),"")'
        sl, ml, arc = O("Stock log return"), O("Market log return"), O("Abnormal return")
        f["Abnormal return"] = f'=IF(AND(ISNUMBER({sl}),ISNUMBER({ml})),{sl}-{ml},"")'
        f["Share return (simple)"] = f'=IF(ISNUMBER({sl}),{s1}/{s0}-1,"")'
        f["Put payoff (share of S0)"] = f'=IF(ISNUMBER({O("Share return (simple)")}),MAX(0,(1-{A["OTM"]})-(1+{O("Share return (simple)")})),"")'
        f["Run-up before (abnormal)"] = (f'=IFERROR(LN(INDEX({PX},{q0},{pc})/INDEX({PX},{q0}-{A["RUN"]},{pc}))'
                                         f'-LN(INDEX({PX},{q0},{MC_})/INDEX({PX},{q0}-{A["RUN"]},{MC_})),"")')
        rng_ = lambda c, k0, k1: f"INDEX({PX},{q0}{k0},{c}):INDEX({PX},{q0}{k1},{c})"
        vol_txt = (f'=IFERROR(STDEV(LN({rng_(pc, "-" + A["VOL"] + "+1", "")}/{rng_(pc, "-" + A["VOL"], "-1")})'
                   f'-LN({rng_(MC_, "-" + A["VOL"] + "+1", "")}/{rng_(MC_, "-" + A["VOL"], "-1")})),"")')
        tk, rd_ = O("Ticker"), O("Report date")
        f["Company past beat rate"] = (f'=IF(COUNTIFS({RE_("A")},{tk},{RE_("B")},"<"&{rd_}-3)>={A["MINH"]},'
                                       f'AVERAGEIFS({RE_("F")},{RE_("A")},{tk},{RE_("B")},"<"&{rd_}-3),"")')
        f["Company's average past |move|"] = (f'=IF(COUNTIFS({RE_("A")},{tk},{RE_("B")},"<"&{rd_}-3,{RE_("I")},">=0")>={A["MINM"]},'
                                              f'AVERAGEIFS({RE_("I")},{RE_("A")},{tk},{RE_("B")},"<"&{rd_}-3),"")')
        f["Analyst surprise (%), capped"] = (f'=IFERROR(MAX(-100,MIN(100,AVERAGEIFS({RE_("E")},{RE_("A")},{tk},{RE_("B")},">="&{rd_}-3,'
                                             f'{RE_("B")},"<="&{rd_}+3))),"")')
        CD, CY = f"${C['Report date']}$5:${C['Report date']}${E1}", f"${C['Outcome y (1 = beat)']}$5:${C['Outcome y (1 = beat)']}${E1}"
        f["Running base rate"] = f'=IF(COUNTIF({CD},"<"&{rd_})>={A["MINB"]},AVERAGEIFS({CY},{CD},"<"&{rd_}),"")'
        p_, y_ = O("p: P(beat) at cut-off"), O("Outcome y (1 = beat)")
        br, cr = O("Running base rate"), O("Company past beat rate")
        f["Brier: crowd"] = f'=IF(ISNUMBER({p_}),({p_}-{y_})^2,"")'
        f["Brier: base rate"] = f'=IF(ISNUMBER({br}),({br}-{y_})^2,"")'
        f["Brier: company record"] = f'=IF(ISNUMBER({cr}),({cr}-{y_})^2,"")'
        f["All three forecasts (1/0)"] = (f'=IF(AND(ISNUMBER({O("Brier: crowd")}),ISNUMBER({O("Brier: base rate")}),'
                                          f'ISNUMBER({O("Brier: company record")})),1,0)')
        f["Change in p over the week"] = f'=IF(AND(ISNUMBER({p_}),ISNUMBER({O("p days earlier (Assumptions)")})),{p_}-{O("p days earlier (Assumptions)")},"")'
        f["Fall (1/0)"] = f'=IF(ISNUMBER({arc}),IF({arc}<={A["FALL"]},1,0),"")'
        f["Large fall (1/0)"] = f'=IF(ISNUMBER({arc}),IF({arc}<={A["BIG"]},1,0),"")'
        f["Snapshot row"] = sp_ptr[i]
        sr = O("Snapshot row")
        sx = lambda k: f"INDEX({SN(k)},{sr})"
        has = lambda *ks: f"COUNT({','.join(sx(k) for k in ks)})={len(ks)}"
        okput = f'AND({has("put95_cost", "put95_strike", "spot")},{sx("put95_strike")}<={sx("spot")})'
        f["Snapshot time (UTC)"] = f'=IF(ISNUMBER({sr}),{sx("snap_utc")},"")'
        f["Option-implied move (snapshot)"] = f'=IF(ISNUMBER({sr}),IF({has("implied_move")},{sx("implied_move")},""),"")'
        f["Put price / share price (snapshot)"] = f'=IF(ISNUMBER({sr}),IF({okput},{sx("put95_cost")},""),"")'
        f["EPS estimate change, 30 days (snapshot)"] = (f'=IFERROR(IF(AND(ISNUMBER({sr}),{has("eps_now", "eps_30d_ago")}),({sx("eps_now")}-{sx("eps_30d_ago")})'
                                                        f'/MAX(ABS({sx("eps_30d_ago")}),0.01),""),"")')
        f["Net revisions up / analysts, 30 days (snapshot)"] = (f'=IFERROR(IF(AND(ISNUMBER({sr}),{has("rev_up_30d", "rev_down_30d", "n_analysts")}),({sx("rev_up_30d")}-{sx("rev_down_30d")})'
                                                                f'/MAX({sx("n_analysts")},1),""),"")')
        f["Estimate range / estimate (snapshot)"] = (f'=IFERROR(IF(AND(ISNUMBER({sr}),{has("eps_high", "eps_low", "eps_avg")}),({sx("eps_high")}-{sx("eps_low")})'
                                                     f'/MAX(ABS({sx("eps_avg")}),0.01),""),"")')
        cf, smp, rr_ = O("Confident (1/0)"), O("Sample"), O("Share return (simple)")
        f["AR if confident"] = f'=IF(AND({cf}=1,ISNUMBER({arc})),{arc},"")'
        f["AR if less sure"] = f'=IF(AND({cf}=0,ISNUMBER({arc})),{arc},"")'
        f["AR if confident, new"] = f'=IF({smp}="New",{O("AR if confident")},"")'
        f["AR if less sure, new"] = f'=IF({smp}="New",{O("AR if less sure")},"")'
        f["AR if confident, Part 1"] = f'=IF({smp}="Part 1",{O("AR if confident")},"")'
        f["AR if less sure, Part 1"] = f'=IF({smp}="Part 1",{O("AR if less sure")},"")'
        f["Share return if confident"] = f'=IF(AND({cf}=1,ISNUMBER({rr_})),{rr_},"")'
        f["AR: confident miss"] = f'=IF({y_}=0,{O("AR if confident")},"")'
        f["AR: other miss"] = f'=IF({y_}=0,{O("AR if less sure")},"")'
        f["AR: confident miss, new"] = f'=IF({y_}=0,{O("AR if confident, new")},"")'
        f["AR: other miss, new"] = f'=IF({y_}=0,{O("AR if less sure, new")},"")'
        f["AR: confident miss, Part 1"] = f'=IF({y_}=0,{O("AR if confident, Part 1")},"")'
        f["AR: other miss, Part 1"] = f'=IF({y_}=0,{O("AR if less sure, Part 1")},"")'
        f["Week (Monday)"] = f'=IF(ISNUMBER({rd_}),{rd_}-WEEKDAY({rd_},2)+1,"")'
        f["EDGAR row"] = edrow.get((str(rec["ticker"]), pd.Timestamp(rec["report_date"]).strftime("%Y-%m-%d")))
        er = O("EDGAR row")
        f["Release (New York time)"] = f'=IF(ISNUMBER({er}),IF(COUNT(INDEX({ED_("M")},{er}))=1,INDEX({ED_("M")},{er}),""),"")'
        f["Release session"] = f'=IF(ISNUMBER({er}),INDEX({ED_("O")},{er}),"")'
        f["Odds read before the release (1/0)"] = (f'=IF(ISNUMBER({er}),IF(COUNT(INDEX({ED_("P")},{er}))=1,'
                                                   f'IF(INDEX({ED_("P")},{er})>{O("Cut-off (unix)")},1,0),""),"")')
        pp, pf = O("Put price / share price (snapshot)"), O("Put payoff (share of S0)")
        f["Hedge P&L: payoff - snapshot put price"] = f'=IF(AND(ISNUMBER({pf}),ISNUMBER({pp})),{pf}-{pp},"")'
        im = O("Option-implied move (snapshot)")
        f["Option-implied move - realised |share return|"] = f'=IF(AND(ISNUMBER({im}),ISNUMBER({rr_})),{im}-ABS({rr_}),"")'
        f["Option history row"] = ohrow.get(rec["slug"])
        oq = O("Option history row")
        ohx = lambda c, oq=oq: f'=IF(ISNUMBER({oq}),IF(COUNT(INDEX({OH_(c)},{oq}))=1,INDEX({OH_(c)},{oq}),""),"")'
        f["Put price / share price (history)"] = ohx("U")
        f["Put value after / share price (history)"] = ohx("V")
        f["Hedge P&L with actual option prices (history)"] = ohx("W")
        f["Implied move (history)"] = ohx("X")
        imh = O("Implied move (history)")
        f["Implied move (history) - realised |share return|"] = f'=IF(AND(ISNUMBER({imh}),ISNUMBER({rr_})),{imh}-ABS({rr_}),"")'
        f["LSEG row"] = lsrow.get(rec["slug"])
        lq = O("LSEG row")
        lsx = lambda c, lq=lq: f'=IF(ISNUMBER({lq}),IF(COUNT(INDEX({LS_(c)},{lq}))=1,INDEX({LS_(c)},{lq}),""),"")'
        f["IV before (LSEG)"] = lsx("H"); f["IV after (LSEG)"] = lsx("I"); f["Implied event move (LSEG)"] = lsx("J")
        iml = O("Implied event move (LSEG)")
        f["Implied event move (LSEG) - realised |share return|"] = f'=IF(AND(ISNUMBER({iml}),ISNUMBER({rr_})),{iml}-ABS({rr_}),"")'
        f["Analyst revision (LSEG)"] = lsx("N")
        rvc = O("Analyst revision (LSEG)")
        f["Revision up (1/0)"] = f'=IF(ISNUMBER({rvc}),IF({rvc}>0,1,0),"")'
        ivb, iva, kk = O("IV before (LSEG)"), O("IV after (LSEG)"), f'(1-{A["OTM"]})'
        t1 = f'({A["IVD"]}/365)'
        hold = f'(INDEX({LS_("E")},{lq})-INDEX({LS_("D")},{lq}))'
        t2 = f'(MAX(1,{A["IVD"]}-{hold})/365)'
        f["Put price / share price (LSEG IV)"] = f'=IF(AND(ISNUMBER({ivb}),N({ivb})>0),{bs_put("1", kk, ivb, t1)},"")'
        f["Put value after / share price (LSEG IV)"] = (f'=IF(AND(ISNUMBER({iva}),N({iva})>0,ISNUMBER({rr_})),'
                                                         f'{bs_put("(1+" + rr_ + ")", kk, iva, t2)},"")')
        pl1, pl2 = O("Put price / share price (LSEG IV)"), O("Put value after / share price (LSEG IV)")
        f["Hedge P&L with LSEG-priced put"] = f'=IF(AND(ISNUMBER({pl1}),ISNUMBER({pl2})),{pl2}-{pl1},"")'
        f["Implied move before the report (LSEG)"] = lsx("R")
        f["StarMine predicted surprise (LSEG)"] = lsx("S")
        aiq = f"Raw_Models!{L(2 + [x[0] for x in MODEL_SPECS].index('q10_crowd_mgr_gb'))}{r}"
        aif = f"Raw_Models!{L(2 + len(MODEL_SPECS))}{r}"
        f["AI 10% worst case (Raw_Models)"] = f'=IF(ISNUMBER({aiq}),{aiq},"")'
        hpl, arc, pc = O("Hedge P&L with LSEG-priced put"), O("Abnormal return"), O("p: P(beat) at cut-off")
        f["In backtest (1/0)"] = f'=IF(AND(ISNUMBER({arc}),ISNUMBER({pc}),ISNUMBER({aif}),ISNUMBER({hpl})),1,"")'
        inb = O("In backtest (1/0)")
        f["Rule trims (1/0)"] = f'=IF({inb}=1,IF({pc}>={A["HP"]},1,0),"")'
        f["AI trims (1/0)"] = f'=IF({inb}=1,{aif},"")'
        rt, at_ = O("Rule trims (1/0)"), O("AI trims (1/0)")
        f["Outcome: hold"] = f'=IF({inb}=1,{arc},"")'
        f["Outcome: rule trims"] = f'=IF({inb}=1,{arc}*(1-{A["TRIM"]}*{rt})-{A["TCOST"]}*{A["TRIM"]}*{rt},"")'
        f["Outcome: AI trims"] = f'=IF({inb}=1,{arc}*(1-{A["TRIM"]}*{at_})-{A["TCOST"]}*{A["TRIM"]}*{at_},"")'
        f["Outcome: rule buys the put"] = f'=IF({inb}=1,{arc}+{rt}*{hpl},"")'
        f["Outcome: AI buys the put"] = f'=IF({inb}=1,{arc}+{at_}*{hpl},"")'
        rk = f'MATCH({O("Ticker")}&"|"&TEXT({O("Report date")},"yyyy-mm-dd"),{RL_("A")},0)'
        f["Last release: guidance (generative AI)"] = f'=IFERROR(IF(INDEX({RL_("H")},{rk})=0,"",INDEX({RL_("H")},{rk})),"")'
        f["Last release: tone (generative AI)"] = f'=IFERROR(IF(ISNUMBER(INDEX({RL_("I")},{rk})),INDEX({RL_("I")},{rk}),""),"")'
        f["Last release: keyword balance"] = f'=IFERROR(IF(ISNUMBER(INDEX({RL_("L")},{rk})),INDEX({RL_("L")},{rk}),""),"")'
        f["Skew (LSEG put - call IV)"] = lsx("V")
        for j, (lab_f, lab_o) in enumerate((("Past-move rule trims (1/0)", "Outcome: past-move rule"),
                                            ("Implied-move rule trims (1/0)", "Outcome: implied-move rule"),
                                            ("Expected-loss rule trims (1/0)", "Outcome: expected-loss rule"))):
            sf = f"Raw_Models!{L(3 + len(MODEL_SPECS) + j)}{r}"
            f[lab_f] = f'=IF({inb}=1,IF(ISNUMBER({sf}),{sf},0),"")'
            sfl = O(lab_f)
            f[lab_o] = f'=IF({inb}=1,{arc}*(1-{A["TRIM"]}*{sfl})-{A["TCOST"]}*{A["TRIM"]}*{sfl},"")'
        for j, (_, lab_f, lab_o, sized) in enumerate(NEW_POLICIES):
            sf = f"Raw_Models!{L(6 + len(MODEL_SPECS) + j)}{r}"
            f[lab_f] = f'=IF({inb}=1,IF(ISNUMBER({sf}),{sf},0),"")'
            sh = O(lab_f) if sized else f'{A["TRIM"]}*{O(lab_f)}'
            f[lab_o] = f'=IF({inb}=1,{arc}*(1-{sh})-{A["TCOST"]}*{sh},"")'
        chg, smc, rvc = O("Change in p over the week"), O("StarMine predicted surprise (LSEG)"), O("Analyst revision (LSEG)")
        f["Crowd lost confidence (1/0)"] = f'=IF(ISNUMBER({chg}),IF({chg}<=-{A["DRIFT"]},1,0),"")'
        f["StarMine expects a miss, confident crowd (1/0)"] = f'=IF(AND({cf}=1,ISNUMBER({smc})),IF({smc}<0,1,0),"")'
        f["Analysts cut estimates, confident crowd (1/0)"] = f'=IF(AND({cf}=1,ISNUMBER({rvc})),IF({rvc}<0,1,0),"")'
        f["Crowd warns (1/0)"] = (f'=IF(OR({O("Crowd lost confidence (1/0)")}=1,{O("StarMine expects a miss, confident crowd (1/0)")}=1),1,'
                                  f'IF(ISNUMBER({chg}),0,""))')
        lc, smd, rvd = O("Crowd lost confidence (1/0)"), O("StarMine expects a miss, confident crowd (1/0)"), O("Analysts cut estimates, confident crowd (1/0)")
        f["AR | crowd lost confidence"] = f'=IF(AND({lc}=1,ISNUMBER({arc})),{arc},"")'
        f["AR | crowd steady or more sure"] = f'=IF(AND({lc}=0,ISNUMBER({arc})),{arc},"")'
        f["AR | confident, StarMine expects a miss"] = f'=IF(AND({smd}=1,ISNUMBER({arc})),{arc},"")'
        f["AR | confident, StarMine agrees"] = f'=IF(AND({smd}=0,ISNUMBER({arc})),{arc},"")'
        f["AR | confident, analysts cut"] = f'=IF(AND({rvd}=1,ISNUMBER({arc})),{arc},"")'
        f["AR | confident, analysts up or flat"] = f'=IF(AND({rvd}=0,ISNUMBER({arc})),{arc},"")'
        sf = f"Raw_Models!{L(6 + len(MODEL_SPECS) + len(NEW_POLICIES))}{r}"
        f["Crowd-warning rule trims (1/0)"] = f'=IF({inb}=1,IF(ISNUMBER({sf}),{sf},0),"")'
        cwf = O("Crowd-warning rule trims (1/0)")
        f["Outcome: crowd-warning rule"] = f'=IF({inb}=1,{arc}*(1-{A["TRIM"]}*{cwf})-{A["TCOST"]}*{A["TRIM"]}*{cwf},"")'
        f["Has 3 past reports (1/0)"] = f'=IF(ISNUMBER({O("Company" + chr(39) + "s average past |move|")}),1,0)'
        for oc in BT_SUB:
            f[f"{oc} | no 3 past reports"] = f'=IF(AND(ISNUMBER({O(oc)}),{O("Has 3 past reports (1/0)")}=0),{O(oc)},"")'
            f[f"{oc} | new reports (holdout)"] = f'=IF(AND(ISNUMBER({O(oc)}),{O("Sample")}="New"),{O(oc)},"")'
        for h, v in f.items():
            EVT[f"{C[h]}{r}"] = v
        vc = f"{C['Volatility before (SD of daily abnormal returns)']}{r}"
        EVT[vc] = ArrayFormula(vc, vol_txt)
    fmt = {"Report date": "yyyy-mm-dd", "Volume (USD)": "#,##0", "Abnormal return": PCT, "Share return (simple)": PCT,
           "Stock log return": PCT, "Market log return": PCT, "Put payoff (share of S0)": PC2, "Run-up before (abnormal)": PCT,
           "Volatility before (SD of daily abnormal returns)": PC2, "Company past beat rate": "0%", "Company's average past |move|": PC1,
           "Running base rate": "0%", "Option-implied move (snapshot)": PC1, "Put price / share price (snapshot)": PC2,
           "EPS estimate change, 30 days (snapshot)": PCT, "Week (Monday)": "yyyy-mm-dd", "Release (New York time)": "yyyy-mm-dd hh:mm",
           "Hedge P&L: payoff - snapshot put price": "+0.00%;-0.00%", "Option-implied move - realised |share return|": "+0.0%;-0.0%",
           "Put price / share price (history)": "0.00%", "Put value after / share price (history)": "0.00%",
           "Hedge P&L with actual option prices (history)": "+0.00%;-0.00%", "Implied move (history)": "0.0%",
           "Implied move (history) - realised |share return|": "+0.0%;-0.0%", "IV before (LSEG)": "0.0%", "IV after (LSEG)": "0.0%",
           "Implied event move (LSEG)": "0.0%", "Implied event move (LSEG) - realised |share return|": "+0.0%;-0.0%",
           "Analyst revision (LSEG)": "+0.0%;-0.0%", "Put price / share price (LSEG IV)": "0.00%",
           "Put value after / share price (LSEG IV)": "0.00%", "Hedge P&L with LSEG-priced put": "+0.00%;-0.00%",
           "Implied move before the report (LSEG)": "0.0%", "StarMine predicted surprise (LSEG)": "+0.0%;-0.0%",
           "AI 10% worst case (Raw_Models)": "+0.0%;-0.0%", "Outcome: hold": "+0.0%;-0.0%", "Outcome: rule trims": "+0.0%;-0.0%",
           "Outcome: AI trims": "+0.0%;-0.0%", "Outcome: rule buys the put": "+0.0%;-0.0%", "Outcome: AI buys the put": "+0.0%;-0.0%",
           "Outcome: past-move rule": "+0.0%;-0.0%", "Outcome: implied-move rule": "+0.0%;-0.0%", "Outcome: expected-loss rule": "+0.0%;-0.0%"}
    fmt.update({h: "+0.0%;-0.0%" for h in ec if h.startswith("Outcome")})
    fmt["Sized trim (share of the position)"] = "0%"
    for h in ec:
        nf = fmt.get(h, PCT if h.startswith("AR") or h.startswith("Share return if") else
                     ("0.000" if h.startswith("p") or h.startswith("Brier") or h.startswith("Change in p") else None))
        if nf:
            for r in range(5, 5 + n):
                EVT[f"{C[h]}{r}"].number_format = nf
    EVT.freeze_panes = "D5"
    R = lambda h: f"Events!${C[h]}$5:${C[h]}${E1}"
    Rp, Ry, Rar, Rs, Rc = R("p: P(beat) at cut-off"), R("Outcome y (1 = beat)"), R("Abnormal return"), R("Sample"), R("Confident (1/0)")

    # ---------- Raw_Models and Model_Calc ----------
    RMOD["A1"] = ("Walk-forward model forecasts written by the script (scikit-learn). Each month is forecast by models trained only on "
                  "earlier months. Same row order as Events."); RMOD["A1"].font = BOLD
    head(RMOD, 4, ["Contract slug"] + [lab for _, lab, *_ in MODEL_SPECS] +
         ["AI policy trims (1/0): predicted 10% worst case at or below the cut-off from earlier months, same turnover as the rule"] +
         [f"{lab} (1/0), cut-off from earlier months" for _, _, lab in SHEET_RULES] +
         ["Pre-registered 1: fall classifier trims (1/0)", "Pre-registered 2: residual AI trims (1/0), main AI flag where no past moves",
          "Pre-registered 3: sized trim, share of the position sold",
          "Part 3: crowd-warning screen trims (1/0), past |move| raised when the crowd warns, cut-off from earlier months"])
    for i, rec in enumerate(models.to_dict("records")):
        put(RMOD, 5 + i, 1, rec["slug"])
        for j, (col, *_ ) in enumerate(MODEL_SPECS):
            put(RMOD, 5 + i, 2 + j, clean(rec[col]), "0.000")
        put(RMOD, 5 + i, 2 + len(MODEL_SPECS), clean(rec["ai_trim"]), "0")
        for j, (k, *_ ) in enumerate(SHEET_RULES):
            put(RMOD, 5 + i, 3 + len(MODEL_SPECS) + j, clean(rec[k]), "0")
        for j, (k, *_ ) in enumerate(NEW_POLICIES):
            put(RMOD, 5 + i, 6 + len(MODEL_SPECS) + j, clean(rec[k]), "0.00")
        put(RMOD, 5 + i, 6 + len(MODEL_SPECS) + len(NEW_POLICIES), clean(rec["cw_trim"]), "0")
    title(MC, "Model_Calc: errors of each model and its benchmark on the same reports (formulas)",
          "Squared errors for probabilities (Brier), absolute errors for the size of the move. A row counts only when the model has "
          "a forecast and the benchmark exists.", "A2:R2")
    mc_h = ["Ticker", "Report date", "y", "p", "Fall (1/0)", "|Abnormal return|", "Company's average past |move|", "Option-implied move",
            "|Share return|"]
    for _, lab, *_ in MODEL_SPECS:
        mc_h += [f"{lab}: error", f"{lab}: benchmark error"]
    mc_h += ["Option-implied move: error", "Option-implied move: benchmark error (company's past move)",
             "Implied move (history): error", "Implied move (history): benchmark error (company's past move)", "Implied move (history)",
             "Implied event move (LSEG)", "LSEG implied move: error", "LSEG implied move: benchmark error (company's past move)",
             "Abnormal return", "Options 10% worst case (LSEG move before the report, normal)", "Options 10% worst case: error",
             "Options 10% worst case: benchmark error (history)", "AI 10% worst case vs options: error",
             "AI 10% worst case vs options: benchmark error (options)", "AI 10% worst case: move below the forecast (1/0)",
             "AI 10% worst case vs spreadsheet: error", "AI 10% worst case vs spreadsheet: benchmark error (k x past move)",
             "AI with the release reading: error", "AI with the release reading: benchmark error (same AI without it)"]
    head(MC, 4, mc_h, h=75)
    MCc = {h: L(1 + j) for j, h in enumerate(mc_h)}
    for i in range(n):
        r, e = 5 + i, 5 + i
        MC[f"A{r}"] = f"=Events!{C['Ticker']}{e}"; MC[f"B{r}"] = f"=Events!{C['Report date']}{e}"; MC[f"B{r}"].number_format = "yyyy-mm-dd"
        MC[f"C{r}"] = f"=Events!{C['Outcome y (1 = beat)']}{e}"; MC[f"D{r}"] = f"=Events!{C['p: P(beat) at cut-off']}{e}"
        MC[f"E{r}"] = f"=Events!{C['Fall (1/0)']}{e}"
        MC[f"F{r}"] = f'=IF(ISNUMBER(Events!{C["Abnormal return"]}{e}),ABS(Events!{C["Abnormal return"]}{e}),"")'
        MC[f"G{r}"] = f"=Events!{C['Company' + chr(39) + 's average past |move|']}{e}"
        MC[f"H{r}"] = f"=Events!{C['Option-implied move (snapshot)']}{e}"
        MC[f"I{r}"] = f'=IF(ISNUMBER(Events!{C["Share return (simple)"]}{e}),ABS(Events!{C["Share return (simple)"]}{e}),"")'
        fall_bench = f"Raw_Models!{L(2 + [s[0] for s in MODEL_SPECS].index('fall_crowd_lin'))}{r}"
        arm = MCc["Abnormal return"]
        MC[f"{arm}{r}"] = f"=Events!{C['Abnormal return']}{e}"
        pin = lambda f_, y_: f'MAX({A["QLOW"]}*({y_}-{f_}),({A["QLOW"]}-1)*({y_}-{f_}))'
        spec_col = {x[0]: j for j, x in enumerate(MODEL_SPECS)}
        for j, (col, lab, tgt, kind, *_ ) in enumerate(MODEL_SPECS):
            pr = f"Raw_Models!{L(2 + j)}{r}"
            ce, cb = MCc[f"{lab}: error"], MCc[f"{lab}: benchmark error"]
            bpr = f"Raw_Models!{L(2 + spec_col[BENCH[col]])}{r}" if col in BENCH else None
            if tgt == "y" and bpr:
                MC[f"{ce}{r}"] = f'=IF(AND(ISNUMBER({pr}),ISNUMBER({bpr}),ISNUMBER(D{r})),({pr}-C{r})^2,"")'
                MC[f"{cb}{r}"] = f'=IF(ISNUMBER({ce}{r}),({bpr}-C{r})^2,"")'
            elif tgt == "big10":
                continue
            elif tgt == "ar":
                bq = bpr or pr
                MC[f"{ce}{r}"] = f'=IF(AND(ISNUMBER({pr}),ISNUMBER({bq}),ISNUMBER({arm}{r})),{pin(pr, arm + str(r))},"")'
                MC[f"{cb}{r}"] = f'=IF(ISNUMBER({ce}{r}),{pin(bq, arm + str(r))},"")'
            elif tgt == "y":
                MC[f"{ce}{r}"] = f'=IF(AND(ISNUMBER({pr}),ISNUMBER(D{r})),({pr}-C{r})^2,"")'
                MC[f"{cb}{r}"] = f'=IF(ISNUMBER({ce}{r}),(D{r}-C{r})^2,"")'
            elif tgt == "fall5":
                MC[f"{ce}{r}"] = f'=IF(AND(ISNUMBER({pr}),ISNUMBER(E{r}),ISNUMBER({fall_bench})),({pr}-E{r})^2,"")'
                MC[f"{cb}{r}"] = f'=IF(ISNUMBER({ce}{r}),({fall_bench}-E{r})^2,"")'
            else:
                MC[f"{ce}{r}"] = f'=IF(AND(ISNUMBER({pr}),ISNUMBER(F{r}),ISNUMBER(G{r})),ABS({pr}-F{r}),"")'
                MC[f"{cb}{r}"] = f'=IF(ISNUMBER({ce}{r}),ABS(G{r}-F{r}),"")'
        lm = MCc["Implied event move (LSEG)"]
        MC[f"{lm}{r}"] = f"=Events!{C['Implied event move (LSEG)']}{e}"
        ce3, cb3 = MCc["LSEG implied move: error"], MCc["LSEG implied move: benchmark error (company's past move)"]
        MC[f"{ce3}{r}"] = f'=IF(AND(ISNUMBER({lm}{r}),ISNUMBER(I{r}),ISNUMBER(G{r})),ABS({lm}{r}-I{r}),"")'
        MC[f"{cb3}{r}"] = f'=IF(ISNUMBER({ce3}{r}),ABS(G{r}-F{r}),"")'
        hm = MCc["Implied move (history)"]
        MC[f"{hm}{r}"] = f"=Events!{C['Implied move (history)']}{e}"
        ce2, cb2 = MCc["Implied move (history): error"], MCc["Implied move (history): benchmark error (company's past move)"]
        MC[f"{ce2}{r}"] = f'=IF(AND(ISNUMBER({hm}{r}),ISNUMBER(I{r}),ISNUMBER(G{r})),ABS({hm}{r}-I{r}),"")'
        MC[f"{cb2}{r}"] = f'=IF(ISNUMBER({ce2}{r}),ABS(G{r}-F{r}),"")'
        ce, cb = MCc["Option-implied move: error"], MCc["Option-implied move: benchmark error (company's past move)"]
        MC[f"{ce}{r}"] = f'=IF(AND(ISNUMBER(H{r}),ISNUMBER(I{r}),ISNUMBER(G{r})),ABS(H{r}-I{r}),"")'
        MC[f"{cb}{r}"] = f'=IF(ISNUMBER({ce}{r}),ABS(G{r}-F{r}),"")'
        oq, imp = MCc["Options 10% worst case (LSEG move before the report, normal)"], f"Events!{C['Implied move before the report (LSEG)']}{e}"
        MC[f"{oq}{r}"] = f'=IF(ISNUMBER({imp}),NORMSINV({A["QLOW"]})*{imp},"")'
        hq = f"Raw_Models!{L(2 + spec_col['q10_hist'])}{r}"
        aq = f"Raw_Models!{L(2 + spec_col['q10_crowd_mgr_gb'])}{r}"
        y_ = f"{arm}{r}"
        e1, b1 = MCc["Options 10% worst case: error"], MCc["Options 10% worst case: benchmark error (history)"]
        MC[f"{e1}{r}"] = f'=IF(AND(ISNUMBER({oq}{r}),ISNUMBER({hq}),ISNUMBER({y_})),{pin(oq + str(r), y_)},"")'
        MC[f"{b1}{r}"] = f'=IF(ISNUMBER({e1}{r}),{pin(hq, y_)},"")'
        e2, b2 = MCc["AI 10% worst case vs options: error"], MCc["AI 10% worst case vs options: benchmark error (options)"]
        MC[f"{e2}{r}"] = f'=IF(AND(ISNUMBER({oq}{r}),ISNUMBER({aq}),ISNUMBER({y_})),{pin(aq, y_)},"")'
        MC[f"{b2}{r}"] = f'=IF(ISNUMBER({e2}{r}),{pin(oq + str(r), y_)},"")'
        e3, b3 = MCc["AI 10% worst case vs spreadsheet: error"], MCc["AI 10% worst case vs spreadsheet: benchmark error (k x past move)"]
        sq = f"Raw_Models!{L(2 + spec_col['q10_past_scaled'])}{r}"
        MC[f"{e3}{r}"] = f'=IF(AND(ISNUMBER({sq}),ISNUMBER({aq}),ISNUMBER({y_})),{pin(aq, y_)},"")'
        MC[f"{b3}{r}"] = f'=IF(ISNUMBER({e3}{r}),{pin(sq, y_)},"")'
        e4, b4 = MCc["AI with the release reading: error"], MCc["AI with the release reading: benchmark error (same AI without it)"]
        nq = f"Raw_Models!{L(2 + spec_col['q10_notext_gb'])}{r}"
        MC[f"{e4}{r}"] = f'=IF(AND(ISNUMBER({nq}),ISNUMBER({aq}),ISNUMBER({y_})),{pin(aq, y_)},"")'
        MC[f"{b4}{r}"] = f'=IF(ISNUMBER({e4}{r}),{pin(nq, y_)},"")'
        cv = MCc["AI 10% worst case: move below the forecast (1/0)"]
        MC[f"{cv}{r}"] = f'=IF(AND(ISNUMBER({aq}),ISNUMBER({y_})),IF({y_}<{aq},1,0),"")'
    MC.freeze_panes = "C5"
    MR = lambda h: f"Model_Calc!${MCc[h]}$5:${MCc[h]}${E1}"

    # ---------- Results ----------
    title(RES, "Results: does the Part 1 finding hold, and on new reports?",
          "Every cell is a formula on Events. 'New' = reports after the Part 1 data run (Assumptions), the out-of-sample test. "
          "The confident threshold and the fall threshold were fixed before the new reports were seen.", "A2:H2")
    r = 4
    put(RES, r, 1, "A. Sample", font=SEC); r += 1
    head(RES, r, ["", "All", "Part 1", "New"]); r += 1
    SMP = [None, "Part 1", "New"]
    cond = lambda s: f',{Rs},"{s}"' if s else ""
    for lab, fn, nf in [("Settled reports", lambda s: f'=COUNTIFS({Ry},">=0"{cond(s)})', "0"),
                        ("Reports with crowd odds", lambda s: f'=COUNTIFS({Rp},">=0"{cond(s)})', "0"),
                        ("Share that beat", lambda s: f'=IFERROR(AVERAGEIFS({Ry},{Ry},">=0"{cond(s)}),"n/a")', PC1),
                        ("Average crowd p", lambda s: f'=IFERROR(AVERAGEIFS({Rp},{Rp},">=0"{cond(s)}),"n/a")', PC1)]:
        put(RES, r, 1, lab)
        for j, s in enumerate(SMP):
            put(RES, r, 2 + j, fn(s), nf)
        r += 1
    r += 1
    put(RES, r, 1, "B. Accuracy: Brier score on reports with all three forecasts (lower is better)", font=SEC); r += 1
    head(RES, r, ["", "All", "Part 1", "New"]); r += 1
    Rall = R("All three forecasts (1/0)")
    b0 = r
    for lab, col in [("Brier, crowd", "Brier: crowd"), ("Brier, running base rate", "Brier: base rate"),
                     ("Brier, company record", "Brier: company record")]:
        put(RES, r, 1, lab)
        for j, s in enumerate(SMP):
            put(RES, r, 2 + j, f'=IFERROR(AVERAGEIFS({R(col)},{Rall},1{cond(s)}),"n/a")', "0.000")
        r += 1
    for lab, k in [("Skill vs base rate (1 - crowd / base rate)", 1), ("Skill vs company record", 2)]:
        put(RES, r, 1, lab)
        for j in range(3):
            c_ = L(2 + j)
            put(RES, r, 2 + j, f'=IFERROR(1-{c_}{b0}/{c_}{b0 + k},"n/a")', PCT)
        r += 1
    put(RES, r, 1, "Reports with all three forecasts")
    for j, s in enumerate(SMP):
        put(RES, r, 2 + j, f'=COUNTIFS({Rall},1{cond(s)})', "0")
    r += 2
    put(RES, r, 1, "C. Calibration: when the crowd says p, how often do companies beat? (all reports)", font=SEC); r += 1
    head(RES, r, ["p from", "p to", "Reports", "Average p", "Share that beat"]); r += 1
    cal0 = r
    for lo, hi in [(0, .2), (.2, .4), (.4, .6), (.6, .8), (.8, .9), (.9, 1.0001)]:
        put(RES, r, 1, lo, font=BLUE); put(RES, r, 2, hi, font=BLUE)
        put(RES, r, 3, f'=COUNTIFS({Rp},">="&A{r},{Rp},"<"&B{r})')
        put(RES, r, 4, f'=IFERROR(AVERAGEIFS({Rp},{Rp},">="&A{r},{Rp},"<"&B{r}),"")', PC1)
        put(RES, r, 5, f'=IFERROR(AVERAGEIFS({Ry},{Rp},">="&A{r},{Rp},"<"&B{r}),"")', PC1)
        r += 1
    ch = ScatterChart(); ch.title = "Calibration"; ch.style = 13; ch.height = 7; ch.width = 11; ch.legend = None
    ch.x_axis.title = "Average crowd p"; ch.y_axis.title = "Share that beat"
    for ax in (ch.x_axis, ch.y_axis):
        ax.scaling.min, ax.scaling.max, ax.delete = 0, 1, False
    s_ = Series(Reference(RES, min_col=5, min_row=cal0, max_row=r - 1), Reference(RES, min_col=4, min_row=cal0, max_row=r - 1))
    s_.marker.symbol = "circle"; s_.graphicalProperties.line.noFill = True; ch.series.append(s_)
    RES.add_chart(ch, "G4")
    r += 1
    put(RES, r, 1, "D. Table 1: average abnormal return by crowd confidence and outcome", font=SEC); r += 1
    for s in SMP:
        head(RES, r, [f"Sample: {s or 'All'}", "Crowd confident", "Crowd less sure", "Difference", "Reports: confident",
                      "Reports: less sure"]); r += 1
        for lab, yv in (("Company beat", 1), ("Company missed", 0)):
            put(RES, r, 1, lab)
            for j, g in enumerate((1, 0)):
                put(RES, r, 2 + j, f'=IFERROR(AVERAGEIFS({Rar},{Ry},{yv},{Rc},{g}{cond(s)}),"n/a")', PCT)
                put(RES, r, 5 + j, f'=COUNTIFS({Ry},{yv},{Rc},{g},{Rar},">-100"{cond(s)})', "0")
            put(RES, r, 4, f'=IFERROR(B{r}-C{r},"n/a")', PCT)
            r += 1
        r += 1
    put(RES, r, 1, "E. Is the gap on a miss real? Welch test, confident misses against other misses", font=SEC); r += 1
    head(RES, r, ["", "All", "Part 1", "New"]); r += 1
    w0 = r
    wl = ["Confident misses", "Other misses", "Average move: confident misses", "Average move: other misses", "Gap",
          "Variance: confident misses", "Variance: other misses", "Standard error of the gap", "Welch t", "Degrees of freedom",
          "95% interval: low", "95% interval: high", "p-value (two-sided)"]
    cols_w = {None: ("AR: confident miss", "AR: other miss"), "Part 1": ("AR: confident miss, Part 1", "AR: other miss, Part 1"),
              "New": ("AR: confident miss, new", "AR: other miss, new")}
    for k, lab in enumerate(wl):
        put(RES, w0 + k, 1, lab)
    for j, s in enumerate(SMP):
        c_ = L(2 + j)
        x1, x2 = R(cols_w[s][0]), R(cols_w[s][1])
        rr0 = w0
        cells = [f"=COUNT({x1})", f"=COUNT({x2})", f'=IFERROR(AVERAGE({x1}),"n/a")', f'=IFERROR(AVERAGE({x2}),"n/a")',
                 f'=IFERROR({c_}{rr0+2}-{c_}{rr0+3},"n/a")', f'=IFERROR(VAR({x1}),"n/a")', f'=IFERROR(VAR({x2}),"n/a")',
                 f'=IFERROR(SQRT({c_}{rr0+5}/{c_}{rr0}+{c_}{rr0+6}/{c_}{rr0+1}),"n/a")', f'=IFERROR({c_}{rr0+4}/{c_}{rr0+7},"n/a")',
                 f'=IFERROR(({c_}{rr0+5}/{c_}{rr0}+{c_}{rr0+6}/{c_}{rr0+1})^2/(({c_}{rr0+5}/{c_}{rr0})^2/({c_}{rr0}-1)'
                 f'+({c_}{rr0+6}/{c_}{rr0+1})^2/({c_}{rr0+1}-1)),"n/a")',
                 f'=IFERROR({c_}{rr0+4}-TINV(0.05,{c_}{rr0+9})*{c_}{rr0+7},"n/a")',
                 f'=IFERROR({c_}{rr0+4}+TINV(0.05,{c_}{rr0+9})*{c_}{rr0+7},"n/a")',
                 f'=IFERROR(TDIST(ABS({c_}{rr0+8}),{c_}{rr0+9},2),"n/a")']
        nfs = ["0", "0", PCT, PCT, PCT, "0.0000", "0.0000", PC2, "0.00", "0.0", PCT, PCT, "0.000"]
        for k, (v, nf) in enumerate(zip(cells, nfs)):
            put(RES, w0 + k, 2 + j, v, nf)
    r = w0 + len(wl) + 1
    for t in ["Reading: the New column is the out-of-sample test. It needs at least a few dozen misses before it can confirm or reject the Part 1 result.",
              "Week-clustered tests, the continuous version and the threshold sweep are on the Tests sheet."]:
        put(RES, r, 1, t, font=NOTE); r += 1
    r += 1
    put(RES, r, 1, "F. Brier decomposition (Murphy): reliability, resolution, uncertainty. Reports with all three forecasts", font=SEC); r += 1
    put(RES, r, 1, "Brier = reliability - resolution + uncertainty (exact up to the binning). Reliability: lower is better (probabilities "
                   "mean what they say). Resolution: higher is better (the forecast separates beats from misses).", font=NOTE); r += 1
    N_all = f"COUNTIFS({Rall},1)"
    ob = f"AVERAGEIFS({Ry},{Rall},1)"
    dec = {}
    for lab, colh in (("Crowd p", "p: P(beat) at cut-off"), ("Company record", "Company past beat rate"), ("Running base rate", "Running base rate")):
        head(RES, r, [f"Forecast: {lab}", "to", "Reports", "Average forecast", "Share that beat", "Reliability term", "Resolution term"]); r += 1
        Rf_ = R(colh); d0 = r
        for k in range(10):
            lo, hi = k / 10, (k + 1) / 10 + (0.0001 if k == 9 else 0)
            put(RES, r, 1, lo, font=BLUE); put(RES, r, 2, hi, font=BLUE)
            put(RES, r, 3, f'=COUNTIFS({Rall},1,{Rf_},">="&A{r},{Rf_},"<"&B{r})', "0")
            put(RES, r, 4, f'=IFERROR(AVERAGEIFS({Rf_},{Rall},1,{Rf_},">="&A{r},{Rf_},"<"&B{r}),"")', PC1)
            put(RES, r, 5, f'=IFERROR(AVERAGEIFS({Ry},{Rall},1,{Rf_},">="&A{r},{Rf_},"<"&B{r}),"")', PC1)
            put(RES, r, 6, f"=IF(C{r}>0,C{r}*(D{r}-E{r})^2,0)", "0.0000")
            put(RES, r, 7, f"=IF(C{r}>0,C{r}*(E{r}-{ob})^2,0)", "0.0000")
            r += 1
        dec[lab] = (d0, r - 1)
        r += 1
    head(RES, r, ["Decomposition", "Crowd p", "Company record", "Running base rate"]); r += 1
    q0 = r
    rows_ = [("Brier score (direct)", lambda c: f'=AVERAGEIFS({R({"Crowd p": "Brier: crowd", "Company record": "Brier: company record", "Running base rate": "Brier: base rate"}[c])},{Rall},1)'),
             ("Reliability (lower is better)", lambda c: f"=SUM(F{dec[c][0]}:F{dec[c][1]})/{N_all}"),
             ("Resolution (higher is better)", lambda c: f"=SUM(G{dec[c][0]}:G{dec[c][1]})/{N_all}"),
             ("Uncertainty (same for all)", lambda c: f"={ob}*(1-{ob})"),
             ("Reliability - resolution + uncertainty (check)", None)]
    for k, (lab, fn) in enumerate(rows_):
        put(RES, r, 1, lab)
        for j, c in enumerate(("Crowd p", "Company record", "Running base rate")):
            cc = L(2 + j)
            put(RES, r, 2 + j, fn(c) if fn else f"={cc}{q0+1}-{cc}{q0+2}+{cc}{q0+3}", "0.0000")
        r += 1
    RES_DEC = q0
    r += 1
    for t in ["Reading: the New column is the out-of-sample test. It needs at least a few dozen misses before it can confirm or reject the Part 1 result.",
              "Bootstrap intervals that keep each earnings week together are on the Bootstrap sheet."][2:]:
        put(RES, r, 1, t, font=NOTE); r += 1
    RES.column_dimensions["A"].width = 58
    for c_ in "BCDEF":
        RES.column_dimensions[c_].width = 15

    # ---------- Tests: week-clustered tests as formulas (Test_Calc holds one row per report, Weeks sums by week) ----------
    EV = lambda h, rr: f"Events!{C[h]}{rr}"
    tests = []
    for smp in ("All", "New"):
        sc = (lambda rr: "TRUE") if smp == "All" else (lambda rr, smp=smp: f'{EV("Sample", rr)}="New"')
        okr = lambda rr, sc=sc: f'ISNUMBER({EV("Abnormal return", rr)}),ISNUMBER({EV("p: P(beat) at cut-off", rr)}),{sc(rr)}'
        tests += [
            (f"{smp}_skill_base", smp, "Crowd beats the running base rate (Brier base rate - Brier crowd)", "mean",
             lambda rr, sc=sc: f'{EV("All three forecasts (1/0)", rr)}=1,{sc(rr)}',
             lambda rr: f'{EV("Brier: base rate", rr)}-{EV("Brier: crowd", rr)}', None, "0.000"),
            (f"{smp}_skill_rec", smp, "Crowd beats the company record (Brier record - Brier crowd)", "mean",
             lambda rr, sc=sc: f'{EV("All three forecasts (1/0)", rr)}=1,{sc(rr)}',
             lambda rr: f'{EV("Brier: company record", rr)}-{EV("Brier: crowd", rr)}', None, "0.000"),
            (f"{smp}_surprise", smp, "Reaction follows the surprise against the crowd: slope of AR on y - p", "slope",
             okr, lambda rr: f'{EV("Outcome y (1 = beat)", rr)}-{EV("p: P(beat) at cut-off", rr)}',
             lambda rr: EV("Abnormal return", rr), "0.0000"),
            (f"{smp}_beats", smp, "Beats are priced in: slope of AR on p, beats only", "slope",
             lambda rr, okr=okr: f'{okr(rr)},{EV("Outcome y (1 = beat)", rr)}=1', lambda rr: EV("p: P(beat) at cut-off", rr),
             lambda rr: EV("Abnormal return", rr), "0.0000"),
            (f"{smp}_misses", smp, "Continuous test: slope of AR on p, misses only", "slope",
             lambda rr, okr=okr: f'{okr(rr)},{EV("Outcome y (1 = beat)", rr)}=0', lambda rr: EV("p: P(beat) at cut-off", rr),
             lambda rr: EV("Abnormal return", rr), "0.0000")]
        ths = [(f"TH{i + 1}", A[f"TH{i + 1}"]) for i in range(len(GAP_THRESHOLDS))] if smp == "All" else [("HP", A["HP"])]
        for key, cell in ths:
            tests.append((f"{smp}_gap_{key}", smp, f"Miss gap: confident minus other misses, crowd p at least the threshold", "slope",
                          lambda rr, okr=okr: f'{okr(rr)},{EV("Outcome y (1 = beat)", rr)}=0',
                          lambda rr, cell=cell: f'IF({EV("p: P(beat) at cut-off", rr)}>={cell},1,0)',
                          lambda rr: EV("Abnormal return", rr), "+0.0%;-0.0%", cell))
        tests.append((f"{smp}_fall", smp, "Falls of 5% or more: confident minus less sure (share of reports)", "slope", okr,
                      lambda rr: f'IF({EV("p: P(beat) at cut-off", rr)}>={A["HP"]},1,0)', lambda rr: EV("Fall (1/0)", rr), "+0.0%;-0.0%"))
    okh = lambda rr: (f'ISNUMBER({EV("Abnormal return", rr)}),ISNUMBER({EV("p: P(beat) at cut-off", rr)}),'
                      f'ISNUMBER({EV("Hedge P&L with actual option prices (history)", rr)})')
    tests += [("All_hedge_conf", "All", "Hedge P&L with actual option prices, confident reports (below 0 = puts cost more than they paid)", "mean",
               lambda rr: f'{okh(rr)},{EV("p: P(beat) at cut-off", rr)}>={A["HP"]}',
               lambda rr: EV("Hedge P&L with actual option prices (history)", rr), None, "+0.00%;-0.00%"),
              ("All_hedge_gap", "All", "Hedge P&L with actual option prices: confident minus less sure", "slope", okh,
               lambda rr: f'IF({EV("p: P(beat) at cut-off", rr)}>={A["HP"]},1,0)',
               lambda rr: EV("Hedge P&L with actual option prices (history)", rr), "+0.00%;-0.00%"),
              ("All_impl_gap", "All", "Implied move minus realised |share return|: confident minus less sure", "slope",
               lambda rr: (f'ISNUMBER({EV("p: P(beat) at cut-off", rr)}),'
                           f'ISNUMBER({EV("Implied move (history) - realised |share return|", rr)})'),
               lambda rr: f'IF({EV("p: P(beat) at cut-off", rr)}>={A["HP"]},1,0)',
               lambda rr: EV("Implied move (history) - realised |share return|", rr), "+0.0%;-0.0%")]
    okl = lambda rr: f'ISNUMBER({EV("p: P(beat) at cut-off", rr)}),ISNUMBER({EV("Implied event move (LSEG)", rr)})'
    okl2 = lambda rr: (f'ISNUMBER({EV("p: P(beat) at cut-off", rr)}),'
                       f'ISNUMBER({EV("Implied event move (LSEG) - realised |share return|", rr)})')
    tests += [("All_lseg_move_gap", "All", "Options expect a smaller move when the crowd is confident? LSEG implied event move: confident minus less sure",
               "slope", okl, lambda rr: f'IF({EV("p: P(beat) at cut-off", rr)}>={A["HP"]},1,0)',
               lambda rr: EV("Implied event move (LSEG)", rr), "+0.0%;-0.0%"),
              ("All_lseg_over", "All", "Options overprice the earnings move: LSEG implied event move - realised |share return|", "mean",
               okl2, lambda rr: EV("Implied event move (LSEG) - realised |share return|", rr), None, "+0.0%;-0.0%"),
              ("All_lseg_over_gap", "All", "Overpricing when the crowd is confident minus less sure (LSEG)", "slope", okl2,
               lambda rr: f'IF({EV("p: P(beat) at cut-off", rr)}>={A["HP"]},1,0)',
               lambda rr: EV("Implied event move (LSEG) - realised |share return|", rr), "+0.0%;-0.0%")]
    okp = lambda rr: f'ISNUMBER({EV("p: P(beat) at cut-off", rr)}),ISNUMBER({EV("Hedge P&L with LSEG-priced put", rr)})'
    okv = lambda rr: (f'ISNUMBER({EV("p: P(beat) at cut-off", rr)}),ISNUMBER({EV("Analyst revision (LSEG)", rr)}),'
                      f'ISNUMBER({EV("Abnormal return", rr)})')
    tests += [("All_lput_conf", "All", "Hedge P&L with the LSEG-priced put, confident reports (below 0 = the put cost more than it paid)",
               "mean", lambda rr: f'{okp(rr)},{EV("p: P(beat) at cut-off", rr)}>={A["HP"]}',
               lambda rr: EV("Hedge P&L with LSEG-priced put", rr), None, "+0.00%;-0.00%"),
              ("All_lput_gap", "All", "Hedge P&L with the LSEG-priced put: confident minus less sure", "slope", okp,
               lambda rr: f'IF({EV("p: P(beat) at cut-off", rr)}>={A["HP"]},1,0)',
               lambda rr: EV("Hedge P&L with LSEG-priced put", rr), "+0.00%;-0.00%"),
              ("All_rev_beat", "All", "Beat rate: analysts raised the estimate minus did not (LSEG)", "slope", okv,
               lambda rr: EV("Revision up (1/0)", rr), lambda rr: EV("Outcome y (1 = beat)", rr), "+0.0%;-0.0%"),
              ("All_rev_conf_gap", "All", "Crowd adds to analysts? Beat rate, confident minus less sure, estimate raised only", "slope",
               lambda rr: f'{okv(rr)},{EV("Revision up (1/0)", rr)}=1',
               lambda rr: f'IF({EV("p: P(beat) at cut-off", rr)}>={A["HP"]},1,0)',
               lambda rr: EV("Outcome y (1 = beat)", rr), "+0.0%;-0.0%")]
    tests.append(("All_gap_before", "All", "Miss gap at the confident threshold, odds read before the release (EDGAR)", "slope",
                  lambda rr: (f'ISNUMBER({EV("Abnormal return", rr)}),ISNUMBER({EV("p: P(beat) at cut-off", rr)}),'
                              f'{EV("Outcome y (1 = beat)", rr)}=0,{EV("Odds read before the release (1/0)", rr)}=1'),
                  lambda rr: f'IF({EV("p: P(beat) at cut-off", rr)}>={A["HP"]},1,0)', lambda rr: EV("Abnormal return", rr), "+0.0%;-0.0%"))

    title(TC, "Test_Calc: one row per report for each test (formulas). In = 1 if the report is in the test. Score = the report's part of "
              "the standard error.", "Rows match Events. Values outside a test are 0 and drop out of the sums.", "A2:L2")
    tc_head = ["Week (Monday)"]
    tcol = {}
    for t_ in tests:
        key, kind = t_[0], t_[3]
        base = len(tc_head) + 1
        if kind == "mean":
            tc_head += [f"{key}: in", f"{key}: d", f"{key}: score"]
            tcol[key] = dict(i=L(base), d=L(base + 1), s=L(base + 2))
        else:
            tc_head += [f"{key}: in", f"{key}: x", f"{key}: y", f"{key}: score"]
            tcol[key] = dict(i=L(base), x=L(base + 1), y=L(base + 2), s=L(base + 3))
    head(TC, 4, tc_head, h=60)
    TCR = lambda c: f"Test_Calc!${c}$5:${c}${E1}"
    title(TST, "Tests: which findings hold, and at what confidence? (formulas)",
          "Standard errors are clustered by report week, so reports in the same earnings week count as one block of evidence. With "
          "fewer weeks than the Assumptions minimum (the new reports), each report is its own block. p-value from the t distribution.", "A2:P2")
    th_ = ["Sample", "Claim", "Threshold", "Reports", "Estimate", "Weeks with reports", "SE by week", "SE by report", "Method", "SE used",
           "t", "Degrees of freedom", "p-value", "Holds at level 1", "Holds at level 2", "Holds at level 3", "Mean of x (or d)", "Mean of y",
           "Sum of squares of x (or n)", "Intercept"]
    head(TST, 4, th_, h=45)
    for j, k in enumerate(("CL1", "CL2", "CL3")):
        put(TST, 3, 14 + j, f"={A[k]}", "0%", BOLD)
    put(TST, 3, 13, "Levels:", font=BOLD)
    trow = {}
    nweeks = int(((ev["report_date"].max() - ev["report_date"].min()).days // 7) + 3) if n else 1
    WE = 4 + nweeks
    title(WK, "Weeks: sums of the Test_Calc columns by report week (formulas)", "Each row is one Monday. Count = reports in the test that "
          "week. Score = sum of the reports' scores that week.", "A2:L2")
    wk_head = ["Week (Monday)"]
    wcol = {}
    for t_ in tests:
        wk_head += [f"{t_[0]}: reports", f"{t_[0]}: score"]
        wcol[t_[0]] = (L(len(wk_head) - 1), L(len(wk_head)))
    head(WK, 4, wk_head, h=60)
    put(WK, 5, 1, f"=MIN({TCR('A')})", "yyyy-mm-dd")
    for q in range(6, WE + 1):
        put(WK, q, 1, f"=A{q - 1}+7", "yyyy-mm-dd")
    WKR = lambda c: f"Weeks!${c}$5:${c}${WE}"
    for k_, t_ in enumerate(tests):
        key, smp, lab, kind = t_[:4]
        q = 5 + k_
        trow[key] = q
        tc = tcol[key]
        put(TST, q, 1, "All reports" if smp == "All" else "New reports only"); put(TST, q, 2, lab)
        if len(t_) > 8:
            put(TST, q, 3, f"={t_[8]}", "0.00")
        elif "gap" in key or "fall" in key:
            put(TST, q, 3, f"={A['HP']}", "0.00")
        put(TST, q, 4, f"=SUM({TCR(tc['i'])})", "0")
        if kind == "mean":
            put(TST, q, 17, f'=IFERROR(SUMPRODUCT({TCR(tc["i"])},{TCR(tc["d"])})/D{q},"n/a")', "0.0000")
            put(TST, q, 5, f"=Q{q}", t_[7]); put(TST, q, 19, f"=D{q}", "0")
            sc_ = tc["s"]
        else:
            put(TST, q, 17, f'=IFERROR(SUMPRODUCT({TCR(tc["i"])},{TCR(tc["x"])})/D{q},"n/a")', "0.0000")
            put(TST, q, 18, f'=IFERROR(SUMPRODUCT({TCR(tc["i"])},{TCR(tc["y"])})/D{q},"n/a")', "0.0000")
            put(TST, q, 19, f'=IFERROR(SUMPRODUCT({TCR(tc["i"])},({TCR(tc["x"])}-Q{q})^2),"n/a")', "0.0000")
            put(TST, q, 5, f'=IFERROR(SUMPRODUCT({TCR(tc["i"])},{TCR(tc["x"])}-Q{q},{TCR(tc["y"])})/S{q},"n/a")', t_[7])
            put(TST, q, 20, f'=IFERROR(R{q}-E{q}*Q{q},"n/a")', "0.0000")
            sc_ = tc["s"]
        wc, ws_ = wcol[key]
        put(TST, q, 6, f'=COUNTIF({WKR(wc)},">0")', "0")
        put(TST, q, 7, f'=IFERROR(SQRT(SUMSQ({WKR(ws_)}))/S{q}*SQRT(F{q}/(F{q}-1)),"n/a")', "0.0000")
        put(TST, q, 8, f'=IFERROR(SQRT(SUMSQ({TCR(sc_)}))/S{q}*SQRT(D{q}/(D{q}-1)),"n/a")', "0.0000")
        put(TST, q, 9, f'=IF(F{q}>={A["MINW"]},"by week","by report")')
        put(TST, q, 10, f'=IF(I{q}="by week",G{q},H{q})', "0.0000")
        put(TST, q, 11, f'=IFERROR(E{q}/J{q},"n/a")', "0.00")
        put(TST, q, 12, f'=IF(I{q}="by week",F{q}-1,D{q}-1)', "0")
        put(TST, q, 13, f'=IFERROR(TDIST(ABS(K{q}),L{q},2),"n/a")', "0.0000")
        for j, k in enumerate(("CL1", "CL2", "CL3")):
            put(TST, q, 14 + j, f'=IF(ISNUMBER(M{q}),IF(M{q}<1-{A[k]},"yes","no"),"n/a")')
        for q2 in range(5, WE + 1):
            WK[f"{wc}{q2}"] = f"=SUMIF({TCR('A')},$A{q2},{TCR(tc['i'])})"
            WK[f"{ws_}{q2}"] = f"=SUMIF({TCR('A')},$A{q2},{TCR(sc_)})"
    for i in range(n):
        rr = 5 + i
        TC[f"A{rr}"] = f"=Events!{C['Week (Monday)']}{rr}"; TC[f"A{rr}"].number_format = "yyyy-mm-dd"
        for t_ in tests:
            key, kind, cfn, xf, yf = t_[0], t_[3], t_[4], t_[5], t_[6]
            tc, q = tcol[key], trow[key]
            TC[f"{tc['i']}{rr}"] = f"=IF(AND({cfn(rr)}),1,0)"
            if kind == "mean":
                TC[f"{tc['d']}{rr}"] = f"=IF({tc['i']}{rr}=1,{xf(rr)},0)"
                TC[f"{tc['s']}{rr}"] = f"=IFERROR({tc['i']}{rr}*({tc['d']}{rr}-Tests!$Q${q}),0)"
            else:
                TC[f"{tc['x']}{rr}"] = f"=IF({tc['i']}{rr}=1,{xf(rr)},0)"
                TC[f"{tc['y']}{rr}"] = f"=IF({tc['i']}{rr}=1,{yf(rr)},0)"
                TC[f"{tc['s']}{rr}"] = (f"=IFERROR({tc['i']}{rr}*({tc['x']}{rr}-Tests!$Q${q})*({tc['y']}{rr}-Tests!$T${q}"
                                        f"-Tests!$E${q}*{tc['x']}{rr}),0)")
    TC.freeze_panes = "B5"; WK.freeze_panes = "B5"
    r = 5 + len(tests) + 1
    for t in ["Reading: 'Holds' means the p-value is below 1 minus the confidence level. The miss gap is shown at five thresholds on purpose: "
              "the threshold was fixed at the confident level before the new reports were seen, and the others show how much the result "
              "depends on where the line is drawn.",
              "The continuous test uses every miss and no threshold. If it is weaker than the threshold test, the effect sits at high p only."]:
        put(TST, r, 1, t, font=NOTE); r += 1
    TST.column_dimensions["A"].width = 16; TST.column_dimensions["B"].width = 70
    for c_ in "CDEFGHIJKLMNOP":
        TST.column_dimensions[c_].width = 11

    # ---------- Tail ----------
    title(TAIL, "Tail: same average, different risk?",
          "Part 1 feedback: the expected reactions are close (about -0.5% and -0.4%), so compare the downside. "
          "All cells are formulas on Events. Thresholds come from Assumptions.", "A2:F2")
    r = 4
    blocks = [("All reports", None, "AR if confident", "AR if less sure"),
              ("Part 1 sample", "Part 1", "AR if confident, Part 1", "AR if less sure, Part 1"),
              ("New reports only (out of sample)", "New", "AR if confident, new", "AR if less sure, new")]
    for lab, s, hc, hl in blocks:
        put(TAIL, r, 1, lab, font=SEC); r += 1
        head(TAIL, r, ["", "Crowd confident", "Crowd less sure", "Difference"]); r += 1
        cnd = lambda g: f"{Rc},{g}{cond(s)}"
        Rf = R("Fall (1/0)")
        stats = [
            ("Reports with a share move", lambda g, h: f'=COUNT({R(h)})', "0"),
            ("Crowd's probability of a miss (1 - average p)", lambda g, h: f'=IFERROR(1-AVERAGEIFS({Rp},{cnd(g)}),"n/a")', PC1),
            ("Share that missed", lambda g, h: f'=IFERROR(1-AVERAGEIFS({Ry},{cnd(g)}),"n/a")', PC1),
            ("Average move", lambda g, h: f'=IFERROR(AVERAGE({R(h)}),"n/a")', PCT),
            ("Average move on a beat", lambda g, h: f'=IFERROR(AVERAGEIFS({Rar},{cnd(g)},{Ry},1),"n/a")', PCT),
            ("Average move on a miss", lambda g, h: f'=IFERROR(AVERAGEIFS({Rar},{cnd(g)},{Ry},0),"n/a")', PCT),
            ("Standard deviation of the move", lambda g, h: f'=IFERROR(STDEV({R(h)}),"n/a")', PC1),
            ("Share that fell 5% or more against the market", lambda g, h: f'=IFERROR(COUNTIF({R(h)},"<="&{A["FALL"]})/COUNT({R(h)}),"n/a")', PC1),
            ("   after a beat", lambda g, h: f'=IFERROR(COUNTIFS({cnd(g)},{Ry},1,{Rf},1)/COUNTIFS({cnd(g)},{Ry},1,{Rf},">=0"),"n/a")', PC1),
            ("   after a miss", lambda g, h: f'=IFERROR(COUNTIFS({cnd(g)},{Ry},0,{Rf},1)/COUNTIFS({cnd(g)},{Ry},0,{Rf},">=0"),"n/a")', PC1),
            ("Falls of 5% or more that came after a beat (share of all falls)",
             lambda g, h: f'=IFERROR(COUNTIFS({cnd(g)},{Ry},1,{Rf},1)/COUNTIFS({cnd(g)},{Rf},1),"n/a")', PC1),
            ("Share that fell 10% or more", lambda g, h: f'=IFERROR(COUNTIF({R(h)},"<="&{A["BIG"]})/COUNT({R(h)}),"n/a")', PC1),
            ("10th percentile of the move", lambda g, h: f'=IFERROR(PERCENTILE({R(h)},0.1),"n/a")', PCT),
            ("Average of the worst 10% of moves", lambda g, h: f'=IFERROR(AVERAGEIF({R(h)},"<="&PERCENTILE({R(h)},0.1)),"n/a")', PCT)]
        for name, fn, nf in stats:
            put(TAIL, r, 1, name)
            put(TAIL, r, 2, fn(1, hc), nf); put(TAIL, r, 3, fn(0, hl), nf)
            put(TAIL, r, 4, f'=IFERROR(B{r}-C{r},"n/a")', nf)
            r += 1
        r += 1
    put(TAIL, r, 1, "Bootstrap intervals for these differences are on the Bootstrap sheet.", font=NOTE)
    TAIL.column_dimensions["A"].width = 62
    for c_ in "BCD":
        TAIL.column_dimensions[c_].width = 16

    # ---------- Hedge ----------
    title(HED, "Hedge: what is a protective put worth before a report?",
          "Value of the hedge = P(miss) x average payoff on a miss + P(beat) x average payoff on a beat - put price. P(miss) is "
          "the crowd's 1 - p. Payoffs come from past reports in the same group (Events). The price comes from the snapshots, "
          "or from the Assumptions input until snapshots exist.", "A2:F2")
    r = 4
    head(HED, r, ["", "Crowd confident", "Crowd less sure"]); r += 1
    h0 = r
    Rpay, Rsp = R("Put payoff (share of S0)"), R("Put price / share price (snapshot)")
    hl = [("Reports with a payoff", lambda g: f'=COUNTIFS({Rc},{g},{Rpay},">=0")', "0"),
          ("Crowd's probability of a miss (1 - average p)", lambda g: f'=IFERROR(1-AVERAGEIFS({Rp},{Rc},{g}),"n/a")', PC1),
          ("Average put payoff when the company missed", lambda g: f'=IFERROR(AVERAGEIFS({Rpay},{Rc},{g},{Ry},0),"n/a")', PC2),
          ("Average put payoff when the company beat", lambda g: f'=IFERROR(AVERAGEIFS({Rpay},{Rc},{g},{Ry},1),"n/a")', PC2),
          ("Expected payoff = P(miss) x payoff on a miss + P(beat) x payoff on a beat", None, PC2),
          ("Average payoff actually paid (check)", lambda g: f'=IFERROR(AVERAGEIFS({Rpay},{Rc},{g}),"n/a")', PC2),
          ("Reports with a snapshot put price", lambda g: f'=COUNTIFS({Rc},{g},{Rsp},">=0")', "0"),
          ("Average put price in snapshots (share of the share price)", lambda g: f'=IFERROR(AVERAGEIFS({Rsp},{Rc},{g}),"no snapshots yet")', PC2),
          ("Price used: snapshots, or the Assumptions input if none", None, PC2),
          ("Value of the hedge per report (expected payoff - price)", None, "+0.00%;-0.00%"),
          ("Break-even price: the most worth paying for the put", None, PC2)]
    for k, (lab, fn, nf) in enumerate(hl):
        put(HED, r, 1, lab)
        for j, g in enumerate((1, 0)):
            c_ = L(2 + j)
            v = {4: f'=IFERROR({c_}{h0+1}*{c_}{h0+2}+(1-{c_}{h0+1})*{c_}{h0+3},"n/a")',
                 8: f'=IF(ISNUMBER({c_}{h0+7}),{c_}{h0+7},{A["COST"]})',
                 9: f'=IFERROR({c_}{h0+4}-{c_}{h0+8},"n/a")', 10: f"={c_}{h0+4}"}.get(k) or fn(g)
            put(HED, r, 2 + j, v, nf)
        r += 1
    H_LM, H_LB, H_PR = h0 + 2, h0 + 3, h0 + 8
    r += 1
    put(HED, r, 1, "Three choices for a manager holding the share through a confident report (past reports)", font=SEC); r += 1
    head(HED, r, ["", "Hold", "Sell half before", "Hold + buy the put"]); r += 1
    Rrc = R("Share return if confident")
    put(HED, r, 1, "Average share return (simple)")
    put(HED, r, 2, f'=IFERROR(AVERAGE({Rrc}),"n/a")', PCT); put(HED, r, 3, f'=IFERROR(B{r}/2,"n/a")', PCT)
    put(HED, r, 4, f'=IFERROR(B{r}+B{h0+5}-B{H_PR},"n/a")', PCT); r += 1
    put(HED, r, 1, "Worst single report")
    put(HED, r, 2, f'=IFERROR(MIN({Rrc}),"n/a")', PCT); put(HED, r, 3, f'=IFERROR(B{r}/2,"n/a")', PCT)
    put(HED, r, 4, f'=IFERROR(MAX(B{r},-{A["OTM"]})-B{H_PR},"n/a")', PCT); r += 1
    put(HED, r, 1, "Most a report can cost"); put(HED, r, 2, "unlimited"); put(HED, r, 3, "half the fall")
    put(HED, r, 4, f'=-{A["OTM"]}-B{H_PR}', PCT); r += 2
    for t in ["Positive value: on average the put pays more than it costs. Negative value: the put still caps the loss, and the price is the "
              "cost of that insurance.",
              "The payoff is measured at the close after the report and ignores remaining time value, so it understates the hedge."]:
        put(HED, r, 1, t, font=NOTE); r += 1
    HED.column_dimensions["A"].width = 70
    for c_ in "BCD":
        HED.column_dimensions[c_].width = 17

    # ---------- Live ----------
    title(LIVE, "Live: decision card for each upcoming report (latest snapshot)",
          "Row pointer = latest snapshot of each upcoming contract (set by the script). Everything else is a formula. "
          "Expected payoff uses the crowd's probability now and the Hedge sheet. Value = expected payoff - today's put price.", "A2:N2")
    lh = ["Snapshot row", "Ticker", "Report date", "Snapshot (UTC)", "p now", "Group", "P(miss) = 1 - p", "Expected put payoff",
          "Put price / share price (mid quote; last trade when the market is closed)", "Value of the hedge", "Option-implied move",
          "Company's average past |move|",
          "Analyst revisions: net up / analysts, 30 days", "Suggested action"]
    head(LIVE, 4, lh)
    today = pd.Timestamp.now(tz="UTC").tz_localize(None).normalize()
    lv = snap[pd.to_datetime(snap["report_date"]) >= today] if len(snap) else snap
    if len(lv):
        lv = lv.sort_values("snap_ts").drop_duplicates("slug", keep="last")
        lv = lv.sort_values("volume", ascending=False).drop_duplicates(["ticker", "report_date"]).sort_values(["report_date", "ticker"])
    for i, idx in enumerate(lv.index if len(lv) else []):
        r = 5 + i
        sx = lambda k: f"INDEX({SN(k)},A{r})"
        has = lambda *ks: f"COUNT({','.join(sx(k) for k in ks)})={len(ks)}"
        put(LIVE, r, 1, int(idx) + 1)
        put(LIVE, r, 2, f"={sx('ticker')}"); put(LIVE, r, 3, f"={sx('report_date')}", "yyyy-mm-dd"); put(LIVE, r, 4, f"={sx('snap_utc')}")
        put(LIVE, r, 5, f'=IF({has("p_mid")},{sx("p_mid")},{sx("p_gamma")})', "0%")
        put(LIVE, r, 6, f'=IF(ISNUMBER(E{r}),IF(E{r}>={A["HP"]},"confident","less sure"),"")')
        put(LIVE, r, 7, f'=IF(ISNUMBER(E{r}),1-E{r},"")', "0%")
        put(LIVE, r, 8, f'=IFERROR(IF(F{r}="confident",G{r}*Hedge!$B${H_LM}+(1-G{r})*Hedge!$B${H_LB},'
                        f'G{r}*Hedge!$C${H_LM}+(1-G{r})*Hedge!$C${H_LB}),"")', PC2)
        put(LIVE, r, 9, f'=IF(AND({has("put95_cost", "put95_strike", "spot")},{sx("put95_strike")}<={sx("spot")}),{sx("put95_cost")},'
                        f'IF(AND({has("put95_last", "put95_strike", "spot")},{sx("put95_strike")}<={sx("spot")}),{sx("put95_last")},""))', PC2)
        put(LIVE, r, 10, f'=IF(AND(ISNUMBER(H{r}),ISNUMBER(I{r})),H{r}-I{r},"")', "+0.00%;-0.00%")
        put(LIVE, r, 11, f'=IF({has("implied_move")},{sx("implied_move")},"")', PC1)
        put(LIVE, r, 12, f'=IF(COUNTIFS({RE_("A")},B{r},{RE_("I")},">=0")>={A["MINM"]},AVERAGEIFS({RE_("I")},{RE_("A")},B{r}),"")', PC1)
        put(LIVE, r, 13, f'=IF({has("rev_up_30d", "rev_down_30d", "n_analysts")},({sx("rev_up_30d")}-{sx("rev_down_30d")})/MAX({sx("n_analysts")},1),"")', "+0.00;-0.00")
        put(LIVE, r, 14, f'=IF(F{r}<>"confident","no flag",IF(NOT(ISNUMBER(J{r})),"flag: confident crowd, no put price",'
                         f'IF(J{r}>0,"flag: put looks cheap, consider hedging","flag: put is dear, consider trimming")))')
    if not len(lv):
        put(LIVE, 5, 1, "No upcoming reports in the snapshots yet. Run the script during earnings season.", font=NOTE)
    for c_, w in zip("ABCDEFGHIJKLMN", (8, 8, 11, 16, 7, 10, 9, 11, 11, 11, 11, 12, 12, 40)):
        LIVE.column_dimensions[c_].width = w

    # ---------- AI layer: Raw_AI_Live, Raw_GenAI, Live columns O to R, GenAI sheet ----------
    RAL["A1"] = ("AI layer for upcoming reports, written by the script: boosted-tree downside model (1-in-10 worst case) and |move| model "
                 "trained on all settled reports, the AI cut-off (same turnover as the rule), and a client note written by generative AI "
                 "from these numbers only."); RAL["A1"].font = BOLD
    alc = ["slug", "ticker", "report_date", "p", "ai_q10", "ai_move", "ai_cutoff", "rule_action", "note", "note_model", "drivers", "guidance", "summary"]
    head(RAL, 4, ["Contract slug", "Ticker", "Report date", "Crowd p", "AI 10% worst case", "AI expected |move|", "AI cut-off",
                  "Rule action", "Client note (generative AI)", "Generative AI model", "Drivers of the worst case (model, one input at a time)",
                  "Last quarter's guidance (generative AI)", "Last quarter's release in one line (generative AI)"])
    for i, rec in enumerate(al.reindex(columns=alc).to_dict("records") if len(al) else []):
        for j, k in enumerate(alc):
            v = clean(rec[k])
            if k == "report_date" and v:
                v = pd.Timestamp(v).to_pydatetime()
            put(RAL, 5 + i, 1 + j, v, {"report_date": "yyyy-mm-dd", "p": "0%", "ai_q10": "+0.0%;-0.0%", "ai_move": PC1,
                                       "ai_cutoff": "+0.0%;-0.0%"}.get(k))
    ALE = 4 + max(len(al), 1)
    AL_ = lambda c_: f"Raw_AI_Live!${c_}$5:${c_}${ALE}"
    for c_, txt in zip("OPQR", ("AI 10% worst case", "AI expected |move|", "AI action (formula: worst case at or below the AI cut-off)",
                               "Client note (generative AI)")):
        LIVE[f"{c_}4"] = txt; LIVE[f"{c_}4"].font = BOLD; LIVE[f"{c_}4"].alignment = Alignment(wrap_text=True, vertical="top")
    for i, idx in enumerate(lv.index if len(lv) else []):
        r = 5 + i
        mt = f'MATCH(INDEX({SN("slug")},A{r}),{AL_("A")},0)'
        put(LIVE, r, 15, f'=IFERROR(INDEX({AL_("E")},{mt}),"")', "+0.0%;-0.0%")
        put(LIVE, r, 16, f'=IFERROR(INDEX({AL_("F")},{mt}),"")', PC1)
        put(LIVE, r, 17, f'=IF(ISNUMBER(O{r}),IF(O{r}<=INDEX({AL_("G")},{mt}),"AI: trim {TRIM:.0%}","AI: hold"),"")')
        put(LIVE, r, 18, f'=IFERROR(IF(INDEX({AL_("I")},{mt})=0,"",INDEX({AL_("I")},{mt})),"")')
        LIVE.cell(row=r, column=18).alignment = Alignment(wrap_text=True, vertical="top")
    for c_, w in zip("OPQR", (11, 11, 14, 70)):
        LIVE.column_dimensions[c_].width = w

    RGA["A1"] = ("Generative AI reading of a random sample of contracts, written by the script: Polymarket (question and rules text) and "
                 "Kalshi KPI contracts (title only: company name, metric, number and quarter in free text). Columns A to G as returned. "
                 "Columns H to O are formulas that compare it with our rule-based parser (Raw_Markets, Raw_K_Markets)."); RGA["A1"].font = BOLD
    gac = ["slug", "ai_ticker", "ai_report_date", "ai_eps_type", "ai_threshold", "model", "source"]
    head(RGA, 4, ["Contract", "AI: ticker", "AI: report date", "AI: EPS type (Polymarket) or metric (Kalshi)", "AI: threshold",
                  "Generative AI model", "Source", "Parser: ticker", "Parser: report date", "Parser: EPS type", "Parser: threshold",
                  "Ticker matches", "Date matches", "EPS type matches", "Threshold matches"], h=45)
    MKE = 4 + max(len(mk), 1)
    KME = 4 + max(len(km), 1)
    MK_ = lambda c_: f"Raw_Markets!${c_}$5:${c_}${MKE}"
    KM_ = lambda c_: f"Raw_K_Markets!${c_}$5:${c_}${KME}"
    if len(ga) and "source" not in ga.columns:
        ga["source"] = "Polymarket"
    for i, rec in enumerate(ga.reindex(columns=gac).to_dict("records") if len(ga) else []):
        q = 5 + i
        for j, k in enumerate(gac):
            put(RGA, q, 1 + j, clean(rec[k]) if k != "source" else (clean(rec[k]) or "Polymarket"))
        mp, mkk = f"MATCH(A{q},{MK_('A')},0)", f"MATCH(A{q},{KM_('A')},0)"
        kal = f'G{q}="Kalshi"'
        put(RGA, q, 8, f'=IFERROR(IF({kal},INDEX({KM_("J")},{mkk}),INDEX({MK_("C")},{mp})),"")')
        put(RGA, q, 9, f'=IF({kal},"",IFERROR(INDEX({MK_("D")},{mp}),""))', "yyyy-mm-dd")
        put(RGA, q, 10, f'=IF({kal},"",IFERROR(INDEX({MK_("F")},{mp}),""))')
        put(RGA, q, 11, f'=IFERROR(IF({kal},INDEX({KM_("E")},{mkk}),INDEX({MK_("E")},{mp})),"")')
        put(RGA, q, 12, f'=IF(AND(B{q}<>"",H{q}<>""),IF(UPPER(SUBSTITUTE(B{q},"-","."))=UPPER(SUBSTITUTE(H{q},"-",".")),1,0),"")')
        put(RGA, q, 13, f'=IF(AND(C{q}<>"",ISNUMBER(I{q})),IF(TEXT(I{q},"yyyy-mm-dd")=C{q}&"",1,0),"")')
        put(RGA, q, 14, f'=IF(AND(D{q}<>"",J{q}<>""),IF(LOWER(SUBSTITUTE(D{q},"-",""))=LOWER(SUBSTITUTE(J{q},"-","")),1,0),"")')
        put(RGA, q, 15, f'=IF(AND(ISNUMBER(E{q}),ISNUMBER(K{q})),IF(ABS(E{q}-K{q})<=MAX(0.005,0.001*ABS(K{q})),1,0),"")')
    GAE = 4 + max(len(ga), 1)
    title(GAI, "GenAI: where the generative AI is needed, and how well it does",
          "1) Read contracts. New prediction-market contracts are listed every week in free text. Polymarket's are regular enough for "
          "hand-written rules. Kalshi's KPI contracts are not: company names, metrics, units ('42.5 million', '870 thousand') and quarters "
          "vary, and the ticker is not in the text. Checked here against our parser. 2) Explain each flag to the client: the note names "
          "the drivers behind the model's number (which inputs moved the forecast most), written from the numbers only.", "A2:F2")
    GAI.row_dimensions[2].height = 75
    head(GAI, 4, ["Contract reading: generative AI against our parser", "Polymarket: checked", "Polymarket: agreement",
                  "Kalshi KPI: checked", "Kalshi KPI: agreement"])
    for k, (lab, c_) in enumerate((("Ticker (Kalshi: from the company name)", "L"), ("Report date", "M"),
                                   ("EPS type (GAAP / non-GAAP)", "N"), ("Threshold (Kalshi: number with its unit)", "O"))):
        rg, sg = f"Raw_GenAI!${c_}$5:${c_}${GAE}", f"Raw_GenAI!$G$5:$G${GAE}"
        put(GAI, 5 + k, 1, lab)
        for j, src in enumerate(("Polymarket", "Kalshi")):
            put(GAI, 5 + k, 2 + 2 * j, f'=COUNTIFS({sg},"{src}",{rg},">=0")', "0")
            put(GAI, 5 + k, 3 + 2 * j, f'=IFERROR(AVERAGEIFS({rg},{sg},"{src}"),"n/a")', PC1)
    put(GAI, 10, 1, "Disagreements are listed on Raw_GenAI (columns L to O = 0). Check each against the contract text: the parser can "
                    "be the one that is wrong.", font=NOTE)
    gq = 12
    put(GAI, gq, 1, "3) Read last quarter's earnings release (SEC EDGAR) before each report. What happened next, by what the AI read:", font=SEC)
    head(GAI, gq + 1, ["Last quarter's guidance (generative AI)", "Reports", "Beat rate", "Average abnormal return",
                       f"Share falling {abs(FALL):.0%} or more", "Average crowd p"])
    Rg = R("Last release: guidance (generative AI)")
    for k, g_ in enumerate(("raised", "maintained", "not given", "lowered", "withdrawn")):
        q = gq + 2 + k
        put(GAI, q, 1, g_); put(GAI, q, 2, f'=COUNTIFS({Rg},"{g_}",{Rar},">-100")', "0")
        put(GAI, q, 3, f'=IFERROR(AVERAGEIFS({Ry},{Rg},"{g_}",{Rar},">-100"),"n/a")', PC1)
        put(GAI, q, 4, f'=IFERROR(AVERAGEIFS({Rar},{Rg},"{g_}"),"n/a")', "+0.0%;-0.0%")
        put(GAI, q, 5, f'=IFERROR(COUNTIFS({Rg},"{g_}",{Rar},"<="&{A["FALL"]})/B{q},"n/a")', PC1)
        put(GAI, q, 6, f'=IFERROR(AVERAGEIFS({Rp},{Rg},"{g_}",{Rar},">-100"),"n/a")', "0.00")
    put(GAI, gq + 7, 1, "Does the reading add to the models? Models sheet: AI with and without the reading, and crowd + reading against "
                        "crowd + a keyword count of the same text (the spreadsheet way).", font=NOTE)
    head(GAI, gq + 9, ["Client notes for this week's reports (Live sheet, column R)", "Ticker", "Report date", "AI action", "Note"])
    with_note = [i for i, v in enumerate(al["note"]) if isinstance(v, str) and v.strip()] if len(al) else []
    for q, i in enumerate(with_note[:40], start=gq + 10):
        put(GAI, q, 2, f"=Raw_AI_Live!B{5 + i}"); put(GAI, q, 3, f"=Raw_AI_Live!C{5 + i}", "yyyy-mm-dd")
        put(GAI, q, 4, f'=IF(Raw_AI_Live!E{5 + i}<=Raw_AI_Live!G{5 + i},"trim {TRIM:.0%}","hold")')
        put(GAI, q, 5, f'=IF(Raw_AI_Live!I{5 + i}=0,"",Raw_AI_Live!I{5 + i})')
        GAI.cell(row=q, column=5).alignment = Alignment(wrap_text=True, vertical="top")
    if not len(al):
        put(GAI, gq + 10, 1, "No upcoming reports scored yet (needs today's snapshot).", font=NOTE)
    elif not with_note:
        put(GAI, gq + 10, 1, "No client notes yet: generative AI was not available in this run (Colab AI or GEMINI_API_KEY).", font=NOTE)
    GAI.column_dimensions["A"].width = 50; GAI.column_dimensions["E"].width = 90
    for c_ in "BCD":
        GAI.column_dimensions[c_].width = 13

    # ---------- Hedge: options against history (needs Live) ----------
    LV1 = 4 + max(len(lv), 1)
    LVR = lambda c_: f"Live!${c_}$5:${c_}${LV1}"
    r = HED.max_row + 2
    put(HED, r, 1, "Options against history: does the options market treat confident reports as safer?", font=SEC); r += 1
    put(HED, r, 1, "If puts are cheaper than the history-based payoff when the crowd is confident, the put is good value there. "
                   "Implied move above the company's past move = options price a bigger move than usual.", font=NOTE); r += 1
    head(HED, r, ["Upcoming reports (Live sheet)", "Crowd confident", "Crowd less sure"]); r += 1
    o0 = r
    up = [("Reports with a put price", lambda g: f'=COUNTIFS({LVR("F")},"{g}",{LVR("I")},">=0")', "0"),
          ("Average put price (share of the share price)", lambda g: f'=IFERROR(AVERAGEIFS({LVR("I")},{LVR("F")},"{g}",{LVR("I")},">=0"),"n/a")', PC2),
          ("Average expected payoff from history (same reports)", lambda g: f'=IFERROR(AVERAGEIFS({LVR("H")},{LVR("F")},"{g}",{LVR("I")},">=0"),"n/a")', PC2),
          ("Price / expected payoff (below 1 = cheap)", None, "0.00"),
          ("Share of reports where the put looks cheap", lambda g: f'=IFERROR(COUNTIFS({LVR("F")},"{g}",{LVR("J")},">0")/B{o0},"n/a")', PC1),
          ("Average option-implied move", lambda g: f'=IFERROR(AVERAGEIFS({LVR("K")},{LVR("F")},"{g}",{LVR("K")},">=0",{LVR("L")},">=0"),"n/a")', PC1),
          ("Average company past move (same reports)", lambda g: f'=IFERROR(AVERAGEIFS({LVR("L")},{LVR("F")},"{g}",{LVR("K")},">=0",{LVR("L")},">=0"),"n/a")', PC1),
          ("Implied move / past move (above 1 = options expect more)", None, "0.00")]
    for k, (lab, fn, nf) in enumerate(up):
        put(HED, r, 1, lab)
        for j, g in enumerate(("confident", "less sure")):
            c_ = L(2 + j)
            v = {3: f'=IFERROR({c_}{o0+1}/{c_}{o0+2},"n/a")', 7: f'=IFERROR({c_}{o0+5}/{c_}{o0+6},"n/a")'}.get(k) or fn(g).replace(f"B{o0}", f"{c_}{o0}")
            put(HED, r, 2 + j, v, nf)
        r += 1
    r += 1
    head(HED, r, ["Settled reports with a snapshot (Events)", "Crowd confident", "Crowd less sure"]); r += 1
    s0 = r
    Rpnl, Rimr = R("Hedge P&L: payoff - snapshot put price"), R("Option-implied move - realised |share return|")
    st_ = [("Reports with a put price", lambda g: f'=COUNTIFS({Rc},{g},{Rsp},">=0")', "0"),
           ("Average put price paid", lambda g: f'=IFERROR(AVERAGEIFS({Rsp},{Rc},{g},{Rsp},">=0"),"n/a")', PC2),
           ("Average payoff received", lambda g: f'=IFERROR(AVERAGEIFS({Rpay},{Rc},{g},{Rsp},">=0"),"n/a")', PC2),
           ("Average hedge P&L (payoff - price)", lambda g: f'=IFERROR(AVERAGEIFS({Rpnl},{Rc},{g}),"n/a")', "+0.00%;-0.00%"),
           ("Share of reports where the put paid more than it cost", lambda g: f'=IFERROR(COUNTIFS({Rc},{g},{Rpnl},">0")/COUNTIFS({Rc},{g},{Rpnl},">-1"),"n/a")', PC1),
           ("Average implied move - realised |share return| (above 0 = options overpriced the move)",
            lambda g: f'=IFERROR(AVERAGEIFS({Rimr},{Rc},{g}),"n/a")', "+0.0%;-0.0%")]
    for lab, fn, nf in st_:
        put(HED, r, 1, lab)
        for j, g in enumerate((1, 0)):
            put(HED, r, 2 + j, fn(g), nf)
        r += 1
    H_OPT, H_SET = o0, s0
    r += 1
    head(HED, r, ["Settled reports with past option prices (Massive)", "Crowd confident", "Crowd less sure"]); r += 1
    hh0 = r
    Rph, Rvh, Rplh, Rimh = (R(x) for x in ("Put price / share price (history)", "Put value after / share price (history)",
                                           "Hedge P&L with actual option prices (history)", "Implied move (history) - realised |share return|"))
    hh = [("Reports with a past put price", lambda g: f'=COUNTIFS({Rc},{g},{Rph},">=0")', "0"),
          ("Average put price paid (close the day before)", lambda g: f'=IFERROR(AVERAGEIFS({Rph},{Rc},{g}),"n/a")', PC2),
          ("Average put value after the report (close the day after)", lambda g: f'=IFERROR(AVERAGEIFS({Rvh},{Rc},{g},{Rph},">=0"),"n/a")', PC2),
          ("Average hedge P&L with actual option prices", lambda g: f'=IFERROR(AVERAGEIFS({Rplh},{Rc},{g}),"n/a")', "+0.00%;-0.00%"),
          ("Share of reports where the put gained", lambda g: f'=IFERROR(COUNTIFS({Rc},{g},{Rplh},">0")/COUNTIFS({Rc},{g},{Rplh},">-1"),"n/a")', PC1),
          ("Average implied move - realised |share return| (above 0 = options overpriced the move)",
           lambda g: f'=IFERROR(AVERAGEIFS({Rimh},{Rc},{g}),"n/a")', "+0.0%;-0.0%")]
    for lab, fn, nf in hh:
        put(HED, r, 1, lab)
        for j, g in enumerate((1, 0)):
            put(HED, r, 2 + j, fn(g), nf)
        r += 1
    H_HIST = hh0
    r += 1
    head(HED, r, ["Settled reports with LSEG implied volatility", "Crowd confident", "Crowd less sure"]); r += 1
    hl0 = r
    Rivb, Riva, Riml, Rimlr = (R(x) for x in ("IV before (LSEG)", "IV after (LSEG)", "Implied event move (LSEG)",
                                              "Implied event move (LSEG) - realised |share return|"))
    hlr = [("Reports with LSEG implied volatility", lambda g: f'=COUNTIFS({Rc},{g},{Riml},">=0")', "0"),
           ("Average implied volatility before the report", lambda g: f'=IFERROR(AVERAGEIFS({Rivb},{Rc},{g},{Riml},">=0"),"n/a")', PC1),
           ("Average implied volatility after the report", lambda g: f'=IFERROR(AVERAGEIFS({Riva},{Rc},{g},{Riml},">=0"),"n/a")', PC1),
           ("Average implied event move", lambda g: f'=IFERROR(AVERAGEIFS({Riml},{Rc},{g}),"n/a")', PC1),
           ("Average implied event move - realised |share return| (above 0 = options overpriced the move)",
            lambda g: f'=IFERROR(AVERAGEIFS({Rimlr},{Rc},{g}),"n/a")', "+0.0%;-0.0%")]
    for lab, fn, nf in hlr:
        put(HED, r, 1, lab)
        for j, g in enumerate((1, 0)):
            put(HED, r, 2 + j, fn(g), nf)
        r += 1
    H_LSEG = hl0
    r += 1
    head(HED, r, [f"Protective put priced from LSEG implied volatility, all settled reports ({IV_DAYS}-day put, strike "
                  f"{1 - PUT_OTM:.0%} of the share price, bought at the close before, sold at the close after)", "Crowd confident",
                  "Crowd less sure"], h=60); r += 1
    hp0 = r
    Rlp, Rlv, Rlpl = R("Put price / share price (LSEG IV)"), R("Put value after / share price (LSEG IV)"), R("Hedge P&L with LSEG-priced put")
    hpr = [("Reports with an LSEG put price", lambda g: f'=COUNTIFS({Rc},{g},{Rlpl},">-1")', "0"),
           ("Average put price (share of the share price)", lambda g: f'=IFERROR(AVERAGEIFS({Rlp},{Rc},{g},{Rlpl},">-1"),"n/a")', PC2),
           ("Average put value after the report", lambda g: f'=IFERROR(AVERAGEIFS({Rlv},{Rc},{g},{Rlpl},">-1"),"n/a")', PC2),
           ("Average hedge P&L (value after - price)", lambda g: f'=IFERROR(AVERAGEIFS({Rlpl},{Rc},{g}),"n/a")', "+0.00%;-0.00%"),
           ("Share of reports where the put gained", lambda g: f'=IFERROR(COUNTIFS({Rc},{g},{Rlpl},">0")/COUNTIFS({Rc},{g},{Rlpl},">-1"),"n/a")', PC1),
           ("Average hedge P&L on misses", lambda g: f'=IFERROR(AVERAGEIFS({Rlpl},{Rc},{g},{Ry},0),"n/a")', "+0.00%;-0.00%"),
           ("Average hedge P&L on beats", lambda g: f'=IFERROR(AVERAGEIFS({Rlpl},{Rc},{g},{Ry},1),"n/a")', "+0.00%;-0.00%")]
    for lab, fn, nf in hpr:
        put(HED, r, 1, lab)
        for j, g in enumerate((1, 0)):
            put(HED, r, 2 + j, fn(g), nf)
        r += 1
    put(HED, r, 1, "Put priced with Black-Scholes at the LSEG 30-day at-the-money implied volatility (no skew, no rates). The value "
                   "after uses the implied volatility after the report and the days left.", font=NOTE); r += 2
    H_LSP = hp0
    head(HED, r, ["Crowd vs analysts: LSEG consensus EPS revision before the report, all settled reports", "Crowd confident",
                  "Crowd less sure"], h=45); r += 1
    hv0 = r
    Rrv, Rup = R("Analyst revision (LSEG)"), R("Revision up (1/0)")
    hvr = [("Reports with an analyst revision", lambda g: f'=COUNTIFS({Rc},{g},{Rup},">=0")', "0"),
           ("Share of reports where analysts raised the estimate", lambda g: f'=IFERROR(AVERAGEIFS({Rup},{Rc},{g}),"n/a")', PC1),
           ("Beat rate when analysts raised the estimate", lambda g: f'=IFERROR(AVERAGEIFS({Ry},{Rc},{g},{Rup},1),"n/a")', PC1),
           ("Beat rate when analysts did not raise it", lambda g: f'=IFERROR(AVERAGEIFS({Ry},{Rc},{g},{Rup},0),"n/a")', PC1),
           ("Average move on a miss, estimate raised", lambda g: f'=IFERROR(AVERAGEIFS({Rar},{Rc},{g},{Rup},1,{Ry},0),"n/a")', PCT),
           ("Average move on a miss, estimate not raised", lambda g: f'=IFERROR(AVERAGEIFS({Rar},{Rc},{g},{Rup},0,{Ry},0),"n/a")', PCT)]
    for lab, fn, nf in hvr:
        put(HED, r, 1, lab)
        for j, g in enumerate((1, 0)):
            put(HED, r, 2 + j, fn(g), nf)
        r += 1
    H_REV = hv0
    put(HED, r + 1, 1, "These rows fill as daily snapshots build up and the snapshotted reports settle.", font=NOTE)

    # ---------- Models ----------
    title(MOD, "Models: can anything beat the crowd out of sample?",
          "Forecasts are on Raw_Models (Python, walk-forward). Errors are on Model_Calc. Every cell here is a formula. A model is kept "
          "only if its error is lower than its benchmark's on the same reports. Manager's data = company beat record, run-up, "
          "volatility, average past move, GAAP flag, contract volume.", "A2:H2")
    head(MOD, 4, ["Model", "Benchmark", "Reports (out of sample)", "Model error", "Benchmark error", "Skill vs benchmark", "Decision"])
    bl = {"y": "Crowd p (Brier)", "fall5": "Crowd p alone, logistic (Brier)", "abs_ar": "Company's average past move (mean abs. error)",
          "ar": "10th percentile of earlier reports (pinball loss)"}
    bl_col = {"beat_crowd_sm_lin": "StarMine model alone (Brier)", "q10_hist": "same model (pinball loss)",
              "q10_resid_gb": "Spreadsheet k x past move (pinball loss)", "q10_kalshi_gb": "Same AI without Kalshi (pinball loss)"}
    r = 5
    mrow = {}
    for col, lab, tgt, *_ in MODEL_SPECS:
        if col in ("fall_crowd_lin", "q10_hist", "fall10_gb"):
            continue
        mrow[col] = r
        ce, cb = MR(f"{lab}: error"), MR(f"{lab}: benchmark error")
        put(MOD, r, 1, lab); put(MOD, r, 2, bl_col.get(col, bl[tgt])); put(MOD, r, 3, f"=COUNT({ce})", "0")
        put(MOD, r, 4, f'=IFERROR(AVERAGE({ce}),"n/a")', "0.0000"); put(MOD, r, 5, f'=IFERROR(AVERAGE({cb}),"n/a")', "0.0000")
        put(MOD, r, 6, f'=IFERROR(1-D{r}/E{r},"n/a")', PCT)
        put(MOD, r, 7, f'=IF(C{r}<50,"too few reports",IF(D{r}<E{r},"keep","drop: benchmark is better"))')
        r += 1
    mrow["options"] = r
    ce, cb = MR("Option-implied move: error"), MR("Option-implied move: benchmark error (company's past move)")
    put(MOD, r, 1, "|Move|: option-implied move (snapshots, against the |share return|)"); put(MOD, r, 2, bl["abs_ar"])
    put(MOD, r, 3, f"=COUNT({ce})", "0"); put(MOD, r, 4, f'=IFERROR(AVERAGE({ce}),"n/a")', "0.0000")
    put(MOD, r, 5, f'=IFERROR(AVERAGE({cb}),"n/a")', "0.0000"); put(MOD, r, 6, f'=IFERROR(1-D{r}/E{r},"n/a")', PCT)
    put(MOD, r, 7, f'=IF(C{r}<30,"collecting snapshots",IF(D{r}<E{r},"options beat the company record","company record is better"))')
    r += 1
    mrow["options_hist"] = r
    ce, cb = MR("Implied move (history): error"), MR("Implied move (history): benchmark error (company's past move)")
    put(MOD, r, 1, "|Move|: option-implied move from past option prices (Massive), against the |share return|"); put(MOD, r, 2, bl["abs_ar"])
    put(MOD, r, 3, f"=COUNT({ce})", "0"); put(MOD, r, 4, f'=IFERROR(AVERAGE({ce}),"n/a")', "0.0000")
    put(MOD, r, 5, f'=IFERROR(AVERAGE({cb}),"n/a")', "0.0000"); put(MOD, r, 6, f'=IFERROR(1-D{r}/E{r},"n/a")', PCT)
    put(MOD, r, 7, f'=IF(C{r}<30,"collecting option history",IF(D{r}<E{r},"options beat the company record","company record is better"))')
    r += 1
    mrow["options_lseg"] = r
    ce, cb = MR("LSEG implied move: error"), MR("LSEG implied move: benchmark error (company's past move)")
    put(MOD, r, 1, "|Move|: implied event move from LSEG implied volatility, against the |share return|"); put(MOD, r, 2, bl["abs_ar"])
    put(MOD, r, 3, f"=COUNT({ce})", "0"); put(MOD, r, 4, f'=IFERROR(AVERAGE({ce}),"n/a")', "0.0000")
    put(MOD, r, 5, f'=IFERROR(AVERAGE({cb}),"n/a")', "0.0000"); put(MOD, r, 6, f'=IFERROR(1-D{r}/E{r},"n/a")', PCT)
    put(MOD, r, 7, f'=IF(C{r}<30,"waiting for LSEG data",IF(D{r}<E{r},"options beat the company record","company record is better"))')
    r += 1
    for key, lab, bench, (ec_, bc_), ok_txt, bad_txt in (
            ("q10_options", "10% worst case: options (LSEG implied move before the report, normal)", bl["ar"],
             ("Options 10% worst case: error", "Options 10% worst case: benchmark error (history)"), "options beat history", "history is better"),
            ("q10_ai_vs_options", "10% worst case: AI downside model against the options", "Options 10% worst case (pinball loss)",
             ("AI 10% worst case vs options: error", "AI 10% worst case vs options: benchmark error (options)"),
             "AI beats the options", "options are better"),
            ("q10_ai_vs_sheet", "10% worst case: AI downside model against the spreadsheet (k x company's past move)",
             "Spreadsheet k x past move (pinball loss)",
             ("AI 10% worst case vs spreadsheet: error", "AI 10% worst case vs spreadsheet: benchmark error (k x past move)"),
             "AI beats the spreadsheet", "spreadsheet is as good or better"),
            ("q10_text_vs_notext", "10% worst case: AI with last quarter's release read by generative AI, against the same AI without it",
             "Same AI without the reading (pinball loss)",
             ("AI with the release reading: error", "AI with the release reading: benchmark error (same AI without it)"),
             "the reading helps", "the reading does not help")):
        mrow[key] = r
        ce, cb = MR(ec_), MR(bc_)
        put(MOD, r, 1, lab); put(MOD, r, 2, bench)
        put(MOD, r, 3, f"=COUNT({ce})", "0"); put(MOD, r, 4, f'=IFERROR(AVERAGE({ce}),"n/a")', "0.0000")
        put(MOD, r, 5, f'=IFERROR(AVERAGE({cb}),"n/a")', "0.0000"); put(MOD, r, 6, f'=IFERROR(1-D{r}/E{r},"n/a")', PCT)
        put(MOD, r, 7, f'=IF(C{r}<50,"too few reports",IF(F{r}>=0.02,"{ok_txt}",IF(F{r}<=-0.02,"{bad_txt}","about the same (within 2%)")))')
        r += 1
    mrow["q10_cover"] = r
    cvr = MR("AI 10% worst case: move below the forecast (1/0)")
    put(MOD, r, 1, "AI downside model: share of reports that fell below the 10% worst case (should be close to 10%)")
    put(MOD, r, 3, f"=COUNT({cvr})", "0"); put(MOD, r, 4, f'=IFERROR(AVERAGE({cvr}),"n/a")', PC1)
    put(MOD, r, 7, f'=IF(C{r}<50,"too few reports",IF(ABS(D{r}-{A["QLOW"]})<=0.03,"well calibrated","check calibration"))')
    r += 2
    put(MOD, r, 1, "Snapshot benchmarks for P(beat): crowd p against analyst revisions (reports with a snapshot)", font=SEC); r += 1
    head(MOD, r, ["", "Reports", "Share that beat", "Average crowd p", "Brier, crowd"]); r += 1
    Rrev = R("Net revisions up / analysts, 30 days (snapshot)")
    for lab, crit in [("Analysts revised up (net revisions above 0)", '">0"'), ("No net revision", "0"),
                      ("Analysts revised down (net revisions below 0)", '"<0"')]:
        put(MOD, r, 1, lab)
        put(MOD, r, 2, f"=COUNTIFS({Rrev},{crit},{Rp},\">=0\")", "0")
        put(MOD, r, 3, f'=IFERROR(AVERAGEIFS({Ry},{Rrev},{crit},{Rp},">=0"),"n/a")', PC1)
        put(MOD, r, 4, f'=IFERROR(AVERAGEIFS({Rp},{Rrev},{crit},{Rp},">=0"),"n/a")', PC1)
        put(MOD, r, 5, f'=IFERROR(AVERAGEIFS({R("Brier: crowd")},{Rrev},{crit}),"n/a")', "0.000")
        r += 1
    put(MOD, r + 1, 1, "Reading: if the crowd's p already follows the revisions, the share that beat should match the average p in each row.", font=NOTE)
    MOD.column_dimensions["A"].width = 62; MOD.column_dimensions["B"].width = 40
    for c_ in "CDEFG":
        MOD.column_dimensions[c_].width = 15

    # ---------- Backtest: does the AI decide better than the spreadsheet rule? ----------
    title(BTS, "Backtest: AI policy against the spreadsheet rule, out of sample",
          f"Every settled report with an AI forecast and an LSEG put price. Hold = keep the position. Rule = trim {TRIM:.0%} when the crowd "
          f"p is at least {HIGH_P} (a spreadsheet can do this). AI = trim {TRIM:.0%} when the AI downside model's 1-in-10 worst case is at "
          "or below a cut-off set from earlier months only, so it trims the same share of reports as the rule did. Put = buy the "
          "LSEG-priced put instead of trimming. Outcome = abnormal return of the position over the report. All cells are formulas on Events.",
          "A2:J2")
    BTS.row_dimensions[2].height = 60
    head(BTS, 4, ["Policy", "Reports", "Share of reports acted on", "Average outcome", "Average of the worst 10%", "Worst outcome",
                  "Return given up vs holding", "Worst-10% improvement vs holding", "Improvement per point of return given up",
                  f"Large falls (abnormal return at or below the Assumptions threshold) where the policy acted",
                  "Average share of the position sold (turnover)"], h=60)
    Rhold = R("Outcome: hold")
    pols = [("Hold everything", "Outcome: hold", None), ("Spreadsheet rule: trim when the crowd is confident", "Outcome: rule trims", "Rule trims (1/0)"),
            ("AI: trim when the predicted downside is worst", "Outcome: AI trims", "AI trims (1/0)"),
            ("Spreadsheet rule: buy the put when the crowd is confident", "Outcome: rule buys the put", "Rule trims (1/0)"),
            ("AI: buy the put when the predicted downside is worst", "Outcome: AI buys the put", "AI trims (1/0)"),
            ("Smarter spreadsheet: trim the biggest past moves", "Outcome: past-move rule", "Past-move rule trims (1/0)"),
            ("Smarter spreadsheet: trim the biggest options-implied moves (LSEG)", "Outcome: implied-move rule", "Implied-move rule trims (1/0)"),
            ("Smarter spreadsheet: trim the biggest crowd P(miss) x past move", "Outcome: expected-loss rule", "Expected-loss rule trims (1/0)"),
            ("Pre-registered 1: trim the highest predicted chance of a 10% fall", "Outcome: fall-classifier policy", "Fall-classifier trims (1/0)"),
            ("Pre-registered 2: AI learns k on top of the past-move spreadsheet", "Outcome: residual-AI policy", "Residual-AI trims (1/0)"),
            ("Pre-registered 3: sized AI trims, same budget, more on the worst", "Outcome: sized AI trims", "Sized trim (share of the position)"),
            ("Part 3: past-move screen, raised when the crowd warns", "Outcome: crowd-warning rule", "Crowd-warning rule trims (1/0)")]
    bt0, r = 5, 5
    bt_row = {}
    for lab, oc, act in pols:
        Ro = R(oc)
        bt_row[oc] = r
        put(BTS, r, 1, lab); put(BTS, r, 2, f"=COUNT({Ro})", "0")
        put(BTS, r, 3, f'=IFERROR(AVERAGE({R(act)}),"n/a")' if act else 0, PC1)
        put(BTS, r, 4, f'=IFERROR(AVERAGE({Ro}),"n/a")', "+0.00%;-0.00%")
        put(BTS, r, 5, f'=IFERROR(AVERAGEIF({Ro},"<="&PERCENTILE({Ro},0.1)),"n/a")', "+0.00%;-0.00%")
        put(BTS, r, 6, f'=IFERROR(MIN({Ro}),"n/a")', "+0.0%;-0.0%")
        put(BTS, r, 7, f'=IFERROR($D${bt0}-D{r},"n/a")', "+0.00%;-0.00%")
        put(BTS, r, 8, f'=IFERROR(E{r}-$E${bt0},"n/a")', "+0.00%;-0.00%")
        put(BTS, r, 9, f'=IFERROR(IF(G{r}>0,H{r}/G{r},IF(H{r}>0,"better on both","n/a")),"n/a")', "0.00")
        put(BTS, r, 10, f'=IFERROR(COUNTIFS({Rhold},"<="&{A["BIG"]},{R(act)},">0")/COUNTIFS({Rhold},"<="&{A["BIG"]}),"n/a")' if act else 0, PC1)
        sold = (0 if not act or "put" in oc else f'=IFERROR(AVERAGE({R(act)}),"n/a")' if act.startswith("Sized") else
                f'=IFERROR({A["TRIM"]}*AVERAGE({R(act)}),"n/a")')
        put(BTS, r, 11, sold, PC1)
        r += 1
    r += 1
    put(BTS, r, 1, "Where a spreadsheet cannot see, and the holdout (reports after the Part 1 run). Same policies, same cut-offs", font=SEC); r += 1
    head(BTS, r, ["Policy", "No 3 past reports: reports", "No 3 past reports: average of the worst 10%",
                  "New reports (holdout): reports", "New reports: average of the worst 10%"]); r += 1
    sub_row = {}
    for lab, oc, _ in pols:
        if oc not in BT_SUB:
            continue
        sub_row[oc] = r
        put(BTS, r, 1, lab)
        for j, sl in enumerate(BT_SETS):
            Rx = R(f"{oc} | {sl}")
            put(BTS, r, 2 + 2 * j, f"=COUNT({Rx})", "0")
            put(BTS, r, 3 + 2 * j, f'=IFERROR(AVERAGEIF({Rx},"<="&PERCENTILE({Rx},0.1)),"n/a")', "+0.00%;-0.00%")
        r += 1
    put(BTS, r, 1, "Reports without 3 past reports: the past-move rule cannot rank them and does not trim them. The residual AI uses "
                   "the main AI flag there. The holdout is small: read it as a consistency check.", font=NOTE); r += 1
    r += 1
    put(BTS, r, 1, "AI against each rule (same turnover). The smarter spreadsheets answer: does the AI only pick volatile stocks?", font=SEC); r += 1
    head(BTS, r, ["Difference", "Estimate", "95% low (bootstrap)", "95% high (bootstrap)", "p-value (bootstrap)", "Holds at 95%"]); r += 1
    BE_ = 4 + max(len(boot), 1)
    bt_est = {}
    ra, ru, rh = bt_row["Outcome: AI trims"], bt_row["Outcome: rule trims"], bt_row["Outcome: hold"]
    rp, ri, rl = bt_row["Outcome: past-move rule"], bt_row["Outcome: implied-move rule"], bt_row["Outcome: expected-loss rule"]
    rf, rr_, rs_ = bt_row["Outcome: fall-classifier policy"], bt_row["Outcome: residual-AI policy"], bt_row["Outcome: sized AI trims"]
    for k, lab, est_f in (("bt_tail_ai_rule", "Average of the worst 10%: AI minus crowd rule (above 0 = AI loses less in bad reports)", f"=E{ra}-E{ru}"),
                          ("bt_mean_ai_rule", "Average outcome: AI minus crowd rule (above 0 = AI gives up less return)", f"=D{ra}-D{ru}"),
                          ("bt_tail_ai_hold", "Average of the worst 10%: AI minus holding everything", f"=E{ra}-E{rh}"),
                          ("bt_tail_ai_past", "Average of the worst 10%: AI minus trimming the biggest past moves", f"=E{ra}-E{rp}"),
                          ("bt_tail_ai_iv", "Average of the worst 10%: AI minus trimming the biggest implied moves", f"=E{ra}-E{ri}"),
                          ("bt_tail_ai_el", "Average of the worst 10%: AI minus trimming the biggest crowd P(miss) x past move", f"=E{ra}-E{rl}"),
                          ("bt_tail_fall10_past", "Pre-registered 1 (fall classifier) minus trimming the biggest past moves", f"=E{rf}-E{rp}"),
                          ("bt_tail_resid_past", "Pre-registered 2 (AI learns k) minus trimming the biggest past moves", f"=E{rr_}-E{rp}"),
                          ("bt_tail_size_past", "Pre-registered 3 (sized trims) minus trimming the biggest past moves", f"=E{rs_}-E{rp}"),
                          ("bt_tail_ai_past_nohist", "No 3 past reports: AI minus the past-move rule",
                           f'=C{sub_row["Outcome: AI trims"]}-C{sub_row["Outcome: past-move rule"]}'),
                          ("bt_tail_cw_past", "Part 3: crowd-warning screen minus trimming the biggest past moves",
                           f'=E{bt_row["Outcome: crowd-warning rule"]}-E{rp}')):
        bt_est[k] = f"Backtest!B{r}"
        cn = f"All: {k}"
        put(BTS, r, 1, lab); put(BTS, r, 2, est_f, "+0.00%;-0.00%")
        if cn in boot.columns:
            cl = L(1 + list(boot.columns).index(cn)); rg = f"Raw_Bootstrap!${cl}$5:${cl}${BE_}"
            put(BTS, r, 3, f'=IFERROR(PERCENTILE({rg},0.025),"n/a")', "+0.00%;-0.00%")
            put(BTS, r, 4, f'=IFERROR(PERCENTILE({rg},0.975),"n/a")', "+0.00%;-0.00%")
            put(BTS, r, 5, f'=IFERROR(2*MIN(COUNTIF({rg},">=0"),COUNTIF({rg},"<=0"))/COUNT({rg}),"n/a")', "0.000")
            put(BTS, r, 6, f'=IF(ISNUMBER(E{r}),IF(E{r}<0.05,"yes","no"),"n/a")')
        r += 1
    r += 1
    for t_ in ("Reading: a policy is better if it improves the worst 10% for less return given up. The AI and the rule trim the same share "
               "of reports, so the difference is which reports they choose.",
               "Trading cost and trim size are blue inputs on Assumptions. The AI cut-off and forecasts are on Raw_Models (walk-forward, Python).",
               "Bootstrap: whole report weeks resampled (Raw_Bootstrap). Put prices: Black-Scholes at the LSEG at-the-money implied volatility."):
        put(BTS, r, 1, t_, font=NOTE); r += 1
    BTS.column_dimensions["A"].width = 62
    for c_ in "BCDEFGHIJK":
        BTS.column_dimensions[c_].width = 15

    # ---------- Crowd_Adds: what the crowd adds to the decision (Part 3) ----------
    CA = wb.create_sheet("Crowd_Adds", index=wb.sheetnames.index("Backtest") + 1)
    title(CA, "Crowd_Adds: does the crowd add anything to the past-move screen? (Part 3)",
          "Two signals the past-move screen cannot see. 1) The crowd loses confidence: p fell by at least the Assumptions threshold over "
          "the week before the report. 2) The crowd disagrees with the professionals: the crowd is confident of a beat but StarMine "
          "expects a miss, or analysts cut their estimates. Every cell is a formula on Events. Results are associations.", "A2:H2")
    CA.row_dimensions[2].height = 60
    head(CA, 4, ["Group", "Reports", "Average move (abnormal)", "Average |move|", "Share falling by the fall threshold or more",
                 "Average of the worst 10%"])
    grp = [("Crowd lost confidence over the week", "AR | crowd lost confidence"),
           ("Crowd steady or more sure", "AR | crowd steady or more sure"),
           ("Confident crowd, StarMine expects a miss", "AR | confident, StarMine expects a miss"),
           ("Confident crowd, StarMine agrees", "AR | confident, StarMine agrees"),
           ("Confident crowd, analysts cut estimates (30 days)", "AR | confident, analysts cut"),
           ("Confident crowd, analysts up or flat", "AR | confident, analysts up or flat")]
    ca_row = {}
    r = 5
    for lab, col in grp:
        Rx = R(col); ca_row[col] = r
        put(CA, r, 1, lab); put(CA, r, 2, f"=COUNT({Rx})", "0")
        put(CA, r, 3, f'=IFERROR(AVERAGE({Rx}),"n/a")', "+0.00%;-0.00%")
        put(CA, r, 4, f'=IFERROR((SUMIF({Rx},">0")-SUMIF({Rx},"<0"))/COUNT({Rx}),"n/a")', "0.00%")
        put(CA, r, 5, f'=IFERROR(COUNTIF({Rx},"<="&{A["FALL"]})/COUNT({Rx}),"n/a")', "0.0%")
        put(CA, r, 6, f'=IFERROR(AVERAGEIF({Rx},"<="&PERCENTILE({Rx},0.1)),"n/a")', "+0.00%;-0.00%")
        r += 1
    r += 1
    put(CA, r, 1, "Tests: is the difference larger than chance? Whole report weeks resampled (Raw_Bootstrap)", font=SEC); r += 1
    head(CA, r, ["Difference", "Estimate", "95% low", "95% high", "p-value (bootstrap)", "Holds at 95%"]); r += 1
    g_ = lambda col, c: f"{c}{ca_row[col]}"
    for k, lab, est_f in (
            ("bt_ca_drift_tail", "Worst 10%: crowd lost confidence minus the other reports",
             f'={g_("AR | crowd lost confidence", "F")}-{g_("AR | crowd steady or more sure", "F")}'),
            ("bt_ca_drift_abs", "Average |move|: crowd lost confidence minus the other reports",
             f'={g_("AR | crowd lost confidence", "D")}-{g_("AR | crowd steady or more sure", "D")}'),
            ("bt_ca_sm_tail", "Worst 10%, confident crowd: StarMine expects a miss minus StarMine agrees",
             f'={g_("AR | confident, StarMine expects a miss", "F")}-{g_("AR | confident, StarMine agrees", "F")}'),
            ("bt_ca_rev_tail", "Worst 10%, confident crowd: analysts cut minus analysts up or flat",
             f'={g_("AR | confident, analysts cut", "F")}-{g_("AR | confident, analysts up or flat", "F")}')):
        bt_est[k] = f"Crowd_Adds!B{r}"
        put(CA, r, 1, lab); put(CA, r, 2, est_f, "+0.00%;-0.00%")
        cn = f"All: {k}"
        if cn in boot.columns:
            cl = L(1 + list(boot.columns).index(cn)); rg = f"Raw_Bootstrap!${cl}$5:${cl}${BE_}"
            put(CA, r, 3, f'=IFERROR(PERCENTILE({rg},0.025),"n/a")', "+0.00%;-0.00%")
            put(CA, r, 4, f'=IFERROR(PERCENTILE({rg},0.975),"n/a")', "+0.00%;-0.00%")
            put(CA, r, 5, f'=IFERROR(2*MIN(COUNTIF({rg},">=0"),COUNTIF({rg},"<=0"))/COUNT({rg}),"n/a")', "0.000")
            put(CA, r, 6, f'=IF(ISNUMBER(E{r}),IF(E{r}<0.05,"yes","no"),"n/a")')
        r += 1
    r += 1
    put(CA, r, 1, "Does it change the decision? Same turnover as the crowd rule, out of sample (Backtest)", font=SEC); r += 1
    head(CA, r, ["Policy", "Reports", "Average outcome", "Average of the worst 10%"]); r += 1
    for lab, oc in (("Hold everything", "Outcome: hold"), ("Past-move screen", "Outcome: past-move rule"),
                    ("Past-move screen, raised when the crowd warns", "Outcome: crowd-warning rule"), ("AI downside model", "Outcome: AI trims")):
        br_ = bt_row[oc]
        put(CA, r, 1, lab); put(CA, r, 2, f"=Backtest!B{br_}", "0")
        put(CA, r, 3, f"=Backtest!D{br_}", "+0.00%;-0.00%"); put(CA, r, 4, f"=Backtest!E{br_}", "+0.00%;-0.00%")
        r += 1
    put(CA, r, 1, "Crowd-warning screen minus past-move screen, worst 10% (bootstrap test on Backtest)"); put(CA, r, 2, f"={bt_est['bt_tail_cw_past']}", "+0.00%;-0.00%")
    r += 2
    for t_ in (f"The crowd warns when it lost confidence over the week, or when it is confident but StarMine expects a miss. The screen ranks "
               f"reports by past |move| x (1 + {CROWD_WARN_WEIGHT}) when the crowd warns. The weight was fixed before the test, not tuned.",
               "A negative worst-10% difference means the group falls harder in bad reports. Above 0 on the last line means the crowd "
               "makes the screen lose less."):
        put(CA, r, 1, t_, font=NOTE); r += 1
    CA.column_dimensions["A"].width = 62
    for c_ in "BCDEF":
        CA.column_dimensions[c_].width = 16

    # ---------- Business: market, price and value for a client (Part 3) ----------
    BUS = wb.create_sheet("Business", index=wb.sheetnames.index("Crowd_Adds") + 1)
    title(BUS, "Business: market, pricing and value for a client (Part 3)",
          "Blue cells are inputs. Market figures have a source. Pricing and costs are the author's assumptions and can be changed. "
          "Client value links to the Backtest sheet.", "A2:D2")
    BB = {}
    r = 4
    def sec_(t_):
        nonlocal r
        put(BUS, r, 1, t_, font=SEC); r += 1
        head(BUS, r, ["Item", "Value", "Source or note"], h=20); r += 1
    def inp(key, lab, v, nf, src_):
        nonlocal r
        put(BUS, r, 1, lab); put(BUS, r, 2, v, nf, BLUE); put(BUS, r, 3, src_)
        BB[key] = f"$B${r}"; r += 1
    def calc(key, lab, fx, nf, note=""):
        nonlocal r
        put(BUS, r, 1, lab); put(BUS, r, 2, fx, nf); put(BUS, r, 3, note)
        BB[key] = f"$B${r}"; r += 1
    sec_("1. Market (sourced)")
    inp("ria", "SEC-registered investment advisers, end of 2025", 16544, "#,##0",
        "Investment Adviser Association and Comply, Investment Adviser Industry Snapshot 2026 (June 2026)")
    inp("ria_ind", "Share of SEC-registered advisers managing assets for individuals", 0.634, "0.0%",
        "Investment Adviser Association and Comply, Investment Adviser Industry Snapshot 2025 (2024 data)")
    inp("mcb", "Banks in Monaco, Q1 2025", 26, "0", "Monaco Statistics (IMSEE), Economy Bulletin Q1 2025")
    inp("mca", "Asset management companies in Monaco, Q1 2025", 70, "0", "Monaco Statistics (IMSEE), Economy Bulletin Q1 2025")
    inp("bbg", "Bloomberg Terminal, one seat, per year (USD)", 31980, "#,##0", "NeuGroup, 9 Oct 2024: 2025 price for single-terminal users")
    inp("lseg", "LSEG Workspace, per user per year, low end (USD)", 12000, "#,##0", "multiples.vc, Financial Data Platform Pricing in 2026")
    inp("pmv", "Kalshi and Polymarket trading volume in 2025 (USD billion)", 44, "0", "The Block, 31 Dec 2025")
    calc("tgt", "Target firms: US advisers for individuals plus Monaco banks and asset managers",
         f"=ROUND({BB['ria']}*{BB['ria_ind']},0)+{BB['mcb']}+{BB['mca']}", "#,##0", "Beachhead: Monaco. Scale market: US advisers")
    r += 1
    sec_("2. Pricing and costs (group assumptions)")
    inp("price", "BeatOdds price per seat per year (EUR)", 2400, "#,##0", "Assumption. A fraction of a terminal seat")
    inp("seats", "Seats per client firm", 3, "0", "Assumption")
    inp("pen1", "Share of target firms signed, year 1", 0.005, "0.0%", "Assumption")
    inp("pen2", "Share of target firms signed, year 2", 0.015, "0.0%", "Assumption")
    inp("pen3", "Share of target firms signed, year 3", 0.03, "0.0%", "Assumption")
    inp("var", "Variable cost per seat per year: data, Gemini calls, hosting (EUR)", 400, "#,##0", "Assumption")
    inp("fix", "Fixed costs per year: two staff, data licences, compliance (EUR)", 250000, "#,##0", "Assumption")
    r += 1
    put(BUS, r, 1, "3. Three-year plan (formulas)", font=SEC); r += 1
    head(BUS, r, ["Item", "Year 1", "Year 2", "Year 3"], h=20); r += 1
    plan = {}
    for lab, key, fx, nf in (("Client firms", "firms", lambda y: f"=ROUND({BB['tgt']}*{BB['pen' + str(y)]},0)", "#,##0"),
                             ("Seats", "st", lambda y: f"={plan['firms'][y]}*{BB['seats']}", "#,##0"),
                             ("Revenue (EUR)", "rev", lambda y: f"={plan['st'][y]}*{BB['price']}", "#,##0"),
                             ("Costs (EUR)", "cost", lambda y: f"={plan['st'][y]}*{BB['var']}+{BB['fix']}", "#,##0"),
                             ("Operating profit (EUR)", "prof", lambda y: f"={plan['rev'][y]}-{plan['cost'][y]}", "#,##0;-#,##0")):
        plan[key] = {}
        put(BUS, r, 1, lab)
        for y in (1, 2, 3):
            plan[key][y] = f"{L(1 + y)}{r}"
            put(BUS, r, 1 + y, fx(y), nf)
        r += 1
    calc("be", "Break-even seats", f"=ROUNDUP({BB['fix']}/({BB['price']}-{BB['var']}),0)", "#,##0", "Fixed costs / (price - variable cost)")
    r += 1
    sec_("4. Value for one client (links to Backtest)")
    inp("book", "Equity book per manager (EUR)", 10000000, "#,##0", "Assumption, same as the slides")
    inp("hold", "Holdings in the book", 20, "0", "Assumption, same as the slides")
    inp("seas", "Earnings seasons per year", 4, "0", "Quarterly reporting")
    ra_, rp_ = bt_row["Outcome: AI trims"], bt_row["Outcome: past-move rule"]
    calc("imp", "Worst-10% improvement vs holding: AI policy (points)", f"=Backtest!H{ra_}", "+0.0%;-0.0%", "Backtest, out of sample")
    calc("impp", "Worst-10% improvement vs holding: past-move screen (points)", f"=Backtest!H{rp_}", "+0.0%;-0.0%", "Backtest, out of sample")
    calc("bad", "Bad reports per season (worst 10% of holdings)", f"={BB['hold']}*0.1", "0.0")
    calc("val", "Loss avoided in bad reports per season (EUR)", f"={BB['imp']}*{BB['book']}/{BB['hold']}*{BB['bad']}", "#,##0")
    calc("valy", "Loss avoided per year (EUR)", f"={BB['val']}*{BB['seas']}", "#,##0")
    calc("ratio", "One seat as a share of the loss avoided per year", f"=IFERROR({BB['price']}/{BB['valy']},\"n/a\")", "0.0%")
    r += 1
    for t_ in ("Loss avoided is measured in the backtest on past reports. It is not a forecast. A past-move screen gets most of it, "
               "so the price is for the full service: odds read across venues, worst case, put pricing and a checked client brief.",
               "Bloomberg and LSEG prices are list prices reported in the press, in USD. BeatOdds is priced in EUR."):
        put(BUS, r, 1, t_, font=NOTE); r += 1
    BUS.column_dimensions["A"].width = 66; BUS.column_dimensions["B"].width = 16
    BUS.column_dimensions["C"].width = 16; BUS.column_dimensions["D"].width = 16

    # ---------- Bootstrap ----------
    RB["A1"] = (f"Bootstrap draws written by the script: {BOOT} draws, whole report weeks resampled with replacement, seed {SEED}. "
                f"Thresholds used: confident p >= {HIGH_P}, fall <= {FALL:.0%}. Empty column: too few reports or misses.")
    RB["A1"].font = BOLD
    head(RB, 4, list(boot.columns))
    for i, rec in enumerate(boot.itertuples(index=False)):
        for j, v in enumerate(rec):
            put(RB, 5 + i, 1 + j, clean(v), "0.0000")
    BE = 4 + max(len(boot), 1)
    title(BT, "Bootstrap: how sure are we? (formulas on Raw_Bootstrap)",
          "Estimate = the same statistic on Tail or Results. Interval = 2.5th to 97.5th percentile of the draws. "
          "p-value = twice the share of draws on the other side of zero.", "A2:H2")
    head(BT, 4, ["Sample", "Statistic", "Estimate", "95% low", "95% high", "p-value (bootstrap)", "Draws"])
    tail_row = {}
    blk = 2 + 14 + 1
    for b_i, s in enumerate(("All", "Part 1", "New")):
        base = 6 + b_i * blk
        tail_row[s] = dict(mean_gap=f"Tail!D{base + 3}", fall_gap=f"Tail!D{base + 7}", tail_gap=f"Tail!D{base + 13}")
    est_cell = {"All": dict(miss_gap=f"Results!B{w0 + 4}", **tail_row["All"]), "New": dict(miss_gap=f"Results!D{w0 + 4}", **tail_row["New"])}
    r = 5
    est_cell["All"].update(bt_est)
    for s in ("All", "New"):
        for k, lab in BOOT_STATS:
            if s == "New" and k.startswith("bt_"):
                continue
            colname = f"{s}: {k}"
            cl = L(1 + list(boot.columns).index(colname))
            rg = f"Raw_Bootstrap!${cl}$5:${cl}${BE}"
            put(BT, r, 1, "All reports" if s == "All" else "New reports only"); put(BT, r, 2, lab)
            put(BT, r, 3, f"={est_cell[s][k]}", PCT)
            put(BT, r, 4, f'=IFERROR(PERCENTILE({rg},0.025),"n/a")', PCT); put(BT, r, 5, f'=IFERROR(PERCENTILE({rg},0.975),"n/a")', PCT)
            put(BT, r, 6, f'=IFERROR(2*MIN(COUNTIF({rg},">=0"),COUNTIF({rg},"<=0"))/COUNT({rg}),"n/a")', "0.000")
            put(BT, r, 7, f"=COUNT({rg})", "0")
            r += 1
    put(BT, r + 1, 1, "If the blue thresholds on Assumptions are changed, rerun the script so the draws use the same thresholds.", font=NOTE)
    BT.column_dimensions["A"].width = 18; BT.column_dimensions["B"].width = 55
    for c_ in "CDEFG":
        BT.column_dimensions[c_].width = 14

    # ---------- Kalshi ----------
    nk = 0
    if len(km):
        km = km.reset_index(drop=True).copy()
        km["sym"] = km["ticker"].map(yf_symbol)
        kcols = ["mticker", "series", "event", "title", "strike", "actual", "y", "volume", "close_ts", "ticker", "report_date", "cut",
                 "date_matched", "sym"]
        RKM["A1"] = f"Kalshi KPI contracts as downloaded (merged across runs): {KALSHI}"; RKM["A1"].font = BOLD
        head(RKM, 4, ["Contract", "Series", "Event", "Title", "Strike", "Actual value", "Outcome (1 = above)", "Volume (contracts)",
                      "Close (unix)", "Stock ticker", "Report date", "Cut-off (unix)", "Report date matched on Yahoo", "Yahoo symbol"])
        for i, rec in enumerate(km[kcols].to_dict("records")):
            for j, k in enumerate(kcols):
                v = clean(rec[k])
                if k == "report_date" and v:
                    v = pd.Timestamp(v).to_pydatetime()
                put(RKM, 5 + i, 1 + j, v, "yyyy-mm-dd" if k == "report_date" else None)
        RKO["A1"] = "Kalshi hourly YES prices as downloaded (candlesticks)"; RKO["A1"].font = BOLD
        head(RKO, 4, ["Contract", "Time (unix, UTC)", "YES price"])
        ko = ko[ko["market"].isin(set(km["mticker"]))].sort_values(["market", "ts"])
        kspan, rr = {}, 5
        for mt, g in ko.groupby("market", sort=False):
            a = rr
            for t, p in zip(g["ts"], g["p"]):
                RKO.cell(row=rr, column=1, value=mt); RKO.cell(row=rr, column=2, value=int(t)); RKO.cell(row=rr, column=3, value=float(p)); rr += 1
            kspan[mt] = (a, rr - 1)
        title(KE, "Kalshi_Events: one row per KPI event with odds (formulas on Raw_K_Markets, Raw_K_Odds, Raw_Prices)",
              "p = YES price at the cut-off (column L on Raw_K_Markets). Base rate = share of earlier KPI events above the strike.", "A2:P2")
        kh = ["Raw_K_Markets row", "Ticker", "Report date", "Strike", "Actual", "Outcome y (1 = above)", "Volume", "Odds first row",
              "Odds last row", "Cut-off (unix)", "p at cut-off", "Running base rate", "Brier: Kalshi", "Brier: base rate", "Both (1/0)",
              "Correct side (1/0)", "Price column", "Abnormal return", "KPI surprise y - p"]
        head(KE, 4, kh)
        kr_ = 5
        rows_k = [(i, rec) for i, rec in enumerate(km.to_dict("records")) if rec["mticker"] in kspan]
        KE1 = 4 + max(len(rows_k), 1)
        for i, rec in rows_k:
            raw, (a, b), r = 5 + i, kspan[rec["mticker"]], kr_
            KE[f"A{r}"] = raw; KE[f"B{r}"] = f"=Raw_K_Markets!J{raw}"; KE[f"C{r}"] = f"=Raw_K_Markets!K{raw}"; KE[f"C{r}"].number_format = "yyyy-mm-dd"
            KE[f"D{r}"] = f"=Raw_K_Markets!E{raw}"; KE[f"E{r}"] = f"=Raw_K_Markets!F{raw}"; KE[f"F{r}"] = f"=Raw_K_Markets!G{raw}"
            KE[f"G{r}"] = f"=Raw_K_Markets!H{raw}"; KE[f"H{r}"] = a; KE[f"I{r}"] = b; KE[f"J{r}"] = f"=Raw_K_Markets!L{raw}"
            KE[f"K{r}"] = f'=IFERROR(INDEX(Raw_K_Odds!$C${a}:$C${b},MATCH(J{r},Raw_K_Odds!$B${a}:$B${b},1)),"")'
            KE[f"L{r}"] = f'=IF(COUNTIF($C$5:$C${KE1},"<"&C{r})>={A["KMINB"]},AVERAGEIFS($F$5:$F${KE1},$C$5:$C${KE1},"<"&C{r}),"")'
            KE[f"M{r}"] = f'=IF(ISNUMBER(K{r}),(K{r}-F{r})^2,"")'; KE[f"N{r}"] = f'=IF(ISNUMBER(L{r}),(L{r}-F{r})^2,"")'
            KE[f"O{r}"] = f'=IF(AND(ISNUMBER(M{r}),ISNUMBER(N{r})),1,0)'
            KE[f"P{r}"] = f'=IF(ISNUMBER(K{r}),IF((K{r}>0.5)=(F{r}=1),1,0),"")'
            KE[f"Q{r}"] = f'=IFERROR(MATCH(Raw_K_Markets!N{raw},{HDR},0),"")'
            a_, b_ = f"MATCH(C{r}-1,{DATES},1)", f"MATCH(C{r},{DATES},1)+1"
            KE[f"R{r}"] = f'=IFERROR(LN(INDEX({PX},{b_},Q{r})/INDEX({PX},{a_},Q{r}))-LN(INDEX({PX},{b_},{MC_})/INDEX({PX},{a_},{MC_})),"")'
            KE[f"S{r}"] = f'=IF(ISNUMBER(K{r}),F{r}-K{r},"")'
            for c_ in "KLMN":
                KE[f"{c_}{r}"].number_format = "0.000"
            KE[f"R{r}"].number_format = PCT
            kr_ += 1
        nk = len(rows_k)
        KR_ = lambda c_: f"Kalshi_Events!${c_}$5:${c_}${KE1}"
        title(KR, "Kalshi KPI contracts: does the crowd forecast operating numbers?",
              "Each contract asks whether a company reports a KPI above a level. All cells are formulas on Kalshi_Events.", "A2:F2")
        r = 4
        head(KR, r, ["", "Value"]); r += 1
        k0 = r
        for lab, v, nf in [("KPI events with odds", f"=COUNT({KR_('K')})", "0"), ("Share above the level", f"=AVERAGE({KR_('F')})", PC1),
                           ("Average Kalshi p", f"=AVERAGE({KR_('K')})", PC1), ("Correct side", f"=AVERAGE({KR_('P')})", PC1),
                           ("Always guessing the more common outcome", f"=MAX(B{k0+1},1-B{k0+1})", PC1),
                           ("Brier, Kalshi (events with both)", f'=IFERROR(AVERAGEIFS({KR_("M")},{KR_("O")},1),"n/a")', "0.000"),
                           ("Brier, running base rate", f'=IFERROR(AVERAGEIFS({KR_("N")},{KR_("O")},1),"n/a")', "0.000"),
                           ("Skill vs base rate", f'=IFERROR(1-B{k0+5}/B{k0+6},"n/a")', PCT),
                           ("Abnormal return on KPI surprise: slope", f'=IFERROR(SLOPE({KR_("R")},{KR_("S")}),"n/a")', "0.0000"),
                           ("t statistic", f'=IFERROR(SLOPE({KR_("R")},{KR_("S")})/(STEYX({KR_("R")},{KR_("S")})/SQRT(DEVSQ({KR_("S")}))),"n/a")', "0.00"),
                           ("R squared", f'=IFERROR(RSQ({KR_("R")},{KR_("S")}),"n/a")', "0.000")]:
            put(KR, r, 1, lab); put(KR, r, 2, v, nf); r += 1
        r += 1
        put(KR, r, 1, "Calibration", font=SEC); r += 1
        head(KR, r, ["p from", "p to", "Events", "Average p", "Share above"]); r += 1
        for lo, hi in [(0, .2), (.2, .4), (.4, .6), (.6, .8), (.8, 1.0001)]:
            put(KR, r, 1, lo, font=BLUE); put(KR, r, 2, hi, font=BLUE)
            put(KR, r, 3, f'=COUNTIFS({KR_("K")},">="&A{r},{KR_("K")},"<"&B{r})')
            put(KR, r, 4, f'=IFERROR(AVERAGEIFS({KR_("K")},{KR_("K")},">="&A{r},{KR_("K")},"<"&B{r}),"")', PC1)
            put(KR, r, 5, f'=IFERROR(AVERAGEIFS({KR_("F")},{KR_("K")},">="&A{r},{KR_("K")},"<"&B{r}),"")', PC1)
            r += 1
        KR.column_dimensions["A"].width = 46
    else:
        put(KR, 1, 1, "No Kalshi data in this run.", font=NOTE)

    # ---------- Summary: what holds, at what confidence ----------
    title(SUM_, "Summary: BeatOdds Part 3. What holds, and what it means for a manager",
          "Every number links to the sheet named in the last column. 'Holds' uses week-clustered tests (Tests sheet) at the confidence "
          "levels on Assumptions. Robust findings first, then the Part 1 illustration with its continuous test, then the decision.", "A2:J2")
    head(SUM_, 4, ["Finding", "Sample", "Estimate", "t", "p-value", "Holds at level 1", "Holds at level 2", "Holds at level 3", "Source"])
    for j, k in enumerate(("CL1", "CL2", "CL3")):
        put(SUM_, 3, 6 + j, f"={A[k]}", "0%", BOLD)
    r = 5

    def srow(text, key, src="Tests"):
        nonlocal r
        q = trow[key]
        put(SUM_, r, 1, text); put(SUM_, r, 2, f"=Tests!A{q}")
        put(SUM_, r, 3, f"=Tests!E{q}", TST.cell(row=q, column=5).number_format)
        for j, c_ in enumerate("KMNOP"):
            put(SUM_, r, 4 + j, f"=Tests!{c_}{q}", "0.00" if c_ == "K" else ("0.0000" if c_ == "M" else None))
        put(SUM_, r, 9, f"Tests row {q}")
        r += 1

    def sec(t):
        nonlocal r
        r += 1; put(SUM_, r, 1, t, font=SEC); r += 1

    sec("1. Robust: the crowd forecasts the result, and the share reacts to the surprise against the crowd")
    for smp in ("All", "New"):
        srow("Crowd beats the running base rate (Brier gain)", f"{smp}_skill_base")
        srow("Crowd beats the company's own record (Brier gain)", f"{smp}_skill_rec")
    srow("Share reacts to the surprise against the crowd (slope on y - p)", "All_surprise")
    srow("Share reacts to the surprise against the crowd (slope on y - p)", "New_surprise")
    srow("Expected beats are priced in (slope of the reaction on p, beats)", "All_beats")
    sec("2. The Part 1 illustration and its continuous test")
    for i in range(len(GAP_THRESHOLDS)):
        srow("Miss gap: confident misses minus other misses, at the threshold on Tests", f"All_gap_TH{i + 1}")
    srow("Miss gap at the confident threshold, new reports", "New_gap_HP")
    srow("Continuous test: does a higher p go with a bigger fall on a miss?", "All_misses")
    srow("Miss gap with odds read before the release (EDGAR check)", "All_gap_before")
    sec("3. A confident crowd does not mean a safer report")
    srow("Falls of 5% or more: confident minus less sure", "All_fall")
    srow("Hedge P&L with past option prices, confident reports", "All_hedge_conf")
    srow("Hedge P&L with past option prices: confident minus less sure", "All_hedge_gap")
    srow("Options overprice the move more when the crowd is confident? (implied - realised)", "All_impl_gap")
    srow("LSEG: options expect a smaller move when the crowd is confident?", "All_lseg_move_gap")
    srow("LSEG: options overprice the earnings move (implied - realised)", "All_lseg_over")
    srow("LSEG: overpricing, confident minus less sure", "All_lseg_over_gap")
    srow("LSEG-priced put: hedge P&L, confident reports", "All_lput_conf")
    srow("LSEG-priced put: hedge P&L, confident minus less sure", "All_lput_gap")
    srow("LSEG analysts: beat rate, estimate raised minus not raised", "All_rev_beat")
    srow("LSEG analysts: crowd still adds when the estimate was raised (confident minus less sure)", "All_rev_conf_gap")
    put(SUM_, r, 1, "Average of the worst 10% of moves: confident minus less sure"); put(SUM_, r, 2, "All reports")
    put(SUM_, r, 3, "=Tail!D19", PCT); put(SUM_, r, 5, "=Bootstrap!F8", "0.0000"); put(SUM_, r, 9, "Tail row 19, Bootstrap row 8"); r += 1
    put(SUM_, r, 1, "Expected put payoff from history: confident / less sure"); put(SUM_, r, 2, "All reports")
    put(SUM_, r, 3, f"=Hedge!B{h0+4}", PC2); put(SUM_, r, 4, f"=Hedge!C{h0+4}", PC2); put(SUM_, r, 9, f"Hedge row {h0+4} (B and C)"); r += 1
    sec("4. The decision: the crowd says whether, a model says how much, the options say the price")
    head(SUM_, r, ["Piece", "Measure", "Value", "Benchmark", "Skill vs benchmark", "Decision", "", "", "Source"]); r += 1
    for lab, col in (("Whether: best model with the crowd vs the crowd alone (Brier)", "beat_crowd_mgr_lin"),
                     ("Whether: manager's data only vs the crowd (Brier)", "beat_mgr_lin"),
                     ("How much: boosted trees vs the company's past move (abs. error)", "move_crowd_mgr_gb"),
                     ("How much: option-implied move vs the company's past move", "options"),
                     ("How much: implied move from past option prices vs the company's past move", "options_hist"),
                     ("How much: LSEG implied event move vs the company's past move", "options_lseg"),
                     ("Downside: AI 10% worst case vs history (pinball loss)", "q10_crowd_mgr_gb"),
                     ("Downside: AI 10% worst case vs the options market", "q10_ai_vs_options"),
                     ("Downside: AI 10% worst case vs a spreadsheet (k x past move)", "q10_ai_vs_sheet"),
                     ("Downside: AI with the release read by generative AI vs without", "q10_text_vs_notext"),
                     ("Whether: crowd + release read by generative AI vs crowd (Brier)", "beat_crowd_text_lin"),
                     ("Whether: crowd + keyword count of the release vs crowd (Brier)", "beat_crowd_kw_lin"),
                     ("Whether: StarMine alone vs the crowd (Brier)", "beat_sm_lin"),
                     ("Whether: crowd + StarMine vs StarMine alone (Brier)", "beat_crowd_sm_lin"),
                     ("Downside: AI learns k on top of the spreadsheet vs the spreadsheet", "q10_resid_gb"),
                     ("Downside: AI + Kalshi KPI odds vs the same AI without them", "q10_kalshi_gb")):
        q = mrow[col]
        put(SUM_, r, 1, lab); put(SUM_, r, 2, f"=Models!C{q}", "0"); put(SUM_, r, 3, f"=Models!D{q}", "0.0000")
        put(SUM_, r, 4, f"=Models!E{q}", "0.0000"); put(SUM_, r, 5, f"=Models!F{q}", PCT); put(SUM_, r, 6, f"=Models!G{q}")
        put(SUM_, r, 9, f"Models row {q}"); r += 1
    for lab, rr_o, nf in (("The price: average put price, upcoming reports", H_OPT + 1, PC2),
                          ("The price: expected payoff from history, same reports", H_OPT + 2, PC2),
                          ("The price: price / expected payoff (below 1 = cheap)", H_OPT + 3, "0.00"),
                          ("The price: realised hedge P&L on settled snapshot reports", H_SET + 3, "+0.00%;-0.00%"),
                          ("The price: hedge P&L with past option prices (Massive)", H_HIST + 3, "+0.00%;-0.00%"),
                          ("The price: implied move - realised move, past option prices", H_HIST + 5, "+0.0%;-0.0%"),
                          ("The price: LSEG implied event move", H_LSEG + 3, PC1),
                          ("The price: LSEG implied event move - realised move", H_LSEG + 4, "+0.0%;-0.0%"),
                          ("The price: LSEG-priced put, average price", H_LSP + 1, PC2),
                          ("The price: LSEG-priced put, average hedge P&L", H_LSP + 3, "+0.00%;-0.00%"),
                          ("Analysts: beat rate when the estimate was raised", H_REV + 2, PC1),
                          ("Analysts: beat rate when it was not raised", H_REV + 3, PC1)):
        put(SUM_, r, 1, lab); put(SUM_, r, 2, "confident / less sure")
        put(SUM_, r, 3, f"=Hedge!B{rr_o}", nf); put(SUM_, r, 4, f"=Hedge!C{rr_o}", nf); put(SUM_, r, 9, f"Hedge row {rr_o}"); r += 1
    sec("5. Why AI: does the AI policy decide better than the spreadsheet rule? (Backtest sheet)")
    head(SUM_, r, ["Policy", "Reports", "Average outcome", "Average of the worst 10%", "Worst-10% improvement vs holding",
                   "Return given up vs holding", "", "", "Source"]); r += 1
    for lab, oc in (("Hold everything", "Outcome: hold"), ("Spreadsheet rule: trim when confident", "Outcome: rule trims"),
                    ("AI: trim the worst predicted downside", "Outcome: AI trims"), ("Rule: buy the put when confident", "Outcome: rule buys the put"),
                    ("AI: buy the put on the worst predicted downside", "Outcome: AI buys the put"),
                    ("Smarter spreadsheet: biggest past moves", "Outcome: past-move rule"),
                    ("Smarter spreadsheet: biggest implied moves (LSEG)", "Outcome: implied-move rule"),
                    ("Smarter spreadsheet: biggest P(miss) x past move", "Outcome: expected-loss rule"),
                    ("Pre-registered 1: fall classifier", "Outcome: fall-classifier policy"),
                    ("Pre-registered 2: AI learns k on top of the spreadsheet", "Outcome: residual-AI policy"),
                    ("Pre-registered 3: sized AI trims", "Outcome: sized AI trims")):
        q = bt_row[oc]
        put(SUM_, r, 1, lab); put(SUM_, r, 2, f"=Backtest!B{q}", "0"); put(SUM_, r, 3, f"=Backtest!D{q}", "+0.00%;-0.00%")
        put(SUM_, r, 4, f"=Backtest!E{q}", "+0.00%;-0.00%"); put(SUM_, r, 5, f"=Backtest!H{q}", "+0.00%;-0.00%")
        put(SUM_, r, 6, f"=Backtest!G{q}", "+0.00%;-0.00%"); put(SUM_, r, 9, f"Backtest row {q}"); r += 1
    head(SUM_, r, ["AI minus each rule", "", "Estimate", "95% low", "95% high", "p-value", "Holds at 95%", "", "Source"]); r += 1
    for k, lab in (("bt_tail_ai_rule", "Worst 10%: AI minus crowd rule"), ("bt_mean_ai_rule", "Average outcome: AI minus crowd rule"),
                   ("bt_tail_ai_hold", "Worst 10%: AI minus holding"), ("bt_tail_ai_past", "Worst 10%: AI minus biggest past moves"),
                   ("bt_tail_ai_iv", "Worst 10%: AI minus biggest implied moves"), ("bt_tail_ai_el", "Worst 10%: AI minus P(miss) x past move"),
                   ("bt_tail_fall10_past", "Worst 10%: pre-registered 1 minus biggest past moves"),
                   ("bt_tail_resid_past", "Worst 10%: pre-registered 2 minus biggest past moves"),
                   ("bt_tail_size_past", "Worst 10%: pre-registered 3 minus biggest past moves"),
                   ("bt_tail_ai_past_nohist", "No 3 past reports: AI minus the past-move rule")):
        q = int(bt_est[k].split("B")[-1])
        put(SUM_, r, 1, lab)
        for j, c_ in enumerate("BCDEF"):
            put(SUM_, r, 3 + j, f"=Backtest!{c_}{q}", "0.000" if c_ == "E" else (None if c_ == "F" else "+0.00%;-0.00%"))
        put(SUM_, r, 9, f"Backtest row {q}"); r += 1
    sec("6. Sample size: how many confident misses does the out-of-sample test need?")
    gapc, vc1, vc2, n1, n2 = f"Results!C{w0+4}", f"Results!C{w0+5}", f"Results!C{w0+6}", f"Results!C{w0}", f"Results!C{w0+1}"
    for lab, v, nf in (("Gap to detect (Part 1 estimate)", f"={gapc}", PCT),
                       ("Confident misses so far: all reports", f"=Results!B{w0}", "0"),
                       ("Confident misses so far: new reports", f"=Results!D{w0}", "0"),
                       ("Needed for a 50% chance of significance at level 2",
                        f'=IFERROR((NORMSINV(1-(1-{A["CL2"]})/2)/ABS({gapc}))^2*({vc1}+{vc2}*{n1}/{n2}),"n/a")', "0"),
                       ("Needed for 80% power at level 2",
                        f'=IFERROR(((NORMSINV(1-(1-{A["CL2"]})/2)+NORMSINV(0.8))/ABS({gapc}))^2*({vc1}+{vc2}*{n1}/{n2}),"n/a")', "0")):
        if lab.startswith("Needed for 80%"):
            pw = r
        put(SUM_, r, 1, lab); put(SUM_, r, 3, v, nf); put(SUM_, r, 9, "Results section E"); r += 1
    r += 1
    Rbr, Rses = R("Odds read before the release (1/0)"), R("Release session")
    for lab, v in (("EDGAR check: reports with a release time", f'=COUNTIF({Rbr},">=0")'),
                   ("EDGAR check: odds read before the release", f'=COUNTIF({Rbr},1)'),
                   ("EDGAR check: odds read after the release (left out of the EDGAR test)", f'=COUNTIF({Rbr},0)'),
                   ("EDGAR check: release before the open / after the close",
                    f'=COUNTIF({Rses},"before the open")&" / "&COUNTIF({Rses},"after the close")')):
        put(SUM_, r, 1, lab); put(SUM_, r, 3, v, "0"); put(SUM_, r, 9, "Events, Raw_EDGAR"); r += 1
    r += 1
    for t in ["Reading: the rule (confident = crowd p at least the threshold) was fixed before the new reports. The new reports are a "
              "consistency check. A full earnings season adds about 12 to 18 confident misses, so two more seasons are needed.",
              "Results are associations over one sample, not causes."]:
        put(SUM_, r, 1, t, font=NOTE); r += 1
    sec("7. Limits, stated plainly (the numbers update with the data)")
    ra_, ru_, rp_, ri_ = bt_row["Outcome: AI trims"], bt_row["Outcome: rule trims"], bt_row["Outcome: past-move rule"], bt_row["Outcome: implied-move rule"]
    lims = [
        f'="New reports: "&Results!D{w0}&" confident misses so far, about "&TEXT(C{pw},"0")&" are needed for an 80% chance to detect the '
        f'Part 1 gap. The miss gap on new reports alone is not confirmed yet."',
        f'="Choosing which reports to trim: the AI (worst 10% "&TEXT(Backtest!E{ra_},"0.0%")&") beats the crowd rule ("&TEXT(Backtest!E{ru_},"0.0%")'
        f'&") and the options-implied move ("&TEXT(Backtest!E{ri_},"0.0%")&"), but is about as good as the simplest spreadsheet that trims '
        f'the biggest past moves ("&TEXT(Backtest!E{rp_},"0.0%")&"). Most of the gain comes from avoiding volatile names."',
        f'="Sizing the downside: the AI worst case scores "&TEXT(Models!F{mrow["q10_ai_vs_sheet"]},"+0.0%;-0.0%")&" against k x past move '
        f'(Models). A one-number spreadsheet does about as well. The AI adds the explanation and the fusion of sources, not a better number."',
        f'="Pre-registered variants (fixed before the run), worst 10% against the past-move rule: fall classifier "'
        f'&TEXT(Backtest!E{bt_row["Outcome: fall-classifier policy"]}-Backtest!E{rp_},"+0.0%;-0.0%")&", AI learns k "'
        f'&TEXT(Backtest!E{bt_row["Outcome: residual-AI policy"]}-Backtest!E{rp_},"+0.0%;-0.0%")&", sized trims "'
        f'&TEXT(Backtest!E{bt_row["Outcome: sized AI trims"]}-Backtest!E{rp_},"+0.0%;-0.0%")&". All three are reported, whatever the result."',
        f'="Beat or miss: no model beats the crowd p alone (Models). The crowd is the forecast, the models are checks."',
        '="Put prices: Black-Scholes at at-the-money implied volatility (no skew), so out-of-the-money puts are priced slightly low."']
    for t in lims:
        put(SUM_, r, 1, t); SUM_.cell(row=r, column=1).alignment = Alignment(wrap_text=True, vertical="top")
        SUM_.merge_cells(start_row=r, start_column=1, end_row=r, end_column=9); SUM_.row_dimensions[r].height = 32; r += 1
    SUM_.column_dimensions["A"].width = 72; SUM_.column_dimensions["B"].width = 18
    for c_ in "CDEFGH":
        SUM_.column_dimensions[c_].width = 12
    SUM_.column_dimensions["I"].width = 26

    for ws in (RM, RO, RP, RE, RS, RED, ROH, RLS, RMOD, RB, RAL, RGA, RKM, RKO, RREL, RFE):
        ws.freeze_panes = "A5"
    for nme, colr in (("Summary", "C00000"), ("Tests", "1F3A5F"), ("Test_Calc", "2E7D32"), ("Weeks", "2E7D32"), ("Results", "1F3A5F"), ("Tail", "1F3A5F"), ("Hedge", "1F3A5F"), ("Backtest", "C00000"), ("Live", "2E7D32"), ("Models", "1F3A5F"), ("GenAI", "1F3A5F"),
                      ("Bootstrap", "1F3A5F"), ("Kalshi_Results", "1F3A5F"), ("Events", "2E7D32"), ("Model_Calc", "2E7D32"),
                      ("Kalshi_Events", "2E7D32"), ("Assumptions", "A6A6A6")):
        W[nme].sheet_properties.tabColor = colr
    try:
        add_extra_sheets(wb, A)
    except Exception as e:
        print(f"   New sheets (Scorecard, AI_Checks, Agent) failed, workbook saved without them: {e}", flush=True)
    wb.calculation.fullCalcOnLoad = True
    wb.save(OUTPUT_FILE)
    return n, nk


# 8. Run
# 9. Scorecard from two saved workbooks, without rerunning STEP 1 and STEP 2 (STEP = "score")
SC_COLS = ["Contract", "Ticker", "Report date", "Flag made (UTC)", "Release (UTC, EDGAR)", "Flag made before the release (1/0)",
             "Crowd p", "AI 10% worst case", "AI cut-off", "AI action", "Rule action", "Put price / share price", "Abnormal return",
             "Share return", "Outcome (1 = beat)"]
SC_PRICE_COLS = ["Close before (share)", "Close after (share)", "Close before (SPY)", "Close after (SPY)", "Price status",
              "Outcome source"]
SC_OUT = f"Group{GROUP_NUMBER}_P3_Workbook.xlsx"


def sc_upload():
    try:
        from google.colab import files
        print("Upload two saved workbooks (select both files in the dialog).", flush=True)
        return files.upload()
    except ImportError:
        import sys
        return {p: open(p, "rb").read() for p in sys.argv[1:3]}


def sc_raw_df(wb, name):
    rows = list(wb[name].iter_rows(min_row=4, values_only=True))
    return pd.DataFrame(rows[1:], columns=[str(c) for c in rows[0]]).dropna(how="all")


def sc_snap_time(data):
    import openpyxl
    wb = openpyxl.load_workbook(io.BytesIO(data), read_only=True)
    r = next(wb["Raw_Snapshots"].iter_rows(min_row=5, max_row=5, values_only=True))
    return str(r[0])


def sc_release_utc(acc):
    try:
        g = pd.Timestamp(str(acc)[:19].replace("T", " "))
    except Exception:
        return None
    d0 = pd.Timestamp(g.year, 3, 8); d0 += pd.Timedelta(days=(6 - d0.weekday()) % 7)
    d1 = pd.Timestamp(g.year, 11, 1); d1 += pd.Timedelta(days=(6 - d1.weekday()) % 7)
    off = pd.Timedelta(hours=4 if d0 <= g < d1 else 5)
    for ny in (g - off, g):
        h = ny.hour + ny.minute / 60
        if 6 <= h < 9.5 or 16 <= h < 17.6:
            return ny + off
    return None


def sc_edgar_times(pairs):
    s = requests.Session()
    s.headers.update({"User-Agent": SEC_USER_AGENT, "Accept-Encoding": "gzip, deflate"})
    tk = s.get("https://www.sec.gov/files/company_tickers.json", timeout=30).json()
    cik = {str(v["ticker"]).upper(): int(v["cik_str"]) for v in tk.values()}
    out = {}
    for t, rd in pairs:
        c = cik.get(t.upper()) or cik.get(t.upper().replace(".", "-"))
        if not c:
            continue
        time.sleep(0.15)
        try:
            js = s.get(f"https://data.sec.gov/submissions/CIK{c:010d}.json", timeout=30).json()["filings"]["recent"]
        except Exception:
            continue
        f = pd.DataFrame(js)
        f = f[(f["form"] == "8-K") & f["items"].astype(str).str.contains("2.02", regex=False)]
        f = f[(pd.to_datetime(f["filingDate"]) >= rd - pd.Timedelta(days=1)) & (pd.to_datetime(f["filingDate"]) <= rd + pd.Timedelta(days=1))]
        for a in f["acceptanceDateTime"]:
            u = sc_release_utc(a)
            if u is not None and abs((u.normalize() - rd).days) <= 1:
                out[(t, rd.strftime("%Y-%m-%d"))] = u
                break
    return out


def sc_outcome(slug, ticker, rd):
    try:
        js = requests.get("https://gamma-api.polymarket.com/markets", params={"slug": slug}, timeout=20).json()
        m = js[0] if isinstance(js, list) and js else {}
        px = [float(x) for x in re.findall(r"[\d.]+", str(m.get("outcomePrices", "")))]
        if m.get("closed") and px and max(px) > 0.98:
            return (1 if px[0] > 0.98 else 0), "Polymarket"
    except Exception:
        pass
    try:
        import yfinance as yf
        e = yf.Ticker(yf_symbol(ticker)).get_earnings_dates(limit=8)
        e.index = pd.to_datetime(e.index).tz_localize(None).normalize()
        r = e[(e.index >= rd - pd.Timedelta(days=1)) & (e.index <= rd + pd.Timedelta(days=1))].dropna(subset=["Reported EPS"])
        if len(r):
            return int(r["Reported EPS"].iloc[0] > r["EPS Estimate"].iloc[0]), "Yahoo (Polymarket not settled)"
    except Exception:
        pass
    return None, "not available yet"


def score_saved_workbooks():
    import openpyxl
    import yfinance as yf
    from openpyxl.styles import Font
    up = sc_upload()
    while len(up) < 2:
        print(f"Received {len(up)} workbook. Upload the other one now (29 October and 30 October are both needed).", flush=True)
        more = sc_upload()
        if not more:
            raise SystemExit("Stopped: two workbooks are needed.")
        up.update(more)
    snaps = {n: sc_snap_time(d) for n, d in up.items()}
    if len(set(snaps.values())) < 2:
        raise SystemExit(f"Both files have the same snapshot ({list(snaps.values())[0]}). Upload the 29 October and the 30 October workbooks.")
    f29 = min(snaps, key=snaps.get)
    f30 = max(snaps, key=snaps.get)
    print(f"Flags from {f29} (snapshot {snaps[f29]} UTC). Scorecard written into {f30} (snapshot {snaps[f30]} UTC).", flush=True)

    w29 = openpyxl.load_workbook(io.BytesIO(up[f29]), read_only=True)
    al, sn = sc_raw_df(w29, "Raw_AI_Live"), sc_raw_df(w29, "Raw_Snapshots")
    al["rd"] = pd.to_datetime(al["Report date"]).dt.normalize()
    flag_ts = pd.Timestamp(snaps[f29])
    now = pd.Timestamp.now(tz="UTC").tz_localize(None)
    sn["slug"] = sn["slug"].astype(str)
    put = dict(zip(sn["slug"], pd.to_numeric(sn["put95_cost"], errors="coerce")))
    due = al[al["rd"] < now.normalize()].copy()
    print(f"Flags in that workbook: {len(al)}. Report day already passed: {len(due)}.", flush=True)
    if due.empty:
        print("No flagged report has passed yet.")
        return
    tick = sorted(set(due["Ticker"].astype(str)))
    start = (due["rd"].min() - pd.Timedelta(days=10)).strftime("%Y-%m-%d")
    px = yf.download([yf_symbol(t) for t in tick] + [MARKET], start=start, auto_adjust=True, progress=False)["Close"]
    px.index = pd.to_datetime(px.index).tz_localize(None).normalize()
    rel = sc_edgar_times([(str(t), d) for t, d in zip(due["Ticker"], due["rd"])])
    print(f"EDGAR release times found: {len(rel)} of {len(due)}", flush=True)
    rows = []
    for x in due.itertuples(index=False):
        t, rd = str(x.Ticker), x.rd
        ru = rel.get((t, rd.strftime("%Y-%m-%d")))
        limit = ru if ru is not None else rd + pd.Timedelta(hours=CUTOFF_HOUR_UTC)
        before = int(flag_ts < limit)
        dates = px.index
        a = dates.searchsorted(rd - pd.Timedelta(days=1), side="right") - 1
        b = dates.searchsorted(rd, side="right")
        sym = yf_symbol(t)
        c = [np.nan] * 4
        status = "window not closed yet"
        if sym in px.columns and a >= 0 and b < len(dates):
            c = [px[sym].iloc[a], px[sym].iloc[b], px[MARKET].iloc[a], px[MARKET].iloc[b]]
            final = dates[b] < now.normalize() or now.hour * 60 + now.minute >= 20 * 60 + 15
            status = "close" if final else f"last trade at {now:%H:%M} UTC on {now:%d %b %Y}"
        q10, cut = float(x._4), float(x._6)
        y, ysrc = sc_outcome(x._0, t, rd) if before else (None, "")
        rows.append([x._0, t, rd.to_pydatetime(), flag_ts.to_pydatetime(), ru.to_pydatetime() if ru is not None else None, before,
                     float(x._3), q10, cut, "trim 50%" if q10 <= cut else "hold", x._7, put.get(str(x._0)),
                     None, None, y] + c + [status, ysrc])
    df = pd.DataFrame(rows, columns=SC_COLS + SC_PRICE_COLS)
    print(f"Flags made before the release: {int(df.iloc[:, 5].sum())}. With prices after the report: "
          f"{int(((df.iloc[:, 5] == 1) & df['Close after (share)'].notna()).sum())}.", flush=True)

    # ---------- write into the submission workbook ----------
    wb = openpyxl.load_workbook(io.BytesIO(up[f30]))
    ws = wb["Raw_Scorecard"]
    asm = {str(r[0].value): f"Assumptions!$B${r[0].row}" for r in wb["Assumptions"].iter_rows(min_row=4, max_col=1) if r[0].value}
    A = lambda part: next(v for k, v in asm.items() if k.startswith(part))
    HP, FALL, TRIM, TCOST, OTM = (A("Confident crowd: p at least"), A("A fall: abnormal return at or below"),
                                  A("Backtest: share of the position sold"), A("Backtest: round-trip trading cost"),
                                  A("Protective put: strike this far below"))
    ws["A1"] = (f"Live scorecard: the flags saved in the workbook of {flag_ts:%d %b %Y %H:%M} UTC (file uploaded before the reports), "
                f"scored with Yahoo Finance prices, SEC EDGAR release times and Polymarket outcomes, retrieved {now:%d %b %Y %H:%M} UTC. "
                "Columns A to L, O and Z to AE as saved or retrieved. Columns M, N and P to Y are formulas. Close after = the close of the first "
                "trading day after the report, or the last trade at the time in column AD if that day has not closed yet.")
    ws["A1"].font = Font(name="Arial", bold=True)
    for j, h in enumerate(SC_PRICE_COLS):
        ws.cell(row=4, column=26 + j, value=h).font = Font(name="Arial", bold=True)
    E = 4 + len(df)
    for i, rec in enumerate(df.itertuples(index=False)):
        q = 5 + i
        for j, v in enumerate(rec):
            if j in (12, 13):
                continue
            if isinstance(v, float) and not np.isfinite(v):
                v = None
            ws.cell(row=q, column=(1 + j) if j < 15 else (26 + j - 15), value=v)
        ws.cell(row=q, column=3).number_format = "yyyy-mm-dd"
        for col in (4, 5):
            ws.cell(row=q, column=col).number_format = "yyyy-mm-dd hh:mm"
        ws.cell(row=q, column=13, value=f'=IF(AND(ISNUMBER(Z{q}),ISNUMBER(AA{q}),ISNUMBER(AB{q}),ISNUMBER(AC{q})),LN(AA{q}/Z{q})-LN(AC{q}/AB{q}),"")').number_format = "+0.0%;-0.0%"
        ws.cell(row=q, column=14, value=f'=IF(AND(ISNUMBER(Z{q}),ISNUMBER(AA{q})),AA{q}/Z{q}-1,"")').number_format = "+0.0%;-0.0%"
        for col, nf in ((7, "0%"), (8, "+0.0%;-0.0%"), (9, "+0.0%;-0.0%"), (12, "0.00%")):
            ws.cell(row=q, column=col).number_format = nf
        f = {16: f'=IF(AND(ISNUMBER(H{q}),ISNUMBER(I{q})),IF(H{q}<=I{q},1,0),"")',
             17: f'=IF(ISNUMBER(G{q}),IF(G{q}>={HP},1,0),"")',
             18: f'=IF(AND(F{q}=1,ISNUMBER(M{q})),1,0)',
             19: f'=IF(R{q}=1,IF(M{q}<={FALL},1,0),"")',
             20: f'=IF(AND(R{q}=1,ISNUMBER(H{q})),IF(M{q}<H{q},1,0),"")',
             21: f'=IF(R{q}=1,M{q},"")',
             22: f'=IF(R{q}=1,M{q}*(1-{TRIM}*N(Q{q}))-{TCOST}*{TRIM}*N(Q{q}),"")',
             23: f'=IF(R{q}=1,M{q}*(1-{TRIM}*N(P{q}))-{TCOST}*{TRIM}*N(P{q}),"")',
             24: f'=IF(AND(R{q}=1,ISNUMBER(N{q})),MAX(0,(1-{OTM})-(1+N{q})),"")',
             25: f'=IF(AND(ISNUMBER(X{q}),ISNUMBER(L{q})),X{q}-L{q},"")'}
        for col, v in f.items():
            c_ = ws.cell(row=q, column=col, value=v)
            if col >= 21:
                c_.number_format = "+0.0%;-0.0%"
    pat = re.compile(r"Raw_Scorecard!\$([A-Z]+)\$5:\$([A-Z]+)\$\d+")
    for sh in ("Scorecard", "Summary"):
        for row in wb[sh].iter_rows():
            for c_ in row:
                if isinstance(c_.value, str) and "Raw_Scorecard!" in c_.value:
                    c_.value = pat.sub(lambda m: f"Raw_Scorecard!${m.group(1)}$5:${m.group(2)}${E}", c_.value)
    sc = wb["Scorecard"]
    sc["A2"] = (f"Flags saved on {flag_ts:%d %b %Y %H:%M} UTC, before the reports, scored once the reports were out. A flag counts only if "
                "it was made before the company's release (SEC EDGAR time, or 10:00 UTC on the report day if unknown). Outcome = "
                "abnormal return over the same window as Events. Few reports so far: a track record that grows each week, not a test.")
    wb.calculation.fullCalcOnLoad = True
    wb.save(SC_OUT)
    print(f"Saved {SC_OUT}. Open it in Excel so the formulas calculate, then look at the Scorecard sheet.", flush=True)
    s = df[(df.iloc[:, 5] == 1) & df["Close after (share)"].notna()].copy()
    if len(s):
        s["ar"] = np.log(s["Close after (share)"] / s["Close before (share)"]) - np.log(s["Close after (SPY)"] / s["Close before (SPY)"])
        tr = s["AI action"].str.startswith("trim")
        print(f"Preview: {len(s)} scored. AI said trim on {int(tr.sum())}: average move {s.loc[tr, 'ar'].mean():+.1%}. "
              f"AI said hold on {int((~tr).sum())}: average move {s.loc[~tr, 'ar'].mean():+.1%}. "
              f"Below the AI worst case: {int((s['ar'] < s['AI 10% worst case']).sum())} of {len(s)}. "
              f"Status: {s['Price status'].iloc[0]}.", flush=True)
        print(s[["Ticker", "Crowd p", "AI 10% worst case", "AI action", "Rule action", "ar", "Price status"]].to_string(index=False))
    try:
        from google.colab import files
        files.download(SC_OUT)
    except ImportError:
        pass


def run():
    global DATA_DIR
    t0 = time.time()
    if STEP == "score":
        score_saved_workbooks()
        return
    os.makedirs(DATA_DIR, exist_ok=True)
    if STEP in (1, 2):
        ask_for_zips()
    if not restore_zips():
        if STEP != 2:
            print("Downloading all data from the public sources (Polymarket, Kalshi, Yahoo Finance, SEC EDGAR). About 20 to 30 minutes.", flush=True)
    download_utc = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M")
    if STEP == 2:
        step2(download_utc, t0)
        return
    if LSEG_ONLY:
        if not os.path.exists(_p(F_MARKETS)) or not os.path.exists(_p(F_ODDS)):
            print(f"LSEG_ONLY needs the data: upload {RAW_ZIP} to the Colab folder first, then run again.", flush=True)
            return
        print("LSEG only: no downloads. Building the request file from the saved data.", flush=True)
        m_, o_, p_, e_, *_ = load_data()
        ev_ = build_events(m_, o_, p_, e_)
        if lseg_request(ev_, p_):
            lseg_interactive()
        else:
            print("   No reports to request.", flush=True)
        build_from_saved(download_utc)
        print(f"Done in {time.time() - t0:.0f} s.", flush=True)
        return
    print("1/6 Polymarket earnings contracts", flush=True)
    mk = fetch_markets()
    mk.to_csv(_p(F_MARKETS), index=False)
    if TAKE_SNAPSHOT:
        print("2/6 Snapshot: odds, option prices and analyst revisions for upcoming reports", flush=True)
        take_snapshot(mk)
    if SNAPSHOT_ONLY:
        save_zip_and_download(workbook=False)
        print(f"Snapshot saved and {RAW_ZIP} downloaded. Done in {time.time() - t0:.0f} s.", flush=True)
        return
    download_all(mk)
    try:
        m_, o_, p_, e_, *_ = load_data()
        ev_ = build_events(m_, o_, p_, e_)
        fetch_edgar(ev_)
    except Exception as e:
        ev_ = None
        print(f"   EDGAR step failed, continuing without it: {e}", flush=True)
    for step, fn in (("Previous earnings releases (EDGAR)", lambda: fetch_prior_releases(ev_, load_data()[4])),
                     ("Sectors (Yahoo)", lambda: fetch_sectors(set(m_["ticker"].dropna()))),
                     ("Generative AI reading the releases", score_releases)):
        try:
            if ev_ is not None:
                fn()
        except Exception as e:
            print(f"   {step} failed, continuing without it: {e}", flush=True)
    n_req = 0
    try:
        if ev_ is not None:
            n_req = lseg_request(ev_, p_)
            if n_req and LSEG_INTERACTIVE:
                lseg_interactive()
    except Exception as e:
        print(f"   LSEG step failed, continuing without it: {e}", flush=True)
    try:
        if ev_ is not None:
            fetch_option_history(ev_, p_)
    except Exception as e:
        print(f"   Option history step failed, continuing without it: {e}", flush=True)
    with open(_p(F_LOG), "a") as fh:
        fh.write(f"{download_utc}\n")
    if n_req and not LSEG_INTERACTIVE:
        save_zip_and_download(workbook=False, download=False)
        try:
            from google.colab import files
            files.download(RAW_ZIP)
            files.download(_p(LSEG_REQUEST))
        except Exception:
            pass
        print("\n" + "=" * 78 + f"\nSTEP 1 done in {time.time() - t0:.0f} s. Raw data saved. {LSEG_REQUEST} is downloading.\n"
              "  1. Open it in Excel with LSEG Workspace signed in. Setup!B6 should show a number (Apple, about 26.21).\n"
              "  2. Wait until columns L, M, T, U, X, AA and AC on the Request sheet show numbers (Setup B6, C6, B7, B8 test them), then save.\n"
              "  3. In a new Colab cell run:   STEP = 2   and   run()   then upload the saved file.\n" + "=" * 78, flush=True)
        return
    build_from_saved(download_utc)
    print(f"Done in {time.time() - t0:.0f} s.", flush=True)


MERGE_KEYS = {F_SNAP: ["snap_utc", "slug"], F_FLAGS: ["slug", "snap_utc"], F_HASH: ["snap_utc"], F_ODDS: ["slug", "ts"],
              F_KODDS: ["market", "ts"], F_AI_REL: ["accession"], F_PRIOR: ["ticker", "report_date"], F_KMAP: ["series"],
              F_AGENT: ["asked_utc"], F_GUARD: ["kind", "key", "text"], F_EDGAR: ["ticker", "report_date"], F_LSEG: ["slug"],
              F_SECTOR: ["ticker"], F_EPS: ["ticker", "date"], F_KM: ["mticker"], F_STRESS: ["slug", "variant"], F_AI_PARSE: ["slug"],
              F_OPTHIST: ["slug"]}
KEEP_FIRST = {F_FLAGS, F_HASH}


def restore_zips():
    import glob
    zs = sorted(set(glob.glob(f"Group{GROUP_NUMBER}_P3_RawData*.zip") + glob.glob(f"Group{GROUP_NUMBER}_P2_RawData*.zip") + glob.glob(f"Group{GROUP_NUMBER}_P1_RawData*.zip")))
    if not zs:
        return False
    os.makedirs(DATA_DIR, exist_ok=True)

    def newest_snap(z):
        try:
            with z.open(next(n for n in z.namelist() if n.endswith(F_SNAP))) as fh:
                return str(pd.read_csv(fh)["snap_utc"].max())
        except Exception:
            return ""
    srcs = []
    for zn in zs:
        try:
            z = zipfile.ZipFile(zn)
        except Exception:
            print(f"   {zn} is not a valid zip, skipped", flush=True)
            continue
        srcs.append((("0" if "_P1_" in zn else "1") + newest_snap(z), zn, z))
    srcs.sort(key=lambda x: x[0])
    names = sorted({os.path.basename(n) for _, _, z in srcs for n in z.namelist() if n.endswith((".csv", ".json"))})
    for nm in names:
        frames = []
        for _, zn, z in srcs:
            hit = [n for n in z.namelist() if os.path.basename(n) == nm]
            if hit:
                with z.open(hit[0]) as fh:
                    frames.append(fh.read())
        if os.path.exists(_p(nm)):
            with open(_p(nm), "rb") as fh:
                frames.append(fh.read())
        if nm in MERGE_KEYS and nm.endswith(".csv"):
            import io
            dfs = [pd.read_csv(io.BytesIO(b)) for b in frames if b]
            d = pd.concat(dfs, ignore_index=True)
            keys = [k for k in MERGE_KEYS[nm] if k in d]
            if keys:
                d = d.drop_duplicates(keys, keep="first" if nm in KEEP_FIRST else "last")
            d.to_csv(_p(nm), index=False)
        else:
            with open(_p(nm), "wb") as fh:
                fh.write(frames[-1])
    for _, _, z in srcs:
        z.close()
    print(f"Restored and merged data from {len(srcs)} zip file(s): {', '.join(zn for _, zn, _ in srcs)}", flush=True)
    return True


def _upload_one(msg):
    """One upload dialog (one file is enough). Returns {} after Cancel, None outside Colab."""
    try:
        from google.colab import files
    except ImportError:
        return None
    print(msg, flush=True)
    try:
        up = files.upload() or {}
    except Exception:
        up = {}
    for name, data in up.items():
        with open(os.path.basename(name), "wb") as fh:
            fh.write(data)
    return up


def ask_for_zips():
    import glob
    if glob.glob(f"Group{GROUP_NUMBER}_P*_RawData*.zip"):
        return
    while True:
        up = _upload_one(f"Upload a saved data zip (for example Group{GROUP_NUMBER}_P2_RawData.zip), one file at a time. "
                         "Press Cancel when there are no more, or to start from zero.")
        if not up:
            return
        print(f"   received {', '.join(up)}", flush=True)


def step2(download_utc, t0):
    print("STEP 2: upload the filled GroupA2_LSEG_request.xlsx.", flush=True)
    while True:
        up = _upload_one("Upload one file: the filled GroupA2_LSEG_request.xlsx. If Colab restarted after STEP 1, upload the RawData zip "
                         "first, then the LSEG file. Cancel = build the workbook without LSEG.")
        if up is None:
            path = input("Path to the filled LSEG file (Enter to skip): ").strip() if sys.stdin and sys.stdin.isatty() else ""
            up = {os.path.basename(path): open(path, "rb").read()} if path else {}
        got_x = False
        for name, data in up.items():
            if name.lower().endswith(".zip"):
                restore_zips()
            elif name.lower().endswith((".xlsx", ".xlsm")):
                with open(_p(LSEG_REQUEST), "wb") as fh:
                    fh.write(data)
                got_x = True
                print(f"   received {name}", flush=True)
        if got_x or not up:
            break
    if not os.path.exists(_p(F_MARKETS)):
        print(f"   No data found: Colab restarted after STEP 1. Run STEP 1 again, or upload the {RAW_ZIP} it downloaded with the LSEG file.", flush=True)
        return
    lseg_collect()
    build_from_saved(download_utc)
    print(f"Done in {time.time() - t0:.0f} s.", flush=True)


def build_from_saved(download_utc):
    print("6/6 Models, bootstrap and workbook", flush=True)
    mk, od, px, est, snap, km, ko = load_data()
    ev = build_events(mk, od, px, est)
    ls = load_lseg()
    ev = add_lseg_features(ev, ls)
    try:
        score_releases()
    except Exception as e:
        print(f"   Generative AI reading the releases failed: {e}", flush=True)
    rel = load_releases()
    ev = add_peer_features(add_text_features(ev, rel), est, px)
    print(f"   last release read by generative AI: {ev['guid_score'].notna().sum()} reports, peers: {ev['peer_mean'].notna().sum()}, "
          f"skew: {ev['skew'].notna().sum()}", flush=True)
    print(f"   settled reports: {len(ev)} (new since {FREEZE_DATE}: {(ev['sample'] == 'New').sum()}), snapshots: {len(snap)}, "
          f"LSEG implied volatility: {ls['iv_entry'].notna().sum()}, LSEG consensus EPS: {ev['rev_lseg'].notna().sum()}, "
          f"StarMine: {ev['smart_surp'].notna().sum()}", flush=True)
    try:
        ev = add_kalshi_features(ev, km, ko)
    except Exception as e:
        print(f"   Kalshi KPI input failed, models run without it: {e}", flush=True)
        ev = ev.assign(k_kpi_p=np.nan, k_kpi_n=0.0)
    save_hedge_stats(ev)
    models = run_models(ev)
    ev = add_backtest(ev, models)
    boot = run_bootstrap(ev)
    try:
        log_flags()
    except Exception as e:
        print(f"   Flag log failed: {e}", flush=True)
    try:
        live_ai(ev, snap, px, est, mk)
    except Exception as e:
        print(f"   Live AI step failed, continuing without it: {e}", flush=True)
    try:
        genai_parse_check(mk, km)
    except Exception as e:
        print(f"   Generative AI parser check failed, continuing without it: {e}", flush=True)
    for step, fn in (("Live scorecard", lambda: score_flags(px, mk)), ("Parser stress test", lambda: stress_test(mk)),
                     ("Portfolio agent demo", lambda: ask(AGENT_DEMO, show=False) if AGENT_DEMO else None),
                     ("Guardrail check", guard_notes)):
        try:
            fn()
        except Exception as e:
            print(f"   {step} failed, continuing without it: {e}", flush=True)
    n, nk = build(mk, od, px, est, snap, km, ko, ev, models, boot, download_utc, load_edgar(), load_opthist(), ls, rel)
    save_zip_and_download(workbook=True, download=False)
    print(f"   {OUTPUT_FILE}: {n} Polymarket reports, {nk} Kalshi KPI events. Raw data: {RAW_ZIP}", flush=True)
    print("   Open the workbook in Excel so the formulas calculate. Start with Summary.", flush=True)
    save_zip_and_download(workbook=True, write=False)


def save_zip_and_download(workbook=True, download=True, write=True):
    if write:
        with zipfile.ZipFile(RAW_ZIP, "w", zipfile.ZIP_DEFLATED) as z:
            for f in (F_MARKETS, F_ODDS, F_PRICES, F_EPS, F_SNAP, F_KM, F_KODDS, F_EDGAR, F_OPTHIST, F_LSEG, F_AI_PARSE, F_AI_LIVE, F_PRIOR, F_AI_REL, F_SECTOR,
                      F_KALL, F_KMAP, F_STRESS, F_FLAGS, F_SCORE, F_GUARD, F_AGENT, F_HASH, F_HEDGE):
                if os.path.exists(_p(f)):
                    z.write(_p(f), arcname=f)
    if not download:
        return
    try:
        from google.colab import files
        files.download(RAW_ZIP)
        if workbook:
            files.download(OUTPUT_FILE)
            if USE_LSEG and not LSEG_INTERACTIVE and os.path.exists(_p(LSEG_REQUEST)):
                files.download(_p(LSEG_REQUEST))
    except Exception:
        pass


if __name__ == "__main__":
    run()
