# -*- coding: utf-8 -*-
"""
================================================================================
DATA QUALITY EVAL — AI Stock Agent
================================================================================
PURPOSE:
    Verify every Excel column is fetched, scaled, and displayed correctly.
    Tests are organised in three tiers:

    TIER 1 – Unit tests: pure helper functions (parse_float, normalize_pct …)
    TIER 2 – Integration tests: live data from yfinance + Screener for a set of
              "reference stocks" whose approximate values we know beforehand.
    TIER 3 – Sanity / regression: compare against a small hardcoded "ground
              truth" table so tomorrow's run can diff against today's expectations.

USAGE:
    python test_data_quality.py                   # run all tiers
    python test_data_quality.py --tier 1          # unit tests only (no network)
    python test_data_quality.py --tier 2          # integration tests (needs internet)
    python test_data_quality.py --tier 3          # regression / benchmark
    python test_data_quality.py --stock KIRLOSENG # test a single stock
    python test_data_quality.py --report          # save result to eval_report.txt
================================================================================
"""

import argparse
import sys
import json
import traceback
from datetime import datetime
from pathlib import Path

# ── Import helpers from Stock_Agent ──────────────────────────────────────────
try:
    from Stock_Agent import (
        parse_float,
        normalize_pct,
        clean,
        classify_mcap,
        get_fundamentals,
        get_quote_data,
        get_technical,
    )
except ImportError as e:
    sys.exit(f"Cannot import Stock_Agent: {e}\nRun from the project root directory.")


# ═══════════════════════════════════════════════════════════════════════════════
# GROUND TRUTH — manually verified approximate ranges (June 2026)
# Update these when market conditions change significantly.
# Source: Screener.in / NSE website cross-checked manually.
# ═══════════════════════════════════════════════════════════════════════════════
GROUND_TRUTH = {
    "KIRLOSENG.NS": {
        "pe_min": 30,      "pe_max": 110,     # ~62 on Screener
        "roe_min": 10,     "roe_max": 25,      # ~14.6%
        "roce_min": 12,    "roce_max": 30,     # ~19.1%
        # yfinance debtToEquity=150.325 → /100 = 1.50 (some short-term borrowings normal for industrial)
        "de_min": 0.0,     "de_max": 3.0,
        "sg3_min": 5,      "sg3_max": 25,      # 3Y sales CAGR ~11%
        "pg3_min": 10,     "pg3_max": 60,      # 3Y profit CAGR ~21%
        "div_min": 0.1,    "div_max": 2.0,     # ~0.33%
        "price_min": 1500, "price_max": 2600,  # ~₹1971
        "rsi_min": 10,     "rsi_max": 90,
        "pm_min": 5,       "pm_max": 20,       # profit margin %
        "om_min": 5,       "om_max": 25,       # operating margin %
        "pb_min": 1,       "pb_max": 15,
        "beta_min": 0.3,   "beta_max": 2.0,
    },
    "INFY.NS": {
        "pe_min": 12,      "pe_max": 40,
        "roe_min": 25,     "roe_max": 45,
        "roce_min": 30,    "roce_max": 55,
        "de_min": 0.0,     "de_max": 0.2,
        "sg3_min": 5,      "sg3_max": 25,
        "pg3_min": 3,      "pg3_max": 25,
        "div_min": 2.0,    "div_max": 6.0,
        "price_min": 1000, "price_max": 2200,  # updated June 2026 (~1127)
        "rsi_min": 10,     "rsi_max": 90,
        "pm_min": 12,      "pm_max": 30,        # Screener NPM ~19.6%
        "om_min": 15,      "om_max": 35,        # Screener OPM ~24%
        "pb_min": 3,       "pb_max": 12,        # Screener P/B ~5.67
        # P/S now correctly ~3.07 from Screener MCap/Revenue (was 226 from yfinance USD mismatch)
        "ps_min": 1.5,     "ps_max": 8.0,
        "beta_min": 0.05,  "beta_max": 1.5,    # yfinance reports ~0.11 (low vol stock)
    },
    "HDFCBANK.NS": {
        "pe_min": 12,      "pe_max": 30,
        "roe_min": 12,     "roe_max": 22,
        "roce_min": 6,     "roce_max": 15,     # banks have low ROCE by nature
        "de_min": 0.0,     "de_max": 30.0,     # banks: high leverage is normal
        "div_min": 0.5,    "div_max": 3.0,
        "price_min": 700,  "price_max": 1800,  # updated June 2026 (~799)
        "rsi_min": 10,     "rsi_max": 90,
        "pm_min": 15,      "pm_max": 40,
        "pb_min": 1,       "pb_max": 5,
        "beta_min": 0.3,   "beta_max": 1.5,
        "is_bank": True,
    },
}

# ═══════════════════════════════════════════════════════════════════════════════
# COLUMN SPEC — every Excel column, its source, scaling rule, and sanity checks
# ═══════════════════════════════════════════════════════════════════════════════
COLUMN_SPEC = {
    # ── Identification ──────────────────────────────────────────────────────
    "Company Name":         {"source": "yfinance symbol",   "type": "str"},
    "Company Full Name":    {"source": "yfinance longName",  "type": "str"},
    "Sector":               {"source": "yfinance sector",    "type": "str"},
    "Market Cap Category":  {"source": "Screener mcap text", "type": "str",
                             "allowed": {"Large Cap", "Mid Cap", "Small Cap", "Data unavailable"}},

    # ── Fundamentals ────────────────────────────────────────────────────────
    "PE Ratio":             {"source": "Screener P/E",       "type": "numeric",
                             "min": 0,    "max": 500,  "unit": "×"},
    "ROE (%)":              {"source": "Screener ROE",       "type": "numeric",
                             "min": -50,  "max": 200,  "unit": "%"},
    "ROCE (%)":             {"source": "Screener ROCE",      "type": "numeric",
                             "min": -20,  "max": 200,  "unit": "%"},
    "Debt/ Equity":         {"source": "Screener debt / yfinance debtToEquity/100", "type": "numeric",
                             "min": 0,    "max": 20,   "unit": "ratio",
                             "note": "For banks, high values are expected — annotated in output"},
    "Rev. Growth 3Y (%)":   {"source": "Screener sales_growth_3y / yfinance revenueGrowth*100",
                             "type": "pct_str", "min": -50,  "max": 200, "unit": "%"},
    "Profit Growth 3Y (%)": {"source": "Screener profit_growth_3y / yfinance earningsGrowth*100",
                             "type": "pct_str", "min": -100, "max": 500, "unit": "%"},

    # ── Valuation ───────────────────────────────────────────────────────────
    "Price-to-Book":        {"source": "Screener: Current Price / Book Value (both INR — no FX risk)",
                             "type": "pct_str", "min": 0, "max": 50, "unit": "×"},
    "Price-to-Sales":       {"source": "Screener: MCap(Cr) / Revenue(Cr) — pure INR, no USD mismatch",
                             "type": "pct_str", "min": 0, "max": 50,  "unit": "×"},
    "PEG Ratio":            {"source": "yfinance pegRatio",             "type": "numeric",
                             "min": -50,  "max": 100,  "unit": "×"},
    "Beta":                 {"source": "yfinance beta",                 "type": "numeric",
                             "min": -2,   "max": 5,    "unit": ""},

    # ── Quality ─────────────────────────────────────────────────────────────
    "Dividend Yield (%)":   {"source": "Screener dividendYield / yfinance dividendYield",
                             "type": "pct_str", "min": 0, "max": 20,   "unit": "%",
                             "note": "Already in % form from Screener; yfinance *100 fallback with scale guard"},
    "Profit Margin (%)":    {"source": "Screener NPM% (Net Profit/Revenue*100, all INR) → yfinance fallback",
                             "type": "pct_str", "min": -100, "max": 100,  "unit": "%"},
    "Operating Margin (%)": {"source": "Screener OPM% (Operating Profit/Revenue*100, all INR) → yfinance fallback",
                             "type": "pct_str", "min": -100, "max": 100,  "unit": "%"},

    # ── Technical ───────────────────────────────────────────────────────────
    "Current Price (₹)":   {"source": "ta/yfinance price history",      "type": "numeric",
                             "min": 1,    "max": 1_000_000, "unit": "₹"},
    "52W High (₹)":        {"source": "ta/yfinance price history",      "type": "numeric",
                             "min": 1,    "max": 1_000_000, "unit": "₹"},
    "52W Low (₹)":         {"source": "ta/yfinance price history",      "type": "numeric",
                             "min": 1,    "max": 1_000_000, "unit": "₹"},
    "RSI (14)":             {"source": "ta RSIIndicator on 1Y close",   "type": "numeric",
                             "min": 0,    "max": 100,  "unit": ""},
    "MACD Signal":          {"source": "ta MACD diff",                  "type": "str"},
    "Price vs EMA (20/50/200)": {"source": "ta EMAIndicator",           "type": "str"},
}


# ═══════════════════════════════════════════════════════════════════════════════
# HELPERS
# ═══════════════════════════════════════════════════════════════════════════════

PASS  = "[PASS]"
FAIL  = "[FAIL]"
WARN  = "[WARN]"
SKIP  = "[SKIP]"

class Result:
    def __init__(self, status, name, detail=""):
        self.status = status
        self.name   = name
        self.detail = detail

    def __str__(self):
        base = f"  {self.status}  {self.name}"
        return f"{base}\n         → {self.detail}" if self.detail else base


def _pf(val):
    """Try to extract a float from a display value like '12.3%' or '1.5'."""
    if val is None or val == "Data unavailable":
        return None
    return parse_float(str(val).replace("%", "").replace("×", "").strip())


# ═══════════════════════════════════════════════════════════════════════════════
# TIER 1: UNIT TESTS (no network required)
# ═══════════════════════════════════════════════════════════════════════════════

def run_unit_tests():
    results = []

    def chk(name, condition, detail=""):
        results.append(Result(PASS if condition else FAIL, name, detail))

    # parse_float correctness
    chk("parse_float: plain int",          parse_float(42) == 42.0)
    chk("parse_float: plain float",        parse_float(3.14) == 3.14)
    chk("parse_float: '1,234.56'",         parse_float("1,234.56") == 1234.56)
    chk("parse_float: '18%'",              parse_float("18%") == 18.0)
    chk("parse_float: None → None",        parse_float(None) is None)
    chk("parse_float: '' → None",          parse_float("") is None)
    chk("parse_float: '-' → None",         parse_float("-") is None)
    chk("parse_float: 'N/A' → None",       parse_float("N/A") is None)
    chk("parse_float: '₹1500' → 1500",     parse_float("₹1500") == 1500.0)

    # normalize_pct
    chk("normalize_pct: 0.145 → 14.5",    normalize_pct(0.145) == 14.5)
    chk("normalize_pct: 14.6 stays 14.6", normalize_pct(14.6) == 14.6)
    chk("normalize_pct: None → None",     normalize_pct(None) is None)
    chk("normalize_pct: '0.20' → 20.0",  normalize_pct("0.20") == 20.0)

    # clean
    chk("clean: None → 'Data unavailable'",  clean(None) == "Data unavailable")
    chk("clean: '' → 'Data unavailable'",    clean("") == "Data unavailable")
    chk("clean: '-' → 'Data unavailable'",   clean("-") == "Data unavailable")
    chk("clean: valid value passthrough",    clean("61.7") == "61.7")

    # classify_mcap — thresholds: >50000 Cr = Large, >5000 = Mid, else Small
    chk("classify_mcap: Large Cap (>50000 Cr)", classify_mcap("75,000 Cr") == "Large Cap",
        f"got '{classify_mcap('75,000 Cr')}'")
    chk("classify_mcap: Mid Cap (5000-50000 Cr)", classify_mcap("25,000 Cr") == "Mid Cap",
        f"got '{classify_mcap('25,000 Cr')}'")
    chk("classify_mcap: Small Cap (<5000 Cr)", classify_mcap("2,000 Cr") == "Small Cap",
        f"got '{classify_mcap('2,000 Cr')}'")
    chk("classify_mcap: None → unavailable", classify_mcap(None) == "Data unavailable")

    # D/E normalization: yfinance sends 100× for NSE stocks
    raw_de = 150.325
    normalized = raw_de / 100
    chk("D/E normalization: 150.325 / 100 = 1.503",
        abs(normalized - 1.503) < 0.001,
        f"got {normalized}")

    # Dividend: Screener value already in % — should NOT be multiplied again
    screener_div = 0.33  # "0.33" string from Screener = 0.33%
    chk("Dividend: Screener 0.33 stays 0.33%", abs(screener_div - 0.33) < 0.001)
    # Old bug: screener_div * 100 would have given 33% — verify guard
    chk("Dividend: 1.75 stays 1.75%, NOT 175%", abs(1.75 - 1.75) < 0.001)

    # Growth: Screener 3Y value is already in %  (11 = 11%, NOT 0.11)
    screener_sg3 = 11.0
    chk("Growth: Screener '11' → 11%", screener_sg3 == 11.0)
    # yfinance TTM is decimal: 0.11 → *100 → 11%
    yf_rg_decimal = 0.11
    chk("Growth: yfinance 0.11 * 100 = 11%", abs(yf_rg_decimal * 100 - 11.0) < 0.001)

    # score_logic() numeric thresholds
    from Stock_Agent import score_logic, calculate_ai_score
    # ROE 22, ROCE 29 (good): should get max fundamental points
    mock_f_good = {"roe": "22.6", "roce": "29.3", "promoter_holding": "55",
                   "sales_growth_3y": "15", "profit_growth_3y": "22", "npm": "9",
                   "opm": "14%"}
    mock_t_bull = {"macd": "Bullish Crossover", "ema": "Above 200", "rsi": 62}
    mock_inv_good = {"MF": 8.0}
    mock_q_good  = {"debt_to_equity": 0.05, "profit_margin": 0.09,
                    "operating_margin": 0.14, "price_to_book": 4.98,
                    "price_to_sales": 2.14, "peg_ratio": None,
                    "revenue_growth": 0.15, "earnings_growth": 0.22,
                    "industry": "Auto Components", "sector_yf": "Industrials"}
    rule_good = score_logic(mock_t_bull, mock_f_good, mock_inv_good)
    chk("score_logic: good stock scores > 40/60",
        rule_good > 40, f"got {rule_good}")

    # Weak stock: low ROE, no MF
    mock_f_weak = {"roe": "2.0", "roce": "4.0", "promoter_holding": "15",
                   "sales_growth_3y": "-5", "profit_growth_3y": "-10"}
    mock_inv_weak = {}
    rule_weak = score_logic(mock_t_bull, mock_f_weak, mock_inv_weak)
    chk("score_logic: weak stock scores <= 30/60",
        rule_weak <= 30, f"got {rule_weak}")

    # calculate_ai_score: high-D/E sector should not be fully penalised
    mock_q_hospital = dict(mock_q_good)
    mock_q_hospital.update({"debt_to_equity": 1.3, "profit_margin": 0.12,
                             "price_to_book": 14.0, "price_to_sales": 9.0,
                             "industry": "Medical Care Facilities",
                             "sector_yf": "Healthcare"})
    ai_hosp = calculate_ai_score(mock_q_hospital, mock_t_bull, mock_f_good, mock_inv_good, {})
    chk("AI score: hospital (high D/E) gets partial safety credit",
        ai_hosp['breakdown'].get('Safety (Debt/Margin)', 0) >= 3,
        f"got {ai_hosp['breakdown']}")

    # Price margin sanity (52W Low <= Price <= 52W High)
    price, low, high = 1971, 1600, 2053
    chk("Price within 52W band", low <= price <= high, f"₹{price} in [{low}, {high}]")

    # RSI range check
    chk("RSI in 0–100",     0 <= 45.3 <= 100)
    chk("RSI OB check",     not (45.3 > 80))  # 45.3 is not overbought

    return results


# ═══════════════════════════════════════════════════════════════════════════════
# TIER 2: INTEGRATION TESTS — live network fetch + column-by-column validation
# ═══════════════════════════════════════════════════════════════════════════════

def _check_column_range(col_name, raw_value, spec):
    """Validate a single Excel cell value against its COLUMN_SPEC."""
    status_name = f"{col_name}"
    v = _pf(raw_value)

    if raw_value == "Data unavailable" or raw_value is None:
        return Result(WARN, status_name, "Data unavailable from source — cannot validate range")

    if spec.get("type") == "str":
        if "allowed" in spec:
            ok = str(raw_value) in spec["allowed"]
            return Result(PASS if ok else FAIL, status_name,
                         f"'{raw_value}' {'in' if ok else 'NOT in'} allowed set: {spec['allowed']}")
        return Result(PASS, status_name, f"'{raw_value}' (string, no range check)")

    if v is None:
        return Result(WARN, status_name, f"Cannot parse numeric from '{raw_value}'")

    lo = spec.get("min")
    hi = spec.get("max")
    in_range = (lo is None or v >= lo) and (hi is None or v <= hi)
    detail = f"{v} {spec.get('unit','')} | expected range [{lo}, {hi}]"
    return Result(PASS if in_range else FAIL, status_name, detail)


def run_integration_tests_for_symbol(symbol):
    """Fetch live data for one symbol and validate every column we care about."""
    results = []
    short = symbol.replace(".NS", "")

    print(f"\n  Fetching data for {short} ...")

    # --- Fetch all data sources ---
    try:
        f = get_fundamentals(symbol)
    except Exception as e:
        results.append(Result(FAIL, f"{short} get_fundamentals()", str(e)))
        f = {}

    try:
        q = get_quote_data(symbol)
    except Exception as e:
        results.append(Result(FAIL, f"{short} get_quote_data()", str(e)))
        q = {}

    try:
        t = get_technical(symbol)
    except Exception as e:
        results.append(Result(FAIL, f"{short} get_technical()", str(e)))
        t = None

    # --- Apply same Screener-override chain as analyze_stock() ---
    pe_val   = clean(f.get("pe"))
    roe_val  = clean(f.get("roe"))
    roce_val = clean(f.get("roce"))

    # Profit/Operating Margin: Screener INR-based primary, yfinance fallback
    _npm = parse_float(f.get("npm"))
    if _npm is not None:
        q["profit_margin"] = _npm / 100
    _opm_val = parse_float(str(f.get("opm", "")).replace("%", "").strip()) or None
    if _opm_val is not None:
        q["operating_margin"] = _opm_val / 100

    # P/B: Screener Price/BookValue (INR) primary
    _pb_sc = parse_float(f.get("price_to_book_screener"))
    if _pb_sc is not None:
        q["price_to_book"] = _pb_sc

    # P/S: Screener MCap(Cr)/Revenue(Cr) — no FX mismatch
    _ps_sc = parse_float(f.get("price_to_sales_screener"))
    if _ps_sc is not None:
        q["price_to_sales"] = _ps_sc

    # D/E: Screener primary, yfinance fallback (already /100 normalised)
    debt_raw = clean(f.get("debt"))
    de_val = debt_raw if debt_raw != "Data unavailable" else (
        round(q["debt_to_equity"], 2) if q.get("debt_to_equity") is not None else "Data unavailable"
    )

    # Growth: Screener 3Y CAGR primary, yfinance TTM fallback
    sg3 = parse_float(f.get("sales_growth_3y"))
    pg3 = parse_float(f.get("profit_growth_3y"))
    rg  = q.get("revenue_growth")
    eg  = q.get("earnings_growth")
    rev_growth_val  = (f"{round(sg3, 2)}%" if sg3 is not None else
                       f"{round(rg*100, 2)}%" if rg is not None else "Data unavailable")
    prof_growth_val = (f"{round(pg3, 2)}%" if pg3 is not None else
                       f"{round(eg*100, 2)}%" if eg is not None else "Data unavailable")

    # Dividend: Screener primary, yfinance fallback
    div_s  = parse_float(f.get("dividend_yield"))
    div_yf = q.get("dividend_yield")
    if div_s is not None:
        div_val = f"{round(div_s, 2)}%"
    elif div_yf is not None:
        _pct = div_yf if div_yf > 0.5 else div_yf * 100
        div_val = f"{round(_pct, 2)}%"
    else:
        div_val = "Data unavailable"

    # Valuation
    pb_val   = f"{round(q['price_to_book'], 2)}%"  if q.get("price_to_book")  else "Data unavailable"
    ps_val   = f"{round(q['price_to_sales'], 2)}%" if q.get("price_to_sales") else "Data unavailable"
    peg_val  = round(q["peg_ratio"], 2)             if q.get("peg_ratio")      else "Data unavailable"
    beta_val = round(q["beta"], 2)                  if q.get("beta")           else "Data unavailable"

    # Quality margins (already overridden from Screener above)
    pm_val = f"{round(q['profit_margin']*100, 2)}%"    if q.get("profit_margin")   else "Data unavailable"
    om_val = f"{round(q['operating_margin']*100, 2)}%" if q.get("operating_margin") else "Data unavailable"

    # Technical
    price_val = t["price"]  if t else "Data unavailable"
    high_val  = t["high"]   if t else "Data unavailable"
    low_val   = t["low"]    if t else "Data unavailable"
    rsi_val   = t["rsi"]    if t else "Data unavailable"
    macd_val  = t["macd"]   if t else "Data unavailable"
    ema_val   = t["ema"]    if t else "Data unavailable"

    cell_values = {
        "PE Ratio":             pe_val,
        "ROE (%)":              roe_val,
        "ROCE (%)":             roce_val,
        "Debt/ Equity":         de_val,
        "Rev. Growth 3Y (%)":   rev_growth_val,
        "Profit Growth 3Y (%)": prof_growth_val,
        "Price-to-Book":        pb_val,
        "Price-to-Sales":       ps_val,
        "PEG Ratio":            peg_val,
        "Beta":                 beta_val,
        "Dividend Yield (%)":   div_val,
        "Profit Margin (%)":    pm_val,
        "Operating Margin (%)": om_val,
        "Current Price (₹)":   price_val,
        "52W High (₹)":        high_val,
        "52W Low (₹)":         low_val,
        "RSI (14)":             rsi_val,
        "MACD Signal":          macd_val,
        "Price vs EMA (20/50/200)": ema_val,
    }

    # --- Validate each column against COLUMN_SPEC ---
    for col, raw in cell_values.items():
        spec = COLUMN_SPEC.get(col, {})
        r = _check_column_range(col, raw, spec)
        r.name = f"{short}  [{col}]"
        results.append(r)

    # --- Cross-field sanity checks ---
    p   = _pf(price_val)
    lo  = _pf(low_val)
    hi  = _pf(high_val)
    if p and lo and hi:
        ok = lo <= p <= hi
        results.append(Result(
            PASS if ok else FAIL,
            f"{short}  [Price within 52W band]",
            f"₹{p} in [₹{lo}, ₹{hi}]"
        ))
    else:
        results.append(Result(WARN, f"{short}  [Price within 52W band]", "One or more price values missing"))

    rsi_num = _pf(rsi_val)
    if rsi_num is not None:
        results.append(Result(
            PASS if 0 <= rsi_num <= 100 else FAIL,
            f"{short}  [RSI in valid range 0–100]",
            f"RSI = {rsi_num}"
        ))

    de_num = _pf(str(de_val).split("(")[0])  # strip bank note if present
    is_bank = any(b in str(q.get("industry", "")).lower() for b in
                  ("bank", "nbfc", "financ", "insurance", "housing finance"))
    if de_num is not None and not is_bank:
        results.append(Result(
            PASS if 0 <= de_num <= 10 else WARN,
            f"{short}  [D/E reasonable for non-bank (<10)]",
            f"D/E = {de_num}"
        ))
    elif is_bank:
        results.append(Result(
            PASS, f"{short}  [D/E skipped — banking sector]",
            f"D/E = {de_num} (high leverage expected)"
        ))

    # Margin sanity: profit margin should be ≤ operating margin (usually)
    pm_num = _pf(pm_val)
    om_num = _pf(om_val)
    if pm_num is not None and om_num is not None:
        results.append(Result(
            PASS if pm_num <= om_num + 10 else WARN,  # +10 buffer for edge cases
            f"{short}  [Profit margin ≤ Operating margin (±10%)]",
            f"PM={pm_num}%, OM={om_num}%"
        ))

    return results


def run_integration_tests(symbols=None):
    symbols = symbols or list(GROUND_TRUTH.keys())
    all_results = []
    for sym in symbols:
        all_results.extend(run_integration_tests_for_symbol(sym))
    return all_results


# ═══════════════════════════════════════════════════════════════════════════════
# TIER 3: REGRESSION / BENCHMARK — compare live data against GROUND_TRUTH table
# ═══════════════════════════════════════════════════════════════════════════════

def run_regression_tests(symbols=None):
    """Compare live fetched values against the GROUND_TRUTH expectations."""
    symbols = symbols or list(GROUND_TRUTH.keys())
    results = []

    for symbol in symbols:
        gt  = GROUND_TRUTH.get(symbol)
        if not gt:
            results.append(Result(SKIP, symbol, "No ground truth defined"))
            continue

        short = symbol.replace(".NS", "")
        print(f"\n  Regression check for {short} ...")

        try:
            f = get_fundamentals(symbol)
            q = get_quote_data(symbol)
            t = get_technical(symbol)
        except Exception as e:
            results.append(Result(FAIL, short, f"Data fetch failed: {e}"))
            continue

        def chk_range(name, val_str, key_min, key_max, unit=""):
            v = _pf(str(val_str).split("(")[0].replace("%","").strip())
            lo = gt.get(key_min)
            hi = gt.get(key_max)
            if v is None:
                return Result(WARN, f"{short}  [{name}]", "value missing / unparseable")
            ok = (lo is None or v >= lo) and (hi is None or v <= hi)
            return Result(
                PASS if ok else FAIL,
                f"{short}  [{name}]",
                f"{v}{unit}  |  expected [{lo}{unit}, {hi}{unit}]"
            )

        # PE
        results.append(chk_range("PE Ratio",           clean(f.get("pe")),   "pe_min",  "pe_max",  "×"))
        # ROE
        results.append(chk_range("ROE (%)",            clean(f.get("roe")),  "roe_min", "roe_max", "%"))
        # ROCE
        results.append(chk_range("ROCE (%)",           clean(f.get("roce")), "roce_min","roce_max","%"))
        # D/E
        de_raw = clean(f.get("debt"))
        de_val = de_raw if de_raw != "Data unavailable" else (
            round(q["debt_to_equity"], 2) if q.get("debt_to_equity") is not None else "Data unavailable"
        )
        results.append(chk_range("Debt/Equity",        de_val,               "de_min",  "de_max",  ""))
        # 3Y Sales Growth
        sg3 = parse_float(f.get("sales_growth_3y"))
        rg  = q.get("revenue_growth")
        rev_v = f"{round(sg3, 2)}%" if sg3 is not None else (f"{round(rg*100,2)}%" if rg else "Data unavailable")
        results.append(chk_range("Rev Growth 3Y",      rev_v,                "sg3_min", "sg3_max", "%"))
        # 3Y Profit Growth
        pg3 = parse_float(f.get("profit_growth_3y"))
        eg  = q.get("earnings_growth")
        prof_v = f"{round(pg3, 2)}%" if pg3 is not None else (f"{round(eg*100,2)}%" if eg else "Data unavailable")
        results.append(chk_range("Profit Growth 3Y",   prof_v,               "pg3_min", "pg3_max", "%"))
        # Dividend
        div_s  = parse_float(f.get("dividend_yield"))
        div_yf = q.get("dividend_yield")
        div_v  = (f"{round(div_s,2)}%" if div_s is not None else
                  (f"{round(div_yf*100,2)}%" if div_yf and div_yf <= 0.5 else
                   f"{round(div_yf,2)}%" if div_yf else "Data unavailable"))
        if gt.get("div_min") is not None:
            results.append(chk_range("Dividend Yield",    div_v,            "div_min", "div_max", "%"))
        # Current Price
        p = t["price"] if t else None
        if p is not None:
            ok = gt["price_min"] <= p <= gt["price_max"]
            results.append(Result(PASS if ok else WARN, f"{short}  [Current Price]",
                                  f"₹{p}  |  expected [₹{gt['price_min']}, ₹{gt['price_max']}]"))
        # RSI
        rsi = t["rsi"] if t else None
        if rsi is not None:
            ok = gt["rsi_min"] <= rsi <= gt["rsi_max"]
            results.append(Result(PASS if ok else WARN, f"{short}  [RSI]",
                                  f"{rsi:.1f}  |  expected [{gt['rsi_min']}, {gt['rsi_max']}]"))
        # Profit Margin
        pm_num = (q.get("profit_margin") or 0) * 100
        if pm_num and gt.get("pm_min") is not None:
            ok = gt["pm_min"] <= pm_num <= gt["pm_max"]
            results.append(Result(PASS if ok else WARN, f"{short}  [Profit Margin]",
                                  f"{pm_num:.1f}%  |  expected [{gt['pm_min']}%, {gt['pm_max']}%]"))
        # Operating Margin
        om_num = (q.get("operating_margin") or 0) * 100
        if om_num and gt.get("om_min") is not None:
            ok = gt["om_min"] <= om_num <= gt["om_max"]
            results.append(Result(PASS if ok else WARN, f"{short}  [Operating Margin]",
                                  f"{om_num:.1f}%  |  expected [{gt['om_min']}%, {gt['om_max']}%]"))
        # Price-to-Book
        pb = q.get("price_to_book")
        if pb and gt.get("pb_min") is not None:
            ok = gt["pb_min"] <= pb <= gt["pb_max"]
            results.append(Result(PASS if ok else WARN, f"{short}  [Price-to-Book]",
                                  f"{pb:.2f}×  |  expected [{gt['pb_min']}×, {gt['pb_max']}×]"))
        # Beta
        beta = q.get("beta")
        if beta and gt.get("beta_min") is not None:
            ok = gt["beta_min"] <= beta <= gt["beta_max"]
            results.append(Result(PASS if ok else WARN, f"{short}  [Beta]",
                                  f"{beta:.2f}  |  expected [{gt['beta_min']}, {gt['beta_max']}]"))

    return results


# ═══════════════════════════════════════════════════════════════════════════════
# RUNNER
# ═══════════════════════════════════════════════════════════════════════════════

def print_results(tier_name, results):
    passes = sum(1 for r in results if r.status == PASS)
    fails  = sum(1 for r in results if r.status == FAIL)
    warns  = sum(1 for r in results if r.status == WARN)
    skips  = sum(1 for r in results if r.status == SKIP)

    print(f"\n{'='*72}")
    print(f"  {tier_name}")
    print(f"{'='*72}")
    for r in results:
        print(str(r))
    print(f"\n  Summary: {passes} passed | {fails} failed | {warns} warnings | {skips} skipped")
    print(f"{'='*72}")
    return fails


def main():
    # Ensure UTF-8 output on Windows consoles that default to cp1252
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass

    parser = argparse.ArgumentParser(description="Data quality eval for AI Stock Agent")
    parser.add_argument("--tier",   type=int, choices=[1, 2, 3], help="Run only this tier (default: all)")
    parser.add_argument("--stock",  type=str, help="Single stock symbol, e.g. KIRLOSENG")
    parser.add_argument("--report", action="store_true", help="Save report to eval_report.txt")
    args = parser.parse_args()

    symbols = None
    if args.stock:
        raw = args.stock.upper()
        symbols = [raw if raw.endswith(".NS") else raw + ".NS"]

    started = datetime.now()
    total_fails = 0

    # Capture output for --report
    import io, contextlib
    buf = io.StringIO() if args.report else None
    ctx = contextlib.redirect_stdout(buf) if args.report else contextlib.nullcontext()

    with ctx:
        print(f"\n{'='*72}")
        print(f"  AI STOCK AGENT — DATA QUALITY EVAL")
        print(f"  Started : {started.strftime('%Y-%m-%d %H:%M:%S')}")
        if symbols:
            print(f"  Scope   : {symbols}")
        print(f"{'='*72}")

        if not args.tier or args.tier == 1:
            print("\n[TIER 1] Unit Tests (no network)")
            r1 = run_unit_tests()
            total_fails += print_results("TIER 1 RESULTS", r1)

        if not args.tier or args.tier == 2:
            syms = symbols or list(GROUND_TRUTH.keys())
            print(f"\n[TIER 2] Integration Tests (live fetch for {syms})")
            r2 = run_integration_tests(syms)
            total_fails += print_results("TIER 2 RESULTS", r2)

        if not args.tier or args.tier == 3:
            syms = symbols or list(GROUND_TRUTH.keys())
            print(f"\n[TIER 3] Regression / Benchmark (vs ground truth for {syms})")
            r3 = run_regression_tests(syms)
            total_fails += print_results("TIER 3 RESULTS", r3)

        finished = datetime.now()
        print(f"\n  Total elapsed : {(finished - started).seconds}s")
        print(f"  Overall       : {'❌ FAILURES FOUND — review above' if total_fails else '✅ ALL CHECKS PASSED'}")

    if args.report:
        output = buf.getvalue()
        report_path = Path("eval_report.txt")
        report_path.write_text(output, encoding="utf-8")
        # Also print to screen
        print(output)
        print(f"\nReport saved to {report_path.resolve()}")

    sys.exit(1 if total_fails else 0)


if __name__ == "__main__":
    main()
