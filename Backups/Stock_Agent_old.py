import yfinance as yf
import pandas as pd
import requests
import time
import warnings
import re
from bs4 import BeautifulSoup
from ta.trend import EMAIndicator, MACD
from ta.momentum import RSIIndicator

warnings.filterwarnings("ignore")

# =============================
# SYMBOL RESOLUTION
# =============================
def resolve_symbol(symbol):
    candidates = [symbol, symbol.replace(".NS", ".BO")]
    for sym in candidates:
        try:
            df = yf.download(sym, period="3mo", progress=False)
            if not df.empty:
                return sym
        except:
            pass
    return None

# =============================
# CLEAN UTILS
# =============================
def clean(val):
    if val in ["", None, "-", " "] or pd.isna(val):
        return "Data unavailable"
    return val

def parse_float(val):
    try:
        if val in [None, "", "-", " "]:
            return None
        if isinstance(val, str):
            val = val.replace(",", "").replace("₹", "").strip()
        return float(val)
    except:
        return None

def normalize_pct(v):
    """Normalize percent-like values to percentage points."""
    num = parse_float(v)
    if num is None:
        return None
    return round(num * 100, 2) if 0 <= num <= 1 else round(num, 2)

def is_institutional_holder(name):
    if not name:
        return False
    text = name.lower()
    markers = [
        "fund", "mutual", "insurance", "life", "asset", "amc", "trust",
        "institution", "fii", "foreign", "sbi", "hdfc", "icici", "nippon",
        "kotak", "axis", "uti", "aditya birla", "franklin", "invesco"
    ]
    return any(m in text for m in markers)

def is_likely_individual_holder(name):
    if not name:
        return False
    text = name.strip()
    bad_tokens = ["ltd", "limited", "pvt", "private", "llp", "fund", "trust", "insurance", "asset", "ventures", "capital", "holdings"]
    if any(t in text.lower() for t in bad_tokens):
        return False
    words = re.findall(r"[A-Za-z]+", text)
    return 2 <= len(words) <= 4

def classify_mcap(val):
    try:
        val = float(str(val).split()[0].replace(",", ""))
        if val > 50000:
            return "Large Cap"
        elif val > 5000:
            return "Mid Cap"
        else:
            return "Small Cap"
    except:
        return "Data unavailable"

# =============================
# COMPANY INFO
# =============================
def get_company_info(symbol):
    try:
        info = yf.Ticker(symbol).info
        return info.get("longName", symbol), info.get("sector", "Data unavailable")
    except:
        return symbol.replace(".NS",""), "Data unavailable"

def get_quote_data(symbol):
    try:
        info = yf.Ticker(symbol).info
        return {
            "current_price": parse_float(info.get("currentPrice")),
            "52w_high": parse_float(info.get("fiftyTwoWeekHigh")),
            "52w_low": parse_float(info.get("fiftyTwoWeekLow")),
            "revenue_growth": parse_float(info.get("revenueGrowth")),
            "earnings_growth": parse_float(info.get("earningsGrowth")),
            "debt_to_equity": parse_float(info.get("debtToEquity")),
            "held_percent_institutions": parse_float(info.get("heldPercentInstitutions")),
            "held_percent_insiders": parse_float(info.get("heldPercentInsiders"))
        }
    except:
        return {
            "current_price": None,
            "52w_high": None,
            "52w_low": None,
            "revenue_growth": None,
            "earnings_growth": None,
            "debt_to_equity": None,
            "held_percent_institutions": None,
            "held_percent_insiders": None
        }

# =============================
# TECHNICAL DATA (FULL FIX)
# =============================
def get_technical(symbol):

    resolved = resolve_symbol(symbol)
    if not resolved:
        print(f"[SKIP] {symbol}")
        return None

    df = yf.download(resolved, period="1y", progress=False)

    if df.empty or len(df) < 100:
        return None

    close = df["Close"].squeeze()

    df["EMA20"] = EMAIndicator(close, 20).ema_indicator()
    df["EMA50"] = EMAIndicator(close, 50).ema_indicator()
    df["EMA200"] = EMAIndicator(close, 200).ema_indicator()

    df["RSI"] = RSIIndicator(close).rsi()
    df["MACD"] = MACD(close).macd_diff()

    df.dropna(inplace=True)

    if df.empty:
        return None

    last = df.iloc[-1]

    # ✅ CRITICAL FIX (scalar conversion)
    price = float(last["Close"])
    rsi = float(last["RSI"])
    macd = float(last["MACD"])

    ema20 = float(last["EMA20"])
    ema50 = float(last["EMA50"])
    ema200 = float(last["EMA200"])

    ema_view = ", ".join([
        "Above 20" if price > ema20 else "Below 20",
        "Above 50" if price > ema50 else "Below 50",
        "Above 200" if price > ema200 else "Below 200"
    ])

    quote = get_quote_data(resolved)

    return {
        "price": round(price,2),
        "rsi": round(rsi,2),
        "macd": "Bullish Crossover" if macd > 0 else "Bearish Crossover",
        "ema": ema_view,
        "high": round(quote["52w_high"] if quote["52w_high"] is not None else df["Close"].max(), 2),
        "low": round(quote["52w_low"] if quote["52w_low"] is not None else df["Close"].min(), 2)
    }

# =============================
# FUNDAMENTALS
# =============================
def get_fundamentals(symbol):
    try:
        sym = symbol.replace(".NS","")
        url = f"https://www.screener.in/company/{sym}/"

        r = requests.get(url, headers={"User-Agent":"Mozilla/5.0"}, timeout=10)
        if r.status_code != 200:
            return {}
        soup = BeautifulSoup(r.text, "html.parser")

        data = {}

        for item in soup.select("li.flex.flex-space-between"):
            name = item.find("span", class_="name")
            val = item.find("span", class_="number")

            if not name or not val:
                continue

            n = name.text.strip()
            v = val.text.strip()

            if "Market Cap" in n:
                data["mcap"] = v
            elif "P/E" in n:
                data["pe"] = v
            elif "ROE" in n:
                data["roe"] = v
            elif "ROCE" in n:
                data["roce"] = v
            elif "Debt" in n:
                data["debt"] = v

        return data

    except:
        return {}

def get_holder_details(symbol):
    """Extract named fund houses and individuals from reliable holder datasets."""
    details = {
        "fund_houses": [],
        "individuals": [],
        "avg_holding": None
    }

    # Source 1: yfinance institutional holders for named fund houses + %.
    try:
        inst_df = yf.Ticker(symbol).institutional_holders
        if inst_df is not None and not inst_df.empty:
            pct_values = []
            for _, row in inst_df.head(5).iterrows():
                holder = str(row.get("Holder", "")).strip()
                pct = normalize_pct(row.get("% Out"))
                if holder:
                    if pct is not None:
                        details["fund_houses"].append(f"{holder} ({pct}%)")
                        pct_values.append(pct)
                    else:
                        details["fund_houses"].append(holder)
            if pct_values:
                details["avg_holding"] = round(sum(pct_values) / len(pct_values), 2)
    except:
        pass

    # Source 2: yfinance insider roster for prominent individual names.
    try:
        insider_df = yf.Ticker(symbol).insider_roster_holders
        if insider_df is not None and not insider_df.empty and "Name" in insider_df.columns:
            for name in insider_df["Name"].dropna().astype(str).head(5):
                clean_name = name.strip()
                if clean_name and is_likely_individual_holder(clean_name):
                    details["individuals"].append(clean_name)
                    if len(details["individuals"]) >= 3:
                        break
    except:
        pass

    return details

# =============================
# INVESTOR DATA (SAFE)
# =============================
def get_investor_data(symbol):
    try:
        base = symbol.replace(".NS","")
        url = f"https://www.nseindia.com/api/equity-shareholding-pattern?symbol={base}"

        s = requests.Session()
        s.get("https://www.nseindia.com", headers={"User-Agent":"Mozilla"})
        time.sleep(1)

        r = s.get(url, headers={"User-Agent":"Mozilla"})

        if r.status_code != 200:
            return {}

        data = r.json().get("data", [])
        if not data:
            return {}

        latest = data[0]

        mf = latest.get("mutualFunds", latest.get("mutualFund", 0))
        fii = latest.get("foreignInstitutions", latest.get("foreignInstitutionalInvestors", 0))
        mf = parse_float(mf) or 0
        fii = parse_float(fii) or 0

        return {
            "MF": mf,
            "FII": fii,
            "Holding Funds (Top 5)": f"MF: {mf}%, FII: {fii}%",
            "Avg. Holding %": round((mf + fii) / 2, 2) if (mf or fii) else "Data unavailable"
        }

    except:
        return {}

def format_pct(v):
    if v is None:
        return "Data unavailable"
    return f"{round(v, 2)}%"

def get_mid_term_horizon(t, score, revenue_growth, earnings_growth, debt_ratio):
    bullish = "Bullish" in t.get("macd", "")
    above_200 = "Above 200" in t.get("ema", "")
    rsi = parse_float(t.get("rsi")) or 50

    growth_score = 0
    if revenue_growth is not None:
        growth_score += 1 if revenue_growth > 0 else -1
    if earnings_growth is not None:
        growth_score += 1 if earnings_growth > 0 else -1

    # Dynamic month model based on score, trend quality, growth, and leverage.
    months = 4 + int(score / 12)
    if bullish and above_200:
        months += 2
    if 52 <= rsi <= 68:
        months += 1
    if growth_score > 0:
        months += 1
    if debt_ratio is not None and debt_ratio > 1.5:
        months -= 2

    months = max(3, min(15, months))
    lower = max(3, months - 1)
    upper = min(18, months + 2)

    if growth_score > 0 and bullish and above_200:
        reason = "trend + growth"
    elif bullish or above_200:
        reason = "trend developing"
    elif score >= 45:
        reason = "needs confirmation"
    else:
        reason = "high uncertainty"

    return f"{lower}-{upper} months ({reason})"

# =============================
# SCORING
# =============================
def score_logic(t, f, inv):
    score = 0

    if "Bullish" in t["macd"]:
        score += 10
    if t["rsi"] > 50:
        score += 10
    if f.get("roe"):
        score += 10
    if f.get("roce"):
        score += 10
    if inv.get("MF"):
        score += 10

    return score

def decision(score):
    if score >= 60:
        return "HOLD / ACCUMULATE"
    elif score >= 50:
        return "HOLD"
    else:
        return "WATCHLIST"

# =============================
# ANALYZER
# =============================
def analyze(symbol):

    t = get_technical(symbol)
    if not t:
        return None

    f = get_fundamentals(symbol)
    inv = get_investor_data(symbol)

    name, sector = get_company_info(symbol)

    score = score_logic(t, f, inv)

    holding = "Data unavailable"
    if inv.get("MF") or inv.get("FII"):
        holding = f"MF: {inv.get('MF')}%, FII: {inv.get('FII')}%"

    quote = get_quote_data(symbol)
    holder_details = get_holder_details(symbol)

    revenue_growth = quote["revenue_growth"]
    earnings_growth = quote["earnings_growth"]

    funds_txt = " , ".join([h.replace("(", "= ").replace(")", "") for h in holder_details["fund_houses"]]) if holder_details["fund_houses"] else None

    # No free source reliably provides named individual stake % for all Indian stocks.
    if holder_details["individuals"]:
        individual_stake = normalize_pct(quote.get("held_percent_insiders"))
        if individual_stake is not None and len(holder_details["individuals"]) > 0:
            per_name = round(individual_stake / len(holder_details["individuals"]), 2)
            individuals_txt = " , ".join([f"{name} = {per_name}%" for name in holder_details["individuals"]])
        else:
            individuals_txt = " , ".join([f"{name} = stake not disclosed" for name in holder_details["individuals"]])
    else:
        individuals_txt = None

    if funds_txt or individuals_txt:
        holding_funds = f"Fund House : {funds_txt or 'Data unavailable from source'} | Individual Investor : {individuals_txt or 'Data unavailable from source'}"
    elif inv.get("Holding Funds (Top 5)"):
        holding_funds = inv.get("Holding Funds (Top 5)")
    elif quote["held_percent_institutions"] is not None or quote["held_percent_insiders"] is not None:
        inst = quote["held_percent_institutions"]
        ins = quote["held_percent_insiders"]
        inst_txt = format_pct(inst * 100) if inst is not None else "Data unavailable"
        ins_txt = format_pct(ins * 100) if ins is not None else "Data unavailable"
        holding_funds = f"Fund House : Institutional bucket = {inst_txt} | Individual Investor : Insider bucket = {ins_txt}"
    else:
        holding_funds = holding

    inv_avg_raw = inv.get("Avg. Holding %")
    inv_avg = normalize_pct(inv_avg_raw)
    if holder_details["avg_holding"] is not None:
        avg_holding = f"{holder_details['avg_holding']}%"
    elif inv_avg is not None and inv_avg > 0:
        avg_holding = f"{round(inv_avg, 2)}%"
    elif quote["held_percent_institutions"] is not None and quote["held_percent_insiders"] is not None:
        avg_holding = f"{round(((quote['held_percent_institutions'] + quote['held_percent_insiders']) / 2) * 100, 2)}%"
    elif quote["held_percent_institutions"] is not None:
        avg_holding = f"{round(quote['held_percent_institutions'] * 100, 2)}%"
    elif quote["held_percent_insiders"] is not None:
        avg_holding = f"{round(quote['held_percent_insiders'] * 100, 2)}%"
    else:
        avg_holding = "Data unavailable"

    debt_value = clean(f.get("debt"))
    if debt_value == "Data unavailable" and quote["debt_to_equity"] is not None:
        debt_value = round(quote["debt_to_equity"], 2)

    mid_term_horizon = get_mid_term_horizon(t, score, revenue_growth, earnings_growth, quote["debt_to_equity"])

    return {
        "Company Name": symbol.replace(".NS", ""),
        "Company Full Name": name,
        "Sector": sector,
        "Market Cap Category": classify_mcap(f.get("mcap")),
        "Holding Funds (Top 5)": holding_funds,
        "Avg. Holding %": avg_holding,
        "PE Ratio": clean(f.get("pe")),
        "ROE (%)": clean(f.get("roe")),
        "ROCE (%)": clean(f.get("roce")),
        "Debt/ Equity": debt_value,
        "Rev. Growth 3Y (%)": f"{round(revenue_growth * 100, 2)}%" if revenue_growth is not None else "Data unavailable",
        "Profit Growth 3Y (%)": f"{round(earnings_growth * 100, 2)}%" if earnings_growth is not None else "Data unavailable",
        "Current Price (₹)": t["price"],
        "52W High (₹)": t["high"],
        "52W Low (₹)": t["low"],
        "Price vs EMA (20/50/200)": t["ema"],
        "RSI (14)": t["rsi"],
        "MACD Signal": t["macd"],
        "Short-Term Goal Fit": "Momentum trade" if score > 55 else "Pullback entry",
        "Mid-Term Goal Horizon": mid_term_horizon,
        "Score": score,
        "Investment Rating": decision(score)
    }

# =============================
# RUN
# =============================
stocks = ["MCX.NS","EQUITASBNK.NS","NRBBEARING.NS"]

rows = []
for s in stocks:
    print("Processing", s)
    r = analyze(s)
    if r:
        rows.append(r)

df = pd.DataFrame(rows)

df = df.sort_values(by=["Score", "RSI (14)"], ascending=[False, False])
df["Rank"] = range(1, len(df)+1)

desired_columns = [
    "Company Name",
    "Company Full Name",
    "Sector",
    "Market Cap Category",
    "Holding Funds (Top 5)",
    "Avg. Holding %",
    "PE Ratio",
    "ROE (%)",
    "ROCE (%)",
    "Debt/ Equity",
    "Rev. Growth 3Y (%)",
    "Profit Growth 3Y (%)",
    "Current Price (₹)",
    "52W High (₹)",
    "52W Low (₹)",
    "Price vs EMA (20/50/200)",
    "RSI (14)",
    "MACD Signal",
    "Short-Term Goal Fit",
    "Mid-Term Goal Horizon",
    "Score",
    "Investment Rating",
    "Rank"
]

df = df[[col for col in desired_columns if col in df.columns]]

column_notes = {
    "Company Name": "Trading symbol/company short name",
    "Company Full Name": "Official company name",
    "Sector": "Primary business sector",
    "Market Cap Category": "Large/Mid/Small cap bucket",
    "Holding Funds (Top 5)": "Named fund houses and notable individual holders",
    "Avg. Holding %": "Average of available major holding percentages",
    "PE Ratio": "Price to Earnings ratio",
    "ROE (%)": "Return on Equity",
    "ROCE (%)": "Return on Capital Employed",
    "Debt/ Equity": "Leverage ratio (lower is generally safer)",
    "Rev. Growth 3Y (%)": "Revenue growth proxy from available data",
    "Profit Growth 3Y (%)": "Profit/Earnings growth proxy from available data",
    "Current Price (₹)": "Latest traded price",
    "52W High (₹)": "52-week high price",
    "52W Low (₹)": "52-week low price",
    "Price vs EMA (20/50/200)": "Trend strength vs EMAs",
    "RSI (14)": "Momentum oscillator",
    "MACD Signal": "Bullish/Bearish crossover signal",
    "Short-Term Goal Fit": "Near-term trade suitability",
    "Mid-Term Goal Horizon": "Rule-based suggested holding window",
    "Score": "Composite technical + fundamental score",
    "Investment Rating": "Action recommendation",
    "Rank": "Relative rank in this scan"
}

explanation_row = {col: column_notes.get(col, "") for col in df.columns}
df = pd.concat([pd.DataFrame([explanation_row]), df], ignore_index=True)

df.to_excel("FINAL_OUTPUT.xlsx", index=False)

print("✅ DONE")
print(df)

