import yfinance as yf
import pandas as pd
import requests
import time
import warnings
from bs4 import BeautifulSoup
from ta.trend import EMAIndicator, MACD
from ta.momentum import RSIIndicator

warnings.filterwarnings("ignore")

# =============================
# SYMBOL RESOLUTION (NS → BO → SKIP)
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
# UTILS
# =============================

def clean_value(val):
    if val in ["", None, " ", "-", "%", "P/E"]:
        return "Data unavailable"
    return val


# =============================
# COMPANY INFO
# =============================

def get_company_info(symbol):
    try:
        info = yf.Ticker(symbol).info
        name = info.get("longName", symbol.replace(".NS",""))
        sector = info.get("sector", "Data unavailable")
        return name, sector
    except:
        return symbol.replace(".NS",""), "Data unavailable"


# =============================
# TECHNICAL DATA (FIXED)
# =============================

def get_technical(symbol):

    resolved = resolve_symbol(symbol)

    if not resolved:
        print(f"[SKIP] No symbol found → {symbol}")
        return None

    df = yf.download(resolved, period="1y", interval="1d", progress=False)

    if df.empty or len(df) < 100:
        print(f"[SKIP] Not enough data → {resolved}")
        return None

    if isinstance(df.columns, pd.MultiIndex):
        df.columns = df.columns.get_level_values(0)

    close = df["Close"].squeeze()

    df["EMA200"] = EMAIndicator(close, 200).ema_indicator()
    df["RSI"] = RSIIndicator(close).rsi()
    df["MACD"] = MACD(close).macd_diff()

    df.dropna(inplace=True)

    if df.empty:
        return None

    last = df.iloc[-1]

    price = float(last["Close"])
    rsi_val = float(last["RSI"])
    macd_val = float(last["MACD"])
    ema200 = float(last["EMA200"])

    return {
        "Price": round(price, 2),
        "RSI": round(rsi_val, 2),
        "MACD": "Bullish" if macd_val > 0 else "Bearish",
        "EMA": "Above" if price > ema200 else "Below"
    }


# =============================
# FUNDAMENTALS (FIXED)
# =============================

def get_fundamentals(symbol):

    try:
        sym = symbol.replace(".NS", "")
        url = f"https://www.screener.in/company/{sym}/"

        r = requests.get(url, headers={"User-Agent":"Mozilla/5.0"})
        soup = BeautifulSoup(r.text, "html.parser")

        data = {}

        for item in soup.select("li.flex.flex-space-between"):
            name = item.find("span", class_="name")
            val = item.find("span", class_="number")

            if not name or not val:
                continue

            name = name.text.strip()
            val = val.text.strip()

            if "P/E" in name:
                data["PE"] = val
            elif "ROE" in name:
                data["ROE"] = val
            elif "ROCE" in name:
                data["ROCE"] = val
            elif "Debt" in name or "Debt to equity" in name:
                data["Debt"] = val

        return data

    except:
        return {}


# =============================
# INVESTOR DATA (FIXED)
# =============================

def get_investor_data(symbol):

    try:
        base = symbol.replace(".NS","")
        url = f"https://www.nseindia.com/api/equity-shareholding-pattern?symbol={base}"

        session = requests.Session()
        headers = {"User-Agent": "Mozilla/5.0"}

        session.get("https://www.nseindia.com", headers=headers)
        time.sleep(1)

        r = session.get(url, headers=headers)

        if r.status_code != 200:
            return {}

        data = r.json().get("data", [])

        if not data:
            return {}

        latest = data[0]
        prev = data[1] if len(data) > 1 else {}

        mf_latest = float(latest.get("mutualFunds", latest.get("mutualFund", 0)))
        mf_prev = float(prev.get("mutualFunds", prev.get("mutualFund", 0)))

        fii_latest = float(latest.get("foreignInstitutions",
                                     latest.get("foreignInstitutionalInvestors", 0)))
        fii_prev = float(prev.get("foreignInstitutions",
                                 prev.get("foreignInstitutionalInvestors", 0)))

        return {
            "MF": mf_latest,
            "FII": fii_latest,
            "MF Trend": "Added" if mf_latest > mf_prev else "Reduced",
            "FII Trend": "Added" if fii_latest > fii_prev else "Reduced"
        }

    except:
        return {}


# =============================
# SCORING
# =============================

def calculate_score(tech, funda, investor):

    score = 0

    if tech["MACD"] == "Bullish":
        score += 10

    if tech["RSI"] > 40:
        score += min((tech["RSI"] - 40), 10)

    if tech["EMA"] == "Above":
        score += 10

    if funda.get("ROE"):
        score += 10

    if funda.get("ROCE"):
        score += 10

    if investor.get("MF Trend") == "Added":
        score += 10

    return round(score)


# =============================
# ANALYZER
# =============================

def analyze_stock(symbol):

    tech = get_technical(symbol)

    if tech is None:
        return None

    funda = get_fundamentals(symbol)
    investor = get_investor_data(symbol)

    name, sector = get_company_info(symbol)

    mf = investor.get("MF")
    fii = investor.get("FII")

    if mf or fii:
        holding = f"MF: {mf}%, FII: {fii}%"
    else:
        holding = "Data unavailable"

    debt = funda.get("Debt", "Data unavailable")

    return {
        "Company": symbol.replace(".NS",""),
        "Company Full Name": name,
        "Sector": sector,
        "Holding Pattern": holding,
        "PE": clean_value(funda.get("PE")),
        "ROE": clean_value(funda.get("ROE")),
        "ROCE": clean_value(funda.get("ROCE")),
        "Debt": debt,
        "Price": tech["Price"],
        "RSI": tech["RSI"],
        "Score": calculate_score(tech, funda, investor)
    }


# =============================
# RUN
# =============================

stocks = [
"CYIENT.NS","BLUESTARCO.NS","WHIRLPOOL.NS","KALPATARU.NS",
"JUBLINGREA.NS","NH.NS","KIRLOSENG.NS","EQUITASBNK.NS",
"NAVINFLUOR.NS","MCX.NS","BRIGADE.NS","IPCALAB.NS",
"PATELENG.NS","ADVAIT.NS","SHANKARA.NS","SDBL.NS",
"TGVSI.NS","ANUP.NS","FAIRCHEMOR.NS","NAVA.NS",
"SALZERELEC.NS","AMAL.NS","HITECH.NS","BIGBLOC.NS",
"KMCSHIL.NS","BMW.NS","SUNDARMFIN.NS","SRF.NS","KIMS.NS"
]

results = []
skipped = []

for s in stocks:
    print(f"Processing {s}")

    try:
        r = analyze_stock(s)
        if r:
            results.append(r)
        else:
            skipped.append(s)
    except Exception as e:
        print(f"[ERROR] {s}: {e}")
        skipped.append(s)

df = pd.DataFrame(results)

df = df.sort_values(by=["Score","RSI"], ascending=[False, False])
df["Rank"] = range(1, len(df) + 1)

df.to_excel("Stock_Agent_Final_Output.xlsx", index=False)

print("\n✅ DONE")
print("Skipped:", skipped)
print(df)