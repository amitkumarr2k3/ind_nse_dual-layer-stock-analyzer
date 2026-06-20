import yfinance as yf
import pandas as pd
import requests
import time
from bs4 import BeautifulSoup
from ta.trend import EMAIndicator, MACD
from ta.momentum import RSIIndicator

# =============================
# ✅ UTILS
# =============================

def clean_value(val):
    if val in ["", None, " ", "-", "%", "P/E"]:
        return "Data unavailable"
    return val

def parse_market_cap(mcap_str):
    try:
        if not mcap_str:
            return None
        mcap_str = mcap_str.replace("₹", "").replace(",", "").replace("Cr.", "").strip()
        return float(mcap_str)
    except:
        return None

def classify_market_cap(mcap):
    if mcap is None:
        return "Data unavailable"
    if mcap > 50000:
        return "Large Cap"
    elif mcap > 5000:
        return "Mid Cap"
    else:
        return "Small Cap"

# =============================
# ✅ COMPANY INFO (DYNAMIC)
# =============================

def get_company_info(symbol):
    try:
        ticker = yf.Ticker(symbol)
        info = ticker.info

        name = info.get("longName", symbol)
        sector = info.get("sector", "Data unavailable")

        return name, sector
    except:
        return symbol.replace(".NS",""), "Data unavailable"

# =============================
# ✅ TECHNICAL DATA
# =============================

def get_technical(symbol):
    try:
        df = yf.download(symbol, period="1y", interval="1d")

        if df.empty:
            return None

        if isinstance(df.columns, pd.MultiIndex):
            df.columns = df.columns.get_level_values(0)

        close = df['Close'].squeeze()

        df['EMA20'] = EMAIndicator(close, 20).ema_indicator()
        df['EMA50'] = EMAIndicator(close, 50).ema_indicator()
        df['EMA200'] = EMAIndicator(close, 200).ema_indicator()

        df['RSI'] = RSIIndicator(close).rsi()
        df['MACD'] = MACD(close).macd_diff()

        df.dropna(inplace=True)
        if df.empty:
            return None

        latest = df.iloc[-1]

        price = float(latest['Close'])
        ema20 = float(latest['EMA20'])
        ema50 = float(latest['EMA50'])
        ema200 = float(latest['EMA200'])

        ema_status = ", ".join([
            "Above 20" if price > ema20 else "Below 20",
            "Above 50" if price > ema50 else "Below 50",
            "Above 200" if price > ema200 else "Below 200"
        ])

        return {
            "Price": round(price, 2),
            "RSI": round(float(latest['RSI']), 2),
            "MACD": "Bullish" if latest['MACD'] > 0 else "Bearish",
            "EMA": ema_status,
            "52W_High": round(df['Close'].max(),2),
            "52W_Low": round(df['Close'].min(),2)
        }

    except:
        return None

# =============================
# ✅ FUNDAMENTALS (SCREENER)
# =============================

def get_fundamentals(symbol):
    try:
        sym = symbol.replace(".NS","")
        url = f"https://www.screener.in/company/{sym}/"
        headers = {"User-Agent": "Mozilla/5.0"}

        r = requests.get(url, headers=headers)
        soup = BeautifulSoup(r.text, "html.parser")

        data = {}

        for item in soup.select("li.flex.flex-space-between"):
            label = item.find("span", class_="name")
            value = item.find("span", class_="number")

            if not label or not value:
                continue

            label = label.text.strip()
            value = value.text.strip()

            if "Market Cap" in label:
                data["MCap"] = value
            elif "P/E" in label:
                data["PE"] = value
            elif "ROE" in label:
                data["ROE"] = value
            elif "ROCE" in label:
                data["ROCE"] = value
            elif "Debt" in label:
                data["Debt"] = value

        return data
    except:
        return {}

# =============================
# ✅ NSE INVESTOR DATA
# =============================

def get_nse_session():
    session = requests.Session()
    headers = {"User-Agent": "Mozilla/5.0"}
    session.get("https://www.nseindia.com", headers=headers)
    return session

def detect_change(prev, curr):
    if curr > prev:
        return "Added"
    elif curr < prev:
        return "Reduced"
    else:
        return "Hold"

def get_nse_shareholding(symbol):
    try:
        base = symbol.replace(".NS","")

        url = f"https://www.nseindia.com/api/equity-shareholding-pattern?symbol={base}"

        session = get_nse_session()
        time.sleep(1)

        r = session.get(url, headers={"User-Agent": "Mozilla/5.0"})
        data = r.json().get("data", [])

        if len(data) < 2:
            return {}

        latest = data[0]
        prev = data[1]

        def safe(obj, keys):
            for k in keys:
                if k in obj:
                    return float(obj[k])
            return 0

        mf_latest = safe(latest, ["mutualFunds","mutualFund"])
        mf_prev = safe(prev, ["mutualFunds","mutualFund"])

        fii_latest = safe(latest, ["foreignInstitutions","fii"])
        fii_prev = safe(prev, ["foreignInstitutions","fii"])

        return {
            "MF": mf_latest,
            "FII": fii_latest,
            "MF Trend": detect_change(mf_prev, mf_latest),
            "FII Trend": detect_change(fii_prev, fii_latest)
        }

    except:
        return {}

# =============================
# ✅ SCORING ENGINE (IMPROVED)
# =============================

def calculate_score(tech, funda, investor):
    score = 0

    # Technical (30)
    score += 10 if tech["MACD"] == "Bullish" else 0
    score += max(min((tech["RSI"] - 40), 10),0)
    score += 10 if "Above 200" in tech["EMA"] else 0

    # Fundamentals (25)
    score += 5 if funda.get("PE") else 0
    score += 10 if funda.get("ROE") else 0
    score += 10 if funda.get("ROCE") else 0

    # Investor (20)
    score += 10 if investor.get("MF Trend") == "Added" else 0
    score += 10 if investor.get("FII Trend") == "Added" else 0

    # Risk (15)
    if funda.get("Debt"):
        score += 15

    # Momentum (10)
    if tech["RSI"] > 60:
        score += 10

    return round(score,0)

def final_action(score):
    if score >= 75:
        return "ENTER"
    elif score >= 65:
        return "ADD"
    elif score >= 55:
        return "HOLD"
    elif score >= 45:
        return "WATCHLIST"
    else:
        return "EXIT"

# =============================
# ✅ MAIN ANALYZER
# =============================

def analyze_stock(symbol):

    tech = get_technical(symbol)
    if tech is None:
        return None

    funda = get_fundamentals(symbol)
    investor = get_nse_shareholding(symbol)

    name, sector = get_company_info(symbol)

    mcap = parse_market_cap(funda.get("MCap"))
    mcap_cat = classify_market_cap(mcap)

    score = calculate_score(tech, funda, investor)
    action = final_action(score)

    mf = investor.get("MF","NA")
    fii = investor.get("FII","NA")

    investor_summary = "Data unavailable"
    if mf != 0:
        investor_summary = f"MF: {mf}% ({investor.get('MF Trend')}), FII: {fii}% ({investor.get('FII Trend')})"

    return {
        "Company": symbol.replace(".NS",""),
        "Company Full Name": name,
        "Sector": sector,
        "Market Cap Category": mcap_cat,
        "Holding Pattern": investor_summary,

        "PE": clean_value(funda.get("PE")),
        "ROE": clean_value(funda.get("ROE")),
        "ROCE": clean_value(funda.get("ROCE")),
        "Debt": clean_value(funda.get("Debt")),

        "Price": tech["Price"],
        "RSI": tech["RSI"],
        "MACD": tech["MACD"],
        "EMA": tech["EMA"],

        "Score": score,
        "Action": action
    }

# =============================
# ✅ RUN
# =============================

stocks = ["CYIENT.NS","BLUESTARCO.NS","WHIRLPOOL.NS","KALPATARU.NS","JUBLINGREA.NS","NH.NS","KIRLOSENG.NS","EQUITASBNK.NS","NAVINFLUOR.NS","MCX.NS","BRIGADE.NS","IPCALAB.NS","PATELENG.NS","ADVAIT.NS","SHANKARA.NS","SDBL.NS","ACE.NS","MAHSEAMLES.NS","NRBBEARING.NS","TIMETECHNO.NS","FIEMIND.NS","LAOPALA.NS","SHARDAMOTR.NS","TDPOWERSYS.NS","TGVSI.NS","ANUP.NS","FAIRCHEMOR.NS","NAVA.NS","SALZERELEC.NS","AMAL.NS","HITECH.NS","BIGBLOC.NS","KMCSHIL.NS","BMW.NS","SUNDARMFIN.NS","SRF.NS","KIMS.NS"]
results = []

for s in stocks:
    print(f"Processing {s}")
    r = analyze_stock(s)
    if r:
        results.append(r)

df = pd.DataFrame(results)
df.fillna("Data unavailable", inplace=True)

# ✅ Proper ranking (NO duplicates)
df = df.sort_values(by=["Score","RSI"], ascending=[False,False])
df["Rank"] = range(1,len(df)+1)

# ✅ Top 5 Filter
df["Top Pick"] = df.apply(
    lambda x: "✅ Top 5" if x["Rank"] <= 5 and x["Score"] >= 60 else "",
    axis=1
)

df.to_excel("Final_Pro_Stock_Output.xlsx", index=False)

print("\n✅ Final Output Generated")
print(df)