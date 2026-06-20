import yfinance as yf
import pandas as pd
from ta.trend import EMAIndicator, MACD
from ta.momentum import RSIIndicator

# =============================
# FUNDAMENTAL DATA
# =============================

import requests
from bs4 import BeautifulSoup

def get_fundamentals(symbol):
    try:
        screener_symbol = symbol.replace(".NS", "")
        url = f"https://www.screener.in/company/{screener_symbol}/"
        headers = {"User-Agent": "Mozilla/5.0"}

        r = requests.get(url, headers=headers)
        soup = BeautifulSoup(r.text, "html.parser")

        data = {}

        # =============================
        # ✅ 1. TOP RATIOS (PE, ROE, etc.)
        # =============================
        ratios = soup.select("li.flex.flex-space-between")

        for item in ratios:
            text = item.text.strip()

            if "Market Cap" in text:
                data["Market Cap"] = text.split()[-2]

            elif "P/E" in text:
                data["PE"] = text.split()[-1]

            elif "ROE" in text:
                data["ROE"] = text.split()[-1]

            elif "ROCE" in text:
                data["ROCE"] = text.split()[-1]

            elif "Debt" in text:
                data["Debt/Equity"] = text.split()[-1]

        # =============================
        # ✅ 2. RATIOS TABLE (ADVANCED)
        # =============================
        tables = soup.find_all("table")

        for table in tables:
            if "Ratios" in table.text:

                rows = table.find_all("tr")

                for row in rows:
                    cols = [c.text.strip() for c in row.find_all("td")]

                    if len(cols) >= 2:
                        if "Sales growth" in cols[0]:
                            data["Rev Growth 3Y"] = cols[1]

                        elif "Profit growth" in cols[0]:
                            data["Profit Growth 3Y"] = cols[1]

        # =============================
        # ✅ 3. SHAREHOLDING (OPTIONAL PREP)
        # =============================
        # (We will expand this later for investor tracking)

        return data

    except Exception as e:
        print(f"[ERROR] Fundamentals failed for {symbol}: {e}")
        return {}

def get_investor_data(symbol):
    try:
        screener_symbol = symbol.replace(".NS","")
        url = f"https://www.screener.in/company/{screener_symbol}/consolidated/"
        headers = {"User-Agent": "Mozilla/5.0"}

        r = requests.get(url, headers=headers)
        soup = BeautifulSoup(r.text, "html.parser")

        # Placeholder (Screener HTML changes frequently)
        funds = 3   # fallback example
        avg_holding = "2.5%"

        trend = "Data unavailable"

        return {
            "Holding Funds": funds,
            "Avg Holding %": avg_holding,
            "Investor Trend": trend
        }

    except:
        return {
            "Holding Funds": "NA",
            "Avg Holding %": "NA",
            "Investor Trend": "NA"
        }

def get_growth_data(symbol):
    try:
        # Placeholder values (replace with real financial source later)
        return {
            "Rev Growth 3Y": "20%",
            "Profit Growth 3Y": "30%"
        }
    except:
        return {
            "Rev Growth 3Y": "NA",
            "Profit Growth 3Y": "NA"
        }
    
def generate_decision(tech):

    rsi = tech["RSI"]
    macd = tech["MACD"]

    if rsi < 45 and macd == "Bearish":
        return "Pullback entry", "20-30% rebound", "HOLD / ACCUMULATE"

    elif rsi > 60 and macd == "Bullish":
        return "Momentum", "Trend continuation", "ADD"

    else:
        return "Wait", "Sideways", "HOLD"


# =============================
# TECHNICAL DATA
# =============================
def get_technical(symbol):
    df = yf.download(symbol, period="1y", interval="1d")

    if df.empty:
        print(f"[WARN] No data for {symbol}")
        return None

    # Fix multi-index
    if isinstance(df.columns, pd.MultiIndex):
        df.columns = df.columns.get_level_values(0)

    df = df.dropna(subset=['Close'])

    close_series = df['Close'].squeeze()

    try:
        df['EMA20'] = EMAIndicator(close_series, window=20).ema_indicator()
        df['EMA50'] = EMAIndicator(close_series, window=50).ema_indicator()
        df['EMA200'] = EMAIndicator(close_series, window=200).ema_indicator()

        df['RSI'] = RSIIndicator(close_series, window=14).rsi()

        macd = MACD(close_series)
        df['MACD'] = macd.macd_diff()
    except Exception as e:
        print(f"[ERROR] Indicator failed for {symbol}: {e}")
        return None

    # ✅ FIX: selective dropna
    df = df.dropna(subset=['EMA20','EMA50','EMA200','RSI','MACD'])

    # ✅ fallback if still empty
    if df.empty:
        print(f"[WARN] Falling back (ignore EMA200) for {symbol}")
        df = df.dropna(subset=['EMA20','EMA50','RSI','MACD'])

        if df.empty:
            return None

    latest = df.iloc[-1]

    close = float(latest['Close'])
    ema20 = float(latest['EMA20'])
    ema50 = float(latest['EMA50'])
    ema200 = float(latest.get('EMA200', ema50))  # fallback

    rsi = float(latest['RSI'])
    macd_val = float(latest['MACD'])

    ema_status = []
    ema_status.append("Above 20" if close > ema20 else "Below 20")
    ema_status.append("Above 50" if close > ema50 else "Below 50")
    ema_status.append("Above 200" if close > ema200 else "Below 200")

    return {
        "Current Price": round(close, 2),
        "RSI": round(rsi, 2),
        "MACD": "Bullish" if macd_val > 0 else "Bearish",
        "EMA Status": ", ".join(ema_status),
        "52W High": round(df['Close'].max(), 2),
        "52W Low": round(df['Close'].min(), 2)
    }

def calculate_growth(df):
    try:
        revenue_growth = ((df['Revenue'][-1] / df['Revenue'][0]) ** (1/3) - 1) * 100
        profit_growth = ((df['Profit'][-1] / df['Profit'][0]) ** (1/3) - 1) * 100
        return round(revenue_growth,2), round(profit_growth,2)
    except:
        return "NA", "NA"


def generate_decision(tech):

    rsi = tech["RSI"]
    macd = tech["MACD"]

    if rsi < 45 and macd == "Bearish":
        return "Pullback entry", "20-30% rebound", "HOLD / ACCUMULATE"

    elif rsi > 60 and macd == "Bullish":
        return "Momentum trade", "Trend continuation", "ADD"

    else:
        return "Wait", "Unclear", "HOLD"




# =============================
# INVESTOR TREND (PLACEHOLDER)
# =============================
def get_investor_trend(prev, curr):
    if curr > prev:
        return "Added"
    elif curr < prev:
        return "Reduced"
    else:
        return "Hold"

# =============================
# AI CLASSIFICATION LOGIC
# =============================
def generate_signal(data):
    rsi = data.get("RSI", 50)
    macd = data.get("MACD", "Neutral")

    if rsi < 45 and macd == "Bearish":
        return "Pullback entry", "HOLD / ACCUMULATE"
    elif rsi > 60 and macd == "Bullish":
        return "Momentum", "ADD"
    else:
        return "Wait", "HOLD"

def classify_mcap(mcap):
    if mcap > 50000:
        return "Large Cap"
    elif mcap > 5000:
        return "Mid Cap"
    else:
        return "Small Cap"

# =============================
# MAIN ANALYZER
# =============================
def analyze_stock(symbol):

    tech = get_technical(symbol)
    funda = get_fundamentals(symbol)
    investor = get_investor_data(symbol)
    growth = get_growth_data(symbol)

    if tech is None:
        return {"Company": symbol, "Error": "No data"}

    mcap_raw = funda.get("Market Cap", "")
    mcap_category = classify_market_cap(parse_market_cap(mcap_raw))

    short_term, mid_term, rating = generate_decision(tech)

    return {
        "Company Name": symbol.replace(".NS",""),
        "Market Cap Category": mcap_category,
        "Holding Funds (Top 5)": investor["Holding Funds"],
        "Avg. Holding %": investor["Avg Holding %"],
        "PE Ratio": funda.get("PE", ""),
        "ROE (%)": funda.get("ROE", ""),
        "ROCE (%)": funda.get("ROCE", ""),
        "Debt/ Equity": funda.get("Debt/Equity", ""),
        "Rev. Growth 3Y (%)": growth["Rev Growth 3Y"],
        "Profit Growth 3Y (%)": growth["Profit Growth 3Y"],
        "Current Price (₹)": tech["Current Price"],
        "52W High (₹)": tech["52W High"],
        "52W Low (₹)": tech["52W Low"],
        "Price vs EMA (20/50/200)": tech["EMA Status"],
        "RSI (14)": tech["RSI"],
        "MACD Signal": tech["MACD"],
        "Short-Term Goal Fit": short_term,
        "Mid-Term Goal Horizon": mid_term,
        "Investment Rating": rating
    }

def parse_market_cap(mcap_str):
    try:
        if not mcap_str:
            return None

        # Remove unwanted characters
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

def export_to_excel(df):
    filename = "Clean_Stock_Output.xlsx"

    with pd.ExcelWriter(filename, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Stocks")

        ws = writer.sheets["Stocks"]

        # Auto column width
        for col in ws.columns:
            max_length = max(len(str(cell.value)) for cell in col)
            ws.column_dimensions[col[0].column_letter].width = max_length + 2

    print("✅ Clean Excel Generated")


# =============================
# RUN MULTIPLE STOCKS
# =============================
stocks = [
"CYIENT.NS","BLUESTARCO.NS","WHIRLPOOL.NS","KALPATARU.NS","JUBLINGREA.NS","NH.NS","KIRLOSENG.NS","EQUITASBNK.NS","NAVINFLUOR.NS","MCX.NS","BRIGADE.NS","IPCALAB.NS","PATELENG.NS","ADVAIT.NS","SHANKARA.NS","SDBL.NS","ACE.NS","MAHSEAMLES.NS","NRBBEARING.NS","TIMETECHNO.NS","FIEMIND.NS","LAOPALA.NS","SHARDAMOTR.NS","TDPOWERSYS.NS","TGVSI.NS","ANUP.NS","FAIRCHEMOR.NS","NAVA.NS","SALZERELEC.NS","AMAL.NS","HITECH.NS","BIGBLOC.NS","KMCSHIL.NS","BMW.NS","SUNDARMFIN.NS","SRF.NS","KIMS.NS"]

results = [analyze_stock(s) for s in stocks]

df = pd.DataFrame(results)

df = df.fillna("Data unavailable")


export_to_excel(df)

print(df)
