# AI-Powered Stock Analysis Agent

## 📋 Overview

An automated daily stock analysis platform that performs **dual-layer analysis**: keeping rule-based recommendations separate from AI-based recommendations to enable meaningful comparison without influence between systems.

**Key Features:**
- ✅ Dual symbol source modes (predefined list in `config.json` or live AI selection from `.github/skill.md`)
- ✅ Dual recommendation engines (Rule-based for reference, AI-based for primary decisions)
- ✅ Comprehensive multi-dimensional analysis (valuation, momentum, trend, quality, safety, growth)
- ✅ Automatic data aggregation from yfinance, Screener.in, NSE APIs
- ✅ Excel export with detailed explanations and AI reasoning
- ✅ Natural-language stock Q&A grounded only in the generated Excel report
- ✅ VS Code MCP server for stock questions against the latest workbook
- ✅ Run summary text file with top picks and weekly change notes
- ✅ Historical CSV snapshots for week-over-week comparison
- ✅ Optional email delivery with top AI picks, top rule-based picks, aligned picks, and weekly summary
- ✅ Logging and audit trail for each run
- ✅ MCP-ready for daily scheduled execution

---

## 🚀 Quick Start

### 1. Install Dependencies
```bash
pip install yfinance pandas requests beautifulsoup4 ta-lib openpyxl mcp
```

### 2. Configure Symbol Source (Edit `config.json`)
```json
{
  "stocks": [
    "MCX.NS",
    "EQUITASBNK.NS",
    "NRBBEARING.NS"
  ],
  "symbol_selection": {
    "mode": "prompt",
    "default_option": "live_ai",
    "live_rules_file": ".github/skill.md",
    "live_top_n": 40,
    "live_min_parameters": 12,
    "live_min_pass_ratio": 0.55,
    "fallback_to_config_stocks": true
  },
  "export_file": "AI_STOCK_ANALYSIS.xlsx",
  "log_file": "stock_analysis.log",
  "history_dir": "analysis_history",
  "run_report_file": "latest_run_report.txt",
  "enable_ai_ranking": true,
  "enable_rule_based": true,
  "ai_thresholds": {
    "strong_buy_min_score": 75,
    "buy_min_score": 60,
    "hold_min_score": 45
  },
  "email": {
    "enabled": false,
    "smtp_host": "smtp.gmail.com",
    "smtp_port": 587,
    "use_tls": true,
    "sender_email": "",
    "sender_password_env": "STOCK_AGENT_EMAIL_PASSWORD",
    "recipients": []
  }
}
```

### 3. Run Analysis
```bash
python Stock_Agent.py
```

### 4. Ask a Grounded Stock Question From Excel
```bash
python Stock_Agent.py --ask-stock MCX --question "Should I buy this stock for long-term accumulation?"
```

This query mode reads only the generated `Analysis` sheet and uses the stored Excel evidence:
- AI Score
- AI Recommendation
- AI Confidence
- AI Justification
- Rule-based rating
- Query match percentages
- Reported fundamentals and technical fields

If GenAI is enabled, the answer is still forced to stay within the workbook evidence. If GenAI is unavailable, the script falls back to a deterministic evidence summary.

**Outputs:**
- `AI_STOCK_ANALYSIS.xlsx` with dual-track analysis
- `latest_run_report.txt` with top 5 AI picks, top 5 rule-based picks, aligned picks, and weekly summary
- `analysis_history/analysis_YYYYMMDD_HHMMSS.csv` archived snapshot for historical comparison

---

## MCP Q&A In VS Code

This repo now includes a local MCP server entry point: `stock_excel_mcp_server.py`.

It exposes 3 tools to VS Code/GitHub Copilot Chat:
- `list_available_stocks`
- `get_stock_snapshot`
- `answer_stock_question`

The workspace also includes `.vscode/mcp.json`, so VS Code can register the server as a stdio MCP server.

Example questions once the MCP server is enabled in VS Code:
- "What does the report say about MCX?"
- "Should I buy EQUITASBNK based on the latest Excel?"
- "Compare CYIENT and BLUESTARCO using the stored AI recommendation and query match percentages."

Grounding rules:
- The server reads only the generated Excel report.
- Recommendation and confidence are anchored to the stored workbook values.
- If the workbook lacks enough evidence, the response says so explicitly instead of guessing.

---

## 📊 Output Columns

### A. Identification (4 columns)
| Column | Description |
|--------|-------------|
| Company Name | Stock symbol without exchange suffix |
| Company Full Name | Full legal name from yfinance |
| Sector | Business sector classification |
| Market Cap Category | Large Cap (>50K Cr) / Mid Cap (5K-50K) / Small Cap (<5K) |

### B. Holdings & Ownership (2 columns)
| Column | Description |
|--------|-------------|
| Holding Funds (Top 5) | Named fund houses + individual shareholders with stakes |
| Avg. Holding % | Average of major institutional/insider holdings |

### C. Fundamental Metrics (7 columns)
| Column | Description |
|--------|-------------|
| PE Ratio | Price-to-Earnings multiple from Screener.in |
| ROE (%) | Return on Equity percentage |
| ROCE (%) | Return on Capital Employed percentage |
| Debt/Equity | Leverage ratio from yfinance |
| Rev. Growth 3Y (%) | 3-year revenue growth rate |
| Profit Growth 3Y (%) | 3-year earnings/profit growth rate |

### D. Enhanced Valuation Metrics (5 columns)
| Column | Description |
|--------|-------------|
| Price-to-Book | P/B ratio (valuation relative to book value) |
| Price-to-Sales | P/S ratio (valuation relative to revenue) |
| PEG Ratio | Price/Earnings-to-Growth ratio |
| Beta | Stock volatility vs. market |
| Dividend Yield (%) | Annual dividend as % of price |

### E. Quality Metrics (2 columns)
| Column | Description |
|--------|-------------|
| Profit Margin (%) | Net profit as % of revenue |
| Operating Margin (%) | Operating profit as % of revenue |

### F. Technical Metrics (6 columns)
| Column | Description |
|--------|-------------|
| Current Price (₹) | Last trading price |
| 52W High (₹) | 52-week high price |
| 52W Low (₹) | 52-week low price |
| Price vs EMA (20/50/200) | Price position vs. exponential moving averages |
| RSI (14) | Relative Strength Index (momentum 0-100) |
| MACD Signal | Bullish/Bearish crossover signal |

### G. Rule-Based Analysis (4 columns) - **FOR REFERENCE**
| Column | Description |
|--------|-------------|
| Rule-Based Score | Traditional scoring (0-60 max) from technical + fundamental signals |
| Short-Term Goal Fit | Momentum trade vs. Pullback entry classification |
| Mid-Term Goal Horizon | Suggested holding period with reasoning |
| Rule-Based Rating | Action: HOLD/ACCUMULATE, HOLD, or WATCHLIST |

### H. AI-Based Analysis (5 columns) - **PRIMARY RECOMMENDATION**
| Column | Description |
|--------|-------------|
| AI Score | Composite AI score (0-100) across 6 dimensions |
| AI Ranking | Ranked recommendation with confidence level |
| AI Recommendation | **STRONG BUY / BUY / HOLD / REDUCE/AVOID** |
| AI Confidence | Confidence level: HIGH / MEDIUM-HIGH / MEDIUM |
| AI Justification | Detailed explanation of recommendation with top scoring factors |

### I. Metadata (1 column)
| Column | Description |
|--------|-------------|
| Last Updated | Timestamp of analysis execution |

---

## 🧠 AI Scoring Logic

The AI Score (0-100) is calculated across **6 independent dimensions**:

### 1. **Valuation** (0-20 points) - Attractiveness of price
- Price-to-Book < 2.0: +10 pts
- Price-to-Sales < 2.0: +10 pts  
- PEG Ratio ≤ 1.0: +10 pts

### 2. **Momentum** (0-20 points) - Short-term price strength
- RSI 40-70: +15 pts (optimal zone)
- RSI 30-40 or 70-80: +8 pts (caution zone)
- MACD Bullish: +5 pts

### 3. **Trend** (0-15 points) - Price positioning vs. EMAs
- Price > EMA200: +10 pts (above long-term trend)
- Price > EMA50: +5 pts (above medium-term trend)

### 4. **Quality** (0-20 points) - Business fundamentals
- ROE > 15%: +10 pts
- ROCE > 15%: +10 pts
- (Lower thresholds: >10% = +5 pts each)

### 5. **Safety** (0-15 points) - Balance sheet strength
- Debt/Equity < 1.0: +10 pts (healthy leverage)
- Profit Margin > 10%: +5 pts

### 6. **Growth** (0-10 points) - Expansion trajectory
- Revenue Growth > 10%: +5 pts
- Earnings Growth > 10%: +5 pts

**Total: 0-100 points**

### Recommendation Mapping
| AI Score | Recommendation | Confidence | Ideal For |
|----------|----------------|-----------|-----------|
| ≥ 75 | **STRONG BUY** | HIGH | Aggressive accumulation |
| 60-74 | **BUY** | MEDIUM-HIGH | Regular portfolio building |
| 45-59 | **HOLD** | MEDIUM | Monitor, accumulate on dips |
| < 45 | **REDUCE/AVOID** | MEDIUM | Wait for better entry or skip |

---

## 📁 File Structure

```
AI-Stock-Agent/
├── Stock_Agent.py           # Main analysis script
├── stock_excel_mcp_server.py # MCP server for Excel-backed stock Q&A
├── config.json              # Configuration (stocks, thresholds, output)
├── latest_run_report.txt     # Latest run summary (auto-generated)
├── analysis_history/         # Archived CSV snapshots (auto-generated)
├── .github/
│   ├── copilot-instructions.md      # Repo-wide Copilot guidance
│   └── agents/stock-analysis.agent.md  # Reusable stock-analysis agent description
├── .vscode/mcp.json         # Workspace MCP server registration for VS Code
├── README.md                # This file
├── stock_analysis.log       # Execution log (auto-generated)
├── AI_STOCK_ANALYSIS.xlsx   # Output analysis (auto-generated)
└── Stock_Agent_old.py       # Backup of previous version
```

---

## ⚙️ Configuration (`config.json`)

```json
{
  "stocks": [
    "MCX.NS",
    "EQUITASBNK.NS"
  ],
  "export_file": "AI_STOCK_ANALYSIS.xlsx",
  "log_file": "stock_analysis.log",
  "history_dir": "analysis_history",
  "run_report_file": "latest_run_report.txt",
  "enable_ai_ranking": true,
  "enable_rule_based": true,
  "ai_thresholds": {
    "strong_buy_min_score": 75,
    "buy_min_score": 60,
    "hold_min_score": 45
  },
  "email": {
    "enabled": false,
    "smtp_host": "smtp.gmail.com",
    "smtp_port": 587,
    "use_tls": true,
    "sender_email": "",
    "sender_password_env": "STOCK_AGENT_EMAIL_PASSWORD",
    "recipients": []
  }
}
```

### Configuration Parameters
| Parameter | Type | Description |
|-----------|------|-------------|
| `stocks` | Array | List of NSE stock symbols (e.g., "MCX.NS") |
| `symbol_selection.mode` | String | `predefined`, `live_ai`, or `prompt` |
| `symbol_selection.default_option` | String | Default option used when `mode` is `prompt` |
| `symbol_selection.live_rules_file` | String | Markdown file containing live selection query rules |
| `symbol_selection.live_top_n` | Integer | Maximum symbols selected in live mode |
| `symbol_selection.live_min_parameters` | Integer | Minimum recognized rules inferred from `.github/skill.md` query (default: 12) |
| `symbol_selection.live_min_pass_ratio` | Float | Minimum parameter pass ratio to qualify in live mode |
| `symbol_selection.fallback_to_config_stocks` | Boolean | Use predefined list if live selection fails |
| `export_file` | String | Excel output filename |
| `log_file` | String | Log file path |
| `history_dir` | String | Folder used to archive CSV snapshots for weekly comparison |
| `run_report_file` | String | Text summary generated after every run |
| `enable_ai_ranking` | Boolean | Enable AI-based recommendations |
| `enable_rule_based` | Boolean | Enable rule-based analysis (for comparison) |
| `ai_thresholds.strong_buy_min_score` | Integer | Minimum score for STRONG BUY (default: 75) |
| `ai_thresholds.buy_min_score` | Integer | Minimum score for BUY (default: 60) |
| `ai_thresholds.hold_min_score` | Integer | Minimum score for HOLD (default: 45) |
| `email.enabled` | Boolean | Send an email after each run when SMTP is configured |
| `email.smtp_host` | String | SMTP server host |
| `email.smtp_port` | Integer | SMTP server port |
| `email.use_tls` | Boolean | Enable STARTTLS before login |
| `email.sender_email` | String | From-address used for SMTP login |
| `email.sender_password_env` | String | Environment variable name that stores the SMTP app password |
| `email.recipients` | Array | One or more recipient email addresses |

### Email Password Setup (Environment Variable)

Do not store SMTP passwords in `config.json`.

Use an app password and set it in an environment variable:

PowerShell (current session):
```powershell
$env:STOCK_AGENT_EMAIL_PASSWORD = "your-16-char-app-password"
python .\Stock_Agent.py
```

PowerShell (persist for your user profile):
```powershell
[System.Environment]::SetEnvironmentVariable("STOCK_AGENT_EMAIL_PASSWORD", "your-16-char-app-password", "User")
```

After setting a persistent variable, restart terminal/VS Code before running.

### Email Report Contents

When `email.enabled` is `true`, each run sends:
- An investor-focused GenAI-style brief headline and market stance
- Key investor insights (opportunity mix, alignment, suggested accumulation posture)
- Top long-term quality candidates (profitability + leverage filters)
- Top 5 AI recommendations
- Top 5 rule-based recommendations
- Top aligned AI and rule-based recommendations
- Risk alerts (cautious names / leverage and overheating checks)
- Weekly change summary from archived history
- "What to monitor next" points tailored for long-term investors

The message is delivered HTML-first with a compact plain text fallback.
Email subject shows symbol source mode:
- `Investor Brief | Predefined Symbols | YYYY-MM-DD`
- `Investor Brief | Live Selection by AI | YYYY-MM-DD`

If no prior snapshot exists yet, the weekly section explicitly says that history is not available yet.

### Copilot Files In `.github`

You do not strictly need a `copilot-instructions.md` file to use this project. The script will run without it.

It is useful when you want always-on repository guidance for Copilot. This repo now includes:
- `.github/copilot-instructions.md` for shared repo instructions
- `.github/agents/stock-analysis.agent.md` for a stock-analysis-specific agent description

If the filename is misspelled, Copilot may ignore it. Use the exact name `copilot-instructions.md`.

---

## 📡 Data Sources

| Source | Data | Method |
|--------|------|--------|
| **yfinance** | Prices, growth rates, valuations, margins, holdings % | API |
| **Screener.in** | P/E, ROE, ROCE, Debt, Market Cap | Web scraping (BeautifulSoup) |
| **NSE India API** | Shareholding pattern (MF%, FII%) | REST API |

---

## 🔄 Workflow

```
1. Load Configuration
  └─ Read thresholds and symbol source mode from config.json

2. Resolve Symbols
  ├─ `predefined` mode: use `stocks` from config
  └─ `live_ai` mode: read query rules from `.github/skill.md`, auto-discover NSE symbols, rank symbols, choose top names

3. For Each Stock
   ├─ Extract Technical Data
   │  └─ Download 1-year price history
   │  └─ Calculate: EMA(20/50/200), RSI(14), MACD
   ├─ Extract Fundamental Data
   │  └─ Screener.in: P/E, ROE, ROCE, Debt, Market Cap
   ├─ Extract Quote Data (yfinance)
   │  └─ Growth rates, valuations, margins, holdings %
   ├─ Extract Investor Data
   │  └─ NSE API: MF%, FII%
   │  └─ Named holders: Fund houses & individuals
   ├─ RULE-BASED ANALYSIS (Reference)
   │  └─ Score = Technical + Fundamental + Investor signals
   │  └─ Recommendation = HOLD/ACCUMULATE vs HOLD vs WATCHLIST
   └─ AI-BASED ANALYSIS (Independent)
      └─ Score = 6-dimensional composite (0-100)
      └─ Recommendation = STRONG BUY / BUY / HOLD / REDUCE/AVOID

4. Rank Stocks
   └─ Sort by AI Score (descending)
   └─ Assign rank (1 = highest AI score)

5. Export Results
   └─ Write to Excel: AI_STOCK_ANALYSIS.xlsx
   └─ Display top 10 stocks with AI vs Rule-Based comparison

6. Logging
   └─ Write execution log to stock_analysis.log
```

---

## 🔧 Usage Examples

### Run Standard Analysis
```bash
python Stock_Agent.py
```
Processes symbols based on `symbol_selection.mode`, then exports to `AI_STOCK_ANALYSIS.xlsx`

### Run Live AI Selection
Set:
```json
{
  "symbol_selection": {
    "mode": "live_ai",
    "live_rules_file": ".github/skill.md",
    "live_min_parameters": 12
  }
}
```

### Prompt User Choice Each Run
Set:
```json
{
  "symbol_selection": {
    "mode": "prompt",
    "default_option": "predefined"
  }
}
```

### Add New Predefined Stocks
Edit `config.json`:
```json
{
  "stocks": [
    "MCX.NS",
    "EQUITASBNK.NS",
    "YOUR_SYMBOL.NS"
  ]
}
```

### Change Output Filename
```json
{
  "export_file": "My_Analysis_2025-01-15.xlsx"
}
```

### Adjust AI Thresholds
```json
{
  "ai_thresholds": {
    "strong_buy_min_score": 80,
    "buy_min_score": 65,
    "hold_min_score": 50
  }
}
```

### Disable Rule-Based (AI-only)
```json
{
  "enable_rule_based": false
}
```

---

## 🧩 `.github/skill.md` Format (Live Mode)

`.github/skill.md` should include:

```md
## Discovery Rules
- universe_scan_limit: 350

## Query
Market Capitalization > 500 and Market Capitalization < 50000
AND Sales growth 3Years > 12
AND Profit growth 3Years > 12
AND Return on equity > 15
AND Return on capital employed > 18
AND Debt to equity < 0.5
AND Interest Coverage Ratio > 4
AND Current ratio > 1.2
AND OPM > 12
AND PEG Ratio < 2
AND Promoter holding > 50
AND Pledged percentage < 3
AND Cash from operations last year > 0
AND Price to earning < 35
```

Notes:
- Keep the query in plain text as shown.
- Script auto-discovers NSE symbols (bounded by `universe_scan_limit`) and applies the query-derived rules.
- If live selection fails and fallback is enabled, script falls back to predefined `stocks`.

---

## 📋 Logging

**Log File:** `stock_analysis.log`

Each run logs:
- Start/end times
- Configuration loaded (stock count)
- Processing status for each symbol
- Analysis completion summary
- Total stocks analyzed

**View Recent Logs:**
```bash
tail -20 stock_analysis.log
```

---

## 🕐 Scheduling (Daily Execution)

### Windows Task Scheduler
```
Task Name: AI Stock Analysis
Trigger: Daily at 4:00 PM (after market close)
Action: python "C:\path\to\Stock_Agent.py"
```

### Linux/Mac (cron)
```bash
# Daily at 4 PM IST (10:30 AM UTC)
0 16 * * * /usr/bin/python3 /path/to/Stock_Agent.py
```

### MCP Agent Integration
For MCP-based scheduling, wrap the analyzer:
```python
# stock_mcp_agent.py
import subprocess
import json
from datetime import datetime

def run_daily_analysis():
    result = subprocess.run(["python", "Stock_Agent.py"], capture_output=True)
    
    if result.returncode == 0:
        return {"status": "success", "timestamp": datetime.now().isoformat()}
    else:
        return {"status": "failed", "error": result.stderr.decode()}

if __name__ == "__main__":
    run_daily_analysis()
```

---

## ⚠️ Troubleshooting

### Error: `PermissionError: [Errno 13] Permission denied: 'AI_STOCK_ANALYSIS.xlsx'`
**Cause:** Excel file is open and locked  
**Solution:** Close Excel or rename `export_file` in config.json

### Error: `yfinance: No data found for symbol`
**Cause:** Stock symbol not valid or NSE-listed  
**Solution:** Verify symbol format (e.g., "MCX.NS" for NSE)

### Error: `requests.ConnectionError: Unable to connect to Screener.in`
**Cause:** Screener.in server unavailable or blocked  
**Solution:** Retry later; Screener.in data is non-critical (fallback to yfinance)

### Warning: `[SKIP] Insufficient data for symbol`
**Cause:** Stock has < 100 days of trading history  
**Solution:** Stock is too new; add to watchlist for future analysis

### No Output Columns
**Cause:** All stocks skipped due to data unavailability  
**Solution:** Verify symbols are NSE-listed and have sufficient trading history

---

## 📈 AI vs Rule-Based Comparison

### When to Use Each

**AI Recommendation (Primary):**
- Holistic multi-dimensional scoring (valuation, momentum, trend, quality, safety, growth)
- Not influenced by traditional rule-based signals
- Best for: Long-term portfolio building, balanced risk-adjusted decisions

**Rule-Based Rating (Reference):**
- Traditional technical + fundamental signals
- Useful for: Momentum validation, comparing against legacy analysis
- Use to understand: Why AI differs from conventional wisdom

### Example Interpretation

| Stock | AI Score | AI Rec | Rule-Based | Interpretation |
|-------|----------|--------|-----------|-----------------|
| Stock A | 78 | STRONG BUY | WATCHLIST | AI sees quality + valuation; rule-based lacks momentum |
| Stock B | 42 | HOLD | ACCUMULATE | AI cautious on overall metrics; rule-based focuses on signals |
| Stock C | 65 | BUY | HOLD | Good balance; both systems agree on moderate rating |

---

## 📚 Technical Details

### Dependencies
- **yfinance**: Stock price data, fundamentals, institutional holdings
- **pandas**: Data manipulation and Excel export
- **BeautifulSoup4**: HTML parsing for Screener.in
- **ta-lib**: Technical analysis (EMA, RSI, MACD)
- **requests**: HTTP requests with rate limiting
- **openpyxl**: Excel file handling

### Performance
- **Processing Time:** ~1-2 min per stock (API calls + calculations)
- **Full Run (37 stocks):** ~45-60 minutes
- **Memory Usage:** ~100-200 MB
- **API Rate Limiting:** 0.5 sec delay between requests

---

## 🤝 Contributing

### Adding Custom Indicators
Extend `calculate_ai_score()` function:
```python
# Add new dimension (0-10 points)
custom_score = 0
if some_condition:
    custom_score += 10
score += custom_score
breakdown["Custom Metric"] = custom_score
```

### Modifying Data Sources
Edit data extraction functions:
- `get_technical()` - Technical indicators
- `get_fundamentals()` - Company metrics
- `get_quote_data()` - Quote-level data
- `get_investor_data()` - Shareholding patterns

---

## 📄 License

This project is provided as-is for personal stock analysis use.

---

## 📞 Support

For issues or questions:
1. Check the **Troubleshooting** section
2. Review **stock_analysis.log** for error details
3. Verify configuration in **config.json**
4. Ensure all dependencies are installed: `pip list`

---

**Last Updated:** 2025-01-15  
**Version:** 2.0 (AI + Rule-Based Dual Analysis)
