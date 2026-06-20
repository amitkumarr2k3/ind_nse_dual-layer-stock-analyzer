# NSE Dual-Track Stock Analyzer

> **Automated dual-engine stock analysis for Indian equities listed on the National Stock Exchange (NSE)**

Runs two fully independent analysis engines side-by-side — a traditional rule-based checklist and a sector-aware AI scoring model — and exports everything into a single Excel workbook for daily review.

---

## What It Does

- Fetches live data for any NSE-listed stock (symbol ending in `.NS`)
- Scores each stock using **two independent systems** so you can compare them:
  - **Rule-Based** (0–60): Quick checklist of momentum, profitability, and institutional holding signals
  - **AI-Based** (0–100): Six-dimensional sector-aware scoring across valuation, quality, growth, safety, technicals, and governance
- Computes a **Composite Rating** (letter grade) blending both, capped to the AI recommendation level
- Exports 40+ columns per stock to Excel including 3Y / 5Y / 10Y growth CAGRs
- Archives a CSV snapshot after every run for week-over-week comparison
- Optionally emails an investor brief with top picks and weekly changes

---

## Scope

**Indian stocks on NSE only.**
Data is sourced from yfinance, Screener.in (Indian financials), and the NSE India shareholding API. It is not designed for US or global equities.

---

## Quick Start

### 1. Install dependencies
```bash
pip install yfinance pandas requests beautifulsoup4 ta openpyxl mcp
```

### 2. Configure your stock list (`config.json`)
```json
{
  "stocks": ["RELIANCE.NS", "TCS.NS", "NMDC.NS"],
  "export_file": "NSE_ANALYSIS.xlsx",
  "enable_ai_ranking": true
}
```
All NSE symbols must end in `.NS`.

### 3. Run the analyzer
```bash
python nse_dual_track_analyzer.py
```

### 4. Ask a question about a stock (grounded in the Excel output)
```bash
python nse_dual_track_analyzer.py --ask-stock NMDC --question "Should I buy for long-term?"
```

Use a different config:
```bash
python nse_dual_track_analyzer.py --config config_maincheck.json
```

---

## Output: Excel Columns

### Identification
| Column | Description |
|--------|-------------|
| Company Name | NSE ticker without `.NS` suffix |
| Company Full Name | Legal registered name |
| Sector | Business sector |
| Market Cap Category | Large / Mid / Small / Micro cap |

### Holdings & Ownership
| Column | Description |
|--------|-------------|
| Promoter Holding (%) | Founder / promoter stake |
| FII Holding (%) | Foreign Institutional Investors |
| DII Holding (%) | Domestic Institutional Investors |
| MF Holding (%) | Mutual Fund holding (from NSE data) |
| Holding Funds (Top 5) | Named fund houses with stakes |
| FII / DII Trend (3Q) | Buying or selling trend over last 3 quarters |

### Fundamentals & Growth
| Column | Description |
|--------|-------------|
| PE Ratio | Price-to-Earnings (Screener.in) |
| ROE (%) | Return on Equity |
| ROCE (%) | Return on Capital Employed |
| Debt / Equity | Leverage ratio |
| Rev. Growth 3Y (%) | 3-year revenue CAGR |
| Rev. Growth 5Y (%) | 5-year revenue CAGR |
| Rev. Growth 10Y/7Y (%) | 10-year CAGR (falls back to 7Y if unavailable) |
| Profit Growth 3Y (%) | 3-year profit CAGR |
| Profit Growth 5Y (%) | 5-year profit CAGR |
| Profit Growth 10Y/7Y (%) | 10-year CAGR (falls back to 7Y if unavailable) |

### Valuation
| Column | Description |
|--------|-------------|
| Price-to-Book | P/B ratio |
| Price-to-Sales | P/S ratio |
| PEG Ratio | PE ÷ Growth (computed) |
| Sector PE Comparison | Cheap / Fair / Pricey vs sector median |
| Beta | Volatility relative to Nifty |

### Quality
| Column | Description |
|--------|-------------|
| Profit Margin (%) | Net profit margin |
| Operating Margin (%) | Operating profit margin |
| Dividend Yield (%) | Annual yield |

### Technical
| Column | Description |
|--------|-------------|
| Current Price (INR) | Last traded price |
| 52W High / 52W Low | 52-week range |
| Price vs EMA (20/50/200) | Position above/below moving averages |
| RSI (14) | Momentum indicator |
| MACD Signal | Bullish / Bearish crossover |

### Rule-Based Analysis (reference)
| Column | Description |
|--------|-------------|
| Rule-Based Score | 0–60 composite score |
| Rule-Based Rating | HOLD/ACCUMULATE · HOLD · WATCHLIST |
| Mid-Term Goal Horizon | Suggested holding window with reason |

### AI-Based Analysis (primary)
| Column | Description |
|--------|-------------|
| AI Score | 0–100 composite score |
| AI Recommendation | STRONG BUY · BUY · HOLD · REDUCE/AVOID |
| AI Confidence | HIGH · MEDIUM · LOW |
| AI Justification | Key scoring factors and cautions |
| Sector Model | Scoring profile applied (General / Bank / Insurance …) |
| Data Quality | Completeness of underlying data |
| Risk Flags | Hard blockers that reduced conviction |
| Why AI View Can Be Wrong | Guardrail on what to watch |

### Composite & Metadata
| Column | Description |
|--------|-------------|
| Composite Rating | Letter grade blending AI (70%) + Rule (30%), capped to AI action |
| Last Updated | Run timestamp |

---

## AI Scoring (0–100)

Six independent dimensions, each scored and summed:

| Dimension | Max | What it measures |
|-----------|-----|-----------------|
| Valuation (sector-aware) | 15 | PE, PB, PS vs sector |
| Technical setup | 15 | RSI, MACD, EMA position |
| Quality (ROE/ROCE/margins) | 22 | Profitability efficiency |
| Growth quality | 18 | Revenue & profit CAGRs |
| Safety / balance sheet | 18 | Debt, interest coverage, cash flow |
| Governance / investability | 12 | Promoter stake, pledge %, data completeness |

**Sector-aware:** Banks, NBFCs, insurance, and housing finance are not penalised for high leverage (structurally normal). All other sectors use standard debt/equity thresholds.

### Recommendation thresholds (configurable)
| Score | Recommendation |
|-------|----------------|
| ≥ 75 | **STRONG BUY** |
| 60–74 | **BUY** |
| 45–59 | **HOLD** |
| < 45 | **REDUCE / AVOID** |

---

## Data Sources

| Source | Data fetched |
|--------|-------------|
| **yfinance** | Prices, technicals, growth rates, margins, valuations |
| **Screener.in** | Indian-native P/E, ROE, ROCE, growth CAGRs (3Y/5Y/10Y) |
| **NSE India API** | Shareholding pattern: FII%, DII%, MF%, Promoter% |

Screener.in data is preferred over yfinance for Indian fundamentals because it uses INR-denominated financials and eliminates USD/INR currency mismatch in P/S and margins.

---

## Project File Structure

```
nse-dual-track-analyzer/
├── nse_dual_track_analyzer.py      # Main analysis script
├── stock_excel_mcp_server.py       # MCP server for Excel-backed Q&A in VS Code
├── config.json                     # Default configuration (edit symbols here)
├── config_maincheck.json           # Full symbol list
├── config_maincheck_small.json     # Small 8-stock test config
├── AI_VS_RULE_BASED_GUIDE.md       # Plain-English explanation for non-technical users
├── MCP_AGENT_SETUP.md              # Daily scheduling and automation guide
├── SKILL_SMALLCAP_SCREENER.md      # Small/micro cap screener rules
├── .github/
│   ├── copilot-instructions.md     # Repo-scoped Copilot guidance
│   └── skill.md                    # Live AI symbol selection query rules
├── analysis_history/               # Auto-generated CSV snapshots (git-ignored)
├── latest_run_report.txt           # Auto-generated run summary (git-ignored)
└── stock_analysis.log              # Execution log (git-ignored)
```

---

## Configuration Reference (`config.json`)

```json
{
  "stocks": ["RELIANCE.NS", "TCS.NS"],
  "symbol_selection": {
    "mode": "predefined",
    "live_rules_file": ".github/skill.md",
    "live_top_n": 40,
    "fallback_to_config_stocks": true
  },
  "export_file": "NSE_ANALYSIS.xlsx",
  "history_dir": "analysis_history",
  "run_report_file": "latest_run_report.txt",
  "enable_ai_ranking": true,
  "ai_thresholds": {
    "strong_buy_min_score": 75,
    "buy_min_score": 60,
    "hold_min_score": 45
  },
  "email": {
    "enabled": false,
    "smtp_host": "smtp.gmail.com",
    "smtp_port": 587,
    "sender_email": "you@gmail.com",
    "sender_password_env": "STOCK_AGENT_EMAIL_PASSWORD",
    "recipients": ["you@gmail.com"]
  }
}
```

**Email password:** Never put passwords in `config.json`. Set an environment variable:
```powershell
$env:STOCK_AGENT_EMAIL_PASSWORD = "your-app-password"
```

---

## Symbol Selection Modes

| Mode | Description |
|------|-------------|
| `predefined` | Uses the `stocks` list in config.json |
| `live_ai` | Dynamically selects top candidates using rules in `.github/skill.md` |
| `prompt` | Asks at runtime which mode to use |

---

## Scheduling (Daily Automation)

See [MCP_AGENT_SETUP.md](MCP_AGENT_SETUP.md) for full instructions.

**Quick Windows Task Scheduler example:**
```
Program:   C:\path\to\.venv\Scripts\python.exe
Arguments: C:\path\to\nse_dual_track_analyzer.py --config config.json
Run at:    4:00 PM IST (market close + 30 min)
```

---

## VS Code MCP Server (Stock Q&A)

Start the local MCP server to ask natural-language questions against the generated Excel:
```bash
python stock_excel_mcp_server.py
```

Then in VS Code Copilot Chat:
- "What does the latest report say about NMDC?"
- "Which stocks are STRONG BUY with HIGH confidence?"
- "Compare RELIANCE and TCS by AI score."

Answers are grounded exclusively in the generated Excel — no hallucination.

---

## Disclaimer

This tool is for **informational and educational purposes only**. It is not financial advice. All data is fetched from public sources and may be delayed or inaccurate. Always do your own research before making investment decisions.

---

## License

MIT License. Free to use, modify, and distribute.

