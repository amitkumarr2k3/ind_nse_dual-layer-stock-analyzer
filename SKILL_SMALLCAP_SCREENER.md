---
name: smallcap-screener
version: "1.0"
description: "Rules to discover best-of-breed NSE small & micro cap stocks (market cap ₹50 Cr – ₹5000 Cr) using 50+ fundamental, technical, and quality parameters."

# Number of top candidates to pass on for full dual-track analysis
output_limit: 30

# Minimum pre-screen skill score to qualify (0–100 scale based on weights below)
min_pass_score: 20

# ─────────────────────────────────────────────────────────────────
# CANDIDATE UNIVERSE
# Broad pool of NSE small & micro cap stocks across sectors.
# The screener will apply filters and return the top `output_limit`.
# ─────────────────────────────────────────────────────────────────
candidate_universe:
  # Financial Services
  - EQUITASBNK.NS
  - UJJIVANSFB.NS
  - KARURVYSYA.NS
  - DCBBANK.NS
  - CREDITACC.NS
  - APTUS.NS
  - HOMEFIRST.NS
  - ANGELONE.NS
  - CDSL.NS
  - KFINTECH.NS
  - CAMS.NS
  - BSE.NS
  - RBLBANK.NS
  - SOUTHBANK.NS
  - ESAFSFB.NS
  # Engineering & Capital Goods
  - ANUP.NS
  - TDPOWERSYS.NS
  - KIRLOSENG.NS
  - SALZERELEC.NS
  - ACE.NS
  - MAHSEAMLES.NS
  - GRINDWELL.NS
  - KAYNES.NS
  - CRAFTSMAN.NS
  - JBMA.NS
  - SUPRAJIT.NS
  - ORIENTELEC.NS
  - LGBBROSLTD.NS
  - BEML.NS
  # Chemicals & Specialty
  - NAVINFLUOR.NS
  - FAIRCHEMOR.NS
  - DEEPAKFERT.NS
  - GNFC.NS
  - EPIGRAL.NS
  - NOCIL.NS
  - APCOTEXIND.NS
  - PCBL.NS
  - HSCL.NS
  - GAEL.NS
  # Consumer, Retail & FMCG
  - SAFARI.NS
  - RADICO.NS
  - JYOTHYLAB.NS
  - INDIAMART.NS
  - JUSTDIAL.NS
  - MHRIL.NS
  - MPSLTD.NS
  - ZYDUSWELL.NS
  # IT / Technology
  - CYIENT.NS
  - NIITLTD.NS
  - TIMETECHNO.NS
  - VIMTALABS.NS
  - EASEMYTRIP.NS
  # Pharma & Healthcare
  - IPCALAB.NS
  - NH.NS
  - KRSNAA.NS
  - AJANTPHARM.NS
  - GRANULES.NS
  - ERIS.NS
  # Materials, Metals & Industrials
  - NRBBEARING.NS
  - JINDALSAW.NS
  - RATNAMANI.NS
  - RAJRATAN.NS
  - FIEMIND.NS
  - CARBORUNIV.NS
  - GPIL.NS
  - PRINCEPIPE.NS
  - AMAL.NS
  - HITECH.NS
  # Infrastructure & Construction
  - PATELENG.NS
  - BRIGADE.NS
  - JUBLINGREA.NS
  - SHANKARA.NS
  - KMCSHIL.NS
  # Diversified / Other
  - MCX.NS
  - LAOPALA.NS
  - SHARDAMOTR.NS
  - TGVSI.NS
  - SDBL.NS
  - ADVAIT.NS
  - BMW.NS
  - NAVA.NS
  - BIGBLOC.NS
  - BLUESTARCO.NS
  - SRF.NS
  - KIMS.NS
  - WHIRLPOOL.NS

# ─────────────────────────────────────────────────────────────────
# HARD FILTER CRITERIA (50+ parameters)
# A stock must pass ALL applicable filters to qualify.
# Parameters marked "(soft)" are used for scoring only, not as
# hard cut-offs, because the data may not always be available.
# ─────────────────────────────────────────────────────────────────
filters:

  # ── 1. MARKET CAP RANGE (defines small/micro cap universe) ──────
  min_market_cap_cr: 50          # ₹50 Cr minimum — avoid nano/illiquid shells
  max_market_cap_cr: 5000        # ₹5000 Cr maximum — true small cap boundary

  # ── 2. PRICE FLOOR ──────────────────────────────────────────────
  min_price_inr: 15              # Exclude sub-₹15 penny stocks (liquidity / manipulation risk)

  # ── 3. PROFITABILITY ────────────────────────────────────────────
  min_eps_positive: true         # Company must be profitable (positive trailing EPS)
  min_book_value_positive: true  # Net tangible assets must be positive
  min_roe_pct: 10                # Return on Equity ≥ 10 % (capital efficiency floor)
  min_roce_pct: 10               # Return on Capital Employed ≥ 10 % (operational efficiency)
  min_profit_margin_pct: 6       # Net profit margin ≥ 6 % (sector-agnostic floor)
  min_operating_margin_pct: 8    # Operating margin ≥ 8 % (pre-interest / pre-tax efficiency)
  min_ebitda_margin_pct: 10      # EBITDA margin ≥ 10 % (soft — used for scoring)

  # ── 4. VALUATION ────────────────────────────────────────────────
  max_pe_ratio: 45               # P/E ≤ 45 — avoid richly priced growth stories
  max_price_to_book: 7           # P/B ≤ 7 — reasonable premium over book value
  max_price_to_sales: 6          # P/S ≤ 6 — revenue-based valuation sanity check
  max_peg_ratio: 3.0             # PEG ≤ 3 — growth-adjusted valuation guard
  max_ev_ebitda: 25              # EV/EBITDA ≤ 25 (soft — used for scoring when available)
  min_earnings_yield_pct: 1.5    # Earnings Yield (1/PE) ≥ 1.5 % (soft)

  # ── 5. GROWTH ───────────────────────────────────────────────────
  min_revenue_growth_pct: 5      # Revenue growth (YoY) ≥ 5 % — growing business
  min_earnings_growth_pct: 0     # Earnings growth ≥ 0 % — at least flat profits
  min_retained_earnings_growth_pct: 3   # Retained earnings trend (soft)
  min_3y_cagr_revenue_pct: 8     # 3-year revenue CAGR ≥ 8 % (aspirational / soft)
  min_3y_cagr_earnings_pct: 5    # 3-year earnings CAGR ≥ 5 % (aspirational / soft)

  # ── 6. LEVERAGE & BALANCE SHEET SAFETY ──────────────────────────
  max_debt_equity: 1.5           # Debt/Equity ≤ 1.5 — not over-leveraged
  min_interest_coverage: 2.5     # Interest coverage ≥ 2.5x (soft — data limited)
  min_current_ratio: 0.8         # Current ratio ≥ 0.8 (liquidity floor, soft)
  max_net_debt_to_ebitda: 4.0    # Net Debt/EBITDA ≤ 4.0 (soft)
  max_working_capital_cycle_days: 200   # Working capital efficiency (soft)

  # ── 7. TECHNICAL MOMENTUM ────────────────────────────────────────
  min_rsi: 28                    # RSI ≥ 28 — avoid severely broken/distressed stocks
  max_rsi: 82                    # RSI ≤ 82 — avoid extremely overbought
  price_above_ema_200: false     # Prefer (not required) price above 200-day EMA
  price_above_ema_50: false      # Prefer price above 50-day EMA (soft preference)
  macd_bullish_preferred: true   # Prefer bullish MACD crossover (soft)

  # ── 8. PRICE RANGE FROM 52-WEEK EXTREMES ─────────────────────────
  min_above_52w_low_pct: 5       # Must be ≥ 5 % above 52W low (not at floor)
  max_below_52w_high_pct: 60     # Must be within 60 % of 52W high (not a collapse)

  # ── 9. VOLATILITY & RISK ─────────────────────────────────────────
  min_beta: 0.2                  # Beta ≥ 0.2 — some market correlation
  max_beta: 2.8                  # Beta ≤ 2.8 — not excessively volatile

  # ── 10. INSTITUTIONAL QUALITY SIGNALS ────────────────────────────
  min_institutional_holding_pct: 2     # ≥ 2 % institutional holding (basic validation)
  min_mf_holding_pct: 0.5              # ≥ 0.5 % mutual fund ownership (soft)
  max_promoter_pledge_pct: 40          # Promoter pledge ≤ 40 % (soft — screener.in data)
  min_promoter_holding_pct: 20         # Promoter holding ≥ 20 % (skin in the game, soft)

  # ── 11. BUSINESS QUALITY (SOFT / ASPIRATIONAL) ───────────────────
  min_asset_turnover: 0.2        # Asset turnover ≥ 0.2 (soft — capital efficiency)
  min_return_on_assets_pct: 5    # ROA ≥ 5 % (soft)
  min_return_on_invested_capital_pct: 10   # ROIC ≥ 10 % (soft)
  min_fcf_yield_pct: 1.0         # Free cash flow yield ≥ 1 % (soft)
  consistent_dividend: false     # Dividend history (bonus for scoring, not required)
  no_audit_qualification: true   # Clean auditor opinion preferred (soft)
  min_revenue_ttm_cr: 30         # Trailing 12-month revenue ≥ ₹30 Cr (operating business)
  sector_excluded: []            # List sectors to exclude (empty = include all)

# ─────────────────────────────────────────────────────────────────
# SCORING WEIGHTS
# Determines how survivors are ranked for the top `output_limit` cut.
# Weights are relative — higher means the metric matters more.
# Total weight ≈ 100 for normalized comparison.
# ─────────────────────────────────────────────────────────────────
scoring_weights:
  weight_roe: 15                 # Return on Equity — capital efficiency
  weight_roce: 15                # Return on Capital Employed — operational efficiency
  weight_revenue_growth: 12      # Revenue growth rate — business momentum
  weight_earnings_growth: 12     # Earnings growth rate — bottom-line expansion
  weight_debt_equity: 10         # Leverage safety — lower D/E is better
  weight_profit_margin: 10       # Net profitability quality
  weight_operating_margin: 8     # Pre-interest efficiency
  weight_rsi_momentum: 8         # Technical momentum (RSI position)
  weight_ema_trend: 8            # Price trend vs EMA stack
  weight_institutional: 5        # Institutional ownership confidence
  weight_valuation_pb: 5         # Price-to-Book attractiveness
  weight_valuation_ps: 4         # Price-to-Sales attractiveness
  weight_peg: 4                  # Growth-adjusted valuation (PEG)
  weight_beta_stability: 3       # Lower beta = more stable (bonus)
  weight_dividend: 2             # Dividend yield (bonus points)
---

# Small & Micro Cap Screener — SKILL Documentation

## What This File Does

This is the **skill configuration** for the AI Live Selection mode of the Stock Analysis Agent.
When you run:

```bash
python Stock_Agent.py --mode skill
```

The agent reads the YAML front-matter above and:

1. **Loads the candidate universe** — ~85 curated NSE small/micro cap stocks across 9 sectors
2. **Pre-screens each candidate** — applies 50+ filter rules using live yfinance data
3. **Scores survivors** — ranks them using the weighted scoring model
4. **Selects the top `output_limit`** — passes the best-ranked symbols into the full dual-track analysis pipeline

---

## Parameter Reference (50+ Criteria)

### Group 1 — Market Cap Range
| Parameter | Value | Purpose |
|-----------|-------|---------|
| `min_market_cap_cr` | ₹50 Cr | Minimum to exclude nano/illiquid shells |
| `max_market_cap_cr` | ₹5000 Cr | Maximum for true small cap boundary |

### Group 2 — Price Floor
| Parameter | Value | Purpose |
|-----------|-------|---------|
| `min_price_inr` | ₹15 | Exclude penny stocks (manipulation/liquidity risk) |

### Group 3 — Profitability (6 parameters)
| Parameter | Value | Purpose |
|-----------|-------|---------|
| `min_eps_positive` | true | Must be profitable (positive trailing EPS) |
| `min_book_value_positive` | true | Positive net tangible assets |
| `min_roe_pct` | 10% | Return on Equity floor |
| `min_roce_pct` | 10% | Return on Capital Employed floor |
| `min_profit_margin_pct` | 6% | Net margin floor |
| `min_operating_margin_pct` | 8% | Operating efficiency floor |
| `min_ebitda_margin_pct` | 10% | EBITDA margin (soft — scoring only) |

### Group 4 — Valuation (6 parameters)
| Parameter | Value | Purpose |
|-----------|-------|---------|
| `max_pe_ratio` | 45 | Avoid richly priced stocks |
| `max_price_to_book` | 7 | Premium over book value cap |
| `max_price_to_sales` | 6 | Revenue valuation sanity |
| `max_peg_ratio` | 3.0 | Growth-adjusted valuation guard |
| `max_ev_ebitda` | 25 | EV/EBITDA cap (soft) |
| `min_earnings_yield_pct` | 1.5% | Inverse PE floor (soft) |

### Group 5 — Growth (5 parameters)
| Parameter | Value | Purpose |
|-----------|-------|---------|
| `min_revenue_growth_pct` | 5% | YoY revenue expansion floor |
| `min_earnings_growth_pct` | 0% | No earnings decline |
| `min_retained_earnings_growth_pct` | 3% | Compounding reinvestment (soft) |
| `min_3y_cagr_revenue_pct` | 8% | 3-year revenue CAGR (soft) |
| `min_3y_cagr_earnings_pct` | 5% | 3-year earnings CAGR (soft) |

### Group 6 — Balance Sheet Safety (5 parameters)
| Parameter | Value | Purpose |
|-----------|-------|---------|
| `max_debt_equity` | 1.5 | Leverage ceiling |
| `min_interest_coverage` | 2.5x | Debt serviceability (soft) |
| `min_current_ratio` | 0.8 | Liquidity floor (soft) |
| `max_net_debt_to_ebitda` | 4.0 | Net leverage cap (soft) |
| `max_working_capital_cycle_days` | 200 | Working capital efficiency (soft) |

### Group 7 — Technical Momentum (5 parameters)
| Parameter | Value | Purpose |
|-----------|-------|---------|
| `min_rsi` | 28 | Avoid severely distressed/broken stocks |
| `max_rsi` | 82 | Avoid extremely overbought |
| `price_above_ema_200` | false (soft) | Long-term trend preference |
| `price_above_ema_50` | false (soft) | Medium-term trend preference |
| `macd_bullish_preferred` | true (soft) | Positive momentum preference |

### Group 8 — 52-Week Position (2 parameters)
| Parameter | Value | Purpose |
|-----------|-------|---------|
| `min_above_52w_low_pct` | 5% | Not at floor (signs of life) |
| `max_below_52w_high_pct` | 60% | Not a collapsed stock |

### Group 9 — Volatility & Risk (2 parameters)
| Parameter | Value | Purpose |
|-----------|-------|---------|
| `min_beta` | 0.2 | Requires some market correlation |
| `max_beta` | 2.8 | Caps excessive volatility |

### Group 10 — Institutional Quality (4 parameters)
| Parameter | Value | Purpose |
|-----------|-------|---------|
| `min_institutional_holding_pct` | 2% | Basic validation signal |
| `min_mf_holding_pct` | 0.5% | Mutual fund interest (soft) |
| `max_promoter_pledge_pct` | 40% | Pledging risk guard (soft) |
| `min_promoter_holding_pct` | 20% | Promoter conviction (soft) |

### Group 11 — Business Quality / Aspirational (8 parameters)
| Parameter | Value | Purpose |
|-----------|-------|---------|
| `min_asset_turnover` | 0.2 | Capital efficiency (soft) |
| `min_return_on_assets_pct` | 5% | Asset productivity (soft) |
| `min_return_on_invested_capital_pct` | 10% | ROIC floor (soft) |
| `min_fcf_yield_pct` | 1.0% | Free cash flow yield (soft) |
| `consistent_dividend` | false | Dividend history bonus (soft) |
| `no_audit_qualification` | true | Clean audit opinion (soft) |
| `min_revenue_ttm_cr` | ₹30 Cr | Operating business floor |
| `sector_excluded` | [] | Sectors to exclude (empty = all) |

---

## Scoring Weights (15 weights = 121 total points)

| Weight Key | Points | What It Rewards |
|-----------|--------|-----------------|
| `weight_roe` | 15 | High return on equity |
| `weight_roce` | 15 | Efficient capital deployment |
| `weight_revenue_growth` | 12 | Fast-growing revenues |
| `weight_earnings_growth` | 12 | Growing bottom line |
| `weight_debt_equity` | 10 | Low leverage |
| `weight_profit_margin` | 10 | Healthy net margins |
| `weight_operating_margin` | 8 | Pre-interest efficiency |
| `weight_rsi_momentum` | 8 | Technical momentum (RSI 35–65) |
| `weight_ema_trend` | 8 | Price above EMA stack |
| `weight_institutional` | 5 | Institutional ownership |
| `weight_valuation_pb` | 5 | Low P/B ratio |
| `weight_valuation_ps` | 4 | Low P/S ratio |
| `weight_peg` | 4 | Low PEG ratio |
| `weight_beta_stability` | 3 | Lower beta stability bonus |
| `weight_dividend` | 2 | Dividend yield bonus |

---

## Customising the Screener

### Change the output count
Edit `output_limit` in the YAML front-matter (e.g. `40` for more candidates).

### Tighten / relax filters
Change threshold values in the `filters` section. For example, to require stronger ROE:
```yaml
min_roe_pct: 15
```

### Add / remove candidate stocks
Edit the `candidate_universe` list. Use NSE symbol format: `SYMBOL.NS`.

### Exclude sectors
```yaml
sector_excluded:
  - "Financial Services"
  - "Energy"
```
*(Note: sector exclusion requires sector data from yfinance, applied during full analysis.)*

---

## Data Availability Notes

- **Hard-filtered (live data):** market cap, price, PE, ROE, debt/equity, revenue growth, earnings growth, profit margin, beta, P/B, P/S, PEG, EPS, institutional holding
- **Soft / scoring only:** ROCE, EBITDA margin, interest coverage, current ratio, working capital, 3Y CAGR, promoter holding/pledge, ROIC, FCF yield, audit quality
- **Fetched in full analysis pass:** RSI, MACD, EMA stack, 52W high/low, Rule-Based Score, AI Score

Soft parameters that are unavailable for a specific stock are skipped (not penalised), so the screener degrades gracefully.
