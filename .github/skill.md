# Live Selection Query

Use this file for live AI symbol discovery. The engine auto-discovers NSE symbols,
filters them using the named queries below, and picks the top N (configured by `live_top_n`).
Each query is evaluated independently; a symbol qualifies based on its best pass-ratio
across all queries.

## Discovery Rules
- universe_scan_limit: 5000

## Query: QUALITY_SMALL_MICRO
- market_cap_gt_500
- market_cap_lt_18000
- is_small_or_micro_cap
- debt_lt_0_5
- roe_gt_15
- roce_gt_18
- sales_growth_3y_gt_12
- profit_growth_3y_gt_12
- operating_margin_gt_12
- interest_coverage_gt_4
- cash_from_operations_positive
- promoter_holding_gt_50
- pledged_percentage_lt_3
- pb_lt_4
- data_completeness_gt_75

## Query: GROWTH_SMALL_MICRO
- market_cap_gt_500
- market_cap_lt_18000
- is_small_or_micro_cap
- debt_lt_0_8
- roe_gt_15
- roce_gt_15
- revenue_growth_gt_12
- earnings_growth_gt_12
- sales_growth_3y_gt_12
- profit_growth_3y_gt_12
- cash_from_operations_positive
- promoter_holding_gt_50
- pledged_percentage_lt_3
- peg_lt_1_5
- price_above_ema50
- rsi_40_70
- data_completeness_gt_75

## Query: MOMENTUM_QUALITY_SMALL_MICRO
- market_cap_gt_500
- market_cap_lt_18000
- is_small_or_micro_cap
- debt_lt_0_5
- roe_gt_15
- roce_gt_18
- sales_growth_3y_gt_12
- profit_growth_3y_gt_12
- cash_from_operations_positive
- promoter_holding_gt_50
- pledged_percentage_lt_3
- price_above_ema50
- price_above_ema200
- macd_bullish
- rsi_40_70
- within_20pct_of_52w_high
- data_completeness_gt_75

## Query: BEST_SMALL_MICRO_CORE
- market_cap_gt_500
- market_cap_lt_18000
- is_small_or_micro_cap
- debt_lt_0_5
- roe_gt_15
- roce_gt_18
- sales_growth_3y_gt_12
- profit_growth_3y_gt_12
- cash_from_operations_positive
- interest_coverage_gt_4
- promoter_holding_gt_50
- pledged_percentage_lt_3
- pe_lt_35
- price_above_ema50
- rsi_40_70
- data_completeness_gt_75

## Notes
- Each query name maps to a list of parameter keys evaluated via LIVE_PARAMETER_CHECKS.
- A symbol's best pass-ratio across all queries determines its ranking.
- If live selection fails and fallback is enabled, predefined symbols from config are used.
