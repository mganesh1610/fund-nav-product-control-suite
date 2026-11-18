# Accounting Standards, Internal Controls & Relational Schema Catalog

## 1. Governing Accounting & Regulatory Frameworks Modeled

1. **ASC 946 (Investment Companies — Fund Accounting & NAV Packages):**
   - Monthly and quarterly Net Asset Value (`NAV`) roll-forwards across multi-currency private equity, private credit, infrastructure, and liquid co-investment funds.
   - Pro-rata LP/GP Investor Capital Statements conserving `100.00%` partnership equity and `$0.00` waterfall-to-NAV rounding drift via deterministic largest-remainder cent allocation.
2. **ASC 820 (Fair Value Measurement & Independent Price Verification — Product Control & Fund Control):**
   - Categorizes fund and trading desk positions across `Level 1`, `Level 2`, and `Level 3` fair-value hierarchies.
   - Enforces a `10.0 bps` Independent Price Verification (`IPV`) tolerance comparing front-office trader marks against sovereign U.S. Treasury par-yield curve benchmarks and European Central Bank (`ECB`) daily FX spot fixings. Stale or off-market marks exceeding `10.0 bps` automatically post an `IPV Valuation Adjustment` prior to back-office General Ledger (`GL`) close.
3. **ASC 830 / IAS 21 (Foreign Currency Matters):**
   - Separates local-currency rate/duration mark-to-market (`MTM`) movements from closing spot foreign-exchange (`FX`) translation adjustments across `EUR`, `GBP`, `CHF`, `JPY`, and `USD`.
4. **Asset Management Fund Controller Policy Controls:**
   - Validates minimum `3.0%` unencumbered cash liquidity buffers, `<= 200 bps` annualized management fee policy caps, and `ACT/365` / monthly fee & administrative expense accruals.

## 2. Relational Table Schemas (`outputs/controllers_fund_and_pnl.sqlite`)

| Table Name | Primary Key | Row Count | Grain & Purpose |
| :--- | :--- | :--- | :--- |
| `fct_fund_nav_close` | `(close_month, fund_id)` | `96` (`32` quarterly) | Monthly & quarterly fund-level NAV roll-forward, MTM valuation, FX translation, management fee/admin accruals, and policy compliance flags across 4 funds (`2024-01` to `2025-12`). |
| `fct_investor_capital_statements` | `(close_month, fund_id, investor_id)` | `384` | Partner-level LP/GP capital account statements reconciling beginning capital, capital calls, distributions, income, MTM, FX, and fee accruals to ending capital (`$0.00` variance to Fund NAV). |
| `fct_daily_product_pnl_ipv` | `(trade_date, position_id)` | `2,952` | Daily trading-desk P&L attribution (`Carry`, `Rate DV01/Convexity`, `FX Reval`, `Unexplained Residual`), front-office vs. independent market IPV checks, GL adjustments, and DV01 risk metrics across 6 positions and 493 aligned trading days. |
