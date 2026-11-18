-- Fund Accounting NAV Roll-Forward, Investor Capital Waterfall & Product Control Daily P&L SQL Queries

-- 1. Fund-Level Period-End Close NAV Roll-Forward Parity Check ($0.00 Variance)
WITH nav_check AS (
    SELECT
        close_month,
        quarter,
        fund_id,
        base_currency,
        beginning_nav_usd,
        capital_calls_usd,
        distributions_usd,
        investment_income_usd,
        mtm_valuation_adj_usd,
        fx_translation_adj_usd,
        total_accruals_usd,
        ending_nav_usd,
        ROUND(
            (beginning_nav_usd + capital_calls_usd - distributions_usd + investment_income_usd
             + mtm_valuation_adj_usd + fx_translation_adj_usd - total_accruals_usd) - ending_nav_usd,
            2
        ) AS nav_rollforward_variance_usd
    FROM fct_fund_nav_close
)
SELECT * FROM nav_check WHERE ABS(nav_rollforward_variance_usd) > 0.00;

-- 2. LP / GP Investor Capital Statement Waterfall vs. Fund Ending NAV Reconciliation
WITH partner_rollup AS (
    SELECT
        close_month,
        fund_id,
        COUNT(DISTINCT investor_id) AS partner_count,
        ROUND(SUM(ownership_pct), 2) AS total_ownership_pct,
        ROUND(SUM(ending_capital_usd), 2) AS sum_investor_ending_capital_usd
    FROM fct_investor_capital_statements
    GROUP BY close_month, fund_id
)
SELECT
    f.close_month,
    f.fund_id,
    p.partner_count,
    p.total_ownership_pct,
    f.ending_nav_usd,
    p.sum_investor_ending_capital_usd,
    ROUND(f.ending_nav_usd - p.sum_investor_ending_capital_usd, 2) AS waterfall_variance_usd
FROM fct_fund_nav_close f
INNER JOIN partner_rollup p
    ON f.close_month = p.close_month
   AND f.fund_id = p.fund_id
ORDER BY f.close_month, f.fund_id;

-- 3. Product Controllers Daily P&L Attribution & Independent Price Verification (IPV) Audit
SELECT
    desk_name,
    currency,
    COUNT(*) AS position_days,
    ROUND(SUM(carry_accrual_pnl_usd), 2) AS cumulative_carry_pnl_usd,
    ROUND(SUM(rate_move_pnl_usd), 2) AS cumulative_rate_move_pnl_usd,
    ROUND(SUM(fx_reval_pnl_usd), 2) AS cumulative_fx_reval_pnl_usd,
    ROUND(SUM(unexplained_pnl_usd), 2) AS cumulative_unexplained_pnl_usd,
    ROUND(SUM(total_daily_pnl_usd), 2) AS cumulative_total_pnl_usd,
    SUM(ipv_exception_flag) AS ipv_stale_marks_flagged,
    ROUND(SUM(ABS(ipv_valuation_adj_usd)), 2) AS total_ipv_gl_adjustments_usd,
    ROUND(MAX(dv01_risk_usd), 2) AS peak_position_dv01_usd
FROM fct_daily_product_pnl_ipv
GROUP BY desk_name, currency
ORDER BY desk_name, currency;
