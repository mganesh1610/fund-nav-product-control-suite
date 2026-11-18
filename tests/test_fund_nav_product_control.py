#!/usr/bin/env python3
"""Automated unit & accounting-control boundary tests for fund-nav-product-control-suite."""

from __future__ import annotations

from pathlib import Path
import sqlite3
import sys
import unittest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from fund_nav_and_pnl_engine import run_controllers_suite  # noqa: E402


class TestFundNavProductControlSuite(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.summary = run_controllers_suite(ROOT)
        cls.db_path = ROOT / "outputs" / "controllers_fund_and_pnl.sqlite"

    def test_01_all_14_internal_control_gates_pass(self) -> None:
        gates = self.summary["internal_control_gates"]
        self.assertEqual(gates["passed"], 14)
        self.assertEqual(gates["total"], 14)
        self.assertEqual(gates["pass_rate_pct"], 100.0)

    def test_02_fund_nav_rollforward_exact_cent_conservation(self) -> None:
        fa = self.summary["fund_accounting_metrics"]
        self.assertEqual(fa["monthly_close_packages"], 96)
        self.assertEqual(fa["quarterly_reporting_packages"], 32)
        self.assertEqual(fa["investor_capital_statements_generated"], 384)
        self.assertEqual(fa["nav_rollforward_max_variance_usd"], 0.0)
        self.assertEqual(fa["investor_waterfall_to_nav_variance_usd"], 0.0)

    def test_03_usd_fund_has_zero_fx_translation_adjustment(self) -> None:
        conn = sqlite3.connect(str(self.db_path))
        non_zero = conn.execute(
            "SELECT COUNT(*) FROM fct_fund_nav_close WHERE base_currency = 'USD' AND ABS(fx_translation_adj_usd) > 0.0"
        ).fetchone()[0]
        conn.close()
        self.assertEqual(non_zero, 0)

    def test_04_asset_management_policy_compliance(self) -> None:
        fa = self.summary["fund_accounting_metrics"]
        self.assertEqual(fa["asset_management_policy_breaches"], 0)

    def test_05_daily_product_pnl_attribution_exact_identity(self) -> None:
        pc = self.summary["product_controllers_metrics"]
        self.assertGreaterEqual(pc["daily_position_valuations"], 2700)
        self.assertEqual(pc["pnl_decomposition_max_variance_usd"], 0.0)
        self.assertGreaterEqual(pc["pnl_attribution_explained_ratio_pct"], 98.0)

    def test_06_independent_price_verification_intercepts_all_stale_marks(self) -> None:
        pc = self.summary["product_controllers_metrics"]
        self.assertGreater(pc["ipv_stale_mark_exceptions_intercepted"], 20)
        self.assertEqual(pc["ipv_unremediated_gl_breaches"], 0)

    def test_07_excel_close_workbook_and_csv_artifacts_exist(self) -> None:
        xlsx_path = ROOT / "outputs" / "fund_nav_and_daily_pnl_close_package.xlsx"
        self.assertTrue(xlsx_path.is_file())
        self.assertGreater(xlsx_path.stat().st_size, 10000)


if __name__ == "__main__":
    unittest.main()
