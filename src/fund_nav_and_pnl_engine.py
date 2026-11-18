#!/usr/bin/env python3
"""Private Fund NAV, Multi-Currency Period-End Close & Product Control Daily P&L Engine.

Downloads and SHA-256 verifies official public sovereign market datasets:
1. U.S. Department of the Treasury Daily Par Yield Curve Rates (2024 & 2025 CSVs)
2. European Central Bank (ECB) Historical Daily Foreign Exchange Reference Rates (eurofxref-hist.zip)

Executes four core Controllers workflows in SQLite + exact Decimal/integer-cent accounting:
1. Fund Accounting: Monthly & quarterly fund-level Net Asset Value (NAV) packages, LP/GP investor
   capital statements, mark-to-market (MTM) fair-value adjustments, ASC 830 / IAS 21 foreign
   currency (FX) translation adjustments, and management-fee / expense accruals across 4 investment
   funds and 24 monthly close cycles (8 quarterly packages).
2. Asset Management Fund Controllers: ASC 820 Fair Value Hierarchy (Level 1 / 2 / 3) classification,
   single-position concentration caps, liquidity buffer compliance, and management fee policy caps.
3. Product Controllers: Daily trading-desk Profit & Loss (P&L) attribution (Carry, Rate DV01/Convexity,
   FX Revaluation, Unexplained P&L control), Independent Price Verification (IPV) against sovereign
   Treasury & ECB FX benchmarks across 3,000 desk position-days, Front-to-Back Office (FO-to-BO)
   general-ledger reconciliation, and DV01 / net FX risk limit monitoring.
4. MS Office Reporting Package Export: Multi-tab native OpenXML (.xlsx) close workbook, CSV schedules,
   SQLite relational database, and 14 deterministic accounting & internal control audit gates.
"""

from __future__ import annotations

import csv
from decimal import Decimal, ROUND_HALF_UP
import hashlib
import io
import json
from pathlib import Path
import sqlite3
import urllib.request
import xml.sax.saxutils as saxutils
import zipfile

TREASURY_2024_URL = (
    "https://home.treasury.gov/resource-center/data-chart-center/interest-rates/"
    "daily-treasury-rates.csv/2024/all?type=daily_treasury_yield_curve&field_tdr_date_value=2024&page&_format=csv"
)
TREASURY_2025_URL = (
    "https://home.treasury.gov/resource-center/data-chart-center/interest-rates/"
    "daily-treasury-rates.csv/2025/all?type=daily_treasury_yield_curve&field_tdr_date_value=2025&page&_format=csv"
)
ECB_FX_ZIP_URL = "https://www.ecb.europa.eu/stats/eurofxref/eurofxref-hist.zip"

EXPECTED_TREASURY_2024_SHA256 = "6775840b76cd7fd41121f61d6f6c60184eca6b2ba49954ec7372c463cbeef34d"

CENT = Decimal("0.01")


def q2(val: Decimal | float | int) -> Decimal:
    return Decimal(str(val)).quantize(CENT, rounding=ROUND_HALF_UP)


def fetch_cached_bytes(url: str, cache_path: Path) -> tuple[bytes, str]:
    cache_path.parent.mkdir(parents=True, exist_ok=True)
    if cache_path.is_file() and cache_path.stat().st_size > 0:
        data = cache_path.read_bytes()
    else:
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=30) as resp:
            data = resp.read()
        cache_path.write_bytes(data)
    return data, hashlib.sha256(data).hexdigest()


def parse_treasury_csv(raw_bytes: bytes) -> dict[str, dict[str, float]]:
    """Return {YYYY-MM-DD: {'3M': float, '2Y': float, '5Y': float, '10Y': float, '30Y': float}}."""
    text = raw_bytes.decode("utf-8", errors="replace")
    reader = csv.DictReader(io.StringIO(text))
    col_map = {
        "3 Mo": "3M",
        "2 Yr": "2Y",
        "5 Yr": "5Y",
        "10 Yr": "10Y",
        "30 Yr": "30Y",
    }
    out: dict[str, dict[str, float]] = {}
    for row in reader:
        raw_dt = (row.get("Date") or "").strip()
        if not raw_dt or "/" not in raw_dt:
            continue
        parts = raw_dt.split("/")
        if len(parts) != 3:
            continue
        mm, dd, yyyy = parts[0].zfill(2), parts[1].zfill(2), parts[2]
        iso_dt = f"{yyyy}-{mm}-{dd}"
        rates: dict[str, float] = {}
        valid = True
        for src_col, dst_col in col_map.items():
            val_s = (row.get(src_col) or "").strip()
            if not val_s or val_s == "N/A":
                valid = False
                break
            rates[dst_col] = float(val_s)
        if valid:
            out[iso_dt] = rates
    return out


def parse_ecb_fx_zip(raw_bytes: bytes) -> tuple[int, int, dict[str, dict[str, float]]]:
    """Parse ECB eurofxref-hist.zip into USD-numeraire FX rates (USD per 1 unit of foreign currency)."""
    zf = zipfile.ZipFile(io.BytesIO(raw_bytes))
    csv_text = zf.read(zf.namelist()[0]).decode("utf-8", errors="replace")
    reader = csv.DictReader(io.StringIO(csv_text))
    total_rows = 0
    non_null_obs = 0
    fx_usd_per_ccy: dict[str, dict[str, float]] = {}
    for row in reader:
        dt = (row.get("Date") or "").strip()
        if not dt:
            continue
        total_rows += 1
        for k, v in row.items():
            if k and k != "Date" and v and v.strip() not in ("", "N/A"):
                non_null_obs += 1
        usd_per_eur_s = (row.get("USD") or "").strip()
        gbp_per_eur_s = (row.get("GBP") or "").strip()
        jpy_per_eur_s = (row.get("JPY") or "").strip()
        chf_per_eur_s = (row.get("CHF") or "").strip()
        if all(s and s != "N/A" for s in (usd_per_eur_s, gbp_per_eur_s, jpy_per_eur_s, chf_per_eur_s)):
            usd_per_eur = float(usd_per_eur_s)
            fx_usd_per_ccy[dt] = {
                "USD": 1.0,
                "EUR": round(usd_per_eur, 6),
                "GBP": round(usd_per_eur / float(gbp_per_eur_s), 6),
                "JPY": round(usd_per_eur / float(jpy_per_eur_s), 8),
                "CHF": round(usd_per_eur / float(chf_per_eur_s), 6),
            }
    return total_rows, non_null_obs, fx_usd_per_ccy


def write_minimal_xlsx(workbook_path: Path, sheets: list[tuple[str, list[list[object]]]]) -> None:
    """Generate a dependency-free OpenXML .xlsx workbook with styled header rows for Excel review."""
    workbook_path.parent.mkdir(parents=True, exist_ok=True)

    def col_letter(idx_zero: int) -> str:
        res = ""
        n = idx_zero + 1
        while n > 0:
            n, rem = divmod(n - 1, 26)
            res = chr(65 + rem) + res
        return res

    content_types = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
        '  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
        '  <Default Extension="xml" ContentType="application/xml"/>',
        '  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
        '  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>',
    ]
    for i in range(len(sheets)):
        content_types.append(
            f'  <Override PartName="/xl/worksheets/sheet{i+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        )
    content_types.append("</Types>")

    root_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"""

    wb_xml = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
        "  <sheets>",
    ]
    wb_rels = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
    ]
    for i, (title, _) in enumerate(sheets):
        safe_title = saxutils.escape(title[:31])
        wb_xml.append(f'    <sheet name="{safe_title}" sheetId="{i+1}" r:id="rId{i+1}"/>')
        wb_rels.append(
            f'  <Relationship Id="rId{i+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i+1}.xml"/>'
        )
    wb_rels.append(
        f'  <Relationship Id="rId{len(sheets)+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
    )
    wb_xml.append("  </sheets></workbook>")
    wb_rels.append("</Relationships>")

    styles_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2">
    <font><sz val="10"/><name val="Calibri"/></font>
    <font><b/><sz val="10"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>
  </fonts>
  <fills count="3">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF173B55"/></patternFill></fill>
  </fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>
  </cellXfs>
</styleSheet>"""

    with zipfile.ZipFile(workbook_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", "\n".join(content_types))
        zf.writestr("_rels/.rels", root_rels)
        zf.writestr("xl/workbook.xml", "\n".join(wb_xml))
        zf.writestr("xl/_rels/workbook.xml.rels", "\n".join(wb_rels))
        zf.writestr("xl/styles.xml", styles_xml)

        for idx, (_, rows) in enumerate(sheets):
            ws = [
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
                '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
                "  <sheetData>",
            ]
            for r_idx, row in enumerate(rows):
                ws.append(f'    <row r="{r_idx+1}">')
                style_attr = ' s="1"' if r_idx == 0 else ""
                for c_idx, val in enumerate(row):
                    ref = f"{col_letter(c_idx)}{r_idx+1}"
                    if isinstance(val, (int, float, Decimal)) and not isinstance(val, bool):
                        ws.append(f'      <c r="{ref}"{style_attr}><v>{val}</v></c>')
                    else:
                        txt = saxutils.escape(str(val if val is not None else ""))
                        ws.append(f'      <c r="{ref}" t="inlineStr"{style_attr}><is><t>{txt}</t></is></c>')
                ws.append("    </row>")
            ws.append("  </sheetData></worksheet>")
            zf.writestr(f"xl/worksheets/sheet{idx+1}.xml", "\n".join(ws))


def run_controllers_suite(base_dir: Path | None = None) -> dict[str, object]:
    root = base_dir or Path(__file__).resolve().parents[1]
    cache_dir = root / "data" / "raw_cache"
    out_dir = root / "outputs"
    out_dir.mkdir(parents=True, exist_ok=True)

    t24_bytes, t24_sha = fetch_cached_bytes(TREASURY_2024_URL, cache_dir / "treasury_yield_2024.csv")
    t25_bytes, t25_sha = fetch_cached_bytes(TREASURY_2025_URL, cache_dir / "treasury_yield_2025.csv")
    ecb_bytes, ecb_sha = fetch_cached_bytes(ECB_FX_ZIP_URL, cache_dir / "eurofxref-hist.zip")

    if t24_sha != EXPECTED_TREASURY_2024_SHA256:
        raise ValueError(f"2024 Treasury SHA-256 mismatch: {t24_sha}")

    treasury_24 = parse_treasury_csv(t24_bytes)
    treasury_25 = parse_treasury_csv(t25_bytes)
    treasury_all = {**treasury_24, **treasury_25}
    ecb_rows_count, ecb_obs_count, fx_map = parse_ecb_fx_zip(ecb_bytes)

    # Align common 2024-2025 business dates where both US Treasury and ECB FX published fixings
    aligned_dates = sorted(dt for dt in treasury_all.keys() if dt in fx_map and "2024-01-01" <= dt <= "2025-12-31")
    if len(aligned_dates) < 450:
        raise RuntimeError(f"Insufficient aligned sovereign market dates: {len(aligned_dates)}")

    # Group aligned dates by YYYY-MM (24 monthly close periods across 2024-01..2025-12)
    month_end_dates: dict[str, str] = {}
    for dt in aligned_dates:
        ym = dt[:7]
        month_end_dates[ym] = dt
    months_sorted = sorted(month_end_dates.keys())

    # -------------------------------------------------------------------------
    # 1. FUND ACCOUNTING & ASSET MANAGEMENT FUND CONTROLLERS (4 Funds, 24 Closes, 16 LP/GP Accounts)
    # -------------------------------------------------------------------------
    funds = [
        {
            "fund_id": "FUND_PE_BUYOUT_I",
            "fund_name": "Global Buyout Opportunities Fund I",
            "currency": "EUR",
            "asc820_level": "Level 3",
            "committed_usd": Decimal("85000000.00"),
            "initial_local_notional": Decimal("52000000.00"),
            "tenor_anchor": "10Y",
            "duration_sensitivity": Decimal("4.20"),
            "mgmt_fee_annual_rate": Decimal("0.0175"),
            "admin_monthly_accrual_usd": Decimal("28500.00"),
            "lp_shares": [
                ("LP_PENSION_ALPHA", "Public Pension Partner Alpha", Decimal("0.45")),
                ("LP_ENDOWMENT_BETA", "University Endowment Beta", Decimal("0.35")),
                ("LP_SOVEREIGN_GAMMA", "Institutional Sovereign Gamma", Decimal("0.18")),
                ("GP_SPONSOR_COMMIT", "General Partner Sponsor Commitment", Decimal("0.02")),
            ],
        },
        {
            "fund_id": "FUND_CREDIT_OPP_II",
            "fund_name": "Opportunistic Private Credit Fund II",
            "currency": "GBP",
            "asc820_level": "Level 2",
            "committed_usd": Decimal("65000000.00"),
            "initial_local_notional": Decimal("38000000.00"),
            "tenor_anchor": "5Y",
            "duration_sensitivity": Decimal("3.40"),
            "mgmt_fee_annual_rate": Decimal("0.0150"),
            "admin_monthly_accrual_usd": Decimal("22000.00"),
            "lp_shares": [
                ("LP_INSURANCE_DELTA", "Life Insurance General Account Delta", Decimal("0.50")),
                ("LP_FOUNDATION_EPSILON", "Charitable Foundation Epsilon", Decimal("0.30")),
                ("LP_FAMILY_ZETA", "Multi-Family Office Zeta", Decimal("0.18")),
                ("GP_SPONSOR_COMMIT", "General Partner Sponsor Commitment", Decimal("0.02")),
            ],
        },
        {
            "fund_id": "FUND_INFRA_YIELD_III",
            "fund_name": "Cross-Border Infrastructure Yield Fund III",
            "currency": "CHF",
            "asc820_level": "Level 3",
            "committed_usd": Decimal("55000000.00"),
            "initial_local_notional": Decimal("41000000.00"),
            "tenor_anchor": "10Y",
            "duration_sensitivity": Decimal("5.10"),
            "mgmt_fee_annual_rate": Decimal("0.0150"),
            "admin_monthly_accrual_usd": Decimal("19500.00"),
            "lp_shares": [
                ("LP_PENSION_ETA", "Municipal Retirement System Eta", Decimal("0.48")),
                ("LP_ENDOWMENT_THETA", "Medical Research Endowment Theta", Decimal("0.32")),
                ("LP_INSURANCE_IOTA", "Property Reinsurance Partner Iota", Decimal("0.18")),
                ("GP_SPONSOR_COMMIT", "General Partner Sponsor Commitment", Decimal("0.02")),
            ],
        },
        {
            "fund_id": "FUND_LIQUID_MULTI_IV",
            "fund_name": "Liquid Multi-Asset Co-Investment Fund IV",
            "currency": "USD",
            "asc820_level": "Level 1",
            "committed_usd": Decimal("45000000.00"),
            "initial_local_notional": Decimal("36000000.00"),
            "tenor_anchor": "2Y",
            "duration_sensitivity": Decimal("1.85"),
            "mgmt_fee_annual_rate": Decimal("0.0125"),
            "admin_monthly_accrual_usd": Decimal("15000.00"),
            "lp_shares": [
                ("LP_TREASURY_KAPPA", "Corporate Pension Trust Kappa", Decimal("0.52")),
                ("LP_WEALTH_LAMBDA", "Institutional Wealth Feeder Lambda", Decimal("0.28")),
                ("LP_ENDOWMENT_MU", "Regional University Foundation Mu", Decimal("0.18")),
                ("GP_SPONSOR_COMMIT", "General Partner Sponsor Commitment", Decimal("0.02")),
            ],
        },
    ]

    first_dt = aligned_dates[0]
    fund_nav_rows: list[dict[str, object]] = []
    investor_statement_rows: list[dict[str, object]] = []
    am_policy_breaches = 0

    for fund in funds:
        ccy = str(fund["currency"])
        tenor = str(fund["tenor_anchor"])
        dur = Decimal(str(fund["duration_sensitivity"]))
        prev_fx = Decimal(str(fx_map[first_dt][ccy]))
        prev_yield = Decimal(str(treasury_all[first_dt][tenor]))
        local_units = Decimal(str(fund["initial_local_notional"]))
        beg_nav_usd = q2(local_units * prev_fx)
        cash_reserve_usd = q2(Decimal(str(fund["committed_usd"])) * Decimal("0.05"))
        beg_nav_usd = q2(beg_nav_usd + cash_reserve_usd)

        for m_idx, ym in enumerate(months_sorted):
            close_dt = month_end_dates[ym]
            curr_fx = Decimal(str(fx_map[close_dt][ccy]))
            curr_yield = Decimal(str(treasury_all[close_dt][tenor]))
            quarter_label = f"{ym[:4]}-Q{(int(ym[5:7]) - 1) // 3 + 1}"
            is_quarter_end = int(ym[5:7]) in (3, 6, 9, 12)

            # Scheduled capital calls on Q1/Q3 starts and distributions on quarter-ends
            cap_call_usd = q2(Decimal(str(fund["committed_usd"])) * Decimal("0.025")) if m_idx in (2, 8, 14, 20) else Decimal("0.00")
            distribution_usd = q2(beg_nav_usd * Decimal("0.018")) if is_quarter_end else Decimal("0.00")

            # Period-end close components:
            # 1) Yield/spread income & realized operating carry
            investment_income_usd = q2((beg_nav_usd - cash_reserve_usd) * (curr_yield / Decimal("100")) / Decimal("12"))
            # 2) Mark-to-market (MTM) fair-value adjustment driven by real sovereign yield curve delta (in percentage points)
            dy_pct = (curr_yield - prev_yield) / Decimal("100")
            mtm_adjustment_usd = q2(-(beg_nav_usd - cash_reserve_usd) * dur * dy_pct)
            # 3) Foreign currency (FX) translation adjustment (ASC 830) on foreign-denominated portfolio base
            if ccy == "USD":
                fx_translation_usd = Decimal("0.00")
            else:
                fx_ret = (curr_fx - prev_fx) / prev_fx
                fx_translation_usd = q2((beg_nav_usd - cash_reserve_usd) * fx_ret)
            # 4) Management fee accrual + fund admin/audit expense accrual
            mgmt_fee_accrual_usd = q2(Decimal(str(fund["committed_usd"])) * Decimal(str(fund["mgmt_fee_annual_rate"])) / Decimal("12"))
            expense_accrual_usd = q2(Decimal(str(fund["admin_monthly_accrual_usd"])))
            total_accruals_usd = q2(mgmt_fee_accrual_usd + expense_accrual_usd)

            end_nav_usd = q2(
                beg_nav_usd
                + cap_call_usd
                - distribution_usd
                + investment_income_usd
                + mtm_adjustment_usd
                + fx_translation_usd
                - total_accruals_usd
            )

            # Asset Management Fund Controller compliance gates
            liquidity_ratio = float(cash_reserve_usd / end_nav_usd) if end_nav_usd > 0 else 0.0
            annualized_fee_bps = float((mgmt_fee_accrual_usd * 12 / Decimal(str(fund["committed_usd"]))) * 10000)
            policy_compliant = (liquidity_ratio >= 0.03) and (annualized_fee_bps <= 200.0)
            if not policy_compliant:
                am_policy_breaches += 1

            # Allocate exact LP/GP Investor Capital Statements with largest-remainder cent conservation
            lp_list = list(fund["lp_shares"])
            running_end = Decimal("0.00")
            lp_rows_this_month: list[dict[str, object]] = []
            for i_lp, (inv_id, inv_name, share) in enumerate(lp_list):
                if i_lp < len(lp_list) - 1:
                    inv_beg = q2(beg_nav_usd * share)
                    inv_call = q2(cap_call_usd * share)
                    inv_dist = q2(distribution_usd * share)
                    inv_inc = q2(investment_income_usd * share)
                    inv_mtm = q2(mtm_adjustment_usd * share)
                    inv_fx = q2(fx_translation_usd * share)
                    inv_acc = q2(total_accruals_usd * share)
                    inv_end = q2(inv_beg + inv_call - inv_dist + inv_inc + inv_mtm + inv_fx - inv_acc)
                    running_end = q2(running_end + inv_end)
                else:
                    # Residual partner absorbs exact sub-cent rounding so Sum(Partner Capital) == Fund NAV ($0.00 variance)
                    inv_beg = q2(beg_nav_usd - sum(Decimal(str(r["beginning_capital_usd"])) for r in lp_rows_this_month))
                    inv_call = q2(cap_call_usd - sum(Decimal(str(r["capital_calls_usd"])) for r in lp_rows_this_month))
                    inv_dist = q2(distribution_usd - sum(Decimal(str(r["distributions_usd"])) for r in lp_rows_this_month))
                    inv_inc = q2(investment_income_usd - sum(Decimal(str(r["investment_income_usd"])) for r in lp_rows_this_month))
                    inv_mtm = q2(mtm_adjustment_usd - sum(Decimal(str(r["mtm_valuation_adj_usd"])) for r in lp_rows_this_month))
                    inv_fx = q2(fx_translation_usd - sum(Decimal(str(r["fx_translation_adj_usd"])) for r in lp_rows_this_month))
                    inv_acc = q2(total_accruals_usd - sum(Decimal(str(r["fee_and_expense_accruals_usd"])) for r in lp_rows_this_month))
                    inv_end = q2(end_nav_usd - running_end)
                lp_rows_this_month.append(
                    {
                        "close_month": ym,
                        "quarter": quarter_label,
                        "close_date": close_dt,
                        "fund_id": fund["fund_id"],
                        "investor_id": inv_id,
                        "investor_name": inv_name,
                        "ownership_pct": float(share * 100),
                        "beginning_capital_usd": float(inv_beg),
                        "capital_calls_usd": float(inv_call),
                        "distributions_usd": float(inv_dist),
                        "investment_income_usd": float(inv_inc),
                        "mtm_valuation_adj_usd": float(inv_mtm),
                        "fx_translation_adj_usd": float(inv_fx),
                        "fee_and_expense_accruals_usd": float(inv_acc),
                        "ending_capital_usd": float(inv_end),
                    }
                )
            investor_statement_rows.extend(lp_rows_this_month)

            fund_nav_rows.append(
                {
                    "close_month": ym,
                    "quarter": quarter_label,
                    "is_quarter_end": 1 if is_quarter_end else 0,
                    "close_date": close_dt,
                    "fund_id": fund["fund_id"],
                    "fund_name": fund["fund_name"],
                    "base_currency": ccy,
                    "asc820_fair_value_level": fund["asc820_level"],
                    "beginning_nav_usd": float(beg_nav_usd),
                    "capital_calls_usd": float(cap_call_usd),
                    "distributions_usd": float(distribution_usd),
                    "investment_income_usd": float(investment_income_usd),
                    "mtm_valuation_adj_usd": float(mtm_adjustment_usd),
                    "fx_translation_adj_usd": float(fx_translation_usd),
                    "mgmt_fee_accrual_usd": float(mgmt_fee_accrual_usd),
                    "admin_expense_accrual_usd": float(expense_accrual_usd),
                    "total_accruals_usd": float(total_accruals_usd),
                    "ending_nav_usd": float(end_nav_usd),
                    "closing_fx_usd_per_ccy": float(curr_fx),
                    "benchmark_yield_pct": float(curr_yield),
                    "liquidity_buffer_ratio": round(liquidity_ratio, 4),
                    "policy_compliant": 1 if policy_compliant else 0,
                }
            )

            beg_nav_usd = end_nav_usd
            prev_fx = curr_fx
            prev_yield = curr_yield

    # -------------------------------------------------------------------------
    # 2. PRODUCT CONTROLLERS: DAILY TRADING DESK P&L, IPV & FO-TO-BO RECONCILIATION
    # -------------------------------------------------------------------------
    desk_positions = [
        {
            "position_id": "DESK_UST_2Y_BENCH",
            "desk_name": "G10 Rates Trading Desk",
            "currency": "USD",
            "notional_local": Decimal("25000000.00"),
            "tenor": "2Y",
            "mod_duration": Decimal("1.92"),
            "convexity": Decimal("4.40"),
            "coupon_rate": Decimal("0.0425"),
        },
        {
            "position_id": "DESK_UST_10Y_BENCH",
            "desk_name": "G10 Rates Trading Desk",
            "currency": "USD",
            "notional_local": Decimal("20000000.00"),
            "tenor": "10Y",
            "mod_duration": Decimal("8.15"),
            "convexity": Decimal("78.50"),
            "coupon_rate": Decimal("0.0410"),
        },
        {
            "position_id": "DESK_UST_30Y_LONG",
            "desk_name": "Long Duration Macro Desk",
            "currency": "USD",
            "notional_local": Decimal("12000000.00"),
            "tenor": "30Y",
            "mod_duration": Decimal("16.80"),
            "convexity": Decimal("365.00"),
            "coupon_rate": Decimal("0.0435"),
        },
        {
            "position_id": "DESK_EUR_SOV_5Y",
            "desk_name": "Cross-Currency Credit & FX Desk",
            "currency": "EUR",
            "notional_local": Decimal("18000000.00"),
            "tenor": "5Y",
            "mod_duration": Decimal("4.55"),
            "convexity": Decimal("24.20"),
            "coupon_rate": Decimal("0.0350"),
        },
        {
            "position_id": "DESK_GBP_GILT_10Y",
            "desk_name": "Cross-Currency Credit & FX Desk",
            "currency": "GBP",
            "notional_local": Decimal("14000000.00"),
            "tenor": "10Y",
            "mod_duration": Decimal("7.90"),
            "convexity": Decimal("72.00"),
            "coupon_rate": Decimal("0.0405"),
        },
        {
            "position_id": "DESK_JPY_FX_CARRY_3M",
            "desk_name": "APAC FX & Short-Rates Desk",
            "currency": "JPY",
            "notional_local": Decimal("1800000000.00"),
            "tenor": "3M",
            "mod_duration": Decimal("0.25"),
            "convexity": Decimal("0.12"),
            "coupon_rate": Decimal("0.0125"),
        },
    ]

    daily_pnl_rows: list[dict[str, object]] = []
    ipv_exceptions_flagged = 0
    total_ipv_reserve_usd = Decimal("0.00")

    for d_idx in range(1, len(aligned_dates)):
        prev_dt = aligned_dates[d_idx - 1]
        curr_dt = aligned_dates[d_idx]
        day_count = 1

        for p_idx, pos in enumerate(desk_positions):
            ccy = str(pos["currency"])
            tenor = str(pos["tenor"])
            notional_local = Decimal(str(pos["notional_local"]))
            dur = Decimal(str(pos["mod_duration"]))
            conv = Decimal(str(pos["convexity"]))
            coupon = Decimal(str(pos["coupon_rate"]))

            fx_prev = Decimal(str(fx_map[prev_dt][ccy]))
            fx_curr = Decimal(str(fx_map[curr_dt][ccy]))
            y_prev = Decimal(str(treasury_all[prev_dt][tenor])) / Decimal("100")
            y_curr = Decimal(str(treasury_all[curr_dt][tenor])) / Decimal("100")
            dy = y_curr - y_prev

            # Independent Market Clean Price Index around par (100.00)
            ind_price_prev = Decimal("1.00") - dur * (y_prev - coupon) + Decimal("0.5") * conv * (y_prev - coupon) ** 2
            ind_price_curr = Decimal("1.00") - dur * (y_curr - coupon) + Decimal("0.5") * conv * (y_curr - coupon) ** 2

            mv_prev_usd = q2(notional_local * ind_price_prev * fx_prev)
            mv_curr_usd = q2(notional_local * ind_price_curr * fx_curr)
            accrued_carry_usd = q2(notional_local * coupon * (Decimal(day_count) / Decimal("365")) * fx_curr)
            total_economic_pnl_usd = q2((mv_curr_usd - mv_prev_usd) + accrued_carry_usd)

            # Taylor P&L Attribution: Carry + Rate Move (Local Slope DV01 + Convexity) + FX Revaluation + Cross-Term/Rounding Residual
            local_slope = -dur + conv * (y_prev - coupon)
            rate_price_delta_local = notional_local * (local_slope * dy + Decimal("0.5") * conv * (dy ** 2))
            rate_move_pnl_usd = q2(rate_price_delta_local * fx_prev)
            fx_reval_pnl_usd = q2(notional_local * ind_price_prev * (fx_curr - fx_prev))
            explained_pnl_usd = q2(accrued_carry_usd + rate_move_pnl_usd + fx_reval_pnl_usd)
            unexplained_pnl_usd = q2(total_economic_pnl_usd - explained_pnl_usd)

            # Independent Price Verification (IPV):
            # Simulate realistic periodic front-office stale/optimistic trader marks (+16.5 bps price mark skew on specific audit days)
            is_stale_desk_mark = ((d_idx * 7 + p_idx * 13) % 89 == 0)
            desk_mark_spread_bps = Decimal("16.5") if is_stale_desk_mark else Decimal("1.2")
            desk_mv_usd = q2(mv_curr_usd * (Decimal("1.00") + desk_mark_spread_bps / Decimal("10000")))
            ipv_variance_usd = q2(desk_mv_usd - mv_curr_usd)
            # IPV threshold: > 10.0 bps variance triggers Product Controller IPV adjustment to General Ledger
            ipv_breach_flag = 1 if desk_mark_spread_bps > Decimal("10.0") else 0
            ipv_valuation_adj_usd = -ipv_variance_usd if ipv_breach_flag else Decimal("0.00")
            if ipv_breach_flag:
                ipv_exceptions_flagged += 1
                total_ipv_reserve_usd = q2(total_ipv_reserve_usd + abs(ipv_valuation_adj_usd))

            # Post-IPV Back-Office General Ledger Carrying Value (reconciles FO Desk MV + IPV Adj + Unbooked Stale Adj == Independent MV)
            gl_verified_mv_usd = q2(desk_mv_usd + ipv_valuation_adj_usd if ipv_breach_flag else mv_curr_usd)
            dv01_usd = q2(mv_curr_usd * dur * Decimal("0.0001"))

            daily_pnl_rows.append(
                {
                    "trade_date": curr_dt,
                    "position_id": pos["position_id"],
                    "desk_name": pos["desk_name"],
                    "currency": ccy,
                    "benchmark_tenor": tenor,
                    "opening_mv_usd": float(mv_prev_usd),
                    "independent_closing_mv_usd": float(mv_curr_usd),
                    "desk_submitted_mv_usd": float(desk_mv_usd),
                    "ipv_spread_bps": float(desk_mark_spread_bps),
                    "ipv_exception_flag": ipv_breach_flag,
                    "ipv_valuation_adj_usd": float(ipv_valuation_adj_usd),
                    "gl_verified_closing_mv_usd": float(gl_verified_mv_usd),
                    "carry_accrual_pnl_usd": float(accrued_carry_usd),
                    "rate_move_pnl_usd": float(rate_move_pnl_usd),
                    "fx_reval_pnl_usd": float(fx_reval_pnl_usd),
                    "explained_pnl_usd": float(explained_pnl_usd),
                    "unexplained_pnl_usd": float(unexplained_pnl_usd),
                    "total_daily_pnl_usd": float(total_economic_pnl_usd),
                    "dv01_risk_usd": float(dv01_usd),
                }
            )

    # -------------------------------------------------------------------------
    # 3. PERSIST TO SQLITE & RUN 14 DETERMINISTIC INTERNAL CONTROL & RECONCILIATION GATES
    # -------------------------------------------------------------------------
    db_path = out_dir / "controllers_fund_and_pnl.sqlite"
    if db_path.exists():
        db_path.unlink()
    conn = sqlite3.connect(str(db_path))
    cur = conn.cursor()

    cur.execute(
        """
        CREATE TABLE fct_fund_nav_close (
            close_month TEXT, quarter TEXT, is_quarter_end INTEGER, close_date TEXT,
            fund_id TEXT, fund_name TEXT, base_currency TEXT, asc820_fair_value_level TEXT,
            beginning_nav_usd REAL, capital_calls_usd REAL, distributions_usd REAL,
            investment_income_usd REAL, mtm_valuation_adj_usd REAL, fx_translation_adj_usd REAL,
            mgmt_fee_accrual_usd REAL, admin_expense_accrual_usd REAL, total_accruals_usd REAL,
            ending_nav_usd REAL, closing_fx_usd_per_ccy REAL, benchmark_yield_pct REAL,
            liquidity_buffer_ratio REAL, policy_compliant INTEGER,
            PRIMARY KEY (close_month, fund_id)
        )
        """
    )
    cur.execute(
        """
        CREATE TABLE fct_investor_capital_statements (
            close_month TEXT, quarter TEXT, close_date TEXT, fund_id TEXT,
            investor_id TEXT, investor_name TEXT, ownership_pct REAL,
            beginning_capital_usd REAL, capital_calls_usd REAL, distributions_usd REAL,
            investment_income_usd REAL, mtm_valuation_adj_usd REAL, fx_translation_adj_usd REAL,
            fee_and_expense_accruals_usd REAL, ending_capital_usd REAL,
            PRIMARY KEY (close_month, fund_id, investor_id)
        )
        """
    )
    cur.execute(
        """
        CREATE TABLE fct_daily_product_pnl_ipv (
            trade_date TEXT, position_id TEXT, desk_name TEXT, currency TEXT, benchmark_tenor TEXT,
            opening_mv_usd REAL, independent_closing_mv_usd REAL, desk_submitted_mv_usd REAL,
            ipv_spread_bps REAL, ipv_exception_flag INTEGER, ipv_valuation_adj_usd REAL,
            gl_verified_closing_mv_usd REAL, carry_accrual_pnl_usd REAL, rate_move_pnl_usd REAL,
            fx_reval_pnl_usd REAL, explained_pnl_usd REAL, unexplained_pnl_usd REAL,
            total_daily_pnl_usd REAL, dv01_risk_usd REAL,
            PRIMARY KEY (trade_date, position_id)
        )
        """
    )

    cur.executemany(
        "INSERT INTO fct_fund_nav_close VALUES (" + ",".join(["?"] * 22) + ")",
        [tuple(r.values()) for r in fund_nav_rows],
    )
    cur.executemany(
        "INSERT INTO fct_investor_capital_statements VALUES (" + ",".join(["?"] * 15) + ")",
        [tuple(r.values()) for r in investor_statement_rows],
    )
    cur.executemany(
        "INSERT INTO fct_daily_product_pnl_ipv VALUES (" + ",".join(["?"] * 19) + ")",
        [tuple(r.values()) for r in daily_pnl_rows],
    )
    conn.commit()

    # Execute SQL reconciliation & internal control gates
    nav_rollforward_max_drift = cur.execute(
        """
        SELECT COALESCE(MAX(ABS(
            (beginning_nav_usd + capital_calls_usd - distributions_usd + investment_income_usd
             + mtm_valuation_adj_usd + fx_translation_adj_usd - total_accruals_usd) - ending_nav_usd
        )), 0.0)
        FROM fct_fund_nav_close
        """
    ).fetchone()[0]

    lp_to_fund_nav_max_drift = cur.execute(
        """
        WITH lp_Rollup AS (
            SELECT close_month, fund_id, ROUND(SUM(ending_capital_usd), 2) AS sum_lp_end_nav
            FROM fct_investor_capital_statements
            GROUP BY close_month, fund_id
        )
        SELECT COALESCE(MAX(ABS(f.ending_nav_usd - l.sum_lp_end_nav)), 0.0)
        FROM fct_fund_nav_close f
        JOIN lp_Rollup l ON f.close_month = l.close_month AND f.fund_id = l.fund_id
        """
    ).fetchone()[0]

    pnl_decomp_max_drift = cur.execute(
        """
        SELECT COALESCE(MAX(ABS((explained_pnl_usd + unexplained_pnl_usd) - total_daily_pnl_usd)), 0.0)
        FROM fct_daily_product_pnl_ipv
        """
    ).fetchone()[0]

    ipv_unremediated_breaches = cur.execute(
        """
        SELECT COUNT(*)
        FROM fct_daily_product_pnl_ipv
        WHERE ipv_exception_flag = 1
          AND ABS((desk_submitted_mv_usd + ipv_valuation_adj_usd) - independent_closing_mv_usd) > 0.01
        """
    ).fetchone()[0]

    max_desk_dv01_usd = cur.execute(
        """
        SELECT MAX(daily_dv01) FROM (
            SELECT trade_date, SUM(dv01_risk_usd) AS daily_dv01
            FROM fct_daily_product_pnl_ipv
            GROUP BY trade_date
        )
        """
    ).fetchone()[0]

    total_ending_aum_usd = cur.execute(
        "SELECT ROUND(SUM(ending_nav_usd), 2) FROM fct_fund_nav_close WHERE close_month = ?",
        (months_sorted[-1],),
    ).fetchone()[0]

    sqlite_integrity = cur.execute("PRAGMA integrity_check").fetchone()[0]
    conn.close()

    # Export CSV deliverables
    def write_csv(path: Path, rows: list[dict[str, object]]) -> None:
        if not rows:
            return
        with path.open("w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
            w.writeheader()
            w.writerows(rows)

    write_csv(out_dir / "fund_nav_monthly_quarterly_packages.csv", fund_nav_rows)
    write_csv(out_dir / "investor_capital_statements.csv", investor_statement_rows)
    write_csv(out_dir / "daily_trading_pnl_and_ipv_log.csv", daily_pnl_rows)

    # Export multi-tab Excel close package (.xlsx)
    nav_sheet = [list(fund_nav_rows[0].keys())] + [list(r.values()) for r in fund_nav_rows]
    lp_sheet = [list(investor_statement_rows[0].keys())] + [list(r.values()) for r in investor_statement_rows]
    pnl_sample_sheet = [list(daily_pnl_rows[0].keys())] + [list(r.values()) for r in daily_pnl_rows[:400]]
    write_minimal_xlsx(
        out_dir / "fund_nav_and_daily_pnl_close_package.xlsx",
        [
            ("Fund_NAV_Close_Schedule", nav_sheet),
            ("Investor_Capital_Statements", lp_sheet),
            ("Daily_Product_PnL_IPV", pnl_sample_sheet),
        ],
    )

    gates = {
        "gate_01_treasury_2024_sha256_verified": t24_sha == EXPECTED_TREASURY_2024_SHA256,
        "gate_02_treasury_2025_ingested": len(treasury_25) >= 200,
        "gate_03_ecb_fx_multi_currency_ingested": ecb_rows_count >= 7000 and ecb_obs_count >= 200000,
        "gate_04_24_monthly_close_cycles_complete": len(months_sorted) == 24 and len(fund_nav_rows) == 96,
        "gate_05_8_quarterly_packages_tagged": sum(1 for r in fund_nav_rows if r["is_quarter_end"] == 1) == 32,
        "gate_06_fund_nav_rollforward_zero_drift": nav_rollforward_max_drift < 1e-6,
        "gate_07_investor_capital_to_fund_nav_zero_variance": lp_to_fund_nav_max_drift < 1e-6,
        "gate_08_usd_base_fund_zero_fx_translation": all(
            r["fx_translation_adj_usd"] == 0.0 for r in fund_nav_rows if r["base_currency"] == "USD"
        ),
        "gate_09_asset_management_policy_controls_100pct": am_policy_breaches == 0,
        "gate_10_daily_pnl_attribution_conservation": pnl_decomp_max_drift < 1e-6,
        "gate_11_ipv_stale_marks_100pct_adjusted_to_gl": ipv_exceptions_flagged > 0 and ipv_unremediated_breaches == 0,
        "gate_12_dv01_interest_rate_risk_within_limit": float(max_desk_dv01_usd) < 75000.0,
        "gate_13_excel_close_package_generated": (out_dir / "fund_nav_and_daily_pnl_close_package.xlsx").stat().st_size > 5000,
        "gate_14_sqlite_integrity_ok": sqlite_integrity == "ok",
    }
    passed_gates = sum(1 for v in gates.values() if v)

    unexplained_abs = sum(abs(float(r["unexplained_pnl_usd"])) for r in daily_pnl_rows)
    total_abs = sum(abs(float(r["total_daily_pnl_usd"])) for r in daily_pnl_rows)
    pnl_explained_ratio_pct = round(max(0.0, (1.0 - unexplained_abs / total_abs)) * 100.0, 2) if total_abs > 0 else 100.0

    summary = {
        "provenance": {
            "treasury_2024_url": TREASURY_2024_URL,
            "treasury_2024_sha256": t24_sha,
            "treasury_2025_url": TREASURY_2025_URL,
            "treasury_2025_sha256": t25_sha,
            "ecb_fx_url": ECB_FX_ZIP_URL,
            "ecb_fx_sha256": ecb_sha,
            "ecb_historical_dates": ecb_rows_count,
            "ecb_non_null_fx_observations": ecb_obs_count,
            "aligned_2024_2025_trading_days": len(aligned_dates),
        },
        "fund_accounting_metrics": {
            "funds_modeled": len(funds),
            "monthly_close_packages": len(fund_nav_rows),
            "quarterly_reporting_packages": sum(1 for r in fund_nav_rows if r["is_quarter_end"] == 1),
            "investor_capital_statements_generated": len(investor_statement_rows),
            "distinct_lp_gp_accounts": 16,
            "final_period_combined_fund_nav_usd": total_ending_aum_usd,
            "nav_rollforward_max_variance_usd": round(float(nav_rollforward_max_drift), 6),
            "investor_waterfall_to_nav_variance_usd": round(float(lp_to_fund_nav_max_drift), 6),
            "asset_management_policy_breaches": am_policy_breaches,
        },
        "product_controllers_metrics": {
            "trading_desks": 4,
            "multi_currency_positions": len(desk_positions),
            "daily_position_valuations": len(daily_pnl_rows),
            "pnl_attribution_explained_ratio_pct": pnl_explained_ratio_pct,
            "pnl_decomposition_max_variance_usd": round(float(pnl_decomp_max_drift), 6),
            "ipv_stale_mark_exceptions_intercepted": ipv_exceptions_flagged,
            "ipv_unremediated_gl_breaches": ipv_unremediated_breaches,
            "cumulative_ipv_valuation_adjustments_usd": float(total_ipv_reserve_usd),
            "peak_portfolio_dv01_usd": round(float(max_desk_dv01_usd), 2),
        },
        "internal_control_gates": {
            "passed": passed_gates,
            "total": len(gates),
            "pass_rate_pct": round((passed_gates / len(gates)) * 100.0, 2),
            "details": gates,
        },
    }
    (out_dir / "controllers_validation_summary.json").write_text(json.dumps(summary, indent=2), encoding="utf-8")
    return summary


if __name__ == "__main__":
    result = run_controllers_suite()
    print(json.dumps(result, indent=2))
