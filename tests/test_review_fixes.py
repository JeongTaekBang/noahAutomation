"""
Regression tests for recent code review fixes.

Covers:
1. db_sync SheetSyncResult.success reflects errors
2. reconcile_so period_code_to_col picks latest year
3. reconcile_so load_ax_sales deduplicates Projects
4. reconcile_ind fill_industry_code marks duplicates as '중복검토'
"""

from pathlib import Path

import pandas as pd
import pytest


# ── Test 1: db_sync SheetSyncResult.success ──────────────────────────


def test_sheet_sync_result_reports_failure_on_errors():
    """SheetSyncResult.success must be False when errors > 0."""
    from po_generator.db_sync import SheetSyncResult

    r = SheetSyncResult(sheet_name="test", table_name="test")
    assert r.success, "fresh result with 0 errors should be success"

    r.errors = 1
    assert not r.success, "result with errors > 0 must report failure"


# ── Test 2: reconcile_so period_code_to_col ──────────────────────────


def test_period_code_to_col_picks_latest_year():
    """When multiple FX columns match the same month, the latest year wins."""
    from reconcile_so import period_code_to_col

    fx_cols = ["2025-03", "2025-06", "2026-03", "2026-06"]
    assert period_code_to_col("P03", fx_cols) == "2026-03"
    assert period_code_to_col("P06", fx_cols) == "2026-06"

    # Single match still works
    assert period_code_to_col("P01", ["2026-01", "2026-02"]) == "2026-01"

    # No match returns None
    assert period_code_to_col("P12", fx_cols) is None


# ── Test 3: reconcile_so AX Sales deduplication ──────────────────────


def test_load_ax_sales_deduplicates_projects(tmp_path):
    """Duplicate AX Project rows must be aggregated (summed), not duplicated."""
    from reconcile_so import load_ax_sales

    df = pd.DataFrame(
        {
            "Project": ["P001", "P001", "P002"],
            "Customer": ["CustA", "CustA", "CustB"],
            "AX": [100, 200, 300],
        }
    )
    f = tmp_path / "ax_sales.xlsx"
    df.to_excel(f, index=False)

    result = load_ax_sales(f)

    # P001 should be aggregated: 100 + 200 = 300
    assert len(result) == 2
    p001 = result[result["Project"] == "P001"]
    assert p001["AX"].iloc[0] == 300


# ── Test 4: reconcile_ind duplicate O.C No. → '중복검토' ─────────────


def test_fill_industry_code_marks_duplicates_for_review():
    """Duplicate O.C No. entries must be flagged '중복검토', NOT filled."""
    from reconcile_ind import fill_industry_code

    ob = pd.DataFrame(
        {
            "발주번호": ["OC-001", "OC-002", "OC-003"],
            "Industry code": [None, None, None],
        }
    )
    mapping = {
        "OC-001": ("SO-001", "IND-A", "Sector-X"),
        "OC-002": ("SO-002", "IND-B", "Sector-Y"),
        "OC-003": ("SO-003", "IND-C", "Sector-Z"),
    }
    dup_oc_set = {"OC-002"}  # OC-002 is a duplicate

    result, total_null, filled, so_missing, po_missing, dup_review = fill_industry_code(
        ob, mapping, "Industry code", dup_oc_set=dup_oc_set,
    )

    # OC-001: normal match, filled
    assert result.loc[0, "매핑상태"] == "매칭"
    assert result.loc[0, "Industry code"] == "IND-A"

    # OC-002: duplicate, NOT filled, marked for review
    assert result.loc[1, "매핑상태"] == "중복검토"
    assert pd.isna(result.loc[1, "Industry code"])

    # OC-003: normal match, filled
    assert result.loc[2, "매핑상태"] == "매칭"
    assert result.loc[2, "Industry code"] == "IND-C"

    # Counts: 3 total null, 2 filled, 0 so_missing, 0 po_missing, 1 dup_review
    assert total_null == 3
    assert filled == 2
    assert so_missing == 0
    assert po_missing == 0
    assert dup_review == 1
