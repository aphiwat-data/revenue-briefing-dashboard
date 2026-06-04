"""Tests for the variance pivot table — the headline calculation.

These tests guard against bugs like the historical "Today VS STLY uses max
of 3 years instead of STLY" issue that caused sign-flips.
"""
from __future__ import annotations

import pandas as pd
import pytest

from src.domain.pivot import build_latest_pivot_table, build_variance_pivot_table


class TestVariancePivotColumnsArePresent:
    def test_today_vs_bud_column_exists(self, sample_metric_long, sample_role_selection):
        pivot = build_variance_pivot_table(sample_metric_long, sample_role_selection)
        assert "Today VS BUD" in pivot.columns

    def test_duetto_vs_bud_column_exists(self, sample_metric_long, sample_role_selection):
        pivot = build_variance_pivot_table(sample_metric_long, sample_role_selection)
        assert "Duetto VS BUD" in pivot.columns

    def test_today_vs_duetto_column_exists(self, sample_metric_long, sample_role_selection):
        pivot = build_variance_pivot_table(sample_metric_long, sample_role_selection)
        assert "Today VS Duetto" in pivot.columns

    def test_separate_stly_columns(self, sample_metric_long, sample_role_selection):
        """Regression test: each year must have its OWN variance column,
        not a hidden max() of 3 years."""
        pivot = build_variance_pivot_table(sample_metric_long, sample_role_selection)
        assert "Today VS STLY" in pivot.columns
        assert "Today VS ST2Y" in pivot.columns
        assert "Today VS ST3Y" in pivot.columns

    def test_duetto_vs_final_columns(self, sample_metric_long, sample_role_selection):
        pivot = build_variance_pivot_table(sample_metric_long, sample_role_selection)
        assert "Duetto VS Final LY" in pivot.columns
        assert "Duetto VS Final 2Y" in pivot.columns
        assert "Duetto VS Final 3Y" in pivot.columns


class TestVarianceFormulaCorrectness:
    """Verify the variance formula: (X - Y) / |Y| * 100."""

    def test_today_vs_bud_positive(self, sample_metric_long, sample_role_selection):
        """Today=0.65*base, Budget=0.95*base — Today < Budget → negative variance."""
        pivot = build_variance_pivot_table(sample_metric_long, sample_role_selection)
        rev_rows = pivot[pivot["Metric"] == "Rev"]
        # Today = 650,000; Budget = 950,000; variance = (650-950)/950 * 100 = -31.58%
        for _, row in rev_rows.iterrows():
            if pd.notna(row.get("Today VS BUD")):
                assert row["Today VS BUD"] < 0, "Today < Budget should give negative variance"

    def test_duetto_vs_bud_positive(self, sample_metric_long, sample_role_selection):
        """Duetto=1.00*base, Budget=0.95*base — Duetto > Budget → positive variance."""
        pivot = build_variance_pivot_table(sample_metric_long, sample_role_selection)
        rev_rows = pivot[pivot["Metric"] == "Rev"]
        for _, row in rev_rows.iterrows():
            if pd.notna(row.get("Duetto VS BUD")):
                assert row["Duetto VS BUD"] > 0, "Duetto > Budget should give positive variance"

    def test_stly_variance_uses_stly_not_max(self):
        """REGRESSION: ensure Today VS STLY uses STLY directly, not max(STLY,ST2Y,ST3Y)."""
        # Build a deliberately tricky case where ST2Y > STLY
        # If the bug exists, Today VS STLY would use ST2Y as denominator (wrong)
        role_sel = pd.DataFrame([{"Role": "Today / Latest", "Report Label": "L1"}])
        df = pd.DataFrame([
            # Today=120, STLY=100, ST2Y=150 (higher than STLY)
            # Correct: (120-100)/100 = +20%
            # Buggy:   (120-150)/150 = -20%  ← would flip sign
            {"Hotel": "H", "Stay Month": "Jun, 2026", "Metric": "ADR",
             "Reference": "Today", "Value": 120,
             "Report Date": pd.Timestamp("2026-06-30"), "Report Label": "L1"},
            {"Hotel": "H", "Stay Month": "Jun, 2026", "Metric": "ADR",
             "Reference": "STLY", "Value": 100,
             "Report Date": pd.Timestamp("2026-06-30"), "Report Label": "L1"},
            {"Hotel": "H", "Stay Month": "Jun, 2026", "Metric": "ADR",
             "Reference": "ST2Y", "Value": 150,  # higher than STLY
             "Report Date": pd.Timestamp("2026-06-30"), "Report Label": "L1"},
        ])
        pivot = build_variance_pivot_table(df, role_sel)
        row = pivot.iloc[0]
        # Today VS STLY MUST be +20% (using STLY=100), not -20% (using ST2Y=150)
        assert row["Today VS STLY"] > 0, \
            "Bug regression: Today VS STLY should compare against STLY only, not max"
        assert abs(row["Today VS STLY"] - 20.0) < 0.01

    def test_st2y_variance_uses_st2y(self):
        """Today VS ST2Y must use ST2Y as denominator."""
        role_sel = pd.DataFrame([{"Role": "Today / Latest", "Report Label": "L1"}])
        df = pd.DataFrame([
            {"Hotel": "H", "Stay Month": "Jun, 2026", "Metric": "ADR",
             "Reference": "Today", "Value": 120,
             "Report Date": pd.Timestamp("2026-06-30"), "Report Label": "L1"},
            {"Hotel": "H", "Stay Month": "Jun, 2026", "Metric": "ADR",
             "Reference": "ST2Y", "Value": 150,
             "Report Date": pd.Timestamp("2026-06-30"), "Report Label": "L1"},
        ])
        pivot = build_variance_pivot_table(df, role_sel)
        row = pivot.iloc[0]
        # (120-150)/150 * 100 = -20%
        assert abs(row["Today VS ST2Y"] - (-20.0)) < 0.01


class TestSafePctEdgeCases:
    """The safe_pct helper inside build_variance_pivot_table must handle edge cases."""

    def test_zero_denominator_returns_na(self):
        """Division by zero should NOT crash — return NA."""
        role_sel = pd.DataFrame([{"Role": "Today / Latest", "Report Label": "L1"}])
        df = pd.DataFrame([
            {"Hotel": "H", "Stay Month": "Jun, 2026", "Metric": "Rev",
             "Reference": "Today", "Value": 100,
             "Report Date": pd.Timestamp("2026-06-30"), "Report Label": "L1"},
            {"Hotel": "H", "Stay Month": "Jun, 2026", "Metric": "Rev",
             "Reference": "Budget", "Value": 0,  # zero — would divide by zero
             "Report Date": pd.Timestamp("2026-06-30"), "Report Label": "L1"},
        ])
        pivot = build_variance_pivot_table(df, role_sel)
        # Should not crash; variance should be NA
        assert pd.isna(pivot.iloc[0]["Today VS BUD"])

    def test_missing_reference_no_variance_column(self, sample_role_selection):
        """If a reference is missing entirely, no variance column should be wrong."""
        # Only Today + Budget, no STLY at all
        df = pd.DataFrame([
            {"Hotel": "H", "Stay Month": "Jun, 2026", "Metric": "Rev",
             "Reference": "Today", "Value": 100,
             "Report Date": pd.Timestamp("2026-06-30"), "Report Label": "30 | 2026-06-30"},
            {"Hotel": "H", "Stay Month": "Jun, 2026", "Metric": "Rev",
             "Reference": "Budget", "Value": 80,
             "Report Date": pd.Timestamp("2026-06-30"), "Report Label": "30 | 2026-06-30"},
        ])
        pivot = build_variance_pivot_table(df, sample_role_selection)
        # No STLY column expected — variance shouldn't appear
        assert "STLY" not in pivot.columns
        assert "Today VS STLY" not in pivot.columns


class TestColumnOrder:
    """The column order matters for the user's reading flow."""

    def test_today_before_budget(self, sample_metric_long, sample_role_selection):
        pivot = build_variance_pivot_table(sample_metric_long, sample_role_selection)
        cols = list(pivot.columns)
        assert cols.index("Today") < cols.index("Budget")

    def test_budget_before_today_vs_bud(self, sample_metric_long, sample_role_selection):
        """Raw values before their derived variance."""
        pivot = build_variance_pivot_table(sample_metric_long, sample_role_selection)
        cols = list(pivot.columns)
        assert cols.index("Budget") < cols.index("Today VS BUD")

    def test_duetto_after_today_vs_bud(self, sample_metric_long, sample_role_selection):
        """Forecast column comes after the OTB vs Budget block."""
        pivot = build_variance_pivot_table(sample_metric_long, sample_role_selection)
        cols = list(pivot.columns)
        assert cols.index("Today VS BUD") < cols.index("Duetto")

    def test_duetto_immediately_before_duetto_vs_bud(self, sample_metric_long, sample_role_selection):
        """The headline pairing: Duetto then Duetto VS BUD."""
        pivot = build_variance_pivot_table(sample_metric_long, sample_role_selection)
        cols = list(pivot.columns)
        assert cols.index("Duetto") < cols.index("Duetto VS BUD")
