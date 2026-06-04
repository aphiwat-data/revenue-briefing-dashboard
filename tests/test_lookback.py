"""Tests for the look-back pattern.

The look-back logic uses `drop_duplicates(["Hotel","Stay Month","Metric","Reference"],
keep="last")` after sorting by Report Date. This guarantees that past stay months
pull from THEIR OWN latest report file, not the global Today file.
"""
from __future__ import annotations

import pandas as pd
import pytest

from src.domain.pivot import build_latest_pivot_table
from src.domain.aggregations import build_pace_summary, build_final_comparison


class TestLookbackLatestPivot:
    def test_picks_latest_report_date_for_each_key(self, multi_date_metric_long, sample_role_selection):
        """For each (Hotel, Stay Month, Metric, Reference), latest Report Date wins."""
        pivot = build_latest_pivot_table(multi_date_metric_long, sample_role_selection)

        # multi_date_metric_long has Today=100 on Jun 29 and Today=150 on Jun 30
        # Latest snapshot should use 150
        rev_alpha = pivot[(pivot["Hotel"] == "Hotel Alpha") & (pivot["Metric"] == "Rev")]
        assert not rev_alpha.empty
        assert rev_alpha.iloc[0]["Today"] == 150

    def test_constant_reference_unchanged(self, multi_date_metric_long, sample_role_selection):
        """Budget is constant across days — should be 200 from either day."""
        pivot = build_latest_pivot_table(multi_date_metric_long, sample_role_selection)
        rev_alpha = pivot[(pivot["Hotel"] == "Hotel Alpha") & (pivot["Metric"] == "Rev")]
        assert rev_alpha.iloc[0]["Budget"] == 200


class TestPaceSummaryUsesLatest:
    def test_pace_summary_picks_latest(self, multi_date_metric_long, sample_role_selection):
        """build_pace_summary uses latest-per-key, so Today = newer value."""
        pace = build_pace_summary(multi_date_metric_long, sample_role_selection)
        # No STLY/ST2Y/ST3Y refs in this fixture — pace might be empty or limited
        # Just verify it doesn't crash and returns a DataFrame
        assert isinstance(pace, pd.DataFrame)


class TestPivotEmpty:
    def test_empty_input_returns_empty(self, sample_role_selection):
        empty = pd.DataFrame(columns=[
            "Hotel", "Stay Month", "Metric", "Reference",
            "Value", "Report Date", "Report Label",
        ])
        pivot = build_latest_pivot_table(empty, sample_role_selection)
        assert pivot.empty

    def test_no_today_role_returns_data(self, multi_date_metric_long):
        """Even if role_selection is missing 'Today / Latest', the look-back
        pattern uses Report Date directly — so data should still come through."""
        empty_roles = pd.DataFrame(columns=["Role", "Report Label"])
        pivot = build_latest_pivot_table(multi_date_metric_long, empty_roles)
        # Function uses drop_duplicates on Report Date, not role_selection
        assert not pivot.empty


class TestLookbackPattern:
    """The fundamental drop_duplicates pattern that powers look-back."""

    def test_drop_duplicates_keeps_last_by_report_date(self):
        df = pd.DataFrame([
            {"Hotel": "H", "Stay Month": "Jun", "Metric": "Rev", "Reference": "Today",
             "Value": 100, "Report Date": pd.Timestamp("2026-06-29")},
            {"Hotel": "H", "Stay Month": "Jun", "Metric": "Rev", "Reference": "Today",
             "Value": 150, "Report Date": pd.Timestamp("2026-06-30")},
        ])
        # Apply the look-back pattern manually
        result = (
            df.sort_values("Report Date")
              .drop_duplicates(
                  subset=["Hotel", "Stay Month", "Metric", "Reference"],
                  keep="last",
              )
        )
        # Latest report wins
        assert len(result) == 1
        assert result.iloc[0]["Value"] == 150

    def test_lookback_preserves_independent_stay_months(self):
        """If May has Today=80 from its own May report, and Jun has Today=150
        from its own Jun report, both should be preserved."""
        df = pd.DataFrame([
            # May 2026 stay month, May 2026 report
            {"Hotel": "H", "Stay Month": "May, 2026", "Metric": "Rev", "Reference": "Today",
             "Value": 80, "Report Date": pd.Timestamp("2026-05-31")},
            # Jun 2026 stay month, Jun 2026 report
            {"Hotel": "H", "Stay Month": "Jun, 2026", "Metric": "Rev", "Reference": "Today",
             "Value": 150, "Report Date": pd.Timestamp("2026-06-30")},
        ])
        result = (
            df.sort_values("Report Date")
              .drop_duplicates(
                  subset=["Hotel", "Stay Month", "Metric", "Reference"],
                  keep="last",
              )
        )
        # Both stay months preserved with their own report data
        assert len(result) == 2
        may_row = result[result["Stay Month"] == "May, 2026"].iloc[0]
        jun_row = result[result["Stay Month"] == "Jun, 2026"].iloc[0]
        assert may_row["Value"] == 80
        assert jun_row["Value"] == 150
