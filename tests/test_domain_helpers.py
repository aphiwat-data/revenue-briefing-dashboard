"""Tests for src/domain/helpers.py — budget variance, stay-month filters."""
from __future__ import annotations

import pandas as pd
import pytest

from src.domain.helpers import (
    apply_stay_month_filter,
    budget_status_from_variance,
    calc_budget_variance,
    normalize_stay_month_selection,
    short_hotel_name,
    stay_month_label,
)


class TestCalcBudgetVariance:
    def test_positive_variance(self):
        # Forecast > Budget → positive
        var, pct = calc_budget_variance(110, 100)
        assert var == 10
        assert abs(pct - 10.0) < 0.01

    def test_negative_variance(self):
        var, pct = calc_budget_variance(80, 100)
        assert var == -20
        assert abs(pct - (-20.0)) < 0.01

    def test_zero_budget_returns_none_pct(self):
        var, pct = calc_budget_variance(100, 0)
        assert pct is None

    def test_none_budget_treated_as_zero(self):
        # Implementation: None budget → treated as 0 → variance = forecast
        # but variance_pct is None (can't divide by 0)
        var, pct = calc_budget_variance(100, None)
        assert var == 100  # forecast - 0
        assert pct is None  # division by zero protected

    def test_none_forecast_treated_as_zero(self):
        var, pct = calc_budget_variance(None, 100)
        assert var == -100  # 0 - budget
        assert abs(pct - (-100.0)) < 0.01


class TestBudgetStatusFromVariance:
    def test_positive_is_above(self):
        result = budget_status_from_variance(10)
        assert "Above" in result or "above" in result

    def test_negative_is_below(self):
        result = budget_status_from_variance(-10)
        assert "Below" in result or "below" in result


class TestApplyStayMonthFilter:
    @pytest.fixture
    def df(self) -> pd.DataFrame:
        return pd.DataFrame({
            "Stay Month": ["May, 2026", "Jun, 2026", "Jul, 2026"],
            "Value": [100, 200, 300],
        })

    def test_all_keeps_everything(self, df):
        result = apply_stay_month_filter(df, "All")
        assert len(result) == 3

    def test_list_filters(self, df):
        result = apply_stay_month_filter(df, ["May, 2026", "Jun, 2026"])
        assert len(result) == 2

    def test_string_filters_single(self, df):
        result = apply_stay_month_filter(df, "May, 2026")
        assert len(result) == 1
        assert result.iloc[0]["Stay Month"] == "May, 2026"

    def test_empty_df_unchanged(self):
        empty = pd.DataFrame(columns=["Stay Month"])
        assert apply_stay_month_filter(empty, "All").empty


class TestNormalizeStayMonthSelection:
    def test_none_becomes_all(self):
        assert normalize_stay_month_selection(None) == "All"

    def test_empty_list_becomes_all(self):
        assert normalize_stay_month_selection([]) == "All"

    def test_all_string_stays_all(self):
        assert normalize_stay_month_selection("All") == "All"

    def test_single_string_becomes_list(self):
        result = normalize_stay_month_selection("May, 2026")
        assert result == ["May, 2026"]

    def test_list_with_all_becomes_all(self):
        assert normalize_stay_month_selection(["May, 2026", "All"]) == "All"


class TestShortHotelName:
    def test_unknown_hotel_returns_self(self):
        # Falls through to the input if not in map
        result = short_hotel_name("Unknown Hotel XYZ")
        assert result == "Unknown Hotel XYZ"

    def test_none_handled(self):
        # Should not crash on None
        result = short_hotel_name(None)
        assert result is not None
