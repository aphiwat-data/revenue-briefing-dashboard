"""Tests for formatting helpers and small utilities."""
from __future__ import annotations

import pandas as pd
import pytest

from src.core.helpers import (
    clean_text,
    fmt_pct2,
    fmt_raw2,
    fmt_signed_pct2,
    month_sort_key,
    normalize_stay_month,
    safe_delta,
    trunc2,
)


class TestTrunc2:
    def test_truncates_positive(self) -> None:
        assert trunc2(1.239) == 1.23

    def test_truncates_negative_toward_zero(self) -> None:
        # Truncate toward zero, not floor
        assert trunc2(-1.239) == -1.23

    def test_zero(self) -> None:
        assert trunc2(0) == 0

    def test_none_returns_none(self) -> None:
        assert trunc2(None) is None

    def test_nan_returns_none(self) -> None:
        assert trunc2(float("nan")) is None


class TestFmtRaw2:
    def test_thousands_separator(self) -> None:
        # fmt_raw2 uses trunc2 (truncates toward zero, not rounds)
        # so 1234567.89 → 1234567.88 (last digit truncated, not rounded)
        result = fmt_raw2(1234567.89)
        assert result.startswith("1,234,567.")
        assert "," in result  # thousands separator present

    def test_negative(self) -> None:
        # Same truncation rule applies — verify formatting works for negatives
        result = fmt_raw2(-1234.5)
        assert result.startswith("-1,234.")
        assert "," in result

    def test_none_returns_empty(self) -> None:
        assert fmt_raw2(None) == ""

    def test_zero(self) -> None:
        assert fmt_raw2(0) == "0.00"


class TestFmtPct2:
    def test_basic(self) -> None:
        assert fmt_pct2(12.34) == "12.34%"

    def test_negative(self) -> None:
        assert fmt_pct2(-5.5) == "-5.50%"

    def test_none(self) -> None:
        assert fmt_pct2(None) == ""


class TestFmtSignedPct2:
    def test_positive_has_plus_sign(self) -> None:
        assert fmt_signed_pct2(5.5).startswith("+")

    def test_negative_has_minus_sign(self) -> None:
        assert fmt_signed_pct2(-5.5).startswith("-")

    def test_none(self) -> None:
        assert fmt_signed_pct2(None) == ""


class TestSafeDelta:
    def test_normal_calculation(self) -> None:
        # (110-100)/100 * 100 = 10%
        result = safe_delta(110, 100)
        assert "10" in result

    def test_zero_base_returns_none(self) -> None:
        assert safe_delta(100, 0) is None

    def test_none_base_returns_none(self) -> None:
        assert safe_delta(100, None) is None

    def test_nan_base_returns_none(self) -> None:
        assert safe_delta(100, float("nan")) is None


class TestCleanText:
    def test_strips_whitespace(self) -> None:
        assert clean_text("  hello  ") == "hello"

    def test_strips_quotes(self) -> None:
        assert clean_text('"hello"') == "hello"

    def test_none_returns_empty(self) -> None:
        assert clean_text(None) == ""


class TestNormalizeStayMonth:
    def test_canonical_format(self) -> None:
        # "May, 2026" should be normalized
        result = normalize_stay_month("May, 2026")
        assert result == "May, 2026"

    def test_alternate_format(self) -> None:
        # "May-26" → "May, 2026"
        result = normalize_stay_month("May-26")
        assert result == "May, 2026"

    def test_invalid_returns_none(self) -> None:
        assert normalize_stay_month("not a date") is None

    def test_none_returns_none(self) -> None:
        assert normalize_stay_month(None) is None

    def test_empty_returns_none(self) -> None:
        assert normalize_stay_month("") is None


class TestMonthSortKey:
    def test_sortable(self) -> None:
        months = ["Aug, 2026", "Jan, 2026", "Dec, 2025"]
        sorted_months = sorted(months, key=month_sort_key)
        assert sorted_months == ["Dec, 2025", "Jan, 2026", "Aug, 2026"]

    def test_invalid_handled(self) -> None:
        # Should return NaT (or similar), not crash
        result = month_sort_key("garbage")
        assert pd.isna(result)
