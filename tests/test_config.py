"""Tests for the YAML-config externalisation.

Verify that REFERENCE_PATTERNS / METRIC_PATTERNS load correctly
from config/duetto_schema.yaml.
"""
from __future__ import annotations

from src.core.constants import (
    FINAL_REFS,
    METRIC_ORDER,
    METRIC_PATTERNS,
    METRIC_TO_DISPLAY,
    REFERENCE_PATTERNS,
    SAME_TIME_REFS,
)


class TestReferencePatterns:
    def test_all_canonical_refs_present(self):
        expected = {"Today", "STLY", "ST2Y", "ST3Y", "Duetto", "Budget",
                    "Final LY", "Final 2Y", "Final 3Y"}
        assert expected.issubset(set(REFERENCE_PATTERNS.keys()))

    def test_duetto_has_forecast_alias(self):
        # 'Forecast' (without 'Duetto' prefix) should match Duetto reference
        assert any("Forecast" in p for p in REFERENCE_PATTERNS["Duetto"])

    def test_each_pattern_is_list(self):
        for ref, patterns in REFERENCE_PATTERNS.items():
            assert isinstance(patterns, list), f"{ref} should map to a list"
            assert all(isinstance(p, str) for p in patterns)


class TestMetricPatterns:
    def test_four_canonical_metrics(self):
        assert {"Occupancy", "Rooms", "ADR", "Revenue"}.issubset(set(METRIC_PATTERNS.keys()))


class TestMetricMapping:
    def test_metric_to_display_complete(self):
        assert METRIC_TO_DISPLAY["Occupancy"] == "Occ"
        assert METRIC_TO_DISPLAY["Rooms"] == "Room"
        assert METRIC_TO_DISPLAY["ADR"] == "ADR"
        assert METRIC_TO_DISPLAY["Revenue"] == "Rev"

    def test_metric_order_matches_display(self):
        assert set(METRIC_ORDER) == {"Occ", "Room", "ADR", "Rev"}


class TestSameTimeAndFinalRefs:
    def test_same_time_refs(self):
        assert SAME_TIME_REFS == ["STLY", "ST2Y", "ST3Y"]

    def test_final_refs(self):
        assert FINAL_REFS == ["Final LY", "Final 2Y", "Final 3Y"]
