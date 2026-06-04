"""Shared pytest fixtures for the dashboard test suite."""
from __future__ import annotations

import pandas as pd
import pytest


@pytest.fixture
def sample_role_selection() -> pd.DataFrame:
    """Minimal role_selection mapping the 4 roles to report labels."""
    return pd.DataFrame([
        {"Role": "Today / Latest",      "Report Label": "30 | 2026-06-30"},
        {"Role": "Yesterday / Previous","Report Label": "29 | 2026-06-29"},
        {"Role": "Last 7D",             "Report Label": "23 | 2026-06-23"},
        {"Role": "1st Month",           "Report Label": "01 | 2026-06-01"},
    ])


@pytest.fixture
def sample_metric_long() -> pd.DataFrame:
    """Long-format metric data covering 2 hotels x 2 stay months x all references."""
    rows = []
    for hotel in ["Hotel Alpha", "Hotel Bravo"]:
        for stay_month in ["Jun, 2026", "Jul, 2026"]:
            # Realistic ranges
            multiplier = 1.0 if hotel == "Hotel Alpha" else 0.8
            for metric, base_val in [
                ("Rev",  1_000_000 * multiplier),
                ("Room", 1500 * multiplier),
                ("ADR",  150 * multiplier),
                ("Occ",  70 * multiplier),
            ]:
                for ref, factor in [
                    ("Today",     0.65),
                    ("Duetto",    1.00),
                    ("Budget",    0.95),
                    ("STLY",      0.90),
                    ("ST2Y",      0.85),
                    ("ST3Y",      0.80),
                    ("Final LY",  0.95),
                    ("Final 2Y",  0.88),
                    ("Final 3Y",  0.82),
                ]:
                    rows.append({
                        "Hotel": hotel,
                        "Stay Month": stay_month,
                        "Metric": metric,
                        "Reference": ref,
                        "Value": base_val * factor,
                        "Report Date": pd.Timestamp("2026-06-30"),
                        "Report Label": "30 | 2026-06-30",
                    })
    return pd.DataFrame(rows)


@pytest.fixture
def multi_date_metric_long() -> pd.DataFrame:
    """Same as sample_metric_long but with multiple Report Dates for look-back tests."""
    base = []
    for report_day, report_date in [
        ("29 | 2026-06-29", pd.Timestamp("2026-06-29")),
        ("30 | 2026-06-30", pd.Timestamp("2026-06-30")),
    ]:
        for hotel in ["Hotel Alpha", "Hotel Bravo"]:
            for stay_month in ["Jun, 2026"]:
                for metric in ["Rev", "ADR"]:
                    for ref in ["Today", "Budget", "Duetto"]:
                        # Today value GROWS over report dates (more bookings)
                        # Budget stays constant
                        # Duetto adjusts slightly
                        if ref == "Today":
                            val = 100 if report_date.day == 29 else 150
                        elif ref == "Budget":
                            val = 200  # constant
                        else:  # Duetto
                            val = 180 if report_date.day == 29 else 195
                        base.append({
                            "Hotel": hotel,
                            "Stay Month": stay_month,
                            "Metric": metric,
                            "Reference": ref,
                            "Value": val,
                            "Report Date": report_date,
                            "Report Label": report_day,
                        })
    return pd.DataFrame(base)
