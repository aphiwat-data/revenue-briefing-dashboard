"""Domain helpers — business logic that doesn't touch Streamlit.

Pure functions for filtering, naming, formatting, and budget variance.
"""
from __future__ import annotations

from typing import Any, Iterable

import pandas as pd

from src.core.helpers import fmt_pct2, fmt_raw2


HOTEL_SHORT_NAMES: dict[str, str] = {
    "The Grass Serviced Suites": "TG",
    "Hotel Amber Pattaya": "Amber PTY",
    "Hotel Amber Sukhumvit 85": "Amber 85",
    "Altera Hotel & Residence Pattaya": "Altera",
    "Arbour Hotel and Residence": "Arbour",
    "Arden Hotel & Residence Pattaya": "Arden",
    "Aster Hotel & Residence Pattaya": "Aster",
}


def short_hotel_name(hotel_name: Any) -> str:
    """Map a full hotel name to its short label. Falls through to the input."""
    return HOTEL_SHORT_NAMES.get(str(hotel_name), str(hotel_name))


def metric_label_order() -> list[str]:
    """Canonical order for metric columns in tables."""
    return ["Occ", "Room", "ADR", "Rev"]


def get_metric_options_with_all() -> list[str]:
    """Metric dropdown options including the 'All Metrics' choice."""
    return ["All Metrics"] + metric_label_order()


def compact_metric_table_height(
    df: pd.DataFrame | None,
    row_height: int = 36,
    max_height: int = 520,
) -> int:
    """Suggested st.dataframe height for a DataFrame, capped at max_height."""
    if df is None or df.empty:
        return 160
    return min(max_height, 48 + row_height * len(df))


def normalize_stay_month_selection(
    selected_months: str | Iterable[str] | None,
) -> str | list[str]:
    """Convert a stay-month selector value to canonical form.

    Returns "All" when all months should be included,
    or a list[str] of explicitly selected months.
    """
    if selected_months is None:
        return "All"

    if isinstance(selected_months, str):
        return "All" if selected_months == "All" else [selected_months]

    selected_list = list(selected_months)

    if not selected_list or "All" in selected_list:
        return "All"

    return selected_list


def stay_month_label(stay_month_selection: str | list[str] | tuple | set) -> str:
    """Human-readable label for a stay-month selection (caption/header use)."""
    if stay_month_selection == "All":
        return "All"
    if isinstance(stay_month_selection, (list, tuple, set)):
        values = list(stay_month_selection)
        if len(values) <= 3:
            return ", ".join(values)
        return f"{len(values)} months selected"
    return str(stay_month_selection)


def apply_stay_month_filter(
    df: pd.DataFrame | None,
    stay_month_selection: str | list[str] | tuple | set,
) -> pd.DataFrame:
    """Filter a DataFrame by stay-month selection. Supports single, list, or 'All'."""
    if df is None or df.empty:
        return df if df is not None else pd.DataFrame()

    if stay_month_selection == "All":
        return df

    if isinstance(stay_month_selection, (list, tuple, set)):
        return df[df["Stay Month"].isin(list(stay_month_selection))].copy()

    return df[df["Stay Month"] == stay_month_selection].copy()


def format_compact_value(x: Any, is_pct: bool = False) -> str:
    """Format a numeric value compactly. is_pct=True adds % suffix."""
    if x is None or pd.isna(x):
        return ""
    return fmt_pct2(x) if is_pct else fmt_raw2(x)


def add_week_columns(df: pd.DataFrame | None) -> pd.DataFrame:
    """Add a 'Report Week' column derived from Report Date (Mon-anchored)."""
    if df is None or df.empty:
        return df if df is not None else pd.DataFrame()
    out = df.copy()
    dt = pd.to_datetime(out["Report Date"])
    out["Report Week"] = dt.dt.to_period("W-MON").astype(str)
    return out


def friendly_week_label(start_date: Any, end_date: Any) -> str:
    """Render a 'DD MMM - DD MMM' label for a week range. Empty if invalid."""
    start = pd.to_datetime(start_date)
    end = pd.to_datetime(end_date)
    if pd.isna(start) or pd.isna(end):
        return ""
    return f"{start.strftime('%d %b')} - {end.strftime('%d %b')}"


def calc_budget_variance(
    forecast: float | None,
    budget: float | None,
) -> tuple[float | None, float | None]:
    """Single source of truth for Budget vs Forecast variance.

    Returns (variance, variance_pct) tuple.

    - Variance     = Forecast − Budget
    - Variance %   = Variance / Budget * 100
    - Returns (variance, None) when budget is 0/None (avoid division by zero)
    """
    if pd.isna(forecast):
        forecast = 0
    if pd.isna(budget):
        budget = 0

    variance = forecast - budget

    if budget is None or pd.isna(budget) or budget == 0:
        variance_pct = None
    else:
        variance_pct = variance / budget * 100

    return variance, variance_pct


def budget_status_from_variance(variance: float | None) -> str:
    """Map a numeric variance to a human-readable status label."""
    if variance is None or pd.isna(variance):
        return "No Budget"
    if variance > 0:
        return "Above Budget"
    if variance < 0:
        return "Below Budget"
    return "On Budget"


def budget_delta_text(variance_pct: float | None) -> str | None:
    """Format a variance % as a Streamlit metric delta string."""
    if variance_pct is None or pd.isna(variance_pct):
        return None
    return fmt_pct2(variance_pct)
