from __future__ import annotations

import pandas as pd

from src.core.helpers import fmt_pct2, fmt_raw2


HOTEL_SHORT_NAMES = {
    "The Grass Serviced Suites": "TG",
    "Hotel Amber Pattaya": "Amber PTY",
    "Hotel Amber Sukhumvit 85": "Amber 85",
    "Altera Hotel & Residence Pattaya": "Altera",
    "Arbour Hotel and Residence": "Arbour",
    "Arden Hotel & Residence Pattaya": "Arden",
    "Aster Hotel & Residence Pattaya": "Aster",
}


def short_hotel_name(hotel_name):
    return HOTEL_SHORT_NAMES.get(str(hotel_name), str(hotel_name))


def metric_label_order():
    return ["Occ", "Room", "ADR", "Rev"]


def get_metric_options_with_all():
    return ["All Metrics"] + metric_label_order()


def compact_metric_table_height(df, row_height=36, max_height=520):
    if df is None or df.empty:
        return 160
    return min(max_height, 48 + row_height * len(df))


def normalize_stay_month_selection(selected_months):
    """
    Convert multiselect result to either:
    - "All" when all months should be included
    - list[str] when user selected specific stay months
    """
    if selected_months is None:
        return "All"

    if isinstance(selected_months, str):
        return "All" if selected_months == "All" else [selected_months]

    selected_months = list(selected_months)

    if not selected_months or "All" in selected_months:
        return "All"

    return selected_months


def stay_month_label(stay_month_selection):
    if stay_month_selection == "All":
        return "All"
    if isinstance(stay_month_selection, (list, tuple, set)):
        values = list(stay_month_selection)
        if len(values) <= 3:
            return ", ".join(values)
        return f"{len(values)} months selected"
    return str(stay_month_selection)


def apply_stay_month_filter(df, stay_month_selection):
    """
    Supports single month, multiple months, or All.
    """
    if df is None or df.empty:
        return df

    if stay_month_selection == "All":
        return df

    if isinstance(stay_month_selection, (list, tuple, set)):
        return df[df["Stay Month"].isin(list(stay_month_selection))].copy()

    return df[df["Stay Month"] == stay_month_selection].copy()


def format_compact_value(x, is_pct=False):
    if x is None or pd.isna(x):
        return ""
    return fmt_pct2(x) if is_pct else fmt_raw2(x)


def add_week_columns(df):
    if df is None or df.empty: return df
    out=df.copy(); dt=pd.to_datetime(out["Report Date"])
    out["Report Week"] = dt.dt.to_period("W-MON").astype(str)
    return out


def friendly_week_label(start_date, end_date):
    start = pd.to_datetime(start_date)
    end = pd.to_datetime(end_date)
    if pd.isna(start) or pd.isna(end):
        return ""
    return f"{start.strftime('%d %b')} - {end.strftime('%d %b')}"


def calc_budget_variance(forecast, budget):
    """
    Single source of truth for Budget vs Forecast.

    Budget is the target/base.
    Forecast is the latest expected result.

    Variance = Forecast - Budget
    Variance % = Variance / Budget * 100
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


def budget_status_from_variance(variance):
    if variance is None or pd.isna(variance):
        return "No Budget"
    if variance > 0:
        return "Above Budget"
    if variance < 0:
        return "Below Budget"
    return "On Budget"


def budget_delta_text(variance_pct):
    """
    Streamlit metric delta text.
    Use already-calculated budget variance %, not safe_delta().
    """
    if variance_pct is None or pd.isna(variance_pct):
        return None
    return fmt_pct2(variance_pct)
