"""
Budget Review page: priority table + sort board.

This module is rendered into the Streamlit app from app.py.
All Streamlit calls (st.*) happen here; pure data logic lives in src/domain/.
"""
from __future__ import annotations

import html
import re

import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go

from src.core.constants import METRIC_ORDER
from src.core.helpers import (
    fmt_pct2,
    fmt_raw2,
    fmt_signed_pct2,
    month_sort_key,
)
from src.domain.helpers import (
    apply_stay_month_filter,
    short_hotel_name,
)

from src.domain.budget import build_budget_review, build_budget_review_summary_view
from src.domain.helpers import calc_budget_variance, budget_delta_text
from src.ui.dataframe_stylers import style_pace_variance_table


def render_budget_sort_board_v32(metric_long, role_selection, selected_hotels, stay_month_selection):
    """
    Revenue-friendly Budget Sort Board for All Month usage.

    Default behavior:
    - Aggregate to Summary by Hotel to avoid huge unreadable all-month charts
    - Show cards/table first
    - Chart only Top/Bottom rows
    """
    st.markdown('<div class="section-title">Budget Sort Board</div>', unsafe_allow_html=True)
    st.caption("Main budget page: Budget vs Forecast, priority cards, and action-focused table.")

    base_metric_data = metric_long[metric_long["Hotel"].isin(selected_hotels)].copy()
    base_metric_data = apply_stay_month_filter(base_metric_data, stay_month_selection)

    budget_df = build_budget_review(base_metric_data, role_selection)

    if budget_df.empty:
        st.info("No Budget data found for selected filters.")
        return pd.DataFrame()

    c1, c2 = st.columns([1, 1])

    metric_choice = c1.selectbox(
        "Metric",
        ["All Metrics", "Rev", "Occ", "Room", "ADR"],
        index=0,
        key="sort_v32_metric",
    )

    sort_choice = c2.selectbox(
        "Sort by",
        ["Budget Variance", "Budget Variance %", "Forecast", "Budget"],
        index=0,
        key="sort_v32_sort",
    )

    summary_view = build_budget_review_summary_view(budget_df, "Summary by Hotel")
    month_view = build_budget_review_summary_view(budget_df, "Detail by Month")

    if metric_choice != "All Metrics":
        summary_view = summary_view[summary_view["Metric"] == metric_choice].copy()
        month_view = month_view[month_view["Metric"] == metric_choice].copy()

    if summary_view.empty and month_view.empty:
        st.info("No rows after selected filters.")
        return summary_view

    # Worst first = ascending (most negative variance at top)
    summary_view = summary_view.sort_values(sort_choice, ascending=True).reset_index(drop=True)
    if not month_view.empty:
        month_view["_month_sort"] = month_view["Stay Month"].apply(month_sort_key)
        month_view = (
            month_view
            .sort_values(
                ["_month_sort", "Hotel", "Metric"],
                ascending=[True, True, True],
                na_position="last",
            )
            .drop(columns=["_month_sort"])
            .reset_index(drop=True)
        )

    total_forecast = summary_view["Forecast"].sum()
    total_budget = summary_view["Budget"].sum()
    total_variance = total_forecast - total_budget
    below_rows = int((summary_view["Budget Variance"] < 0).sum())
    above_rows = int((summary_view["Budget Variance"] > 0).sum())

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Budget", fmt_raw2(total_budget))
    k2.metric("Forecast", fmt_raw2(total_forecast))
    k3.metric("Variance vs Budget", fmt_raw2(total_variance), budget_delta_text(calc_budget_variance(total_forecast, total_budget)[1]))
    k4.metric("Below Budget Rows", below_rows)

    st.markdown("#### Summary by Hotel")
    st.caption("Aggregated view for quick prioritization across selected months.")
    priority_budget_view = render_priority_budget_table(summary_view, show_heading=False, key_suffix="summary")

    st.markdown("#### Detail by Month")
    st.caption("Monthly breakdown for checking which stay months drive the total variance.")
    render_priority_budget_table(month_view, show_heading=False, key_suffix="month")

    return summary_view

def render_priority_budget_table(view, show_heading=True, key_suffix="default"):
    """
    Presentation-friendly priority budget table.
    Replaces long priority cards with a compact table.
    """
    if show_heading:
        st.markdown("#### Priority budget table")

    table_df = view.copy()
    multi_hotel = table_df["Hotel"].nunique() > 1 if "Hotel" in table_df.columns else False
    table_sort = "Worst variance first"

    if key_suffix == "month" and {"Stay Month", "Metric"}.issubset(table_df.columns):
        table_sort = "Month order"
        table_df["_month_sort"] = table_df["Stay Month"].apply(month_sort_key)
        table_df["Metric"] = pd.Categorical(table_df["Metric"], categories=METRIC_ORDER, ordered=True)
        sort_cols = ["_month_sort"]
        if "Hotel" in table_df.columns:
            sort_cols.append("Hotel")
        sort_cols.append("Metric")
        table_df = (
            table_df
            .sort_values(sort_cols, ascending=True, na_position="last")
            .drop(columns=["_month_sort"])
            .reset_index(drop=True)
        )
    elif multi_hotel:
        table_sort = st.selectbox(
            "Table sort",
            ["Worst variance first", "Best variance first", "Hotel order"],
            index=0,
            key=f"priority_budget_table_sort_{key_suffix}",
        )

    if table_sort == "Worst variance first":
        table_df = table_df.sort_values("Budget Variance", ascending=True).reset_index(drop=True)
    elif table_sort == "Best variance first":
        table_df = table_df.sort_values("Budget Variance", ascending=False).reset_index(drop=True)
    else:
        sort_cols = [c for c in ["Hotel", "Stay Month", "Metric"] if c in table_df.columns]
        table_df = table_df.sort_values(sort_cols).reset_index(drop=True) if sort_cols else table_df

    preferred_cols = [
        "Hotel",
        "Stay Month",
        "Metric",
        "Budget",
        "Forecast",
        "Today OTB",
        "OTB vs Budget %",
        "OTB vs Forecast %",
        "Budget Variance",
        "Budget Variance %",
        "Status",
    ]
    cols = [c for c in preferred_cols if c in table_df.columns]
    show = table_df[cols].copy()
    raw = table_df[cols].copy()

    for col in ["Budget", "Forecast", "Today OTB", "Budget Variance"]:
        if col in show.columns:
            show[col] = show[col].apply(lambda x: "" if pd.isna(x) else fmt_raw2(x))

    for col in ["Budget Variance %", "OTB vs Budget %", "OTB vs Forecast %"]:
        if col in show.columns:
            show[col] = show[col].apply(lambda x: "" if pd.isna(x) else fmt_signed_pct2(x))

    def style_priority_budget(_):
        styles = pd.DataFrame("", index=show.index, columns=show.columns)

        if "Budget Variance" not in raw.columns:
            return styles

        for idx, val in raw["Budget Variance"].items():
            if pd.isna(val):
                bg, text, border = "#dbeafe", "#1e3a8a", "#2563eb"
            elif val > 0:
                bg, text, border = "#bbf7d0", "#14532d", "#15803d"
            elif val < 0:
                bg, text, border = "#fecaca", "#7f1d1d", "#b91c1c"
            else:
                bg, text, border = "#fef08a", "#713f12", "#ca8a04"

            for col in ["Budget Variance", "Budget Variance %", "Status"]:
                if col in styles.columns:
                    styles.loc[idx, col] = (
                        f"background-color:{bg}; color:{text}; "
                        f"font-weight:900; border-left:4px solid {border};"
                    )
        for value_col in ["OTB vs Budget %", "OTB vs Forecast %"]:
            if value_col not in raw.columns:
                continue
            for idx, val in raw[value_col].items():
                bg, text, border = (
                    ("#bbf7d0", "#14532d", "#15803d")
                    if pd.notna(val) and val > 0
                    else ("#fecaca", "#7f1d1d", "#b91c1c")
                    if pd.notna(val) and val < 0
                    else ("#fef08a", "#713f12", "#ca8a04")
                )
                styles.loc[idx, value_col] = (
                    f"background-color:{bg}; color:{text}; "
                    f"font-weight:900; border-left:4px solid {border};"
                )

        return styles

    # Display-only column rename - internal `raw`/`show` keep old names
    # for the styler closure; we rename ONLY the dataframe shown to user.
    _display_rename = {
        "Forecast":           "Duetto",
        "Today OTB":          "Today",
        "OTB vs Budget %":    "Today VS BUD %",
        "OTB vs Forecast %":  "Today VS Duetto %",
        "Budget Variance":    "Duetto VS BUD",
        "Budget Variance %":  "Duetto VS BUD %",
    }
    show_display = show.rename(columns=_display_rename)

    # Re-define styler against new column names so it still finds the cells
    _new_to_old = {v: k for k, v in _display_rename.items()}
    def style_priority_budget_display(_):
        styles = pd.DataFrame("", index=show_display.index, columns=show_display.columns)
        if "Duetto VS BUD" not in raw.columns and "Budget Variance" not in raw.columns:
            return styles
        raw_var = raw["Budget Variance"] if "Budget Variance" in raw.columns else None
        if raw_var is not None:
            for idx, val in raw_var.items():
                if pd.isna(val):
                    bg, text, border = "#dbeafe", "#1e3a8a", "#2563eb"
                elif val > 0:
                    bg, text, border = "#bbf7d0", "#14532d", "#15803d"
                elif val < 0:
                    bg, text, border = "#fecaca", "#7f1d1d", "#b91c1c"
                else:
                    bg, text, border = "#fef08a", "#713f12", "#ca8a04"
                for col in ["Duetto VS BUD", "Duetto VS BUD %", "Status"]:
                    if col in styles.columns:
                        styles.loc[idx, col] = (
                            f"background-color:{bg}; color:{text}; "
                            f"font-weight:900; border-left:4px solid {border};"
                        )
        for display_col, raw_col in [
            ("Today VS BUD %", "OTB vs Budget %"),
            ("Today VS Duetto %", "OTB vs Forecast %"),
        ]:
            if raw_col not in raw.columns or display_col not in styles.columns:
                continue
            for idx, val in raw[raw_col].items():
                if pd.notna(val) and val > 0:
                    bg, text, border = "#bbf7d0", "#14532d", "#15803d"
                elif pd.notna(val) and val < 0:
                    bg, text, border = "#fecaca", "#7f1d1d", "#b91c1c"
                else:
                    bg, text, border = "#fef08a", "#713f12", "#ca8a04"
                styles.loc[idx, display_col] = (
                    f"background-color:{bg}; color:{text}; "
                    f"font-weight:900; border-left:4px solid {border};"
                )
        return styles

    st.dataframe(
        show_display.style.apply(style_priority_budget_display, axis=None),
        use_container_width=True,
        hide_index=True,
        height=min(540, 48 + 38 * len(show_display)),
    )

    return table_df
