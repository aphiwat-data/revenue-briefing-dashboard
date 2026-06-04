"""
Forecast Movement table (used in Overview).

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

from src.domain.aggregations import build_movement_summary
from src.domain.budget import build_forecast_movement_v31


def render_forecast_movement_table_only(metric_data, role_selection):
    """
    Presentation-friendly Forecast Movement page.

    Design decision:
    - Table only, no cards and no bar chart.
    - Default metric = All Metrics so Revenue Team can present all key metrics together.
    - Movement periods = 1 Day / 7 Days / First Day of Month.
    """
    st.markdown('<div class="section-title">Duetto Movement</div>', unsafe_allow_html=True)
    st.caption("Latest Duetto vs 1 Day / 7 Days / First Day of Month. Green = picked up, Red = dropped.")

    # Prefer the v31 builder (per-period breakdown); fall back to summary on error.
    try:
        movement = build_forecast_movement_v31(metric_data, role_selection)
    except Exception:
        try:
            movement = build_movement_summary(metric_data, role_selection)
        except Exception:
            movement = pd.DataFrame()

    if movement is None or movement.empty:
        st.info("No movement data. Upload multiple report dates to compare.")
        return pd.DataFrame()

    # -- Stay Month pills filter ------------------------------
    if "Stay Month" in movement.columns:
        mv_months = sorted(movement["Stay Month"].dropna().unique(), key=month_sort_key)
        if mv_months:
            sel_mv_months = st.pills(
                "Stay Month",
                options=mv_months,
                selection_mode="multi",
                default=list(mv_months),   # default = all months shown
                key="movement_month_pills",
                label_visibility="collapsed",
            )
            if sel_mv_months:
                movement = movement[movement["Stay Month"].isin(sel_mv_months)].copy()

    c1, c2, c3 = st.columns([1, 1, 1])

    metric_filter = c1.selectbox(
        "Metric",
        ["All Metrics", "Rev", "Occ", "Room", "ADR"],
        index=0,
        key="movement_table_only_metric",
        help="All Metrics shows everything at once.",
    )

    period_filter = c2.selectbox(
        "Compare with",
        ["All Periods", "1 Day", "7 Days", "First Day of Month"],
        index=0,
        key="movement_table_only_period",
    )

    sort_mode = c3.selectbox(
        "Sort",
        ["Hotel order", "Worst movement % first", "Best movement % first"],
        index=0,   # default = Hotel order
        key="movement_table_only_sort",
    )

    view = movement.copy()

    # Normalize possible older column names
    rename_map = {
        "Latest D4cast": "Latest Forecast",
        "Base D4cast": "Base Forecast",
        "D4cast Diff": "Movement",
        "D4cast Diff %": "Movement %",
        "Compare": "Period",
    }
    for old, new in rename_map.items():
        if old in view.columns and new not in view.columns:
            view = view.rename(columns={old: new})

    if metric_filter != "All Metrics":
        view = view[view["Metric"] == metric_filter].copy()

    if period_filter != "All Periods":
        view = view[view["Period"] == period_filter].copy()

    if view.empty:
        st.info("No movement data for selected filters.")
        return movement

    # -- Canonical sort keys -----------------------------------
    _M_ORD = {"Occ": 0, "Room": 1, "ADR": 2, "Rev": 3}
    _P_ORD = {"1 Day": 0, "7 Days": 1, "First Day of Month": 2}
    view["_mo"] = view["Metric"].map(_M_ORD).fillna(99)
    view["_po"] = view["Period"].map(_P_ORD).fillna(99) if "Period" in view.columns else 0

    if sort_mode == "Worst movement % first" and "Movement %" in view.columns:
        view = view.sort_values(
            ["Movement %", "Hotel", "_mo", "_po"],
            ascending=[True, True, True, True],
        ).reset_index(drop=True)
    elif sort_mode == "Best movement % first" and "Movement %" in view.columns:
        view = view.sort_values(
            ["Movement %", "Hotel", "_mo", "_po"],
            ascending=[False, True, True, True],
        ).reset_index(drop=True)
    else:
        view = view.sort_values(
            ["Hotel", "Stay Month", "_mo", "_po"],
        ).reset_index(drop=True)

    view = view.drop(columns=["_mo", "_po"], errors="ignore")

    # -- Top KPI row - colored ---------------------------------
    total_move = view["Movement"].sum() if "Movement" in view.columns else 0
    up_rows   = int((view["Movement"] > 0).sum()) if "Movement" in view.columns else 0
    down_rows = int((view["Movement"] < 0).sum()) if "Movement" in view.columns else 0

    def _kpi_html(label, value, color, bg):
        return (
            f'<div style="background:{bg};border-left:4px solid {color};border-radius:6px;'
            f'padding:10px 14px;">'
            f'<div style="font-size:0.72rem;font-weight:700;text-transform:uppercase;'
            f'letter-spacing:0.06em;color:#666;">{label}</div>'
            f'<div style="font-size:1.25rem;font-weight:700;color:{color};margin-top:3px;">{value}</div>'
            f'</div>'
        )

    if total_move > 0:
        move_color, move_bg = "#15803d", "#f6ffed"
    elif total_move < 0:
        move_color, move_bg = "#b91c1c", "#fff1f0"
    else:
        move_color, move_bg = "#d48806", "#fffbe6"

    k1, k2, k3 = st.columns(3)
    k1.markdown(_kpi_html("Total Movement", fmt_raw2(total_move), move_color, move_bg), unsafe_allow_html=True)
    k2.markdown(_kpi_html("Rows Up up",   str(up_rows),   "#15803d", "#f6ffed"), unsafe_allow_html=True)
    k3.markdown(_kpi_html("Rows Down down", str(down_rows), "#b91c1c", "#fff1f0"), unsafe_allow_html=True)

    st.markdown("#### Movement table")

    # Keep only presentation-friendly columns
    preferred_cols = [
        "Hotel",
        "Stay Month",
        "Metric",
        "Period",
        "Latest Forecast",
        "Base Forecast",
        "Movement %",
        "Status",
    ]
    cols = [c for c in preferred_cols if c in view.columns]
    show = view[cols].copy()
    raw = view[cols].copy()

    for col in ["Latest Forecast", "Base Forecast"]:
        if col in show.columns:
            show[col] = show[col].apply(lambda x: "" if pd.isna(x) else fmt_raw2(x))
    if "Movement %" in show.columns:
        show["Movement %"] = show["Movement %"].apply(lambda x: "" if pd.isna(x) else fmt_signed_pct2(x))

    def style_movement_table(_):
        styles = pd.DataFrame("", index=show.index, columns=show.columns)

        if "Movement %" not in raw.columns:
            return styles

        for idx, val in raw["Movement %"].items():
            if pd.isna(val):
                bg = "#dbeafe"
                text = "#1e3a8a"
                border = "#2563eb"
            elif val > 0:
                bg = "#bbf7d0"
                text = "#14532d"
                border = "#15803d"
            elif val < 0:
                bg = "#fecaca"
                text = "#7f1d1d"
                border = "#b91c1c"
            else:
                bg = "#fef08a"
                text = "#713f12"
                border = "#ca8a04"

            for col in ["Movement %", "Status"]:
                if col in styles.columns:
                    styles.loc[idx, col] = (
                        f"background-color:{bg}; color:{text}; "
                        f"font-weight:900; border-left:4px solid {border};"
                    )

        return styles

    st.dataframe(
        show.style.apply(style_movement_table, axis=None),
        use_container_width=True,
        hide_index=True,
        height=min(700, 48 + 34 * len(show)),
    )

    with st.expander("How to read Duetto Movement"):
        st.markdown("""
        **Movement = Latest Forecast - Base Forecast**

        - **1 Day** = Latest Forecast vs previous report forecast  
        - **7 Days** = Latest Forecast vs report around 7 days ago  
        - **First Day of Month** = Latest Forecast vs first report of the month  

        **Green** means forecast moved up.  
        **Red** means forecast moved down.  
        **Yellow** means flat / no movement.
        """)

    return movement
