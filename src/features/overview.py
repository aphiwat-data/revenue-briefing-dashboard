"""
Overview page: Variance Pivot + comparison charts.

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

from src.domain.pivot import build_variance_pivot_table
from src.services.excel_export import to_duetto_pivot_svg_bytes
from src.ui.dataframe_stylers import (
    style_final_variance_table,
    style_latest_pivot_table,
    style_pace_variance_table,
    style_variance_pivot,
)


def _render_comparison_summary_cards(df, value_col, value_label):
    valid = df[(df["Budget"].notna()) & (df["Budget"] != 0)].copy()
    if valid.empty:
        return

    total_value = float(valid[value_col].sum())
    total_budget = float(valid["Budget"].sum())
    variance = total_value - total_budget
    variance_pct = variance / total_budget * 100 if total_budget else None
    above = int((valid[value_col] >= valid["Budget"]).sum())
    below = int((valid[value_col] < valid["Budget"]).sum())

    def variance_card(container):
        if variance > 0:
            bg, border, text, status = "#dcfce7", "#15803d", "#14532d", "Above Budget"
        elif variance < 0:
            bg, border, text, status = "#fee2e2", "#b91c1c", "#7f1d1d", "Below Budget"
        else:
            bg, border, text, status = "#fef9c3", "#ca8a04", "#713f12", "On Budget"

        container.markdown(
            f"""
            <div style="
                background:{bg};
                border:1px solid {border};
                border-left:6px solid {border};
                border-radius:8px;
                padding:13px 16px;
                min-height:96px;
            ">
                <div style="font-size:0.72rem;font-weight:800;text-transform:uppercase;letter-spacing:0.06em;color:{text};">
                    Variance vs Budget
                </div>
                <div style="font-size:1.45rem;font-weight:900;color:{text};line-height:1.25;margin-top:6px;">
                    {fmt_raw2(variance)}
                </div>
                <div style="font-size:0.86rem;font-weight:800;color:{text};margin-top:2px;">
                    {fmt_signed_pct2(variance_pct)}  {status}
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    c1, c2, c3, c4 = st.columns(4)
    c1.metric(value_label, fmt_raw2(total_value))
    c2.metric("Budget", fmt_raw2(total_budget))
    variance_card(c3)
    c4.metric("Properties", f"{above} above / {below} below", delta=f"of {len(valid)} hotels", delta_color="off")

def _render_forecast_vs_budget(pivot_df):
    """
    Forecast VS Budget snapshot - KPI cards only.
    Metric selector: Revenue Summary (default) or any individual metric.
    """
    st.markdown('<div class="section-title">Duetto VS Budget</div>', unsafe_allow_html=True)

    if pivot_df is None or pivot_df.empty:
        return
    if "Duetto" not in pivot_df.columns or "Budget" not in pivot_df.columns:
        return

    available_metrics = [m for m in METRIC_ORDER if not pivot_df[pivot_df["Metric"] == m].empty]
    _mc1, _mc2 = st.columns([1.25, 3])
    metric_sel_fvb = _mc1.selectbox(
        "Metric",
        ["Revenue Summary"] + available_metrics,
        index=0,
        key="fvb_metric_sel",
    )
    _mc2.markdown(
        '<div style="font-size:0.78rem;color:#6b7280;padding-top:30px;">'
        'Bars show Forecast. Black line is Budget. Green = above budget, red = below budget.'
        '</div>',
        unsafe_allow_html=True,
    )

    flt = (
        pivot_df[pivot_df["Metric"] == "Rev"].copy()
        if metric_sel_fvb == "Revenue Summary"
        else pivot_df[pivot_df["Metric"] == metric_sel_fvb].copy()
    )
    if flt.empty:
        return

    by_hotel = flt.groupby("Hotel", as_index=False).agg({"Duetto": "sum", "Budget": "sum"})
    by_hotel = by_hotel[(by_hotel["Budget"].notna()) & (by_hotel["Budget"] != 0)]
    if by_hotel.empty:
        return
    _render_comparison_summary_cards(by_hotel, "Duetto", "Forecast")

def _render_otb_comparison_chart(pivot_df):
    """
    Compare On The Book - KPI cards only.
    Metric selector: Revenue Summary (default) or individual metric.
    """
    st.markdown('<div class="section-title">Compare - Today VS Budget</div>', unsafe_allow_html=True)

    if pivot_df is None or pivot_df.empty or "Today" not in pivot_df.columns:
        st.info("No On-The-Book data available.")
        return

    available_metrics = [m for m in METRIC_ORDER if not pivot_df[pivot_df["Metric"] == m].empty]
    _oc1, _oc2 = st.columns([1.25, 3])
    metric_sel_otb = _oc1.selectbox(
        "Metric",
        ["Revenue Summary"] + available_metrics,
        index=0,
        key="otb_chart_metric_sel",
    )
    _oc2.markdown(
        '<div style="font-size:0.78rem;color:#6b7280;padding-top:30px;">'
        'Bars show On The Book. Black line is Budget. Green = above budget, red = below budget.'
        '</div>',
        unsafe_allow_html=True,
    )

    flt = (
        pivot_df[pivot_df["Metric"] == "Rev"].copy()
        if metric_sel_otb == "Revenue Summary"
        else pivot_df[pivot_df["Metric"] == metric_sel_otb].copy()
    )
    if flt.empty:
        return

    agg_cols = {c: "sum" for c in ["Today", "Budget"] if c in flt.columns}
    if "Budget" not in agg_cols:
        st.info("No Budget data available for On-The-Book comparison.")
        return
    by_hotel = flt.groupby("Hotel", as_index=False).agg(agg_cols)
    by_hotel = by_hotel[(by_hotel["Budget"].notna()) & (by_hotel["Budget"] != 0)]
    if by_hotel.empty:
        return
    _render_comparison_summary_cards(by_hotel, "Today", "On The Book")

def render_compact_hotel_tabs(pivot_df, report_file_month=None):
    """
    Duetto pivot table - variance-enhanced view.
    Columns: Today | Budget | Today VS BUD | Duetto | Duetto VS BUD | Today VS Duetto |
             STLY | Today VS STLY | Final LY | Duetto VS Final LY | ... + more.
    """
    if pivot_df is None or pivot_df.empty:
        st.info("No pivot data for selected filters.")
        return

    st.markdown('<div class="section-title">Duetto Pivot - by Stay Month</div>', unsafe_allow_html=True)

    c_view, c_image, c_legend = st.columns([1.3, 1.25, 3.1])
    with c_view:
        view_mode = st.pills(
            "View",
            options=["Hotel tabs", "All hotels"],
            default="Hotel tabs",
            key="compact_pivot_view_pills",
            label_visibility="collapsed",
        )
        if not view_mode:
            view_mode = "Hotel tabs"
    file_stamp = (
        str(report_file_month).replace(", ", "_").replace(" ", "_")
        if report_file_month
        else "selected_filters"
    )
    with c_image:
        st.download_button(
            "Export Image",
            data=to_duetto_pivot_svg_bytes(pivot_df),
            file_name=f"duetto_pivot_by_stay_month_{file_stamp}.svg",
            mime="image/svg+xml",
            type="primary",
            width="stretch",
        )
    with c_legend:
        has_var = any(" VS " in str(c) for c in pivot_df.columns)
        if has_var:
            legend_html = (
                '<div style="font-size:0.76rem;color:#6b7280;line-height:1.6;padding-top:8px;text-align:right;">'
                '<b style="color:#1e40af;">Today</b> = on-the-book &nbsp;&nbsp; '
                '<b style="color:#15803d;">Duetto</b> = forecast &nbsp;&nbsp; '
                '<b style="color:#713f12;">Budget</b> = target &nbsp;&nbsp; '
                '<b style="color:#166534;">VS %</b> = variance (green +, red -)'
                '</div>'
            )
        else:
            legend_html = (
                '<div style="font-size:0.76rem;color:#6b7280;line-height:1.6;padding-top:8px;text-align:right;">'
                '<b style="color:#1e40af;">Duetto</b> = forecast &nbsp;&nbsp; '
                '<b style="color:#713f12;">Budget</b> = target &nbsp;&nbsp; '
                'STLY/ST2Y/ST3Y = same-time &nbsp;&nbsp; Final LY/2Y/3Y = actuals'
                '</div>'
            )
        st.markdown(legend_html, unsafe_allow_html=True)

    _style_fn = style_variance_pivot if any(" VS " in str(c) for c in pivot_df.columns) else style_latest_pivot_table

    # Pin "Stay Month" + "Metric" (the context columns) to the left when the
    # table scrolls horizontally - so the user keeps seeing which row is what.
    def _pinned_config(df_cols):
        cfg = {}
        for col in ("Stay Month", "Metric"):
            if col in df_cols:
                cfg[col] = st.column_config.Column(pinned=True)
        return cfg

    if view_mode == "Hotel tabs":
        hotels = sorted(pivot_df["Hotel"].dropna().unique())
        labels = [short_hotel_name(h) for h in hotels]
        tabs = st.tabs(labels)
        for tab, hotel in zip(tabs, hotels):
            with tab:
                sub = pivot_df[pivot_df["Hotel"] == hotel].drop(columns=["Hotel"]).reset_index(drop=True)
                st.dataframe(
                    _style_fn(sub),
                    use_container_width=True,
                    hide_index=True,
                    height=min(650, 40 + 38 * len(sub)),
                    column_config=_pinned_config(sub.columns),
                )
    else:
        show = pivot_df.copy()
        show["Hotel"] = show["Hotel"].apply(short_hotel_name)
        st.dataframe(
            _style_fn(show),
            use_container_width=True,
            hide_index=True,
            height=min(750, 40 + 30 * len(show)),
            column_config=_pinned_config(show.columns),
        )
