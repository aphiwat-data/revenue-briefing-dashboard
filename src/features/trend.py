"""
Trend page: combo chart + forecast trend + daily momentum.

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

from plotly.subplots import make_subplots


def render_forecast_trend_by_month_v3(metric_data):
    st.markdown('<div class="section-title">Duetto Trend by Stay Month</div>', unsafe_allow_html=True)
    st.caption("Use this to see forecast trend by selected stay month. Point labels can be shown only on latest points to avoid clutter.")

    d4 = metric_data[metric_data["Reference"] == "Duetto"].copy()
    if d4.empty:
        st.info("No forecast trend data.")
        return pd.DataFrame()

    month_options = sorted(d4["Stay Month"].dropna().unique(), key=month_sort_key)

    c1, c2, c3, c4 = st.columns([1, 1.5, 1, 0.9])

    metric_choice = c1.selectbox(
        "Metric",
        ["Rev", "Occ", "Room", "ADR", "All Metrics"],
        index=0,
        key="trend_v3_metric",
    )

    trend_months = c2.multiselect(
        "Select Stay Months",
        options=month_options,
        default=month_options,   # default = all months
        key="trend_v3_month_filter",
        help="Select one or more stay months for this trend chart.",
    )

    label_mode = c3.selectbox(
        "Point labels",
        ["Latest point only", "All points", "Hide labels"],
        index=0,
        key="trend_v3_labels",
    )

    view_mode = c4.selectbox(
        "View",
        ["Value", "Daily % Change"],
        index=0,
        key="trend_v3_view_mode",
        help="Daily % Change = (today - yesterday) / yesterday x 100 per stay month",
    )

    view = d4.copy()

    if metric_choice != "All Metrics":
        view = view[view["Metric"] == metric_choice].copy()

    view = view[view["Stay Month"].isin(trend_months)].copy()

    if view.empty:
        st.info("No trend data for selected filters.")
        return pd.DataFrame()

    trend = (
        view.groupby(["Report Date", "Stay Month"], as_index=False)["Value"]
        .sum()
        .sort_values(["Stay Month", "Report Date"])
    )

    # -- Daily % Change mode -----------------------------------
    is_pct_mode = (view_mode == "Daily % Change")
    if is_pct_mode:
        trend["Pct Change"] = (
            trend.groupby("Stay Month")["Value"]
            .pct_change() * 100
        )
        trend = trend.dropna(subset=["Pct Change"]).copy()
        plot_col   = "Pct Change"
        y_title    = "Daily % Change"
        chart_title = f"Daily % Change - Forecast by Stay Month ({metric_choice})"
    else:
        plot_col   = "Value"
        y_title    = "Forecast"
        chart_title = f"Forecast Trend by Stay Month ({metric_choice})"

    # -- Point labels -----------------------------------------
    def _fmt_label(val):
        if pd.isna(val):
            return ""
        return f"{val:+.2f}%" if is_pct_mode else fmt_raw2(val)

    if label_mode == "All points":
        trend["Label"] = trend[plot_col].apply(_fmt_label)
    elif label_mode == "Latest point only":
        latest_date = trend["Report Date"].max()
        trend["Label"] = trend.apply(
            lambda r: _fmt_label(r[plot_col]) if r["Report Date"] == latest_date else "",
            axis=1,
        )
    else:
        trend["Label"] = ""

    fig = px.line(
        trend,
        x="Report Date",
        y=plot_col,
        color="Stay Month",
        markers=True,
        text="Label",
        title=chart_title,
        render_mode="webgl",
    )
    fig.update_traces(
        textposition="top center",
        line=dict(width=2),
        marker=dict(size=6),
    )
    fig.update_layout(
        plot_bgcolor="rgba(0,0,0,0)",
        paper_bgcolor="rgba(0,0,0,0)",
        xaxis=dict(
            title="Report Date",
            tickformat="%d %b",
            showgrid=True,
            gridcolor="#f0f0f0",
            zeroline=False,
        ),
        yaxis=dict(
            title=y_title,
            showgrid=True,
            gridcolor="#f0f0f0",
            zeroline=is_pct_mode,
            zerolinecolor="#999",
            zerolinewidth=1,
            ticksuffix="%" if is_pct_mode else "",
        ),
        legend=dict(
            title="Stay Month",
            orientation="v",
            x=1.01,
            y=1,
            bgcolor="rgba(0,0,0,0)",
            bordercolor="#e4e4e4",
            borderwidth=1,
        ),
        margin=dict(l=20, r=20, t=50, b=20),
        height=400,
        hovermode="x unified",
    )
    st.plotly_chart(
        fig,
        use_container_width=True,
        config={
            "displaylogo": False,
            "modeBarButtonsToRemove": ["lasso2d", "select2d", "autoScale2d"],
            "scrollZoom": False,
        },
        key="forecast_trend_v3",
    )

    with st.expander("Trend data"):
        show = trend.copy()
        show["Value"] = show["Value"].apply(fmt_raw2)
        st.dataframe(show, use_container_width=True, hide_index=True, height=min(520, 48 + 34 * len(show)))

    return trend

def render_trend_comparison(metric_data, role_selection):
    """
    Trend Comparison - single combo chart with dual Y-axis.
         Metric A -> bars on LEFT axis (e.g. Revenue)
         Metric B -> line on RIGHT axis (e.g. ADR)
         Baseline drawn as dashed reference for each metric.
    """
    st.markdown(
        '<div class="section-title">Trend Comparison - Two Metrics vs Baseline</div>',
        unsafe_allow_html=True,
    )
    st.caption(
        "Bars + line compare either **Forecast VS Budget** or **OTB VS Budget**. "
        "Each baseline appears as a dashed reference."
    )

    if metric_data is None or metric_data.empty:
        st.info("No data.")
        return

    # Use the LATEST report per (Hotel, Stay Month, Metric, Reference) instead of
    # a single "today_label". This way a selected past stay month (e.g. May while
    # we're on June's report) pulls data from its OWN latest report file.
    if metric_data is None or metric_data.empty:
        st.info("No data.")
        return
    if "Report Date" in metric_data.columns:
        latest = (
            metric_data.sort_values("Report Date")
            .drop_duplicates(
                subset=["Hotel", "Stay Month", "Metric", "Reference"],
                keep="last",
            )
            .copy()
        )
    else:
        latest = metric_data.copy()
    if latest.empty:
        st.info("No data available for the current filter.")
        return

    # Available metrics (already short names: Rev / ADR / Occ / Room)
    available_metrics = [m for m in ["Rev", "ADR", "Occ", "Room"] if m in latest["Metric"].unique()]
    if len(available_metrics) < 2:
        st.info("Need at least 2 metrics in data.")
        return

    # Available baselines
    available_refs = latest["Reference"].dropna().unique().tolist()
    baseline_options = [
        r for r in ["Budget", "Duetto", "STLY", "ST2Y", "ST3Y", "Final LY", "Final 2Y", "Final 3Y"]
        if r in available_refs
    ]
    if not baseline_options:
        st.info("No baseline references available.")
        return

    # -- Controls ---------------------------------------------
    # Comparison mode pill - what to compare against Budget:
    #   "Forecast VS Budget" (default when available - Duetto vs locked Budget)
    #   "OTB VS Budget"      (Today/on-the-book vs locked Budget)
    cmp_modes = []
    latest_refs = latest["Reference"].unique()
    if "Duetto" in latest_refs and "Budget" in latest_refs:
        cmp_modes.append("Forecast VS Budget")
    if "Today" in latest_refs and "Budget" in latest_refs:
        cmp_modes.append("OTB VS Budget")
    if not cmp_modes:
        st.info("No Forecast/OTB and Budget reference available.")
        return

    c_mode, c_a, c_b = st.columns([1.2, 1, 1])
    with c_mode:
        st.markdown(
            '<p style="font-size:0.76rem;font-weight:600;color:#6b7280;'
            'text-transform:uppercase;letter-spacing:0.06em;margin:0 0 4px 0;">'
            'Compare</p>',
            unsafe_allow_html=True,
        )
        cmp_choice = st.pills(
            "Compare",
            options=cmp_modes,
            selection_mode="single",
            default=cmp_modes[0],
            key="trend_cmp_mode",
            label_visibility="collapsed",
        )
        if not cmp_choice:
            cmp_choice = cmp_modes[0]
    # Map mode -> reference names in data
    actual_ref = "Duetto" if cmp_choice == "Forecast VS Budget" else "Today"
    actual_label = "Forecast" if actual_ref == "Duetto" else "OTB"
    baseline = "Budget"

    default_a = "Rev" if "Rev" in available_metrics else available_metrics[0]
    with c_a:
        metric_a = st.selectbox(
            "Bars (left axis)",
            available_metrics,
            index=available_metrics.index(default_a),
            key="trend_cmp_metric_a",
        )
    b_opts = [m for m in available_metrics if m != metric_a]
    default_b = "ADR" if "ADR" in b_opts else b_opts[0]
    with c_b:
        metric_b = st.selectbox(
            "Line (right axis)",
            b_opts,
            index=b_opts.index(default_b),
            key="trend_cmp_metric_b",
        )

    # -- Data builder -----------------------------------------
    # Actual = Forecast (Duetto) or OTB (Today); baseline = Budget
    def get_trend(metric):
        sub = latest[latest["Metric"] == metric]
        actual = sub[sub["Reference"] == actual_ref].groupby("Stay Month")["Value"].sum()
        base = sub[sub["Reference"] == baseline].groupby("Stay Month")["Value"].sum()
        months = sorted(set(actual.index) | set(base.index), key=month_sort_key)
        return (
            months,
            [actual.get(m, None) for m in months],
            [base.get(m, None) for m in months],
        )

    months_a, actual_a, base_a = get_trend(metric_a)
    months_b, actual_b, base_b = get_trend(metric_b)

    if not months_a and not months_b:
        st.info("No stay months in data.")
        return

    # ---------------------------------------------------------
    # DUETTO-STYLE CUMULATIVE CURVE
    #    Filled area (green) for Metric A - left Y axis
    #    Smooth line (amber) for Metric B - right Y axis
    #    Dashed line baselines on each axis
    # ---------------------------------------------------------
    _metric_full = {"Rev": "Revenue", "ADR": "ADR", "Occ": "Occupancy", "Room": "Rooms"}
    full_a = f"{actual_label} {_metric_full.get(metric_a, metric_a)}"
    full_b = f"{actual_label} {_metric_full.get(metric_b, metric_b)}"

    months_union = sorted(set(months_a) | set(months_b), key=month_sort_key)

    def _align(months_src, values, target):
        if not months_src:
            return [None] * len(target)
        idx = dict(zip(months_src, values))
        return [idx.get(m, None) for m in target]

    actual_a_v = _align(months_a, actual_a, months_union)
    base_a_v   = _align(months_a, base_a,   months_union)
    actual_b_v = _align(months_b, actual_b, months_union)
    base_b_v   = _align(months_b, base_b,   months_union)

    INK = "#0f172a"

    # Duetto-inspired palette
    GREEN_LINE    = "#0f766e"   # teal-700 (Rev line on top of area)
    GREEN_FILL    = "rgba(134,239,172,0.55)"   # green-300 with transparency
    GREEN_BASE    = "#0d9488"   # teal-600 dashed baseline
    AMBER_LINE    = "#f59e0b"   # amber-500 ADR line
    AMBER_BASE    = "#d97706"   # amber-600 ADR baseline dotted

    def _smart_fmt(v):
        if v is None or pd.isna(v):
            return ""
        v = float(v)
        if abs(v) >= 1_000_000:
            return f"{v/1_000_000:,.2f}M"
        if abs(v) >= 10_000:
            return f"{v/1_000:,.0f}K"
        if abs(v) >= 1_000:
            return f"{v/1_000:,.1f}K"
        return f"{v:,.1f}"

    # -- Aggregate totals for KPI cards ------------------------
    def _total(vals):
        return sum(v for v in vals if v is not None and pd.notna(v))

    def _avg(vals):
        clean = [v for v in vals if v is not None and pd.notna(v)]
        return sum(clean) / len(clean) if clean else 0

    is_avg_a = metric_a in ("ADR", "Occ")
    is_avg_b = metric_b in ("ADR", "Occ")
    total_a   = _avg(actual_a) if is_avg_a else _total(actual_a)
    total_a_b = _avg(base_a)   if is_avg_a else _total(base_a)
    total_b   = _avg(actual_b) if is_avg_b else _total(actual_b)
    total_b_b = _avg(base_b)   if is_avg_b else _total(base_b)
    var_a_pct = (total_a - total_a_b) / abs(total_a_b) * 100 if total_a_b else None
    var_b_pct = (total_b - total_b_b) / abs(total_b_b) * 100 if total_b_b else None

    # -- KPI cards ---------------------------------------------
    def _kpi_card(full_label, kind, value, base_value, var_pct, accent):
        if var_pct is None:
            var_chip = '<span style="color:#9ca3af;font-size:11px;">no baseline</span>'
        else:
            arrow = "^" if var_pct >= 0 else "^"
            var_color = "#15803d" if var_pct >= 0 else "#b91c1c"
            var_bg = "#dcfce7" if var_pct >= 0 else "#fee2e2"
            var_chip = (
                f'<span style="background:{var_bg};color:{var_color};'
                f'font-size:12px;font-weight:700;padding:4px 12px;border-radius:14px;'
                f'font-variant-numeric:tabular-nums;">'
                f'{arrow} {abs(var_pct):.1f}% vs {baseline}</span>'
            )
        return (
            f'<div style="background:#fff;border:1px solid #e5e7eb;border-left:4px solid {accent};'
            f'border-radius:10px;padding:16px 20px;'
            f'box-shadow:0 1px 2px rgba(15,23,42,0.04);">'
            f'<div style="font-size:11px;font-weight:700;text-transform:uppercase;'
            f'letter-spacing:0.08em;color:#6b7280;">{full_label}  {kind}</div>'
            f'<div style="display:flex;align-items:baseline;gap:14px;margin-top:8px;">'
            f'<span style="font-size:28px;font-weight:800;color:{INK};'
            f'letter-spacing:-0.02em;font-variant-numeric:tabular-nums;">'
            f'{_smart_fmt(value)}</span>'
            f'<span style="font-size:13px;color:#9ca3af;">'
            f'baseline {_smart_fmt(base_value)}</span>'
            f'</div>'
            f'<div style="margin-top:10px;">{var_chip}</div>'
            f'</div>'
        )

    kind_a = "Avg" if is_avg_a else "Total"
    kind_b = "Avg" if is_avg_b else "Total"
    k1, k2 = st.columns(2)
    k1.markdown(_kpi_card(full_a, kind_a, total_a, total_a_b, var_a_pct, GREEN_LINE), unsafe_allow_html=True)
    k2.markdown(_kpi_card(full_b, kind_b, total_b, total_b_b, var_b_pct, AMBER_LINE), unsafe_allow_html=True)
    st.markdown('<div style="height:10px;"></div>', unsafe_allow_html=True)

    # -- Duetto-style combo chart -----------------------------
    from plotly.subplots import make_subplots
    fig = make_subplots(specs=[[{"secondary_y": True}]])

    # If only 1 stay month is in scope, the spline curve and area-fill can't
    # render visibly with a single point - we make the markers larger so the
    # single point is clearly visible. Style stays identical for 2+ months.
    n_pts = len(months_union)
    _marker_a = 14 if n_pts == 1 else 7
    _marker_b = 14 if n_pts == 1 else 7
    _marker_base = 14 if n_pts == 1 else 6
    _base_mode = "markers" if n_pts == 1 else "lines"

    # Metric A - filled area (left Y)
    fig.add_trace(
        go.Scatter(
            x=months_union,
            y=actual_a_v,
            name=full_a,
            mode="lines+markers",
            line=dict(color=GREEN_LINE, width=2.5, shape="spline", smoothing=0.6),
            marker=dict(size=_marker_a, color=GREEN_LINE, line=dict(width=2, color="#fff")),
            fill="tozeroy",
            fillcolor=GREEN_FILL,
            hovertemplate=f"<b>%{{x}}</b><br>{full_a}: %{{y:,.0f}}<extra></extra>",
        ),
        secondary_y=False,
    )

    # Metric A baseline - dashed line (left Y)
    if any(v is not None for v in base_a_v):
        fig.add_trace(
            go.Scatter(
                x=months_union,
                y=base_a_v,
                name=f"{full_a} {baseline}",
                mode=_base_mode,
                line=dict(color=GREEN_BASE, width=2, dash="dash"),
                marker=dict(
                    symbol="line-ew",
                    size=_marker_base * 3,
                    color=GREEN_BASE,
                    line=dict(width=3, color=GREEN_BASE),
                ),
                hovertemplate=f"<b>%{{x}}</b><br>{full_a} {baseline}: %{{y:,.0f}}<extra></extra>",
            ),
            secondary_y=False,
        )

    # Metric B - smooth line (right Y)
    fig.add_trace(
        go.Scatter(
            x=months_union,
            y=actual_b_v,
            name=full_b,
            mode="lines+markers",
            line=dict(color=AMBER_LINE, width=3, shape="spline", smoothing=0.7),
            marker=dict(size=_marker_b, color=AMBER_LINE, line=dict(width=2, color="#fff")),
            hovertemplate=f"<b>%{{x}}</b><br>{full_b}: %{{y:,.0f}}<extra></extra>",
        ),
        secondary_y=True,
    )

    # Metric B baseline - dotted line (right Y)
    if any(v is not None for v in base_b_v):
        fig.add_trace(
            go.Scatter(
                x=months_union,
                y=base_b_v,
                name=f"{full_b} {baseline}",
                mode=_base_mode,
                line=dict(color=AMBER_BASE, width=1.5, dash="dot"),
                marker=dict(
                    symbol="line-ew",
                    size=_marker_base * 3,
                    color=AMBER_BASE,
                    line=dict(width=2.5, color=AMBER_BASE),
                ),
                opacity=0.85,
                hovertemplate=f"<b>%{{x}}</b><br>{full_b} {baseline}: %{{y:,.0f}}<extra></extra>",
            ),
            secondary_y=True,
        )

    # Subtle hint when only 1 month - chart can't draw a "trend" with 1 point
    if n_pts == 1:
        st.caption(
            "Info:  1 stay month -  Trend - curve - "
            "( 2 -).  marker --  "
            "KPI cards -"
        )

    fig.update_layout(
        plot_bgcolor="#ffffff",
        paper_bgcolor="#ffffff",
        height=440,
        margin=dict(l=10, r=10, t=30, b=60),
        legend=dict(
            orientation="h", y=-0.18, x=0.5, xanchor="center",
            font=dict(size=12, color="#374151"),
            bgcolor="rgba(0,0,0,0)",
        ),
        hovermode="x unified",
        font=dict(family="Inter, -apple-system, sans-serif"),
    )
    fig.update_xaxes(
        showgrid=False,
        tickfont=dict(size=12, color="#1f2937"),
        showline=True, linecolor="#e5e7eb", linewidth=1,
    )
    fig.update_yaxes(
        title=dict(
            text=f"<b>{full_a}</b>",
            font=dict(size=12, color=GREEN_LINE),
            standoff=8,
        ),
        showgrid=True, gridcolor="#f3f4f6",
        zeroline=False,
        tickformat=",.0f",
        tickfont=dict(size=10, color=GREEN_LINE),
        secondary_y=False,
    )
    fig.update_yaxes(
        title=dict(
            text=f"<b>{full_b}</b>",
            font=dict(size=12, color=AMBER_LINE),
            standoff=8,
        ),
        showgrid=False,
        zeroline=False,
        tickformat=",.0f",
        tickfont=dict(size=10, color=AMBER_LINE),
        secondary_y=True,
    )

    st.plotly_chart(
        fig,
        use_container_width=True,
        config={
            "displaylogo": False,
            "modeBarButtonsToRemove": ["lasso2d", "select2d", "autoScale2d"],
        },
    )

    # -- Auto-diagnosis ----------------------------------------
    if metric_a == "Rev" and metric_b == "ADR" and var_a_pct is not None and var_b_pct is not None:
        if var_a_pct > 0 and var_b_pct > 0:
            narrative = (
                f"Revenue **+{var_a_pct:.1f}%** and ADR **+{var_b_pct:.1f}%** vs {baseline} - "
                f"growth is **rate-driven** (selling at higher prices)."
            )
            tone = "good"
        elif var_a_pct > 0 and var_b_pct <= 0:
            narrative = (
                f"Revenue **+{var_a_pct:.1f}%** despite ADR at **{var_b_pct:+.1f}%** vs {baseline} - "
                f"growth is **volume-driven** (more rooms sold at lower rates)."
            )
            tone = "info"
        elif var_a_pct <= 0 and var_b_pct > 0:
            narrative = (
                f"Revenue **{var_a_pct:+.1f}%** while ADR **+{var_b_pct:.1f}%** vs {baseline} - "
                f"rate is up but **volume is dragging** revenue down."
            )
            tone = "warn"
        else:
            narrative = (
                f"Revenue **{var_a_pct:+.1f}%** and ADR **{var_b_pct:+.1f}%** vs {baseline} - "
                f"**both rate and volume** under pressure."
            )
            tone = "bad"
        bg_map = {"good": "#f0fdf4", "info": "#eff6ff", "warn": "#fffbeb", "bad": "#fef2f2"}
        bd_map = {"good": "#22c55e", "info": "#3b82f6", "warn": "#f59e0b", "bad": "#ef4444"}
        fg_map = {"good": "#166534", "info": "#1e40af", "warn": "#92400e", "bad": "#991b1b"}
        st.markdown(
            f'<div style="background:{bg_map[tone]};border-left:4px solid {bd_map[tone]};'
            f'color:{fg_map[tone]};padding:12px 18px;border-radius:8px;'
            f'font-size:13px;line-height:1.5;margin-top:6px;">'
            f'<b style="text-transform:uppercase;font-size:10px;letter-spacing:0.08em;'
            f'opacity:0.8;">Diagnosis</b><br>'
            f'{narrative}'
            f'</div>',
            unsafe_allow_html=True,
        )

def render_revenue_momentum_pct_chart(metric_data):
    """
    Daily % Change momentum chart.
    - Bars: selected metric's daily % change
    - Dashed reference line: ADR daily % change (when selected metric != ADR)
    - Filter by stay months (default = first 3 months, can show all selected)
    """
    st.markdown('<div class="section-title">Daily Duetto Movement (%)</div>', unsafe_allow_html=True)
    st.caption(
        "Bars show day-over-day Duetto change: (Today Duetto - Yesterday Duetto) / Yesterday Duetto x 100. "
        "Default view starts with 3 stay months and can show every selected month, split month by month. Green = increase, red = decrease. "
        "Dashed line = ADR day-over-day change."
    )

    d4 = metric_data[metric_data["Reference"] == "Duetto"].copy()
    if d4.empty:
        st.info("No forecast data.")
        return

    month_options = sorted(d4["Stay Month"].dropna().unique(), key=month_sort_key)

    default_months = month_options[:3]

    c1, c2 = st.columns([1, 2.5])
    metric_choice = c1.selectbox(
        "Metric",
        ["Rev", "Occ", "Room", "ADR"],
        index=0,
        key="momentum_pct_metric",
    )
    sel_months = c2.multiselect(
        "Stay Months",
        options=month_options,
        default=default_months,
        key="momentum_pct_months",
    )

    if not sel_months:
        st.caption("Select at least one stay month.")
        return

    view_months = sel_months

    # -- Build daily % change for one stay month -----------------
    def _daily_pct(metric, stay_month):
        sub = d4[(d4["Metric"] == metric) & (d4["Stay Month"] == stay_month)].copy()
        if sub.empty:
            return pd.DataFrame()
        # Sum across selected hotels within the stay month -> one row per report date.
        agg = (
            sub.groupby("Report Date", as_index=False)["Value"]
            .sum()
            .sort_values("Report Date")
        )
        agg["Pct Change"] = agg["Value"].pct_change() * 100
        return agg.dropna(subset=["Pct Change"]).reset_index(drop=True)

    def _render_momentum_month(stay_month, container_key):
        grp = _daily_pct(metric_choice, stay_month)
        adr_ref = _daily_pct("ADR", stay_month) if metric_choice != "ADR" else pd.DataFrame()

        if grp.empty:
            st.info(f"Need at least 2 report dates to compute daily % change for {stay_month}.")
            return

        grp["Direction"] = grp["Pct Change"].apply(
            lambda x: "Up" if x > 0 else ("Down" if x < 0 else "Flat")
        )
        grp["Label"] = grp["Pct Change"].apply(lambda x: f"{x:+.1f}%")

        fig = px.bar(
            grp,
            x="Report Date",
            y="Pct Change",
            color="Direction",
            color_discrete_map={
                "Up":   "#15803d",
                "Down": "#b91c1c",
                "Flat": "#d48806",
            },
            text="Label",
            title=f"{stay_month} - {metric_choice} daily movement",
        )
        if not adr_ref.empty:
            fig.add_trace(go.Scatter(
                x=adr_ref["Report Date"],
                y=adr_ref["Pct Change"],
                name="ADR (reference)",
                mode="lines+markers",
                line=dict(color="#1e293b", width=2, dash="dash"),
                marker=dict(size=6, color="#1e293b"),
                hovertemplate="<b>%{x|%d %b}</b><br>ADR daily %: %{y:+.2f}%<extra></extra>",
            ))

        fig.update_traces(
            textposition="outside",
            textfont=dict(size=10, color="#374151"),
            marker_line_width=0,
            cliponaxis=False,
            selector=dict(type="bar"),
        )
        fig.update_layout(
            plot_bgcolor="rgba(0,0,0,0)",
            paper_bgcolor="rgba(0,0,0,0)",
            margin=dict(l=20, r=20, t=45, b=20),
            height=360,
            hovermode="x unified",
            bargap=0.35,
            showlegend=(not adr_ref.empty),
            legend=dict(
                orientation="h", y=-0.15, x=0.5, xanchor="center",
                font=dict(size=11, color="#374151"),
            ),
        )
        fig.update_xaxes(
            title="Report Date",
            tickformat="%d %b",
            showgrid=False,
        )
        fig.update_yaxes(
            zeroline=True,
            zerolinecolor="#94a3b8",
            zerolinewidth=1.5,
            showgrid=True,
            gridcolor="#f1f5f9",
            ticksuffix="%",
            title="",
        )

        st.plotly_chart(
            fig,
            use_container_width=True,
            config={
                "displaylogo": False,
                "modeBarButtonsToRemove": ["lasso2d", "select2d", "autoScale2d"],
                "scrollZoom": False,
            },
            key=f"revenue_momentum_pct_chart_{container_key}",
        )

        # -- Summary stats -------------------------------------
        pos_days = int((grp["Pct Change"] > 0).sum())
        neg_days = int((grp["Pct Change"] < 0).sum())
        avg_pct  = grp["Pct Change"].mean()
        max_row  = grp.loc[grp["Pct Change"].idxmax()]
        min_row  = grp.loc[grp["Pct Change"].idxmin()]

        s1, s2, s3, s4, s5 = st.columns(5)
        s1.metric("Days Up",   pos_days)
        s2.metric("Days Down", neg_days)
        s3.metric("Avg Daily %", f"{avg_pct:+.2f}%")
        s4.metric("Best Day",  f"{max_row['Pct Change']:+.1f}%",
                  max_row["Report Date"].strftime("%d %b"))
        s5.metric("Worst Day", f"{min_row['Pct Change']:+.1f}%",
                  min_row["Report Date"].strftime("%d %b"))

    if len(view_months) == 1:
        _render_momentum_month(view_months[0], re.sub(r"[^A-Za-z0-9]", "_", str(view_months[0])))
    else:
        tabs = st.tabs(view_months)
        for tab, stay_month in zip(tabs, view_months):
            with tab:
                _render_momentum_month(stay_month, re.sub(r"[^A-Za-z0-9]", "_", str(stay_month)))
