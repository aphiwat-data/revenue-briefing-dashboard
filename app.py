"""
G5 Forecast Revenue Dashboard v4 (Pro UX/UI Edition)

Professional Streamlit dashboard for hotel revenue morning review.

Run:
    python -m streamlit run g5_d4cast_revenue_dashboard_v3.py

Install:
    python -m pip install streamlit pandas openpyxl plotly
"""

import re
import html

import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go

from src.core.assets import LOGO_B64
from src.core.constants import *
from src.core.helpers import (
    fmt_pct2,
    fmt_raw2,
    fmt_signed_pct2,
    month_sort_key,
    safe_delta,
    trunc2,
)
from src.core.page import setup_page
from src.domain.helpers import (
    add_week_columns,
    apply_stay_month_filter,
    budget_delta_text,
    budget_status_from_variance,
    calc_budget_variance,
    compact_metric_table_height,
    format_compact_value,
    friendly_week_label,
    get_metric_options_with_all,
    metric_label_order,
    normalize_stay_month_selection,
    short_hotel_name,
    stay_month_label,
)
from src.data.ingest import (
    build_file_catalog_from_folder,
    build_file_catalog_from_uploads,
    build_metric_long,
    build_ref_col_map,
    parse_record,
    select_role_files,
)
from src.services.excel_export import to_excel_bytes
from src.ui.styles import inject_styles
from src.ui.sidebar import hotel_key, selected_property_chips_html


setup_page(st)
inject_styles(st)


# ============================================================
# Data Aggregation & Logic (Enhanced with Emojis)
# ============================================================



def risk_level(diff_pct):
    if pd.isna(diff_pct): return "Unknown"
    if diff_pct <= -5: return "High"
    if diff_pct <= -2: return "Medium"
    return "Low"

def build_movement_summary(metric_data, role_selection):
    role_map = {row["Role"]: row["Report Label"] for _, row in role_selection.iterrows() if pd.notna(row["Report Label"])}
    latest_label = role_map.get("Today / Latest")
    base_map = {"vs Yesterday": role_map.get("Yesterday / Previous"), "vs 7D": role_map.get("Last 7D"), "vs 1st Month": role_map.get("1st Month")}
    
    rows = []
    latest_df = metric_data[(metric_data["Report Label"] == latest_label) & (metric_data["Reference"] == "Duetto")].copy()
    
    for keys, group in latest_df.groupby(["Hotel", "Stay Month", "Metric"]):
        h, sm, m = keys
        lv = group["Value"].sum()
        for compare, base_label in base_map.items():
            if base_label is None:
                bv, diff, diff_pct, status = None, None, None, "No Base"
            else:
                bv = metric_data[(metric_data["Hotel"] == h) & (metric_data["Stay Month"] == sm) & (metric_data["Metric"] == m) & (metric_data["Report Label"] == base_label) & (metric_data["Reference"] == "Duetto")]["Value"].sum()
                if pd.isna(bv) or bv == 0:
                    diff, diff_pct, status = None, None, "No Base"
                else:
                    diff = lv - bv
                    diff_pct = diff / bv * 100
                    status = "Up" if diff > 0 else "Down" if diff < 0 else "Flat"
            rows.append({"Hotel": h, "Stay Month": sm, "Metric": m, "Compare": compare, "Latest D4cast": lv, "Base Forecast": bv, "Forecast Diff": diff, "Forecast Diff %": diff_pct, "Status": status, "Risk": risk_level(diff_pct)})
    
    out = pd.DataFrame(rows)
    if not out.empty:
        out["Compare"] = pd.Categorical(out["Compare"], ["vs Yesterday", "vs 7D", "vs 1st Month"], ordered=True)
        out = out.sort_values(["Hotel", "Stay Month", "Metric", "Compare"]).reset_index(drop=True)
    return out

def build_pace_summary(metric_data, role_selection):
    # Per (Hotel, Stay Month, Metric, Reference) keep the row from the latest
    # available Report Date - supports look-back at past stay months.
    if metric_data is None or metric_data.empty:
        return pd.DataFrame()
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
    rows = []
    for keys, group in latest.groupby(["Hotel", "Stay Month", "Metric"]):
        h, sm, m = keys
        def val(ref): x = group[group["Reference"] == ref]["Value"]; return x.sum() if not x.empty else None
        today, stly, st2y, st3y = val("Today"), val("STLY"), val("ST2Y"), val("ST3Y")
        cands = [(r, v) for r, v in [("STLY", stly), ("ST2Y", st2y), ("ST3Y", st3y)] if pd.notna(v)]
        pace_ref, pace_value = max(cands, key=lambda x: x[1]) if cands else (None, None)
        
        if not pace_value or pd.isna(today):
            diff, diff_pct, status = None, None, "No Pace"
        else:
            diff = today - pace_value
            diff_pct = diff / pace_value * 100
            status = "Ahead" if diff > 0 else "Behind" if diff < 0 else "On Pace"
        rows.append({"Hotel": h, "Stay Month": sm, "Metric": m, "Today": today, "STLY": stly, "ST2Y": st2y, "ST3Y": st3y, "Recommended Pace": pace_ref, "Recommended Pace Value": pace_value, "Pace Diff": diff, "Pace Diff %": diff_pct, "Status": status, "Risk": risk_level(diff_pct)})
    return pd.DataFrame(rows).sort_values(["Hotel", "Stay Month", "Metric"]).reset_index(drop=True) if rows else pd.DataFrame()

def build_final_comparison(metric_data, role_selection):
    # Per (Hotel, Stay Month, Metric, Reference) keep the row from the latest
    # available Report Date - supports look-back at past stay months.
    if metric_data is None or metric_data.empty:
        return pd.DataFrame()
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
    rows = []
    for keys, group in latest.groupby(["Hotel", "Stay Month", "Metric"]):
        d4 = group[group["Reference"] == "Duetto"]["Value"].sum()
        if pd.isna(d4) or d4 == 0: continue
        for final_ref in FINAL_REFS:
            base = group[group["Reference"] == final_ref]["Value"].sum()
            if pd.isna(base) or base == 0: continue
            diff = d4 - base
            rows.append({"Hotel": keys[0], "Stay Month": keys[1], "Metric": keys[2], "Forecast": d4, "Base Final": final_ref, "Final Value": base, "Diff": diff, "Diff %": diff / base * 100, "Status": "Higher" if diff > 0 else "Lower" if diff < 0 else "Equal"})
    return pd.DataFrame(rows).sort_values(["Hotel", "Stay Month", "Metric", "Base Final"]).reset_index(drop=True) if rows else pd.DataFrame()


# ============================================================
# Daily Briefing Excel - one-click morning meeting deck
# ============================================================


def build_portfolio_snapshot(metric_data, role_selection):
    """
    Portfolio totals per metric: OTB / Forecast / Budget / STLY,
    plus Forecast vs Budget % and OTB vs STLY %.
    """
    role_map = {
        row["Role"]: row["Report Label"]
        for _, row in role_selection.iterrows()
        if pd.notna(row["Report Label"])
    }
    today_label = role_map.get("Today / Latest")
    if not today_label:
        return pd.DataFrame()
    latest = metric_data[metric_data["Report Label"] == today_label]
    if latest.empty:
        return pd.DataFrame()

    rows = []
    for metric in ["Rev", "ADR", "Occ", "Room"]:
        sub = latest[latest["Metric"] == metric]
        if sub.empty:
            continue
        use_mean = metric in ("ADR", "Occ")
        agg = (lambda s: s.mean()) if use_mean else (lambda s: s.sum())
        def _v(ref):
            xs = sub[sub["Reference"] == ref]["Value"]
            return agg(xs) if not xs.empty else None
        otb, fct, bgt, stly = _v("Today"), _v("Duetto"), _v("Budget"), _v("STLY")

        def _pct(a, b):
            if a is None or b is None or pd.isna(a) or pd.isna(b) or b == 0:
                return None
            return (a - b) / abs(b) * 100

        rows.append({
            "Metric":             metric,
            "Today":              otb,
            "Duetto":             fct,
            "Budget":             bgt,
            "STLY":               stly,
            "Duetto VS BUD %":    _pct(fct, bgt),
            "Today VS BUD %":     _pct(otb, bgt),
            "Today VS STLY %":    _pct(otb, stly),
        })
    return pd.DataFrame(rows)


def build_hotel_scorecard(metric_data, role_selection):
    """
    Per-hotel scorecard: total Forecast / Budget / Var% for Rev,
    plus ADR (avg) and Occ (avg) OTB. One row per hotel.
    """
    role_map = {
        row["Role"]: row["Report Label"]
        for _, row in role_selection.iterrows()
        if pd.notna(row["Report Label"])
    }
    today_label = role_map.get("Today / Latest")
    if not today_label:
        return pd.DataFrame()
    latest = metric_data[metric_data["Report Label"] == today_label]
    if latest.empty:
        return pd.DataFrame()

    hotels = sorted(latest["Hotel"].dropna().unique())
    rows = []
    for h in hotels:
        sub = latest[latest["Hotel"] == h]
        def _get(metric, ref, mean=False):
            xs = sub[(sub["Metric"] == metric) & (sub["Reference"] == ref)]["Value"]
            if xs.empty:
                return None
            return xs.mean() if mean else xs.sum()

        rev_otb = _get("Rev", "Today")
        rev_fct = _get("Rev", "Duetto")
        rev_bgt = _get("Rev", "Budget")
        adr_otb = _get("ADR", "Today", mean=True)
        adr_bgt = _get("ADR", "Budget", mean=True)
        occ_otb = _get("Occ", "Today", mean=True)
        occ_bgt = _get("Occ", "Budget", mean=True)

        def _pct(a, b):
            if a is None or b is None or pd.isna(a) or pd.isna(b) or b == 0:
                return None
            return (a - b) / abs(b) * 100

        rows.append({
            "Hotel":                  h,
            "Rev Today":              rev_otb,
            "Rev Duetto":             rev_fct,
            "Rev Budget":             rev_bgt,
            "Rev Duetto VS BUD %":    _pct(rev_fct, rev_bgt),
            "ADR Today":              adr_otb,
            "ADR Budget":             adr_bgt,
            "ADR Today VS BUD %":     _pct(adr_otb, adr_bgt),
            "Occ Today %":            occ_otb,
            "Occ Budget %":           occ_bgt,
            "Occ Today VS BUD pts":   (occ_otb - occ_bgt) if (occ_otb is not None and occ_bgt is not None) else None,
        })
    return pd.DataFrame(rows)


def build_daily_briefing_sheets(metric_data, role_selection, report_file_month):
    """
    Assemble the full multi-sheet daily briefing for the morning meeting.

    Sheet order is curated for the meeting flow:
      1. Portfolio Snapshot     - the headline numbers
      2. Hotel Scorecard        - per-hotel summary
      3. Variance Pivot         - detailed pivot with all variance %
      4. Same-Time Pace         - OTB vs STLY / ST2Y / ST3Y
      5. Historical Final       - Forecast vs Final LY / 2Y / 3Y
      6. Forecast Movement      - vs 1D / 7D / First-day
      7. Hotel Momentum         - daily forecast pickup per hotel
      8. Role Selection         - file roles validation
    """
    sheets = {}

    # 1. Portfolio Snapshot
    sheets["Portfolio Snapshot"] = build_portfolio_snapshot(metric_data, role_selection)

    # 2. Hotel Scorecard
    sheets["Hotel Scorecard"] = build_hotel_scorecard(metric_data, role_selection)

    # 3. Variance Pivot (the big OTB/Bgt/Duetto/STLY/Final variance table)
    if "build_variance_pivot_table" in globals():
        sheets["Variance Pivot"] = build_variance_pivot_table(metric_data, role_selection)

    # 4. Same-Time Pace
    if "build_pace_summary" in globals():
        sheets["Same-Time Pace"] = build_pace_summary(metric_data, role_selection)

    # 5. Historical Final
    if "build_final_comparison" in globals():
        sheets["Historical Final"] = build_final_comparison(metric_data, role_selection)

    # 6. Forecast Movement
    if "build_forecast_movement_v31" in globals():
        sheets["Duetto Movement"] = build_forecast_movement_v31(metric_data, role_selection)
    elif "build_movement_summary" in globals():
        sheets["Duetto Movement"] = build_movement_summary(metric_data, role_selection)

    # 7. Hotel Momentum (daily forecast pickup per hotel, per stay month)
    d4 = metric_data[metric_data["Reference"] == "Duetto"].copy()
    if not d4.empty:
        d4 = d4.sort_values(["Hotel", "Stay Month", "Metric", "Report Date"])
        d4["Previous Value"]  = d4.groupby(["Hotel", "Stay Month", "Metric"])["Value"].shift(1)
        d4["Daily Pickup"]    = d4["Value"] - d4["Previous Value"]
        d4["Daily Pickup %"]  = d4["Daily Pickup"] / d4["Previous Value"] * 100
        sheets["Hotel Momentum"] = d4[[
            "Hotel", "Stay Month", "Metric", "Report Date", "Report Label",
            "Value", "Previous Value", "Daily Pickup", "Daily Pickup %",
        ]].reset_index(drop=True)

    # 8. Role Selection
    sheets["Role Selection"] = role_selection

    return sheets


def render_hotel_table(df, view_mode, key_prefix, column_config=None):
    """
    Render dataframe either as one list table or separated hotel tabs.
    This supports the tab-view layout requested in the dashboard note.
    """
    if df is None or df.empty:
        st.info("No data.")
        return

    if view_mode == "Hotel tabs":
        hotels = sorted(df["Hotel"].dropna().unique()) if "Hotel" in df.columns else []
        if not hotels:
            st.dataframe(df, use_container_width=True, hide_index=True, column_config=column_config)
            return

        tabs = st.tabs(hotels)
        for tab, hotel in zip(tabs, hotels):
            with tab:
                sub = df[df["Hotel"] == hotel].copy()
                st.dataframe(sub, use_container_width=True, hide_index=True, column_config=column_config)
    else:
        st.dataframe(df, use_container_width=True, hide_index=True, column_config=column_config)







def build_latest_pivot_table(metric_data, role_selection):
    """
    Build compact pivot table like the reference screenshot:
    Month | Metric | Today | STLY | ST2Y | ST3Y | Duetto | Budget | Final LY | Final 2Y | Final 3Y

    Per-(Hotel, Stay Month, Metric, Reference) we keep the LATEST available
    Report Date row. This ensures that a past stay month (e.g. May while
    we're on June's report) still shows data taken from its OWN latest
    report file - fixing the "No data in latest report" issue when looking
    back at past months.
    Keeps metric order: Occ, Room, ADR, Rev.
    """
    if metric_data is None or metric_data.empty:
        return pd.DataFrame()

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
        return pd.DataFrame()

    pivot = (
        latest.pivot_table(
            index=["Hotel", "Stay Month", "Metric"],
            columns="Reference",
            values="Value",
            aggfunc="sum",
        )
        .reset_index()
    )

    ref_order = ["Today", "STLY", "ST2Y", "ST3Y", "Duetto", "Budget", "Final LY", "Final 2Y", "Final 3Y"]
    available_refs = [c for c in ref_order if c in pivot.columns]

    pivot["Metric"] = pd.Categorical(pivot["Metric"], categories=METRIC_ORDER, ordered=True)
    pivot["Stay Month Sort"] = pivot["Stay Month"].apply(month_sort_key)
    pivot = pivot.sort_values(["Hotel", "Stay Month Sort", "Metric"]).drop(columns=["Stay Month Sort"])

    return pivot[["Hotel", "Stay Month", "Metric"] + available_refs].reset_index(drop=True)


def build_variance_pivot_table(metric_data, role_selection):
    """
    Extended pivot: raw values PLUS colour-coded variance % columns.

    Added columns (only when both sides are available):
        Today VS BUD       - Today (OTB) vs Budget %
        Duetto VS BUD      - Duetto vs Budget %                 -> headline variance
        Today VS Duetto    - Today vs Duetto (pace toward forecast)
        Today VS STLY      - Today vs best same-time (STLY / ST2Y / ST3Y) %
        Duetto VS Final LY - Duetto vs Final LY %
        Duetto VS Final 2Y - Duetto vs Final 2Y %
        Duetto VS Final 3Y - Duetto vs Final 3Y %

    Column order:
      Today | Budget | Today VS BUD | Duetto | Duetto VS BUD | Today VS Duetto |
      STLY | Today VS STLY | Final LY | Duetto VS Final LY | Final 2Y | ... | Final 3Y | ...
    """
    pivot = build_latest_pivot_table(metric_data, role_selection)
    if pivot.empty:
        return pivot

    def safe_pct(num, denom):
        try:
            n, d = float(num), float(denom)
            if d == 0 or pd.isna(n) or pd.isna(d):
                return pd.NA
            return (n - d) / abs(d) * 100
        except (TypeError, ValueError):
            return pd.NA

    if "Today" in pivot.columns and "Budget" in pivot.columns:
        pivot["Today VS BUD"] = pivot.apply(lambda r: safe_pct(r["Today"], r["Budget"]), axis=1)

    # Headline: Duetto vs Budget %
    if "Duetto" in pivot.columns and "Budget" in pivot.columns:
        pivot["Duetto VS BUD"] = pivot.apply(lambda r: safe_pct(r["Duetto"], r["Budget"]), axis=1)

    if "Today" in pivot.columns and "Duetto" in pivot.columns:
        pivot["Today VS Duetto"] = pivot.apply(lambda r: safe_pct(r["Today"], r["Duetto"]), axis=1)

    # -- Same-Time variances - each year compared INDEPENDENTLY --
    # FIX (Bug: Arbour STLY sign flip):
    #   Old code computed "Today VS STLY" against MAX(STLY, ST2Y, ST3Y), which
    #   could flip signs when ST2Y/ST3Y were higher than STLY. Now we compute
    #   each variance against its OWN reference year - STLY vs STLY only,
    #   ST2Y vs ST2Y only, ST3Y vs ST3Y only. No more "best-of-3" surprise.
    if "Today" in pivot.columns:
        for ref_col, var_col in [
            ("STLY", "Today VS STLY"),
            ("ST2Y", "Today VS ST2Y"),
            ("ST3Y", "Today VS ST3Y"),
        ]:
            if ref_col in pivot.columns:
                pivot[var_col] = pivot.apply(
                    lambda r, _c=ref_col: safe_pct(r["Today"], r[_c]),
                    axis=1,
                )

    if "Duetto" in pivot.columns:
        for ref_col, var_col in [
            ("Final LY", "Duetto VS Final LY"),
            ("Final 2Y", "Duetto VS Final 2Y"),
            ("Final 3Y", "Duetto VS Final 3Y"),
        ]:
            if ref_col in pivot.columns:
                pivot[var_col] = pivot.apply(
                    lambda r, _c=ref_col: safe_pct(r["Duetto"], r[_c]),
                    axis=1,
                )

    # Build column order
    base = [c for c in ["Hotel", "Stay Month", "Metric"] if c in pivot.columns]
    ordered = []
    for col in ["Today", "Budget", "Today VS BUD", "Duetto", "Duetto VS BUD", "Today VS Duetto"]:
        if col in pivot.columns:
            ordered.append(col)
    # Same-time benchmarks - each with its OWN variance (no more best-of-3)
    for ref_col, var_col in [
        ("STLY", "Today VS STLY"),
        ("ST2Y", "Today VS ST2Y"),
        ("ST3Y", "Today VS ST3Y"),
    ]:
        if ref_col in pivot.columns:
            ordered.append(ref_col)
        if var_col in pivot.columns:
            ordered.append(var_col)
    # Historical finals
    for ref_col, var_col in [
        ("Final LY", "Duetto VS Final LY"),
        ("Final 2Y", "Duetto VS Final 2Y"),
        ("Final 3Y", "Duetto VS Final 3Y"),
    ]:
        if ref_col in pivot.columns:
            ordered.append(ref_col)
        if var_col in pivot.columns:
            ordered.append(var_col)

    all_cols = base + ordered
    return pivot[[c for c in all_cols if c in pivot.columns]].reset_index(drop=True)


def style_latest_pivot_table(df):
    """
    Pro-readable compact table styling:
     Duetto column = always tinted blue (focal - current forecast)
     Budget column = always tinted amber (target)
     Rev row      = headline background (stronger orange)
     Duetto vs Budget on Rev row = green if above target, red if below
    """
    if df is None or df.empty:
        return df

    numeric_cols = [c for c in ["Today", "STLY", "ST2Y", "ST3Y", "Duetto", "Budget", "Final LY", "Final 2Y", "Final 3Y"] if c in df.columns]
    fmt = {c: fmt_raw2 for c in numeric_cols if "fmt_raw2" in globals()}

    def row_style(row):
        styles = pd.Series("", index=row.index)
        metric = str(row.get("Metric", ""))

        # -- Base banding by metric row ------------------------
        if metric == "Occ":
            styles[:] = "background-color: #f9fafb"
        elif metric == "Rev":
            styles[:] = "background-color: #fff7ed; font-weight: 500"

        # -- Always emphasise the focal columns ----------------
        if "Duetto" in row.index:
            base = "background-color: #eff6ff; font-weight: 600; color: #1e40af"
            if metric == "Rev":
                base = "background-color: #dbeafe; font-weight: 700; color: #1e3a8a"
            styles["Duetto"] = base

        if "Budget" in row.index:
            base = "background-color: #fefce8; color: #713f12"
            if metric == "Rev":
                base = "background-color: #fef3c7; font-weight: 600; color: #713f12"
            styles["Budget"] = base

        # -- Headline call-out: Duetto vs Budget on Rev row ----
        if metric == "Rev" and "Duetto" in row.index and "Budget" in row.index:
            d = row.get("Duetto")
            b = row.get("Budget")
            if pd.notna(d) and pd.notna(b):
                if d < b:
                    styles["Duetto"] = "background-color: #fecaca; font-weight: 700; color: #991b1b"
                elif d > b:
                    styles["Duetto"] = "background-color: #bbf7d0; font-weight: 700; color: #166534"

        return styles

    return df.style.format(fmt).apply(row_style, axis=1)


def style_variance_pivot(df):
    """
    Style the variance pivot table.
     Today  -> blue tint (focal - on-the-book)
     Duetto -> green tint (forecast)
     Budget -> amber tint (target)
     VS variance columns -> green if positive, red if negative, yellow if zero
     Rev row -> slightly warmer background for row emphasis
    """
    if df is None or df.empty:
        return df

    var_cols = [c for c in df.columns if " VS " in str(c)]
    raw_cols = [c for c in ["Today", "STLY", "ST2Y", "ST3Y", "Duetto", "Budget",
                             "Final LY", "Final 2Y", "Final 3Y"] if c in df.columns]

    def _fmt_pct(x):
        try:
            v = float(x)
            return f"{v:+.1f}%"
        except (TypeError, ValueError):
            return "-"

    fmt = {c: fmt_raw2 for c in raw_cols}
    for c in var_cols:
        fmt[c] = _fmt_pct

    def row_style(row):
        styles = pd.Series("", index=row.index)
        metric = str(row.get("Metric", ""))

        if metric == "Rev":
            styles[:] = "background-color: #fffbf5"
        elif metric == "Occ":
            styles[:] = "background-color: #f9fafb"

        # Key column highlighting
        if "Today" in row.index:
            base = "background-color: #eff6ff; font-weight: 600; color: #1e40af"
            if metric == "Rev":
                base = "background-color: #dbeafe; font-weight: 700; color: #1e3a8a"
            styles["Today"] = base

        if "Duetto" in row.index:
            base = "background-color: #f0fdf4; font-weight: 600; color: #15803d"
            if metric == "Rev":
                base = "background-color: #dcfce7; font-weight: 700; color: #14532d"
            styles["Duetto"] = base

        if "Budget" in row.index:
            styles["Budget"] = "background-color: #fefce8; color: #713f12"

        # Variance columns
        for col in var_cols:
            if col not in row.index:
                continue
            try:
                v = float(row[col])
                if v > 0:
                    styles[col] = "background-color: #bbf7d0; color: #166534; font-weight: 700"
                elif v < 0:
                    styles[col] = "background-color: #fecaca; color: #991b1b; font-weight: 700"
                else:
                    styles[col] = "background-color: #fef9c3; color: #92400e; font-weight: 600"
            except (TypeError, ValueError):
                pass

        return styles

    return df.style.format(fmt, na_rep="-").apply(row_style, axis=1)


def style_final_variance_table(df):
    """
    Color-code the Historical Final Comparison table by Status (Higher / Lower / Equal).
    Green = Forecast above historical final, Red = below.
    """
    if df is None or df.empty:
        return df

    def row_style(row):
        styles = pd.Series("", index=row.index)
        status = str(row.get("Status", ""))
        if status == "Higher":
            color = "background-color: #bbf7d0; color: #166534; font-weight: 600"
        elif status == "Lower":
            color = "background-color: #fecaca; color: #991b1b; font-weight: 600"
        else:
            color = "background-color: #fef9c3; color: #92400e"
        for col in ["Variance", "Variance %", "Status"]:
            if col in styles.index:
                styles[col] = color
        return styles

    return df.style.apply(row_style, axis=1)


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


def render_compact_hotel_tabs(pivot_df):
    """
    Duetto pivot table - variance-enhanced view.
    Columns: Today | Budget | Today VS BUD | Duetto | Duetto VS BUD | Today VS Duetto |
             STLY | Today VS STLY | Final LY | Duetto VS Final LY | ... + more.
    """
    if pivot_df is None or pivot_df.empty:
        st.info("No pivot data for selected filters.")
        return

    st.markdown('<div class="section-title">Duetto Pivot - by Stay Month</div>', unsafe_allow_html=True)

    c_view, c_legend = st.columns([1.6, 3])
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


def build_hotel_leaderboard(metric_long, role_selection, hotels, stay_month_selection, leaderboard_metric):
    """
    Hotel leaderboard for comparing property performance.

    User can choose metric:
    Occ / Room / ADR / Rev

    Leaderboard includes:
    - Latest Forecast
    - Previous Forecast
    - Daily PU
    - Daily PU %
    - 7D PU
    - 7D PU %
    - MTD PU
    - MTD PU %
    - Status
    """
    if metric_long is None or metric_long.empty:
        return pd.DataFrame()

    role_map = {
        row["Role"]: row["Report Label"]
        for _, row in role_selection.iterrows()
        if pd.notna(row["Report Label"])
    }

    today_label = role_map.get("Today / Latest")
    yday_label = role_map.get("Yesterday / Previous")
    seven_label = role_map.get("Last 7D")
    first_label = role_map.get("1st Month")

    df = metric_long[
        (metric_long["Hotel"].isin(hotels))
        & (metric_long["Metric"] == leaderboard_metric)
        & (metric_long["Reference"] == "Duetto")
    ].copy()

    df = apply_stay_month_filter(df, stay_month_selection)

    if df.empty:
        return pd.DataFrame()

    group_cols = ["Hotel"]

    def get_by_label(label, value_name):
        if label is None:
            return pd.DataFrame({"Hotel": [], value_name: []})
        out = (
            df[df["Report Label"] == label]
            .groupby(group_cols, as_index=False)["Value"]
            .sum()
            .rename(columns={"Value": value_name})
        )
        return out

    latest = get_by_label(today_label, "Latest D4cast")
    previous = get_by_label(yday_label, "Previous Forecast")
    seven = get_by_label(seven_label, "7D Base Forecast")
    first = get_by_label(first_label, "1st Month Base Forecast")

    out = latest.merge(previous, on="Hotel", how="left")
    out = out.merge(seven, on="Hotel", how="left")
    out = out.merge(first, on="Hotel", how="left")

    out["Daily PU"] = out["Latest D4cast"] - out["Previous Forecast"]
    out["Daily PU %"] = out["Daily PU"] / out["Previous Forecast"] * 100

    out["7D PU"] = out["Latest D4cast"] - out["7D Base Forecast"]
    out["7D PU %"] = out["7D PU"] / out["7D Base Forecast"] * 100

    out["MTD PU"] = out["Latest D4cast"] - out["1st Month Base Forecast"]
    out["MTD PU %"] = out["MTD PU"] / out["1st Month Base Forecast"] * 100

    def status_from_daily(x):
        if pd.isna(x):
            return "No Base"
        if x > 0:
            return "Up"
        if x < 0:
            return "Down"
        return "Flat"

    out["Status"] = out["Daily PU"].apply(status_from_daily)
    out["Abs Daily PU"] = out["Daily PU"].abs()
    out["Rank"] = out["Latest D4cast"].rank(method="dense", ascending=False).astype(int)

    return out.sort_values("Latest D4cast", ascending=False).reset_index(drop=True)


def render_hotel_leaderboard(metric_long, role_selection, selected_hotels, stay_month_selection):
    st.markdown('<div class="section-title">Hotel Leaderboard</div>', unsafe_allow_html=True)
    st.caption("Compare hotels by selected metric. Use colors to quickly spot winners, drops, and movement.")

    c1, c2, c3 = st.columns([1, 1.2, 1])

    leaderboard_metric = c1.selectbox(
        "Leaderboard metric",
        ["Occ", "Room", "ADR", "Rev"],
        index=0,
        key="leaderboard_metric",
    )

    rank_by = c2.selectbox(
        "Rank by",
        [
            "Latest D4cast",
            "Daily PU %",
            "7D PU %",
            "MTD PU %",
        ],
        index=1,
        key="leaderboard_rank_by",
    )

    display_count = c3.selectbox(
        "Show",
        ["Top 5", "Top 8", "Top 10", "Top 15", "All"],
        index=1,
        key="leaderboard_display_count",
    )

    leaderboard = build_hotel_leaderboard(
        metric_long=metric_long,
        role_selection=role_selection,
        hotels=selected_hotels,
        stay_month_selection=stay_month_selection,
        leaderboard_metric=leaderboard_metric,
    )

    if leaderboard.empty:
        st.info("No leaderboard data for selected filters.")
        return pd.DataFrame()

    ascending = rank_by in ["Daily PU %", "7D PU %", "MTD PU %"] and st.checkbox(
        "Show worst first",
        value=True,
        key="leaderboard_worst_first",
    )

    leaderboard = leaderboard.sort_values(rank_by, ascending=ascending).reset_index(drop=True)

    if display_count != "All":
        n = int(display_count.replace("Top ", ""))
        leaderboard_view = leaderboard.head(n).copy()
    else:
        leaderboard_view = leaderboard.copy()

    k1, k2, k3, k4 = st.columns(4)

    total_latest = leaderboard["Latest D4cast"].sum()
    total_daily_pu = leaderboard["Daily PU"].sum()
    hotels_up = (leaderboard["Daily PU"] > 0).sum()
    hotels_down = (leaderboard["Daily PU"] < 0).sum()

    k1.metric(f"Total {leaderboard_metric} Latest Forecast", fmt_raw2(total_latest))
    k2.metric("Total Daily PU", fmt_raw2(total_daily_pu))
    k3.metric("Hotels Up", int(hotels_up))
    k4.metric("Hotels Down", int(hotels_down))

    chart_df = leaderboard_view.copy()
    chart_df["Color Status"] = chart_df["Daily PU"].apply(
        lambda x: "Up" if pd.notna(x) and x > 0 else "Down" if pd.notna(x) and x < 0 else "Flat"
    )

    color_map = {
        "Up": "#15803d",
        "Down": "#b91c1c",
        "Flat": "#ca8a04",
    }

    fig = px.bar(
        chart_df,
        x=rank_by,
        y="Hotel",
        orientation="h",
        color="Color Status",
        color_discrete_map=color_map,
        hover_data={
            "Latest D4cast": ":,.2f",
            "Previous Forecast": ":,.2f",
            "Daily PU": ":,.2f",
            "Daily PU %": ":.2f",
            "7D PU": ":,.2f",
            "7D PU %": ":.2f",
            "MTD PU": ":,.2f",
            "MTD PU %": ":.2f",
            "Color Status": False,
        },
        title=f"Hotel Leaderboard by {rank_by} ({leaderboard_metric})",
    )

    fig.update_layout(
        plot_bgcolor="rgba(0,0,0,0)",
        paper_bgcolor="rgba(0,0,0,0)",
        height=max(420, 48 * len(chart_df)),
        xaxis=dict(showgrid=True, gridcolor="#f1f5f9"),
        yaxis=dict(categoryorder="total ascending"),
        legend_title_text="Daily PU Status",
        margin=dict(l=20, r=20, t=60, b=20),
    )

    fig.update_traces(
        marker_line_width=0,
        texttemplate="%{x:,.2f}",
        textposition="outside",
        cliponaxis=False,
    )

    st.plotly_chart(
        fig,
        use_container_width=True,
        config={"displayModeBar": True, "displaylogo": False},
    )

    st.markdown("#### Leaderboard Table")

    table_view = leaderboard_view[[
        "Hotel",
        "Latest D4cast",
        "Previous Forecast",
        "Daily PU %",
        "7D PU %",
        "MTD PU %",
        "Status",
    ]].copy()

    st.dataframe(
        table_view,
        use_container_width=True,
        hide_index=True,
        height=min(560, 44 + 36 * len(table_view)),
        column_config={
            "Latest D4cast": st.column_config.NumberColumn("Latest D4cast", format="%,.2f"),
            "Previous Forecast": st.column_config.NumberColumn("Previous Forecast", format="%,.2f"),
            "Daily PU %": st.column_config.NumberColumn("Daily PU %", format="%+.2f%%"),
            "7D PU %": st.column_config.NumberColumn("7D PU %", format="%+.2f%%"),
            "MTD PU %": st.column_config.NumberColumn("MTD PU %", format="%+.2f%%"),
            "Status": st.column_config.TextColumn("Status"),
        },
    )

    return leaderboard


def build_hotel_compare_matrix(metric_long, role_selection, hotels, stay_month_selection, compare_metric):
    """
    Build hotel comparison matrix for color-coded leaderboard.
    This is the old-style useful comparison view:
    each hotel is a row, each key metric is a column.
    """
    lb = build_hotel_leaderboard(
        metric_long=metric_long,
        role_selection=role_selection,
        hotels=hotels,
        stay_month_selection=stay_month_selection,
        leaderboard_metric=compare_metric,
    )

    if lb.empty:
        return pd.DataFrame()

    cols = [
        "Hotel",
        "Latest D4cast",
        "Daily PU",
        "Daily PU %",
        "7D PU",
        "7D PU %",
        "MTD PU",
        "MTD PU %",
        "Status",
    ]

    available = [c for c in cols if c in lb.columns]
    out = lb[available].copy()

    # Default order: worst Daily PU first, because morning meeting usually wants risk first.
    if "Daily PU" in out.columns:
        out = out.sort_values("Daily PU", ascending=True).reset_index(drop=True)

    return out


def style_compare_matrix(df):
    """
    Color-coded comparison table:
    green = positive, red = negative, blue = flat/neutral.
    """
    if df is None or df.empty:
        return df

    numeric_cols = [
        "Latest D4cast",
        "Daily PU",
        "Daily PU %",
        "7D PU",
        "7D PU %",
        "MTD PU",
        "MTD PU %",
    ]

    fmt = {}
    for c in numeric_cols:
        if c in df.columns:
            if "%" in c:
                fmt[c] = lambda x: "" if pd.isna(x) else fmt_pct2(x)
            else:
                fmt[c] = lambda x: "" if pd.isna(x) else fmt_raw2(x)

    def color_cell(val, col_name):
        if pd.isna(val):
            return "background-color: #f8fafc; color: #64748b"

        # Latest Forecast is a magnitude metric, not positive/negative performance.
        if col_name == "Latest D4cast":
            return "background-color: #eef2ff; color: #1e293b; font-weight: 600"

        if isinstance(val, (int, float)):
            if val > 0:
                return "background-color: #bbf7d0; color: #14532d; font-weight: 600"
            if val < 0:
                return "background-color: #fecaca; color: #7f1d1d; font-weight: 600"
            return "background-color: #fef08a; color: #713f12; font-weight: 600"

        return ""

    def apply_style(data):
        styles = pd.DataFrame("", index=data.index, columns=data.columns)

        for c in data.columns:
            if c in numeric_cols:
                styles[c] = data[c].apply(lambda x: color_cell(x, c))

        if "Status" in data.columns:
            styles["Status"] = data["Status"].apply(
                lambda x: "background-color: #bbf7d0; color: #14532d; font-weight: 700"
                if "Up" in str(x)
                else "background-color: #fecaca; color: #7f1d1d; font-weight: 700"
                if "Down" in str(x)
                else "background-color: #fef08a; color: #713f12; font-weight: 700"
            )

        return styles

    return df.style.format(fmt).apply(apply_style, axis=None)


def compact_number_display_df(df, cols):
    """
    For cramped tables: convert selected numeric columns to truncated 2-decimal strings.
    This prevents visual clipping like ',403.00' when column is too narrow.
    """
    out = df.copy()
    for c in cols:
        if c in out.columns:
            out[c] = out[c].apply(lambda x: "" if pd.isna(x) else fmt_raw2(x))
    return out








def render_metric_sections(
    df,
    title,
    render_func,
    metric_col="Metric",
    default_expand_all=True,
):
    """
    Render All Metrics safely without making one huge wide/clipped table.
    Each metric gets its own section in business-friendly order:
    Occ -> Room -> ADR -> Rev.
    """
    st.markdown(f"#### {title}")

    if df is None or df.empty:
        st.info("No data.")
        return

    if metric_col not in df.columns:
        render_func(df)
        return

    tabs = st.tabs(metric_label_order())

    for tab, metric in zip(tabs, metric_label_order()):
        with tab:
            sub = df[df[metric_col] == metric].copy()
            if sub.empty:
                st.info(f"No {metric} data.")
            else:
                render_func(sub)


def build_hotel_compare_matrix_all(metric_long, role_selection, hotels, stay_month_selection):
    """
    Build comparison matrix for all metrics.
    Output is long-form, not ultra-wide, so numbers do not get clipped.
    """
    frames = []
    for m in metric_label_order():
        one = build_hotel_compare_matrix(
            metric_long=metric_long,
            role_selection=role_selection,
            hotels=hotels,
            stay_month_selection=stay_month_selection,
            compare_metric=m,
        )
        if not one.empty:
            one.insert(1, "Metric", m)
            frames.append(one)

    if not frames:
        return pd.DataFrame()

    return pd.concat(frames, ignore_index=True)


def render_no_clip_dataframe(df, height=None):
    """
    Render dataframe with readable width.
    Streamlit sometimes clips numeric text in narrow numeric columns.
    For presentation tables, pre-format numeric values as strings.
    """
    if df is None or df.empty:
        st.info("No data.")
        return

    show = df.copy()
    numeric_cols = [
        "Latest D4cast",
        "Previous Forecast",
        "Base Forecast",
        "Forecast Diff",
        "Forecast Diff %",
        "Daily PU",
        "Daily PU %",
        "7D PU",
        "7D PU %",
        "MTD PU",
        "MTD PU %",
        "Today",
        "STLY",
        "ST2Y",
        "ST3Y",
        "Recommended Pace Value",
        "Pace Diff",
        "Pace Diff %",
        "Forecast",
        "Final Value",
        "Diff",
        "Diff %",
    ]

    for c in numeric_cols:
        if c in show.columns:
            if "%" in c:
                show[c] = show[c].apply(lambda x: "" if pd.isna(x) else fmt_pct2(x))
            else:
                show[c] = show[c].apply(lambda x: "" if pd.isna(x) else fmt_raw2(x))

    st.dataframe(
        show,
        use_container_width=True,
        hide_index=True,
        height=height if height is not None else compact_metric_table_height(show),
    )


def render_color_matrix_no_clip(matrix):
    """
    Color table, but numeric columns are displayed as strings to prevent clipping.
    Keep color based on original numeric values.
    """
    if matrix is None or matrix.empty:
        st.info("No leaderboard data.")
        return

    raw = matrix.copy()
    show = matrix.copy()

    numeric_cols = [
        "Latest D4cast",
        "Daily PU",
        "Daily PU %",
        "7D PU",
        "7D PU %",
        "MTD PU",
        "MTD PU %",
    ]

    for c in numeric_cols:
        if c in show.columns:
            if "%" in c:
                show[c] = show[c].apply(lambda x: "" if pd.isna(x) else fmt_pct2(x))
            else:
                show[c] = show[c].apply(lambda x: "" if pd.isna(x) else fmt_raw2(x))

    def apply_style(_):
        styles = pd.DataFrame("", index=show.index, columns=show.columns)

        for c in numeric_cols:
            if c not in raw.columns:
                continue
            for idx, val in raw[c].items():
                if pd.isna(val):
                    styles.loc[idx, c] = "background-color: #f8fafc; color: #64748b"
                elif c == "Latest D4cast":
                    styles.loc[idx, c] = "background-color: #eef2ff; color: #1e293b; font-weight: 600"
                elif val > 0:
                    styles.loc[idx, c] = "background-color: #bbf7d0; color: #14532d; font-weight: 600"
                elif val < 0:
                    styles.loc[idx, c] = "background-color: #fecaca; color: #7f1d1d; font-weight: 600"
                else:
                    styles.loc[idx, c] = "background-color: #fef08a; color: #713f12; font-weight: 600"

        if "Status" in show.columns:
            for idx, val in show["Status"].items():
                if "Up" in str(val):
                    styles.loc[idx, "Status"] = "background-color: #bbf7d0; color: #14532d; font-weight: 700"
                elif "Down" in str(val):
                    styles.loc[idx, "Status"] = "background-color: #fecaca; color: #7f1d1d; font-weight: 700"
                else:
                    styles.loc[idx, "Status"] = "background-color: #fef08a; color: #713f12; font-weight: 700"

        return styles

    st.dataframe(
        show.style.apply(apply_style, axis=None),
        use_container_width=True,
        hide_index=True,
        height=compact_metric_table_height(show, row_height=38, max_height=620),
    )










def make_recommended_pace_compact(df):
    """
    Compact view for Recommended Pace.
    Avoids wide tables that require horizontal scrolling.
    """
    if df is None or df.empty:
        return pd.DataFrame()

    out = df.copy()

    cols = []
    for _, r in out.iterrows():
        benchmark_text = []
        for ref in ["STLY", "ST2Y", "ST3Y"]:
            if ref in out.columns:
                benchmark_text.append(f"{ref}: {format_compact_value(r.get(ref))}")

        cols.append({
            "Hotel": r.get("Hotel"),
            "Stay Month": r.get("Stay Month"),
            "Metric": r.get("Metric"),
            "Today": format_compact_value(r.get("Today")),
            "Recommended Pace": r.get("Recommended Pace"),
            "Rec. Pace Value": format_compact_value(r.get("Recommended Pace Value")),
            "Variance": format_compact_value(r.get("Pace Diff")),
            "Variance %": format_compact_value(r.get("Pace Diff %"), is_pct=True),
            "Status": r.get("Status"),
            "Benchmarks": " | ".join(benchmark_text),
        })

    result = pd.DataFrame(cols)
    if "Metric" in result.columns:
        result["Metric"] = pd.Categorical(result["Metric"], categories=METRIC_ORDER, ordered=True)
        result = result.sort_values(["Hotel", "Stay Month", "Metric"]).reset_index(drop=True)

    return result


def make_movement_compact(df):
    """
    Compact view for Movement.
    """
    if df is None or df.empty:
        return pd.DataFrame()

    rows = []
    for _, r in df.iterrows():
        rows.append({
            "Hotel": r.get("Hotel"),
            "Stay Month": r.get("Stay Month"),
            "Metric": r.get("Metric"),
            "Compare": r.get("Compare"),
            "Latest D4cast": format_compact_value(r.get("Latest D4cast")),
            "Base Forecast": format_compact_value(r.get("Base Forecast")),
            "Variance": format_compact_value(r.get("Forecast Diff")),
            "Variance %": format_compact_value(r.get("Forecast Diff %"), is_pct=True),
            "Status": r.get("Status"),
            "Risk": r.get("Risk"),
        })

    result = pd.DataFrame(rows)
    if "Metric" in result.columns:
        result["Metric"] = pd.Categorical(result["Metric"], categories=METRIC_ORDER, ordered=True)
        result = result.sort_values(["Hotel", "Stay Month", "Metric", "Compare"]).reset_index(drop=True)

    return result


def make_final_compact(df):
    """
    Compact view for D4cast vs Final.
    """
    if df is None or df.empty:
        return pd.DataFrame()

    rows = []
    for _, r in df.iterrows():
        rows.append({
            "Hotel": r.get("Hotel"),
            "Stay Month": r.get("Stay Month"),
            "Metric": r.get("Metric"),
            "Base Final": r.get("Base Final"),
            "Forecast": format_compact_value(r.get("Forecast")),
            "Final Value": format_compact_value(r.get("Final Value")),
            "Variance": format_compact_value(r.get("Diff")),
            "Variance %": format_compact_value(r.get("Diff %"), is_pct=True),
            "Status": r.get("Status"),
        })

    result = pd.DataFrame(rows)
    if "Metric" in result.columns:
        result["Metric"] = pd.Categorical(result["Metric"], categories=METRIC_ORDER, ordered=True)
        result = result.sort_values(["Hotel", "Stay Month", "Metric", "Base Final"]).reset_index(drop=True)

    return result


def render_compact_by_hotel(df, view_mode, key_prefix, height_cap=620):
    """
    Compact table renderer with no horizontal scroll.
    Uses hotel tabs if requested. Each tab only shows one hotel's compact table.
    """
    if df is None or df.empty:
        st.info("No data.")
        return

    if view_mode == "Hotel tabs" and "Hotel" in df.columns:
        hotels = sorted(df["Hotel"].dropna().unique())
        tabs = st.tabs([short_hotel_name(h) if "short_hotel_name" in globals() else str(h) for h in hotels])

        for tab, hotel in zip(tabs, hotels):
            with tab:
                sub = df[df["Hotel"] == hotel].drop(columns=["Hotel"]).reset_index(drop=True)
                if key_prefix == "pace":
                    st.dataframe(
                        style_pace_variance_table(sub),
                        use_container_width=True,
                        hide_index=True,
                        height=min(height_cap, 48 + 38 * len(sub)),
                    )
                elif key_prefix == "final":
                    st.dataframe(
                        style_final_variance_table(sub),
                        use_container_width=True,
                        hide_index=True,
                        height=min(height_cap, 48 + 38 * len(sub)),
                    )
                else:
                    st.dataframe(
                        sub,
                        use_container_width=True,
                        hide_index=True,
                        height=min(height_cap, 48 + 38 * len(sub)),
                    )
    else:
        show = df.copy()
        if "Hotel" in show.columns and "short_hotel_name" in globals():
            show["Hotel"] = show["Hotel"].apply(short_hotel_name)

        if key_prefix == "pace":
            st.dataframe(
                style_pace_variance_table(show),
                use_container_width=True,
                hide_index=True,
                height=min(height_cap, 48 + 34 * len(show)),
            )
        elif key_prefix == "final":
            st.dataframe(
                style_final_variance_table(show),
                use_container_width=True,
                hide_index=True,
                height=min(height_cap, 48 + 34 * len(show)),
            )
        else:
            st.dataframe(
                show,
                use_container_width=True,
                hide_index=True,
                height=min(height_cap, 48 + 34 * len(show)),
            )


def render_pace_cards(df):
    """
    Presentation-friendly cards for Recommended Pace.
    Useful when the table is still too wide.
    """
    if df is None or df.empty:
        st.info("No pace data.")
        return

    compact = make_recommended_pace_compact(df)

    hotels = sorted(compact["Hotel"].dropna().unique())
    tabs = st.tabs([short_hotel_name(h) if "short_hotel_name" in globals() else str(h) for h in hotels])

    for tab, hotel in zip(tabs, hotels):
        with tab:
            sub = compact[compact["Hotel"] == hotel].copy()
            for _, row in sub.iterrows():
                status = str(row.get("Status", ""))
                border = "#15803d" if "Ahead" in status or "Up" in status else "#b91c1c" if "Behind" in status or "Down" in status else "#ca8a04"

                st.markdown(
                    f"""
                    <div style="
                        border-left: 6px solid {border};
                        border-radius: 12px;
                        padding: 14px 16px;
                        margin-bottom: 10px;
                        background: #ffffff;
                        box-shadow: 0 1px 4px rgba(15,23,42,0.08);
                    ">
                        <div style="font-weight:700; font-size:1.02rem; margin-bottom:6px;">
                            {row['Stay Month']}  {row['Metric']}  {row['Status']}
                        </div>
                        <div style="display:grid; grid-template-columns: repeat(4, minmax(0,1fr)); gap:10px;">
                            <div><span style="color:#64748b;">Today</span><br><b>{row['Today']}</b></div>
                            <div><span style="color:#64748b;">Recommended</span><br><b>{row['Recommended Pace']}</b></div>
                            <div><span style="color:#64748b;">Variance</span><br><b>{row['Variance']}</b></div>
                            <div><span style="color:#64748b;">Variance %</span><br><b>{row['Variance %']}</b></div>
                        </div>
                        <div style="margin-top:8px; color:#64748b; font-size:0.88rem;">
                            {row['Benchmarks']}
                        </div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )


def build_budget_vs_d4cast(metric_data, role_selection):
    """Latest Forecast vs Locked Budget for Revenue Team focus."""
    if metric_data is None or metric_data.empty:
        return pd.DataFrame()
    role_map = {row["Role"]: row["Report Label"] for _, row in role_selection.iterrows() if pd.notna(row["Report Label"])}
    latest_label = role_map.get("Today / Latest")
    latest = metric_data[metric_data["Report Label"] == latest_label].copy()
    if latest.empty:
        return pd.DataFrame()
    d4 = (latest[latest["Reference"] == "Duetto"]
          .groupby(["Hotel", "Stay Month", "Metric"], as_index=False)["Value"].sum()
          .rename(columns={"Value": "Forecast"}))
    budget = (latest[latest["Reference"] == "Budget"]
              .groupby(["Hotel", "Stay Month", "Metric"], as_index=False)["Value"].sum()
              .rename(columns={"Value": "Budget"}))
    out = d4.merge(budget, on=["Hotel", "Stay Month", "Metric"], how="left")
    out["Variance"] = out["Forecast"] - out["Budget"]
    out["Variance %"] = out["Variance"] / out["Budget"] * 100
    out["Status"] = out["Variance"].apply(budget_status_from_variance)
    if not out.empty:
        out["Metric"] = pd.Categorical(out["Metric"], categories=METRIC_ORDER, ordered=True)
        out = out.sort_values(["Metric", "Variance"]).reset_index(drop=True)
    return out


def render_budget_vs_d4cast(metric_data, role_selection):
    st.markdown('<div class="section-title">Budget vs Forecast Focus</div>', unsafe_allow_html=True)
    st.caption("Purpose: support Revenue Team by showing which hotels are above or below locked budget. Budget variance is the main decision point.")
    budget_df = build_budget_vs_d4cast(metric_data, role_selection)
    if budget_df.empty:
        st.info("No Budget vs Forecast data found. Check whether Budget columns exist in the report.")
        return pd.DataFrame()
    c1,c2,c3 = st.columns([1,1,1])
    metric = c1.selectbox("Metric", ["Rev","Occ","Room","ADR","All Metrics"], index=0, key="budget_focus_metric")
    sort_by = c2.selectbox("Sort by", ["Variance", "Variance %", "Forecast", "Budget"], index=0, key="budget_focus_sort")
    order = c3.selectbox("Order", ["Worst first", "Best first"], index=0, key="budget_focus_order")
    view = budget_df.copy()
    if metric != "All Metrics":
        view = view[view["Metric"] == metric].copy()
    if sort_by in view.columns:
        view = view.sort_values(sort_by, ascending=(order=="Worst first")).reset_index(drop=True)
    total_d4, total_budget = view["Forecast"].sum(), view["Budget"].sum()
    total_var = total_d4-total_budget
    below = int((view["Variance"]<0).sum())
    k1,k2,k3,k4=st.columns(4)
    k1.metric("Total Forecast", fmt_raw2(total_d4))
    k2.metric("Total Budget", fmt_raw2(total_budget))
    k3.metric("Variance", fmt_raw2(total_var), safe_delta(total_d4,total_budget))
    k4.metric("Below Budget Rows", below)
    raw=view.copy(); show=view.copy()
    for col in ["Forecast","Budget","Variance"]:
        if col in show: show[col]=show[col].apply(lambda x: "" if pd.isna(x) else fmt_raw2(x))
    if "Variance %" in show: show["Variance %"]=show["Variance %"].apply(lambda x: "" if pd.isna(x) else fmt_pct2(x))
    def style_budget(_):
        styles=pd.DataFrame("", index=show.index, columns=show.columns)
        for idx,val in raw["Variance"].items():
            color="#bbf7d0" if pd.notna(val) and val>0 else "#fecaca" if pd.notna(val) and val<0 else "#fef08a"
            for col in ["Variance","Variance %","Status"]:
                if col in styles.columns: styles.loc[idx,col]=f"background-color: {color}; font-weight: 700"
        return styles
    st.dataframe(show.style.apply(style_budget, axis=None), use_container_width=True, hide_index=True, height=min(620,48+38*len(show)))
    chart=raw.copy()
    chart["Direction"]=chart["Variance"].apply(lambda x: "Above Budget" if pd.notna(x) and x>0 else "Below Budget" if pd.notna(x) and x<0 else "On Budget")
    fig=px.bar(chart, x="Variance", y="Hotel", orientation="h", color="Direction", color_discrete_map={"Above Budget":"#15803d","Below Budget":"#b91c1c","On Budget":"#ca8a04"}, hover_data={"Forecast":":,.2f","Budget":":,.2f","Variance %":":.2f","Direction":False}, title=f"Budget Variance by Hotel ({metric})")
    fig.update_layout(plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)", height=max(420,46*chart["Hotel"].nunique()), yaxis=dict(categoryorder="total ascending"), margin=dict(l=20,r=20,t=60,b=20))
    fig.update_traces(texttemplate="%{x:,.2f}", textposition="outside", cliponaxis=False)
    st.plotly_chart(fig, use_container_width=True, config={"displaylogo":False})
    return budget_df




def build_weekly_movement(metric_data):
    """Weekly PU = Week End Forecast - Week Start Forecast."""
    if metric_data is None or metric_data.empty: return pd.DataFrame()
    d4=metric_data[metric_data["Reference"]=="Duetto"].copy()
    if d4.empty: return pd.DataFrame()
    d4=add_week_columns(d4).sort_values(["Hotel","Stay Month","Metric","Report Week","Report Date"])
    rows=[]
    for keys,g in d4.groupby(["Hotel","Stay Month","Metric","Report Week"]):
        hotel,stay,metric,week=keys; g=g.sort_values("Report Date")
        sv,ev=g.iloc[0]["Value"],g.iloc[-1]["Value"]
        diff=ev-sv; pct=diff/sv*100 if pd.notna(sv) and sv!=0 else None
        rows.append({"Hotel":hotel,"Stay Month":stay,"Metric":metric,"Report Week":week,"Start Date":g.iloc[0]["Report Date"],"End Date":g.iloc[-1]["Report Date"],"Week Start Forecast":sv,"Week End Forecast":ev,"Weekly PU":diff,"Weekly PU %":pct,"Status":"Up" if diff>0 else "Down" if diff<0 else "Flat"})
    out=pd.DataFrame(rows)
    if out.empty: return out
    out["Metric"]=pd.Categorical(out["Metric"], categories=METRIC_ORDER, ordered=True)
    return out.sort_values(["Report Week","Metric","Weekly PU"]).reset_index(drop=True)


def render_weekly_movement(metric_data):
    st.markdown('<div class="section-title">Weekly Movement View</div>', unsafe_allow_html=True)
    st.caption("Purpose: divide report movement by week. Weekly PU = Week End Forecast - Week Start Forecast.")
    weekly=build_weekly_movement(metric_data)
    if weekly.empty:
        st.info("No weekly movement data. Upload multiple report dates to enable weekly comparison.")
        return pd.DataFrame()
    c1,c2,c3=st.columns([1,1,1])
    metric=c1.selectbox("Metric", ["Rev","Occ","Room","ADR","All Metrics"], index=0, key="weekly_metric")
    week_opts=sorted(weekly["Report Week"].dropna().unique())
    week=c2.selectbox("Report Week", ["All Weeks"]+week_opts, index=0, key="weekly_week")
    order=c3.selectbox("Order", ["Worst first","Best first"], index=0, key="weekly_order")
    view=weekly.copy()
    if metric!="All Metrics": view=view[view["Metric"]==metric].copy()
    if week!="All Weeks": view=view[view["Report Week"]==week].copy()
    view=view.sort_values("Weekly PU", ascending=(order=="Worst first")).reset_index(drop=True)
    k1,k2,k3=st.columns(3)
    k1.metric("Total Weekly PU", fmt_raw2(view["Weekly PU"].sum()))
    k2.metric("Rows Up", int((view["Weekly PU"]>0).sum()))
    k3.metric("Rows Down", int((view["Weekly PU"]<0).sum()))
    raw=view.copy(); show=view.copy()
    for col in ["Week Start Forecast","Week End Forecast","Weekly PU"]:
        show[col]=show[col].apply(lambda x: "" if pd.isna(x) else fmt_raw2(x))
    show["Weekly PU %"]=show["Weekly PU %"].apply(lambda x: "" if pd.isna(x) else fmt_pct2(x))
    def style_week(_):
        styles=pd.DataFrame("", index=show.index, columns=show.columns)
        for idx,val in raw["Weekly PU"].items():
            color="#bbf7d0" if val>0 else "#fecaca" if val<0 else "#fef08a"
            for col in ["Weekly PU","Weekly PU %","Status"]:
                if col in styles.columns: styles.loc[idx,col]=f"background-color: {color}; font-weight: 700"
        return styles
    st.dataframe(show.style.apply(style_week, axis=None), use_container_width=True, hide_index=True, height=min(620,48+36*len(show)))
    chart=raw.copy(); chart["Direction"]=chart["Weekly PU"].apply(lambda x:"Up" if x>0 else "Down" if x<0 else "Flat")
    fig=px.bar(chart, x="Weekly PU", y="Hotel", color="Direction", orientation="h", color_discrete_map={"Up":"#15803d","Down":"#b91c1c","Flat":"#ca8a04"}, hover_data={"Report Week":True,"Stay Month":True,"Week Start Forecast":":,.2f","Week End Forecast":":,.2f","Weekly PU %":":.2f"}, title=f"Weekly Movement by Hotel ({metric})")
    fig.update_layout(plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)", height=max(420,46*chart["Hotel"].nunique()), yaxis=dict(categoryorder="total ascending"), margin=dict(l=20,r=20,t=60,b=20))
    fig.update_traces(texttemplate="%{x:,.2f}", textposition="outside", cliponaxis=False)
    st.plotly_chart(fig, use_container_width=True, config={"displaylogo":False})
    return weekly


def render_metric_dictionary():
    with st.expander("Metric definitions", expanded=False):
        st.markdown("""
        **Latest Forecast** = latest Duetto forecast from the newest report file.  
        **Budget** = locked budget target. This is the main Revenue Team focus.  
        **Variance** = Forecast - Budget/Base. Positive = above target, negative = below target.  
        **Daily PU** = Latest Forecast - previous report Forecast.  
        **7D PU** = Latest Forecast - report around 7 days ago.  
        **MTD PU** = Latest Forecast - first report of the month.  
        **Weekly PU** = Week End Forecast - Week Start Forecast.  
        **Recommended Pace** = best same-time benchmark from STLY / ST2Y / ST3Y.  
        **D4cast vs Final** = current forecast compared with previous-year final actual.
        """)




def build_weekly_movement_v2(metric_data):
    """
    Weekly revenue movement for Revenue Team.

    Weekly Movement = Week End Forecast - Week Start Forecast.
    This is split by Hotel / Stay Month / Metric / Week.
    """
    if metric_data is None or metric_data.empty:
        return pd.DataFrame()

    d4 = metric_data[metric_data["Reference"] == "Duetto"].copy()
    if d4.empty:
        return pd.DataFrame()

    d4["Report Date"] = pd.to_datetime(d4["Report Date"])
    d4["Week Start"] = d4["Report Date"].dt.to_period("W-MON").apply(lambda p: p.start_time.normalize())
    d4["Week End"] = d4["Report Date"].dt.to_period("W-MON").apply(lambda p: p.end_time.normalize())
    d4["Week Label"] = d4.apply(lambda r: friendly_week_label(r["Week Start"], r["Week End"]), axis=1)

    rows = []
    group_cols = ["Hotel", "Stay Month", "Metric", "Week Start", "Week End", "Week Label"]

    for keys, group in d4.sort_values("Report Date").groupby(group_cols):
        hotel, stay_month, metric, week_start, week_end, week_label = keys
        g = group.sort_values("Report Date")
        start_value = g.iloc[0]["Value"]
        end_value = g.iloc[-1]["Value"]
        first_report = g.iloc[0]["Report Date"]
        last_report = g.iloc[-1]["Report Date"]

        weekly_move = end_value - start_value
        weekly_move_pct = weekly_move / start_value * 100 if pd.notna(start_value) and start_value != 0 else None

        rows.append({
            "Hotel": hotel,
            "Stay Month": stay_month,
            "Metric": metric,
            "Week": week_label,
            "First Report": first_report,
            "Last Report": last_report,
            "Start Forecast": start_value,
            "End Forecast": end_value,
            "Weekly Movement": weekly_move,
            "Weekly Movement %": weekly_move_pct,
            "Status": "Up" if weekly_move > 0 else "Down" if weekly_move < 0 else "Flat",
        })

    out = pd.DataFrame(rows)
    if out.empty:
        return out

    out["Metric"] = pd.Categorical(out["Metric"], categories=METRIC_ORDER, ordered=True)
    out = out.sort_values(["Week", "Metric", "Weekly Movement"]).reset_index(drop=True)
    return out


def render_weekly_movement_v2(metric_data):
    st.markdown('<div class="section-title">Weekly Revenue Movement</div>', unsafe_allow_html=True)
    st.caption("Purpose: see which hotel moved up/down by week. Weekly Movement = End-of-week forecast - Start-of-week forecast.")

    weekly = build_weekly_movement_v2(metric_data)
    if weekly.empty:
        st.info("Upload multiple report dates to enable weekly movement.")
        return pd.DataFrame()

    c1, c2, c3 = st.columns([1, 1, 1])
    metric_filter = c1.selectbox(
        "Metric",
        ["Rev", "Occ", "Room", "ADR", "All Metrics"],
        index=0,
        key="weekly_v2_metric",
    )

    week_options = weekly["Week"].dropna().unique().tolist()
    week_filter = c2.selectbox(
        "Week",
        ["All Weeks"] + week_options,
        index=0,
        key="weekly_v2_week",
    )

    focus = c3.selectbox(
        "Focus",
        ["Worst movement", "Best movement", "Highest end forecast"],
        index=0,
        key="weekly_v2_focus",
    )

    view = weekly.copy()
    if metric_filter != "All Metrics":
        view = view[view["Metric"] == metric_filter].copy()
    if week_filter != "All Weeks":
        view = view[view["Week"] == week_filter].copy()

    if view.empty:
        st.info("No weekly data for selected filters.")
        return weekly

    if focus == "Worst movement":
        view = view.sort_values("Weekly Movement", ascending=True).reset_index(drop=True)
    elif focus == "Best movement":
        view = view.sort_values("Weekly Movement", ascending=False).reset_index(drop=True)
    else:
        view = view.sort_values("End Forecast", ascending=False).reset_index(drop=True)

    total_move = view["Weekly Movement"].sum()
    up_rows = (view["Weekly Movement"] > 0).sum()
    down_rows = (view["Weekly Movement"] < 0).sum()

    k1, k2, k3 = st.columns(3)
    k1.metric("Total Weekly Movement", fmt_raw2(total_move))
    k2.metric("Rows Up", int(up_rows))
    k3.metric("Rows Down", int(down_rows))

    # Heatmap-style matrix: Hotel x Week, color = movement.
    st.markdown("#### Weekly movement heatmap")

    heat = view.copy()
    heat["Hotel Short"] = heat["Hotel"].apply(short_hotel_name) if "short_hotel_name" in globals() else heat["Hotel"]
    heat["Display"] = heat["Weekly Movement"].apply(fmt_raw2)

    fig_heat = px.density_heatmap(
        heat,
        x="Week",
        y="Hotel Short",
        z="Weekly Movement",
        histfunc="sum",
        color_continuous_scale=["#b91c1c", "#facc15", "#15803d"],
        title=f"Weekly Movement Heatmap ({metric_filter})",
    )
    fig_heat.update_layout(
        plot_bgcolor="rgba(0,0,0,0)",
        height=max(420, 42 * heat["Hotel Short"].nunique()),
        xaxis_title="Week",
        yaxis_title="Hotel",
        margin=dict(l=20, r=20, t=60, b=20),
        coloraxis_colorbar=dict(title="Movement"),
    )
    st.plotly_chart(fig_heat, use_container_width=True, config={"displaylogo": False}, key="weekly_movement_heatmap_chart")

    st.markdown("#### Weekly movement table")

    show = view[[
        "Hotel", "Stay Month", "Metric", "Week",
        "Start Forecast", "End Forecast", "Weekly Movement", "Weekly Movement %",
        "Status"
    ]].copy()

    raw = view.copy()

    for c in ["Start Forecast", "End Forecast", "Weekly Movement"]:
        show[c] = show[c].apply(lambda x: "" if pd.isna(x) else fmt_raw2(x))
    show["Weekly Movement %"] = show["Weekly Movement %"].apply(lambda x: "" if pd.isna(x) else fmt_pct2(x))

    def style_week(_):
        styles = pd.DataFrame("", index=show.index, columns=show.columns)
        for idx, val in raw["Weekly Movement"].items():
            color = "#bbf7d0" if val > 0 else "#fecaca" if val < 0 else "#fef08a"
            for col in ["Weekly Movement", "Weekly Movement %", "Status"]:
                if col in styles.columns:
                    styles.loc[idx, col] = f"background-color:{color}; font-weight:700"
        return styles

    st.dataframe(
        show.style.apply(style_week, axis=None),
        use_container_width=True,
        hide_index=True,
        height=min(620, 48 + 36 * len(show)),
    )

    return weekly


def build_budget_review(metric_data, role_selection):
    if metric_data is None or metric_data.empty:
        return pd.DataFrame()

    # Per (Hotel, Stay Month, Metric, Reference) keep the row from the latest
    # available Report Date - supports look-back at past stay months.
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
        return pd.DataFrame()

    d4 = latest[latest["Reference"] == "Duetto"].groupby(
        ["Hotel", "Stay Month", "Metric"], as_index=False
    )["Value"].sum().rename(columns={"Value": "Forecast"})

    budget = latest[latest["Reference"] == "Budget"].groupby(
        ["Hotel", "Stay Month", "Metric"], as_index=False
    )["Value"].sum().rename(columns={"Value": "Budget"})

    today = latest[latest["Reference"] == "Today"].groupby(
        ["Hotel", "Stay Month", "Metric"], as_index=False
    )["Value"].sum().rename(columns={"Value": "Today OTB"})

    out = d4.merge(budget, on=["Hotel", "Stay Month", "Metric"], how="left")
    out = out.merge(today, on=["Hotel", "Stay Month", "Metric"], how="left")
    budget_calc = out.apply(lambda r: calc_budget_variance(r["Forecast"], r["Budget"]), axis=1).apply(pd.Series)
    out["Budget Variance"] = budget_calc[0]
    out["Budget Variance %"] = budget_calc[1]
    out["OTB vs Budget %"] = out.apply(
        lambda r: ((r["Today OTB"] - r["Budget"]) / r["Budget"] * 100)
        if pd.notna(r["Today OTB"]) and pd.notna(r["Budget"]) and r["Budget"] != 0
        else None,
        axis=1,
    )
    out["OTB vs Forecast %"] = out.apply(
        lambda r: ((r["Today OTB"] - r["Forecast"]) / r["Forecast"] * 100)
        if pd.notna(r["Today OTB"]) and pd.notna(r["Forecast"]) and r["Forecast"] != 0
        else None,
        axis=1,
    )
    out["Status"] = out["Budget Variance"].apply(budget_status_from_variance)
    return out


def render_budget_review(metric_data, role_selection, key_prefix="budget_review"):
    st.markdown('<div class="section-title">Budget Review</div>', unsafe_allow_html=True)
    st.caption("Revenue purpose: compare latest forecast against Budget. This is the primary decision view for revenue brief.")

    budget_df = build_budget_review(metric_data, role_selection)
    if budget_df.empty:
        st.info("No Budget data found.")
        return pd.DataFrame()

    c1, c2, c3 = st.columns([1, 1, 1])
    metric_choice = c1.selectbox("Metric", ["Rev", "Occ", "Room", "ADR", "All Metrics"], index=0, key=f"{key_prefix}_metric")
    sort_by = c2.selectbox("Sort by", ["Budget Variance", "Budget Variance %", "Forecast", "Budget"], index=0, key=f"{key_prefix}_sort")
    order = c3.selectbox("Order", ["Worst first", "Best first"], index=0, key=f"{key_prefix}_order")

    view = budget_df.copy()
    if metric_choice != "All Metrics":
        view = view[view["Metric"] == metric_choice].copy()

    if sort_by in view.columns:
        view = view.sort_values(sort_by, ascending=(order == "Worst first")).reset_index(drop=True)

    total_forecast = view["Forecast"].sum()
    total_budget = view["Budget"].sum()
    total_var = total_forecast - total_budget

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Forecast", fmt_raw2(total_forecast))
    k2.metric("Budget", fmt_raw2(total_budget))
    k3.metric("Budget Variance", fmt_raw2(total_var), budget_delta_text(calc_budget_variance(total_forecast, total_budget)[1]))
    k4.metric("Below Budget Rows", int((view["Budget Variance"] < 0).sum()))

    raw = view.copy()
    show = view[["Hotel", "Stay Month", "Metric", "Budget", "Forecast", "Budget Variance", "Budget Variance %", "Status"]].copy()
    for c in ["Budget", "Forecast", "Budget Variance"]:
        show[c] = show[c].apply(lambda x: "" if pd.isna(x) else fmt_raw2(x))
    show["Budget Variance %"] = show["Budget Variance %"].apply(lambda x: "" if pd.isna(x) else fmt_pct2(x))

    def style_budget(_):
        styles = pd.DataFrame("", index=show.index, columns=show.columns)
        for idx, val in raw["Budget Variance"].items():
            color = "#bbf7d0" if val > 0 else "#fecaca" if val < 0 else "#fef08a"
            for col in ["Budget Variance", "Budget Variance %", "Status"]:
                styles.loc[idx, col] = f"background-color:{color}; font-weight:700"
        return styles

    st.dataframe(
        show.style.apply(style_budget, axis=None),
        use_container_width=True,
        hide_index=True,
        height=min(560, 48 + 36 * len(show)),
    )

    chart = raw.copy()
    if chart.empty:
        return budget_df

    chart["Direction"] = chart["Budget Variance"].apply(lambda x: "Above Budget" if x > 0 else "Below Budget" if x < 0 else "On Budget")
    color_map = {"Above Budget": "#15803d", "Below Budget": "#b91c1c", "On Budget": "#ca8a04"}

    fig = px.bar(
        chart,
        x="Budget Variance",
        y="Hotel",
        orientation="h",
        color="Direction",
        color_discrete_map=color_map,
        hover_data={"Forecast": ":,.2f", "Budget": ":,.2f", "Budget Variance %": ":.2f"},
        title=f"Budget Variance by Hotel ({metric_choice})",
    )
    fig.update_layout(
        plot_bgcolor="rgba(0,0,0,0)",
        height=max(420, 46 * chart["Hotel"].nunique()),
        yaxis=dict(categoryorder="total ascending"),
        xaxis=dict(showgrid=True, gridcolor="#f1f5f9"),
        margin=dict(l=20, r=20, t=60, b=20),
    )
    fig.update_traces(texttemplate="%{x:,.2f}", textposition="outside", cliponaxis=False)
    st.plotly_chart(
        fig,
        use_container_width=True,
        config={"displaylogo": False},
        key=f"{key_prefix}_budget_variance_chart",
    )

    return budget_df


def render_forecast_trend_by_month(metric_data):
    st.markdown('<div class="section-title">Revenue Forecast Trend</div>', unsafe_allow_html=True)
    st.caption("Trend is split by Stay Month so the line colors clearly show which month is moving.")

    d4 = metric_data[metric_data["Reference"] == "Duetto"].copy()
    if d4.empty:
        st.info("No forecast trend data.")
        return pd.DataFrame()

    c1, c2 = st.columns([1, 1])
    metric_choice = c1.selectbox("Metric", ["Rev", "Occ", "Room", "ADR", "All Metrics"], index=0, key="trend_metric")
    trend_mode = c2.selectbox("Group line by", ["Stay Month", "Hotel", "Stay Month + Hotel"], index=0, key="trend_group")

    view = d4.copy()
    if metric_choice != "All Metrics":
        view = view[view["Metric"] == metric_choice].copy()

    if trend_mode == "Stay Month":
        trend = view.groupby(["Report Date", "Stay Month"], as_index=False)["Value"].sum()
        color_col = "Stay Month"
        line_group = "Stay Month"
    elif trend_mode == "Hotel":
        trend = view.groupby(["Report Date", "Hotel"], as_index=False)["Value"].sum()
        color_col = "Hotel"
        line_group = "Hotel"
    else:
        view["Line"] = view["Stay Month"].astype(str) + " | " + view["Hotel"].astype(str)
        trend = view.groupby(["Report Date", "Line"], as_index=False)["Value"].sum()
        color_col = "Line"
        line_group = "Line"

    trend = trend.sort_values("Report Date")

    fig = px.line(
        trend,
        x="Report Date",
        y="Value",
        color=color_col,
        line_group=line_group,
        markers=True,
        title=f"Forecast Trend by {trend_mode} ({metric_choice})",
    )
    fig.update_layout(
        plot_bgcolor="rgba(0,0,0,0)",
        xaxis=dict(title="Report Date", tickformat="%d %b", showgrid=True, gridcolor="#f1f5f9"),
        yaxis=dict(title="Forecast", showgrid=True, gridcolor="#f1f5f9"),
        legend_title_text=trend_mode,
        margin=dict(l=20, r=20, t=60, b=20),
    )
    st.plotly_chart(fig, use_container_width=True, config={"displaylogo": False}, key="forecast_trend_by_month_chart")

    return trend




def _latest_budget_totals_for_metric(d4_data, metric_data, role_selection, metric_name):
    """
    Header KPI logic for Revenue Team:
    Latest Budget vs Forecast, not vs previous Forecast.
    """
    latest_label = role_selection.loc[
        role_selection["Role"] == "Today / Latest", "Report Label"
    ].iloc[0]

    if pd.isna(latest_label):
        return 0, 0, None, None

    latest = metric_data[
        (metric_data["Report Label"] == latest_label)
        & (metric_data["Metric"] == metric_name)
    ].copy()

    forecast = latest[latest["Reference"] == "Duetto"]["Value"].sum()
    budget = latest[latest["Reference"] == "Budget"]["Value"].sum()

    variance, variance_pct = calc_budget_variance(forecast, budget)

    return forecast, budget, variance, variance_pct



def render_executive_budget_cards(metric_data, role_selection):
    st.markdown("### Revenue Budget Snapshot")
    budget_df = build_budget_review(metric_data, role_selection)
    if budget_df.empty:
        st.info("No Budget data found for executive cards.")
        return

    rev_df = budget_df[budget_df["Metric"] == "Rev"].copy()
    if rev_df.empty:
        rev_df = budget_df.copy()

    forecast = rev_df["Forecast"].sum()
    budget = rev_df["Budget"].sum()
    variance = forecast - budget
    variance_pct = variance / budget * 100 if pd.notna(budget) and budget != 0 else None
    below = int((rev_df["Budget Variance"] < 0).sum())

    cols = st.columns([1.25, 1.25, 1.2, 1])
    cols[0].metric("Revenue Budget", fmt_raw2(budget))
    cols[1].metric("Revenue Forecast", fmt_raw2(forecast))
    cols[2].metric("Variance vs Budget", fmt_raw2(variance), fmt_pct2(variance_pct))
    cols[3].metric("Below Budget Rows", below)

    st.markdown(
        """
        <div style="margin-top:2px; margin-bottom:10px; color:#64748b; font-size:0.92rem;">
        Executive view focuses on <b>Rev</b>. Use detailed mode for Occ, Room, and ADR.
        </div>
        """,
        unsafe_allow_html=True,
    )


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


# ============================================================
# Trend Comparison - two metrics side-by-side with baselines
# ============================================================
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


# ============================================================
# Revenue Momentum % Change chart
# Formula: (revenue_t - revenue_{t-1}) / revenue_{t-1} x 100 per day
# ============================================================
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


def build_forecast_movement_v31(metric_data, role_selection):
    role_map = {
        row["Role"]: row["Report Label"]
        for _, row in role_selection.iterrows()
        if pd.notna(row["Report Label"])
    }
    latest_label = role_map.get("Today / Latest")
    base_roles = {
        "1 Day": role_map.get("Yesterday / Previous"),
        "7 Days": role_map.get("Last 7D"),
        "First Day of Month": role_map.get("1st Month"),
    }

    latest_df = metric_data[
        (metric_data["Report Label"] == latest_label)
        & (metric_data["Reference"] == "Duetto")
    ].copy()

    rows = []
    for keys, group in latest_df.groupby(["Hotel", "Stay Month", "Metric"]):
        hotel, stay_month, metric = keys
        latest_value = group["Value"].sum()

        for period, base_label in base_roles.items():
            if base_label is None:
                base_value = None
                movement = None
                movement_pct = None
                status = "No Base"
            else:
                base_value = metric_data[
                    (metric_data["Hotel"] == hotel)
                    & (metric_data["Stay Month"] == stay_month)
                    & (metric_data["Metric"] == metric)
                    & (metric_data["Report Label"] == base_label)
                    & (metric_data["Reference"] == "Duetto")
                ]["Value"].sum()

                if pd.isna(base_value) or base_value == 0:
                    movement = None
                    movement_pct = None
                    status = "No Base"
                else:
                    movement = latest_value - base_value
                    movement_pct = movement / base_value * 100
                    status = "Up" if movement > 0 else "Down" if movement < 0 else "Flat"

            rows.append({
                "Hotel": hotel,
                "Stay Month": stay_month,
                "Metric": metric,
                "Period": period,
                "Latest Forecast": latest_value,
                "Base Forecast": base_value,
                "Movement": movement,
                "Movement %": movement_pct,
                "Status": status,
                "Risk": risk_level(movement_pct),
            })

    out = pd.DataFrame(rows)
    if out.empty:
        return out

    out["Period"] = pd.Categorical(out["Period"], ["1 Day", "7 Days", "First Day of Month"], ordered=True)
    out["Metric"] = pd.Categorical(out["Metric"], categories=METRIC_ORDER, ordered=True)
    return out.sort_values(["Period", "Metric", "Movement"]).reset_index(drop=True)


def render_forecast_movement_v31(metric_data, role_selection):
    st.markdown('<div class="section-title">Forecast Movement</div>', unsafe_allow_html=True)
    st.caption("Movement tracking only: Latest Forecast compared with 1 Day, 7 Days, and First Day of Month.")

    movement = build_forecast_movement_v31(metric_data, role_selection)
    if movement.empty:
        st.info("No movement data. Upload multiple report dates to compare.")
        return pd.DataFrame()

    c1, c2, c3 = st.columns([1, 1, 1])
    metric_filter = c1.selectbox("Metric", ["Rev", "Occ", "Room", "ADR", "All Metrics"], index=0, key="forecast_movement_metric_v31")
    period_filter = c2.selectbox("Compare with", ["1 Day", "7 Days", "First Day of Month", "All Periods"], index=0, key="forecast_movement_period_v31")
    view_mode = c3.selectbox("View", ["Summary cards", "Color table", "Bar chart"], index=0, key="forecast_movement_view_v31")

    view = movement.copy()
    if metric_filter != "All Metrics":
        view = view[view["Metric"] == metric_filter].copy()
    if period_filter != "All Periods":
        view = view[view["Period"] == period_filter].copy()

    if view.empty:
        st.info("No movement data for selected filters.")
        return movement

    total_movement = view["Movement"].sum()
    up_rows = int((view["Movement"] > 0).sum())
    down_rows = int((view["Movement"] < 0).sum())

    k1, k2, k3 = st.columns(3)
    k1.metric("Total Movement", fmt_raw2(total_movement))
    k2.metric("Rows Up", up_rows)
    k3.metric("Rows Down", down_rows)

    if view_mode == "Summary cards":
        card_df = view.sort_values("Movement", ascending=True).copy()
        show_mode = st.selectbox("Show", ["Worst 5", "Worst 10", "All"], index=0, key="forecast_movement_show_v31")
        if show_mode != "All":
            card_df = card_df.head(int(show_mode.replace("Worst ", "")))

        for _, row in card_df.iterrows():
            move = row["Movement"]
            color = "#bbf7d0" if pd.notna(move) and move > 0 else "#fecaca" if pd.notna(move) and move < 0 else "#fef08a"
            border = "#15803d" if pd.notna(move) and move > 0 else "#b91c1c" if pd.notna(move) and move < 0 else "#ca8a04"
            st.markdown(
                f"""
                <div style="background:{color}; border-left:6px solid {border}; border-radius:14px; padding:14px 16px; margin-bottom:10px;">
                    <div style="font-weight:800; font-size:1.02rem;">{row['Hotel']}  {row['Metric']}  {row['Stay Month']}  {row['Period']}</div>
                    <div style="display:grid; grid-template-columns: repeat(4, minmax(0, 1fr)); gap:12px; margin-top:8px;">
                        <div><span style="color:#64748b;">Latest</span><br><b>{fmt_raw2(row['Latest Forecast'])}</b></div>
                        <div><span style="color:#64748b;">Base</span><br><b>{fmt_raw2(row['Base Forecast'])}</b></div>
                        <div><span style="color:#64748b;">Movement</span><br><b>{fmt_raw2(row['Movement'])}</b></div>
                        <div><span style="color:#64748b;">Movement %</span><br><b>{fmt_pct2(row['Movement %'])}</b></div>
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )

    elif view_mode == "Color table":
        show = view[["Hotel", "Stay Month", "Metric", "Period", "Latest Forecast", "Base Forecast", "Movement", "Movement %", "Status", "Risk"]].copy()
        raw = view.copy()
        for c in ["Latest Forecast", "Base Forecast", "Movement"]:
            show[c] = show[c].apply(lambda x: "" if pd.isna(x) else fmt_raw2(x))
        show["Movement %"] = show["Movement %"].apply(lambda x: "" if pd.isna(x) else fmt_pct2(x))

        def style_move(_):
            styles = pd.DataFrame("", index=show.index, columns=show.columns)
            for idx, val in raw["Movement"].items():
                color = "#bbf7d0" if pd.notna(val) and val > 0 else "#fecaca" if pd.notna(val) and val < 0 else "#fef08a"
                for col in ["Movement", "Movement %", "Status"]:
                    if col in styles.columns:
                        styles.loc[idx, col] = f"background-color:{color}; font-weight:700"
            return styles

        st.dataframe(show.style.apply(style_move, axis=None), use_container_width=True, hide_index=True, height=min(620, 48 + 36 * len(show)))

    else:
        chart = view.copy()
        chart["Direction"] = chart["Movement"].apply(lambda x: "Up" if pd.notna(x) and x > 0 else "Down" if pd.notna(x) and x < 0 else "Flat")
        color_map = {"Up": "#15803d", "Down": "#b91c1c", "Flat": "#ca8a04"}
        fig = px.bar(
            chart,
            x="Movement",
            y="Hotel",
            color="Direction",
            orientation="h",
            color_discrete_map=color_map,
            hover_data={
                "Stay Month": True,
                "Metric": True,
                "Period": True,
                "Latest Forecast": ":,.2f",
                "Base Forecast": ":,.2f",
                "Movement %": ":.2f",
                "Direction": False,
            },
            title=f"Forecast Movement by Hotel ({metric_filter}, {period_filter})",
        )
        fig.update_layout(
            plot_bgcolor="rgba(0,0,0,0)",
            height=max(420, 46 * chart["Hotel"].nunique()),
            yaxis=dict(categoryorder="total ascending"),
            xaxis=dict(showgrid=True, gridcolor="#f1f5f9"),
            margin=dict(l=20, r=20, t=60, b=20),
        )
        fig.update_traces(texttemplate="%{x:,.2f}", textposition="outside", cliponaxis=False)
        st.plotly_chart(fig, use_container_width=True, config={"displaylogo": False}, key="forecast_movement_bar_v31")

    with st.expander("Full movement detail"):
        full = movement.copy()
        for c in ["Latest Forecast", "Base Forecast", "Movement"]:
            full[c] = full[c].apply(lambda x: "" if pd.isna(x) else fmt_raw2(x))
        full["Movement %"] = full["Movement %"].apply(lambda x: "" if pd.isna(x) else fmt_pct2(x))
        st.dataframe(full, use_container_width=True, hide_index=True, height=520)

    return movement



def build_budget_review_summary_view(budget_df, view_level):
    """
    Make Budget Review/Sort Board easier for All Month usage.

    view_level:
    - Summary by Hotel: aggregate selected months into one row per Hotel + Metric
    - Detail by Month: keep Hotel + Stay Month + Metric detail
    """
    if budget_df is None or budget_df.empty:
        return pd.DataFrame()

    df = budget_df.copy()

    if view_level == "Summary by Hotel":
        group_cols = ["Hotel", "Metric"]
    else:
        group_cols = ["Hotel", "Stay Month", "Metric"]

    sum_cols = [c for c in ["Forecast", "Budget", "Today OTB"] if c in df.columns]
    out = df.groupby(group_cols, as_index=False)[sum_cols].sum()

    budget_calc = out.apply(lambda r: calc_budget_variance(r["Forecast"], r["Budget"]), axis=1).apply(pd.Series)
    out["Budget Variance"] = budget_calc[0]
    out["Budget Variance %"] = budget_calc[1]
    out["OTB vs Budget %"] = out.apply(
        lambda r: ((r["Today OTB"] - r["Budget"]) / r["Budget"] * 100)
        if "Today OTB" in out.columns and pd.notna(r["Today OTB"]) and pd.notna(r["Budget"]) and r["Budget"] != 0
        else None,
        axis=1,
    )
    out["OTB vs Forecast %"] = out.apply(
        lambda r: ((r["Today OTB"] - r["Forecast"]) / r["Forecast"] * 100)
        if "Today OTB" in out.columns and pd.notna(r["Today OTB"]) and pd.notna(r["Forecast"]) and r["Forecast"] != 0
        else None,
        axis=1,
    )
    out["Status"] = out["Budget Variance"].apply(budget_status_from_variance)

    if "Metric" in out.columns:
        out["Metric"] = pd.Categorical(out["Metric"], categories=METRIC_ORDER, ordered=True)

    return out


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



def render_revenue_color_legend():
    st.markdown(
        """
        <div class="rev-legend-wrap">
            <span class="rev-pill rev-green">Green = Above Budget / Up / Positive</span>
            <span class="rev-pill rev-red">Red = Below Budget / Down / Risk</span>
            <span class="rev-pill rev-yellow">Yellow = Flat / On Budget / Neutral</span>
            <span class="rev-pill rev-blue">Blue = Forecast / Information</span>
        </div>
        """,
        unsafe_allow_html=True,
    )

def revenue_status_colors(value):
    """
    Return strong UI colors for positive / negative / flat.
    """
    if pd.isna(value):
        return "#dbeafe", "#2563eb", "#1e3a8a", "rev-card-info"
    if value > 0:
        return "#bbf7d0", "#15803d", "#14532d", "rev-card-good"
    if value < 0:
        return "#fecaca", "#b91c1c", "#7f1d1d", "rev-card-bad"
    return "#fef08a", "#ca8a04", "#713f12", "rev-card-flat"

def style_revenue_variance_table(show_df, raw_df, value_col, status_cols=None):
    """
    Stronger coloring for revenue tables.
    """
    if status_cols is None:
        status_cols = []

    def apply_style(_):
        styles = pd.DataFrame("", index=show_df.index, columns=show_df.columns)

        if value_col not in raw_df.columns:
            return styles

        for idx, val in raw_df[value_col].items():
            bg, border, text, _ = revenue_status_colors(val)

            target_cols = [value_col] + status_cols
            pct_col = f"{value_col} %"
            if pct_col in show_df.columns:
                target_cols.append(pct_col)

            # Common alternative percentage column names
            for alt in ["Budget Variance %", "Movement %", "Weekly Movement %", "Daily Movement %", "Variance %"]:
                if alt in show_df.columns and alt not in target_cols:
                    target_cols.append(alt)

            for col in target_cols:
                if col in styles.columns:
                    styles.loc[idx, col] = (
                        f"background-color:{bg}; color:{text}; "
                        f"font-weight:900; border-left:4px solid {border};"
                    )

        return styles

    return show_df.style.apply(apply_style, axis=None)



# ============================================================
# Leaderboard by Stay Month
# ============================================================

def build_leaderboard_by_month(
    metric_long, role_selection,
    stay_month_selection, lb_metric, rank_by
):
    """
    Rank ALL hotels per Stay Month by the chosen metric.
    rank_by = "Budget Variance %" -> rank 1 = most behind budget (needs attention)
    rank_by = "Today OTB" / "Forecast" -> rank 1 = highest value (best performer)
    selected_hotels are highlighted at render time, not filtered here.
    """
    base = metric_long.copy()                         # ALL properties
    base = apply_stay_month_filter(base, stay_month_selection)
    budget_df = build_budget_review(base, role_selection)

    if budget_df.empty:
        return pd.DataFrame()

    if lb_metric != "All Metrics":
        budget_df = budget_df[budget_df["Metric"] == lb_metric].copy()

    if budget_df.empty:
        return pd.DataFrame()

    budget_df = budget_df.copy()
    budget_df["Hotel Short"] = budget_df["Hotel"].apply(short_hotel_name)
    budget_df["_month_sort"] = budget_df["Stay Month"].apply(month_sort_key)

    # Rank within each Stay Month
    # Budget Variance % ascending=True  -> most negative = rank 1 (most behind)
    # OTB / Forecast    ascending=False -> highest = rank 1 (strongest)
    asc_rank = (rank_by == "Budget Variance %")
    rank_col = rank_by if rank_by in budget_df.columns else "Today OTB"
    budget_df["Rank"] = (
        budget_df.groupby("Stay Month")[rank_col]
        .rank(method="dense", ascending=asc_rank)
        .astype(int)
    )

    budget_df = budget_df.sort_values(
        ["_month_sort", "Rank"], ascending=[True, True]
    ).drop(columns=["_month_sort"])

    return budget_df


def render_leaderboard_by_month(
    metric_long, role_selection, selected_hotels, stay_month_selection
):
    """Plotly-table leaderboard, one table per Stay Month.
    Shows ALL properties; selected_hotels are highlighted in blue.
    """
    st.markdown(
        '<div class="section-title">Hotel Leaderboard by Stay Month</div>',
        unsafe_allow_html=True,
    )
    st.caption(
        "All properties ranked per stay month. "
        "Highlighted rows = your selected properties. "
        "Budget Variance % cells: green = above budget  red = below  yellow = flat. "
        "Rank 1 = highest OTB/Forecast, or most behind budget when ranked by Variance %."
    )

    c1, c2, c3 = st.columns([1, 1, 1])
    lb_metric = c1.selectbox(
        "Metric",
        ["Rev", "Occ", "Room", "ADR"],
        index=0,
        key="lb_metric",
    )
    rank_by = c2.selectbox(
        "Rank by",
        ["Today OTB", "Budget Variance %", "Forecast"],
        index=0,
        key="lb_rank_by",
    )
    max_months = c3.selectbox(
        "Show months",
        ["Current + Next 2", "All months"],
        index=0,
        key="lb_max_months",
    )

    lb_data = build_leaderboard_by_month(
        metric_long, role_selection,
        stay_month_selection, lb_metric, rank_by,
    )

    if lb_data.empty:
        st.info("No leaderboard data for selected filters.")
        return pd.DataFrame()

    stay_months = sorted(
        lb_data["Stay Month"].dropna().unique(), key=month_sort_key
    )
    if max_months == "Current + Next 2":
        stay_months = stay_months[:3]

    # Fast lookup set for highlight check
    selected_set = set(selected_hotels)

    # -- Cell color helpers -----------------------------------
    def _var_color(val):
        if pd.isna(val):  return "#dbeafe"
        if val > 0:       return "#bbf7d0"
        if val < 0:       return "#fecaca"
        return "#fef08a"

    def _fmt(val):
        return fmt_raw2(val) if pd.notna(val) else "-"

    def _fmt_pct(val):
        return fmt_signed_pct2(val) if pd.notna(val) else "-"

    # -- One Plotly table per Stay Month ----------------------
    for stay_month in stay_months:
        grp = lb_data[lb_data["Stay Month"] == stay_month].sort_values("Rank")
        if grp.empty:
            continue

        n = len(grp)

        # Determine which rows belong to selected hotels
        is_sel = [h in selected_set for h in grp["Hotel"].tolist()]

        # -- Background colors -----------------------------
        # Selected rows: blue tint across data columns
        # Non-selected: default muted tones
        HL_RANK  = "#dbeafe"   # blue-100
        HL_NAME  = "#bfdbfe"   # blue-200 (slightly stronger for the name)
        HL_DATA  = "#dbeafe"   # blue-100 for value columns

        # Gold / silver / bronze cell tints for rank 1 / 2 / 3 (non-selected rows)
        _RANK_BG = {1: "#fef3c7", 2: "#f1f5f9", 3: "#fde8d0"}
        _RANK_FC = {1: "#92400e", 2: "#475569", 3: "#9a3412"}

        colors_rank = [
            HL_RANK if s else _RANK_BG.get(r, "#f8fafc")
            for r, s in zip(grp["Rank"], is_sel)
        ]
        colors_name = [HL_NAME if s else "#f9fafb"  for s in is_sel]
        colors_otb  = [HL_DATA if s else "#f8fafc"  for s in is_sel]
        colors_bgt  = [HL_DATA if s else "#f8fafc"  for s in is_sel]
        colors_fct  = [HL_DATA if s else "#f8fafc"  for s in is_sel]
        # Variance columns keep semantic color regardless
        colors_var  = [_var_color(v) for v in grp["Budget Variance %"]]

        # -- Font colors ------------------------------------
        fc_rank = [
            "#1d4ed8" if s else _RANK_FC.get(r, "#475569")
            for r, s in zip(grp["Rank"], is_sel)
        ]
        fc_name = ["#1d4ed8" if s else "#1e293b" for s in is_sel]
        fc_data = ["#1d4ed8" if s else "#334155" for s in is_sel]
        fc_var  = ["#374151"] * n

        # -- Hotel name: bold for selected rows -------------
        hotel_names = [
            f"<b>{name}</b>" if s else name
            for name, s in zip(grp["Hotel Short"].tolist(), is_sel)
        ]

        # -- Clean rank numbers -----------------------------
        rank_display = [str(r) for r in grp["Rank"]]

        fig = go.Figure(data=[go.Table(
            columnwidth=[44, 150, 105, 105, 105, 105, 105],
            header=dict(
                values=[
                    "<b>#</b>",
                    "<b>Hotel</b>",
                    "<b>Today OTB</b>",
                    "<b>Budget</b>",
                    "<b>Forecast</b>",
                    "<b>vs Budget %</b>",
                    "<b>vs Budget</b>",
                ],
                fill_color="#1e293b",
                font=dict(color="white", size=12),
                align=["center", "left", "right", "right", "right", "center", "right"],
                height=34,
            ),
            cells=dict(
                values=[
                    rank_display,
                    hotel_names,
                    [_fmt(v) for v in grp["Today OTB"]],
                    [_fmt(v) for v in grp["Budget"]],
                    [_fmt(v) for v in grp["Forecast"]],
                    [_fmt_pct(v) for v in grp["Budget Variance %"]],
                    [_fmt(v) for v in grp["Budget Variance"]],
                ],
                fill_color=[
                    colors_rank,
                    colors_name,
                    colors_otb,
                    colors_bgt,
                    colors_fct,
                    colors_var,
                    colors_var,
                ],
                font=dict(
                    size=12,
                    color=[fc_rank, fc_name, fc_data, fc_data, fc_data, fc_var, fc_var],
                ),
                align=["center", "left", "right", "right", "right", "center", "right"],
                height=32,
            ),
        )])

        fig.update_layout(
            margin=dict(l=0, r=0, t=0, b=4),
            height=max(100, 38 + 32 * n),
            paper_bgcolor="rgba(0,0,0,0)",
        )

        safe_key = re.sub(r"[^A-Za-z0-9]", "_", f"{stay_month}_{lb_metric}_{rank_by}")
        st.markdown(f"#### {stay_month}")
        st.plotly_chart(
            fig,
            use_container_width=True,
            config={"displaylogo": False},
            key=f"lb_{safe_key}",
        )

    return lb_data


def render_mega_leaderboard(metric_long, role_selection, hotels, stay_month_selection, report_file_month):
    """
    Leaderboard table by stay month.
    Shows all hotels; selected hotels are highlighted in blue.
    Rows = hotels, columns = key metrics. Room Nights excluded.
    """
    role_map = {
        row["Role"]: row["Report Label"]
        for _, row in role_selection.iterrows()
        if pd.notna(row["Report Label"])
    }
    today_label = role_map.get("Today / Latest")
    first_label = role_map.get("1st Month")

    if not today_label:
        st.warning("Today's report not available.")
        return

    # Show ALL hotels in metric_long, highlight only the selected ones at render time.
    selected_set = set(hotels) if hotels else set()
    base = metric_long.copy()
    available_months = sorted(base["Stay Month"].dropna().unique(), key=month_sort_key)
    start_idx = available_months.index(report_file_month) if report_file_month in available_months else 0
    leaderboard_months = available_months[start_idx : min(start_idx + 6, len(available_months))]
    base = base[base["Stay Month"].isin(leaderboard_months)].copy()
    if base.empty:
        st.info("No data for the current filter.")
        return

    stay_months = sorted(base["Stay Month"].dropna().unique(), key=month_sort_key)
    if not stay_months:
        st.info("No stay months in data.")
        return

    # -- Month tags (CM / M+1 / M-1 ) -------------------------
    cm_dt = pd.to_datetime(report_file_month, format="%b, %Y", errors="coerce")
    month_tags = []
    for sm in stay_months:
        sm_dt = pd.to_datetime(sm, format="%b, %Y", errors="coerce")
        if pd.isna(cm_dt) or pd.isna(sm_dt):
            month_tags.append((str(sm), sm))
            continue
        delta = (sm_dt.year - cm_dt.year) * 12 + (sm_dt.month - cm_dt.month)
        tag = "CM" if delta == 0 else (f"M+{delta}" if delta > 0 else f"M{delta}")
        month_tags.append((tag, sm))

    # Pre-compute LATEST snapshot per (Hotel, Stay Month, Metric, Reference) -
    # so look-back stay months pull from their own latest available report file.
    if "Report Date" in base.columns:
        base_latest = (
            base.sort_values("Report Date")
            .drop_duplicates(
                subset=["Hotel", "Stay Month", "Metric", "Reference"],
                keep="last",
            )
        )
    else:
        base_latest = base

    # -- Data helpers ------------------------------------------
    def get_value(metric, sm, report_label, reference="Today"):
        # If report_label is the today_label sentinel, use the latest-snapshot
        # view (works for past stay months). For 1st-of-month, filter exactly.
        if report_label == today_label or report_label is None:
            sub = base_latest[
                (base_latest["Stay Month"] == sm)
                & (base_latest["Metric"] == metric)
                & (base_latest["Reference"] == reference)
            ]
        else:
            sub = base[
                (base["Report Label"] == report_label)
                & (base["Stay Month"] == sm)
                & (base["Metric"] == metric)
                & (base["Reference"] == reference)
            ]
        return sub.groupby("Hotel")["Value"].sum() if not sub.empty else pd.Series(dtype=float)

    def get_var_pct(metric, sm, primary_ref="Today"):
        """Variance % vs Budget for `primary_ref` (Today or Duetto)."""
        primary = get_value(metric, sm, today_label, primary_ref)
        bgt = get_value(metric, sm, today_label, "Budget")
        if primary.empty or bgt.empty:
            return pd.Series(dtype=float)
        common = primary.index.intersection(bgt.index)
        if not len(common):
            return pd.Series(dtype=float)
        denom = bgt.loc[common].abs().replace(0, pd.NA)
        return ((primary.loc[common] - bgt.loc[common]) / denom * 100).dropna()

    def get_var_raw_k(metric, sm, primary_ref="Today"):
        """Variance vs Budget in thousands for `primary_ref` (Today or Duetto)."""
        primary = get_value(metric, sm, today_label, primary_ref)
        bgt = get_value(metric, sm, today_label, "Budget")
        if primary.empty or bgt.empty:
            return pd.Series(dtype=float)
        common = primary.index.intersection(bgt.index)
        if not len(common):
            return pd.Series(dtype=float)
        return (primary.loc[common] - bgt.loc[common]) / 1000

    # (header, fetch_fn, is_signed, formatter, sort_descending)
    column_specs = [
        ("ADR 1st",                lambda sm: get_value("ADR", sm, first_label),                False, "num",        True),
        ("ADR Today",              lambda sm: get_value("ADR", sm, today_label, "Today"),       False, "num",        True),
        ("ADR Duetto",             lambda sm: get_value("ADR", sm, today_label, "Duetto"),      False, "num",        True),
        ("Occ 1st",                lambda sm: get_value("Occ", sm, first_label),                False, "pct",        True),
        ("Occ Today",              lambda sm: get_value("Occ", sm, today_label, "Today"),       False, "pct",        True),
        ("Occ Duetto",             lambda sm: get_value("Occ", sm, today_label, "Duetto"),      False, "pct",        True),
        ("ADR Today VS BUD %",     lambda sm: get_var_pct("ADR", sm, "Today"),                  True,  "pct_signed", True),
        ("ADR Duetto VS BUD %",    lambda sm: get_var_pct("ADR", sm, "Duetto"),                 True,  "pct_signed", True),
        ("Rev Today VS BUD (K)",   lambda sm: get_var_raw_k("Rev", sm, "Today"),                True,  "num_signed", True),
        ("Rev Duetto VS BUD (K)",  lambda sm: get_var_raw_k("Rev", sm, "Duetto"),               True,  "num_signed", True),
    ]
    rank_options = ["Rev Today VS BUD (K)", "Rev Duetto VS BUD (K)", "Occ Today"]
    st.markdown('<div class="rank-button-group">', unsafe_allow_html=True)
    rank_by = st.pills(
        "Rank by",
        options=rank_options,
        selection_mode="single",
        default="Rev Today VS BUD (K)",
        key="mega_leaderboard_rank_by",
        label_visibility="collapsed",
    )
    st.markdown('</div>', unsafe_allow_html=True)
    if not rank_by:
        rank_by = "Rev Today VS BUD (K)"

    def fmt_val(v, kind):
        if pd.isna(v):
            return ""
        if kind == "num":        return f"{v:,.0f}"
        if kind == "pct":        return f"{v:.1f}%"
        if kind == "pct_signed": return f"{v:+.1f}%"
        if kind == "num_signed": return f"{v:+,.0f}"
        return str(v)

    def month_frame(sm):
        hotels_for_month = sorted(
            base[base["Stay Month"] == sm]["Hotel"].dropna().unique(),
            key=lambda h: short_hotel_name(h),
        )
        rows = []
        for hotel in hotels_for_month:
            row = {
                "Selected": hotel in selected_set,
                "Hotel": short_hotel_name(hotel),
            }
            for header, fn, _is_signed, _fmt_type, _sort_desc in column_specs:
                series = fn(sm)
                row[header] = series.get(hotel, pd.NA) if not series.empty else pd.NA
            rows.append(row)
        out = pd.DataFrame(rows)
        if out.empty:
            return out
        sort_col = rank_by if rank_by in out.columns else "Rev Today VS BUD (K)"
        out = out.sort_values(sort_col, ascending=False, na_position="last").reset_index(drop=True)
        out.insert(1, "Rank", range(1, len(out) + 1))
        return out

    # Top-3 tints - gold / silver / bronze backgrounds (no emoji, kept professional)
    _MEDAL_BG = {1: "#fef3c7", 2: "#e2e8f0", 3: "#fed7aa"}
    _MEDAL_FG = {1: "#78350f", 2: "#1f2937", 3: "#7c2d12"}

    def style_leaderboard_table(df):
        data = df.copy()
        selected = data["Selected"].fillna(False).tolist() if "Selected" in data.columns else [False] * len(data)
        ranks = data["Rank"].tolist() if "Rank" in data.columns else [None] * len(data)
        show = data.drop(columns=["Selected"], errors="ignore")

        def apply_style(_):
            styles = pd.DataFrame("", index=show.index, columns=show.columns)
            for idx in show.index:
                # -- Subtle zebra striping for unselected rows ---
                if not selected[idx] and idx % 2 == 1:
                    for col in show.columns:
                        styles.loc[idx, col] = "background-color:#fafbfc"

                # -- Rank medal coloring (top 3) -----------------
                rk = ranks[idx] if idx < len(ranks) else None
                if "Rank" in show.columns and pd.notna(rk) and int(rk) in _MEDAL_BG:
                    r = int(rk)
                    styles.loc[idx, "Rank"] = (
                        f"background-color:{_MEDAL_BG[r]}; "
                        f"color:{_MEDAL_FG[r]}; "
                        f"font-weight:900; text-align:center;"
                    )

                # -- Selected hotel highlight --------------------
                if selected[idx]:
                    for col in show.columns:
                        styles.loc[idx, col] = "background-color:#eff6ff; color:#1e3a8a; font-weight:700"
                    if "Hotel" in show.columns:
                        styles.loc[idx, "Hotel"] = (
                            "background-color:#dbeafe; color:#1d4ed8; "
                            "font-weight:900; border-left:4px solid #2563eb"
                        )
                    # Preserve medal bg even when selected
                    if "Rank" in show.columns and pd.notna(rk) and int(rk) in _MEDAL_BG:
                        r = int(rk)
                        styles.loc[idx, "Rank"] = (
                            f"background-color:{_MEDAL_BG[r]}; "
                            f"color:{_MEDAL_FG[r]}; "
                            f"font-weight:900; text-align:center; "
                            f"box-shadow: inset 0 0 0 2px #2563eb;"
                        )

                # -- Variance cell coloring ----------------------
                for col in [
                    "ADR Today VS BUD %",  "ADR Duetto VS BUD %",
                    "Rev Today VS BUD (K)", "Rev Duetto VS BUD (K)",
                ]:
                    if col not in show.columns:
                        continue
                    val = data.loc[idx, col]
                    if pd.isna(val):
                        continue
                    if val > 0:
                        styles.loc[idx, col] = (
                            "background-color:#bbf7d0; color:#14532d; "
                            "font-weight:900; text-align:right;"
                        )
                    elif val < 0:
                        styles.loc[idx, col] = (
                            "background-color:#fecaca; color:#7f1d1d; "
                            "font-weight:900; text-align:right;"
                        )
                    else:
                        styles.loc[idx, col] = (
                            "background-color:#fef08a; color:#713f12; "
                            "font-weight:800; text-align:right;"
                        )
            return styles

        def _fmt_rank(v):
            if pd.isna(v):
                return ""
            return f"{int(v)}"

        fmt = {
            "Rank": _fmt_rank,
            "ADR 1st":                lambda v: "" if pd.isna(v) else f"{v:,.0f}",
            "ADR Today":              lambda v: "" if pd.isna(v) else f"{v:,.0f}",
            "ADR Duetto":             lambda v: "" if pd.isna(v) else f"{v:,.0f}",
            "Occ 1st":                lambda v: "" if pd.isna(v) else f"{v:.1f}%",
            "Occ Today":              lambda v: "" if pd.isna(v) else f"{v:.1f}%",
            "Occ Duetto":             lambda v: "" if pd.isna(v) else f"{v:.1f}%",
            "ADR Today VS BUD %":     lambda v: "" if pd.isna(v) else f"{v:+.1f}%",
            "ADR Duetto VS BUD %":    lambda v: "" if pd.isna(v) else f"{v:+.1f}%",
            "Rev Today VS BUD (K)":   lambda v: "" if pd.isna(v) else f"{v:+,.0f}",
            "Rev Duetto VS BUD (K)":  lambda v: "" if pd.isna(v) else f"{v:+,.0f}",
        }

        styler = show.style.format(fmt, na_rep="").apply(apply_style, axis=None)
        # Header polish
        styler = styler.set_table_styles([
            {"selector": "thead th", "props": [
                ("background", "linear-gradient(180deg,#fafbfc 0%,#f1f5f9 100%)"),
                ("color", "#475569"),
                ("font-weight", "700"),
                ("font-size", "11px"),
                ("text-transform", "uppercase"),
                ("letter-spacing", "0.06em"),
                ("border-bottom", "1px solid #e2e8f0"),
                ("padding", "10px 8px"),
                ("text-align", "center"),
            ]},
            {"selector": "tbody td", "props": [
                ("font-size", "12px"),
                ("padding", "6px 10px"),
                ("border-bottom", "1px solid #f1f5f9"),
            ]},
            {"selector": "", "props": [
                ("border-collapse", "separate"),
                ("border-spacing", "0"),
                ("border-radius", "10px"),
                ("overflow", "hidden"),
                ("box-shadow", "0 1px 3px rgba(15,23,42,0.05), 0 0 0 1px rgba(15,23,42,0.06)"),
            ]},
        ])
        return styler

    def _month_header_html(tag, sm, n_hotels):
        """Pretty pill-style header for each stay-month section."""
        return (
            '<div style="display:flex;align-items:center;gap:12px;margin:20px 0 8px 0;">'
            f'<span style="display:inline-flex;align-items:center;justify-content:center;'
            f'min-width:46px;height:30px;padding:0 12px;border-radius:8px;'
            f'background:linear-gradient(180deg,#dbeafe 0%,#bfdbfe 100%);'
            f'color:#1e3a8a;font-weight:800;font-size:13px;letter-spacing:0.04em;'
            f'box-shadow:0 1px 2px rgba(30,58,138,0.15);">{html.escape(tag)}</span>'
            f'<span style="font-weight:700;font-size:14px;color:#0f172a;">{html.escape(str(sm))}</span>'
            f'<span style="font-size:11px;color:#94a3b8;font-weight:500;">'
            f' {n_hotels} hotels ranked</span>'
            '</div>'
        )

    st.caption(
        "One table per stay month. Selected properties highlighted in blue. "
        f"Top 3 ranks (gold / silver / bronze cells) sorted by {rank_by}  "
        "Green / red columns show budget variance."
    )
    for tag, sm in month_tags:
        month_df = month_frame(sm)
        if month_df.empty:
            st.markdown(_month_header_html(tag, sm, 0), unsafe_allow_html=True)
            st.info("No leaderboard data for this stay month.")
            continue
        st.markdown(_month_header_html(tag, sm, len(month_df)), unsafe_allow_html=True)
        st.dataframe(
            style_leaderboard_table(month_df),
            use_container_width=True,
            hide_index=True,
            height=min(520, 48 + 36 * len(month_df)),
        )


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

    # Use whichever movement builder exists.
    if "build_forecast_movement_v31" in globals():
        movement = build_forecast_movement_v31(metric_data, role_selection)
    elif "build_duetto_movement_summary" in globals():
        movement = build_duetto_movement_summary(metric_data, role_selection)
    else:
        movement = build_movement_summary(metric_data, role_selection) if "build_movement_summary" in globals() else pd.DataFrame()

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



def render_budget_first_kpi_cards(metric_data, role_selection, selected_metric):
    """
    Revenue KPI cards using Budget as the base/target.

    Correct revenue order:
    1. Budget
    2. Forecast
    3. Variance vs Budget = Forecast - Budget
    4. Variance vs Budget % = Variance / Budget * 100
    """
    st.caption("Budget is the target/base. Variance % = (Latest Forecast - Budget) / Budget x 100.")

    metrics_to_show = metric_label_order() if selected_metric == "All Metrics" else [selected_metric]

    for m in metrics_to_show:
        forecast, budget, variance, variance_pct = _latest_budget_totals_for_metric(
            d4_data=None,
            metric_data=metric_data,
            role_selection=role_selection,
            metric_name=m,
        )

        st.markdown(f"#### {m}")

        cols = st.columns([1.25, 1.25, 1.1, 1.1])

        cols[0].metric("Budget", fmt_raw2(budget))
        cols[1].metric("Latest Forecast", fmt_raw2(forecast))
        cols[2].metric("Variance vs Budget", fmt_raw2(variance), fmt_pct2(variance_pct))
        cols[3].metric("Variance vs Budget %", fmt_pct2(variance_pct))


def render_budget_first_executive_cards(metric_data, role_selection):
    """
    Executive KPI cards focused on Rev, with Budget first.
    """
    budget_df = build_budget_review(metric_data, role_selection)
    if budget_df.empty:
        st.info("No Budget data found.")
        return

    rev_df = budget_df[budget_df["Metric"] == "Rev"].copy()
    if rev_df.empty:
        rev_df = budget_df.copy()

    budget = rev_df["Budget"].sum()
    forecast = rev_df["Forecast"].sum()
    variance = forecast - budget
    variance_pct = variance / budget * 100 if pd.notna(budget) and budget != 0 else None

    cols = st.columns([1.25, 1.25, 1.1, 1.1])
    cols[0].metric("Revenue Budget", fmt_raw2(budget))
    cols[1].metric("Revenue Forecast", fmt_raw2(forecast))
    cols[2].metric("Variance vs Budget", fmt_raw2(variance), fmt_pct2(variance_pct))
    cols[3].metric("Variance vs Budget %", fmt_pct2(variance_pct))









def _latest_budget_totals_for_metric_v39(metric_data, role_selection, metric_name):
    latest_label = role_selection.loc[
        role_selection["Role"] == "Today / Latest", "Report Label"
    ].iloc[0]

    if pd.isna(latest_label):
        return 0, 0, 0, None

    latest = metric_data[
        (metric_data["Report Label"] == latest_label)
        & (metric_data["Metric"] == metric_name)
    ].copy()

    forecast = latest[latest["Reference"] == "Duetto"]["Value"].sum()
    budget = latest[latest["Reference"] == "Budget"]["Value"].sum()

    variance, variance_pct = calc_budget_variance(forecast, budget)

    return forecast, budget, variance, variance_pct


KPI_AXIS_MAP = {
    "Budget": "Budget",
    "Forecast": "Forecast",
    "On The Book": "Today OTB",
}
# Order determines base axis: earlier = base.
# Forecast before OTB -> "On The Book vs Forecast" (OTB is compare, Forecast is base)
KPI_AXIS_ORDER = list(KPI_AXIS_MAP)


def sum_kpi_axis_values(budget_df, selected_axes):
    return {
        axis: budget_df[KPI_AXIS_MAP[axis]].sum(min_count=1)
        if KPI_AXIS_MAP[axis] in budget_df.columns
        else None
        for axis in selected_axes
    }


def kpi_pair_variance(axis_a, axis_b, axis_values):
    if "Budget" in [axis_a, axis_b]:
        base_axis = "Budget"
        compare_axis = axis_b if axis_a == "Budget" else axis_a
    else:
        axis_a_index = KPI_AXIS_ORDER.index(axis_a)
        axis_b_index = KPI_AXIS_ORDER.index(axis_b)
        base_axis, compare_axis = (axis_a, axis_b) if axis_a_index < axis_b_index else (axis_b, axis_a)

    base_value = axis_values.get(base_axis)
    compare_value = axis_values.get(compare_axis)
    variance = None if pd.isna(compare_value) or pd.isna(base_value) else compare_value - base_value
    variance_pct = (
        variance / base_value * 100
        if variance is not None and pd.notna(base_value) and base_value != 0
        else None
    )
    return compare_axis, base_axis, variance, variance_pct


def render_kpi_variance_card(container, label, variance, variance_pct, show_abs=False):
    bg, border, text, _ = revenue_status_colors(variance)
    value = fmt_raw2(variance) if show_abs else fmt_signed_pct2(variance_pct)
    detail = fmt_signed_pct2(variance_pct) if show_abs else "Variance %"
    container.markdown(
        f"""
        <div style="
            background:{bg};
            color:{text};
            border-left:6px solid {border};
            border-radius:8px;
            min-height:120px;
            padding:14px 16px;
            margin-bottom:10px;
        ">
            <div style="font-size:0.88rem; font-weight:700;">{label}</div>
            <div style="font-size:clamp(1.05rem, 1.45vw, 1.75rem); font-weight:800; margin-top:8px;">{value}</div>
            <div style="font-size:0.86rem; font-weight:700; margin-top:6px;">{detail}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_kpi_axis_cards(budget_df, selected_axes):
    axis_values = sum_kpi_axis_values(budget_df, selected_axes)
    pairs = [
        (selected_axes[left], selected_axes[right])
        for left in range(len(selected_axes))
        for right in range(left + 1, len(selected_axes))
    ]

    cards = [("value", axis) for axis in selected_axes]
    if len(selected_axes) == 2:
        cards.extend([("variance_abs", pairs[0]), ("variance_pct", pairs[0])])
    else:
        cards.extend([("variance_pct", pair) for pair in pairs])

    cols = st.columns(len(cards))
    for col, card in zip(cols, cards):
        card_type, data = card
        if card_type == "value":
            col.metric(data, fmt_raw2(axis_values.get(data)))
            continue

        compare_axis, base_axis, variance, variance_pct = kpi_pair_variance(data[0], data[1], axis_values)
        label = f"{compare_axis} vs {base_axis}" if card_type == "variance_abs" else f"{compare_axis} vs {base_axis} %"
        render_kpi_variance_card(
            col,
            label,
            variance,
            variance_pct,
            show_abs=(card_type == "variance_abs"),
        )



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



def style_pace_variance_table(df):
    """
    Color Same-Time Pace Benchmark table by Variance %.
    Green = ahead, Red = behind, Yellow = on pace.
    Risk column is intentionally not shown because Status + Variance % already explain the situation.
    """
    if df is None or df.empty:
        return df

    show = df.copy()
    raw = df.copy()

    # Drop Risk from UI if still present from older code paths.
    if "Risk" in show.columns:
        show = show.drop(columns=["Risk"])
    if "Risk" in raw.columns:
        raw = raw.drop(columns=["Risk"])

    def parse_pct(v):
        if pd.isna(v):
            return None
        if isinstance(v, str):
            cleaned = v.replace("%", "").replace(",", "").strip()
            if cleaned == "":
                return None
            try:
                return float(cleaned)
            except Exception:
                return None
        try:
            return float(v)
        except Exception:
            return None

    def apply_style(_):
        styles = pd.DataFrame("", index=show.index, columns=show.columns)

        if "Variance %" not in raw.columns:
            return styles

        for idx, val in raw["Variance %"].items():
            num = parse_pct(val)

            if num is None:
                bg, text, border = "#dbeafe", "#1e3a8a", "#2563eb"
            elif num > 0:
                bg, text, border = "#bbf7d0", "#14532d", "#15803d"
            elif num < 0:
                bg, text, border = "#fecaca", "#7f1d1d", "#b91c1c"
            else:
                bg, text, border = "#fef08a", "#713f12", "#ca8a04"

            for col in ["Variance %", "Variance", "Status"]:
                if col in styles.columns:
                    styles.loc[idx, col] = (
                        f"background-color:{bg}; color:{text}; "
                        f"font-weight:900; border-left:4px solid {border};"
                    )

        return styles

    return show.style.apply(apply_style, axis=None)


# ============================================================
# Main UI Execution
# ============================================================

# -- Page title slot (filled after sidebar so we know the report date) -
# Reserve top-of-page space here, fill it once `report_file_month` is set.
_title_slot = st.empty()

# -- Sidebar ---------------------------------------------------
# RULE: never call st.stop() in main area before with st.sidebar: completes.
# Reason: stop in main area = sidebar never renders = widgets disappear.
# Correct pattern: let sidebar finish first, then guard-check in main area.
# -------------------------------------------------------------

def main() -> None:
    # Default values - overwritten by sidebar once data is loaded
    file_catalog = None
    mode = "Upload"

    with st.sidebar:
        # -- Brand header -----------------------------------------
        if LOGO_B64:
            st.markdown(
                f"""
                <div style="
                    display: flex;
                    align-items: center;
                    gap: 10px;
                    padding: 14px 4px 12px 4px;
                    border-bottom: 1px solid #e4e4e4;
                    margin-bottom: 10px;
                ">
                    <img
                        src="data:image/png;base64,{LOGO_B64}"
                        style="
                            width: 34px;
                            height: 34px;
                            border-radius: 7px;
                            object-fit: cover;
                            flex-shrink: 0;
                        "
                    />
                    <div style="line-height: 1.35;">
                        <div style="font-size:0.95rem;font-weight:700;
                                    color:#111;letter-spacing:0.03em;">
                            ATMIND GROUP
                        </div>
                        <div style="font-size:0.78rem;color:#888;">
                            Revenue Dashboard
                        </div>
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )

        st.markdown("## Data source")
        if st.session_state.pop("force_upload_mode", False):
            st.session_state["data_source_mode"] = "Upload"

        mode = st.radio(
            "source_mode",
            ["Folder", "Upload"],
            horizontal=True,
            label_visibility="collapsed",
            key="data_source_mode",
        )

        # -- Load data ------------------------------------------
        if mode == "Folder":
            folder_path = st.text_input(
                "Path",
                value=r"G:\My Drive\Ecom\Report\G5 - Weekly Pace Review",
            )
            if st.button("->  Refresh Data", use_container_width=True, type="primary"):
                st.cache_data.clear()
                st.rerun()
            try:
                with st.spinner("Loading"):
                    file_catalog = build_file_catalog_from_folder(folder_path)
                st.caption(f"{len(file_catalog)} files found")
            except Exception as e:
                st.error(str(e))
                # file_catalog stays None - main area guard will stop rendering
        else:
            uploaded = st.file_uploader(
                "Files",
                type=["zip", "csv", "xlsx", "xls"],
                accept_multiple_files=True,
                label_visibility="visible",
                help="Drop daily G5 CSV/Excel files, or one ZIP.",
            )
            if uploaded:
                file_catalog = build_file_catalog_from_uploads(uploaded)
                st.session_state["uploaded_file_catalog"] = file_catalog
            elif st.session_state.get("uploaded_file_catalog") is not None:
                file_catalog = st.session_state["uploaded_file_catalog"]
            else:
                st.caption("Drop files above to begin.")
                # file_catalog stays None - main area guard will stop rendering

        # -- Data processing (filters live on main page) ------------
        if file_catalog is not None:
            latest_report_month = file_catalog["Report Date"].max().strftime("%b, %Y")
            report_file_month = latest_report_month  # Always auto-use latest report

            # Role identification stays anchored to the CURRENT report month
            role_selection, month_file_catalog = select_role_files(file_catalog, report_file_month)

            # -- Look-back: also load the previous month's report files so
            # the Stay Month dropdown can include past months. Role mapping
            # (Today / Yesterday / 7D / 1st Month) still comes from the
            # current month only.
            try:
                _cur_dt = pd.to_datetime(report_file_month, format="%b, %Y")
                _prev_start = (_cur_dt - pd.DateOffset(months=1))
                _prev_end   = _prev_start + pd.offsets.MonthEnd(0)
                _prev_files = file_catalog[
                    (file_catalog["Report Date"] >= _prev_start)
                    & (file_catalog["Report Date"] <= _prev_end)
                ].copy()
            except Exception:
                _prev_files = pd.DataFrame(columns=file_catalog.columns)

            # Merge prev + current month catalogs. The two date ranges are disjoint
            # by construction so concat without dedup is safe - and this avoids the
            # None-File-Path collision that would dedupe uploaded files.
            if _prev_files.empty:
                extended_catalog = month_file_catalog.copy()
            else:
                extended_catalog = pd.concat(
                    [_prev_files, month_file_catalog], ignore_index=True
                )
                # Safe dedup using (File Name + Report Date) - works for both
                # folder mode and upload mode (where File Path may be None).
                if {"File Name", "Report Date"}.issubset(extended_catalog.columns):
                    extended_catalog = extended_catalog.drop_duplicates(
                        subset=["File Name", "Report Date"]
                    )
            selected_file_catalog = extended_catalog.sort_values("Report Date").reset_index(drop=True)
            selected_file_catalog["Report Order"] = range(1, len(selected_file_catalog) + 1)

            with st.spinner("Processing"):
                combined_df = pd.concat(
                    [parse_record(row) for _, row in selected_file_catalog.iterrows()],
                    ignore_index=True,
                )
                ref_col_map = build_ref_col_map(combined_df)
                if not ref_col_map.get("Duetto"):
                    st.error("No Forecast / Duetto columns found.")
                    file_catalog = None  # reset -> main area guard will stop
                else:
                    metric_long = build_metric_long(combined_df, ref_col_map)

        if file_catalog is not None:
            all_hotels = sorted(metric_long["Hotel"].dropna().unique())
            all_stay_months = sorted(metric_long["Stay Month"].dropna().unique(), key=month_sort_key)
            # Default hotels = Altera properties (or all if none found)
            _default_hotels = [h for h in all_hotels if "altera" in str(h).lower()] or list(all_hotels)
            # Default stay window = (report month - 1) -> report month + 2
            # so when a new report month is uploaded the team can still look back
            # at the previous month's pace.
            _rpt_idx = all_stay_months.index(report_file_month) if report_file_month in all_stay_months else 0
            _default_start = all_stay_months[max(_rpt_idx - 1, 0)]
            _default_end   = all_stay_months[min(_rpt_idx + 2, len(all_stay_months) - 1)]
            # Metric: always All - each tab handles its own filtering
            selected_metric = "All Metrics"
            st.caption(f"Report: {report_file_month}    {len(file_catalog)} files")

    # -- Fill the title slot now that sidebar has set report_file_month --
    with _title_slot.container():
        _ttl_left, _ttl_right = st.columns([3, 2])
        with _ttl_left:
            st.markdown("## Revenue Briefing")
        with _ttl_right:
            if file_catalog is not None:
                _latest_dt = file_catalog["Report Date"].max()
                _n_files = len(file_catalog)
                _latest_label = (
                    _latest_dt.strftime("%a, %d %b %Y")
                    if pd.notna(_latest_dt) else "-"
                )
                _is_today_data = (
                    pd.notna(_latest_dt)
                    and _latest_dt.normalize() == pd.Timestamp.today().normalize()
                )
                _freshness_bg = "#dcfce7" if _is_today_data else "#eff6ff"
                _freshness_fg = "#15803d" if _is_today_data else "#1e40af"
                _freshness_dot = "#16a34a" if _is_today_data else "#3b82f6"
                _freshness_text = "Today's data" if _is_today_data else "Latest available"
                st.markdown(
                    f'''
                    <div style="display:flex;justify-content:flex-end;
                                align-items:flex-start;padding-top:6px;">
                      <div style="background:#ffffff;border:1px solid #e5e7eb;
                                  border-radius:10px;padding:10px 16px;
                                  box-shadow:0 1px 3px rgba(15,23,42,0.05);
                                  min-width:240px;">
                        <div style="display:flex;align-items:center;gap:6px;
                                    margin-bottom:4px;">
                          <span style="width:7px;height:7px;border-radius:50%;
                                       background:{_freshness_dot};
                                       box-shadow:0 0 0 3px {_freshness_dot}22;"></span>
                          <span style="font-size:10px;font-weight:700;
                                       text-transform:uppercase;letter-spacing:0.08em;
                                       color:{_freshness_fg};background:{_freshness_bg};
                                       padding:2px 8px;border-radius:10px;">
                            {_freshness_text}
                          </span>
                        </div>
                        <div style="font-size:14px;font-weight:700;color:#0f172a;
                                    letter-spacing:-0.01em;line-height:1.3;
                                    font-variant-numeric:tabular-nums;">
                          {_latest_label}
                        </div>
                        <div style="font-size:11px;color:#6b7280;margin-top:3px;
                                    font-weight:500;">
                          Report month <b style="color:#374151;">{report_file_month}</b>
                          &nbsp;&nbsp; {_n_files} file{"" if _n_files == 1 else "s"}
                        </div>
                      </div>
                    </div>
                    ''',
                    unsafe_allow_html=True,
                )
            else:
                st.markdown(
                    '<div style="display:flex;justify-content:flex-end;'
                    'padding-top:14px;color:#9ca3af;font-size:12px;">'
                    'No data loaded</div>',
                    unsafe_allow_html=True,
                )

    # -- Main area: guard - no data loaded ------------------------
    if file_catalog is None:
        st.markdown(
            """
            <div style="
                margin: 48px auto 0 auto;
                max-width: 400px;
                text-align: center;
                padding: 48px 32px;
                border: 1px dashed #d9d9d9;
                border-radius: 8px;
                background: #fafafa;
            ">
                <p style="font-size:1.6rem;margin:0 0 14px 0;color:#d9d9d9;">&#8679;</p>
                <p style="font-size:0.95rem;font-weight:600;color:#1a1a1a;margin:0 0 6px 0;">
                    Upload report files to begin
                </p>
                <p style="font-size:0.82rem;color:#8c8c8c;margin:0;line-height:1.5;">
                    Use the <b>sidebar</b> to upload CSV / Excel files<br>
                    or a single ZIP containing daily reports
                </p>
            </div>
            """,
            unsafe_allow_html=True,
        )
        empty_left, empty_center, empty_right = st.columns([1, 1.2, 1])
        with empty_center:
            main_uploaded = st.file_uploader(
                "Upload File",
                type=["zip", "csv", "xlsx", "xls"],
                accept_multiple_files=True,
                label_visibility="collapsed",
                help="Upload daily G5 CSV/Excel files, or one ZIP.",
                key="main_empty_state_upload",
            )
            if main_uploaded:
                st.session_state["force_upload_mode"] = True
                st.session_state["uploaded_file_catalog"] = build_file_catalog_from_uploads(main_uploaded)
                st.rerun()
        st.stop()

    # -- Top page navigation (above filters - always visible) -----
    _NAV_PAGES = ["Leaderboard", "Overview", "Budget Review", "Trend", "Export"]
    st.markdown('<div class="main-page-nav">', unsafe_allow_html=True)
    selected_page = st.pills(
        "main_page_nav",
        options=_NAV_PAGES,
        selection_mode="single",
        default="Leaderboard",
        label_visibility="collapsed",
        key="main_page_nav",
    )
    st.markdown('</div>', unsafe_allow_html=True)
    if not selected_page:
        selected_page = "Leaderboard"

    st.markdown(
        '<hr style="margin:6px 0 14px 0;border:none;border-top:1px solid #e8eaed;">',
        unsafe_allow_html=True,
    )

    # -- Inline filters (Properties + Stay Month) -----------------
    _fc1, _fc2, _fc3 = st.columns([3, 1.6, 1.6])

    with _fc1:
        st.markdown(
            '<p style="font-size:0.76rem;font-weight:600;color:#6b7280;'
            'text-transform:uppercase;letter-spacing:0.06em;margin:0 0 4px 0;">Properties</p>',
            unsafe_allow_html=True,
        )
        selected_hotels = st.multiselect(
            "Properties",
            options=all_hotels,
            default=[h for h in _default_hotels if h in all_hotels] or list(all_hotels),
            label_visibility="collapsed",
            placeholder="Select properties",
            key="main_hotel_multiselect",
        )

    with _fc2:
        st.markdown(
            '<p style="font-size:0.76rem;font-weight:600;color:#6b7280;'
            'text-transform:uppercase;letter-spacing:0.06em;margin:0 0 4px 0;">Stay Month - From</p>',
            unsafe_allow_html=True,
        )
        _start_idx = all_stay_months.index(_default_start) if _default_start in all_stay_months else 0
        stay_start = st.selectbox(
            "Stay Month From",
            all_stay_months,
            index=_start_idx,
            label_visibility="collapsed",
            key="main_stay_start",
        )

    with _fc3:
        st.markdown(
            '<p style="font-size:0.76rem;font-weight:600;color:#6b7280;'
            'text-transform:uppercase;letter-spacing:0.06em;margin:0 0 4px 0;">Stay Month - To</p>',
            unsafe_allow_html=True,
        )
        _end_idx = all_stay_months.index(_default_end) if _default_end in all_stay_months else len(all_stay_months) - 1
        stay_end = st.selectbox(
            "Stay Month To",
            all_stay_months,
            index=_end_idx,
            label_visibility="collapsed",
            key="main_stay_end",
        )

    st.markdown(
        '<hr style="margin:8px 0 14px 0;border:none;border-top:1px solid #e8eaed;">',
        unsafe_allow_html=True,
    )

    # Build stay_month_selection from start/end dropdowns
    _si = all_stay_months.index(stay_start) if stay_start in all_stay_months else 0
    _ei = all_stay_months.index(stay_end) if stay_end in all_stay_months else len(all_stay_months) - 1
    if _ei < _si:
        _ei = _si
    selected_stay_months_raw = all_stay_months[_si : _ei + 1]
    stay_month_selection = normalize_stay_month_selection(selected_stay_months_raw)

    # -- Guard: no hotels selected --------------------------------
    if not selected_hotels:
        st.warning("Select at least one property above.")
        st.stop()

    # -- Filter data -----------------------------------------------
    metric_data = metric_long[metric_long["Hotel"].isin(selected_hotels)].copy()
    metric_data = apply_stay_month_filter(metric_data, stay_month_selection)

    kpi_metric_data = metric_long[metric_long["Hotel"].isin(selected_hotels)].copy()
    kpi_metric_data = apply_stay_month_filter(kpi_metric_data, stay_month_selection)

    if metric_data.empty:
        st.warning("No data for the current filter selection.")
        st.stop()

    # -- Pre-build summaries ---------------------------------------
    movement_summary = build_movement_summary(metric_data, role_selection)
    pace_summary     = build_pace_summary(metric_data, role_selection)
    final_comparison = build_final_comparison(metric_data, role_selection)

    d4 = metric_data[metric_data["Reference"] == "Duetto"].copy()
    momentum_summary = (
        d4.groupby(["Report Date", "Report Label"])["Value"].sum()
        .reset_index().sort_values("Report Date")
        if not d4.empty else pd.DataFrame()
    )
    hotel_momentum = (
        d4.groupby(["Report Date", "Report Label", "Hotel", "Stay Month"])["Value"].sum()
        .reset_index().sort_values(["Hotel", "Stay Month", "Report Date"])
        if not d4.empty else pd.DataFrame()
    )

    # -- Page content (driven by top-nav pills) -------------------
    if selected_page == "Overview":
        # Build variance pivot (OTB + variances)
        latest_pivot = build_variance_pivot_table(metric_data, role_selection)

        # -- 1. Compare On The Book (KPI cards) -------------------
        _render_otb_comparison_chart(latest_pivot)

        # -- 2. Forecast VS Budget (KPI cards) --------------------
        _render_forecast_vs_budget(latest_pivot)

        # -- 3. Forecast Pivot with variance columns ---------------
        render_compact_hotel_tabs(latest_pivot)

        # -- 4. Same-Time Pace Benchmark (always visible) ---------
        st.markdown('<div class="section-title">Same-Time Pace Benchmark</div>', unsafe_allow_html=True)
        st.caption("Today VS STLY / ST2Y / ST3Y - how today's on-the-book compares to same-time prior years.")
        pace_view = pace_summary.copy()
        if not pace_view.empty and "Stay Month" in pace_view.columns:
            current_month_key = month_sort_key(report_file_month)
            pace_view["_month_sort"] = pace_view["Stay Month"].apply(month_sort_key)
            pace_view = (
                pace_view[pace_view["_month_sort"] >= current_month_key]
                .sort_values(["_month_sort", "Hotel", "Metric"])
                .drop(columns=["_month_sort"])
                .reset_index(drop=True)
            )

        if pace_view.empty:
            st.info("No pace data.")
        else:
            _pace_c1, _pace_c2 = st.columns([1.2, 1])
            view_mode_pace = _pace_c1.selectbox(
                "Pace view", ["Hotel tabs", "List view"], index=0, key="analysis_pace_view",
            )
            pace_layout = _pace_c2.radio(
                "Layout", ["Cards", "Compact table"], index=0, horizontal=True, key="pace_compact_layout",
            )
            pace_display = make_recommended_pace_compact(pace_view)
            if pace_layout == "Cards":
                render_pace_cards(pace_view)
            else:
                render_compact_by_hotel(pace_display, view_mode_pace, "pace")

        # -- 5. Historical Final Comparison (collapsed, coloured) -
        st.markdown('<div class="section-title">Historical Final Comparison</div>', unsafe_allow_html=True)
        with st.expander("Expand - Duetto vs Final LY / 2Y / 3Y", expanded=False):
            st.caption("Duetto vs Final LY / Final 2Y / Final 3Y. Historical context only - not a budget target.")
            if final_comparison.empty:
                st.info("No final comparison data.")
            else:
                view_mode_final = st.selectbox(
                    "Final view", ["Hotel tabs", "List view"], index=0, key="analysis_final_view",
                )
                final_display = make_final_compact(final_comparison)
                render_compact_by_hotel(final_display, view_mode_final, "final")

        # -- 6. Forecast Movement (collapsed by default) -----------
        with st.expander("Duetto Movement", expanded=False):
            forecast_movement_summary = render_forecast_movement_table_only(metric_data, role_selection)


    elif selected_page == "Budget Review":
        leaderboard_summary = render_budget_sort_board_v32(
            metric_long=metric_long,
            role_selection=role_selection,
            selected_hotels=selected_hotels,
            stay_month_selection=stay_month_selection,
        )


    elif selected_page == "Leaderboard":
        st.markdown(
            '<div class="section-title">Leaderboard - Key Metrics by Stay Month</div>',
            unsafe_allow_html=True,
        )
        st.caption(
            "One clean table per stay month. Selected properties are highlighted in blue. "
            "Green = above budget, red = below budget. Room Nights excluded."
        )
        render_mega_leaderboard(
            metric_long=metric_long,
            role_selection=role_selection,
            hotels=selected_hotels,
            stay_month_selection=stay_month_selection,
            report_file_month=report_file_month,
        )


    elif selected_page == "Trend":
        # ----------------------------------------------------------
        # TREND TAB - narrative flow:
        #   1) Strategic - two metrics vs baseline (high-level read)
        #   2) Forecast trend over report days (absolute values)
        #   3) Daily % change (momentum view)
        #   4) Hotel-level drill-down (bubble chart)
        # ----------------------------------------------------------

        # 1. Strategic - compare two metrics with shared baseline
        render_trend_comparison(metric_data, role_selection)

        st.divider()

        # 2. Absolute-value trend by stay month
        trend_summary = render_forecast_trend_by_month_v3(metric_data)

        st.divider()

        # 3. Daily % change bar chart (momentum)
        render_revenue_momentum_pct_chart(metric_data)

        st.divider()

        # 4. Hotel-level momentum (drill-down)
        st.markdown('<div class="section-title">Hotel Momentum - Per-Property Drill-Down</div>', unsafe_allow_html=True)
        st.caption("Bubble size = forecast magnitude  Bubble color = daily % change. Use this to spot which hotels are driving the overall trend.")

        if hotel_momentum.empty:
            st.info("No hotel momentum data.")
        else:
            bubble = hotel_momentum.copy()
            # Bubble-specific Stay Month filter
            if "Stay Month" in bubble.columns:
                bubble_month_options = sorted(bubble["Stay Month"].dropna().unique(), key=month_sort_key)
                bubble_month = st.selectbox(
                    "Bubble Stay Month",
                    ["All Months"] + bubble_month_options,
                    index=0,
                    key="bubble_stay_month_filter_v31",
                )
                if bubble_month != "All Months":
                    bubble = bubble[bubble["Stay Month"] == bubble_month].copy()

            bubble = bubble.sort_values(["Hotel", "Stay Month", "Report Date"])
            bubble["Previous Forecast"] = bubble.groupby(["Hotel", "Stay Month"])["Value"].shift(1)
            bubble["Daily Change"] = bubble["Value"] - bubble["Previous Forecast"]
            bubble["Daily Change %"] = bubble["Daily Change"] / bubble["Previous Forecast"] * 100
            bubble["Bubble Size"] = bubble["Value"].abs()

            latest_snapshot = (
                bubble.sort_values(["Hotel", "Report Date"])
                .groupby(["Hotel", "Stay Month"], as_index=False)
                .tail(1)
                .copy()
            )
            latest_snapshot["Abs Daily PU"] = latest_snapshot["Daily Change"].abs()

            ctl1, ctl2, ctl3 = st.columns([1.1, 1.1, 1.1])

            max_hotels = int(max(1, latest_snapshot["Hotel"].nunique()))
            top_options_raw = [5, 8, 10, 15, 20]
            top_options = [f"Top {n}" for n in top_options_raw if n <= max_hotels]
            if not top_options:
                top_options = [f"Top {max_hotels}"]
            top_options.append("All selected hotels")

            default_display = "Top 8" if "Top 8" in top_options else top_options[0]

            display_hotels = ctl1.selectbox(
                "Display hotels",
                top_options,
                index=top_options.index(default_display),
                key="bubble_display_hotels_dropdown",
            )

            focus_mode = ctl2.selectbox(
                "Focus by",
                [
                    "Biggest movement",
                    "Biggest drop",
                    "Biggest gain",
                    "Highest Forecast",
                ],
                index=0,
                key="bubble_focus_mode_dropdown",
            )

            size_mode = ctl3.selectbox(
                "Bubble size",
                [
                    "Abs Daily PU",
                    "Latest D4cast",
                ],
                index=0,
                key="bubble_size_mode_dropdown",
            )

            # Rank hotels based on what the user wants to focus on.
            if focus_mode == "Biggest movement":
                ranked_hotels = latest_snapshot.sort_values("Abs Daily PU", ascending=False)
            elif focus_mode == "Biggest drop":
                ranked_hotels = latest_snapshot.sort_values("Daily Change", ascending=True)
            elif focus_mode == "Biggest gain":
                ranked_hotels = latest_snapshot.sort_values("Daily Change", ascending=False)
            else:
                ranked_hotels = latest_snapshot.sort_values("Value", ascending=False)

            if display_hotels == "All selected hotels":
                visible_hotels = ranked_hotels["Hotel"].tolist()
            else:
                top_n = int(display_hotels.replace("Top ", ""))
                visible_hotels = ranked_hotels.head(top_n)["Hotel"].tolist()

            bubble_view = bubble[bubble["Hotel"].isin(visible_hotels)].copy()

            if size_mode == "Abs Daily PU":
                bubble_view["Bubble Size"] = bubble_view["Daily Change"].abs().fillna(0)
                # Avoid invisible first-date bubbles.
                if bubble_view["Bubble Size"].max() == 0:
                    bubble_view["Bubble Size"] = bubble_view["Value"].abs()
            else:
                bubble_view["Bubble Size"] = bubble_view["Value"].abs()

            bubble_view["Latest D4cast"] = bubble_view["Value"]
            bubble_view["Daily PU"] = bubble_view["Daily Change"]
            bubble_view["Daily PU %"] = bubble_view["Daily Change %"]

            fig_hotel = px.scatter(
                bubble_view,
                x="Report Date",
                y="Hotel",
                size="Bubble Size",
                color="Daily Change %",
                hover_data={
                    "Report Label": True,
                    "Latest D4cast": ":,.2f",
                    "Previous Forecast": ":,.2f",
                    "Daily PU": ":,.2f",
                    "Daily PU %": ":.2f",
                    "Bubble Size": False,
                },
                title="Hotel-level Forecast Momentum / Daily %",
                color_continuous_scale=["#b91c1c", "#facc15", "#15803d"],
                color_continuous_midpoint=0,
                size_max=34,
            )

            fig_hotel.update_layout(
                plot_bgcolor="rgba(0,0,0,0)",
                paper_bgcolor="rgba(0,0,0,0)",
                xaxis=dict(
                    title="Report Date",
                    showgrid=True,
                    gridcolor="#f1f5f9",
                    tickformat="%d %b",
                ),
                yaxis=dict(
                    title="Hotel",
                    showgrid=True,
                    gridcolor="#f8fafc",
                    categoryorder="array",
                    categoryarray=list(reversed(bubble_view["Hotel"].dropna().unique())),
                ),
                height=max(420, 46 * bubble_view["Hotel"].nunique()),
                margin=dict(l=20, r=20, t=60, b=20),
                coloraxis_colorbar=dict(title="Daily %", ticksuffix="%"),
            )

            fig_hotel.update_traces(
                marker=dict(line=dict(width=0.7, color="white"), opacity=0.86),
                selector=dict(mode="markers"),
            )

            st.plotly_chart(
                fig_hotel,
                use_container_width=True,
                config={
                    "displayModeBar": True,
                    "displaylogo": False,
                    "modeBarButtonsToRemove": ["lasso2d", "select2d"],
                },
            )

            st.markdown('<div class="section-title">Hotel Momentum Summary</div>', unsafe_allow_html=True)

            latest_rows = (
                bubble.sort_values(["Hotel", "Report Date"])
                .groupby(["Hotel", "Stay Month"], as_index=False)
                .tail(1)
                .copy()
            )

            latest_rows["Status"] = latest_rows["Daily Change"].apply(
                lambda x: "Up" if pd.notna(x) and x > 0 else "Down" if pd.notna(x) and x < 0 else "Flat"
            )

            summary_cols = [
                "Hotel",
                "Report Date",
                "Value",
                "Previous Forecast",
                "Daily Change %",
                "Status",
            ]

            summary_view = latest_rows[summary_cols].copy()
            summary_view = summary_view.rename(columns={
                "Value": "Latest D4cast",
                "Daily Change %": "Daily PU %",
            })
            summary_view = summary_view.sort_values("Daily PU %", ascending=True).reset_index(drop=True)

            def color_momentum_row(row):
                styles = pd.Series("", index=row.index)
                daily_pu = row.get("Daily PU %")
                if pd.notna(daily_pu) and daily_pu > 0:
                    for col in ["Daily PU %", "Status"]:
                        if col in styles.index:
                            styles[col] = "background-color: #bbf7d0; font-weight: 700"
                elif pd.notna(daily_pu) and daily_pu < 0:
                    for col in ["Daily PU %", "Status"]:
                        if col in styles.index:
                            styles[col] = "background-color: #fecaca; font-weight: 700"
                else:
                    if "Status" in styles.index:
                        styles["Status"] = "background-color: #fef08a; font-weight: 700"
                return styles

            st.dataframe(
                summary_view.style.format({
                    "Latest D4cast": fmt_raw2,
                    "Previous Forecast": fmt_raw2,
                    "Daily PU %": fmt_signed_pct2,
                }).apply(color_momentum_row, axis=1),
                use_container_width=True,
                hide_index=True,
                height=min(520, 44 + 36 * len(summary_view)),
            )

            with st.expander("Full hotel-level daily data"):
                full_view = bubble[[
                    "Report Date",
                    "Report Label",
                    "Hotel",
                    "Value",
                    "Previous Forecast",
                    "Daily Change",
                    "Daily Change %",
                ]].sort_values(["Hotel", "Report Date"]).reset_index(drop=True)

                full_view = full_view.rename(columns={
                    "Value": "Forecast",
                    "Daily Change": "Daily PU",
                    "Daily Change %": "Daily PU %",
                })

                st.dataframe(
                    full_view,
                    use_container_width=True,
                    hide_index=True,
                    height=520,
                    column_config={
                        "Report Date": st.column_config.DateColumn("Report Date", format="DD MMM YYYY"),
                        "Forecast": st.column_config.NumberColumn("Forecast", format="%,.2f"),
                        "Previous Forecast": st.column_config.NumberColumn("Previous Forecast", format="%,.2f"),
                        "Daily PU": st.column_config.NumberColumn("Daily PU", format="%,.2f"),
                        "Daily PU %": st.column_config.NumberColumn("Daily PU %", format="%.2f%%"),
                    },
                )







    elif selected_page == "Export":
        # -- Primary: Daily Briefing Excel -------------------------
        st.markdown('<div class="section-title">Daily Briefing Excel - One-Click Morning Deck</div>', unsafe_allow_html=True)
        st.caption(
            "Single multi-sheet workbook with every key view from this dashboard - "
            "ready to share with GMs or open at the morning meeting."
        )

        def trigger_download_toast():
            st.toast("File downloaded successfully.")

        # Build all sheets
        briefing_sheets = build_daily_briefing_sheets(metric_data, role_selection, report_file_month)

        # -- Sheet preview strip -----------------------------------
        preview_html = '<div style="display:flex;flex-wrap:wrap;gap:8px;margin:8px 0 16px 0;">'
        sheet_icons = {
            "Portfolio Snapshot": "",
            "Hotel Scorecard":    "",
            "Variance Pivot":     "",
            "Same-Time Pace":     "",
            "Historical Final":   "",
            "Duetto Movement":    "",
            "Hotel Momentum":     "",
            "Role Selection":     "",
        }
        for name, df in briefing_sheets.items():
            n_rows = 0 if (df is None or df.empty) else len(df)
            empty = n_rows == 0
            bg = "#f9fafb" if empty else "#eff6ff"
            bd = "#e5e7eb" if empty else "#bfdbfe"
            fg = "#9ca3af" if empty else "#1e40af"
            preview_html += (
                f'<div style="background:{bg};border:1px solid {bd};border-radius:8px;'
                f'padding:8px 12px;font-size:12px;display:inline-flex;align-items:center;gap:8px;">'
                f'<span style="font-size:14px;opacity:0.7;">{sheet_icons.get(name, "")}</span>'
                f'<span style="color:{fg};font-weight:600;">{html.escape(name)}</span>'
                f'<span style="color:#94a3b8;font-size:11px;font-variant-numeric:tabular-nums;">'
                f'{n_rows:,} row{"" if n_rows == 1 else "s"}</span>'
                f'</div>'
            )
        preview_html += '</div>'
        st.markdown(preview_html, unsafe_allow_html=True)

        # -- Primary download button -------------------------------
        file_stamp = report_file_month.replace(", ", "_").replace(" ", "_")
        st.download_button(
            "  Download Daily Briefing Excel",
            data=to_excel_bytes(briefing_sheets),
            file_name=f"g5_daily_briefing_{file_stamp}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            on_click=trigger_download_toast,
            type="primary",
            use_container_width=True,
        )

        # -- Secondary: quick CSVs ---------------------------------
        st.markdown('<div class="section-title">Quick CSV Exports</div>', unsafe_allow_html=True)
        st.caption("Individual CSVs for spreadsheet pivots or sharing a single view.")

        csv_c1, csv_c2 = st.columns(2)
        with csv_c1:
            st.download_button(
                "Duetto Movement (CSV)",
                data=movement_summary.to_csv(index=False).encode("utf-8"),
                file_name=f"g5_duetto_movement_{file_stamp}.csv",
                mime="text/csv",
                on_click=trigger_download_toast,
                type="secondary",
                use_container_width=True,
            )
        with csv_c2:
            var_pivot = briefing_sheets.get("Variance Pivot")
            if isinstance(var_pivot, pd.DataFrame) and not var_pivot.empty:
                st.download_button(
                    "Variance Pivot (CSV)",
                    data=var_pivot.to_csv(index=False).encode("utf-8"),
                    file_name=f"g5_variance_pivot_{file_stamp}.csv",
                    mime="text/csv",
                    on_click=trigger_download_toast,
                    type="secondary",
                    use_container_width=True,
                )

        # -- Role validation (kept at bottom for transparency) -----
        with st.expander("Report Roles Validation", expanded=False):
            st.caption("Confirms which files were assigned to Today / Yesterday / 7D / 1st Month roles.")
            st.dataframe(role_selection, use_container_width=True, hide_index=True)


if __name__ == "__main__":
    main()
