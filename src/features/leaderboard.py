"""
Leaderboard page: 10-column mega ranking per stay month.

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
