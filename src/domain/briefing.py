"""
Daily briefing builders.

build_portfolio_snapshot — per-metric totals + variance for the morning headline.
build_hotel_scorecard    — one row per hotel with Rev/ADR/Occ + variance.
build_daily_briefing_sheets — orchestrates the 8-sheet Excel briefing.
"""
from __future__ import annotations

import pandas as pd

from src.domain.aggregations import build_pace_summary, build_final_comparison, build_movement_summary
from src.domain.budget import build_forecast_movement_v31
from src.domain.pivot import build_variance_pivot_table

def build_portfolio_snapshot(metric_data: pd.DataFrame, role_selection: pd.DataFrame) -> pd.DataFrame:
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

def build_hotel_scorecard(metric_data: pd.DataFrame, role_selection: pd.DataFrame) -> pd.DataFrame:
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

def build_daily_briefing_sheets(
    metric_data: pd.DataFrame,
    role_selection: pd.DataFrame,
    report_file_month: str,
) -> dict[str, pd.DataFrame]:
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
    sheets["Variance Pivot"] = build_variance_pivot_table(metric_data, role_selection)

    # 4. Same-Time Pace
    sheets["Same-Time Pace"] = build_pace_summary(metric_data, role_selection)

    # 5. Historical Final
    sheets["Historical Final"] = build_final_comparison(metric_data, role_selection)

    # 6. Forecast Movement (prefer the v31 builder; fall back to summary on error)
    try:
        sheets["Duetto Movement"] = build_forecast_movement_v31(metric_data, role_selection)
    except Exception:
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
