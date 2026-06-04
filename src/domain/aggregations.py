"""
Cross-table aggregations and risk scoring.

Each aggregation applies the look-back pattern (latest snapshot per
Hotel/Stay-Month/Metric/Reference) so past stay months work correctly.
"""
from __future__ import annotations

import pandas as pd

from src.core.constants import FINAL_REFS, METRIC_ORDER

def risk_level(diff_pct: float | None) -> str:
    if pd.isna(diff_pct): return "Unknown"
    if diff_pct <= -5: return "High"
    if diff_pct <= -2: return "Medium"
    return "Low"

def build_movement_summary(metric_data: pd.DataFrame, role_selection: pd.DataFrame) -> pd.DataFrame:
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

def build_pace_summary(metric_data: pd.DataFrame, role_selection: pd.DataFrame) -> pd.DataFrame:
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

def build_final_comparison(metric_data: pd.DataFrame, role_selection: pd.DataFrame) -> pd.DataFrame:
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

def make_recommended_pace_compact(df: pd.DataFrame) -> pd.DataFrame:
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

def make_final_compact(df: pd.DataFrame) -> pd.DataFrame:
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
