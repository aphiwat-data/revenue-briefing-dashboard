"""
Budget review pipeline.

build_budget_review              — per (hotel, stay month, metric) variance vs budget.
build_budget_review_summary_view — aggregates to hotel / stay-month / metric levels.
build_forecast_movement_v31      — forecast vs 1-day / 7-day / first-of-month.
"""
from __future__ import annotations

import pandas as pd

from src.core.constants import METRIC_ORDER
from src.domain.aggregations import risk_level
from src.domain.helpers import calc_budget_variance, budget_status_from_variance


def build_budget_review(metric_data: pd.DataFrame, role_selection: pd.DataFrame) -> pd.DataFrame:
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

def build_budget_review_summary_view(budget_df: pd.DataFrame, view_level: str) -> pd.DataFrame:
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

def build_forecast_movement_v31(metric_data: pd.DataFrame, role_selection: pd.DataFrame) -> pd.DataFrame:
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
