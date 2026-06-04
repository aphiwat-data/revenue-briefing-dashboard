"""
Pivot table builders for Duetto data.

build_latest_pivot_table — wide view of all references for the latest snapshot.
build_variance_pivot_table — adds color-coded variance % columns.

Both apply the look-back pattern: per (Hotel, Stay Month, Metric, Reference)
they take the row from the latest available Report Date — so past stay months
pull from their own latest report, not the global "Today" file.
"""
from __future__ import annotations

import pandas as pd

from src.core.constants import METRIC_ORDER
from src.core.helpers import month_sort_key

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
