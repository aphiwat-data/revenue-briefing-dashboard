"""
Pandas Styler functions for dashboard tables.

Each function takes a DataFrame and returns a Styler with color-coded cells:
- Today/Duetto/Budget columns: tinted backgrounds (blue/green/amber)
- Variance ("VS") columns: green if positive, red if negative, yellow if zero
"""
from __future__ import annotations

import pandas as pd

from src.core.helpers import fmt_raw2

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
