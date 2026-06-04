"""
Shared render helpers used by multiple pages.

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
from src.domain.aggregations import make_recommended_pace_compact
from src.domain.helpers import (
    apply_stay_month_filter,
    short_hotel_name,
)

from src.ui.dataframe_stylers import (
    style_final_variance_table,
    style_pace_variance_table,
)


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
