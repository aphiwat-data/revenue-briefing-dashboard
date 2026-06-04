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
from src.domain.pivot import (
    build_latest_pivot_table,
    build_variance_pivot_table,
)
from src.domain.aggregations import (
    build_final_comparison,
    build_movement_summary,
    build_pace_summary,
    make_final_compact,
    make_recommended_pace_compact,
    risk_level,
)
from src.domain.briefing import (
    build_daily_briefing_sheets,
    build_hotel_scorecard,
    build_portfolio_snapshot,
)
from src.domain.budget import (
    build_budget_review,
    build_budget_review_summary_view,
    build_forecast_movement_v31,
)
from src.ui.dataframe_stylers import (
    style_final_variance_table,
    style_latest_pivot_table,
    style_pace_variance_table,
    style_variance_pivot,
)
from src.ui.sidebar import hotel_key, selected_property_chips_html
from src.features.overview import (
    _render_comparison_summary_cards,
    _render_forecast_vs_budget,
    _render_otb_comparison_chart,
    render_compact_hotel_tabs,
)
from src.features.shared import (
    render_compact_by_hotel,
    render_pace_cards,
)
from src.features.trend import (
    render_forecast_trend_by_month_v3,
    render_revenue_momentum_pct_chart,
    render_trend_comparison,
)
from src.features.leaderboard import (
    render_mega_leaderboard,
)
from src.features.budget_review import (
    render_budget_sort_board_v32,
    render_priority_budget_table,
)
from src.features.movement import (
    render_forecast_movement_table_only,
)


setup_page(st)
inject_styles(st)


# ============================================================
# Data Aggregation & Logic (Enhanced with Emojis)
# ============================================================



# ============================================================
# Daily Briefing Excel - one-click morning meeting deck
# ============================================================




























# ============================================================
# Trend Comparison - two metrics side-by-side with baselines
# ============================================================
# ============================================================
# Revenue Momentum % Change chart
# Formula: (revenue_t - revenue_{t-1}) / revenue_{t-1} x 100 per day
# ============================================================


# ============================================================
# Leaderboard by Stay Month
# ============================================================









KPI_AXIS_MAP = {
    "Budget": "Budget",
    "Forecast": "Forecast",
    "On The Book": "Today OTB",
}
# Order determines base axis: earlier = base.
# Forecast before OTB -> "On The Book vs Forecast" (OTB is compare, Forecast is base)
KPI_AXIS_ORDER = list(KPI_AXIS_MAP)




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
