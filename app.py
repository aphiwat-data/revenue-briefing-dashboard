"""
G5 Forecast Revenue Dashboard v4 (Pro UX/UI Edition)

Professional Streamlit dashboard for hotel revenue morning review.

Run:
    python -m streamlit run g5_d4cast_revenue_dashboard_v3.py

Install:
    python -m pip install streamlit pandas openpyxl plotly
"""

import io
import zipfile
import tempfile
import re
from pathlib import Path
from datetime import datetime

import pandas as pd
import streamlit as st
import plotly.express as px


# ============================================================
# Page setup & Custom CSS
# ============================================================

st.set_page_config(
    page_title="Daily Revenue Briefing Dashboard",
    page_icon="📈",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
    /* ================================================================
       G5 Revenue Dashboard — Design System
       Inspired by: Ant Design data tools / Linear / Retool
       Palette: neutral grays + single blue accent + semantic data colors
       ================================================================ */

    /* ── Layout & spacing ────────────────────────────────── */
    .block-container {
        padding-top: 2.5rem;
        padding-bottom: 2.5rem;
        max-width: 97%;
    }

    /* ── Sidebar ─────────────────────────────────────────── */
    [data-testid="stSidebar"] {
        background: #fafafa;
        border-right: 1px solid #f0f0f0;
    }
    /* Section label inside sidebar */
    [data-testid="stSidebar"] .stMarkdown h2 {
        font-size: 0.65rem !important;
        font-weight: 700 !important;
        text-transform: uppercase !important;
        letter-spacing: 0.1em !important;
        color: #bfbfbf !important;
        margin: 1.1rem 0 0.45rem !important;
        padding-bottom: 0.3rem !important;
        border-bottom: 1px solid #f0f0f0 !important;
    }
    [data-testid="stSidebar"] .stMarkdown h4 {
        font-size: 0.78rem !important;
        font-weight: 600 !important;
        color: #595959 !important;
        margin: 0.85rem 0 0.3rem !important;
    }
    [data-testid="stSidebar"] label {
        font-size: 0.8rem !important;
        color: #595959 !important;
    }
    [data-testid="stSidebar"] .stCaption,
    [data-testid="stSidebar"] div[data-testid="stCaptionContainer"] {
        color: #bfbfbf !important;
        font-size: 0.75rem !important;
    }
    [data-testid="stSidebar"] hr {
        border-color: #f0f0f0 !important;
        margin: 0.75rem 0 !important;
    }
    [data-testid="stSidebar"] .stButton > button {
        font-size: 0.8rem !important;
        font-weight: 500 !important;
        border-radius: 5px !important;
        height: 32px !important;
        padding: 0 10px !important;
    }

    /* ── Page title (native st.markdown "## ...") ──────────── */
    /* ใช้ h2 native เพื่อให้ Streamlit จัดการ spacing เองบน Cloud */
    .block-container h2 {
        font-size: 1.3rem !important;
        font-weight: 600 !important;
        color: #1a1a1a !important;
        letter-spacing: -0.01em !important;
    }
    .filter-bar {
        display: inline-flex;
        align-items: center;
        gap: 6px;
        background: #fafafa;
        border: 1px solid #f0f0f0;
        border-radius: 4px;
        padding: 4px 10px;
        font-size: 0.78rem;
        color: #595959;
        margin: 10px 0 14px 0;
    }
    .filter-bar b { color: #1a1a1a; font-weight: 600; }
    .filter-sep { color: #d9d9d9; margin: 0 2px; }

    /* ── KPI cards ───────────────────────────────────────── */
    div[data-testid="stMetric"] {
        background: #ffffff;
        border: 1px solid #f0f0f0;
        border-radius: 6px;
        padding: 14px 16px 12px 16px;
    }
    div[data-testid="stMetricLabel"] {
        font-size: 0.68rem !important;
        font-weight: 600 !important;
        text-transform: uppercase !important;
        letter-spacing: 0.08em !important;
        color: #8c8c8c !important;
    }
    div[data-testid="stMetricValue"] {
        font-size: clamp(1.05rem, 1.35vw, 1.5rem) !important;
        font-weight: 600 !important;
        color: #1a1a1a !important;
        letter-spacing: -0.01em !important;
    }
    div[data-testid="stMetricDelta"] {
        font-size: 0.76rem !important;
        font-weight: 500 !important;
    }

    /* ── Compare toggle chips ────────────────────────────── */
    .stCheckbox > label {
        background: #fafafa !important;
        border: 1px solid #e0e0e0 !important;
        border-radius: 4px !important;
        padding: 3px 10px 3px 6px !important;
        font-size: 0.81rem !important;
        font-weight: 500 !important;
        color: #595959 !important;
        transition: border-color 0.15s, background 0.15s, color 0.15s !important;
        cursor: pointer !important;
        display: inline-flex !important;
        align-items: center !important;
    }
    .stCheckbox > label:has(input:checked) {
        background: #e6f4ff !important;
        border-color: #1677ff !important;
        color: #1677ff !important;
    }

    /* ── Section title ───────────────────────────────────── */
    .section-title {
        font-size: 0.875rem;
        font-weight: 600;
        color: #1a1a1a;
        margin: 1.25rem 0 0.6rem 0;
        padding-bottom: 0.5rem;
        border-bottom: 1px solid #f0f0f0;
    }

    /* ── Tabs ────────────────────────────────────────────── */
    .stTabs [data-baseweb="tab-list"] {
        gap: 0 !important;
        border-bottom: 1px solid #f0f0f0 !important;
        background: transparent !important;
        padding: 0 !important;
    }
    .stTabs [data-baseweb="tab"] {
        background: transparent !important;
        border-bottom: 2px solid transparent !important;
        border-radius: 0 !important;
        padding: 9px 18px !important;
        font-size: 0.82rem !important;
        font-weight: 500 !important;
        color: #8c8c8c !important;
        margin-bottom: -1px !important;
        transition: color 0.15s !important;
    }
    .stTabs [aria-selected="true"] {
        color: #1677ff !important;
        border-bottom-color: #1677ff !important;
        font-weight: 600 !important;
        background: transparent !important;
    }
    .stTabs [data-baseweb="tab-panel"] {
        padding-top: 18px !important;
    }

    /* ── Tables ──────────────────────────────────────────── */
    div[data-testid="stDataFrame"] {
        border: 1px solid #f0f0f0 !important;
        border-radius: 6px !important;
        overflow: hidden !important;
    }

    /* ── Buttons ─────────────────────────────────────────── */
    .stButton > button {
        border-radius: 5px !important;
        font-size: 0.82rem !important;
        font-weight: 500 !important;
        transition: all 0.15s !important;
    }
    .stButton > button[kind="primary"] {
        background: #1677ff !important;
        border-color: #1677ff !important;
    }
    .stButton > button[kind="primary"]:hover {
        background: #4096ff !important;
        border-color: #4096ff !important;
    }

    /* ── Radio ───────────────────────────────────────────── */
    .stRadio [data-testid="stWidgetLabel"] {
        font-size: 0.75rem !important;
        font-weight: 600 !important;
        text-transform: uppercase !important;
        letter-spacing: 0.07em !important;
        color: #8c8c8c !important;
    }

    /* ── Divider between compare row items ───────────────── */
    .compare-sep {
        font-size: 1.1rem;
        color: #e0e0e0;
        text-align: center;
        margin-top: 5px;
        user-select: none;
    }

    /* ── Color legend row ────────────────────────────────── */
    .legend-row {
        display: flex;
        align-items: center;
        gap: 14px;
        margin: 10px 0 14px 0;
        font-size: 0.77rem;
        color: #8c8c8c;
    }
    .legend-dot {
        display: inline-block;
        width: 8px; height: 8px;
        border-radius: 50%;
        margin-right: 5px;
        vertical-align: middle;
    }

    /* ── Status cards used in HTML ───────────────────────── */
    .rev-card-good  { background: #f6ffed; border-left: 3px solid #52c41a; color: #135200; padding: 10px 14px; border-radius: 4px; margin: 3px 0; }
    .rev-card-bad   { background: #fff2f0; border-left: 3px solid #ff4d4f; color: #820014; padding: 10px 14px; border-radius: 4px; margin: 3px 0; }
    .rev-card-flat  { background: #fffbe6; border-left: 3px solid #faad14; color: #614700; padding: 10px 14px; border-radius: 4px; margin: 3px 0; }
    .rev-card-info  { background: #e6f4ff; border-left: 3px solid #1677ff; color: #003eb3; padding: 10px 14px; border-radius: 4px; margin: 3px 0; }

    /* ── Expanders ───────────────────────────────────────── */
    .streamlit-expanderHeader {
        font-size: 0.82rem !important;
        font-weight: 500 !important;
        color: #595959 !important;
        background: #fafafa !important;
        border: 1px solid #f0f0f0 !important;
        border-radius: 5px !important;
    }

    /* ── Selectbox / inputs ──────────────────────────────── */
    .stSelectbox > div > div { border-radius: 5px !important; }
    .stTextInput > div > div > input { border-radius: 5px !important; }

    /* ── Misc ────────────────────────────────────────────── */
    div[data-testid="stHorizontalBlock"] { gap: 0.75rem; }
    .stCaption, div[data-testid="stCaptionContainer"] {
        color: #8c8c8c !important;
        font-size: 0.77rem !important;
    }
    #MainMenu { visibility: hidden; }
    footer { visibility: hidden; }
    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# Constants
# ============================================================

SUPPORTED_EXTENSIONS = [".csv", ".xlsx", ".xls", ".zip"]

MONTH_FORMATS_TRY = [
    "%b, %Y", "%b-%y", "%b, %y", "%b,%y", "%b %Y", "%b %y",
    "%B, %Y", "%B-%y", "%B %Y", "%B %y",
]

REFERENCE_PATTERNS = {
    "Today": ["Today"],
    "STLY": ["STLY (DOW)", "STLY"],
    "ST2Y": ["ST2Y (DOW)", "ST2Y"],
    "ST3Y": ["ST3Y (DOW)", "ST3Y"],
    "Duetto": ["Duetto Forecast", "Duetto", "Forecast", "Forecast"],
    "Budget": ["Locked Budget", "Budget"],
    "Final LY": ["Final LY (DOW)", "Final LY"],
    "Final 2Y": ["Final 2Y (DOW)", "Final 2Y"],
    "Final 3Y": ["Final 3Y (DOW)", "Final 3Y"],
}

METRIC_PATTERNS = {
    "Occupancy": ["Occupancy (Physical)", "Occupancy"],
    "Rooms": ["Rooms (Commit)", "Rooms"],
    "ADR": ["ADR (Commit)", "ADR"],
    "Revenue": ["Room Revenue (Commit)", "Room Revenue", "Revenue"],
}

METRIC_TO_DISPLAY = {
    "Occupancy": "Occ",
    "Rooms": "Room",
    "ADR": "ADR",
    "Revenue": "Rev",
}

DISPLAY_TO_METRIC = {v: k for k, v in METRIC_TO_DISPLAY.items()}
METRIC_ORDER = ["Occ", "Room", "ADR", "Rev"]
SAME_TIME_REFS = ["STLY", "ST2Y", "ST3Y"]
FINAL_REFS = ["Final LY", "Final 2Y", "Final 3Y"]


# ============================================================
# Generic helpers
# ============================================================

def clean_text(x):
    if x is None:
        return ""
    return str(x).strip().strip('"').strip()

def normalize_stay_month(raw):
    if raw is None or (isinstance(raw, float) and pd.isna(raw)):
        return None
    s = str(raw).strip()
    if not s:
        return None

    for fmt in MONTH_FORMATS_TRY:
        try:
            dt = datetime.strptime(s, fmt)
            return dt.strftime("%b, %Y")
        except ValueError:
            continue

    m = re.match(r"^([A-Za-z]+)[\s\-,\,]+(\d{2,4})$", s)
    if m:
        month_part = m.group(1).strip().title()
        year_part = m.group(2).strip()
        try:
            yr = int(year_part)
            if yr < 100: yr += 2000
            try:
                dt = datetime.strptime(f"{month_part} 1 {yr}", "%b %d %Y")
            except ValueError:
                dt = datetime.strptime(f"{month_part} 1 {yr}", "%B %d %Y")
            return dt.strftime("%b, %Y")
        except ValueError:
            return None
    return None

def month_sort_key(x):
    return pd.to_datetime(x, format="%b, %Y", errors="coerce")

def extract_date_from_filename(file_name):
    s = str(file_name)
    patterns = [r"\d{4}-\d{2}-\d{2}", r"\d{4}_\d{2}_\d{2}", r"\d{8}", r"\d{1,2}-\d{1,2}-\d{4}"]
    for pattern in patterns:
        m = re.search(pattern, s)
        if not m: continue
        date_text = m.group(0)
        if re.fullmatch(r"\d{8}", date_text):
            date_text = f"{date_text[:4]}-{date_text[4:6]}-{date_text[6:]}"
        date_text = date_text.replace("_", "-")
        dt = pd.to_datetime(date_text, errors="coerce", dayfirst=False)
        if pd.notna(dt):
            return dt.normalize()
    return pd.NaT

def file_modified_date(path):
    return pd.to_datetime(Path(path).stat().st_mtime, unit="s").normalize()

def trunc2(x):
    if x is None or pd.isna(x):
        return None
    # Truncate toward zero, not round.
    return int(float(x) * 100) / 100


def fmt_raw2(x):
    x = trunc2(x)
    if x is None:
        return ""
    return f"{x:,.2f}"


def fmt_pct2(x):
    x = trunc2(x)
    if x is None:
        return ""
    return f"{x:,.2f}%"

def fmt_signed_pct2(x):
    x = trunc2(x)
    if x is None:
        return ""
    return f"{x:+,.2f}%"


def safe_delta(current, base):
    if base is None or pd.isna(base) or base == 0:
        return None
    return fmt_pct2((current - base) / base * 100)


# ============================================================
# File catalog
# ============================================================

@st.cache_data(show_spinner=False)
def build_file_catalog_from_folder(folder_path_text):
    folder_path = Path(folder_path_text)
    if not folder_path.exists():
        raise FileNotFoundError(f"Folder not found: {folder_path}")
    files = [p for p in folder_path.iterdir() if p.is_file() and p.suffix.lower() in SUPPORTED_EXTENSIONS]
    if not files:
        raise FileNotFoundError(f"No CSV/Excel files found in {folder_path}")

    rows = []
    for p in files:
        report_date = extract_date_from_filename(p.name)
        if pd.isna(report_date):
            report_date = file_modified_date(p)
        rows.append({
            "Source": "Folder", "File Path": str(p), "File Bytes": None, "File Name": p.name,
            "Suffix": p.suffix.lower(), "Report Date": report_date,
            "Modified Time": pd.to_datetime(p.stat().st_mtime, unit="s"),
        })

    df = pd.DataFrame(rows).sort_values(["Report Date", "File Name"]).reset_index(drop=True)
    df["File Order"] = range(1, len(df) + 1)
    df["Report Label"] = df.apply(lambda r: f"{int(r['File Order']):02d} | {r['Report Date'].strftime('%Y-%m-%d')}", axis=1)
    return df


def build_file_catalog_from_uploads(uploaded_files):
    """
    Build file catalog from manually uploaded files.

    Supports:
    - Multiple CSV/XLSX/XLS files
    - One or more ZIP files containing CSV/XLSX/XLS reports

    ZIP uploads are expanded in memory, so parse_record receives only real
    report files (.csv/.xlsx/.xls), never the .zip itself.
    """
    rows = []

    for upload_order, uploaded_file in enumerate(uploaded_files, start=1):
        uploaded_name = uploaded_file.name
        suffix = Path(uploaded_name).suffix.lower()

        if suffix == ".zip":
            try:
                with zipfile.ZipFile(io.BytesIO(uploaded_file.getvalue())) as z:
                    for zip_info in z.infolist():
                        if zip_info.is_dir():
                            continue

                        inner_path = zip_info.filename
                        inner_name = Path(inner_path).name
                        inner_suffix = Path(inner_name).suffix.lower()

                        if inner_suffix not in [".csv", ".xlsx", ".xls"]:
                            continue
                        if inner_name.startswith("~$") or inner_name.startswith("."):
                            continue

                        report_date = extract_date_from_filename(inner_path)
                        if pd.isna(report_date):
                            report_date = pd.Timestamp.today().normalize()

                        rows.append({
                            "Source": "Upload",
                            "File Path": None,
                            "File Bytes": z.read(zip_info),
                            "File Name": inner_name,
                            "Original Path": inner_path,
                            "Suffix": inner_suffix,
                            "Report Date": report_date,
                            "Modified Time": pd.NaT,
                            "Original Upload Order": upload_order,
                            "Upload Container": uploaded_name,
                        })
            except zipfile.BadZipFile:
                raise ValueError(f"Invalid ZIP file: {uploaded_name}")

        elif suffix in [".csv", ".xlsx", ".xls"]:
            report_date = extract_date_from_filename(uploaded_name)
            if pd.isna(report_date):
                report_date = pd.Timestamp.today().normalize()

            rows.append({
                "Source": "Upload",
                "File Path": None,
                "File Bytes": uploaded_file.getvalue(),
                "File Name": uploaded_name,
                "Original Path": uploaded_name,
                "Suffix": suffix,
                "Report Date": report_date,
                "Modified Time": pd.NaT,
                "Original Upload Order": upload_order,
                "Upload Container": uploaded_name,
            })

    if not rows:
        raise ValueError("No valid CSV/XLSX/XLS files found. Upload daily files or a .zip folder containing them.")

    df = pd.DataFrame(rows).sort_values(["Report Date", "File Name"]).reset_index(drop=True)
    df["File Order"] = range(1, len(df) + 1)
    df["Report Label"] = df.apply(
        lambda r: f"{int(r['File Order']):02d} | {r['Report Date'].strftime('%Y-%m-%d')}",
        axis=1,
    )
    return df

def select_role_files(file_catalog, report_file_month):
    start = pd.to_datetime(report_file_month, format="%b, %Y")
    end = start + pd.offsets.MonthEnd(0)
    month_files = file_catalog[(file_catalog["Report Date"] >= start) & (file_catalog["Report Date"] <= end)].copy()
    
    if month_files.empty:
        raise ValueError(f"No report files found for {report_file_month}")

    latest = month_files.sort_values("Report Date").iloc[-1]
    latest_date = latest["Report Date"]
    previous_candidates = month_files[month_files["Report Date"] < latest_date].copy()

    today = latest
    yesterday = previous_candidates.sort_values("Report Date").iloc[-1] if not previous_candidates.empty else None

    target_7d = latest_date - pd.Timedelta(days=7)
    seven = None
    if not previous_candidates.empty:
        temp = previous_candidates.copy()
        temp["Distance To 7D"] = (temp["Report Date"] - target_7d).abs()
        seven = temp.sort_values(["Distance To 7D", "Report Date"]).iloc[0]

    first = month_files.sort_values("Report Date").iloc[0]
    rows = []

    def add(role, row):
        if row is None:
            rows.append({"Role": role, "Report Label": None, "Report Date": None, "File Name": None, "Status": "❌ Missing"})
        else:
            rows.append({"Role": role, "Report Label": row["Report Label"], "Report Date": row["Report Date"], "File Name": row["File Name"], "Status": "✅ OK"})

    add("Today / Latest", today)
    add("Yesterday / Previous", yesterday)
    add("Last 7D", seven)
    add("1st Month", first)

    return pd.DataFrame(rows), month_files.copy()


# ============================================================
# Parsing (Keeping Original Solid Logic)
# ============================================================

def parse_csv_bytes(file_bytes):
    content = file_bytes.decode("utf-8", errors="ignore")
    lines = content.splitlines()
    title = clean_text(lines[0].split(",")[-1]) if lines else ""
    generated = clean_text(lines[1].split(",")[-1]) if len(lines) > 1 else ""
    csv_body = "\n".join(lines[2:]) if len(lines) > 2 else ""
    df = pd.read_csv(io.StringIO(csv_body))
    return df, title, generated

def parse_csv_path(path):
    return parse_csv_bytes(Path(path).read_bytes())

def parse_excel_bytes(file_bytes, suffix=".xlsx"):
    engine = "openpyxl" if suffix == ".xlsx" else None
    raw = pd.read_excel(io.BytesIO(file_bytes), header=None, engine=engine)
    header_idx = None
    for idx in range(min(25, len(raw))):
        row_join = " ".join(raw.iloc[idx].astype(str).str.lower().tolist())
        if ("hotel" in row_join) and ("stay" in row_join or "month" in row_join):
            header_idx = idx
            break
    if header_idx is None:
        return pd.read_excel(io.BytesIO(file_bytes), engine=engine), "", ""
    title = clean_text(raw.iloc[0].dropna().astype(str).tolist()[-1]) if len(raw) > 0 and not raw.iloc[0].dropna().empty else ""
    generated = clean_text(raw.iloc[1].dropna().astype(str).tolist()[-1]) if len(raw) > 1 and not raw.iloc[1].dropna().empty else ""
    headers = raw.iloc[header_idx].tolist()
    df = raw.iloc[header_idx + 1:].copy()
    df.columns = headers
    df = df.dropna(how="all")
    return df, title, generated

def parse_excel_path(path):
    suffix = Path(path).suffix.lower()
    return parse_excel_bytes(Path(path).read_bytes(), suffix=suffix)

def standardize_df(df, report_label, report_date, report_order, file_name, title="", generated=""):
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    df = df.rename(columns={df.columns[0]: "Hotel", df.columns[1]: "Stay Month"})
    df["Stay Month"] = df["Stay Month"].apply(lambda x: normalize_stay_month(x) or x)
    df.insert(0, "Report Order", report_order)
    df.insert(1, "Report Label", report_label)
    df.insert(2, "Report Date", report_date)
    df.insert(3, "Report File", file_name)
    df.insert(4, "Report Title", title)
    df.insert(5, "Generated", generated)
    return df


def parse_record(row):
    suffix = str(row["Suffix"]).lower()

    if suffix == ".zip":
        raise ValueError(
            f"ZIP file reached parser unexpectedly: {row.get('File Name', '')}. "
            "The ZIP must be expanded in build_file_catalog_from_uploads first."
        )

    if suffix not in [".csv", ".xlsx", ".xls"]:
        raise ValueError(f"Unsupported file type: {suffix} | file={row.get('File Name', '')}")

    if row["Source"] == "Folder":
        path = Path(row["File Path"])
        if suffix == ".csv":
            df, title, generated = parse_csv_path(path)
        else:
            df, title, generated = parse_excel_path(path)
    else:
        file_bytes = row["File Bytes"]
        if file_bytes is None or (isinstance(file_bytes, float) and pd.isna(file_bytes)):
            raise ValueError(f"Missing file bytes for uploaded file: {row.get('File Name', '')}")
        if suffix == ".csv":
            df, title, generated = parse_csv_bytes(file_bytes)
        else:
            df, title, generated = parse_excel_bytes(file_bytes, suffix=suffix)

    return standardize_df(
        df,
        row["Report Label"],
        row["Report Date"],
        int(row["Report Order"]),
        row["File Name"],
        title,
        generated,
    )

# ============================================================
# Data Aggregation & Logic (Enhanced with Emojis)
# ============================================================

def build_ref_col_map(df):
    mapping = {}
    for ref, ref_keys in REFERENCE_PATTERNS.items():
        mapping[ref] = {}
        for metric, metric_keys in METRIC_PATTERNS.items():
            for c in df.columns:
                c_lower = str(c).lower()
                if any(str(r).lower() in c_lower for r in ref_keys) and any(str(m).lower() in c_lower for m in metric_keys):
                    mapping[ref][metric] = c
                    break
    return mapping

def build_metric_long(combined_df, ref_col_map):
    rows = []
    for _, r in combined_df.iterrows():
        hotel = r["Hotel"]
        if pd.isna(hotel) or str(hotel).strip() in ["", "Total"]: continue
        for ref, metric_cols in ref_col_map.items():
            for metric_full, col in metric_cols.items():
                val = pd.to_numeric(r.get(col), errors="coerce")
                if pd.isna(val): continue
                metric = METRIC_TO_DISPLAY[metric_full]
                if metric == "Occ": val *= 100
                rows.append({"Report Label": r["Report Label"], "Report Date": r["Report Date"], "Report File": r["Report File"], "Hotel": hotel, "Stay Month": r["Stay Month"], "Reference": ref, "Metric": metric, "Value": val})
    return pd.DataFrame(rows).sort_values(["Report Date", "Hotel", "Stay Month", "Reference", "Metric"]).reset_index(drop=True) if rows else pd.DataFrame()

def risk_level(diff_pct):
    if pd.isna(diff_pct): return "⚪ Unknown"
    if diff_pct <= -5: return "🔴 High"
    if diff_pct <= -2: return "🟠 Medium"
    return "🟢 Low"

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
                bv, diff, diff_pct, status = None, None, None, "⚪ No Base"
            else:
                bv = metric_data[(metric_data["Hotel"] == h) & (metric_data["Stay Month"] == sm) & (metric_data["Metric"] == m) & (metric_data["Report Label"] == base_label) & (metric_data["Reference"] == "Duetto")]["Value"].sum()
                if pd.isna(bv) or bv == 0:
                    diff, diff_pct, status = None, None, "⚪ No Base"
                else:
                    diff = lv - bv
                    diff_pct = diff / bv * 100
                    status = "🟢 Up" if diff > 0 else "🔴 Down" if diff < 0 else "🟡 Flat"
            rows.append({"Hotel": h, "Stay Month": sm, "Metric": m, "Compare": compare, "Latest D4cast": lv, "Base Forecast": bv, "Forecast Diff": diff, "Forecast Diff %": diff_pct, "Status": status, "Risk": risk_level(diff_pct)})
    
    out = pd.DataFrame(rows)
    if not out.empty:
        out["Compare"] = pd.Categorical(out["Compare"], ["vs Yesterday", "vs 7D", "vs 1st Month"], ordered=True)
        out = out.sort_values(["Hotel", "Stay Month", "Metric", "Compare"]).reset_index(drop=True)
    return out

def build_pace_summary(metric_data, role_selection):
    role_map = {row["Role"]: row["Report Label"] for _, row in role_selection.iterrows() if pd.notna(row["Report Label"])}
    latest = metric_data[metric_data["Report Label"] == role_map.get("Today / Latest")].copy()
    rows = []
    for keys, group in latest.groupby(["Hotel", "Stay Month", "Metric"]):
        h, sm, m = keys
        def val(ref): x = group[group["Reference"] == ref]["Value"]; return x.sum() if not x.empty else None
        today, stly, st2y, st3y = val("Today"), val("STLY"), val("ST2Y"), val("ST3Y")
        cands = [(r, v) for r, v in [("STLY", stly), ("ST2Y", st2y), ("ST3Y", st3y)] if pd.notna(v)]
        pace_ref, pace_value = max(cands, key=lambda x: x[1]) if cands else (None, None)
        
        if not pace_value or pd.isna(today):
            diff, diff_pct, status = None, None, "⚪ No Pace"
        else:
            diff = today - pace_value
            diff_pct = diff / pace_value * 100
            status = "🟢 Ahead" if diff > 0 else "🔴 Behind" if diff < 0 else "🟡 On Pace"
        rows.append({"Hotel": h, "Stay Month": sm, "Metric": m, "Today": today, "STLY": stly, "ST2Y": st2y, "ST3Y": st3y, "Recommended Pace": pace_ref, "Recommended Pace Value": pace_value, "Pace Diff": diff, "Pace Diff %": diff_pct, "Status": status, "Risk": risk_level(diff_pct)})
    return pd.DataFrame(rows).sort_values(["Hotel", "Stay Month", "Metric"]).reset_index(drop=True) if rows else pd.DataFrame()

def build_final_comparison(metric_data, role_selection):
    role_map = {row["Role"]: row["Report Label"] for _, row in role_selection.iterrows() if pd.notna(row["Report Label"])}
    latest = metric_data[metric_data["Report Label"] == role_map.get("Today / Latest")].copy()
    rows = []
    for keys, group in latest.groupby(["Hotel", "Stay Month", "Metric"]):
        d4 = group[group["Reference"] == "Duetto"]["Value"].sum()
        if pd.isna(d4) or d4 == 0: continue
        for final_ref in FINAL_REFS:
            base = group[group["Reference"] == final_ref]["Value"].sum()
            if pd.isna(base) or base == 0: continue
            diff = d4 - base
            rows.append({"Hotel": keys[0], "Stay Month": keys[1], "Metric": keys[2], "Forecast": d4, "Base Final": final_ref, "Final Value": base, "Diff": diff, "Diff %": diff / base * 100, "Status": "🟢 Higher" if diff > 0 else "🔴 Lower" if diff < 0 else "🟡 Equal"})
    return pd.DataFrame(rows).sort_values(["Hotel", "Stay Month", "Metric", "Base Final"]).reset_index(drop=True) if rows else pd.DataFrame()


def to_excel_bytes(sheets):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for name, df in sheets.items():
            df.to_excel(writer, sheet_name=name[:31], index=False)
    output.seek(0)
    return output.getvalue()



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



HOTEL_SHORT_NAMES = {
    "The Grass Serviced Suites": "TG",
    "Hotel Amber Pattaya": "Amber PTY",
    "Hotel Amber Sukhumvit 85": "Amber 85",
    "Altera Hotel & Residence Pattaya": "Altera",
    "Arbour Hotel and Residence": "Arbour",
    "Arden Hotel & Residence Pattaya": "Arden",
    "Aster Hotel & Residence Pattaya": "Aster",
}


def short_hotel_name(hotel_name):
    return HOTEL_SHORT_NAMES.get(str(hotel_name), str(hotel_name))


def build_latest_pivot_table(metric_data, role_selection):
    """
    Build compact pivot table like the reference screenshot:
    Month | Metric | Today | STLY | ST2Y | ST3Y | Duetto | Budget | Final LY | Final 2Y | Final 3Y

    Uses only Today / Latest report file.
    Keeps metric order: Occ, Room, ADR, Rev.
    """
    if metric_data is None or metric_data.empty:
        return pd.DataFrame()

    role_map = {
        row["Role"]: row["Report Label"]
        for _, row in role_selection.iterrows()
        if pd.notna(row["Report Label"])
    }
    latest_label = role_map.get("Today / Latest")

    latest = metric_data[metric_data["Report Label"] == latest_label].copy()
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


def style_latest_pivot_table(df):
    """
    Color compact table:
    - Duetto Rev lower than Budget Rev = red
    - Duetto Rev higher than Budget Rev = blue
    - Metric order already handled upstream
    """
    if df is None or df.empty:
        return df

    numeric_cols = [c for c in ["Today", "STLY", "ST2Y", "ST3Y", "Duetto", "Budget", "Final LY", "Final 2Y", "Final 3Y"] if c in df.columns]
    fmt = {c: fmt_raw2 for c in numeric_cols if "fmt_raw2" in globals()}

    def row_style(row):
        styles = pd.Series("", index=row.index)
        metric = str(row.get("Metric", ""))

        # Light banding by metric
        if metric == "Occ":
            styles[:] = "background-color: #f8fafc"
        elif metric == "Rev":
            styles[:] = "background-color: #fff7ed"

        # Focus cell: Duetto vs Budget on Rev row
        if metric == "Rev" and "Duetto" in row.index and "Budget" in row.index:
            d = row.get("Duetto")
            b = row.get("Budget")
            if pd.notna(d) and pd.notna(b):
                if d < b:
                    styles["Duetto"] = "background-color: #fecaca; font-weight: 700"
                elif d > b:
                    styles["Duetto"] = "background-color: #bfdbfe; font-weight: 700"

        return styles

    return df.style.format(fmt).apply(row_style, axis=1)


def render_compact_hotel_tabs(pivot_df):
    """
    Efficient tab layout matching user's screenshot.
    """
    if pivot_df is None or pivot_df.empty:
        st.info("No pivot data for selected filters.")
        return

    view_mode = st.radio(
        "View",
        ["Hotel tabs", "All hotels table"],
        horizontal=True,
        key="compact_pivot_view",
    )

    if view_mode == "Hotel tabs":
        hotels = sorted(pivot_df["Hotel"].dropna().unique())
        labels = [short_hotel_name(h) for h in hotels]

        tabs = st.tabs(labels)

        for tab, hotel in zip(tabs, hotels):
            with tab:
                sub = pivot_df[pivot_df["Hotel"] == hotel].drop(columns=["Hotel"]).reset_index(drop=True)
                st.dataframe(
                    style_latest_pivot_table(sub),
                    use_container_width=True,
                    hide_index=True,
                    height=min(650, 40 + 38 * len(sub)),
                )
    else:
        show = pivot_df.copy()
        show["Hotel"] = show["Hotel"].apply(short_hotel_name)
        st.dataframe(
            style_latest_pivot_table(show),
            use_container_width=True,
            hide_index=True,
            height=min(750, 40 + 30 * len(show)),
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
            return "⚪ No Base"
        if x > 0:
            return "🟢 Up"
        if x < 0:
            return "🔴 Down"
        return "🟡 Flat"

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


def render_color_leaderboard(metric_long, role_selection, selected_hotels, stay_month_selection):
    st.markdown('<div class="section-title">Hotel Comparison Leaderboard</div>', unsafe_allow_html=True)
    st.caption("Color-coded comparison by hotel. Red = drop / negative pickup, green = gain / positive pickup.")

    c1, c2, c3 = st.columns([1, 1, 1])

    compare_metric = c1.selectbox(
        "Metric",
        get_metric_options_with_all(),
        index=0,
        key="color_leaderboard_metric",
    )

    sort_by = c2.selectbox(
        "Sort by",
        [
            "Daily PU",
            "Daily PU %",
            "Latest D4cast",
            "7D PU",
            "7D PU %",
            "MTD PU",
            "MTD PU %",
        ],
        index=0,
        key="color_leaderboard_sort_by",
    )

    sort_mode = c3.selectbox(
        "Order",
        ["Worst first", "Best first"],
        index=0,
        key="color_leaderboard_sort_order",
    )

    if compare_metric == "All Metrics":
        all_matrix = build_hotel_compare_matrix_all(
            metric_long=metric_long,
            role_selection=role_selection,
            hotels=selected_hotels,
            stay_month_selection=stay_month_selection,
        )

        if all_matrix.empty:
            st.info("No leaderboard data for selected filters.")
            return pd.DataFrame()

        st.info("All Metrics view is split into metric tabs to prevent number clipping.")

        metric_tabs = st.tabs(metric_label_order())
        for mt, metric_name in zip(metric_tabs, metric_label_order()):
            with mt:
                matrix = all_matrix[all_matrix["Metric"] == metric_name].drop(columns=["Metric"]).copy()
                if sort_by in matrix.columns:
                    matrix = matrix.sort_values(sort_by, ascending=(sort_mode == "Worst first")).reset_index(drop=True)
                render_color_matrix_no_clip(matrix)

                chart_value = sort_by if sort_by in matrix.columns else "Daily PU"
                chart_df = matrix.copy()
                chart_df["Direction"] = chart_df[chart_value].apply(
                    lambda x: "Up" if pd.notna(x) and x > 0 else "Down" if pd.notna(x) and x < 0 else "Flat"
                )
                color_map = {"Up": "#15803d", "Down": "#b91c1c", "Flat": "#ca8a04"}
                fig = px.bar(
                    chart_df,
                    x=chart_value,
                    y="Hotel",
                    orientation="h",
                    color="Direction",
                    color_discrete_map=color_map,
                    title=f"{metric_name} Leaderboard by {chart_value}",
                )
                fig.update_layout(
                    plot_bgcolor="rgba(0,0,0,0)",
                    paper_bgcolor="rgba(0,0,0,0)",
                    height=max(360, 46 * len(chart_df)),
                    yaxis=dict(categoryorder="total ascending"),
                    xaxis=dict(showgrid=True, gridcolor="#f1f5f9"),
                    margin=dict(l=20, r=20, t=55, b=20),
                )
                fig.update_traces(texttemplate="%{x:,.2f}", textposition="outside", cliponaxis=False)
                st.plotly_chart(fig, use_container_width=True, config={"displaylogo": False})

        return all_matrix

    matrix = build_hotel_compare_matrix(
        metric_long=metric_long,
        role_selection=role_selection,
        hotels=selected_hotels,
        stay_month_selection=stay_month_selection,
        compare_metric=compare_metric,
    )

    if matrix.empty:
        st.info("No leaderboard data for selected filters.")
        return pd.DataFrame()

    if sort_by in matrix.columns:
        matrix = matrix.sort_values(sort_by, ascending=(sort_mode == "Worst first")).reset_index(drop=True)

    # KPI cards
    total_daily_pu = matrix["Daily PU"].sum() if "Daily PU" in matrix.columns else None
    hotels_up = (matrix["Daily PU"] > 0).sum() if "Daily PU" in matrix.columns else 0
    hotels_down = (matrix["Daily PU"] < 0).sum() if "Daily PU" in matrix.columns else 0

    if "Daily PU" in matrix.columns and not matrix["Daily PU"].dropna().empty:
        worst_row = matrix.sort_values("Daily PU", ascending=True).iloc[0]
        best_row = matrix.sort_values("Daily PU", ascending=False).iloc[0]
    else:
        worst_row = None
        best_row = None

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Total Daily PU", fmt_raw2(total_daily_pu))
    k2.metric("Hotels Up", int(hotels_up))
    k3.metric("Hotels Down", int(hotels_down))
    k4.metric("Worst Drop", str(worst_row["Hotel"]) if worst_row is not None else "-")

    render_color_matrix_no_clip(matrix)

    st.markdown("#### Quick comparison chart")

    chart_df = matrix.copy()
    chart_value = sort_by if sort_by in chart_df.columns else "Daily PU"

    chart_df["Direction"] = chart_df[chart_value].apply(
        lambda x: "Up" if pd.notna(x) and x > 0 else "Down" if pd.notna(x) and x < 0 else "Flat"
    )

    color_map = {"Up": "#15803d", "Down": "#b91c1c", "Flat": "#ca8a04"}

    fig = px.bar(
        chart_df,
        x=chart_value,
        y="Hotel",
        orientation="h",
        color="Direction",
        color_discrete_map=color_map,
        hover_data={
            "Latest D4cast": ":,.2f" if "Latest D4cast" in chart_df.columns else False,
            "Daily PU": ":,.2f" if "Daily PU" in chart_df.columns else False,
            "Daily PU %": ":.2f" if "Daily PU %" in chart_df.columns else False,
            "7D PU": ":,.2f" if "7D PU" in chart_df.columns else False,
            "MTD PU": ":,.2f" if "MTD PU" in chart_df.columns else False,
            "Direction": False,
        },
        title=f"{compare_metric} Leaderboard by {chart_value}",
    )

    fig.update_layout(
        plot_bgcolor="rgba(0,0,0,0)",
        paper_bgcolor="rgba(0,0,0,0)",
        height=max(420, 48 * len(chart_df)),
        yaxis=dict(categoryorder="total ascending"),
        xaxis=dict(showgrid=True, gridcolor="#f1f5f9"),
        legend_title_text="Direction",
        margin=dict(l=20, r=20, t=60, b=20),
    )

    fig.update_traces(texttemplate="%{x:,.2f}", textposition="outside", cliponaxis=False)

    st.plotly_chart(
        fig,
        use_container_width=True,
        config={"displayModeBar": True, "displaylogo": False},
    )

    return matrix


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


def metric_label_order():
    return ["Occ", "Room", "ADR", "Rev"]


def get_metric_options_with_all():
    return ["All Metrics"] + metric_label_order()


def compact_metric_table_height(df, row_height=36, max_height=520):
    if df is None or df.empty:
        return 160
    return min(max_height, 48 + row_height * len(df))


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


def normalize_stay_month_selection(selected_months):
    """
    Convert multiselect result to either:
    - "All" when all months should be included
    - list[str] when user selected specific stay months
    """
    if selected_months is None:
        return "All"

    if isinstance(selected_months, str):
        return "All" if selected_months == "All" else [selected_months]

    selected_months = list(selected_months)

    if not selected_months or "All" in selected_months:
        return "All"

    return selected_months


def stay_month_label(stay_month_selection):
    if stay_month_selection == "All":
        return "All"
    if isinstance(stay_month_selection, (list, tuple, set)):
        values = list(stay_month_selection)
        if len(values) <= 3:
            return ", ".join(values)
        return f"{len(values)} months selected"
    return str(stay_month_selection)


def apply_stay_month_filter(df, stay_month_selection):
    """
    Supports single month, multiple months, or All.
    """
    if df is None or df.empty:
        return df

    if stay_month_selection == "All":
        return df

    if isinstance(stay_month_selection, (list, tuple, set)):
        return df[df["Stay Month"].isin(list(stay_month_selection))].copy()

    return df[df["Stay Month"] == stay_month_selection].copy()


def format_compact_value(x, is_pct=False):
    if x is None or pd.isna(x):
        return ""
    return fmt_pct2(x) if is_pct else fmt_raw2(x)


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
                            {row['Stay Month']} · {row['Metric']} · {row['Status']}
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


def add_week_columns(df):
    if df is None or df.empty: return df
    out=df.copy(); dt=pd.to_datetime(out["Report Date"])
    out["Report Week"] = dt.dt.to_period("W-MON").astype(str)
    return out


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
        rows.append({"Hotel":hotel,"Stay Month":stay,"Metric":metric,"Report Week":week,"Start Date":g.iloc[0]["Report Date"],"End Date":g.iloc[-1]["Report Date"],"Week Start Forecast":sv,"Week End Forecast":ev,"Weekly PU":diff,"Weekly PU %":pct,"Status":"🟢 Up" if diff>0 else "🔴 Down" if diff<0 else "🟡 Flat"})
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
    with st.expander("📌 Metric definitions", expanded=False):
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


def friendly_week_label(start_date, end_date):
    start = pd.to_datetime(start_date)
    end = pd.to_datetime(end_date)
    if pd.isna(start) or pd.isna(end):
        return ""
    return f"{start.strftime('%d %b')} - {end.strftime('%d %b')}"


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
            "Status": "🟢 Up" if weekly_move > 0 else "🔴 Down" if weekly_move < 0 else "🟡 Flat",
        })

    out = pd.DataFrame(rows)
    if out.empty:
        return out

    out["Metric"] = pd.Categorical(out["Metric"], categories=METRIC_ORDER, ordered=True)
    out = out.sort_values(["Week", "Metric", "Weekly Movement"]).reset_index(drop=True)
    return out


def render_weekly_movement_v2(metric_data):
    st.markdown('<div class="section-title">Weekly Revenue Movement</div>', unsafe_allow_html=True)
    st.caption("Purpose: see which hotel moved up/down by week. Weekly Movement = End-of-week forecast − Start-of-week forecast.")

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

    role_map = {row["Role"]: row["Report Label"] for _, row in role_selection.iterrows() if pd.notna(row["Report Label"])}
    latest_label = role_map.get("Today / Latest")
    latest = metric_data[metric_data["Report Label"] == latest_label].copy()

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



def render_color_leaderboard(metric_long, role_selection, selected_hotels, stay_month_selection):
    st.markdown('<div class="section-title">Hotel Performance Sort Board</div>', unsafe_allow_html=True)
    st.caption("Revenue-focused board. Default is Budget Variance, because revenue team mainly needs to see above/below budget.")

    # Build base data from all selected metric_long, not only selected_metric.
    base_metric_data = metric_long[metric_long["Hotel"].isin(selected_hotels)].copy()
    base_metric_data = apply_stay_month_filter(base_metric_data, stay_month_selection)

    budget_df = build_budget_review(base_metric_data, role_selection)

    if budget_df.empty:
        st.info("No Budget data found for leaderboard.")
        return pd.DataFrame()

    c1, c2, c3 = st.columns([1, 1, 1])
    metric_choice = c1.selectbox("Metric", ["Rev", "Occ", "Room", "ADR", "All Metrics"], index=0, key="perf_metric")
    sort_choice = c2.selectbox(
        "Sort by",
        ["Budget Variance", "Budget Variance %", "Forecast", "Budget"],
        index=0,
        key="perf_sort",
        help="Pickup metrics were removed from the default board because budget variance is more useful for revenue brief.",
    )
    order = c3.selectbox("Order", ["Worst first", "Best first"], index=0, key="perf_order")

    view = budget_df.copy()
    if metric_choice != "All Metrics":
        view = view[view["Metric"] == metric_choice].copy()

    view = view.sort_values(sort_choice, ascending=(order == "Worst first")).reset_index(drop=True)

    total_forecast = view["Forecast"].sum()
    total_budget = view["Budget"].sum()
    total_var = total_forecast - total_budget

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Budget", fmt_raw2(total_budget))
    k2.metric("Forecast", fmt_raw2(total_forecast))
    k3.metric("Variance vs Budget", fmt_raw2(total_var), budget_delta_text(calc_budget_variance(total_forecast, total_budget)[1]))
    k4.metric("Below Budget Rows", int((view["Budget Variance"] < 0).sum()))

    raw = view.copy()
    show = view[["Hotel", "Stay Month", "Metric", "Budget", "Forecast", "Budget Variance", "Budget Variance %", "Status"]].copy()

    for col in ["Budget", "Forecast", "Budget Variance"]:
        show[col] = show[col].apply(lambda x: "" if pd.isna(x) else fmt_raw2(x))
    show["Budget Variance %"] = show["Budget Variance %"].apply(lambda x: "" if pd.isna(x) else fmt_pct2(x))

    def style_perf(_):
        styles = pd.DataFrame("", index=show.index, columns=show.columns)
        for idx, val in raw["Budget Variance"].items():
            color = "#bbf7d0" if val > 0 else "#fecaca" if val < 0 else "#fef08a"
            for col in ["Budget Variance", "Budget Variance %", "Status"]:
                if col in styles.columns:
                    styles.loc[idx, col] = f"background-color:{color}; font-weight:700"
        return styles

    st.dataframe(
        show.style.apply(style_perf, axis=None),
        use_container_width=True,
        hide_index=True,
        height=min(600, 48 + 36 * len(show)),
    )

    chart = raw.copy()
    chart["Direction"] = chart["Budget Variance"].apply(
        lambda x: "Above Budget" if x > 0 else "Below Budget" if x < 0 else "On Budget"
    )
    color_map = {"Above Budget": "#15803d", "Below Budget": "#b91c1c", "On Budget": "#ca8a04"}

    fig = px.bar(
        chart,
        x=sort_choice,
        y="Hotel",
        orientation="h",
        color="Direction",
        color_discrete_map=color_map,
        hover_data={"Forecast": ":,.2f", "Budget": ":,.2f", "Budget Variance %": ":.2f"},
        title=f"Budget Sort Board by {sort_choice} ({metric_choice})",
    )
    fig.update_layout(
        plot_bgcolor="rgba(0,0,0,0)",
        height=max(420, 46 * chart["Hotel"].nunique()),
        yaxis=dict(categoryorder="total ascending"),
        xaxis=dict(showgrid=True, gridcolor="#f1f5f9"),
        margin=dict(l=20, r=20, t=60, b=20),
    )
    fig.update_traces(texttemplate="%{x:,.2f}", textposition="outside", cliponaxis=False)
    st.plotly_chart(fig, use_container_width=True, config={"displaylogo": False}, key="hotel_sort_board_budget_chart")

    return view



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


def render_budget_first_kpi_section_v39(metric_data, role_selection, selected_metric):
    """
    Revenue-focused KPI cards:
    Budget vs Forecast.
    This replaces old vs Yesterday / vs 7D / vs 1st Month cards.
    """
    st.caption("Revenue focus: KPI cards compare Latest Forecast against Budget.")

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

        cols[0].metric("Latest Forecast", fmt_raw2(forecast))
        cols[1].metric("Budget", fmt_raw2(budget))

        delta_text = fmt_pct2(variance_pct) if variance_pct is not None else None

        cols[2].metric(
            "Variance vs Budget",












            
            fmt_raw2(variance),
            delta_text,
        )

        cols[3].metric(
            "Variance vs Budget %",
            fmt_pct2(variance_pct),
        )



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


def render_budget_first_kpi_section_v39(metric_data, role_selection, selected_metric):
    card_mode = st.radio(
        "KPI view",
        ["Executive cards", "Detailed metric cards"],
        horizontal=True,
        key="kpi_view_mode",
        help="Executive cards are recommended for revenue briefing.",
    )

    if card_mode == "Executive cards":
        render_executive_budget_cards(metric_data, role_selection)
        return

    st.caption("Detailed view: latest forecast compared with budget.")
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


def render_forecast_trend_by_month_v3(metric_data):
    st.markdown('<div class="section-title">Forecast Trend by Stay Month</div>', unsafe_allow_html=True)
    st.caption("Use this to see forecast trend by selected stay month. Point labels can be shown only on latest points to avoid clutter.")

    d4 = metric_data[metric_data["Reference"] == "Duetto"].copy()
    if d4.empty:
        st.info("No forecast trend data.")
        return pd.DataFrame()

    month_options = sorted(d4["Stay Month"].dropna().unique(), key=month_sort_key)

    c1, c2, c3 = st.columns([1, 1.5, 1])

    metric_choice = c1.selectbox(
        "Metric",
        ["Rev", "Occ", "Room", "ADR", "All Metrics"],
        index=0,
        key="trend_v3_metric",
    )

    trend_months = c2.multiselect(
        "Select Stay Months",
        options=month_options,
        default=month_options,
        key="trend_v3_month_filter",
        help="Select one or more stay months for this trend chart.",
    )

    label_mode = c3.selectbox(
        "Point labels",
        ["Latest point only", "All points", "Hide labels"],
        index=0,
        key="trend_v3_labels",
    )

    view = d4.copy()

    if metric_choice != "All Metrics":
        view = view[view["Metric"] == metric_choice].copy()

    view = view[view["Stay Month"].isin(trend_months)].copy()

    if view.empty:
        st.info("No trend data for selected filters.")
        return pd.DataFrame()

    trend = view.groupby(["Report Date", "Stay Month"], as_index=False)["Value"].sum().sort_values("Report Date")

    if label_mode == "All points":
        trend["Label"] = trend["Value"].apply(fmt_raw2)
    elif label_mode == "Latest point only":
        latest_date = trend["Report Date"].max()
        trend["Label"] = trend.apply(lambda r: fmt_raw2(r["Value"]) if r["Report Date"] == latest_date else "", axis=1)
    else:
        trend["Label"] = ""

    fig = px.line(
        trend,
        x="Report Date",
        y="Value",
        color="Stay Month",
        markers=True,
        text="Label",
        title=f"Forecast Trend by Stay Month ({metric_choice})",
    )
    fig.update_traces(textposition="top center")
    fig.update_layout(
        plot_bgcolor="rgba(0,0,0,0)",
        xaxis=dict(title="Report Date", tickformat="%d %b", showgrid=True, gridcolor="#f1f5f9"),
        yaxis=dict(title="Forecast", showgrid=True, gridcolor="#f1f5f9"),
        legend_title_text="Stay Month",
        margin=dict(l=20, r=20, t=60, b=20),
    )
    st.plotly_chart(fig, use_container_width=True, config={"displaylogo": False}, key="forecast_trend_v3")

    with st.expander("Trend data"):
        show = trend.copy()
        show["Value"] = show["Value"].apply(fmt_raw2)
        st.dataframe(show, use_container_width=True, hide_index=True, height=min(520, 48 + 34 * len(show)))

    return trend


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
                status = "⚪ No Base"
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
                    status = "⚪ No Base"
                else:
                    movement = latest_value - base_value
                    movement_pct = movement / base_value * 100
                    status = "🟢 Up" if movement > 0 else "🔴 Down" if movement < 0 else "🟡 Flat"

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
                    <div style="font-weight:800; font-size:1.02rem;">{row['Hotel']} · {row['Metric']} · {row['Stay Month']} · {row['Period']}</div>
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
    st.caption("Main budget page: Budget vs Forecast, priority cards, full table, and optional chart. This replaces the old Budget Review tab.")

    base_metric_data = metric_long[metric_long["Hotel"].isin(selected_hotels)].copy()
    base_metric_data = apply_stay_month_filter(base_metric_data, stay_month_selection)

    budget_df = build_budget_review(base_metric_data, role_selection)

    if budget_df.empty:
        st.info("No Budget data found for selected filters.")
        return pd.DataFrame()

    c1, c2, c3, c4 = st.columns([1, 1, 1, 1])

    metric_choice = c1.selectbox(
        "Metric",
        ["Rev", "Occ", "Room", "ADR", "All Metrics"],
        index=0,
        key="sort_v32_metric",
    )

    view_level = c2.selectbox(
        "Level",
        ["Summary by Hotel", "Detail by Month"],
        index=0,
        key="sort_v32_level",
        help="Use Summary by Hotel for All Month view. Use Detail by Month only when you need monthly breakdown.",
    )

    sort_choice = c3.selectbox(
        "Sort by",
        ["Budget Variance", "Budget Variance %", "Forecast", "Budget"],
        index=0,
        key="sort_v32_sort",
    )

    order = c4.selectbox(
        "Order",
        ["Worst first", "Best first"],
        index=0,
        key="sort_v32_order",
    )

    view = build_budget_review_summary_view(budget_df, view_level)

    if metric_choice != "All Metrics":
        view = view[view["Metric"] == metric_choice].copy()

    if view.empty:
        st.info("No rows after selected filters.")
        return view

    view = view.sort_values(sort_choice, ascending=(order == "Worst first")).reset_index(drop=True)

    total_forecast = view["Forecast"].sum()
    total_budget = view["Budget"].sum()
    total_variance = total_forecast - total_budget
    below_rows = int((view["Budget Variance"] < 0).sum())
    above_rows = int((view["Budget Variance"] > 0).sum())

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Budget", fmt_raw2(total_budget))
    k2.metric("Forecast", fmt_raw2(total_forecast))
    k3.metric("Variance vs Budget", fmt_raw2(total_variance), budget_delta_text(calc_budget_variance(total_forecast, total_budget)[1]))
    k4.metric("Below Budget Rows", below_rows)

    priority_budget_view = render_priority_budget_table(view, default_rows="Worst 10")

    with st.expander("Optional chart: budget variance"):
        st.caption("Chart is optional because All Month view is easier to read as priority cards/table.")

        chart_limit = st.selectbox(
            "Chart rows",
            ["Worst 5", "Worst 10", "Best 5", "Best 10"],
            index=0,
            key="sort_v33_chart_rows",
        )

        chart = view.copy()
        if "Worst" in chart_limit:
            n = int(chart_limit.replace("Worst ", ""))
            chart = chart.sort_values("Budget Variance", ascending=True).head(n)
        else:
            n = int(chart_limit.replace("Best ", ""))
            chart = chart.sort_values("Budget Variance", ascending=False).head(n)

        chart["Direction"] = chart["Budget Variance"].apply(
            lambda x: "Above Budget" if x > 0 else "Below Budget" if x < 0 else "On Budget"
        )
        chart["Label Name"] = chart["Hotel"].astype(str)
        if view_level == "Detail by Month" and "Stay Month" in chart.columns:
            chart["Label Name"] = chart["Hotel"].astype(str) + " | " + chart["Stay Month"].astype(str)

        color_map = {"Above Budget": "#15803d", "Below Budget": "#b91c1c", "On Budget": "#ca8a04"}

        fig = px.bar(
            chart,
            x="Budget Variance",
            y="Label Name",
            orientation="h",
            color="Direction",
            color_discrete_map=color_map,
            hover_data={
                "Forecast": ":,.2f",
                "Budget": ":,.2f",
                "Budget Variance %": ":.2f",
                "Metric": True,
                "Direction": False,
            },
            title=f"Top Priority Budget Variance ({metric_choice}, {view_level})",
        )
        fig.update_layout(
            plot_bgcolor="rgba(0,0,0,0)",
            height=max(360, 48 * len(chart)),
            yaxis=dict(categoryorder="total ascending", title=""),
            xaxis=dict(showgrid=True, gridcolor="#f1f5f9"),
            margin=dict(l=20, r=20, t=60, b=20),
            showlegend=True,
        )
        fig.update_traces(
            texttemplate="%{x:,.0f}",
            textposition="outside",
            cliponaxis=False,
        )
        st.plotly_chart(fig, use_container_width=True, config={"displaylogo": False}, key="sort_v33_optional_chart")

    with st.expander("Full budget detail table", expanded=False):
        show = view.copy()

        raw = view.copy()

        for col in ["Budget", "Forecast", "Today OTB", "Budget Variance"]:
            show[col] = show[col].apply(lambda x: "" if pd.isna(x) else fmt_raw2(x))
        for col in ["Budget Variance %", "OTB vs Budget %", "OTB vs Forecast %"]:
            if col in show.columns:






                show[col] = show[col].apply(lambda x: "" if pd.isna(x) else fmt_signed_pct2(x))

        def style_budget_table(_):
            styles = pd.DataFrame("", index=show.index, columns=show.columns)
            for idx, val in raw["Budget Variance"].items():
                color = "#bbf7d0" if val > 0 else "#fecaca" if val < 0 else "#fef08a"
                for c in ["Budget Variance", "Budget Variance %", "Status"]:
                    if c in styles.columns:
                        styles.loc[idx, c] = f"background-color:{color}; font-weight:700"
            for value_col in ["OTB vs Budget %", "OTB vs Forecast %"]:
                if value_col not in raw.columns:
                    continue
                for idx, val in raw[value_col].items():
                    color = "#bbf7d0" if pd.notna(val) and val > 0 else "#fecaca" if pd.notna(val) and val < 0 else "#fef08a"

                    styles.loc[idx, value_col] = f"background-color:{color}; font-weight:700"
            return styles

        st.dataframe(
            show.style.apply(style_budget_table, axis=None),
            use_container_width=True,
            hide_index=True,
            height=min(620, 48 + 36 * len(show)),
        )

    return view



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



def render_forecast_movement_table_only(metric_data, role_selection):
    """
    Presentation-friendly Forecast Movement page.

    Design decision:
    - Table only, no cards and no bar chart.
    - Default metric = All Metrics so Revenue Team can present all key metrics together.
    - Movement periods = 1 Day / 7 Days / First Day of Month.
    """
    st.markdown('<div class="section-title">Forecast Movement</div>', unsafe_allow_html=True)
    st.caption("Table-only view for presentation. Shows how latest forecast moved versus 1 Day, 7 Days, and First Day of Month.")

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

    c1, c2, c3 = st.columns([1, 1, 1])

    metric_filter = c1.selectbox(
        "Metric",
        ["All Metrics", "Rev", "Occ", "Room", "ADR"],
        index=0,
        key="movement_table_only_metric",
        help="Default is All Metrics so the team can present all key metrics together.",
    )

    period_filter = c2.selectbox(
        "Compare with",
        ["All Periods", "1 Day", "7 Days", "First Day of Month"],
        index=0,
        key="movement_table_only_period",
    )

    sort_mode = c3.selectbox(
        "Sort",


        ["Worst movement % first", "Best movement % first", "Hotel order"],
        index=0,
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

    # Sort for presentation
    if sort_mode == "Worst movement % first" and "Movement %" in view.columns:
        view = view.sort_values(["Movement %", "Hotel", "Metric"], ascending=[True, True, True]).reset_index(drop=True)
    elif sort_mode == "Best movement % first" and "Movement %" in view.columns:


        view = view.sort_values(["Movement %", "Hotel", "Metric"], ascending=[False, True, True]).reset_index(drop=True)
    else:
        view = view.sort_values(["Hotel", "Stay Month", "Metric", "Period"]).reset_index(drop=True)

    # Top KPI row
    total_move = view["Movement"].sum() if "Movement" in view.columns else 0
    up_rows = int((view["Movement"] > 0).sum()) if "Movement" in view.columns else 0
    down_rows = int((view["Movement"] < 0).sum()) if "Movement" in view.columns else 0

    k1, k2, k3 = st.columns(3)
    k1.metric("Total Movement", fmt_raw2(total_move))
    k2.metric("Rows Up", up_rows)
    k3.metric("Rows Down", down_rows)

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

    with st.expander("How to read Forecast Movement"):
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
    st.caption("Budget is the target/base. Variance % = (Latest Forecast - Budget) / Budget × 100.")

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


def render_budget_first_kpi_section_v39(metric_data, role_selection, selected_metric):
    """
    Keep layout simple:
    - Executive cards for Rev summary
    - Detailed cards optional
    """
    kpi_mode = st.radio(
        "KPI view",
        ["Executive Rev", "Detailed by Metric"],
        horizontal=True,
        index=0,
        key="budget_first_kpi_mode",
    )

    if kpi_mode == "Executive Rev":
        render_budget_first_executive_cards(metric_data, role_selection)
    else:
        render_budget_first_kpi_cards(metric_data, role_selection, selected_metric)



def calc_budget_variance(forecast, budget):
    """
    Single source of truth for Budget vs Forecast.

    Budget is the target/base.
    Forecast is the latest expected result.

    Variance = Forecast - Budget
    Variance % = Variance / Budget * 100
    """
    if pd.isna(forecast):
        forecast = 0
    if pd.isna(budget):
        budget = 0

    variance = forecast - budget

    if budget is None or pd.isna(budget) or budget == 0:
        variance_pct = None
    else:
        variance_pct = variance / budget * 100

    return variance, variance_pct


def budget_status_from_variance(variance):
    if variance is None or pd.isna(variance):
        return "⚪ No Budget"
    if variance > 0:
        return "🟢 Above Budget"
    if variance < 0:
        return "🔴 Below Budget"
    return "🟡 On Budget"


def budget_delta_text(variance_pct):
    """
    Streamlit metric delta text.
    Use already-calculated budget variance %, not safe_delta().
    """
    if variance_pct is None or pd.isna(variance_pct):
        return None
    return fmt_pct2(variance_pct)


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
    "On The Book": "Today OTB",
    "Budget": "Budget",
    "Forecast": "Forecast",
}
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


def render_budget_first_kpi_section_v39(metric_data, role_selection, selected_metric):
    """
    KPI section — compare row (chip checkboxes) + view toggle + metric cards.
    Layout: [Compare chips] spacer [View toggle]
    """
    # ── Top control row ───────────────────────────────────────
    ctrl_left, ctrl_right = st.columns([3, 2])

    with ctrl_left:
        st.markdown(
            '<p style="font-size:0.68rem;font-weight:700;text-transform:uppercase;'
            'letter-spacing:0.09em;color:#8c8c8c;margin:0 0 4px 0;">Compare</p>',
            unsafe_allow_html=True,
        )
        c_otb, _s1, c_bgt, _s2, c_fct = st.columns([1.4, 0.12, 1.1, 0.12, 1.2])
        otb_on = c_otb.checkbox("On The Book", value=False, key="kpi_chk_otb")
        _s1.markdown('<p class="compare-sep">│</p>', unsafe_allow_html=True)
        bgt_on = c_bgt.checkbox("Budget", value=True, key="kpi_chk_budget")
        _s2.markdown('<p class="compare-sep">│</p>', unsafe_allow_html=True)
        fct_on = c_fct.checkbox("Forecast", value=True, key="kpi_chk_forecast")

    with ctrl_right:
        kpi_mode = st.radio(
            "View",
            ["Revenue summary", "By metric"],
            horizontal=True,
            index=0,
            key="budget_first_kpi_mode_v39",
        )

    selected_axes = (
        (["On The Book"] if otb_on else []) +
        (["Budget"]     if bgt_on else []) +
        (["Forecast"]   if fct_on else [])
    )

    if len(selected_axes) < 2:
        st.caption("Check at least 2 items to compare.")
        return

    budget_df = build_budget_review(metric_data, role_selection)
    if budget_df.empty:
        st.info("No budget data available for the current filters.")
        return

    # ── Cards ─────────────────────────────────────────────────
    if kpi_mode == "Revenue summary":
        rev_df = budget_df[budget_df["Metric"] == "Rev"].copy()
        render_kpi_axis_cards(rev_df if not rev_df.empty else budget_df, selected_axes)
    else:
        metrics_to_show = metric_label_order() if selected_metric == "All Metrics" else [selected_metric]
        for metric_name in metrics_to_show:
            st.markdown(
                f'<p style="font-size:0.78rem;font-weight:600;color:#595959;'
                f'margin:10px 0 4px 0;">{metric_name}</p>',
                unsafe_allow_html=True,
            )
            metric_df = budget_df[budget_df["Metric"] == metric_name].copy()
            if metric_df.empty:
                st.caption(f"No data for {metric_name}.")
                continue
            render_kpi_axis_cards(metric_df, selected_axes)



def render_priority_budget_table(view, default_rows="Worst 10"):
    """
    Presentation-friendly priority budget table.
    Replaces long priority cards with a compact table.
    """
    st.markdown("#### Priority budget table")

    c1, c2 = st.columns([1, 1])

    show_mode = c1.selectbox(
        "Show rows",
        ["Worst 5", "Worst 10", "Worst 20", "All"],
        index=["Worst 5", "Worst 10", "Worst 20", "All"].index(default_rows) if default_rows in ["Worst 5", "Worst 10", "Worst 20", "All"] else 1,
        key="priority_budget_table_rows",
    )

    table_sort = c2.selectbox(
        "Table sort",
        ["Worst variance first", "Best variance first", "Hotel order"],
        index=0,
        key="priority_budget_table_sort",
    )

    table_df = view.copy()

    if table_sort == "Worst variance first":
        table_df = table_df.sort_values("Budget Variance", ascending=True).reset_index(drop=True)
    elif table_sort == "Best variance first":
        table_df = table_df.sort_values("Budget Variance", ascending=False).reset_index(drop=True)
    else:
        sort_cols = [c for c in ["Hotel", "Stay Month", "Metric"] if c in table_df.columns]
        table_df = table_df.sort_values(sort_cols).reset_index(drop=True) if sort_cols else table_df

    if show_mode != "All":
        n = int(show_mode.replace("Worst ", ""))
        table_df = table_df.head(n)

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

    st.dataframe(
        show.style.apply(style_priority_budget, axis=None),
        use_container_width=True,
        hide_index=True,
        height=min(540, 48 + 38 * len(show)),
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

# ── Page title — use native st.markdown h2, never custom <p>
# ── (custom HTML <p> ชนขอบ Streamlit toolbar บน Cloud)
# ─────────────────────────────────────────────────────────────
st.markdown("## Revenue Briefing")
st.caption("G5 Hotels · D4cast daily forecast review")

# ── Sidebar ───────────────────────────────────────────────────
# RULE: ห้าม st.stop() ใน main area ก่อน with st.sidebar: ทำงาน
# เพราะ stop ใน main area = sidebar ไม่ render = widget หาย
# วิธีที่ถูก: ให้ sidebar รันเสร็จก่อน แล้วค่อย check ใน main area
# ─────────────────────────────────────────────────────────────

# Default values — จะถูก overwrite ใน sidebar ถ้าข้อมูลโหลดได้
file_catalog = None
mode = "Upload"

def hotel_key(hotel_name):
    safe = re.sub(r"[^A-Za-z0-9_]+", "_", str(hotel_name))
    return f"hotel_checkbox_{safe}"

with st.sidebar:
    st.markdown("## Data source")
    mode = st.radio(
        "source_mode",
        ["Folder", "Upload"],
        horizontal=True,
        label_visibility="collapsed",
        key="data_source_mode",
    )

    # ── Load data ──────────────────────────────────────────
    if mode == "Folder":
        folder_path = st.text_input(
            "Path",
            value=r"G:\My Drive\Ecom\Report\G5 - Weekly Pace Review",
        )
        if st.button("Refresh", use_container_width=True, type="primary"):
            st.cache_data.clear()
            st.rerun()
        try:
            with st.spinner("Loading…"):
                file_catalog = build_file_catalog_from_folder(folder_path)
            st.caption(f"{len(file_catalog)} files found")
        except Exception as e:
            st.error(str(e))
            # file_catalog ยัง None — จะ stop ใน main area ด้านล่าง
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
        else:
            st.caption("Drop files above to begin.")
            # file_catalog ยัง None — จะ stop ใน main area ด้านล่าง

    # ── Filters (แสดงเฉพาะเมื่อข้อมูลพร้อม) ──────────────
    if file_catalog is not None:
        st.divider()
        st.markdown("## Filters")

        report_file_months = sorted(file_catalog["Report Date"].dt.strftime("%b, %Y").unique())
        latest_report_month = file_catalog["Report Date"].max().strftime("%b, %Y")
        report_file_month = st.selectbox(
            "Report month",
            report_file_months,
            index=report_file_months.index(latest_report_month),
        )

        role_selection, month_file_catalog = select_role_files(file_catalog, report_file_month)
        selected_file_catalog = month_file_catalog.sort_values("Report Date").reset_index(drop=True)
        selected_file_catalog["Report Order"] = range(1, len(selected_file_catalog) + 1)

        with st.spinner("Processing…"):
            combined_df = pd.concat(
                [parse_record(row) for _, row in selected_file_catalog.iterrows()],
                ignore_index=True,
            )
            ref_col_map = build_ref_col_map(combined_df)
            if not ref_col_map.get("Duetto"):
                st.error("No Forecast / Duetto columns found.")
                file_catalog = None  # reset → main area จะ stop
            else:
                metric_long = build_metric_long(combined_df, ref_col_map)

    if file_catalog is not None:
        all_hotels = sorted(metric_long["Hotel"].dropna().unique())
        all_stay_months = sorted(metric_long["Stay Month"].dropna().unique(), key=month_sort_key)
        default_report_month = (
            report_file_month if report_file_month in all_stay_months
            else (all_stay_months[0] if all_stay_months else None)
        )

        # Stay month
        st.markdown("#### Stay month")
        stay_month_mode = st.selectbox(
            "stay_month_mode",
            ["Report month only", "All months", "Custom"],
            index=0,
            label_visibility="collapsed",
        )

        if stay_month_mode == "Report month only":
            selected_stay_months_raw = [default_report_month] if default_report_month else []
            st.caption(default_report_month or "—")
        elif stay_month_mode == "All months":
            selected_stay_months_raw = all_stay_months
            stay_month_selection = "All"
            st.caption(f"All ({len(all_stay_months)} months)")
        else:
            custom_default = [default_report_month] if default_report_month else all_stay_months[:1]
            selected_stay_months_raw = st.multiselect(
                "Pick months",
                options=all_stay_months,
                default=custom_default,
                label_visibility="collapsed",
            )
            if not selected_stay_months_raw:
                st.warning("Select at least one month.")
                selected_stay_months_raw = custom_default
            st.caption(stay_month_label(selected_stay_months_raw))

        if stay_month_mode != "All months":
            stay_month_selection = normalize_stay_month_selection(selected_stay_months_raw)

        # Hotels
        st.markdown("#### Hotels")
        ca, cb = st.columns(2)
        if ca.button("All", use_container_width=True, key="hotel_select_all_btn"):
            for h in all_hotels:
                st.session_state[hotel_key(h)] = True
            st.rerun()
        if cb.button("None", use_container_width=True, key="hotel_clear_all_btn"):
            for h in all_hotels:
                st.session_state[hotel_key(h)] = False
            st.rerun()

        selected_hotels = []
        with st.expander(f"Properties ({len(all_hotels)})", expanded=True):
            for hotel in all_hotels:
                key = hotel_key(hotel)
                if key not in st.session_state:
                    st.session_state[key] = True
                if st.checkbox(str(hotel), key=key):
                    selected_hotels.append(hotel)
        st.caption(f"{len(selected_hotels)} / {len(all_hotels)} selected")

        # Metric
        st.markdown("#### Metric")
        selected_metric = st.selectbox(
            "metric_sel",
            get_metric_options_with_all(),
            index=0,
            label_visibility="collapsed",
        )

# ── Main area: guard — no data loaded ────────────────────────
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
    st.stop()

# ── Guard: no hotels selected ────────────────────────────────
if not selected_hotels:
    st.warning("Select at least one property in the sidebar.")
    st.stop()

# ── Filter data ───────────────────────────────────────────────
metric_data = metric_long[metric_long["Hotel"].isin(selected_hotels)].copy()
if selected_metric != "All Metrics":
    metric_data = metric_data[metric_data["Metric"] == selected_metric].copy()
metric_data = apply_stay_month_filter(metric_data, stay_month_selection)

if metric_data.empty:
    st.warning("No data for the current filter selection.")
    st.stop()

# ── Pre-build summaries ───────────────────────────────────────
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

# ── Filter context bar ────────────────────────────────────────
st.markdown(
    f'<div class="filter-bar">'
    f'<span class="filter-sep">Metric</span> <b>{selected_metric}</b>'
    f'<span class="filter-sep"> · </span>'
    f'<span class="filter-sep">Stay</span> <b>{stay_month_label(stay_month_selection)}</b>'
    f'<span class="filter-sep"> · </span>'
    f'<span class="filter-sep">Properties</span> <b>{len(selected_hotels)}</b>'
    f'</div>',
    unsafe_allow_html=True,
)

# ── KPI section ───────────────────────────────────────────────
render_budget_first_kpi_section_v39(metric_data, role_selection, selected_metric)

# ── Color legend ──────────────────────────────────────────────
st.markdown(
    '<div class="legend-row">'
    '<span><span class="legend-dot" style="background:#52c41a;"></span>Above / Up</span>'
    '<span><span class="legend-dot" style="background:#ff4d4f;"></span>Below / Down</span>'
    '<span><span class="legend-dot" style="background:#faad14;"></span>Flat / Neutral</span>'
    '<span><span class="legend-dot" style="background:#1677ff;"></span>Forecast / Info</span>'
    '</div>',
    unsafe_allow_html=True,
)
render_metric_dictionary()

# ── Main tabs ─────────────────────────────────────────────────
tab0, tab_movement, tab_leaderboard, tab1, tab_analysis, tab5 = st.tabs([
    "Forecast Pivot",
    "Movement",
    "Budget Board",
    "Trend",
    "Advanced",
    "Export",
])


with tab0:
    latest_pivot = build_latest_pivot_table(metric_data, role_selection)
    render_compact_hotel_tabs(latest_pivot)




with tab_movement:
    forecast_movement_summary = render_forecast_movement_table_only(metric_data, role_selection)


with tab_leaderboard:
    leaderboard_summary = render_budget_sort_board_v32(
        metric_long=metric_long,
        role_selection=role_selection,
        selected_hotels=selected_hotels,
        stay_month_selection=stay_month_selection,
    )



with tab1:
    trend_summary = render_forecast_trend_by_month_v3(metric_data)

    st.markdown('<div class="section-title">Hotel Momentum</div>', unsafe_allow_html=True)

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

        latest_date_bubble = bubble["Report Date"].max()
        latest_by_hotel = (
            bubble[bubble["Report Date"] == latest_date_bubble]
            .groupby("Hotel", as_index=False)["Value"]
            .sum()
            .sort_values("Value", ascending=False)
        )

        # Latest snapshot per hotel, used for ranking/filtering the bubble chart.
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
            color="Daily Change",
            hover_data={
                "Report Label": True,
                "Latest D4cast": ":,.2f",
                "Previous Forecast": ":,.2f",
                "Daily PU": ":,.2f",
                "Daily PU %": ":.2f",
                "Bubble Size": False,
            },
            title=f"Hotel-level Forecast Momentum / Daily PU ({selected_metric})",
            color_continuous_scale=["#b91c1c", "#facc15", "#15803d"],
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
            coloraxis_colorbar=dict(title="Daily PU"),
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

        st.markdown("#### Hotel momentum summary")

        latest_rows = (
            bubble.sort_values(["Hotel", "Report Date"])
            .groupby(["Hotel", "Stay Month"], as_index=False)
            .tail(1)
            .copy()
        )

        latest_rows["Status"] = latest_rows["Daily Change"].apply(
            lambda x: "🟢 Up" if pd.notna(x) and x > 0 else "🔴 Down" if pd.notna(x) and x < 0 else "🟡 Flat"
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



with tab_analysis:
    st.markdown('<div class="section-title">Advanced Analysis</div>', unsafe_allow_html=True)

    # -------------------------
    # Recommended Pace
    # -------------------------
    st.markdown("### 1) Same-Time Pace Benchmark")
    st.caption("Compares today/latest forecast against the best same-time benchmark among STLY / ST2Y / ST3Y. Variance % is colored for presentation clarity.")

    if pace_summary.empty:
        st.info("No pace data.")
    else:
        view_mode_pace = st.selectbox(
            "Recommended Pace view",
            ["Hotel tabs", "List view"],
            index=0,
            key="analysis_pace_view",
        )
        pace_cols = {
            "Hotel": st.column_config.TextColumn("Hotel", width="medium"),
            "Today": st.column_config.NumberColumn("Today", format="%,.2f"),
            "STLY": st.column_config.NumberColumn("STLY", format="%,.2f"),
            "ST2Y": st.column_config.NumberColumn("ST2Y", format="%,.2f"),
            "ST3Y": st.column_config.NumberColumn("ST3Y", format="%,.2f"),
            "Recommended Pace Value": st.column_config.NumberColumn("Rec. Pace Value", format="%,.2f"),
            "Pace Diff": st.column_config.NumberColumn("Variance", format="%,.2f"),
            "Pace Diff %": st.column_config.NumberColumn("Variance %", format="%.2f%%"),
        }
        pace_display = make_recommended_pace_compact(pace_summary)

        pace_layout = st.radio(
                "Layout",
                ["Compact table", "Cards"],
                horizontal=True,
                key="pace_compact_layout",
            )

        if pace_layout == "Cards":
            render_pace_cards(pace_summary)
        else:
            render_compact_by_hotel(pace_display, view_mode_pace, "pace")

    st.divider()

    # -------------------------
    # D4cast vs Final
    # -------------------------
    st.markdown("### 2) Historical Final Comparison")
    st.caption("Forecast compared with Historical Final values: Final LY / Final 2Y / Final 3Y. Use this as historical context, not budget decision.")

    if final_comparison.empty:
        st.info("No final comparison data.")
    else:
        view_mode_final = st.selectbox(
            "D4cast vs Final view",
            ["Hotel tabs", "List view"],
            index=0,
            key="analysis_final_view",
        )
        final_cols = {
            "Hotel": st.column_config.TextColumn("Hotel", width="medium"),
            "Forecast": st.column_config.NumberColumn("Forecast", format="%,.2f"),
            "Final Value": st.column_config.NumberColumn("Final Value", format="%,.2f"),
            "Diff": st.column_config.NumberColumn("Variance", format="%,.2f"),
            "Diff %": st.column_config.NumberColumn("Variance %", format="%.2f%%"),
        }
        final_display = make_final_compact(final_comparison)
        render_compact_by_hotel(final_display, view_mode_final, "final")



with tab5:
    st.markdown('<div class="section-title">Export & Settings</div>', unsafe_allow_html=True)
    st.markdown("""
    - ✅ Use `python -m streamlit run <file>.py` for local presentation.
    - ✅ Make sure the Google Drive / local folder path points to the daily G5 report folder.
    - ✅ Confirm Today / Yesterday / 7D / 1st Month roles below before presenting.
    """)
    st.markdown('<div class="section-title">Report Roles Validation</div>', unsafe_allow_html=True)
    st.dataframe(role_selection, use_container_width=True, hide_index=True)
    
    st.markdown('<div class="section-title">📥 Export Data</div>', unsafe_allow_html=True)
    
    def trigger_download_toast():
        st.toast("✅ File downloaded successfully!", icon="🎉")

    c1, c2 = st.columns(2)
    with c1:
        sheets = {
            "Role Selection": role_selection,
            "Budget Sort Board": build_budget_review(metric_data, role_selection),
            "D4cast Momentum": momentum_summary,
            "Forecast Movement": build_forecast_movement_v31(metric_data, role_selection),
            "Movement Table": movement_summary,
            "Recommended Pace": pace_summary,
            "D4cast vs Final": final_comparison,
        }
        st.download_button(
            "📊 Download Full Excel Report",
            data=to_excel_bytes(sheets),
            file_name=f"g5_d4cast_{report_file_month.replace(', ', '_')}_{selected_metric}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            on_click=trigger_download_toast,
            use_container_width=True
        )
    with c2:
        st.download_button(
            "📝 Download Movement CSV",
            data=movement_summary.to_csv(index=False).encode("utf-8"),
            file_name="g5_d4cast_movement.csv",
            mime="text/csv",
            on_click=trigger_download_toast,
            use_container_width=True
        )
