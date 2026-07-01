from __future__ import annotations

import io
import tempfile
import zipfile
from pathlib import Path

import pandas as pd
import streamlit as st

from src.core.constants import (
    METRIC_PATTERNS,
    METRIC_TO_DISPLAY,
    REFERENCE_PATTERNS,
    SUPPORTED_EXTENSIONS,
)
from src.core.helpers import clean_text, extract_date_from_filename, file_modified_date, normalize_stay_month


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
    """
    Pick 4 role files from the catalog:
      - Today / Latest       : newest file in the report month
      - Yesterday / Previous : the file just before Today's date (cross-month OK)
      - Last 7D              : file closest to Today - 7 days (cross-month OK)
      - 1st Month            : first file within the report month

    Yesterday / Last 7D search the FULL catalog (not just the current month)
    so that early-month reports still have a valid base for movement calc
    (e.g. Jul 1's yesterday = Jun 30).
    """
    start = pd.to_datetime(report_file_month, format="%b, %Y")
    end = start + pd.offsets.MonthEnd(0)
    month_files = file_catalog[(file_catalog["Report Date"] >= start) & (file_catalog["Report Date"] <= end)].copy()

    if month_files.empty:
        raise ValueError(f"No report files found for {report_file_month}")

    latest = month_files.sort_values("Report Date").iloc[-1]
    latest_date = latest["Report Date"]

    # Yesterday / 7D can come from PREVIOUS months too (cross-month lookback)
    all_previous = file_catalog[file_catalog["Report Date"] < latest_date].copy()

    today = latest
    yesterday = (
        all_previous.sort_values("Report Date").iloc[-1]
        if not all_previous.empty else None
    )

    target_7d = latest_date - pd.Timedelta(days=7)
    seven = None
    if not all_previous.empty:
        temp = all_previous.copy()
        temp["Distance To 7D"] = (temp["Report Date"] - target_7d).abs()
        seven = temp.sort_values(["Distance To 7D", "Report Date"]).iloc[0]

    # 1st Month stays scoped to the current report month
    first = month_files.sort_values("Report Date").iloc[0]
    rows = []

    def add(role, row):
        if row is None:
            rows.append({"Role": role, "Report Label": None, "Report Date": None, "File Name": None, "Status": "Missing"})
        else:
            rows.append({"Role": role, "Report Label": row["Report Label"], "Report Date": row["Report Date"], "File Name": row["File Name"], "Status": "OK"})

    add("Today / Latest", today)
    add("Yesterday / Previous", yesterday)
    add("Last 7D", seven)
    add("1st Month", first)

    return pd.DataFrame(rows), month_files.copy()


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
