"""Type-safe helper utilities used across the project.

These functions are pure and have no side effects — safe to import anywhere.
"""
from __future__ import annotations

import re
from datetime import datetime
from pathlib import Path
from typing import Any

import pandas as pd

from src.core.constants import MONTH_FORMATS_TRY


def clean_text(x: Any) -> str:
    """Strip whitespace and surrounding quotes from a value.

    Returns empty string for None.
    """
    if x is None:
        return ""
    return str(x).strip().strip('"').strip()


def normalize_stay_month(raw: Any) -> str | None:
    """Normalize a stay-month string to canonical 'Mon, YYYY' format.

    Accepts many input formats (May-26, May 2026, May, 26, ...).
    Returns None for unparseable input.
    """
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
            if yr < 100:
                yr += 2000
            try:
                dt = datetime.strptime(f"{month_part} 1 {yr}", "%b %d %Y")
            except ValueError:
                dt = datetime.strptime(f"{month_part} 1 {yr}", "%B %d %Y")
            return dt.strftime("%b, %Y")
        except ValueError:
            return None
    return None


def month_sort_key(x: Any) -> pd.Timestamp:
    """Return a sortable Timestamp from a 'Mon, YYYY' string. NaT on failure."""
    return pd.to_datetime(x, format="%b, %Y", errors="coerce")


def extract_date_from_filename(file_name: str) -> pd.Timestamp:
    """Extract a date (Timestamp at midnight) from a filename like
    'report_2026-06-30.csv' or '20260630_report.xlsx'. NaT if no date found.
    """
    s = str(file_name)
    patterns = [r"\d{4}-\d{2}-\d{2}", r"\d{4}_\d{2}_\d{2}", r"\d{8}", r"\d{1,2}-\d{1,2}-\d{4}"]
    for pattern in patterns:
        m = re.search(pattern, s)
        if not m:
            continue
        date_text = m.group(0)
        if re.fullmatch(r"\d{8}", date_text):
            date_text = f"{date_text[:4]}-{date_text[4:6]}-{date_text[6:]}"
        date_text = date_text.replace("_", "-")
        dt = pd.to_datetime(date_text, errors="coerce", dayfirst=False)
        if pd.notna(dt):
            return dt.normalize()
    return pd.NaT


def file_modified_date(path: str | Path) -> pd.Timestamp:
    """Return file's modified date as a Timestamp at midnight."""
    return pd.to_datetime(Path(path).stat().st_mtime, unit="s").normalize()


def trunc2(x: Any) -> float | None:
    """Truncate (toward zero) to 2 decimal places. Returns None for NaN/None."""
    if x is None or pd.isna(x):
        return None
    # Truncate toward zero, not round.
    return int(float(x) * 100) / 100


def fmt_raw2(x: Any) -> str:
    """Format as a thousands-separated number with 2 decimals.

    Returns empty string for None.
    Uses truncation (not rounding) — '1.239' → '1.23'.
    """
    x = trunc2(x)
    if x is None:
        return ""
    return f"{x:,.2f}"


def fmt_pct2(x: Any) -> str:
    """Format as a percentage with 2 decimals and % suffix. Empty for None."""
    x = trunc2(x)
    if x is None:
        return ""
    return f"{x:,.2f}%"


def fmt_signed_pct2(x: Any) -> str:
    """Format as a signed (+/-) percentage with 2 decimals. Empty for None."""
    x = trunc2(x)
    if x is None:
        return ""
    return f"{x:+,.2f}%"


def safe_delta(current: float | None, base: float | None) -> str | None:
    """Return a formatted percentage delta string, or None if base is invalid."""
    if base is None or pd.isna(base) or base == 0:
        return None
    return fmt_pct2((current - base) / base * 100)
