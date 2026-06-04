"""Centralised constants.

Schema-related constants (REFERENCE_PATTERNS, METRIC_PATTERNS, …) are
loaded from `config/duetto_schema.yaml` if it exists, otherwise the
in-code defaults below are used. This lets non-engineers extend the
Duetto column-name matching by editing YAML alone.
"""
from __future__ import annotations

from pathlib import Path

# --- File handling ---------------------------------------------------------
SUPPORTED_EXTENSIONS: list[str] = [".csv", ".xlsx", ".xls", ".zip"]

MONTH_FORMATS_TRY: list[str] = [
    "%b, %Y", "%b-%y", "%b, %y", "%b,%y", "%b %Y", "%b %y",
    "%B, %Y", "%B-%y", "%B %Y", "%B %y",
]

# --- Default Duetto schema (used if config/duetto_schema.yaml is absent) ---
_DEFAULT_REFERENCE_PATTERNS: dict[str, list[str]] = {
    "Today":    ["Today"],
    "STLY":     ["STLY (DOW)", "STLY"],
    "ST2Y":     ["ST2Y (DOW)", "ST2Y"],
    "ST3Y":     ["ST3Y (DOW)", "ST3Y"],
    "Duetto":   ["Duetto Forecast", "Duetto", "Forecast"],
    "Budget":   ["Locked Budget", "Budget"],
    "Final LY": ["Final LY (DOW)", "Final LY"],
    "Final 2Y": ["Final 2Y (DOW)", "Final 2Y"],
    "Final 3Y": ["Final 3Y (DOW)", "Final 3Y"],
}

_DEFAULT_METRIC_PATTERNS: dict[str, list[str]] = {
    "Occupancy": ["Occupancy (Physical)", "Occupancy"],
    "Rooms":     ["Rooms (Commit)", "Rooms"],
    "ADR":       ["ADR (Commit)", "ADR"],
    "Revenue":   ["Room Revenue (Commit)", "Room Revenue", "Revenue"],
}

_DEFAULT_METRIC_TO_DISPLAY: dict[str, str] = {
    "Occupancy": "Occ",
    "Rooms":     "Room",
    "ADR":       "ADR",
    "Revenue":   "Rev",
}

_DEFAULT_METRIC_ORDER: list[str] = ["Occ", "Room", "ADR", "Rev"]


def _load_schema_from_yaml() -> dict | None:
    """Try to load schema overrides from config/duetto_schema.yaml.

    Returns None if PyYAML isn't installed or the file is missing —
    callers fall back to the in-code defaults.
    """
    schema_path = Path(__file__).resolve().parents[2] / "config" / "duetto_schema.yaml"
    if not schema_path.exists():
        return None
    try:
        import yaml  # type: ignore
    except ImportError:
        return None
    try:
        with schema_path.open("r", encoding="utf-8") as f:
            return yaml.safe_load(f)
    except Exception:
        return None


_yaml_schema = _load_schema_from_yaml()

if _yaml_schema:
    REFERENCE_PATTERNS: dict[str, list[str]] = _yaml_schema.get("references", _DEFAULT_REFERENCE_PATTERNS)
    METRIC_PATTERNS: dict[str, list[str]] = _yaml_schema.get("metrics", _DEFAULT_METRIC_PATTERNS)
    METRIC_TO_DISPLAY: dict[str, str] = _yaml_schema.get("metric_to_display", _DEFAULT_METRIC_TO_DISPLAY)
    METRIC_ORDER: list[str] = _yaml_schema.get("metric_order", _DEFAULT_METRIC_ORDER)
else:
    REFERENCE_PATTERNS = _DEFAULT_REFERENCE_PATTERNS
    METRIC_PATTERNS = _DEFAULT_METRIC_PATTERNS
    METRIC_TO_DISPLAY = _DEFAULT_METRIC_TO_DISPLAY
    METRIC_ORDER = _DEFAULT_METRIC_ORDER

# Always-derived constants (no override)
DISPLAY_TO_METRIC: dict[str, str] = {v: k for k, v in METRIC_TO_DISPLAY.items()}
SAME_TIME_REFS: list[str] = ["STLY", "ST2Y", "ST3Y"]
FINAL_REFS: list[str] = ["Final LY", "Final 2Y", "Final 3Y"]
