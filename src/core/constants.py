from __future__ import annotations


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
