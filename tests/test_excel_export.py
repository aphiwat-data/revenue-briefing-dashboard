from __future__ import annotations

import io

import pandas as pd
from openpyxl import load_workbook

from src.services.excel_export import to_styled_duetto_pivot_excel_bytes


def _fill(cell) -> str:
    return str(cell.fill.fgColor.rgb or cell.fill.fgColor.indexed)


def test_styled_duetto_pivot_excel_keeps_variance_colors():
    df = pd.DataFrame(
        [
            {
                "Hotel": "G5 Test Hotel",
                "Stay Month": "Jun 2026",
                "Metric": "Rev",
                "Today": 120,
                "Budget": 100,
                "Duetto": 130,
                "Today VS BUD": 20,
                "Duetto VS BUD": -10,
            }
        ]
    )

    wb = load_workbook(io.BytesIO(to_styled_duetto_pivot_excel_bytes(df)))
    ws = wb["All Hotels"]

    assert ws.freeze_panes == "D2"
    assert "G5 Test Hotel" in wb.sheetnames
    assert _fill(ws["G2"]).endswith("BBF7D0")
    assert _fill(ws["H2"]).endswith("FECACA")
    assert ws["G2"].number_format == '+0.0"%"'


def test_styled_duetto_pivot_excel_writes_empty_state():
    wb = load_workbook(io.BytesIO(to_styled_duetto_pivot_excel_bytes(pd.DataFrame())))

    assert wb["All Hotels"]["A1"].value == "No pivot data for selected filters."
