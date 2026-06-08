from __future__ import annotations

import io
import re

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


def to_excel_bytes(sheets_dict: dict[str, pd.DataFrame]) -> bytes:
    """Pack a {sheet_name: DataFrame} dict into XLSX bytes."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for name, df in sheets_dict.items():
            safe_name = str(name)[:31]   # Excel sheet name max 31 chars
            if df is None or (isinstance(df, pd.DataFrame) and df.empty):
                pd.DataFrame({"(empty)": []}).to_excel(writer, sheet_name=safe_name, index=False)
            else:
                df.to_excel(writer, sheet_name=safe_name, index=False)
    output.seek(0)
    return output.getvalue()


def to_styled_duetto_pivot_excel_bytes(pivot_df: pd.DataFrame) -> bytes:
    """Build a one-sheet styled XLSX for the Duetto Pivot - by Stay Month table."""
    output = io.BytesIO()
    wb = Workbook()
    wb.remove(wb.active)

    if pivot_df is None or pivot_df.empty:
        _write_styled_pivot_sheet(wb, "Duetto Pivot", pd.DataFrame())
    else:
        _write_styled_pivot_sheet(wb, "Duetto Pivot", pivot_df)

    wb.save(output)
    output.seek(0)
    return output.getvalue()


def _safe_sheet_name(name: str) -> str:
    safe = re.sub(r"[\[\]\:\*\?\/\\]", " ", name).strip()
    return (safe or "Sheet")[:31]


def _write_styled_pivot_sheet(wb: Workbook, sheet_name: str, df: pd.DataFrame) -> None:
    ws = wb.create_sheet(_safe_sheet_name(sheet_name))
    _apply_one_page_export_setup(ws)

    if df is None or df.empty:
        ws.append(["No pivot data for selected filters."])
        ws["A1"].font = Font(bold=True, color="6B7280")
        ws.column_dimensions["A"].width = 34
        return

    columns = list(df.columns)
    ws.append(columns)

    header_fill = PatternFill("solid", fgColor="111827")
    header_font = Font(bold=True, color="FFFFFF")
    thin_border = Border(bottom=Side(style="thin", color="D1D5DB"))

    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border

    for _, row in df.iterrows():
        ws.append([row.get(col) for col in columns])

    ws.freeze_panes = "D2" if "Hotel" in columns else "C2"
    ws.auto_filter.ref = ws.dimensions

    for row_idx in range(2, ws.max_row + 1):
        values = {col: ws.cell(row=row_idx, column=idx + 1).value for idx, col in enumerate(columns)}
        metric = str(values.get("Metric", ""))
        for col_idx, col in enumerate(columns, start=1):
            cell = ws.cell(row=row_idx, column=col_idx)
            _apply_pivot_cell_style(cell, col, metric, values)

    for col_idx, col in enumerate(columns, start=1):
        letter = get_column_letter(col_idx)
        max_len = max(
            len(str(ws.cell(row=row_idx, column=col_idx).value or ""))
            for row_idx in range(1, ws.max_row + 1)
        )
        ws.column_dimensions[letter].width = min(max(max_len + 2, 11), 24)

    ws.sheet_view.showGridLines = False


def _apply_one_page_export_setup(ws) -> None:
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.page_margins.left = 0.2
    ws.page_margins.right = 0.2
    ws.page_margins.top = 0.25
    ws.page_margins.bottom = 0.25


def _apply_pivot_cell_style(cell, col: str, metric: str, row_values: dict[str, object]) -> None:
    base_fill = None
    if metric == "Rev":
        base_fill = "FFFBF5"
        cell.font = Font(bold=True if col in {"Today", "Duetto"} else False, color="111827")
    elif metric == "Occ":
        base_fill = "F9FAFB"

    if base_fill:
        cell.fill = PatternFill("solid", fgColor=base_fill)

    if col in {"Hotel", "Stay Month", "Metric"}:
        cell.font = Font(bold=(col == "Metric"), color="111827")
        cell.alignment = Alignment(horizontal="left", vertical="center")
    else:
        cell.alignment = Alignment(horizontal="right", vertical="center")

    if col in {"Today", "STLY", "ST2Y", "ST3Y", "Duetto", "Budget", "Final LY", "Final 2Y", "Final 3Y"}:
        cell.number_format = '#,##0.00'
    if " VS " in str(col):
        cell.number_format = '+0.0"%"'

    if col == "Today":
        _set_cell_colors(cell, "DBEAFE" if metric == "Rev" else "EFF6FF", "1E40AF", True)
    elif col == "Duetto":
        _set_cell_colors(cell, "DCFCE7" if metric == "Rev" else "F0FDF4", "15803D", True)
    elif col == "Budget":
        _set_cell_colors(cell, "FEFCE8", "713F12", metric == "Rev")

    if " VS " in str(col):
        try:
            value = float(row_values.get(col))
        except (TypeError, ValueError):
            value = None
        if value is not None:
            if value > 0:
                _set_cell_colors(cell, "BBF7D0", "166534", True)
            elif value < 0:
                _set_cell_colors(cell, "FECACA", "991B1B", True)
            else:
                _set_cell_colors(cell, "FEF9C3", "92400E", True)

    cell.border = Border(bottom=Side(style="thin", color="E5E7EB"))


def _set_cell_colors(cell, fill: str, font_color: str, bold: bool = False) -> None:
    cell.fill = PatternFill("solid", fgColor=fill)
    cell.font = Font(bold=bold, color=font_color)
