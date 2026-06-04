from __future__ import annotations

import io

import pandas as pd


def to_excel_bytes(sheets_dict):
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
