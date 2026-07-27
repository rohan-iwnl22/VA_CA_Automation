"""Write processed findings into the report template."""

from __future__ import annotations

from copy import copy

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import get_column_letter


# Template column order (must match column_mapper.py output)
TEMPLATE_COLUMNS = [
    "Sr. no",
    "Vulnerbility Title",
    "Description",
    "Risk",
    "Host",
    "Port",
    "Recommendation ",
    "Reference",
    "CVE",
]

THIN_BORDER = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin"),
)

WRAP_COLUMNS = {"Vulnerbility Title", "Description", "Recommendation ", "Reference"}

DATA_FONT = Font(name="Cambria", size=11)


def write_va_report_header(ws, metadata) -> None:
    """Write the VA Report header metadata block (rows 5-10, column C)."""
    ws["C5"] = metadata.client_name
    ws["C6"] = metadata.security_tester
    ws["C7"] = metadata.reviewed_by
    ws["C8"] = metadata.report_date
    ws["C9"] = metadata.report_version
    ws["C10"] = metadata.scanner_name


def write_va_data_rows(ws, df: pd.DataFrame) -> int:
    """Write processed VA data rows starting at row 14.

    Returns the last row written to.
    """
    data_start_row = 14

    for i, row in df.iterrows():
        excel_row = data_start_row + i
        for j, col_name in enumerate(TEMPLATE_COLUMNS):
            cell = ws.cell(row=excel_row, column=j + 1)

            value = row.get(col_name)
            if pd.isna(value) or value == "":
                cell.value = None
            else:
                cell.value = value

            _apply_data_cell_style(cell, col_name)

    return data_start_row + len(df) - 1 if len(df) > 0 else data_start_row - 1


def _apply_data_cell_style(cell, col_name: str) -> None:
    """Apply Cambria 11pt font, thin border, and wrap_text for appropriate columns."""
    cell.font = DATA_FONT
    cell.border = THIN_BORDER

    if col_name in WRAP_COLUMNS:
        cell.alignment = Alignment(wrap_text=True, vertical="top")
    else:
        cell.alignment = Alignment(vertical="top")


def clone_row_style_from_template(ws, source_row: int = 13) -> dict:
    """Capture the style properties from a template row for cloning.

    In a pristine template, row 13 has headers. We capture styles from
    the header-adjacent style to apply to data rows.
    """
    styles = {}
    for col in range(1, 10):
        cell = ws.cell(row=source_row, column=col)
        styles[col] = {
            "font": copy(cell.font) if cell.font else None,
            "border": copy(cell.border) if cell.border else None,
            "alignment": copy(cell.alignment) if cell.alignment else None,
        }
    return styles


def write_introduction_fields(ws, metadata) -> None:
    """Write dynamic fields to the Introduction sheet."""
    if metadata.scanner_name:
        ws["B14"] = metadata.scanner_name
    if metadata.scanner_version:
        ws["B15"] = metadata.scanner_version
    if metadata.report_owner:
        ws["B19"] = metadata.report_owner
