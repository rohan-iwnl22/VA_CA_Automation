"""POST /api/final-report — Compare first audit with retest and add Retest Status."""

from __future__ import annotations

import io
import tempfile
from pathlib import Path

from openpyxl import load_workbook
from fastapi import APIRouter, Depends, File, UploadFile
from fastapi.responses import StreamingResponse

from ..deps import get_current_user

router = APIRouter(tags=["final-report"])

HEADER_ROW = 13
DATA_START_ROW = 14


def _detect_report_type(wb) -> str:
    """Detect if the workbook is VA or CA based on sheet names."""
    sheet_names = wb.sheetnames
    for name in sheet_names:
        lower = name.lower().replace(" ", "").replace("_", "")
        if "vareport" in lower or lower == "vareport":
            return "va"
        if "careport" in lower or lower == "careport":
            return "ca"
    return "va"


def _get_column_indices(ws, report_type: str) -> dict:
    """Return column indices (1-based) for vulnerability title, host, and the last data column."""
    if report_type == "va":
        return {"title_col": 2, "host_col": 5, "last_data_col": 9}
    else:
        return {"title_col": 2, "host_col": 3, "last_data_col": 6}


def _build_retest_lookup(ws, report_type: str) -> set:
    """Build a set of (title, host) tuples from the retest file."""
    cols = _get_column_indices(ws, report_type)
    lookup = set()
    for row in range(DATA_START_ROW, ws.max_row + 1):
        title = ws.cell(row=row, column=cols["title_col"]).value
        host = ws.cell(row=row, column=cols["host_col"]).value
        if title is not None and host is not None:
            lookup.add((str(title).strip(), str(host).strip()))
    return lookup


def _add_retest_status(ws, retest_lookup: set, report_type: str) -> None:
    """Add Retest Status column to the first audit worksheet."""
    cols = _get_column_indices(ws, report_type)
    retest_col = cols["last_data_col"] + 1

    # Write header
    header_cell = ws.cell(row=HEADER_ROW, column=retest_col)
    header_cell.value = "Retest Status"

    # Write data rows
    for row in range(DATA_START_ROW, ws.max_row + 1):
        title = ws.cell(row=row, column=cols["title_col"]).value
        host = ws.cell(row=row, column=cols["host_col"]).value
        cell = ws.cell(row=row, column=retest_col)

        if title is not None and host is not None:
            key = (str(title).strip(), str(host).strip())
            cell.value = "Open" if key in retest_lookup else "Closed"
        else:
            cell.value = "Closed"


@router.post("/final-report")
async def generate_final_report(
    first_audit: UploadFile = File(...),
    retest_file: UploadFile = File(...),
    current_user: dict = Depends(get_current_user),
):
    """Compare first audit with retest file and add Retest Status column.

    Accepts two .xlsx files. The first audit file is the base, and the retest
    file contains retested vulnerabilities. A 'Retest Status' column is added
    to the first audit: 'Open' if the vulnerability+host pair exists in the
    retest file, 'Closed' otherwise.
    """
    # Save uploaded files
    first_content = await first_audit.read()
    retest_content = await retest_file.read()

    first_wb = load_workbook(io.BytesIO(first_content))
    retest_wb = load_workbook(io.BytesIO(retest_content))

    # Detect report type from first audit
    report_type = _detect_report_type(first_wb)

    # Build lookup from retest file
    retest_ws = retest_wb.active
    retest_type = _detect_report_type(retest_wb)
    retest_lookup = _build_retest_lookup(retest_ws, retest_type)

    # Add retest status to first audit
    first_ws = first_wb.active
    _add_retest_status(first_ws, retest_lookup, report_type)

    # Save to buffer
    buf = io.BytesIO()
    first_wb.save(buf)
    buf.seek(0)

    # Determine output filename
    base_name = Path(first_audit.filename).stem if first_audit.filename else "final_report"
    output_name = f"{base_name}_Final.xlsx"

    return StreamingResponse(
        buf,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={output_name}"},
    )
