"""Clone pristine report templates into working files."""

from __future__ import annotations

import shutil
from pathlib import Path

from openpyxl import load_workbook


class TemplatePristencyError(Exception):
    """Raised when a template file is not in the expected pristine state."""


def clone_template(template_path: Path, output_path: Path) -> Path:
    """Copy the pristine template to a working output path.

    Verifies that the template is pristine (no data rows, no pivot tables)
    before copying.

    Returns the path to the working copy.
    """
    if not template_path.exists():
        raise FileNotFoundError(f"Template not found: {template_path}")

    _assert_template_is_pristine(template_path)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(template_path, output_path)

    return output_path


def _assert_template_is_pristine(template_path: Path) -> None:
    """Fail loudly if the template contains unexpected data rows or pivot artifacts."""
    wb = load_workbook(template_path, read_only=True)

    try:
        # Check VA Report sheet has no data rows beyond header
        if "VA Report" in wb.sheetnames:
            ws = wb["VA Report"]
            if ws.max_row and ws.max_row > 13:
                raise TemplatePristencyError(
                    f"Template VA Report sheet has {ws.max_row} rows "
                    f"(expected max row 13 for pristine template). "
                    f"Template may contain leftover data."
                )

        # Check Summary sheet has no data rows beyond headers
        if "Summary" in wb.sheetnames:
            ws = wb["Summary"]
            if ws.max_row and ws.max_row > 7:
                raise TemplatePristencyError(
                    f"Template Summary sheet has {ws.max_row} rows "
                    f"(expected max row 7 for pristine template). "
                    f"Template may contain leftover data."
                )

        # Check no pivot tables exist (openpyxl can detect these)
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            if hasattr(ws, "_pivots") and ws._pivots:
                raise TemplatePristencyError(
                    f"Template sheet '{sheet_name}' contains existing pivot tables. "
                    f"Expected a pristine template with no pivot artifacts."
                )
    finally:
        wb.close()


def load_working_copy(working_path: Path):
    """Load the working copy workbook for editing."""
    return load_workbook(working_path)
