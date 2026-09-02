"""Load and validate the Nessus raw export workbook."""

from __future__ import annotations

from pathlib import Path

import pandas as pd


EXPECTED_COLUMNS = [
    "Plugin ID",
    "CVE",
    "CVSS v2.0 Base Score",
    "Risk",
    "Host",
    "Protocol",
    "Port",
    "Name",
    "Synopsis",
    "Description",
    "Solution",
    "See Also",
    "Plugin Output",
    "CVSS v4.0 Base Score",
    "CVSS v3.0 Base Score",
    "VPR Score",
    "EPSS Score",
]

PREFERRED_SHEET_NAMES = ["RAW File", "RAW Sever"]


class SchemaError(Exception):
    """Raised when the raw workbook schema does not match expectations."""


class SheetNotFoundError(Exception):
    """Raised when the expected sheet name is not found."""


def _find_raw_sheet(xls: pd.ExcelFile) -> str:
    """Find the sheet containing raw Nessus data.

    Checks preferred names first, then falls back to any sheet with the
    correct 17-column schema.
    """
    for name in PREFERRED_SHEET_NAMES:
        if name in xls.sheet_names:
            return name

    for name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=name, engine="openpyxl", nrows=0, dtype=str)
        if len(df.columns) == len(EXPECTED_COLUMNS):
            actual = list(df.columns)
            if actual == EXPECTED_COLUMNS:
                return name

    raise SheetNotFoundError(
        f"No sheet with the expected {len(EXPECTED_COLUMNS)}-column schema found. "
        f"Available sheets: {xls.sheet_names}"
    )


def load_raw_file(file_path: Path | str) -> pd.DataFrame:
    """Load the raw Nessus data sheet from a workbook.

    Automatically detects the correct sheet by checking preferred names
    first, then falling back to schema validation.

    Parameters
    ----------
    file_path : Path or str
        Path to the .xlsx workbook.

    Returns
    -------
    pd.DataFrame
        Raw data with the expected 17-column schema.

    Raises
    ------
    SheetNotFoundError
        If no sheet with the expected schema is found.
    SchemaError
        If the column names or count do not match the expected schema.
    """
    file_path = Path(file_path)
    if not file_path.exists():
        raise FileNotFoundError(f"Raw file not found: {file_path}")

    xls = pd.ExcelFile(file_path, engine="openpyxl")
    try:
        sheet_name = _find_raw_sheet(xls)
        df = pd.read_excel(xls, sheet_name=sheet_name, engine="openpyxl", dtype=str)
    finally:
        xls.close()

    df = df.fillna("")
    _validate_schema(df)
    return df


def _validate_schema(df: pd.DataFrame) -> None:
    """Validate that the DataFrame has the expected 17-column schema."""
    actual_columns = list(df.columns)

    if len(actual_columns) != len(EXPECTED_COLUMNS):
        raise SchemaError(
            f"Expected {len(EXPECTED_COLUMNS)} columns, found {len(actual_columns)}. "
            f"Actual: {actual_columns}"
        )

    for i, (actual, expected) in enumerate(zip(actual_columns, EXPECTED_COLUMNS)):
        if actual != expected:
            raise SchemaError(
                f"Column {i + 1} mismatch: expected '{expected}', found '{actual}'"
            )
