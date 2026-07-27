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

RAW_SHEET_NAME = "RAW File"


class SchemaError(Exception):
    """Raised when the raw workbook schema does not match expectations."""


class SheetNotFoundError(Exception):
    """Raised when the expected sheet name is not found."""


def load_raw_file(file_path: Path | str) -> pd.DataFrame:
    """Load the RAW File sheet from a Nessus export workbook.

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
        If the workbook does not contain a sheet named exactly 'RAW File'.
    SchemaError
        If the column names or count do not match the expected schema.
    """
    file_path = Path(file_path)
    if not file_path.exists():
        raise FileNotFoundError(f"Raw file not found: {file_path}")

    xls = pd.ExcelFile(file_path, engine="openpyxl")
    if RAW_SHEET_NAME not in xls.sheet_names:
        raise SheetNotFoundError(
            f"Expected sheet '{RAW_SHEET_NAME}' not found. "
            f"Available sheets: {xls.sheet_names}"
        )

    df = pd.read_excel(xls, sheet_name=RAW_SHEET_NAME, engine="openpyxl", dtype=str)
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
