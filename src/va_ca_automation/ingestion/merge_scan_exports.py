"""Merge multiple vulnerability-scan export files into a single RAW_File.xlsx format."""

from __future__ import annotations

import logging
import sys
from pathlib import Path

import pandas as pd

logger = logging.getLogger("va_ca_automation")

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

TEXT_COLUMNS = [
    "CVE", "Risk", "Host", "Protocol", "Name", "Synopsis",
    "Description", "Solution", "See Also", "Plugin Output",
]
INT_COLUMNS = ["Plugin ID", "Port"]
FLOAT_COLUMNS = [
    "CVSS v2.0 Base Score", "CVSS v4.0 Base Score",
    "CVSS v3.0 Base Score", "VPR Score", "EPSS Score",
]


def load_one_file(path: Path) -> pd.DataFrame | None:
    """Load a single export file and validate/normalize its columns.

    IMPORTANT: this scan format uses the literal text values "None" (a Risk
    level) and "n/a" (a Solution placeholder) as real, meaningful data -- not
    as missing-value markers. pandas' default NA-detection treats both of
    those strings (along with "NA", "NULL", "nan", etc.) as missing data and
    will silently blank them out. We disable that default behavior entirely
    (keep_default_na=False, na_values=[]) and only decide what counts as
    "empty" ourselves, afterwards.
    """
    try:
        read_kwargs = dict(keep_default_na=False, na_values=[], dtype=str)
        if path.suffix.lower() == ".csv":
            df = pd.read_csv(path, **read_kwargs)
        elif path.suffix.lower() in (".xlsx", ".xls"):
            df = pd.read_excel(path, **read_kwargs)
        else:
            return None
    except Exception as exc:
        logger.warning("Could not read %s: %s", path.name, exc)
        return None

    df.columns = [str(c).strip() for c in df.columns]

    if set(df.columns) != set(EXPECTED_COLUMNS):
        missing = set(EXPECTED_COLUMNS) - set(df.columns)
        extra = set(df.columns) - set(EXPECTED_COLUMNS)
        logger.warning(
            "%s does not match expected schema. Missing: %s | Unexpected: %s",
            path.name,
            missing or "none",
            extra or "none",
        )
        return None

    df = df[EXPECTED_COLUMNS]

    # Only a genuinely empty cell ("") should become blank. "None" and "n/a"
    # are real values in this dataset and must survive untouched.
    for col in TEXT_COLUMNS:
        df[col] = df[col].apply(lambda v: None if v == "" else v)

    # Numeric columns: an empty string legitimately means "no score" -> blank.
    # pd.to_numeric's errors="coerce" turns "" (and any non-numeric junk) into
    # NaN, which is exactly what we want here.
    for col in INT_COLUMNS:
        df[col] = pd.to_numeric(df[col], errors="coerce").astype("Int64")
    for col in FLOAT_COLUMNS:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    return df


def merge_scan_exports(
    files: list[Path],
    dedupe: bool = False,
) -> pd.DataFrame | None:
    """Merge multiple scan export files into a single DataFrame.

    Parameters
    ----------
    files : list[Path]
        List of .csv or .xlsx files to merge.
    dedupe : bool
        If True, drop exact duplicate rows across files.

    Returns
    -------
    pd.DataFrame or None
        Combined DataFrame with all validated rows, or None if no valid files.
    """
    frames = []
    skipped = []

    for path in files:
        df = load_one_file(path)
        if df is not None:
            logger.info("[OK] %s: %d rows", path.name, len(df))
            frames.append(df)
        else:
            skipped.append(path.name)

    if skipped:
        logger.warning("Skipped %d file(s): %s", len(skipped), ", ".join(skipped))

    if not frames:
        logger.error("No valid files to merge")
        return None

    combined = pd.concat(frames, ignore_index=True)

    if dedupe:
        before = len(combined)
        combined = combined.drop_duplicates()
        logger.info("Deduped: %d -> %d rows", before, len(combined))

    return combined


def merge_to_excel(
    files: list[Path],
    output_path: Path,
    dedupe: bool = False,
) -> dict | None:
    """Merge files and write to a single-sheet .xlsx.

    Parameters
    ----------
    files : list[Path]
        List of .csv or .xlsx files to merge.
    output_path : Path
        Path to write the combined .xlsx.
    dedupe : bool
        If True, drop exact duplicate rows.

    Returns
    -------
    dict or None
        Summary with keys: files_merged, skipped, rows_per_file, total_rows.
        Returns None if merge failed.
    """
    frames = []
    rows_per_file = {}
    skipped = []

    for path in files:
        df = load_one_file(path)
        if df is not None:
            logger.info("[OK] %s: %d rows", path.name, len(df))
            frames.append(df)
            rows_per_file[path.name] = len(df)
        else:
            skipped.append(path.name)

    if skipped:
        logger.warning("Skipped %d file(s): %s", len(skipped), ", ".join(skipped))

    if not frames:
        logger.error("No valid files to merge")
        return None

    combined = pd.concat(frames, ignore_index=True)

    if dedupe:
        before = len(combined)
        combined = combined.drop_duplicates()
        logger.info("Deduped: %d -> %d rows", before, len(combined))

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        combined.to_excel(writer, index=False, sheet_name="RAW File")

    logger.info("Wrote %d rows to %s", len(combined), output_path)

    return {
        "files_merged": len(frames),
        "skipped": len(skipped),
        "rows_per_file": rows_per_file,
        "total_rows": len(combined),
    }
