"""Raw-to-template column mapping helpers."""

from __future__ import annotations

import pandas as pd


# Mapping from raw column names to template column names
RAW_TO_TEMPLATE = {
    "Name": "Vulnerbility Title",       # sic - preserve template typo
    "Description": "Description",
    "Risk": "Risk",
    "Host": "Host",
    "Port": "Port",
    "Solution": "Recommendation ",       # trailing space preserved
    "See Also": "Reference",
    "CVE": "CVE",
}


def map_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Map raw DataFrame columns to template column names.

    Drops columns not used in the final report and renames remaining columns.
    """
    mapped = pd.DataFrame()

    for raw_col, template_col in RAW_TO_TEMPLATE.items():
        if raw_col in df.columns:
            mapped[template_col] = df[raw_col]
        else:
            mapped[template_col] = ""

    # Replace empty/NaN values with "N/A" for Reference and CVE
    for col in ["Reference", "CVE"]:
        if col in mapped.columns:
            mapped[col] = mapped[col].fillna("N/A")
            mapped[col] = mapped[col].replace("", "N/A")

    # Ensure Port is numeric
    if "Port" in mapped.columns:
        mapped["Port"] = pd.to_numeric(mapped["Port"], errors="coerce").fillna(0).astype(int)

    return mapped
