"""Risk-based filtering helpers."""

from __future__ import annotations

import pandas as pd


# VA pipeline excludes Risk = "None"; keeps Critical, High, Medium, Low
VA_EXCLUDE_RISKS = {"None"}

# Informational SSL findings to drop even if they carry a non-None risk value
EXCLUDE_NAMES = {
    "SSL Certificate Cannot Be Trusted",
    "SSL Self-Signed Certificate",
}


def filter_va_candidates(df: pd.DataFrame) -> pd.DataFrame:
    """Filter a VA-classified DataFrame to exclude 'None' risk rows and informational SSL findings.

    Parameters
    ----------
    df : pd.DataFrame
        DataFrame with 'Risk' and 'Name' columns, already classified as VA-type rows.

    Returns
    -------
    pd.DataFrame
        Filtered copy with Risk = 'None' rows and excluded names removed.
    """
    df = df[~df["Risk"].isin(VA_EXCLUDE_RISKS)].copy()
    df = df[~df["Name"].isin(EXCLUDE_NAMES)].copy()
    return df


CA_EXCLUDE_RISKS = {"PASSED"}


def filter_ca_candidates(df: pd.DataFrame) -> pd.DataFrame:
    """Filter a CA-classified DataFrame to exclude 'PASSED' rows.

    Parameters
    ----------
    df : pd.DataFrame
        DataFrame with a 'Risk' column, already classified as CA-type rows.

    Returns
    -------
    pd.DataFrame
        Filtered copy with Risk = 'PASSED' rows removed.
    """
    return df[~df["Risk"].isin(CA_EXCLUDE_RISKS)].copy()
