"""Risk-based filtering helpers."""

from __future__ import annotations

import pandas as pd


# VA pipeline excludes Risk = "None"; keeps Critical, High, Medium, Low
VA_EXCLUDE_RISKS = {"None"}


def filter_va_candidates(df: pd.DataFrame) -> pd.DataFrame:
    """Filter a VA-classified DataFrame to exclude 'None' risk rows.

    Parameters
    ----------
    df : pd.DataFrame
        DataFrame with a 'Risk' column, already classified as VA-type rows.

    Returns
    -------
    pd.DataFrame
        Filtered copy with Risk = 'None' rows removed.
    """
    return df[~df["Risk"].isin(VA_EXCLUDE_RISKS)].copy()


def filter_ca_candidates(df: pd.DataFrame) -> pd.DataFrame:
    """Placeholder for CA filtering (deferred to future phase).

    Currently returns the input unchanged.
    """
    return df.copy()
