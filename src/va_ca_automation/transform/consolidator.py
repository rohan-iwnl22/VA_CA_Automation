"""Consolidate rows by Vulnerability Name, joining unique IP addresses."""

from __future__ import annotations

import pandas as pd


def consolidate_by_vuln_name(df: pd.DataFrame) -> pd.DataFrame:
    """Group rows by Vulnerability Name and join unique Host IPs with comma separator.

    For rows sharing the same Vulnerability Name, this function:
    - Joins all unique Host values with ", " into a single comma-separated string
    - Takes the first non-null value for Description, Risk, Port, Recommendation,
      Reference, and CVE columns
    - Returns one row per unique Vulnerability Name

    Parameters
    ----------
    df : pd.DataFrame
        DataFrame with columns including 'Vulnerability Name', 'Host',
        and optionally 'Description', 'Risk', 'Port', 'Recommendation',
        'Reference', 'CVE'.

    Returns
    -------
    pd.DataFrame
        Grouped DataFrame with one row per unique Vulnerability Name.
    """
    if df.empty:
        return df.copy()

    join_columns = ["Description", "Risk", "Port", "Recommendation", "Reference", "CVE"]

    def _first_non_null(s):
        non_null = s.dropna()
        if non_null.empty:
            return ""
        return non_null.iloc[0]

    def _unique_hosts(hosts):
        seen = []
        for h in hosts.dropna():
            if h not in seen:
                seen.append(h)
        return ", ".join(seen) if seen else ""

    agg_rules = {}
    for col in join_columns:
        if col in df.columns:
            agg_rules[col] = _first_non_null

    if "Host" in df.columns:
        agg_rules["Host"] = _unique_hosts

    grouped = df.groupby("Vulnerability Name", sort=False).agg(agg_rules)
    grouped = grouped.reset_index()

    return grouped
