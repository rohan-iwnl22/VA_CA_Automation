"""Text-join Host IPs for rows sharing the same Vulnerability Title."""

from __future__ import annotations

import pandas as pd

RISK_WEIGHTS = {"Critical": 0, "High": 1, "Medium": 2, "Low": 3}


def text_join_hosts(df: pd.DataFrame) -> pd.DataFrame:
    """Group rows by Vulnerbility Title and join unique Host IPs with comma separator.

    For rows sharing the same Vulnerbility Title, this function:
    - Joins all unique Host values with ", " into a single comma-separated string
    - Takes the first non-null value for Description, Risk, Port, Recommendation,
      Reference, and CVE columns
    - Resets the index so the result is a flat DataFrame

    Parameters
    ----------
    df : pd.DataFrame
        DataFrame with columns including 'Vulnerbility Title', 'Host',
        and optionally 'Description', 'Risk', 'Port', 'Recommendation ',
        'Reference', 'CVE'.

    Returns
    -------
    pd.DataFrame
        Grouped DataFrame with one row per unique Vulnerbility Title.
    """
    if df.empty:
        return df.copy()

    join_columns = ["Description", "Risk", "Port", "Recommendation ", "Reference", "CVE"]

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

    grouped = df.groupby("Vulnerbility Title", sort=False).agg(agg_rules)
    grouped = grouped.reset_index()

    grouped["_risk_weight"] = grouped["Risk"].map(RISK_WEIGHTS)
    grouped = grouped.sort_values(by=["_risk_weight"], kind="stable").drop(
        columns=["_risk_weight"]
    )
    grouped = grouped.reset_index(drop=True)
    grouped.insert(0, "Sr. no", range(1, len(grouped) + 1))

    return grouped
