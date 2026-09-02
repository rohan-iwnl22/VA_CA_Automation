"""Host grouping and risk sorting helpers."""

from __future__ import annotations

import pandas as pd

RISK_WEIGHTS = {"Critical": 0, "High": 1, "Medium": 2, "Low": 3}


def sort_va_data(df: pd.DataFrame) -> pd.DataFrame:
    """Sort VA data by host grouping then risk severity.

    1. Group rows into contiguous blocks by Host (alphabetical order).
    2. Within each host block, order by risk weight: Critical > High > Medium > Low.
    3. Stable tiebreak preserves original row order within same host + risk.
    4. Assign sequential Sr. no after sorting.
    """
    df = df.copy()

    # Assign host order based on alphabetical sort
    sorted_hosts = sorted(df["Host"].unique())
    host_order = {h: i for i, h in enumerate(sorted_hosts)}

    df["_host_order"] = df["Host"].map(host_order)
    df["_risk_weight"] = df["Risk"].map(RISK_WEIGHTS)

    df_sorted = df.sort_values(
        by=["_host_order", "_risk_weight"], kind="stable"
    ).drop(columns=["_host_order", "_risk_weight"])

    df_sorted = df_sorted.reset_index(drop=True)
    df_sorted.insert(0, "Sr. no", range(1, len(df_sorted) + 1))

    return df_sorted
