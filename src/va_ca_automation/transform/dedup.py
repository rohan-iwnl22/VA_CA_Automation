"""Deduplication helpers — two-stage process."""

from __future__ import annotations

import re
from typing import Any

import pandas as pd

from ..logging.pipeline_logger import PipelineLogger

VERSION_PATTERN = re.compile(r'(\d+(?:\.\d+){1,4})')


def stage1_exact_dedup(df: pd.DataFrame) -> pd.DataFrame:
    """Stage 1: exact duplicate removal on (Name, Description, Risk, Host).

    Keeps the first occurrence per key (stable).
    """
    key_fields = ["Name", "Description", "Risk", "Host"]
    return df.drop_duplicates(subset=key_fields, keep="first").copy()


def _extract_version(name_text: str) -> str | None:
    """Extract the last dotted-numeric version token from a name string."""
    matches = VERSION_PATTERN.findall(name_text)
    if not matches:
        return None
    return matches[-1]


def _version_to_tuple(version_str: str) -> tuple[int, ...]:
    """Convert a version string to a zero-padded integer tuple for comparison."""
    parts = version_str.split(".")
    nums = [int(p) for p in parts]
    while len(nums) < 4:
        nums.append(0)
    return tuple(nums)


def _base_title(name_text: str) -> str:
    """Strip the version token from the name to produce a grouping key."""
    return VERSION_PATTERN.sub("", name_text).strip()


def stage2_version_collapse(
    df: pd.DataFrame, plogger: PipelineLogger
) -> pd.DataFrame:
    """Stage 2: version-collapse dedup.

    Groups rows by (base_title, Risk, Host). Within each group with multiple
    versioned rows, keeps only the row with the highest version number.
    Rows with no parsable version token pass through unchanged.
    """
    df = df.copy()
    df["_version_raw"] = df["Name"].apply(_extract_version)
    df["_base_title"] = df["Name"].apply(_base_title)

    kept_rows: list[pd.DataFrame] = []
    collapse_log: list[dict[str, Any]] = []

    for (base_title_val, risk_val, host_val), group in df.groupby(
        ["_base_title", "Risk", "Host"]
    ):
        no_version_rows = group[group["_version_raw"].isna()]
        versioned_rows = group[group["_version_raw"].notna()]

        if not no_version_rows.empty:
            kept_rows.append(no_version_rows)

        if versioned_rows.empty:
            continue

        if len(versioned_rows) == 1:
            kept_rows.append(versioned_rows)
            continue

        versioned_rows = versioned_rows.copy()
        versioned_rows["_version_tuple"] = versioned_rows["_version_raw"].apply(
            _version_to_tuple
        )
        versioned_rows_sorted = versioned_rows.sort_values(
            "_version_tuple", ascending=False
        )
        winner = versioned_rows_sorted.iloc[[0]]
        losers = versioned_rows_sorted.iloc[1:]

        kept_rows.append(winner)
        collapse_log.append(
            {
                "base_title": base_title_val,
                "host": host_val,
                "risk": risk_val,
                "kept_version": winner["_version_raw"].iloc[0],
                "dropped_versions": list(losers["_version_raw"]),
            }
        )

    if kept_rows:
        result = pd.concat(kept_rows, ignore_index=True)
    else:
        result = pd.DataFrame(columns=df.columns)

    result = result.drop(columns=["_version_raw", "_base_title"], errors="ignore")

    plogger.log_version_collapse(collapse_log)

    return result
