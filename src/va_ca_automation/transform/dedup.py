"""Deduplication helpers — two-stage process."""

from __future__ import annotations

import re
from typing import Any

import pandas as pd

from ..logging.pipeline_logger import PipelineLogger

VERSION_PATTERN = re.compile(r'(\d+(?:\.\d+){1,4})')
VERSION_FULL_PATTERN = re.compile(r'(\d+(?:\.\d+){1,4}[a-zA-Z_]*)')
RHSA_PATTERN = re.compile(r'\(RHSA-(\d+):(\d+)\)')
CPU_DATE_PATTERN = re.compile(
    r'\(?\s*'
    r'(?:(?:January|February|March|April|May|June|July|August|September|October|November|December)'
    r'|(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*?)'
    r'\s*\d{4}\s*CPU\s*\)?'
)
ORACLE_JAVA_VERSION_RE = re.compile(
    r'(?:\d+\.\d+\.x\s*<\s*)?\d+\.\d+(?:\.\d+)?(?:_\d+)?(?:\s*/\s*(?:\d+\.\d+\.x\s*<\s*)?\d+\.\d+(?:\.\d+)?(?:_\d+)?)+'
)

MONTH_MAP = {
    "jan": 1, "january": 1,
    "feb": 2, "february": 2,
    "mar": 3, "march": 3,
    "apr": 4, "april": 4,
    "may": 5,
    "jun": 6, "june": 6,
    "jul": 7, "july": 7,
    "aug": 8, "august": 8,
    "sep": 9, "september": 9,
    "oct": 10, "october": 10,
    "nov": 11, "november": 11,
    "dec": 12, "december": 12,
}


def stage1_exact_dedup(df: pd.DataFrame) -> pd.DataFrame:
    """Stage 1: exact duplicate removal on (Name, Description, Risk, Host).

    Keeps the first occurrence per key (stable).
    """
    key_fields = ["Name", "Description", "Risk", "Host"]
    return df.drop_duplicates(subset=key_fields, keep="first").copy()


def stage1b_name_host_dedup(df: pd.DataFrame) -> pd.DataFrame:
    """Stage 1b: collapse same (Name, Host) pairs to one row.

    The same vulnerability may appear on the same host on multiple ports,
    producing slightly different Description text per port.  This stage
    collapses those down to a single row per unique (vulnerability name,
    host) pair so that the downstream version-collapse only has to deal
    with genuinely different versions.

    Keeps the first occurrence per key (stable).
    """
    key_fields = ["Name", "Host"]
    return df.drop_duplicates(subset=key_fields, keep="first").copy()


def _extract_rhsa(name_text: str) -> tuple[int, int] | None:
    """Extract the RHSA advisory ID as a (year, number) tuple."""
    match = RHSA_PATTERN.search(name_text)
    if not match:
        return None
    return (int(match.group(1)), int(match.group(2)))


def _extract_version(name_text: str) -> str | None:
    """Extract the last dotted-numeric version token from a name string.

    Uses VERSION_FULL_PATTERN to capture trailing letters/underscores
    (e.g. '1.0.2p', '1.0.2zn') so that patch-level comparison works.
    """
    matches = VERSION_FULL_PATTERN.findall(name_text)
    if not matches:
        return None
    return matches[-1]


def _version_to_tuple(version_str: str) -> tuple[int, ...]:
    """Convert a version string to a zero-padded integer tuple for comparison.

    Handles suffix letters by converting them to their ordinal values,
    so '1.0.2p' becomes (1, 0, 2, 112) and '1.0.2zn' becomes
    (1, 0, 2, 122, 110), allowing correct ordering.
    """
    parts = re.split(r'[._]', version_str)
    result: list[int] = []
    for p in parts:
        m = re.match(r'(\d+)(.*)', p)
        if m:
            result.append(int(m.group(1)))
            suffix = m.group(2)
            if suffix:
                for ch in suffix:
                    result.append(ord(ch))
        else:
            result.append(0)
    while len(result) < 8:
        result.append(0)
    return tuple(result[:8])


def _strip_version(name_text: str) -> str:
    """Strip the version token from the name to produce a grouping key."""
    return VERSION_PATTERN.sub("", name_text).strip()


def _strip_rhsa(name_text: str) -> str:
    """Strip the RHSA advisory ID from the name to produce a grouping key."""
    return RHSA_PATTERN.sub("", name_text).strip()


def _extract_cpu_date(name_text: str) -> tuple[int, int] | None:
    """Extract a CPU date as a (year, month) tuple from a vulnerability name.

    Matches patterns like "(January 2026 CPU)", "(July 2025 CPU)",
    "(Oct 2025 CPU)", etc.
    """
    match = CPU_DATE_PATTERN.search(name_text)
    if not match:
        return None
    matched_text = match.group(0)
    month_match = re.search(
        r'(?:January|February|March|April|May|June|July|August|September|October|November|December'
        r'|Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*',
        matched_text,
    )
    year_match = re.search(r'(\d{4})', matched_text)
    if not month_match or not year_match:
        return None
    month_str = month_match.group(0).lower()[:3]
    month_num = MONTH_MAP.get(month_str, 0)
    year_num = int(year_match.group(1))
    return (year_num, month_num)


def _extract_identifier(name_text: str) -> tuple[int, ...] | None:
    """Extract the best comparable identifier from a vulnerability name.

    Priority: RHSA advisory ID > CPU date > dotted version number.
    Returns a tuple for comparison, or None if no identifier found.
    """
    rhsa = _extract_rhsa(name_text)
    if rhsa is not None:
        return rhsa

    cpu = _extract_cpu_date(name_text)
    if cpu is not None:
        return cpu

    ver = _extract_version(name_text)
    if ver is not None:
        return _version_to_tuple(ver)

    return None


def _make_base_title(name_text: str) -> str:
    """Strip all version/RHSA/CPU-date tokens to produce a grouping key.

    Aggressively normalises vulnerability names so that different versions
    or CPU patches of the same underlying vulnerability map to the same
    grouping key.
    """
    title = name_text
    # Strip Oracle Java multi-version patterns (1.7.0_221 / 1.8.0_211 / ...)
    title = ORACLE_JAVA_VERSION_RE.sub("", title)
    # Strip dotted-underscore patterns like 1.7.x < 1.7.0_211
    title = re.sub(r'\d+\.\d+\.x\s*<\s*', '', title)
    # Strip RHSA advisory IDs
    title = _strip_rhsa(title)
    # Strip full version tokens including trailing letters (1.0.2p, 1.0.2zn)
    title = VERSION_FULL_PATTERN.sub("", title)
    # Strip CPU date patterns (January 2026 CPU)
    title = CPU_DATE_PATTERN.sub("", title)
    # Strip standalone (Unix) / (Unix prefix suffixes
    title = re.sub(r'\(Unix\s*\)?', '', title)
    # Strip CVE references
    title = re.sub(r'CVE-\d+-\d+', '', title)
    # Normalize vulnerability description variants
    title = re.sub(r'\bMultiple\s+Vulnerabilit\w*', 'Vulnerability', title)
    title = re.sub(r'\bInformation\s+Disclosure\b', 'Vulnerability', title)
    # Collapse multiple spaces
    title = re.sub(r'\s+', ' ', title).strip()
    return title


def stage2_version_collapse(
    df: pd.DataFrame, plogger: PipelineLogger
) -> pd.DataFrame:
    """Stage 2: unified version-collapse dedup (per host).

    Groups rows by (base_title, Host). Within each group with
    multiple versioned/RHSA-tagged rows, keeps only the row with the
    highest identifier (RHSA advisory ID or version number).

    Handles ALL vulnerability name patterns: RHSA advisories, dotted
    version numbers, or any combination. Rows with no parsable
    identifier pass through unchanged.

    The same vulnerability name may appear on different hosts (each
    keeping its own latest), but within a single host only one row
    per unique base vulnerability name is kept.
    """
    df = df.copy()
    df["_id_raw"] = df["Name"].apply(_extract_identifier)
    df["_base_title"] = df["Name"].apply(_make_base_title)

    kept_rows: list[pd.DataFrame] = []
    collapse_log: list[dict[str, Any]] = []

    for (base_val, host_val), group in df.groupby(
        ["_base_title", "Host"]
    ):
        no_id_rows = group[group["_id_raw"].isna()]
        id_rows = group[group["_id_raw"].notna()]

        if not no_id_rows.empty:
            kept_rows.append(no_id_rows)

        if id_rows.empty:
            continue

        if len(id_rows) == 1:
            kept_rows.append(id_rows)
            continue

        id_rows = id_rows.copy()
        id_rows_sorted = id_rows.sort_values(
            "_id_raw", ascending=False
        )
        winner = id_rows_sorted.iloc[[0]]
        losers = id_rows_sorted.iloc[1:]

        kept_rows.append(winner)
        collapse_log.append(
            {
                "base_title": base_val,
                "host": host_val,
                "kept_name": winner["Name"].iloc[0],
                "dropped_names": list(losers["Name"]),
            }
        )

    if kept_rows:
        result = pd.concat(kept_rows, ignore_index=True)
    else:
        result = pd.DataFrame(columns=df.columns)

    result = result.drop(columns=["_id_raw", "_base_title"], errors="ignore")

    plogger.log_version_collapse(collapse_log)

    return result
