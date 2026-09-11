"""Deduplication helpers — two-stage process + KB dedup."""

from __future__ import annotations

import re
from typing import Any

import pandas as pd

from ..logging.pipeline_logger import PipelineLogger

VERSION_PATTERN = re.compile(r'(\d+(?:\.\d+){1,4})')
VERSION_FULL_PATTERN = re.compile(r'(\d+(?:\.\d+){1,4}[a-zA-Z0-9_]*)')
RHSA_PATTERN = re.compile(r'\(RHSA-(\d+):(\d+)\)')
CPU_DATE_PATTERN = re.compile(
    r'\(?\s*'
    r'(?:(?:January|February|March|April|May|June|July|August|September|October|November|December)'
    r'|(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*?)'
    r'\s*\d{4}\s*CPU\s*\)?'
)
# Matches "(Month Year)" without CPU suffix — used by KB patches
# e.g. "(May 2026)", "(July 2025)", "(April 2024 CPU)" is handled above
MONTH_YEAR_PATTERN = re.compile(
    r'\(\s*'
    r'(?:January|February|March|April|May|June|July|August|September|October|November|December'
    r'|Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*?'
    r'\s*\d{4}\s*\)'
)
KB_PATTERN = re.compile(r'^KB(\d+)', re.IGNORECASE)
# Matches multi-range version patterns like "10.x < 10.22 / 11.x < 11.17 / 12.x < 12.12"
MULTI_RANGE_PATTERN = re.compile(
    r'(?:\d+\.x\s*<\s*\d+\.\d+(?:\s*/\s*)?)+'
)
# Matches single "N.x < N.M" range
SINGLE_RANGE_PATTERN = re.compile(r'(\d+)\.x\s*<\s*(\d+\.\d+)')
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
    """Extract the best version identifier from a vulnerability name.

    Handles:
    - Multi-range patterns like "10.x < 10.22 / 11.x < 11.17" → extracts highest (11.17)
    - Single version tokens like "8.0.46" → returns as-is
    - Version patterns with letters like "9.3p2"
    """
    # Try multi-range pattern first (PostgreSQL style: 10.x < 10.22 / 11.x < 11.17)
    multi_match = MULTI_RANGE_PATTERN.search(name_text)
    if multi_match:
        range_text = multi_match.group(0)
        pairs = SINGLE_RANGE_PATTERN.findall(range_text)
        if pairs:
            # Find the highest version number from all ranges
            best = ""
            for _major, ver in pairs:
                if _version_tuple_compare(ver, best) > 0:
                    best = ver
            return best if best else None

    # Fallback: extract last dotted-numeric version token
    matches = VERSION_FULL_PATTERN.findall(name_text)
    if not matches:
        return None
    return matches[-1]


def _version_tuple_compare(a: str, b: str) -> int:
    """Compare two version strings. Returns >0 if a>b, <0 if a<b, 0 if equal."""
    a_parts = [int(x) for x in re.split(r'[._]', a) if x.isdigit()]
    b_parts = [int(x) for x in re.split(r'[._]', b) if x.isdigit()]
    for av, bv in zip(a_parts, b_parts):
        if av != bv:
            return av - bv
    return len(a_parts) - len(b_parts)


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
    """Extract a CPU or month-year date as a (year, month) tuple.

    Matches patterns like "(January 2026 CPU)", "(July 2025 CPU)",
    "(May 2026)", "(June 2024)", etc.
    """
    # Try CPU pattern first (has "CPU" suffix)
    match = CPU_DATE_PATTERN.search(name_text)
    if not match:
        # Try month-year pattern without CPU suffix (used by KB patches)
        match = MONTH_YEAR_PATTERN.search(name_text)
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


def _extract_kb_number(name_text: str) -> int | None:
    """Extract KB number from name like 'KB5099538: ...'."""
    match = KB_PATTERN.search(name_text)
    if match:
        return int(match.group(1))
    return None


def _make_kb_base_key(name_text: str) -> str | None:
    """Extract a grouping key for KB patches.

    Returns the OS/platform description after stripping the KB number
    and date, so that multiple KBs for the same OS on the same host
    are grouped together. Returns None if not a KB patch.
    """
    if not KB_PATTERN.search(name_text):
        return None
    title = name_text
    # Strip the KB prefix
    title = KB_PATTERN.sub("", title)
    # Strip the month-year date in parentheses
    title = MONTH_YEAR_PATTERN.sub("", title)
    # Strip CPU date pattern
    title = CPU_DATE_PATTERN.sub("", title)
    # Strip CVE references
    title = re.sub(r'CVE-\d+-\d+', '', title)
    # Collapse whitespace
    title = re.sub(r'\s+', ' ', title).strip()
    # Strip leading colon or dash
    title = re.sub(r'^[\s:.-]+', '', title)
    return title


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
    """Strip all version/RHSA/CPU-date/KB-date tokens to produce a grouping key.

    Aggressively normalises vulnerability names so that different versions
    or CPU patches of the same underlying vulnerability map to the same
    grouping key.
    """
    title = name_text
    # Strip Oracle Java multi-version patterns (1.7.0_221 / 1.8.0_211 / ...)
    title = ORACLE_JAVA_VERSION_RE.sub("", title)
    # Strip dotted-underscore patterns like 1.7.x < 1.7.0_211
    title = re.sub(r'\d+\.\d+\.x\s*<\s*', '', title)
    # Strip multi-range patterns like "10.x < 10.22 / 11.x < 11.17"
    title = MULTI_RANGE_PATTERN.sub("", title)
    # Strip standalone N.x patterns (e.g. "10.x", "11.x")
    title = re.sub(r'\d+\.x\b', '', title)
    # Strip comparison operators before version numbers (<, <=, >, >=)
    title = re.sub(r'[<>=]+\s*', '', title)
    # Strip RHSA advisory IDs
    title = _strip_rhsa(title)
    # Strip full version tokens including trailing letters (1.0.2p, 1.0.2zn)
    title = VERSION_FULL_PATTERN.sub("", title)
    # Strip CPU date patterns (January 2026 CPU)
    title = CPU_DATE_PATTERN.sub("", title)
    # Strip month-year patterns (May 2026), (July 2025) — used by KB patches
    title = MONTH_YEAR_PATTERN.sub("", title)
    # Strip KB number prefix
    title = KB_PATTERN.sub("", title)
    # Strip standalone (Unix) / (Unix prefix suffixes
    title = re.sub(r'\(Unix\s*\)?', '', title)
    # Strip CVE references
    title = re.sub(r'CVE-\d+-\d+', '', title)
    # Normalize vulnerability description variants
    title = re.sub(r'\bMultiple\s+Vulnerabilit\w*', '', title)
    title = re.sub(r'\bInformation\s+Disclosure\b', '', title)
    title = re.sub(r'\bSQL\s*Injection\b', '', title, flags=re.IGNORECASE)
    title = re.sub(r'\bSQLi\b', '', title)
    title = re.sub(r'\bVulnerability\b', '', title)
    # Collapse multiple spaces
    title = re.sub(r'\s+', ' ', title).strip()
    # Strip leading colon or dash
    title = re.sub(r'^[\s:.-]+', '', title)
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
