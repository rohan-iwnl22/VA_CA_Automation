"""Merge IPs sharing the same vulnerability into a single row."""

from __future__ import annotations

import pandas as pd


def merge_ips_by_vulnerability(df: pd.DataFrame) -> pd.DataFrame:
    """Group rows by Vulnerability Title and merge all Host IPs into one comma-separated string.

    For each unique vulnerability, this function:
    - Collects all unique Host IPs and joins them with ", "
    - Keeps the first non-null value for Description, Risk, Port,
      Recommendation, Reference, and CVE columns
    - Returns one row per unique vulnerability with merged IPs

    Parameters
    ----------
    df : pd.DataFrame
        DataFrame with columns including 'Vulnerbility Title', 'Host',
        and optionally 'Description', 'Risk', 'Port', 'Recommendation ',
        'Reference', 'CVE'.

    Returns
    -------
    pd.DataFrame
        Grouped DataFrame with one row per vulnerability and merged IPs.
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
            h_str = str(h).strip()
            if h_str and h_str not in seen:
                seen.append(h_str)
        return ", ".join(seen) if seen else ""

    agg_rules = {}
    for col in join_columns:
        if col in df.columns:
            agg_rules[col] = _first_non_null

    if "Host" in df.columns:
        agg_rules["Host"] = _unique_hosts

    grouped = df.groupby("Vulnerbility Title", sort=False).agg(agg_rules)
    grouped = grouped.reset_index()

    return grouped


def export_merged_ips(df: pd.DataFrame, output_path: str | None = None) -> pd.DataFrame:
    """Process raw VA data and export a merged-IPs report.

    Runs the full pipeline steps (filter, dedup, map, merge) and writes
    the result to an Excel file with one sheet showing each vulnerability
    and its comma-separated list of affected IPs.

    Parameters
    ----------
    df : pd.DataFrame
        Raw or partially processed DataFrame with 'Name', 'Host', 'Risk',
        'Description', 'Solution', 'See Also', 'CVE', 'Port' columns.
    output_path : str, optional
        Path to save the output Excel file. If None, only the DataFrame
        is returned.

    Returns
    -------
    pd.DataFrame
        Merged DataFrame with columns: Vulnerbility Title, Host, Risk,
        Description, Port, Recommendation, Reference, CVE.
    """
    from ..ingestion.schema_validator import (
        classify_rows,
        normalize_whitespace_columns,
        validate_and_normalize_risk,
    )
    from ..transform.column_mapper import map_columns
    from ..transform.dedup import stage1_exact_dedup, stage1b_name_host_dedup, stage2_version_collapse
    from ..transform.filters import filter_va_candidates
    from ..logging.pipeline_logger import PipelineLogger

    df = normalize_whitespace_columns(df, ["Risk", "Host", "Name"])
    df = validate_and_normalize_risk(df, PipelineLogger())

    va_rows, _, _ = classify_rows(df)
    va_filtered = filter_va_candidates(va_rows)
    va_stage1 = stage1_exact_dedup(va_filtered)
    va_stage1b = stage1b_name_host_dedup(va_stage1)
    va_stage2 = stage2_version_collapse(va_stage1b, PipelineLogger())
    va_mapped = map_columns(va_stage2)

    merged = merge_ips_by_vulnerability(va_mapped)

    if output_path:
        merged.to_excel(output_path, index=False, sheet_name="Merged IPs")

    return merged
