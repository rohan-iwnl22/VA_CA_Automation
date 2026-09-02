"""CA pipeline orchestration."""

from __future__ import annotations

import logging
import os
import re
from pathlib import Path

import pandas as pd

from ..excel_writer.chart_builder import build_ca_pie_chart
from ..excel_writer.data_writer import write_ca_data_rows, write_ca_report_header
from ..excel_writer.summary_builder import (
    build_ca_risk_summary,
    build_ca_scope_table,
    format_ca_summary_title,
    write_ca_risk_summary_table,
    write_ca_scope_table,
)
from ..excel_writer.template_cloner import clone_template, load_working_copy
from ..logging.pipeline_logger import PipelineLogger
from ..metadata.engagement_metadata import EngagementMetadata
from ..naming.filename_builder import build_filename, ensure_unique_path
from ..transform.filters import filter_ca_candidates

logger = logging.getLogger("va_ca_automation")


# =========================================================
# EXTRACT TITLE
# =========================================================

def _extract_title(description: str) -> str:
    """Extract clean title from Description field.

    Handles patterns like:
    - "Title: [FAILED]"
    - "Title: [WARNING]"
    - "Title: [WARNINGS]"
    """
    if pd.isna(description):
        return ""

    description = str(description).strip()

    match = re.match(
        r"^(.*?)\s*:\s*\[(?:FAILED|WARNING|WARNINGS)\]",
        description,
        re.IGNORECASE | re.DOTALL
    )

    if match:
        return match.group(1).strip().strip('"')

    # Fallback: first line
    first_line = description.splitlines()[0].strip()
    return first_line.strip('"')


# =========================================================
# EXTRACT DESCRIPTION
# =========================================================

def _extract_description(description: str) -> str:
    """Extract clean description from Description field.

    Extracts text between [FAILED]/[WARNING] marker and 'Solution:'.
    """
    if pd.isna(description):
        return ""

    description = str(description)

    start_marker = r"\[(?:FAILED|WARNING|WARNINGS)\]"
    end_marker = "Solution:"

    start_match = re.search(
        start_marker,
        description,
        re.IGNORECASE
    )

    if start_match:
        start_pos = start_match.end()

        end_match = re.search(
            re.escape(end_marker),
            description[start_pos:],
            re.IGNORECASE
        )

        if end_match:
            end_pos = start_pos + end_match.start()
            return description[start_pos:end_pos].strip()

        return description[start_pos:].strip()

    return description.strip()


# =========================================================
# COMBINE UNIQUE HOSTS
# =========================================================

def _combine_hosts(host_series: pd.Series) -> str:
    """Combine unique hosts into comma-separated string."""
    unique_hosts = []

    for host in host_series:
        if pd.isna(host):
            continue

        host = str(host).strip()

        if host and host not in unique_hosts:
            unique_hosts.append(host)

    return ", ".join(unique_hosts)


# =========================================================
# PROCESS CA DATA
# =========================================================

def _process_ca_data(ca_rows: pd.DataFrame) -> pd.DataFrame:
    """Process CA rows using the user's CA logic.

    Steps:
    1. Remove blank rows
    2. Remove blank Description rows
    3. Filter for FAILED/WARNING/WARNINGS
    4. Extract clean title from Description
    5. Extract clean description from Description
    6. Clean Host and Risk columns
    7. Deduplicate by Host + Title
    """
    df = ca_rows.copy()

    # Remove completely blank rows
    df = df.dropna(how="all").copy()
    logger.info("CA: Rows after removing blank rows: %d", len(df))

    # Remove blank Description rows
    df = df[
        df["Description"].notna()
        &
        df["Description"]
        .astype(str)
        .str.strip()
        .ne("")
    ].copy()
    logger.info("CA: Rows after removing blank descriptions: %d", len(df))

    # Normalize WARNINGS → WARNING
    df["Risk"] = df["Risk"].astype(str).str.strip().str.upper().replace("WARNINGS", "WARNING")

    # Keep only FAILED/WARNING
    df = df[df["Risk"].isin(["FAILED", "WARNING"])].copy()
    logger.info("CA: FAILED/WARNING findings: %d", len(df))

    # Reset index
    df = df.reset_index(drop=True)

    # Create clean title
    df["_Title"] = df["Description"].apply(_extract_title)

    # Create clean description
    df["_CleanDescription"] = df["Description"].apply(_extract_description)

    # Clean Host
    df["Host"] = df["Host"].astype(str).str.strip()

    # Clean Risk
    df["Risk"] = df["Risk"].astype(str).str.strip()

    # Remove duplicates (same Host + same Title)
    before_duplicates = len(df)
    df = df.drop_duplicates(
        subset=["Host", "_Title"],
        keep="first"
    ).copy()
    after_duplicates = len(df)
    logger.info("CA: Duplicates removed: %d", before_duplicates - after_duplicates)

    # Reset index
    df = df.reset_index(drop=True)

    return df


# =========================================================
# CREATE NORMAL CA REPORT
# =========================================================

def _create_normal_ca_report(df: pd.DataFrame) -> pd.DataFrame:
    """Create the normal CA report DataFrame."""
    ca_report = pd.DataFrame()

    ca_report["Sr.No."] = range(1, len(df) + 1)
    ca_report["Title"] = df["_Title"]
    ca_report["Host"] = df["Host"]
    ca_report["Description"] = df["_CleanDescription"]
    ca_report["Solution"] = df["Solution"]
    ca_report["Risk"] = df["Risk"]

    return ca_report


# =========================================================
# CREATE TEXTJOIN CA REPORT
# =========================================================

def _create_textjoin_ca_report(ca_report: pd.DataFrame) -> pd.DataFrame:
    """Create the TextJoin CA report DataFrame.

    Groups identical findings by Title+Description+Solution+Risk
    and combines all affected Hosts.
    """
    textjoined_report = (
        ca_report
        .groupby(
            ["Title", "Description", "Solution", "Risk"],
            dropna=False,
            sort=False
        )
        .agg({"Host": _combine_hosts})
        .reset_index()
    )

    # Reorder columns
    textjoined_report = textjoined_report[
        ["Title", "Host", "Description", "Solution", "Risk"]
    ]

    # Add new Sr.No.
    textjoined_report.insert(
        0,
        "Sr.No.",
        range(1, len(textjoined_report) + 1)
    )

    return textjoined_report


# =========================================================
# BUILD CA FILENAME
# =========================================================

def _build_ca_filename(metadata: EngagementMetadata) -> str:
    """Build the CA report filename.

    Pattern: Configuration_Audit_<Scope>_<Phase>_Audit_Report_<Client>_<Entity(s)>_<Year>_V<Major>.<Minor>.xlsx
    """
    return build_filename(metadata, report_type="Configuration_Audit")


# =========================================================
# SAVE REPORTS
# =========================================================

def _safe_save_workbook(wb, path: Path) -> None:
    """Save workbook, handling PermissionError."""
    try:
        wb.save(path)
    except PermissionError:
        logger.warning("File %s is locked. Attempting to remove and retry...", path)
        try:
            os.remove(path)
            wb.save(path)
        except OSError as e:
            logger.error("Cannot remove locked file: %s", e)
            new_path = path.parent / f"{path.stem}_locked{path.suffix}"
            logger.warning("Saving to alternative path: %s", new_path)
            wb.save(new_path)


# =========================================================
# RUN CA PIPELINE
# =========================================================

def run_ca_pipeline(
    raw_df: pd.DataFrame,
    metadata: EngagementMetadata,
    ca_template_path: Path | str,
    output_dir: Path | str,
    log_file: Path | None = None,
    generate_text_join: bool = True,
) -> Path:
    """Execute the full CA pipeline: filter -> extract -> dedup -> write to template -> save.

    Parameters
    ----------
    raw_df : pd.DataFrame
        Raw data DataFrame (already loaded from the raw file).
    metadata : EngagementMetadata
        Engagement metadata for populating the report.
    ca_template_path : Path
        Path to the pristine CA report template (.xlsx).
    output_dir : Path
        Directory where the final reports will be saved.
    log_file : Path, optional
        If provided, structured log entries will be written here.
    generate_text_join : bool, optional
        If True, also generate a separate TextJoin report file.

    Returns
    -------
    Path
        Path to the generated normal CA report file.
    """
    from ..ingestion.schema_validator import (
        classify_rows,
        normalize_whitespace_columns,
        validate_and_normalize_risk,
    )

    ca_template_path = Path(ca_template_path)
    output_dir = Path(output_dir)

    plogger = PipelineLogger(log_file=log_file)

    # 1. NORMALIZE
    raw_df = normalize_whitespace_columns(raw_df, ["Risk", "Host", "Name"])
    raw_df = validate_and_normalize_risk(raw_df, plogger)

    # 2. CLASSIFY
    va_rows, ca_rows, unknown_rows = classify_rows(raw_df)
    plogger.log_stage_count("ca_candidates_raw", len(ca_rows))

    # 3. FILTER (CA only: exclude PASSED)
    ca_filtered = filter_ca_candidates(ca_rows)
    plogger.log_stage_count("ca_candidates_after_filter", len(ca_filtered))

    if ca_filtered.empty:
        logger.warning("No CA findings to process after filtering.")
        return None

    # 4. PROCESS CA DATA
    ca_processed = _process_ca_data(ca_filtered)
    plogger.log_stage_count("final_ca_rows", len(ca_processed))

    if ca_processed.empty:
        logger.warning("No CA findings to process after deduplication.")
        return None

    # 5. CREATE NORMAL CA REPORT DATAFRAME
    ca_normal_report = _create_normal_ca_report(ca_processed)

    # 6. CREATE TEXTJOIN CA REPORT DATAFRAME
    ca_textjoin_report = _create_textjoin_ca_report(ca_normal_report)

    # 7. BUILD FILENAME
    ca_filename = _build_ca_filename(metadata)
    ca_output_path = output_dir / ca_filename
    ca_output_path = ensure_unique_path(ca_output_path)

    # 8. CREATE OUTPUT FOLDER
    os.makedirs(output_dir, exist_ok=True)

    # 9. CLONE CA TEMPLATE AND WRITE NORMAL REPORT
    working_path = clone_template(ca_template_path, ca_output_path)
    logger.info("CA working copy created: %s", working_path)

    wb = load_working_copy(working_path)
    try:
        ca_ws = wb["CA_Report"]

        # 10. WRITE HEADER METADATA
        write_ca_report_header(ca_ws, metadata)

        # 11. WRITE DATA ROWS (starting at row 14, skipping headers at row 13)
        write_ca_data_rows(ca_ws, ca_normal_report)

        # 12. WRITE SUMMARY SHEET
        summary_ws = wb["Summary"]
        format_ca_summary_title(summary_ws)
        scope_df = build_ca_scope_table(ca_normal_report, metadata)
        write_ca_scope_table(summary_ws, scope_df)
        risk_summary = build_ca_risk_summary(ca_normal_report)
        write_ca_risk_summary_table(summary_ws, risk_summary, start_row=15)
        build_ca_pie_chart(summary_ws, risk_summary, chart_anchor="I8", data_start_row=15)

        # 13. SAVE NORMAL CA REPORT
        _safe_save_workbook(wb, working_path)
        plogger.log_output_file(str(working_path))
        logger.info("CA Audit Report saved: %s (%d findings)", working_path, len(ca_normal_report))
    finally:
        wb.close()

    # 13. TEXTJOIN CA REPORT
    if generate_text_join:
        tj_stem = ca_output_path.stem + "_TextJoin"
        tj_path = ca_output_path.parent / f"{tj_stem}{ca_output_path.suffix}"
        tj_path = clone_template(ca_template_path, tj_path)

        wb_tj = load_working_copy(tj_path)
        try:
            ca_ws_tj = wb_tj["CA_Report"]
            write_ca_report_header(ca_ws_tj, metadata)
            write_ca_data_rows(ca_ws_tj, ca_textjoin_report)

            # Write Summary sheet for TextJoin report
            summary_ws_tj = wb_tj["Summary"]
            format_ca_summary_title(summary_ws_tj)
            scope_df_tj = build_ca_scope_table(ca_textjoin_report, metadata)
            write_ca_scope_table(summary_ws_tj, scope_df_tj)
            risk_summary_tj = build_ca_risk_summary(ca_textjoin_report)
            write_ca_risk_summary_table(summary_ws_tj, risk_summary_tj, start_row=15)
            build_ca_pie_chart(summary_ws_tj, risk_summary_tj, chart_anchor="I8", data_start_row=15)

            _safe_save_workbook(wb_tj, tj_path)
            plogger.log_output_file(str(tj_path))
            logger.info("CA TextJoin Report saved: %s (%d unique findings)", tj_path, len(ca_textjoin_report))
        finally:
            wb_tj.close()

    plogger.log_summary()
    plogger.flush()

    return ca_output_path
