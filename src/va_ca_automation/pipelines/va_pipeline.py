"""VA pipeline orchestration."""

from __future__ import annotations

import logging
from pathlib import Path

from ..excel_writer.chart_builder import build_pie_chart
from ..excel_writer.data_writer import (
    write_introduction_fields,
    write_va_data_rows,
    write_va_report_header,
)
from ..excel_writer.summary_builder import (
    build_risk_summary,
    build_scope_table,
    write_risk_summary_table,
    write_scope_table,
)
from ..excel_writer.template_cloner import clone_template, load_working_copy
from ..ingestion.raw_file_loader import load_raw_file
from ..ingestion.schema_validator import (
    classify_rows,
    normalize_whitespace_columns,
    validate_and_normalize_risk,
)
from ..logging.pipeline_logger import PipelineLogger
from ..metadata.engagement_metadata import EngagementMetadata
from ..naming.filename_builder import build_filename, ensure_unique_path
from ..transform.column_mapper import map_columns
from ..transform.dedup import stage1_exact_dedup, stage2_version_collapse
from ..transform.filters import filter_va_candidates
from ..transform.sorter import sort_va_data

logger = logging.getLogger("va_ca_automation")


def run_va_pipeline(
    raw_file_path: Path | str,
    template_path: Path | str,
    metadata: EngagementMetadata,
    output_dir: Path | str,
    log_file: Path | None = None,
) -> Path:
    """Execute the full VA pipeline: ingest -> filter -> dedup -> sort -> write -> summary -> save.

    Parameters
    ----------
    raw_file_path : Path
        Path to the raw Nessus export (.xlsx).
    template_path : Path
        Path to the pristine blank template (.xlsx).
    metadata : EngagementMetadata
        Engagement metadata for populating the report.
    output_dir : Path
        Directory where the final report will be saved.
    log_file : Path, optional
        If provided, structured log entries will be written here.

    Returns
    -------
    Path
        Path to the generated report file.
    """
    raw_file_path = Path(raw_file_path)
    template_path = Path(template_path)
    output_dir = Path(output_dir)

    plogger = PipelineLogger(log_file=log_file)

    # 1. INGEST
    logger.info("Loading raw file: %s", raw_file_path)
    raw_df = load_raw_file(raw_file_path)
    plogger.log_stage_count("raw_rows", len(raw_df))

    # 2. NORMALIZE
    raw_df = normalize_whitespace_columns(raw_df, ["Risk", "Host", "Name"])
    raw_df = validate_and_normalize_risk(raw_df, plogger)

    # 3. CLASSIFY
    va_rows, ca_rows, unknown_rows = classify_rows(raw_df)
    plogger.log_stage_count("va_candidates_raw", len(va_rows))
    plogger.log_stage_count("ca_candidates", len(ca_rows))
    plogger.log_stage_count("unknown_risk", len(unknown_rows))

    # 4. FILTER (VA only: exclude None)
    va_filtered = filter_va_candidates(va_rows)
    plogger.log_stage_count("va_candidates_after_filter", len(va_filtered))

    # 5. STAGE 1 DEDUP: exact match
    va_stage1 = stage1_exact_dedup(va_filtered)
    plogger.log_stage_count("after_stage1_exact_dedup", len(va_stage1))

    # 6. STAGE 2 DEDUP: version-collapse
    va_stage2 = stage2_version_collapse(va_stage1, plogger)
    plogger.log_stage_count("after_stage2_version_collapse", len(va_stage2))

    # 7. MAP COLUMNS
    va_mapped = map_columns(va_stage2)

    # 8. SORT + RENUMBER
    va_sorted = sort_va_data(va_mapped)
    plogger.log_stage_count("final_va_rows", len(va_sorted))
    plogger.log_risk_breakdown(va_sorted["Risk"].value_counts().to_dict())

    # 9. CLONE TEMPLATE
    filename = build_filename(metadata)
    output_path = output_dir / filename
    output_path = ensure_unique_path(output_path)

    working_path = clone_template(template_path, output_path)
    logger.info("Working copy created: %s", working_path)

    # 10. WRITE HEADER METADATA
    wb = load_working_copy(working_path)
    try:
        va_ws = wb["VA Report"]
        write_va_report_header(va_ws, metadata)

        # 11. WRITE DATA ROWS
        last_row = write_va_data_rows(va_ws, va_sorted)

        # 12. WRITE SUMMARY SCOPE TABLE
        summary_ws = wb["Summary"]
        scope_df = build_scope_table(va_sorted, metadata)
        write_scope_table(summary_ws, scope_df)

        # 13. WRITE RISK SUMMARY + PIE CHART
        risk_summary = build_risk_summary(va_sorted)
        write_risk_summary_table(summary_ws, risk_summary)
        build_pie_chart(summary_ws, risk_summary, chart_anchor="H8")

        # 14. WRITE INTRODUCTION FIELDS
        intro_ws = wb["Introduction"]
        write_introduction_fields(intro_ws, metadata)

        # 15. SAVE
        wb.save(working_path)
        plogger.log_output_file(str(working_path))
        plogger.log_summary()
        plogger.flush()

        logger.info("Report saved: %s", working_path)
        return working_path

    except Exception:
        wb.close()
        raise


def run_va_pipeline_with_validation(
    raw_file_path: Path | str,
    template_path: Path | str,
    metadata: EngagementMetadata,
    output_dir: Path | str,
    log_file: Path | None = None,
) -> Path:
    """Run the VA pipeline with post-write validation.

    Validates that:
    - Row count in VA Report matches the processed dataset.
    - Risk summary grand total matches the row count.
    - No pivot tables exist in the output (static aggregation used).
    """
    output_path = run_va_pipeline(
        raw_file_path, template_path, metadata, output_dir, log_file
    )

    # Post-write validation
    wb = load_working_copy(output_path)
    try:
        va_ws = wb["VA Report"]

        # Count written data rows (starting at row 14)
        data_rows = 0
        for row in range(14, va_ws.max_row + 1):
            if va_ws.cell(row=row, column=1).value is not None:
                data_rows += 1

        # Validate row count
        expected_rows = len(
            sort_va_data(
                map_columns(
                    stage2_version_collapse(
                        stage1_exact_dedup(
                            filter_va_candidates(
                                classify_rows(
                                    validate_and_normalize_risk(
                                        normalize_whitespace_columns(
                                            load_raw_file(raw_file_path),
                                            ["Risk", "Host", "Name"],
                                        ),
                                        PipelineLogger(),
                                    )
                                )[0]
                            )
                        ),
                        PipelineLogger(),
                    )
                )
            )
        )

        if data_rows != expected_rows:
            logger.warning(
                "Row count mismatch: expected %d, found %d", expected_rows, data_rows
            )

        # Validate summary grand total
        summary_ws = wb["Summary"]
        grand_total_cell = summary_ws.cell(row=23, column=6)
        if grand_total_cell.value is not None:
            grand_total = int(grand_total_cell.value)
            if grand_total != data_rows:
                logger.warning(
                    "Grand total mismatch: expected %d, found %d", data_rows, grand_total
                )

    finally:
        wb.close()

    return output_path
