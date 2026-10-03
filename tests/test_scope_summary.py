"""Summary "List of IPs in scope" must include hosts dropped from the report sheets."""

from __future__ import annotations

from datetime import date
from pathlib import Path

import pandas as pd
from openpyxl import Workbook, load_workbook

from va_ca_automation.excel_writer.data_writer import CA_TEMPLATE_COLUMNS, TEMPLATE_COLUMNS
from va_ca_automation.ingestion.raw_file_loader import EXPECTED_COLUMNS
from va_ca_automation.metadata.engagement_metadata import EngagementMetadata
from va_ca_automation.naming.filename_builder import build_textjoin_filename
from va_ca_automation.pipelines.ca_pipeline import run_ca_pipeline
from va_ca_automation.pipelines.va_pipeline import run_va_pipeline

ALL_HOSTS = ["10.0.0.1", "10.0.0.2", "10.0.0.3", "10.0.0.4"]


def _make_template(path: Path, report_sheet: str, header_columns: list[str]) -> Path:
    """Create a minimal template with the sheets the pipelines expect.

    A tiny stand-in for the multi-MB production templates keeps this test fast
    while still exercising the full pipeline.
    """
    wb = Workbook()
    report_ws = wb.active
    report_ws.title = report_sheet
    for col, name in enumerate(header_columns, start=1):
        report_ws.cell(row=13, column=col, value=name)

    wb.create_sheet("Summary")
    wb.create_sheet("Introduction")
    wb.save(path)
    wb.close()
    return path


def _metadata() -> EngagementMetadata:
    return EngagementMetadata(
        client_name="Test Client",
        security_tester="Tester",
        reviewed_by="Reviewer",
        report_date=date(2026, 1, 15),
        report_version="1.0",
        default_device_type="Server",
    )


def _raw_row(**overrides) -> dict:
    """One row with the full 17-column Nessus schema expected by load_raw_file."""
    row = {col: "" for col in EXPECTED_COLUMNS}
    row.update(
        {
            "Plugin ID": "100000",
            "Risk": "Critical",
            "Host": "10.0.0.1",
            "Protocol": "tcp",
            "Port": "443",
            "Name": "Vuln A",
            "Description": "desc",
            "Solution": "fix",
        }
    )
    row.update(overrides)
    return row


def _write_raw_file(tmp_path: Path) -> Path:
    """Raw upload with hosts that filtering removes from the report sheets.

    - 10.0.0.1: reportable VA + CA finding
    - 10.0.0.2: only Risk 'None' -> dropped from VA Report
    - 10.0.0.3: only Risk 'PASSED' -> dropped from CA_Report
    - 10.0.0.4: reportable VA + CA finding
    """
    raw = pd.DataFrame(
        [
            _raw_row(**{"Risk": "Critical", "Host": "10.0.0.1", "Name": "Vuln A"}),
            _raw_row(**{"Risk": "None", "Host": "10.0.0.2", "Name": "Info finding"}),
            _raw_row(
                **{
                    "Risk": "PASSED",
                    "Host": "10.0.0.3",
                    "Name": "Baseline check",
                    "Description": "Baseline check: [PASSED] All good Solution: No action",
                    "Solution": "No action",
                }
            ),
            _raw_row(
                **{
                    "Risk": "FAILED",
                    "Host": "10.0.0.4",
                    "Name": "Weak cipher",
                    "Description": (
                        "Weak cipher suite: [FAILED] Cipher is weak Solution: Disable it"
                    ),
                    "Solution": "Disable it",
                }
            ),
        ]
    )
    raw_path = tmp_path / "raw.xlsx"
    raw.to_excel(raw_path, index=False, sheet_name="RAW File")
    return raw_path


def _read_scope_hosts(path: Path, sheet: str = "Summary") -> list[str]:
    wb = load_workbook(path, data_only=True)
    try:
        ws = wb[sheet]
        hosts = []
        row = 8
        while row < ws.max_row + 1:
            value = ws.cell(row=row, column=2).value
            if value is None or str(value).strip() == "":
                break
            hosts.append(str(value).strip())
            row += 1
        return hosts
    finally:
        wb.close()


def _read_report_hosts(path: Path, sheet: str, host_col_index: int) -> set[str]:
    wb = load_workbook(path, data_only=True)
    try:
        ws = wb[sheet]
        hosts = set()
        for row in ws.iter_rows(min_row=14, values_only=True):
            if row[0] is None:
                continue
            hosts.add(str(row[host_col_index]).strip())
        return hosts
    finally:
        wb.close()


def test_va_scope_includes_host_dropped_from_report(tmp_path):
    raw_path = _write_raw_file(tmp_path)
    template_path = _make_template(tmp_path / "va_template.xlsx", "VA Report", TEMPLATE_COLUMNS)
    va_path = run_va_pipeline(
        raw_file_path=raw_path,
        template_path=template_path,
        metadata=_metadata(),
        output_dir=tmp_path,
        generate_text_join=True,
    )

    report_hosts = _read_report_hosts(va_path, "VA Report", host_col_index=4)
    assert "10.0.0.2" not in report_hosts

    assert _read_scope_hosts(va_path) == ALL_HOSTS

    va_tj_path = va_path.with_name(build_textjoin_filename(va_path.name))
    assert _read_scope_hosts(va_tj_path) == ALL_HOSTS


def test_ca_scope_includes_host_without_findings(tmp_path):
    raw_path = _write_raw_file(tmp_path)
    template_path = _make_template(
        tmp_path / "ca_template.xlsx", "CA_Report", CA_TEMPLATE_COLUMNS
    )
    ca_path = run_ca_pipeline(
        raw_df=pd.read_excel(raw_path),
        metadata=_metadata(),
        ca_template_path=template_path,
        output_dir=tmp_path,
        generate_text_join=True,
    )

    assert ca_path is not None
    report_hosts = _read_report_hosts(ca_path, "CA_Report", host_col_index=2)
    assert "10.0.0.3" not in report_hosts

    assert _read_scope_hosts(ca_path) == ALL_HOSTS

    ca_tj_path = ca_path.with_name(build_textjoin_filename(ca_path.name))
    assert _read_scope_hosts(ca_tj_path) == ALL_HOSTS
