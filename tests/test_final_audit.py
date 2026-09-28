"""Tests for final-audit data writer helpers."""

import io

import openpyxl
import pandas as pd
import pytest

from va_ca_automation.excel_writer.data_writer import (
    FINAL_AUDIT_COLUMNS,
    HEADER_ROW,
    RETEST_CLOSED_FILL,
    RETEST_OPEN_FILL,
    RETEST_FONT_WHITE,
    build_rescan_lookup,
    extract_va_report_data,
    write_final_audit_rows,
)


def _make_va_workbook(rows: list[list]) -> openpyxl.Workbook:
    """Create a minimal VA Report workbook with header row 13 and data rows 14+."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "VA Report"
    headers = [
        "Sr. no", "Vulnerbility Title", "Description", "Risk",
        "Host", "Port", "Recommendation ", "Reference", "CVE",
    ]
    for i, h in enumerate(headers, 1):
        ws.cell(row=HEADER_ROW, column=i, value=h)
    for r, row_data in enumerate(rows, 14):
        for c, val in enumerate(row_data, 1):
            ws.cell(row=r, column=c, value=val)
    return wb


class TestExtractVaReportData:
    def test_extracts_data_rows(self):
        wb = _make_va_workbook([
            [1, "Weak TLS", "Desc", "High", "10.0.0.1", "443", "Fix", "ref", "CVE-1"],
            [2, "SQL Injection", "Desc", "Critical", "10.0.0.2", "80", "Fix", "ref", "CVE-2"],
        ])
        ws = wb["VA Report"]
        data = extract_va_report_data(ws)
        assert len(data) == 2
        assert data[0]["Vulnerbility Title"] == "Weak TLS"
        assert data[1]["Host"] == "10.0.0.2"
        wb.close()

    def test_skips_empty_rows(self):
        wb = _make_va_workbook([
            [1, "Weak TLS", "Desc", "High", "10.0.0.1", "443", "Fix", "ref", "CVE-1"],
        ])
        ws = wb["VA Report"]
        # Add an empty row at 15
        data = extract_va_report_data(ws)
        assert len(data) == 1
        wb.close()

    def test_empty_workbook(self):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "VA Report"
        data = extract_va_report_data(ws)
        assert data == []
        wb.close()


class TestBuildRescanLookup:
    def test_builds_lookup_set(self):
        rescan_data = [
            {"Vulnerbility Title": "Weak TLS", "Host": "10.0.0.1"},
            {"Vulnerbility Title": "SQL Injection", "Host": "10.0.0.2"},
        ]
        lookup = build_rescan_lookup(rescan_data)
        assert ("weak tls", "10.0.0.1") in lookup
        assert ("sql injection", "10.0.0.2") in lookup

    def test_normalizes_whitespace_and_case(self):
        rescan_data = [
            {"Vulnerbility Title": "  Weak TLS  ", "Host": " 10.0.0.1 "},
        ]
        lookup = build_rescan_lookup(rescan_data)
        assert ("weak tls", "10.0.0.1") in lookup

    def test_skips_none_values(self):
        rescan_data = [
            {"Vulnerbility Title": None, "Host": "10.0.0.1"},
            {"Vulnerbility Title": "Weak TLS", "Host": None},
        ]
        lookup = build_rescan_lookup(rescan_data)
        assert len(lookup) == 0

    def test_empty_rescan(self):
        lookup = build_rescan_lookup([])
        assert len(lookup) == 0

    def test_deduplicates(self):
        rescan_data = [
            {"Vulnerbility Title": "Weak TLS", "Host": "10.0.0.1"},
            {"Vulnerbility Title": "Weak TLS", "Host": "10.0.0.1"},
        ]
        lookup = build_rescan_lookup(rescan_data)
        assert len(lookup) == 1


class TestWriteFinalAuditRows:
    def test_open_status_when_match(self):
        wb = _make_va_workbook([
            [1, "Weak TLS", "Desc", "High", "10.0.0.1", "443", "Fix", "ref", "CVE-1"],
        ])
        ws = wb["VA Report"]
        rescan_lookup = {("weak tls", "10.0.0.1")}
        write_final_audit_rows(ws, [
            {"Vulnerbility Title": "Weak TLS", "Host": "10.0.0.1", "Risk": "High"},
        ], rescan_lookup)
        cell = ws.cell(row=14, column=10)
        assert cell.value == "OPEN"
        assert cell.fill == RETEST_OPEN_FILL
        wb.close()

    def test_closed_status_when_no_match(self):
        wb = _make_va_workbook([
            [1, "Weak TLS", "Desc", "High", "10.0.0.1", "443", "Fix", "ref", "CVE-1"],
        ])
        ws = wb["VA Report"]
        rescan_lookup = set()
        write_final_audit_rows(ws, [
            {"Vulnerbility Title": "Weak TLS", "Host": "10.0.0.1", "Risk": "High"},
        ], rescan_lookup)
        cell = ws.cell(row=14, column=10)
        assert cell.value == "CLOSED"
        assert cell.fill == RETEST_CLOSED_FILL
        wb.close()

    def test_closed_when_title_is_none(self):
        wb = _make_va_workbook([
            [1, None, "Desc", "High", "10.0.0.1", "443", "Fix", "ref", "CVE-1"],
        ])
        ws = wb["VA Report"]
        rescan_lookup = {(None, "10.0.0.1")}
        write_final_audit_rows(ws, [
            {"Vulnerbility Title": None, "Host": "10.0.0.1", "Risk": "High"},
        ], rescan_lookup)
        cell = ws.cell(row=14, column=10)
        assert cell.value == "CLOSED"
        wb.close()

    def test_white_bold_font_on_status(self):
        wb = _make_va_workbook([
            [1, "Weak TLS", "Desc", "High", "10.0.0.1", "443", "Fix", "ref", "CVE-1"],
        ])
        ws = wb["VA Report"]
        rescan_lookup = {("weak tls", "10.0.0.1")}
        write_final_audit_rows(ws, [
            {"Vulnerbility Title": "Weak TLS", "Host": "10.0.0.1", "Risk": "High"},
        ], rescan_lookup)
        cell = ws.cell(row=14, column=10)
        assert cell.font.bold is True
        assert cell.font.color.rgb == "00FFFFFF"
        wb.close()

    def test_multiple_rows(self):
        wb = _make_va_workbook([
            [1, "Weak TLS", "Desc", "High", "10.0.0.1", "443", "Fix", "ref", "CVE-1"],
            [2, "SQL Injection", "Desc", "Critical", "10.0.0.2", "80", "Fix", "ref", "CVE-2"],
        ])
        ws = wb["VA Report"]
        rescan_lookup = {("weak tls", "10.0.0.1")}
        data = [
            {"Vulnerbility Title": "Weak TLS", "Host": "10.0.0.1", "Risk": "High"},
            {"Vulnerbility Title": "SQL Injection", "Host": "10.0.0.2", "Risk": "Critical"},
        ]
        write_final_audit_rows(ws, data, rescan_lookup)
        assert ws.cell(row=14, column=10).value == "OPEN"
        assert ws.cell(row=15, column=10).value == "CLOSED"
        wb.close()

    def test_case_insensitive_matching_via_build_lookup(self):
        """build_rescan_lookup normalises to lowercase, write_final_audit_rows uses lowercase keys."""
        rescan_data = [
            {"Vulnerbility Title": "  WEAK TLS  ", "Host": " 10.0.0.1 "},
        ]
        lookup = build_rescan_lookup(rescan_data)
        assert ("weak tls", "10.0.0.1") in lookup

        wb = _make_va_workbook([
            [1, "Weak TLS", "Desc", "High", "10.0.0.1", "443", "Fix", "ref", "CVE-1"],
        ])
        ws = wb["VA Report"]
        write_final_audit_rows(ws, [
            {"Vulnerbility Title": "Weak TLS", "Host": "10.0.0.1", "Risk": "High"},
        ], lookup)
        cell = ws.cell(row=14, column=10)
        assert cell.value == "OPEN"
        wb.close()
