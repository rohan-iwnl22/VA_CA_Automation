"""Tests for naming and metadata modules."""

from datetime import date

import pytest

from va_ca_automation.metadata.engagement_metadata import EngagementMetadata, HostMetadata
from va_ca_automation.naming.filename_builder import build_filename, ensure_unique_path


class TestEngagementMetadata:
    def test_default_values(self):
        meta = EngagementMetadata(
            client_name="Test Client",
            security_tester="Tester",
            reviewed_by="Reviewer",
            report_date=date(2026, 1, 1),
            report_version=1.0,
        )
        assert meta.scanner_name == "Nessus "
        assert meta.scanner_version == "10.11.4"
        assert meta.scope_label == "Server"
        assert meta.phase_label == "First"

    def test_host_metadata_lookup(self):
        meta = EngagementMetadata(
            client_name="Test",
            security_tester="T",
            reviewed_by="R",
            report_date=date(2026, 1, 1),
            report_version=1.0,
            host_metadata={
                "10.0.0.1": HostMetadata("10.0.0.1", "Authenticated", "Server"),
            },
        )
        assert meta.get_host_scan_type("10.0.0.1") == "Authenticated"
        assert meta.get_host_device_type("10.0.0.1") == "Server"
        assert meta.get_host_scan_type("10.0.0.99") == ""


class TestFilenameBuilder:
    def test_builds_correct_filename(self):
        meta = EngagementMetadata(
            client_name="Trust Investment Advisors Private Limited",
            security_tester="Tester",
            reviewed_by="Reviewer",
            report_date=date(2026, 7, 9),
            report_version=1.2,
            scope_label="Server",
            phase_label="First",
            entity_codes=["TSS", "SCPL"],
        )
        filename = build_filename(meta)
        assert filename.startswith("VA_Server_First_Audit_Report_")
        assert "Trust_Investment_Advisors_Private_Limited" in filename
        assert "2026" in filename
        assert "V1_2.xlsx" in filename

    def test_no_entity_codes(self):
        meta = EngagementMetadata(
            client_name="Simple Client",
            security_tester="T",
            reviewed_by="R",
            report_date=date(2026, 1, 1),
            report_version=1.0,
        )
        filename = build_filename(meta)
        assert filename.endswith(".xlsx")
        assert "Audit_Report" in filename

    def test_sanitizes_special_characters(self):
        meta = EngagementMetadata(
            client_name="Client & Co. / Ltd.",
            security_tester="T",
            reviewed_by="R",
            report_date=date(2026, 1, 1),
            report_version=1.0,
        )
        filename = build_filename(meta)
        assert "&" not in filename
        assert "/" not in filename
        assert "." not in filename.replace(".xlsx", "")


class TestEnsureUniquePath:
    def test_returns_same_if_not_exists(self, tmp_path):
        path = tmp_path / "report.xlsx"
        result = ensure_unique_path(path)
        assert result == path

    def test_appends_counter(self, tmp_path):
        path = tmp_path / "report.xlsx"
        path.touch()
        result = ensure_unique_path(path)
        assert result.name == "report_1.xlsx"

    def test_increments_counter(self, tmp_path):
        path = tmp_path / "report.xlsx"
        path.touch()
        (tmp_path / "report_1.xlsx").touch()
        result = ensure_unique_path(path)
        assert result.name == "report_2.xlsx"
