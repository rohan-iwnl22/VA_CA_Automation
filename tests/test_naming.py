"""Tests for naming and metadata modules."""

from datetime import date

import pytest

from va_ca_automation.metadata.engagement_metadata import EngagementMetadata, HostMetadata
from va_ca_automation.naming.filename_builder import (
    NamingConventionError,
    build_filename,
    build_textjoin_filename,
    ensure_unique_path,
    validate_filename,
)


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
    def _meta(self, **kwargs) -> EngagementMetadata:
        defaults = dict(
            client_name="Trust Investment Advisors Private Limited",
            security_tester="Tester",
            reviewed_by="Reviewer",
            report_date=date(2026, 7, 9),
            report_version="1.2",
            scope_label="Server",
            report_type="First",
        )
        defaults.update(kwargs)
        return EngagementMetadata(**defaults)

    def test_va_filename_pattern(self):
        filename = build_filename(self._meta(report_type="Final"))
        assert filename == (
            "VA_Server_Final_Audit_Report_"
            "Trust_Investment_Advisors_Private_Limited_SCPL_2026_V1.2.xlsx"
        )

    def test_ca_filename_pattern(self):
        filename = build_filename(
            self._meta(report_type="First"), report_type="Configuration_Audit"
        )
        assert filename == (
            "Configuration_Audit_Server_First_Audit_Report_"
            "Trust_Investment_Advisors_Private_Limited_SCPL_2026_V1.2.xlsx"
        )

    def test_device_type_input_takes_precedence(self):
        filename = build_filename(self._meta(default_device_type="Firewall"))
        assert filename.startswith("VA_Firewall_First_Audit_Report_")

    def test_version_keeps_dot(self):
        filename = build_filename(self._meta(report_version="1.0"))
        assert filename.endswith("_V1.0.xlsx")

    def test_version_accepts_leading_v(self):
        filename = build_filename(self._meta(report_version="v1.0"))
        assert filename.endswith("_V1.0.xlsx")

    def test_scpl_always_present_and_entity_codes_ignored(self):
        filename = build_filename(self._meta(entity_codes=["TSS", "XYZ"]))
        assert "_SCPL_" in filename
        assert "TSS" not in filename
        assert "XYZ" not in filename

    def test_invalid_report_type_raises(self):
        with pytest.raises(NamingConventionError):
            build_filename(self._meta(report_type="Retest"))

    def test_invalid_version_raises(self):
        with pytest.raises(NamingConventionError):
            build_filename(self._meta(report_version="final"))

    def test_empty_client_raises(self):
        with pytest.raises(NamingConventionError):
            build_filename(self._meta(client_name=""))

    def test_no_device_token_raises(self):
        with pytest.raises(NamingConventionError):
            build_filename(self._meta(scope_label="", default_device_type=""))

    def test_invalid_report_type_argument_raises(self):
        with pytest.raises(NamingConventionError):
            build_filename(self._meta(), report_type="Word")

    def test_sanitizes_special_characters(self):
        filename = build_filename(self._meta(client_name="Client & Co. / Ltd."))
        assert "&" not in filename
        assert "/" not in filename
        assert filename == "VA_Server_First_Audit_Report_Client_Co_Ltd_SCPL_2026_V1.2.xlsx"


class TestBuildTextjoinFilename:
    def test_inserts_before_extension(self):
        name = "VA_Server_Final_Audit_Report_Client_SCPL_2026_V1.0.xlsx"
        assert build_textjoin_filename(name) == (
            "VA_Server_Final_Audit_Report_Client_SCPL_2026_V1.0.textjoin.xlsx"
        )

    def test_idempotent(self):
        name = "VA_Server_Final_Audit_Report_Client_SCPL_2026_V1.0.textjoin.xlsx"
        assert build_textjoin_filename(name) == name


class TestValidateFilename:
    def test_accepts_normal_and_textjoin(self):
        assert validate_filename("VA_Server_First_Audit_Report_Client_SCPL_2026_V1.0.xlsx")
        assert validate_filename(
            "Configuration_Audit_Server_Final_Audit_Report_Client_SCPL_2026_V1.0.textjoin.xlsx"
        )

    def test_rejects_non_conforming(self):
        assert not validate_filename("report.xlsx")
        assert not validate_filename("VA_Server_Retest_Report_Client_SCPL_2026_V1.0.xlsx")
        assert not validate_filename("VA_Server_First_Audit_Report_Client_2026_V1.0.xlsx")


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
