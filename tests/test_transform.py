"""Tests for the transform layer."""

import pandas as pd
import pytest

from va_ca_automation.logging.pipeline_logger import PipelineLogger
from va_ca_automation.transform.column_mapper import RAW_TO_TEMPLATE, map_columns
from va_ca_automation.transform.dedup import (
    _extract_cpu_date,
    _extract_identifier,
    _extract_rhsa,
    _extract_version,
    _make_base_title,
    _version_to_tuple,
    stage1_exact_dedup,
    stage1b_name_host_dedup,
    stage2_version_collapse,
)
from va_ca_automation.transform.filters import VA_EXCLUDE_RISKS, filter_va_candidates
from va_ca_automation.transform.sorter import RISK_WEIGHTS, sort_va_data
from va_ca_automation.transform.text_join import text_join_hosts


@pytest.fixture
def va_df():
    """Create a VA-type DataFrame for testing."""
    return pd.DataFrame(
        {
            "Risk": ["Critical", "High", "Medium", "Low", "Critical", "High"],
            "Host": ["10.0.0.1", "10.0.0.1", "10.0.0.2", "10.0.0.2", "10.0.0.1", "10.0.0.2"],
            "Name": ["Vuln A", "Vuln B", "Vuln C", "Vuln D", "Vuln E", "Vuln F"],
            "Description": ["Desc A", "Desc B", "Desc C", "Desc D", "Desc E", "Desc F"],
            "Solution": ["Fix A", "Fix B", "Fix C", "Fix D", "Fix E", "Fix F"],
            "See Also": ["http://a.com", "", "http://c.com", "", "", "http://f.com"],
            "CVE": ["CVE-2024-0001", "", "CVE-2024-0003", "", "", ""],
            "Port": ["443", "80", "0", "22", "443", "80"],
        }
    )


class TestFilterVaCandidates:
    def test_excludes_none(self):
        df = pd.DataFrame({"Risk": ["Critical", "None", "High", "None", "Medium"]})
        result = filter_va_candidates(df)
        assert len(result) == 3
        assert "None" not in result["Risk"].values

    def test_keeps_all_va_severities(self):
        df = pd.DataFrame({"Risk": ["Critical", "High", "Medium", "Low"]})
        result = filter_va_candidates(df)
        assert len(result) == 4


class TestStage1ExactDedup:
    def test_removes_exact_duplicates(self):
        df = pd.DataFrame(
            {
                "Name": ["Vuln A", "Vuln A", "Vuln B"],
                "Description": ["Desc A", "Desc A", "Desc B"],
                "Risk": ["Critical", "Critical", "High"],
                "Host": ["10.0.0.1", "10.0.0.1", "10.0.0.1"],
                "Port": ["443", "443", "80"],
            }
        )
        result = stage1_exact_dedup(df)
        assert len(result) == 2

    def test_keeps_different_hosts(self):
        df = pd.DataFrame(
            {
                "Name": ["Vuln A", "Vuln A"],
                "Description": ["Desc A", "Desc A"],
                "Risk": ["Critical", "Critical"],
                "Host": ["10.0.0.1", "10.0.0.2"],
                "Port": ["443", "443"],
            }
        )
        result = stage1_exact_dedup(df)
        assert len(result) == 2

    def test_keeps_first_occurrence(self):
        df = pd.DataFrame(
            {
                "Name": ["Vuln A", "Vuln A"],
                "Description": ["Desc A", "Desc A"],
                "Risk": ["Critical", "Critical"],
                "Host": ["10.0.0.1", "10.0.0.1"],
                "Port": ["443", "80"],
            }
        )
        result = stage1_exact_dedup(df)
        assert len(result) == 1
        assert result.iloc[0]["Port"] == "443"


class TestStage1bNameHostDedup:
    def test_collapses_same_name_different_desc(self):
        df = pd.DataFrame(
            {
                "Name": ["Vuln A", "Vuln A", "Vuln A"],
                "Description": ["Desc 1", "Desc 2", "Desc 3"],
                "Risk": ["Critical", "Critical", "Critical"],
                "Host": ["10.0.0.1", "10.0.0.1", "10.0.0.1"],
                "Port": ["443", "80", "22"],
            }
        )
        result = stage1b_name_host_dedup(df)
        assert len(result) == 1

    def test_keeps_different_hosts(self):
        df = pd.DataFrame(
            {
                "Name": ["Vuln A", "Vuln A"],
                "Description": ["Desc 1", "Desc 2"],
                "Risk": ["Critical", "Critical"],
                "Host": ["10.0.0.1", "10.0.0.2"],
                "Port": ["443", "443"],
            }
        )
        result = stage1b_name_host_dedup(df)
        assert len(result) == 2

    def test_keeps_different_names(self):
        df = pd.DataFrame(
            {
                "Name": ["Vuln A", "Vuln B"],
                "Description": ["Desc 1", "Desc 2"],
                "Risk": ["Critical", "High"],
                "Host": ["10.0.0.1", "10.0.0.1"],
                "Port": ["443", "80"],
            }
        )
        result = stage1b_name_host_dedup(df)
        assert len(result) == 2

    def test_collapses_across_risk_levels(self):
        df = pd.DataFrame(
            {
                "Name": ["Vuln A", "Vuln A"],
                "Description": ["Desc 1", "Desc 2"],
                "Risk": ["Critical", "Medium"],
                "Host": ["10.0.0.1", "10.0.0.1"],
                "Port": ["443", "80"],
            }
        )
        result = stage1b_name_host_dedup(df)
        assert len(result) == 1

    def test_multi_host_multi_name(self):
        df = pd.DataFrame(
            {
                "Name": ["Vuln A", "Vuln A", "Vuln B", "Vuln B"],
                "Description": ["D1", "D2", "D3", "D4"],
                "Risk": ["Critical", "Critical", "High", "High"],
                "Host": ["10.0.0.1", "10.0.0.2", "10.0.0.1", "10.0.0.2"],
                "Port": ["443", "80", "22", "443"],
            }
        )
        result = stage1b_name_host_dedup(df)
        assert len(result) == 4


class TestExtractVersion:
    def test_extracts_dotted_version(self):
        assert _extract_version("Adobe Flash Player <= 32.0.0.387 Multiple") == "32.0.0.387"

    def test_extracts_last_version(self):
        assert _extract_version("Plugin 1.0 before 2.0") == "2.0"

    def test_returns_none_for_no_version(self):
        assert _extract_version("Fortinet Format String Bug") is None

    def test_extracts_version_with_qualifier(self):
        assert _extract_version("App <= 7.2.1 Security Update") == "7.2.1"


class TestVersionToTuple:
    def test_pads_to_eight(self):
        assert _version_to_tuple("7.2")[:2] == (7, 2)

    def test_full_version(self):
        assert _version_to_tuple("32.0.0.390")[:4] == (32, 0, 0, 390)

    def test_comparison_works(self):
        assert _version_to_tuple("32.0.0.390") > _version_to_tuple("32.0.0.387")
        assert _version_to_tuple("7.2.1") > _version_to_tuple("7.2")

    def test_suffix_letters_order_correctly(self):
        assert _version_to_tuple("1.0.2zn") > _version_to_tuple("1.0.2zm")
        assert _version_to_tuple("1.0.2p") < _version_to_tuple("1.0.2q")


class TestExtractRhsa:
    def test_extracts_rhsa_id(self):
        assert _extract_rhsa("RHEL 8 : kernel (RHSA-2026:3963)") == (2026, 3963)

    def test_extracts_rhsa_id_lower_number(self):
        assert _extract_rhsa("RHEL 8 : kernel (RHSA-2026:3083)") == (2026, 3083)

    def test_returns_none_for_no_rhsa(self):
        assert _extract_rhsa("Adobe Flash Player <= 32.0.0.387") is None

    def test_extracts_rhsa_different_years(self):
        assert _extract_rhsa("RHEL 8 : kernel (RHSA-2025:1234)") == (2025, 1234)


class TestExtractIdentifier:
    def test_rhsa_takes_priority(self):
        assert _extract_identifier("App 1.0 (RHSA-2026:3963)") == (2026, 3963)

    def test_falls_back_to_version(self):
        assert _extract_identifier("App <= 32.0.0.387")[:4] == (32, 0, 0, 387)

    def test_returns_none_for_no_identifier(self):
        assert _extract_identifier("Fortinet Format String Bug") is None

    def test_cpu_date_after_rhsa(self):
        assert _extract_identifier("App (RHSA-2026:3963) (January 2026 CPU)") == (2026, 3963)

    def test_cpu_date_fallback(self):
        assert _extract_identifier("Oracle Java SE (January 2026 CPU)") == (2026, 1)


class TestExtractCpuDate:
    def test_extracts_full_month(self):
        assert _extract_cpu_date("Oracle Java SE (January 2026 CPU)") == (2026, 1)

    def test_extracts_abbreviated_month(self):
        assert _extract_cpu_date("Oracle Java SE (Oct 2025 CPU)") == (2025, 10)

    def test_extracts_without_parens(self):
        assert _extract_cpu_date("Oracle Java SE July 2025 CPU") == (2025, 7)

    def test_returns_none_for_no_cpu_date(self):
        assert _extract_cpu_date("Adobe Flash Player <= 32.0.0.387") is None

    def test_extracts_july(self):
        assert _extract_cpu_date("Vulnerability (July 2025 CPU)") == (2025, 7)

    def test_extracts_december(self):
        assert _extract_cpu_date("Vulnerability (December 2024 CPU)") == (2024, 12)

    def test_comparison_works(self):
        assert _extract_cpu_date("Vuln (January 2026 CPU)") > _extract_cpu_date("Vuln (July 2025 CPU)")
        assert _extract_cpu_date("Vuln (October 2025 CPU)") > _extract_cpu_date("Vuln (July 2025 CPU)")


class TestMakeBaseTitle:
    def test_strips_rhsa(self):
        result = _make_base_title("RHEL 8 : kernel (RHSA-2026:3963)")
        assert "RHSA-2026:3963" not in result
        assert "RHEL 8 : kernel" in result

    def test_strips_version(self):
        result = _make_base_title("Adobe Flash Player <= 32.0.0.387 Multiple")
        assert "32.0.0.387" not in result
        assert "Adobe Flash Player" in result

    def test_no_version_unchanged(self):
        name = "Fortinet Format String Bug (FG-IR-23-137)"
        assert _make_base_title(name) == name

    def test_strips_cpu_date(self):
        result = _make_base_title("Oracle Java SE Multiple Vulnerabilities (January 2026 CPU)")
        assert "January 2026 CPU" not in result
        assert "Oracle Java SE" in result

    def test_strips_cpu_date_abbreviated(self):
        result = _make_base_title("Oracle Java SE (July 2025 CPU)")
        assert "July 2025 CPU" not in result
        assert "Oracle Java SE" in result

    def test_multiple_cpu_dates_same_base(self):
        base1 = _make_base_title("Oracle Java SE (January 2026 CPU)")
        base2 = _make_base_title("Oracle Java SE (July 2025 CPU)")
        base3 = _make_base_title("Oracle Java SE (October 2025 CPU)")
        assert base1 == base2 == base3


class TestStage2VersionCollapse:
    def test_keeps_highest_version(self):
        df = pd.DataFrame(
            {
                "Name": [
                    "Adobe Flash Player <= 32.0.0.387 Multiple Vulnerabilities",
                    "Adobe Flash Player <= 32.0.0.390 Multiple Vulnerabilities",
                ],
                "Risk": ["Critical", "Critical"],
                "Host": ["10.0.0.1", "10.0.0.1"],
                "Description": ["Desc old", "Desc new"],
                "CVE": ["CVE-2024-0001", "CVE-2024-0002"],
            }
        )
        plogger = PipelineLogger()
        result = stage2_version_collapse(df, plogger)
        assert len(result) == 1
        assert "32.0.0.390" in result.iloc[0]["Name"]

    def test_no_version_passes_through(self):
        df = pd.DataFrame(
            {
                "Name": ["Fortinet Format String Bug (FG-IR-23-137)"],
                "Risk": ["High"],
                "Host": ["10.0.0.9"],
                "Description": ["Desc"],
                "CVE": [""],
            }
        )
        plogger = PipelineLogger()
        result = stage2_version_collapse(df, plogger)
        assert len(result) == 1

    def test_different_base_titles_not_collapsed(self):
        df = pd.DataFrame(
            {
                "Name": [
                    "Adobe Flash Player <= 32.0.0.387 Multiple",
                    "Adobe AIR <= 33.0.0.1 Security Update",
                ],
                "Risk": ["Critical", "Critical"],
                "Host": ["10.0.0.1", "10.0.0.1"],
                "Description": ["Desc 1", "Desc 2"],
                "CVE": ["", ""],
            }
        )
        plogger = PipelineLogger()
        result = stage2_version_collapse(df, plogger)
        assert len(result) == 2

    def test_three_versions_keeps_highest(self):
        df = pd.DataFrame(
            {
                "Name": [
                    "App <= 1.0 Release",
                    "App <= 1.2 Release",
                    "App <= 1.1.5 Release",
                ],
                "Risk": ["Medium", "Medium", "Medium"],
                "Host": ["10.0.0.1", "10.0.0.1", "10.0.0.1"],
                "Description": ["v1.0", "v1.2", "v1.1.5"],
                "CVE": ["", "", ""],
            }
        )
        plogger = PipelineLogger()
        result = stage2_version_collapse(df, plogger)
        assert len(result) == 1
        assert "1.2" in result.iloc[0]["Name"]

    def test_rhsa_keeps_latest(self):
        df = pd.DataFrame(
            {
                "Name": [
                    "RHEL 8 : kernel (RHSA-2026:3083)",
                    "RHEL 8 : kernel (RHSA-2026:3963)",
                ],
                "Risk": ["Critical", "Critical"],
                "Host": ["10.0.0.1", "10.0.0.1"],
                "Description": ["Desc old", "Desc new"],
                "CVE": ["CVE-2024-0001", "CVE-2024-0002"],
            }
        )
        plogger = PipelineLogger()
        result = stage2_version_collapse(df, plogger)
        assert len(result) == 1
        assert "RHSA-2026:3963" in result.iloc[0]["Name"]

    def test_rhsa_different_hosts_not_collapsed(self):
        df = pd.DataFrame(
            {
                "Name": [
                    "RHEL 8 : kernel (RHSA-2026:3083)",
                    "RHEL 8 : kernel (RHSA-2026:3963)",
                ],
                "Risk": ["Critical", "Critical"],
                "Host": ["10.0.0.1", "10.0.0.2"],
                "Description": ["Desc 1", "Desc 2"],
                "CVE": ["", ""],
            }
        )
        plogger = PipelineLogger()
        result = stage2_version_collapse(df, plogger)
        assert len(result) == 2

    def test_rhsa_different_risks_collapsed(self):
        """Same base vuln name + same host collapses even with different Risk."""
        df = pd.DataFrame(
            {
                "Name": [
                    "RHEL 8 : kernel (RHSA-2026:3083)",
                    "RHEL 8 : kernel (RHSA-2026:3963)",
                ],
                "Risk": ["Critical", "High"],
                "Host": ["10.0.0.1", "10.0.0.1"],
                "Description": ["Desc 1", "Desc 2"],
                "CVE": ["", ""],
            }
        )
        plogger = PipelineLogger()
        result = stage2_version_collapse(df, plogger)
        assert len(result) == 1
        assert result.iloc[0]["Name"] == "RHEL 8 : kernel (RHSA-2026:3963)"

    def test_rhsa_three_keeps_latest(self):
        df = pd.DataFrame(
            {
                "Name": [
                    "RHEL 8 : kernel (RHSA-2026:0213)",
                    "RHEL 8 : kernel (RHSA-2026:3083)",
                    "RHEL 8 : kernel (RHSA-2026:3963)",
                ],
                "Risk": ["Medium", "Medium", "Medium"],
                "Host": ["10.0.0.1", "10.0.0.1", "10.0.0.1"],
                "Description": ["Desc 1", "Desc 2", "Desc 3"],
                "CVE": ["", "", ""],
            }
        )
        plogger = PipelineLogger()
        result = stage2_version_collapse(df, plogger)
        assert len(result) == 1
        assert "RHSA-2026:3963" in result.iloc[0]["Name"]

    def test_rhsa_with_no_rhsa_mixed(self):
        df = pd.DataFrame(
            {
                "Name": [
                    "RHEL 8 : kernel (RHSA-2026:3083)",
                    "RHEL 8 : kernel (RHSA-2026:3963)",
                    "RHEL 8 : kernel",
                ],
                "Risk": ["Critical", "Critical", "Critical"],
                "Host": ["10.0.0.1", "10.0.0.1", "10.0.0.1"],
                "Description": ["Desc 1", "Desc 2", "Desc 3"],
                "CVE": ["", "", ""],
            }
        )
        plogger = PipelineLogger()
        result = stage2_version_collapse(df, plogger)
        assert len(result) == 2

    def test_multi_host_rhsa_keeps_latest_per_host(self):
        df = pd.DataFrame(
            {
                "Name": [
                    "RHEL 8 : kernel (RHSA-2026:0213)",
                    "RHEL 8 : kernel (RHSA-2026:3083)",
                    "RHEL 8 : kernel (RHSA-2026:3963)",
                    "RHEL 8 : kernel (RHSA-2026:1500)",
                ],
                "Risk": ["Critical", "Critical", "Critical", "Critical"],
                "Host": ["10.0.0.1", "10.0.0.1", "10.0.0.2", "10.0.0.2"],
                "Description": ["Desc 1", "Desc 2", "Desc 3", "Desc 4"],
                "CVE": ["", "", "", ""],
            }
        )
        plogger = PipelineLogger()
        result = stage2_version_collapse(df, plogger)
        assert len(result) == 2
        host1_rows = result[result["Host"] == "10.0.0.1"]
        host2_rows = result[result["Host"] == "10.0.0.2"]
        assert len(host1_rows) == 1
        assert "RHSA-2026:3083" in host1_rows.iloc[0]["Name"]
        assert len(host2_rows) == 1
        assert "RHSA-2026:3963" in host2_rows.iloc[0]["Name"]

    def test_different_vuln_names_kept_separately(self):
        df = pd.DataFrame(
            {
                "Name": [
                    "RHEL 8 : kernel (RHSA-2026:3083)",
                    "RHEL 8 : openssh (RHSA-2026:3963)",
                ],
                "Risk": ["Critical", "Critical"],
                "Host": ["10.0.0.1", "10.0.0.1"],
                "Description": ["Desc 1", "Desc 2"],
                "CVE": ["", ""],
            }
        )
        plogger = PipelineLogger()
        result = stage2_version_collapse(df, plogger)
        assert len(result) == 2

    def test_rhsa_multi_host_keeps_latest_per_host(self):
        df = pd.DataFrame(
            {
                "Name": [
                    "RHEL 8 : kernel (RHSA-2026:0213)",
                    "RHEL 8 : kernel (RHSA-2026:3083)",
                    "RHEL 8 : kernel (RHSA-2026:3963)",
                    "RHEL 8 : kernel (RHSA-2026:1500)",
                ],
                "Risk": ["Critical", "Critical", "Critical", "Critical"],
                "Host": ["10.0.0.1", "10.0.0.1", "10.0.0.2", "10.0.0.2"],
                "Description": ["Desc 1", "Desc 2", "Desc 3", "Desc 4"],
                "CVE": ["", "", "", ""],
            }
        )
        plogger = PipelineLogger()
        result = stage2_version_collapse(df, plogger)
        assert len(result) == 2
        host1_rows = result[result["Host"] == "10.0.0.1"]
        host2_rows = result[result["Host"] == "10.0.0.2"]
        assert len(host1_rows) == 1
        assert "RHSA-2026:3083" in host1_rows.iloc[0]["Name"]
        assert len(host2_rows) == 1
        assert "RHSA-2026:3963" in host2_rows.iloc[0]["Name"]

    def test_cpu_date_keeps_latest(self):
        """Same vuln name with different CPU dates on same host keeps latest."""
        df = pd.DataFrame(
            {
                "Name": [
                    "Oracle Java SE Multiple Vulnerabilities (July 2025 CPU)",
                    "Oracle Java SE Multiple Vulnerabilities (January 2026 CPU)",
                    "Oracle Java SE Multiple Vulnerabilities (October 2025 CPU)",
                ],
                "Risk": ["Critical", "High", "High"],
                "Host": ["192.168.32.46", "192.168.32.46", "192.168.32.46"],
                "Description": ["Desc Jul", "Desc Jan", "Desc Oct"],
                "CVE": ["CVE-1", "CVE-2", "CVE-3"],
            }
        )
        plogger = PipelineLogger()
        result = stage2_version_collapse(df, plogger)
        assert len(result) == 1
        assert "January 2026 CPU" in result.iloc[0]["Name"]

    def test_cpu_date_different_hosts_not_collapsed(self):
        """Same CPU-dated vuln on different hosts keeps one per host."""
        df = pd.DataFrame(
            {
                "Name": [
                    "Oracle Java SE (July 2025 CPU)",
                    "Oracle Java SE (January 2026 CPU)",
                ],
                "Risk": ["Critical", "High"],
                "Host": ["10.0.0.1", "10.0.0.2"],
                "Description": ["Desc 1", "Desc 2"],
                "CVE": ["", ""],
            }
        )
        plogger = PipelineLogger()
        result = stage2_version_collapse(df, plogger)
        assert len(result) == 2

    def test_cpu_date_multi_host_keeps_latest_per_host(self):
        """CPU-dated vulns on multiple hosts each keep their own latest."""
        df = pd.DataFrame(
            {
                "Name": [
                    "Oracle Java SE (July 2025 CPU)",
                    "Oracle Java SE (January 2026 CPU)",
                    "Oracle Java SE (October 2025 CPU)",
                    "Oracle Java SE (January 2026 CPU)",
                ],
                "Risk": ["Critical", "Critical", "Critical", "Critical"],
                "Host": ["10.0.0.1", "10.0.0.1", "10.0.0.2", "10.0.0.2"],
                "Description": ["Desc 1", "Desc 2", "Desc 3", "Desc 4"],
                "CVE": ["", "", "", ""],
            }
        )
        plogger = PipelineLogger()
        result = stage2_version_collapse(df, plogger)
        assert len(result) == 2
        host1_rows = result[result["Host"] == "10.0.0.1"]
        host2_rows = result[result["Host"] == "10.0.0.2"]
        assert len(host1_rows) == 1
        assert "January 2026 CPU" in host1_rows.iloc[0]["Name"]
        assert len(host2_rows) == 1
        assert "January 2026 CPU" in host2_rows.iloc[0]["Name"]

    def test_cpu_date_with_rhsa_mixed(self):
        """CPU date and RHSA on same host collapse to same base title 'Oracle Java SE'."""
        df = pd.DataFrame(
            {
                "Name": [
                    "Oracle Java SE (January 2026 CPU)",
                    "Oracle Java SE (RHSA-2026:3963)",
                ],
                "Risk": ["Critical", "Critical"],
                "Host": ["10.0.0.1", "10.0.0.1"],
                "Description": ["Desc 1", "Desc 2"],
                "CVE": ["", ""],
            }
        )
        plogger = PipelineLogger()
        result = stage2_version_collapse(df, plogger)
        assert len(result) == 1
        assert "RHSA-2026:3963" in result.iloc[0]["Name"]


class TestColumnMapper:
    def test_maps_columns(self):
        df = pd.DataFrame(
            {
                "Name": ["Vuln A"],
                "Description": ["Desc A"],
                "Risk": ["Critical"],
                "Host": ["10.0.0.1"],
                "Port": ["443"],
                "Solution": ["Fix A"],
                "See Also": ["http://ref.com"],
                "CVE": ["CVE-2024-0001"],
            }
        )
        result = map_columns(df)
        assert "Vulnerbility Title" in result.columns  # sic - typo preserved
        assert "Recommendation " in result.columns     # trailing space preserved
        assert result.iloc[0]["Vulnerbility Title"] == "Vuln A"
        assert result.iloc[0]["Recommendation "] == "Fix A"

    def test_empty_strings_become_na(self):
        df = pd.DataFrame(
            {
                "Name": ["Vuln A"],
                "Description": ["Desc A"],
                "Risk": ["Critical"],
                "Host": ["10.0.0.1"],
                "Port": ["443"],
                "Solution": ["Fix A"],
                "See Also": [""],
                "CVE": [""],
            }
        )
        result = map_columns(df)
        assert result.iloc[0]["Reference"] == "N/A"
        assert result.iloc[0]["CVE"] == "N/A"

    def test_port_is_numeric(self):
        df = pd.DataFrame(
            {
                "Name": ["Vuln A"],
                "Description": ["Desc A"],
                "Risk": ["Critical"],
                "Host": ["10.0.0.1"],
                "Port": ["abc"],
                "Solution": ["Fix A"],
                "See Also": [""],
                "CVE": [""],
            }
        )
        result = map_columns(df)
        assert result.iloc[0]["Port"] == 0  # NaN -> 0


class TestSorter:
    def test_sorts_by_host_then_risk(self):
        df = pd.DataFrame(
            {
                "Risk": ["Low", "Critical", "High", "Medium"],
                "Host": ["10.0.0.1", "10.0.0.1", "10.0.0.1", "10.0.0.1"],
                "Name": ["D", "A", "B", "C"],
            }
        )
        result = sort_va_data(df)
        assert result["Risk"].tolist() == ["Critical", "High", "Medium", "Low"]
        assert result["Sr. no"].tolist() == [1, 2, 3, 4]

    def test_groups_by_host(self):
        df = pd.DataFrame(
            {
                "Risk": ["Critical", "High", "Critical"],
                "Host": ["10.0.0.2", "10.0.0.1", "10.0.0.1"],
                "Name": ["A", "B", "C"],
            }
        )
        result = sort_va_data(df)
        # Alphabetical host order: 10.0.0.1 first, then 10.0.0.2
        assert result.iloc[0]["Host"] == "10.0.0.1"
        assert result.iloc[1]["Host"] == "10.0.0.1"
        assert result.iloc[2]["Host"] == "10.0.0.2"

    def test_sr_no_is_sequential(self):
        df = pd.DataFrame(
            {
                "Risk": ["Critical", "High", "Medium", "Low"],
                "Host": ["10.0.0.1", "10.0.0.1", "10.0.0.2", "10.0.0.2"],
                "Name": ["A", "B", "C", "D"],
            }
        )
        result = sort_va_data(df)
        assert result["Sr. no"].tolist() == [1, 2, 3, 4]


class TestTextJoinHosts:
    def test_empty_df_returns_copy(self):
        df = pd.DataFrame()
        result = text_join_hosts(df)
        assert result.empty
        assert result is not df

    def test_single_row_unchanged(self):
        df = pd.DataFrame(
            {
                "Vulnerbility Title": ["Vuln A"],
                "Host": ["10.0.0.1"],
                "Description": ["Desc A"],
                "Risk": ["Critical"],
            }
        )
        result = text_join_hosts(df)
        assert len(result) == 1
        assert result.iloc[0]["Host"] == "10.0.0.1"

    def test_joins_unique_hosts(self):
        df = pd.DataFrame(
            {
                "Vulnerbility Title": ["Vuln A", "Vuln A", "Vuln A"],
                "Host": ["10.0.0.1", "10.0.0.2", "10.0.0.1"],
                "Description": ["Desc", "Desc", "Desc"],
                "Risk": ["Critical", "Critical", "Critical"],
            }
        )
        result = text_join_hosts(df)
        assert len(result) == 1
        assert result.iloc[0]["Host"] == "10.0.0.1, 10.0.0.2"

    def test_different_titles_not_joined(self):
        df = pd.DataFrame(
            {
                "Vulnerbility Title": ["Vuln A", "Vuln B"],
                "Host": ["10.0.0.1", "10.0.0.2"],
                "Description": ["Desc A", "Desc B"],
                "Risk": ["Critical", "High"],
            }
        )
        result = text_join_hosts(df)
        assert len(result) == 2

    def test_first_non_null_for_other_columns(self):
        df = pd.DataFrame(
            {
                "Vulnerbility Title": ["Vuln A", "Vuln A"],
                "Host": ["10.0.0.1", "10.0.0.2"],
                "Description": ["Desc A", ""],
                "Risk": ["Critical", "Critical"],
                "Port": [443, 80],
            }
        )
        result = text_join_hosts(df)
        assert len(result) == 1
        assert result.iloc[0]["Description"] == "Desc A"
        assert result.iloc[0]["Port"] == 443

    def test_preserves_all_risk_levels(self):
        df = pd.DataFrame(
            {
                "Vulnerbility Title": ["Vuln A", "Vuln A"],
                "Host": ["10.0.0.1", "10.0.0.2"],
                "Description": ["Desc", "Desc"],
                "Risk": ["Critical", "High"],
            }
        )
        result = text_join_hosts(df)
        assert len(result) == 1
        assert result.iloc[0]["Risk"] == "Critical"

    def test_all_hosts_unique(self):
        df = pd.DataFrame(
            {
                "Vulnerbility Title": ["Vuln A", "Vuln A", "Vuln A"],
                "Host": ["10.0.0.1", "10.0.0.2", "10.0.0.3"],
                "Description": ["Desc", "Desc", "Desc"],
                "Risk": ["Critical", "Critical", "Critical"],
            }
        )
        result = text_join_hosts(df)
        assert result.iloc[0]["Host"] == "10.0.0.1, 10.0.0.2, 10.0.0.3"
