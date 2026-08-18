"""Tests for the transform layer."""

import pandas as pd
import pytest

from va_ca_automation.logging.pipeline_logger import PipelineLogger
from va_ca_automation.transform.column_mapper import RAW_TO_TEMPLATE, map_columns
from va_ca_automation.transform.dedup import (
    _base_title,
    _extract_version,
    _version_to_tuple,
    stage1_exact_dedup,
    stage2_version_collapse,
)
from va_ca_automation.transform.filters import VA_EXCLUDE_RISKS, filter_va_candidates
from va_ca_automation.transform.sorter import RISK_WEIGHTS, sort_va_data


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
    def test_pads_to_four(self):
        assert _version_to_tuple("7.2") == (7, 2, 0, 0)

    def test_full_version(self):
        assert _version_to_tuple("32.0.0.390") == (32, 0, 0, 390)

    def test_comparison_works(self):
        assert _version_to_tuple("32.0.0.390") > _version_to_tuple("32.0.0.387")
        assert _version_to_tuple("7.2.1") > _version_to_tuple("7.2")


class TestBaseTitle:
    def test_strips_version(self):
        result = _base_title("Adobe Flash Player <= 32.0.0.387 Multiple Vulnerabilities")
        assert "32.0.0.387" not in result
        assert "Adobe Flash Player" in result

    def test_no_version_unchanged(self):
        name = "Fortinet Format String Bug (FG-IR-23-137)"
        assert _base_title(name) == name


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
        # First-seen order: 10.0.0.2 (first row), then 10.0.0.1
        assert result.iloc[0]["Host"] == "10.0.0.2"
        assert result.iloc[1]["Host"] == "10.0.0.1"
        assert result.iloc[2]["Host"] == "10.0.0.1"

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
