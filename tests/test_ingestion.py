"""Tests for the ingestion layer."""

from pathlib import Path

import pandas as pd
import pytest

from va_ca_automation.ingestion.raw_file_loader import (
    EXPECTED_COLUMNS,
    RAW_SHEET_NAME,
    SchemaError,
    SheetNotFoundError,
    load_raw_file,
)
from va_ca_automation.ingestion.schema_validator import (
    ALL_KNOWN_RISKS,
    KNOWN_CA_RISKS,
    KNOWN_VA_RISKS,
    classify_rows,
    normalize_whitespace_columns,
    validate_and_normalize_risk,
)
from va_ca_automation.logging.pipeline_logger import PipelineLogger


@pytest.fixture
def sample_raw_df():
    """Create a minimal raw DataFrame matching the expected schema."""
    n = 6
    data = {col: [""] * n for col in EXPECTED_COLUMNS}
    data["Risk"] = ["Critical", "High", "Medium", "Low", "None", "PASSED"]
    data["Host"] = ["10.0.0.1", "10.0.0.2", "10.0.0.1", "10.0.0.3", "10.0.0.4", "10.0.0.5"]
    data["Name"] = ["Vuln A", "Vuln B", "Vuln C", "Vuln D", "Info E", "Check F"]
    data["Description"] = ["Desc A", "Desc B", "Desc C", "Desc D", "Info Desc", "Check Desc"]
    return pd.DataFrame(data)


class TestNormalizeWhitespace:
    def test_strips_whitespace(self):
        df = pd.DataFrame({"Risk": [" Critical ", "high"], "Host": [" 10.0.0.1 ", "10.0.0.2"]})
        result = normalize_whitespace_columns(df, ["Risk", "Host"])
        assert result["Risk"].tolist() == ["Critical", "high"]
        assert result["Host"].tolist() == ["10.0.0.1", "10.0.0.2"]

    def test_does_not_mutate_original(self):
        df = pd.DataFrame({"Risk": [" Critical "]})
        original_value = df["Risk"].iloc[0]
        normalize_whitespace_columns(df, ["Risk"])
        assert df["Risk"].iloc[0] == original_value


class TestValidateAndNormalizeRisk:
    def test_title_cases_risk(self):
        df = pd.DataFrame({"Risk": ["critical", "HIGH", "Medium", "none"]})
        plogger = PipelineLogger()
        result = validate_and_normalize_risk(df, plogger)
        assert result["Risk"].tolist() == ["Critical", "High", "Medium", "None"]

    def test_logs_unknown_risks(self):
        df = pd.DataFrame({"Risk": ["Critical", "Informational", "Informational"]})
        plogger = PipelineLogger()
        validate_and_normalize_risk(df, plogger)
        unknown = [e for e in plogger._entries if e.get("event") == "unknown_risk"]
        assert len(unknown) == 1
        assert unknown[0]["risk_value"] == "Informational"
        assert unknown[0]["count"] == 2


class TestClassifyRows:
    def test_classifies_va_and_ca(self, sample_raw_df):
        va, ca, unknown = classify_rows(sample_raw_df)
        assert len(va) == 5  # Critical, High, Medium, Low, None (None is VA-routed, filtered later)
        assert len(ca) == 1   # PASSED
        assert len(unknown) == 0

    def test_none_is_va_risk(self):
        assert "None" in KNOWN_VA_RISKS

    def test_all_known_risks(self):
        assert KNOWN_VA_RISKS == {"Critical", "High", "Medium", "Low", "None"}
        assert KNOWN_CA_RISKS == {"PASSED", "FAILED", "WARNING"}
