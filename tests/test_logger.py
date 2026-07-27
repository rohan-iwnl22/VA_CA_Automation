"""Tests for the pipeline logger."""

import json
from pathlib import Path

from va_ca_automation.logging.pipeline_logger import PipelineLogger


class TestPipelineLogger:
    def test_log_stage_count(self):
        plogger = PipelineLogger()
        plogger.log_stage_count("raw_rows", 1000)
        assert plogger.get_stage_counts()["raw_rows"] == 1000

    def test_log_unknown_risk(self):
        plogger = PipelineLogger()
        plogger.log_unknown_risk("Informational", 5)
        entries = [e for e in plogger._entries if e.get("event") == "unknown_risk"]
        assert len(entries) == 1
        assert entries[0]["count"] == 5

    def test_log_version_collapse(self):
        plogger = PipelineLogger()
        details = [
            {"base_title": "Vuln A", "host": "10.0.0.1", "risk": "Critical",
             "kept_version": "2.0", "dropped_versions": ["1.0"]}
        ]
        plogger.log_version_collapse(details)
        assert len(plogger.get_version_collapse_log()) == 1

    def test_log_risk_breakdown(self):
        plogger = PipelineLogger()
        plogger.log_risk_breakdown({"Critical": 21, "High": 32})
        entries = [e for e in plogger._entries if e.get("event") == "risk_breakdown"]
        assert len(entries) == 1

    def test_flush_writes_jsonl(self, tmp_path):
        log_file = tmp_path / "test.log"
        plogger = PipelineLogger(log_file=log_file)
        plogger.log_stage_count("test", 42)
        plogger.flush()
        assert log_file.exists()
        with open(log_file, encoding="utf-8") as f:
            lines = f.readlines()
        assert len(lines) == 1
        entry = json.loads(lines[0])
        assert entry["stage"] == "test"
        assert entry["row_count"] == 42
