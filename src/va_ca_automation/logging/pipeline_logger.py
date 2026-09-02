"""Structured logging for pipeline stages."""

from __future__ import annotations

import json
import logging
import time
from pathlib import Path
from typing import Any

logger = logging.getLogger("va_ca_automation")


class PipelineLogger:
    """Stage-by-stage structured logger for QA/auditability."""

    def __init__(self, log_file: Path | None = None) -> None:
        self._stage_counts: dict[str, int] = {}
        self._version_collapse_log: list[dict[str, Any]] = []
        self._rhsa_collapse_log: list[dict[str, Any]] = []
        self._unknown_risks: list[dict[str, Any]] = []
        self._start_time = time.monotonic()
        self._log_file = log_file
        self._entries: list[dict[str, Any]] = []

    def log_stage_count(self, stage: str, count: int) -> None:
        """Record a row count at a pipeline stage."""
        self._stage_counts[stage] = count
        entry = {"stage": stage, "row_count": count}
        self._entries.append(entry)
        logger.info("Stage '%s': %d rows", stage, count)

    def log_unknown_risk(self, risk_value: str, count: int) -> None:
        """Record an unknown Risk value encountered during ingestion."""
        self._unknown_risks.append({"risk_value": risk_value, "count": count})
        entry = {"event": "unknown_risk", "risk_value": risk_value, "count": count}
        self._entries.append(entry)
        logger.warning("Unknown Risk value '%s' encountered %d times", risk_value, count)

    def log_version_collapse(self, details: list[dict[str, Any]]) -> None:
        """Record version-collapse dedup details for QA."""
        self._version_collapse_log = details
        for d in details:
            entry = {"event": "version_collapse", **d}
            self._entries.append(entry)
        if details:
            logger.info("Version-collapse: %d groups had versions collapsed", len(details))

    def log_rhsa_collapse(self, details: list[dict[str, Any]]) -> None:
        """Record RHSA-advisory collapse dedup details for QA."""
        self._rhsa_collapse_log = details
        for d in details:
            entry = {"event": "rhsa_collapse", **d}
            self._entries.append(entry)
        if details:
            logger.info("RHSA-collapse: %d groups had advisories collapsed", len(details))

    def log_risk_breakdown(self, breakdown: dict[str, int]) -> None:
        """Record the final risk breakdown."""
        entry = {"event": "risk_breakdown", "counts": breakdown}
        self._entries.append(entry)
        logger.info("Risk breakdown: %s", breakdown)

    def log_output_file(self, path: str) -> None:
        """Record the output file path."""
        entry = {"event": "output_file", "path": path}
        self._entries.append(entry)
        logger.info("Output file: %s", path)

    def log_summary(self) -> None:
        """Write a final summary line with elapsed time."""
        elapsed = time.monotonic() - self._start_time
        entry = {"event": "pipeline_complete", "elapsed_seconds": round(elapsed, 2)}
        self._entries.append(entry)
        logger.info("Pipeline completed in %.2f seconds", elapsed)

    def get_stage_counts(self) -> dict[str, int]:
        """Return the recorded stage counts."""
        return dict(self._stage_counts)

    def get_version_collapse_log(self) -> list[dict[str, Any]]:
        """Return the version-collapse details."""
        return list(self._version_collapse_log)

    def get_rhsa_collapse_log(self) -> list[dict[str, Any]]:
        """Return the RHSA-collapse details."""
        return list(self._rhsa_collapse_log)

    def flush(self) -> None:
        """Write all entries to the log file if configured."""
        if self._log_file is None:
            return
        self._log_file.parent.mkdir(parents=True, exist_ok=True)
        with open(self._log_file, "w", encoding="utf-8") as f:
            for entry in self._entries:
                f.write(json.dumps(entry, default=str) + "\n")
