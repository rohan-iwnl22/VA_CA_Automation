"""Construct output filenames from engagement metadata."""

from __future__ import annotations

import re
from datetime import date
from pathlib import Path

from ..metadata.engagement_metadata import EngagementMetadata


def _sanitize_filename(name: str) -> str:
    """Remove characters that are unsafe in filenames."""
    unsafe = re.compile(r'[/\\:*?"<>|&.]')
    return unsafe.sub("", name)


def build_filename(metadata: EngagementMetadata, report_type: str = "VA") -> str:
    """Build the output filename following the naming convention.

    Pattern: <ReportType>_<Scope>_<Phase>_Audit_Report_<ClientLegalName>_<EntityCode(s)>_<Year>_V<Major>.<Minor>.xlsx

    For CA reports, report_type should be "Configuration_Audit".
    """
    scope = _sanitize_filename(metadata.scope_label)
    phase = _sanitize_filename(metadata.phase_label)
    client_clean = _sanitize_filename(metadata.client_name.replace(" ", "_"))
    year = str(metadata.report_date.year)
    version_str = str(metadata.report_version).replace(".", "_")
    version_part = f"V{version_str}"

    parts = [report_type, scope, phase, "Audit_Report", client_clean]

    if metadata.entity_codes:
        for code in metadata.entity_codes:
            parts.append(_sanitize_filename(code))

    parts.append(year)
    parts.append(version_part)

    filename = "_".join(parts) + ".xlsx"
    return filename


def ensure_unique_path(path: Path) -> Path:
    """Return a unique path by appending a counter if the file already exists."""
    if not path.exists():
        return path
    stem = path.stem
    suffix = path.suffix
    parent = path.parent
    counter = 1
    while True:
        candidate = parent / f"{stem}_{counter}{suffix}"
        if not candidate.exists():
            return candidate
        counter += 1
