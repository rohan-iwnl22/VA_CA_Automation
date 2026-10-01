"""Construct output filenames from engagement metadata."""

from __future__ import annotations

import re
from pathlib import Path

from ..metadata.engagement_metadata import EngagementMetadata

# Fixed entity code embedded in every generated CA/VA filename.
ENTITY_CODE = "SCPL"

# Report type token -> pipeline report type argument.
REPORT_TYPE_TOKENS = ("VA", "Configuration_Audit")

_PHASES = ("First", "Final")

# <ReportType>_<Device>_<First|Final>_Audit_Report_<Client>_SCPL_<Year>_V<version>[.textjoin].xlsx
_FILENAME_RE = re.compile(
    r"^(?:VA|Configuration_Audit)"
    r"_[^/\\:*?\"<>|&]+"
    r"_(?:First|Final)_Audit_Report_"
    r".+"
    rf"_{ENTITY_CODE}_"
    r"\d{4}"
    r"_V\d+(?:\.\d+)*"
    r"(?:\.textjoin)?\.xlsx$"
)

_VERSION_RE = re.compile(r"^\d+(?:\.\d+)*$")


class NamingConventionError(ValueError):
    """Raised when a filename or a filename token violates the naming convention."""


def _sanitize_filename(name: str) -> str:
    """Remove characters that are unsafe in filenames."""
    unsafe = re.compile(r'[/\\:*?"<>|&.]')
    return unsafe.sub("", name)


def _device_token(metadata: EngagementMetadata) -> str:
    """Device token: user-supplied device_type, falling back to scope label."""
    device = (metadata.default_device_type or "").strip() or (metadata.scope_label or "").strip()
    token = _sanitize_filename(device.replace(" ", "_"))
    if not token:
        raise NamingConventionError(
            "device type is required for the filename (device_type or scope must be non-empty)"
        )
    return token


def _phase_token(metadata: EngagementMetadata) -> str:
    """First/Final token taken from the report_type form input."""
    token = (metadata.report_type or "").strip()
    token = token[:1].upper() + token[1:].lower() if token else token
    if token not in _PHASES:
        raise NamingConventionError(
            f"report_type must be one of {_PHASES}, got {metadata.report_type!r}"
        )
    return token


def _version_token(metadata: EngagementMetadata) -> str:
    """Version token like 'V1.0' (uppercase V, dot kept)."""
    raw = str(metadata.report_version).strip()
    raw = re.sub(r"^[vV]", "", raw)
    if not _VERSION_RE.match(raw):
        raise NamingConventionError(
            f"report_number must look like '1.0' or '1.1', got {metadata.report_version!r}"
        )
    return f"V{raw}"


def build_filename(metadata: EngagementMetadata, report_type: str = "VA") -> str:
    """Build the output filename following the naming convention.

    Pattern: <ReportType>_<Device>_<First|Final>_Audit_Report_<Client>_SCPL_<Year>_V<Version>.xlsx

    Examples:
        VA Report:   VA_Server_Final_Audit_Report_Client_Name_SCPL_2026_V1.0.xlsx
        CA Report:   Configuration_Audit_Server_First_Audit_Report_Client_Name_SCPL_2026_V1.0.xlsx

    For CA reports, report_type should be "Configuration_Audit".
    Device token comes from metadata.default_device_type (falls back to scope_label).
    First/Final token comes from metadata.report_type.
    Raises NamingConventionError if any token or the final name violates the convention.
    """
    if report_type not in REPORT_TYPE_TOKENS:
        raise NamingConventionError(
            f"report_type must be one of {REPORT_TYPE_TOKENS}, got {report_type!r}"
        )

    device = _device_token(metadata)
    phase = _phase_token(metadata)
    client_clean = _sanitize_filename((metadata.client_name or "").replace(" ", "_"))
    if not client_clean:
        raise NamingConventionError("client_name is required for the filename")
    year = str(metadata.report_date.year)
    version_part = _version_token(metadata)

    parts = [report_type, device, phase, "Audit_Report", client_clean, ENTITY_CODE, year, version_part]
    filename = re.sub(r"_{2,}", "_", "_".join(parts)) + ".xlsx"

    if not _FILENAME_RE.match(filename):
        raise NamingConventionError(f"filename violates the naming convention: {filename}")
    return filename


def build_textjoin_filename(filename: str) -> str:
    """Insert the '.textjoin' marker before the extension.

    'VA_Server_..._V1.0.xlsx' -> 'VA_Server_..._V1.0.textjoin.xlsx'
    """
    path = Path(filename)
    if path.stem.endswith(".textjoin"):
        return filename
    return f"{path.stem}.textjoin{path.suffix}"


def validate_filename(filename: str) -> bool:
    """Return True if the filename matches the naming convention (TextJoin allowed)."""
    return bool(_FILENAME_RE.match(filename))


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
