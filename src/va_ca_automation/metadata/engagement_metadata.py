"""Typed engagement metadata container."""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from typing import Optional


@dataclass
class HostMetadata:
    """Per-host scan type and device type metadata."""

    ip_address: str
    scan_type: str
    device_type: str


@dataclass
class EngagementMetadata:
    """All per-engagement dynamic fields needed to populate the report template."""

    client_name: str
    security_tester: str
    reviewed_by: str
    report_date: date
    report_version: str
    scanner_name: str = "Nessus "
    scanner_version: str = "10.11.4"
    report_owner: str = ""
    scope_label: str = "Server"
    phase_label: str = "First"
    entity_codes: list[str] = field(default_factory=list)
    default_device_type: str = ""
    host_metadata: dict[str, HostMetadata] = field(default_factory=dict)

    # --- New fields for Word document sections ---
    report_type: str = "First"             # "First" or "Final"
    report_number: str = "1.0"             # e.g. "1.0", "1.1"
    client_short_name: str = ""            # Short name for Document ID
    assessment_start_date: str = ""        # YYYY-MM-DD
    assessment_finish_date: str = ""       # YYYY-MM-DD
    final_retesting_start: str = ""        # YYYY-MM-DD (Final only)
    final_retesting_finish: str = ""       # YYYY-MM-DD (Final only)
    released_date: str = ""                # YYYY-MM-DD
    spokesperson_name: str = ""
    spokesperson_designation: str = ""
    spokesperson_email: str = ""
    senior_name: str = ""                  # Vinit, Abhishek, Sravan, Chirag
    approved_by: str = "Mr. Vijay Sawant"

    # --- Report Release Date table ---
    report_release_date: str = ""          # YYYY-MM-DD
    period: str = ""                       # e.g. "May 2026"
    document_id: str = ""                  # e.g. "SCPL / TIAPL/ VAPT / 04"

    # --- Document Change History ---
    change_history_version: str = ""       # e.g. "1.0"
    change_history_date: str = ""          # YYYY-MM-DD
    change_history_remarks: str = ""       # e.g. "First Audit Report"

    # --- Document Distribution List ---
    distribution_name: str = ""
    distribution_organization: str = ""
    distribution_designation: str = ""
    distribution_email: str = ""

    # --- Details of Auditing Team (Row 1 & 2: user input, Row 3: default) ---
    auditor_1_name: str = ""
    auditor_1_designation: str = ""
    auditor_1_email: str = ""
    auditor_1_qualifications: str = ""
    auditor_1_cert_in: str = "Yes"

    auditor_2_name: str = ""
    auditor_2_designation: str = ""
    auditor_2_email: str = ""
    auditor_2_qualifications: str = ""
    auditor_2_cert_in: str = "Yes"

    # --- Computed properties ---

    @property
    def document_version(self) -> str:
        """Return document version based on report type."""
        if self.report_type == "First":
            return "1.0"
        return self.report_number

    @property
    def document_title(self) -> str:
        """Return document title based on report type."""
        if self.report_type == "First":
            return "First Audit Report"
        return "Final Audit Report"

    @property
    def assessment_date_range(self) -> str:
        """Return assessment date range."""
        return f"{self.assessment_start_date} to {self.assessment_finish_date}"

    @property
    def first_audit_dates(self) -> str:
        """Return first audit dates or NA for Final reports."""
        if self.report_type == "First":
            return f"{self.assessment_start_date} to {self.assessment_finish_date}"
        return "NA"

    @property
    def final_retesting_dates(self) -> str:
        """Return final retesting dates or 'Revalidation not performed' for First reports."""
        if self.report_type == "Final":
            return f"{self.final_retesting_start} to {self.final_retesting_finish}"
        return "Revalidation not performed"

    def get_host_scan_type(self, ip: str) -> str:
        """Return scan type for a host, or empty string if not supplied."""
        meta = self.host_metadata.get(ip)
        return meta.scan_type if meta else ""

    def get_host_device_type(self, ip: str) -> str:
        """Return device type for a host, falling back to default_device_type."""
        meta = self.host_metadata.get(ip)
        if meta and meta.device_type:
            return meta.device_type
        return self.default_device_type
