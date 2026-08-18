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
