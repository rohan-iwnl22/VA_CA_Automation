"""Top-level application entrypoint."""

from __future__ import annotations

import argparse
import logging
import sys
from datetime import date
from pathlib import Path

from .metadata.engagement_metadata import EngagementMetadata, HostMetadata
from .pipelines.va_pipeline import run_va_pipeline

# Resolve project root (src/va_ca_automation -> src -> project root)
_PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent
DEFAULT_TEMPLATE_PATH = _PROJECT_ROOT / "templates" / "va_report_template.xlsx"


def setup_logging(verbose: bool = False) -> None:
    """Configure logging for the application."""
    level = logging.DEBUG if verbose else logging.INFO
    logging.basicConfig(
        level=level,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )


def build_metadata_from_args(args: argparse.Namespace) -> EngagementMetadata:
    """Build EngagementMetadata from parsed CLI arguments."""
    host_metadata = {}
    if args.host_metadata:
        for entry in args.host_metadata:
            parts = entry.split(":")
            if len(parts) == 3:
                ip, scan_type, device_type = parts
                host_metadata[ip.strip()] = HostMetadata(
                    ip_address=ip.strip(),
                    scan_type=scan_type.strip(),
                    device_type=device_type.strip(),
                )

    return EngagementMetadata(
        client_name=args.client_name,
        security_tester=args.tester,
        reviewed_by=args.reviewer,
        report_date=date.fromisoformat(args.report_date) if args.report_date else date.today(),
        report_version=args.version,
        scanner_name=args.scanner_name or "Nessus ",
        scanner_version=args.scanner_version or "10.11.4",
        report_owner=args.report_owner or "",
        scope_label=args.scope or "Server",
        phase_label=args.phase or "First",
        entity_codes=args.entity_codes or [],
        default_device_type=args.device_type or "",
        host_metadata=host_metadata,
    )


def main(argv: list[str] | None = None) -> int:
    """Run the application."""
    parser = argparse.ArgumentParser(
        prog="va-ca-automation",
        description="VA/CA Report Automation - Process Nessus exports into client-ready reports",
    )

    parser.add_argument(
        "raw_file",
        type=Path,
        help="Path to the raw Nessus export (.xlsx)",
    )
    parser.add_argument(
        "-t", "--template",
        type=Path,
        default=None,
        help="Path to the blank template (.xlsx). Defaults to templates/va_report_template.xlsx",
    )
    parser.add_argument(
        "-o", "--output-dir",
        type=Path,
        default=Path("output"),
        help="Output directory for the generated report (default: output/)",
    )
    parser.add_argument(
        "--client-name",
        required=True,
        help="Client legal name for the report",
    )
    parser.add_argument(
        "--tester",
        required=True,
        help="Security tester name",
    )
    parser.add_argument(
        "--reviewer",
        required=True,
        help="Reviewer name",
    )
    parser.add_argument(
        "--report-date",
        default=None,
        help="Report date in YYYY-MM-DD format (default: today)",
    )
    parser.add_argument(
        "--version",
        type=str,
        default="1.0",
        help="Report version number (default: 1.0)",
    )
    parser.add_argument(
        "--scanner-name",
        default=None,
        help="Scanner name (default: Nessus )",
    )
    parser.add_argument(
        "--scanner-version",
        default=None,
        help="Scanner version (default: 10.11.4)",
    )
    parser.add_argument(
        "--report-owner",
        default=None,
        help="Report owner name for Introduction sheet",
    )
    parser.add_argument(
        "--scope",
        default=None,
        help="Scope label (e.g., Server, Firewall)",
    )
    parser.add_argument(
        "--device-type",
        default=None,
        help="Device type for all hosts in the summary table (e.g., Server, Firewall)",
    )
    parser.add_argument(
        "--phase",
        default=None,
        help="Phase label (e.g., First, Retest)",
    )
    parser.add_argument(
        "--entity-codes",
        nargs="*",
        default=None,
        help="Entity codes (e.g., TSS SCPL)",
    )
    parser.add_argument(
        "--host-metadata",
        nargs="*",
        default=None,
        help="Per-host metadata as IP:ScanType:DeviceType (e.g., 192.168.1.1:Authenticated:Server)",
    )
    parser.add_argument(
        "--log-file",
        type=Path,
        default=None,
        help="Path to write structured log entries",
    )
    parser.add_argument(
        "-v", "--verbose",
        action="store_true",
        help="Enable verbose/debug logging",
    )

    args = parser.parse_args(argv)
    setup_logging(args.verbose)

    # Resolve template path
    template_path = args.template
    if template_path is None:
        template_path = DEFAULT_TEMPLATE_PATH

    if not args.raw_file.exists():
        print(f"Error: Raw file not found: {args.raw_file}", file=sys.stderr)
        return 1

    if not template_path.exists():
        print(f"Error: Template file not found: {template_path}", file=sys.stderr)
        print("Please provide a blank template using --template or place it at templates/va_report_template.xlsx", file=sys.stderr)
        return 1

    metadata = build_metadata_from_args(args)

    try:
        output_path = run_va_pipeline(
            raw_file_path=args.raw_file,
            template_path=template_path,
            metadata=metadata,
            output_dir=args.output_dir,
            log_file=args.log_file,
        )
        print(f"Report generated: {output_path}")
        return 0
    except Exception as e:
        logging.getLogger("va_ca_automation").error("Pipeline failed: %s", e, exc_info=True)
        print(f"Error: {e}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
