"""POST /api/word — Generate Word report from VA/CA Excel reports."""

from __future__ import annotations

import logging
import tempfile
from pathlib import Path

import pandas as pd
from datetime import date, datetime
from fastapi import APIRouter, Depends, File, Form, UploadFile
from fastapi.responses import JSONResponse
from openpyxl import load_workbook

from ...metadata.engagement_metadata import EngagementMetadata
from ...naming.filename_builder import NamingConventionError, build_word_filename
from ...word_writer.word_report_builder import build_word_report
from ..deps import get_current_user
from ..temp_registry import create_session, store_file

router = APIRouter(tags=["word"])
logger = logging.getLogger("va_ca_automation")


@router.post("/word/read-metadata")
async def read_excel_metadata(
    file: UploadFile = File(...),
    current_user: dict = Depends(get_current_user),
):
    """Read client name and other metadata from an uploaded Excel file."""
    temp_dir = Path(tempfile.mkdtemp())
    try:
        content = await file.read()
        suffix = Path(file.filename).suffix if file.filename else ".xlsx"
        tmp_path = temp_dir / (file.filename or f"upload{suffix}")
        tmp_path.write_bytes(content)

        meta = {}
        read_path, error = _prepare_workbook(tmp_path)
        if not error:
            classified = _classify_report(read_path)
            if classified and classified[0] == "va":
                meta = _read_va_metadata(read_path, classified[1])
            elif classified and classified[0] == "ca":
                meta = _read_ca_metadata(read_path, classified[1])

        return JSONResponse(content=meta)
    except Exception as e:
        logger.warning("Could not read metadata: %s", e)
        return JSONResponse(content={})
    finally:
        import shutil
        shutil.rmtree(temp_dir, ignore_errors=True)


def _normalize_sheet_name(name: str) -> str:
    """Normalize a sheet name for tolerant matching (case, spacing, underscores)."""
    return " ".join(str(name).strip().lower().replace("_", " ").split())


def _get_sheet(workbook, wanted: str, actual: str | None = None):
    """Return a worksheet by actual name, or by normalized wanted name."""
    if actual is not None:
        for sheet in workbook.sheetnames:
            if sheet == actual:
                return workbook[sheet]
    wanted_norm = _normalize_sheet_name(wanted)
    for sheet in workbook.sheetnames:
        if _normalize_sheet_name(sheet) == wanted_norm:
            return workbook[sheet]
    raise KeyError(wanted)


def _sheet_name_kind(norm_name: str) -> str | None:
    """Classify a normalized sheet name as 'va' or 'ca'."""
    if norm_name in ("va report", "va") or norm_name.startswith("va report"):
        return "va"
    if norm_name in ("ca report", "ca") or norm_name.startswith("ca report"):
        return "ca"
    return None


def _sheet_content_kind(worksheet) -> str | None:
    """Classify a sheet by its header text, for files with renamed sheets."""
    cells: list[str] = []
    try:
        for row in worksheet.iter_rows(
            min_row=1, max_row=min(worksheet.max_row, _HEADER_SCAN_ROWS),
            max_col=_SCAN_MAX_COL, values_only=True,
        ):
            for value in row:
                if value is not None:
                    cells.append(str(value).strip().lower())
    except Exception:
        return None
    blob = " | ".join(cells)
    if "vulnerbility title" in blob or "vulnerability title" in blob:
        return "va"
    if "solution" in blob and "impact" in blob:
        return "ca"
    if "compliance" in blob and "host" in blob:
        return "ca"
    return None


def _classify_report(path: Path) -> tuple[str, str] | None:
    """Return (kind, sheet_name) for an uploaded Excel report, or None.

    Sheet-name matching is tolerant of casing, spacing and underscores, and a
    content fallback recognises VA/CA sheets that were renamed.
    """
    try:
        wb = load_workbook(path, read_only=True, data_only=True)
        try:
            for sheet in wb.sheetnames:
                kind = _sheet_name_kind(_normalize_sheet_name(sheet))
                if kind:
                    return kind, sheet
            for sheet in wb.sheetnames:
                kind = _sheet_content_kind(wb[sheet])
                if kind:
                    return kind, sheet
        finally:
            wb.close()
    except Exception:
        return None
    return None


def _report_kind(path: Path) -> str | None:
    """Return 'va', 'ca', or None for an uploaded Excel report.

    Detection is tolerant of casing, spacing and underscore differences, so
    'VA Report', 'VA_Report', 'va report', 'VA REPORT' all count as VA, and
    'CA_Report', 'CA report', 'ca_report' all count as CA. Sheets renamed to
    something else are still recognised from their column headers.
    """
    classified = _classify_report(path)
    return classified[0] if classified else None


def _prepare_workbook(path: Path) -> tuple[Path, str | None]:
    """Return (path openpyxl can read, error message or None).

    Legacy .xls uploads (BIFF, e.g. CONFIG~1.XLS) cannot be read by openpyxl,
    so they are converted to .xlsx first. Modern .xlsx uploads pass through.
    """
    if path.suffix.lower() != ".xls":
        return path, None
    try:
        sheets = pd.read_excel(path, sheet_name=None, header=None, engine="xlrd")
    except Exception as exc:
        return path, f"could not be read as a legacy .xls file ({type(exc).__name__})"
    if not sheets:
        return path, "contains no sheets"
    converted = path.with_name(f"{path.stem}.converted.xlsx")
    with pd.ExcelWriter(converted, engine="openpyxl") as writer:
        for sheet_name, frame in sheets.items():
            frame.to_excel(
                writer, sheet_name=str(sheet_name)[:31], index=False, header=False
            )
    return converted, None


def _is_va_report(path: Path) -> bool:
    """Check if an Excel file is a VA report by looking for a 'VA Report' sheet."""
    return _report_kind(path) == "va"


def _is_ca_report(path: Path) -> bool:
    """Check if an Excel file is a CA report by looking for a 'CA_Report' sheet."""
    return _report_kind(path) == "ca"


def _read_va_metadata(va_path: Path, sheet_name: str | None = None) -> dict:
    """Read engagement metadata from VA Excel report header (rows 5-10, column C)."""
    meta = {}
    try:
        wb = load_workbook(va_path, read_only=True, data_only=True)
        ws = _get_sheet(wb, "VA Report", sheet_name)
        meta["client_name"] = ws["C5"].value or ""
        meta["security_tester"] = ws["C6"].value or ""
        meta["reviewed_by"] = ws["C7"].value or ""
        raw_date = ws["C8"].value
        if raw_date:
            if hasattr(raw_date, "strftime"):
                meta["report_date"] = raw_date.strftime("%Y-%m-%d")
            else:
                meta["report_date"] = str(raw_date)
        else:
            meta["report_date"] = ""
        meta["report_version"] = ws["C9"].value or ""
        wb.close()
    except Exception as e:
        logger.warning("Could not read metadata from VA Excel: %s", e)
    return meta


def _read_ca_metadata(ca_path: Path, sheet_name: str | None = None) -> dict:
    """Read engagement metadata from CA Excel report header (rows 5-10, column C)."""
    meta = {}
    try:
        wb = load_workbook(ca_path, read_only=True, data_only=True)
        ws = _get_sheet(wb, "CA_Report", sheet_name)
        meta["client_name"] = ws["C5"].value or ""
        meta["security_tester"] = ws["C6"].value or ""
        meta["reviewed_by"] = ws["C7"].value or ""
        raw_date = ws["C8"].value
        if raw_date:
            if hasattr(raw_date, "strftime"):
                meta["report_date"] = raw_date.strftime("%Y-%m-%d")
            else:
                meta["report_date"] = str(raw_date)
        else:
            meta["report_date"] = ""
        meta["report_version"] = ws["C9"].value or ""
        wb.close()
    except Exception as e:
        logger.warning("Could not read metadata from CA Excel: %s", e)
    return meta


def _read_pivot_summary(path: Path, expected: dict[str, int]) -> dict[str, int]:
    """Read a severity/status pivot from a Summary sheet.

    The block is located by the 'Row Labels' header text (per the Word
    generation guide), not fixed cell coordinates, and missing categories
    count as 0.
    """
    counts = dict(expected)
    try:
        wb = load_workbook(path, data_only=True)
        try:
            ws = _get_sheet(wb, "Summary")
            anchor = None
            for row in ws.iter_rows():
                for cell in row:
                    if cell.value is not None and str(cell.value).strip().lower() == "row labels":
                        anchor = cell
                        break
                if anchor is not None:
                    break
            if anchor is None:
                return counts

            label_col = anchor.column
            value_col = label_col + 1
            row_idx = anchor.row + 1
            while True:
                raw_label = ws.cell(row=row_idx, column=label_col).value
                if raw_label is None or str(raw_label).strip() == "":
                    break
                label = str(raw_label).strip()
                raw_value = ws.cell(row=row_idx, column=value_col).value
                try:
                    count = int(raw_value) if raw_value is not None else 0
                except (TypeError, ValueError):
                    count = 0
                for key in counts:
                    if label.lower() == key.lower():
                        counts[key] = count
                        break
                else:
                    if label.lower() == "grand total":
                        counts["Grand Total"] = count
                row_idx += 1

            if counts.get("Grand Total", 0) == 0:
                counts["Grand Total"] = sum(
                    v for k, v in counts.items() if k != "Grand Total"
                )
        finally:
            wb.close()
    except Exception as e:
        logger.warning("Could not read pivot summary from %s: %s", path.name, e)
    return counts


# ── Robust finding-table reader ────────────────────────────────────────────
# The header row is located by its labels instead of assuming row 13, and a
# row is treated as data whenever any recognised cell is filled, so files
# with a different layout (raw exports, extra title rows, blank Sr. no) do
# not lose findings.

_HEADER_SCAN_ROWS = 40
_SCAN_MAX_COL = 30
_EMPTY_RUN_LIMIT = 50

_VA_ALIASES = {
    "sr": ("sr. no", "sr.no", "sr no", "s.r. no", "serial no", "serial number", "#"),
    "title": ("vulnerbility title", "vulnerability title", "plugin name", "plugin title"),
    "description": ("description", "synopsis", "plugin text"),
    "risk": ("risk", "severity", "risk factor"),
    "host": ("host", "ip", "ip address", "asset", "dns name"),
    "port": ("port", "port/protocol", "protocol"),
    "recommendation": ("recommendation", "solution", "remediation"),
    "reference": ("reference", "references", "see also"),
    "cve": ("cve", "cve id", "cves"),
}

_CA_ALIASES = {
    "sr": _VA_ALIASES["sr"],
    "title": ("title", "check name", "control name", "control", "name"),
    "host": _VA_ALIASES["host"],
    "description": ("description", "synopsis", "plugin text"),
    "solution": ("solution", "remediation", "recommendation"),
    "impact": ("impact",),
    "risk": ("status", "risk", "severity", "result"),
}

_VA_WORD_COLUMNS = [
    "Sr. no", "Vulnerability Title", "Description", "Risk",
    "Host", "Port", "Recommendation ", "Reference", "CVE",
]
_CA_WORD_COLUMNS = ["Sr.No.", "Title", "Host", "Description", "Solution", "Impact", "Risk"]

_VA_CANONICAL = {
    "sr": "Sr. no", "title": "Vulnerability Title", "description": "Description",
    "risk": "Risk", "host": "Host", "port": "Port",
    "recommendation": "Recommendation ", "reference": "Reference", "cve": "CVE",
}
_CA_CANONICAL = {
    "sr": "Sr.No.", "title": "Title", "host": "Host", "description": "Description",
    "solution": "Solution", "impact": "Impact", "risk": "Risk",
}


def _norm_label(value) -> str:
    """Normalize a header cell for alias matching."""
    if value is None:
        return ""
    text = str(value).strip().lower().replace("_", " ")
    return " ".join(text.split()).rstrip(".: ")


def _map_header(row: tuple, aliases: dict[str, tuple[str, ...]]) -> dict[str, int]:
    """Map canonical field names to column indexes for one header row."""
    mapping: dict[str, int] = {}
    for idx, value in enumerate(row):
        label = _norm_label(value)
        if not label:
            continue
        for canon, names in aliases.items():
            if canon not in mapping and label in names:
                mapping[canon] = idx
    return mapping


def _read_findings(path: Path, sheet_name: str | None, kind: str) -> pd.DataFrame:
    """Read one VA/CA finding table, locating its header row automatically."""
    aliases = _VA_ALIASES if kind == "va" else _CA_ALIASES
    canonical = _VA_CANONICAL if kind == "va" else _CA_CANONICAL
    columns = _VA_WORD_COLUMNS if kind == "va" else _CA_WORD_COLUMNS
    fallback_sheet = "VA Report" if kind == "va" else "CA_Report"

    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        ws = _get_sheet(wb, fallback_sheet, sheet_name)

        header_row_idx = None
        mapping: dict[str, int] = {}
        for row_idx, row in enumerate(
            ws.iter_rows(
                min_row=1, max_row=min(ws.max_row, _HEADER_SCAN_ROWS),
                max_col=_SCAN_MAX_COL, values_only=True,
            ),
            start=1,
        ):
            candidate = _map_header(row, aliases)
            if "title" not in candidate or "host" not in candidate:
                continue
            if kind == "ca" and not ({"solution", "impact", "risk"} & candidate.keys()):
                continue
            header_row_idx = row_idx
            mapping = candidate
            break

        if header_row_idx is None:
            return pd.DataFrame(columns=columns)

        max_data_row = min(ws.max_row, header_row_idx + 50000)
        records: list[dict] = []
        empty_run = 0
        for row in ws.iter_rows(
            min_row=header_row_idx + 1, max_row=max_data_row,
            max_col=_SCAN_MAX_COL, values_only=True,
        ):
            if not any(v is not None and str(v).strip() != "" for v in row):
                empty_run += 1
                if empty_run >= _EMPTY_RUN_LIMIT:
                    break
                continue
            empty_run = 0
            record = {
                canonical[canon]: (row[idx] if idx < len(row) else None)
                for canon, idx in mapping.items()
            }
            if not any(
                v is not None and str(v).strip() != "" for v in record.values()
            ):
                continue
            records.append(record)
    finally:
        wb.close()

    if not records:
        return pd.DataFrame(columns=columns)

    frame = pd.DataFrame(records, columns=columns)
    sr_column = columns[0]
    sr_values = list(frame[sr_column])
    if "sr" not in mapping or all(v is None or str(v).strip() == "" for v in sr_values):
        frame[sr_column] = range(1, len(frame) + 1)
    return frame


def _read_va_data(va_path: Path, sheet_name: str | None = None) -> tuple[pd.DataFrame, dict[str, int]]:
    """Read VA data and risk summary from a generated VA Excel report."""
    va_word_df = pd.DataFrame()
    va_risk_summary = {
        "Critical": 0,
        "High": 0,
        "Medium": 0,
        "Low": 0,
        "Grand Total": 0,
    }
    try:
        va_word_df = _read_findings(va_path, sheet_name, "va")
        va_risk_summary = _read_pivot_summary(
            va_path, {"Critical": 0, "High": 0, "Medium": 0, "Low": 0, "Grand Total": 0}
        )
    except Exception as e:
        logger.warning("Could not read VA Excel for Word report: %s", e)

    return va_word_df, va_risk_summary


def _read_ca_data(ca_path: Path, sheet_name: str | None = None) -> tuple[pd.DataFrame, dict[str, int]]:
    """Read CA data and risk summary from a generated CA Excel report."""
    ca_word_df = pd.DataFrame()
    ca_risk_summary = {"FAILED": 0, "WARNING": 0, "Grand Total": 0}
    try:
        ca_word_df = _read_findings(ca_path, sheet_name, "ca")
        ca_risk_summary = _read_pivot_summary(
            ca_path, {"FAILED": 0, "WARNING": 0, "Grand Total": 0}
        )
    except Exception as e:
        logger.warning("Could not read CA Excel for Word report: %s", e)

    return ca_word_df, ca_risk_summary


@router.post("/word")
async def generate_word(
    files: list[UploadFile] = File(...),
    client_name: str = Form(""),
    client_short_name: str = Form(""),
    security_tester: str = Form(""),
    reviewed_by: str = Form(""),
    device_type: str = Form(""),
    scope: str = Form("Server"),
    phase: str = Form("First"),
    report_type: str = Form("First"),
    report_number: str = Form("1.0"),
    report_date: str = Form(""),
    assessment_start_date: str = Form(""),
    assessment_finish_date: str = Form(""),
    final_retesting_start: str = Form(""),
    final_retesting_finish: str = Form(""),
    released_date: str = Form(""),
    spokesperson_name: str = Form(""),
    spokesperson_designation: str = Form(""),
    spokesperson_email: str = Form(""),
    senior_name: str = Form(""),
    approved_by: str = Form("Mr. Vijay Sawant"),
    # New fields
    report_release_date: str = Form(""),
    period: str = Form(""),
    document_id: str = Form(""),
    change_history_version: str = Form(""),
    change_history_date: str = Form(""),
    change_history_remarks: str = Form(""),
    distribution_name: str = Form(""),
    distribution_organization: str = Form(""),
    distribution_designation: str = Form(""),
    distribution_email: str = Form(""),
    # Auditing Team - Row 1
    auditor_1_name: str = Form(""),
    auditor_1_designation: str = Form(""),
    auditor_1_email: str = Form(""),
    auditor_1_qualifications: str = Form(""),
    auditor_1_cert_in: str = Form("Yes"),
    # Auditing Team - Row 2
    auditor_2_name: str = Form(""),
    auditor_2_designation: str = Form(""),
    auditor_2_email: str = Form(""),
    auditor_2_qualifications: str = Form(""),
    auditor_2_cert_in: str = Form("Yes"),
    current_user: dict = Depends(get_current_user),
):
    """Generate Word report from uploaded VA/CA Excel reports. Returns download URL."""
    temp_dir = Path(tempfile.mkdtemp())
    cleanup_paths: list[Path] = []

    try:
        rejected: list[str] = []
        uploads: list[tuple[Path, str]] = []  # (readable path, original filename)
        for upload in files:
            content = await upload.read()
            suffix = Path(upload.filename).suffix if upload.filename else ".xlsx"
            raw_name = upload.filename or f"upload_{len(uploads)}{suffix}"
            # Keep the original name when possible, but never let two uploads
            # with the same name overwrite each other in the temp directory.
            safe_name = Path(raw_name.replace("\\", "/")).name or f"upload_{len(uploads)}{suffix}"
            tmp_path = temp_dir / safe_name
            if tmp_path.exists():
                tmp_path = temp_dir / f"{len(uploads)}_{safe_name}"
            tmp_path.write_bytes(content)
            cleanup_paths.append(tmp_path)

            read_path, error = _prepare_workbook(tmp_path)
            if read_path is not tmp_path:
                cleanup_paths.append(read_path)
            if error:
                rejected.append(f"'{safe_name}': {error}")
                continue
            uploads.append((read_path, safe_name))

        va_files: list[tuple[Path, str, str | None]] = []  # (path, upload name, sheet)
        ca_files: list[tuple[Path, str, str | None]] = []
        for p, display_name in uploads:
            classified = _classify_report(p)
            if classified and classified[0] == "va":
                va_files.append((p, display_name, classified[1]))
                continue
            if classified and classified[0] == "ca":
                ca_files.append((p, display_name, classified[1]))
                continue
            try:
                wb = load_workbook(p, read_only=True, data_only=True)
                names = [str(s) for s in wb.sheetnames]
                wb.close()
                rejected.append(
                    f"'{display_name}': sheets are {names} (expected VA/CA scan columns or a "
                    "'VA Report' / 'CA_Report' sheet)"
                )
            except Exception as exc:
                rejected.append(
                    f"'{display_name}': could not be read as an Excel file ({type(exc).__name__})"
                )

        if not va_files and not ca_files:
            detail = (
                "No valid VA or CA report files found. Upload .xlsx files containing a "
                "sheet named 'VA Report' (VA) or 'CA_Report' (CA)."
            )
            if rejected:
                detail += " Problems found: " + "; ".join(rejected)
            return JSONResponse(
                status_code=400,
                content={"detail": detail},
            )

        va_path, _va_name, va_sheet = va_files[0] if va_files else (None, None, None)
        ca_path, _ca_name, ca_sheet = ca_files[0] if ca_files else (None, None, None)
        excel_meta = _read_va_metadata(va_path, va_sheet) if va_path else {}
        if not excel_meta.get("client_name") and ca_path:
            ca_meta = _read_ca_metadata(ca_path, ca_sheet)
            if ca_meta.get("client_name"):
                excel_meta = {**excel_meta, **ca_meta}

        if not client_name and excel_meta.get("client_name"):
            client_name = excel_meta["client_name"]
        if not security_tester and excel_meta.get("security_tester"):
            security_tester = excel_meta["security_tester"]
        if not reviewed_by and excel_meta.get("reviewed_by"):
            reviewed_by = excel_meta["reviewed_by"]
        if not report_date and excel_meta.get("report_date"):
            report_date = excel_meta["report_date"]

        try:
            parsed_report_date = date.fromisoformat(report_date) if report_date else date.today()
        except ValueError:
            parsed_report_date = date.today()

        metadata = EngagementMetadata(
            client_name=client_name,
            security_tester=security_tester,
            reviewed_by=reviewed_by,
            report_date=parsed_report_date,
            report_version=report_number,
            scope_label=scope,
            phase_label=phase,
            default_device_type=device_type,
            report_type=report_type,
            client_short_name=client_short_name,
            assessment_start_date=assessment_start_date,
            assessment_finish_date=assessment_finish_date,
            final_retesting_start=final_retesting_start,
            final_retesting_finish=final_retesting_finish,
            released_date=released_date,
            spokesperson_name=spokesperson_name,
            spokesperson_designation=spokesperson_designation,
            spokesperson_email=spokesperson_email,
            senior_name=senior_name,
            approved_by=approved_by,
            # New fields
            report_release_date=report_release_date,
            period=period,
            document_id=document_id,
            change_history_version=change_history_version,
            change_history_date=change_history_date,
            change_history_remarks=change_history_remarks,
            distribution_name=distribution_name,
            distribution_organization=distribution_organization,
            distribution_designation=distribution_designation,
            distribution_email=distribution_email,
            # Auditing Team - Row 1
            auditor_1_name=auditor_1_name,
            auditor_1_designation=auditor_1_designation,
            auditor_1_email=auditor_1_email,
            auditor_1_qualifications=auditor_1_qualifications,
            auditor_1_cert_in=auditor_1_cert_in,
            # Auditing Team - Row 2
            auditor_2_name=auditor_2_name,
            auditor_2_designation=auditor_2_designation,
            auditor_2_email=auditor_2_email,
            auditor_2_qualifications=auditor_2_qualifications,
            auditor_2_cert_in=auditor_2_cert_in,
        )

        project_root = Path(__file__).resolve().parent.parent.parent.parent.parent
        word_template_path = project_root / "templates" / "Word file.docx"

        va_dfs: list[pd.DataFrame] = []
        va_risk_summary = {"Critical": 0, "High": 0, "Medium": 0, "Low": 0, "Grand Total": 0}
        va_used: list[str] = []
        for f_path, f_name, f_sheet in va_files:
            df, rs = _read_va_data(f_path, f_sheet)
            if not df.empty:
                va_dfs.append(df)
                va_used.append(f_name)
            else:
                rejected.append(f"'{f_name}': recognised as VA but has no finding rows")
            for key in va_risk_summary:
                va_risk_summary[key] += rs.get(key, 0)

        ca_dfs: list[pd.DataFrame] = []
        ca_risk_summary = {"FAILED": 0, "WARNING": 0, "Grand Total": 0}
        ca_used: list[str] = []
        for f_path, f_name, f_sheet in ca_files:
            df, rs = _read_ca_data(f_path, f_sheet)
            if not df.empty:
                ca_dfs.append(df)
                ca_used.append(f_name)
            else:
                rejected.append(f"'{f_name}': recognised as CA but has no finding rows")
            for key in ca_risk_summary:
                ca_risk_summary[key] += rs.get(key, 0)

        if va_dfs:
            va_word_df = pd.concat(va_dfs, ignore_index=True)
        else:
            va_word_df = pd.DataFrame()
        if ca_dfs:
            ca_word_df = pd.concat(ca_dfs, ignore_index=True)
        else:
            ca_word_df = pd.DataFrame()

        va_risk_summary["Grand Total"] = sum(
            v for k, v in va_risk_summary.items() if k != "Grand Total"
        )
        ca_risk_summary["Grand Total"] = sum(
            v for k, v in ca_risk_summary.items() if k != "Grand Total"
        )

        # Excel report (cell C5) takes precedence over the form value.
        filename_client_name = excel_meta.get("client_name") or metadata.client_name
        try:
            word_filename = build_word_filename(
                filename_client_name, metadata.report_date, metadata.report_version
            )
        except NamingConventionError as exc:
            return JSONResponse(status_code=400, content={"detail": str(exc)})
        output_dir = Path(tempfile.mkdtemp())
        word_output_path = output_dir / word_filename

        word_path = build_word_report(
            template_path=word_template_path,
            output_path=word_output_path,
            metadata=metadata,
            va_df=va_word_df,
            ca_df=ca_word_df,
            va_risk_summary=va_risk_summary,
            ca_risk_summary=ca_risk_summary,
            va_dfs=va_dfs,
            ca_dfs=ca_dfs,
        )

        session_id = create_session()
        store_file(session_id, "word", word_path)

        return JSONResponse(
            content={
                "session_id": session_id,
                "files": {"word_report": f"/api/download/{session_id}/word"},
                "used": {"va": va_used, "ca": ca_used},
                "ignored": rejected,
            }
        )
    finally:
        import shutil
        for f in cleanup_paths:
            f.unlink(missing_ok=True)
        shutil.rmtree(temp_dir, ignore_errors=True)
