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
from ...naming.filename_builder import _sanitize_filename
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
        if _is_va_report(tmp_path):
            meta = _read_va_metadata(tmp_path)
        elif _is_ca_report(tmp_path):
            meta = _read_ca_metadata(tmp_path)

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


def _get_sheet(workbook, wanted: str):
    """Return a worksheet by normalized name (case/space/underscore tolerant)."""
    wanted_norm = _normalize_sheet_name(wanted)
    for sheet in workbook.sheetnames:
        if _normalize_sheet_name(sheet) == wanted_norm:
            return workbook[sheet]
    raise KeyError(wanted)


def _report_kind(path: Path) -> str | None:
    """Return 'va', 'ca', or None for an uploaded Excel report.

    Detection is tolerant of casing, spacing and underscore differences, so
    'VA Report', 'VA_Report', 'va report', 'VA REPORT' all count as VA, and
    'CA_Report', 'CA report', 'ca_report' all count as CA.
    """
    try:
        wb = load_workbook(path, read_only=True, data_only=True)
        try:
            names = {_normalize_sheet_name(s) for s in wb.sheetnames}
        finally:
            wb.close()
        if "va report" in names:
            return "va"
        if "ca report" in names:
            return "ca"
        return None
    except Exception:
        return None


def _is_va_report(path: Path) -> bool:
    """Check if an Excel file is a VA report by looking for a 'VA Report' sheet."""
    return _report_kind(path) == "va"


def _is_ca_report(path: Path) -> bool:
    """Check if an Excel file is a CA report by looking for a 'CA_Report' sheet."""
    return _report_kind(path) == "ca"


def _read_va_metadata(va_path: Path) -> dict:
    """Read engagement metadata from VA Excel report header (rows 5-10, column C)."""
    meta = {}
    try:
        wb = load_workbook(va_path, read_only=True, data_only=True)
        ws = _get_sheet(wb, "VA Report")
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


def _read_ca_metadata(ca_path: Path) -> dict:
    """Read engagement metadata from CA Excel report header (rows 5-10, column C)."""
    meta = {}
    try:
        wb = load_workbook(ca_path, read_only=True, data_only=True)
        ws = _get_sheet(wb, "CA_Report")
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


def _read_va_data(va_path: Path) -> tuple[pd.DataFrame, dict[str, int]]:
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
        va_wb = load_workbook(va_path, read_only=True, data_only=True)
        va_ws = _get_sheet(va_wb, "VA Report")
        va_data = []
        for row in va_ws.iter_rows(min_row=14, max_col=9, values_only=True):
            if row[0] is not None:
                va_data.append(row)
        if va_data:
            va_cols = [
                "Sr. no",
                "Vulnerbility Title",
                "Description",
                "Risk",
                "Host",
                "Port",
                "Recommendation ",
                "Reference",
                "CVE",
            ]
            va_word_df = pd.DataFrame(va_data, columns=va_cols)
        va_wb.close()

        va_risk_summary = _read_pivot_summary(
            va_path, {"Critical": 0, "High": 0, "Medium": 0, "Low": 0, "Grand Total": 0}
        )
    except Exception as e:
        logger.warning("Could not read VA Excel for Word report: %s", e)

    return va_word_df, va_risk_summary


def _read_ca_data(ca_path: Path) -> tuple[pd.DataFrame, dict[str, int]]:
    """Read CA data and risk summary from a generated CA Excel report."""
    ca_word_df = pd.DataFrame()
    ca_risk_summary = {"FAILED": 0, "WARNING": 0, "Grand Total": 0}
    try:
        ca_wb = load_workbook(ca_path, read_only=True, data_only=True)
        ca_ws = _get_sheet(ca_wb, "CA_Report")
        ca_data = []
        for row in ca_ws.iter_rows(min_row=14, max_col=7, values_only=True):
            if row[0] is not None:
                ca_data.append(row)
        if ca_data:
            ca_cols = [
                "Sr.No.",
                "Title",
                "Host",
                "Description",
                "Solution",
                "Impact",
                "Risk",
            ]
            ca_word_df = pd.DataFrame(ca_data, columns=ca_cols)
        ca_wb.close()

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
    saved_files = []

    try:
        for upload in files:
            content = await upload.read()
            suffix = Path(upload.filename).suffix if upload.filename else ".xlsx"
            tmp_path = temp_dir / (upload.filename or f"upload_{len(saved_files)}{suffix}")
            tmp_path.write_bytes(content)
            saved_files.append(tmp_path)

        va_files = []
        ca_files = []
        rejected: list[str] = []
        for p in saved_files:
            kind = _report_kind(p)
            if kind == "va":
                va_files.append(p)
                continue
            if kind == "ca":
                ca_files.append(p)
                continue
            try:
                wb = load_workbook(p, read_only=True, data_only=True)
                names = [str(s) for s in wb.sheetnames]
                wb.close()
                rejected.append(
                    f"'{p.name}': sheets are {names} (expected a 'VA Report' or 'CA_Report' sheet)"
                )
            except Exception as exc:
                rejected.append(
                    f"'{p.name}': could not be read as an Excel file ({type(exc).__name__})"
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

        va_path = va_files[0] if va_files else None
        excel_meta = _read_va_metadata(va_path) if va_path else {}

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
        for f in va_files:
            df, rs = _read_va_data(f)
            if not df.empty:
                va_dfs.append(df)
            for key in va_risk_summary:
                va_risk_summary[key] += rs.get(key, 0)

        ca_dfs: list[pd.DataFrame] = []
        ca_risk_summary = {"FAILED": 0, "WARNING": 0, "Grand Total": 0}
        for f in ca_files:
            df, rs = _read_ca_data(f)
            if not df.empty:
                ca_dfs.append(df)
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

        scope_clean = _sanitize_filename(metadata.scope_label)
        phase_clean = _sanitize_filename(metadata.phase_label)
        client_clean = _sanitize_filename(metadata.client_name.replace(" ", "_"))
        year = str(metadata.report_date.year)
        version_str = str(metadata.report_version).replace(".", "_")
        word_parts = [
            "VA_CA",
            scope_clean,
            phase_clean,
            "Audit_Report",
            client_clean,
        ]
        if metadata.entity_codes:
            for code in metadata.entity_codes:
                word_parts.append(_sanitize_filename(code))
        word_parts.append(year)
        word_parts.append(f"V{version_str}")
        word_filename = "_".join(word_parts) + ".docx"
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
            }
        )
    finally:
        import shutil
        for f in saved_files:
            f.unlink(missing_ok=True)
        temp_dir.rmdir()
