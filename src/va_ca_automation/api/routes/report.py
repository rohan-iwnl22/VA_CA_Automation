"""POST /api/report — Generate VA/CA Excel reports (individual downloads)."""

from __future__ import annotations

import tempfile
import time
from pathlib import Path

from datetime import date

from fastapi import APIRouter, Depends, File, Form, UploadFile
from fastapi.responses import JSONResponse

from ...ingestion.raw_file_loader import load_raw_file
from ...metadata.engagement_metadata import EngagementMetadata
from ...pipelines.ca_pipeline import run_ca_pipeline
from ...pipelines.va_pipeline import run_va_pipeline
from ..deps import get_current_user
from ..temp_registry import create_session, store_file

router = APIRouter(tags=["report"])


def _build_metadata(
    client_name: str,
    client_short_name: str,
    security_tester: str,
    reviewed_by: str,
    device_type: str,
    scope: str,
    phase: str,
    report_type: str,
    report_number: str,
    report_date: date,
    assessment_start_date: str,
    assessment_finish_date: str,
    final_retesting_start: str,
    final_retesting_finish: str,
    released_date: str,
    spokesperson_name: str,
    spokesperson_designation: str,
    spokesperson_email: str,
    senior_name: str,
    approved_by: str,
) -> EngagementMetadata:
    """Build EngagementMetadata from form fields."""
    return EngagementMetadata(
        client_name=client_name,
        security_tester=security_tester,
        reviewed_by=reviewed_by,
        report_date=report_date,
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
    )


@router.post("/report")
async def generate_report(
    file: UploadFile = File(...),
    client_name: str = Form(""),
    client_short_name: str = Form(""),
    security_tester: str = Form(""),
    reviewed_by: str = Form(""),
    device_type: str = Form(""),
    scope: str = Form("Server"),
    phase: str = Form("First"),
    report_type: str = Form("First"),
    report_number: str = Form("1.0"),
    report_date: date = Form(...),
    assessment_start_date: str = Form(""),
    assessment_finish_date: str = Form(""),
    final_retesting_start: str = Form(""),
    final_retesting_finish: str = Form(""),
    released_date: str = Form(""),
    spokesperson_name: str = Form(""),
    spokesperson_designation: str = Form(""),
    spokesperson_email: str = Form(""),
    senior_name: str = Form(""),
    approved_by: str = Form("Default"),
    current_user: dict = Depends(get_current_user),
):
    """Generate VA and CA Excel reports. Returns download URLs for each file."""
    content = await file.read()
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    tmp.write(content)
    tmp.close()

    output_dir = Path(tempfile.mkdtemp())

    try:
        metadata = _build_metadata(
            client_name,
            client_short_name,
            security_tester,
            reviewed_by,
            device_type,
            scope,
            phase,
            report_type,
            report_number,
            report_date,
            assessment_start_date,
            assessment_finish_date,
            final_retesting_start,
            final_retesting_finish,
            released_date,
            spokesperson_name,
            spokesperson_designation,
            spokesperson_email,
            senior_name,
            approved_by,
        )

        project_root = Path(__file__).resolve().parent.parent.parent.parent.parent
        template_path = project_root / "templates" / "va_report_template.xlsx"
        ca_template_path = project_root / "templates" / "ca_report_template.xlsx"

        # Run VA pipeline
        va_path = run_va_pipeline(
            raw_file_path=Path(tmp.name),
            template_path=template_path,
            metadata=metadata,
            output_dir=output_dir,
            generate_text_join=True,
        )
        va_tj_path = va_path.with_name(va_path.stem + "_TextJoin" + va_path.suffix)

        # Run CA pipeline
        raw_df = load_raw_file(Path(tmp.name))
        ca_path = run_ca_pipeline(
            raw_df=raw_df,
            metadata=metadata,
            ca_template_path=ca_template_path,
            output_dir=output_dir,
            generate_text_join=True,
        )
        ca_tj_path = (
            ca_path.with_name(ca_path.stem + "_TextJoin" + ca_path.suffix)
            if ca_path
            else None
        )

        # Create session and store file paths
        session_id = create_session()
        store_file(session_id, "va_normal", va_path)
        store_file(session_id, "va_textjoin", va_tj_path)
        if ca_path:
            store_file(session_id, "ca_normal", ca_path)
        if ca_tj_path:
            store_file(session_id, "ca_textjoin", ca_tj_path)

        files = {
            "va_normal": f"/api/download/{session_id}/va_normal",
            "va_textjoin": f"/api/download/{session_id}/va_textjoin",
        }
        if ca_path:
            files["ca_normal"] = f"/api/download/{session_id}/ca_normal"
        if ca_tj_path:
            files["ca_textjoin"] = f"/api/download/{session_id}/ca_textjoin"

        return JSONResponse(content={"session_id": session_id, "files": files})
    finally:
        for _ in range(3):
            try:
                Path(tmp.name).unlink(missing_ok=True)
                break
            except PermissionError:
                time.sleep(0.1)
