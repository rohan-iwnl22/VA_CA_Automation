# Task 06: Create API Routes (Merge, Report, Word)

## Objective
Create the three main API endpoints that wrap existing pipeline functions. Reports are downloaded individually, not as a ZIP.

## Files to Create
- `src/va_ca_automation/api/routes/__init__.py`
- `src/va_ca_automation/api/routes/merge_csv.py`
- `src/va_ca_automation/api/routes/report.py`
- `src/va_ca_automation/api/routes/word.py`

## Endpoint Details

### `POST /api/merge-csv`
- **Input:** Multiple file uploads (`.xlsx` or `.csv`)
- **Process:** Load each file with `load_raw_file()`, concatenate with `pd.concat()`
- **Output:** Single merged `.xlsx` file as download
- **Auth:** Required (JWT)

### `POST /api/report`
- **Input:** Merged file upload + JSON metadata (form fields)
- **Process:**
  1. Save uploaded file to temp directory
  2. Build `EngagementMetadata` from request data
  3. Run `run_va_pipeline()` → generates VA normal + TextJoin
  4. Run `run_ca_pipeline()` → generates CA normal + TextJoin
  5. Store generated file paths in a temp registry (dict with session ID)
- **Output:** JSON response with download URLs for each file:
  ```json
  {
    "session_id": "abc123",
    "files": {
      "va_normal": "/api/download/abc123/va_normal",
      "va_textjoin": "/api/download/abc123/va_textjoin",
      "ca_normal": "/api/download/abc123/ca_normal",
      "ca_textjoin": "/api/download/abc123/ca_textjoin"
    }
  }
  ```
- **Auth:** Required (JWT)

### `GET /api/download/{session_id}/{file_type}`
- **Input:** Session ID + file type (va_normal, va_textjoin, ca_normal, ca_textjoin)
- **Process:** Look up file path in temp registry, return file
- **Output:** Individual `.xlsx` file download
- **Auth:** Required (JWT)

### `POST /api/word`
- **Input:** Merged file upload + JSON metadata (form fields)
- **Process:**
  1. Save uploaded file to temp directory
  2. Build `EngagementMetadata` from request data
  3. Run VA pipeline to generate Excel reports
  4. Read back Excel data for Word population
  5. Run `build_word_report()`
  6. Store generated file path in temp registry
- **Output:** JSON response with download URL:
  ```json
  {
    "session_id": "def456",
    "files": {
      "word_report": "/api/download/def456/word"
    }
  }
  ```
- **Auth:** Required (JWT)

### `GET /api/download/{session_id}/word`
- **Input:** Session ID
- **Process:** Look up Word file path in temp registry, return file
- **Output:** Individual `.docx` file download
- **Auth:** Required (JWT)

## Flow Diagram

```
User uploads files + fills form
        │
        ▼
POST /api/report ──▶ Generates 4 Excel files
        │              Saves to temp dir
        │              Stores path in registry
        ▼
Returns JSON with download URLs
        │
        ▼
User clicks download for each file
        │
        ▼
GET /api/download/{session_id}/va_normal  ──▶ Returns .xlsx
GET /api/download/{session_id}/va_textjoin ──▶ Returns .xlsx
GET /api/download/{session_id}/ca_normal  ──▶ Returns .xlsx
GET /api/download/{session_id}/ca_textjoin ──▶ Returns .xlsx
```

## Implementation Steps

### Step 1: Create `routes/__init__.py` (empty)

### Step 2: Create temp file registry

Add to `src/va_ca_automation/api/db.py` or create a new `src/va_ca_automation/api/temp_registry.py`:

```python
"""In-memory temp file registry for report downloads."""

from __future__ import annotations
import uuid
import tempfile
from pathlib import Path
from datetime import datetime, timedelta

# {session_id: {"files": {type: path}, "created_at": datetime}}
_registry: dict[str, dict] = {}


def create_session() -> str:
    """Create a new session and return its ID."""
    session_id = str(uuid.uuid4())[:8]
    _registry[session_id] = {
        "files": {},
        "created_at": datetime.utcnow(),
    }
    return session_id


def store_file(session_id: str, file_type: str, file_path: Path) -> None:
    """Store a generated file in the registry."""
    if session_id in _registry:
        _registry[session_id]["files"][file_type] = str(file_path)


def get_file(session_id: str, file_type: str) -> Path | None:
    """Retrieve a file path from the registry."""
    session = _registry.get(session_id)
    if session:
        path_str = session["files"].get(file_type)
        if path_str:
            path = Path(path_str)
            if path.exists():
                return path
    return None


def cleanup_session(session_id: str) -> None:
    """Delete all files in a session and remove from registry."""
    session = _registry.pop(session_id, None)
    if session:
        for file_path in session["files"].values():
            Path(file_path).unlink(missing_ok=True)


def cleanup_expired(max_age_minutes: int = 30) -> None:
    """Remove sessions older than max_age."""
    now = datetime.utcnow()
    expired = [
        sid for sid, data in _registry.items()
        if (now - data["created_at"]) > timedelta(minutes=max_age_minutes)
    ]
    for sid in expired:
        cleanup_session(sid)
```

### Step 3: Create `routes/merge_csv.py`

```python
"""POST /api/merge-csv — Merge multiple raw scan files."""

from __future__ import annotations
import io
import tempfile
from pathlib import Path
from fastapi import APIRouter, Depends, File, UploadFile
from fastapi.responses import StreamingResponse

from ..deps import get_current_user
from ...ingestion.raw_file_loader import load_raw_file
import pandas as pd

router = APIRouter(tags=["merge"])


@router.post("/merge-csv")
async def merge_csv(
    files: list[UploadFile] = File(...),
    current_user: dict = Depends(get_current_user),
):
    """Merge multiple raw scan files into one."""
    frames = []
    for upload in files:
        content = await upload.read()
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=Path(upload.filename).suffix)
        tmp.write(content)
        tmp.close()
        try:
            df = load_raw_file(Path(tmp.name))
            frames.append(df)
        finally:
            Path(tmp.name).unlink(missing_ok=True)

    if not frames:
        return {"error": "No valid files provided"}

    merged = pd.concat(frames, ignore_index=True)

    # Write to buffer
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        merged.to_excel(writer, index=False, sheet_name="RAW File")
    buf.seek(0)

    return StreamingResponse(
        buf,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=merged_raw.xlsx"},
    )
```

### Step 4: Create `routes/report.py`

```python
"""POST /api/report — Generate VA/CA Excel reports (individual downloads)."""

from __future__ import annotations
import tempfile
from pathlib import Path
from fastapi import APIRouter, Depends, File, UploadFile, Form
from fastapi.responses import JSONResponse

from ..deps import get_current_user
from ..temp_registry import create_session, store_file
from ...metadata.engagement_metadata import EngagementMetadata
from ...pipelines.va_pipeline import run_va_pipeline
from ...pipelines.ca_pipeline import run_ca_pipeline
from ...ingestion.raw_file_loader import load_raw_file

router = APIRouter(tags=["report"])


def _build_metadata(
    client_name, client_short_name, security_tester, reviewed_by,
    device_type, scope, phase, report_type, report_number,
    assessment_start_date, assessment_finish_date,
    final_retesting_start, final_retesting_finish, released_date,
    spokesperson_name, spokesperson_designation, spokesperson_email,
    senior_name, approved_by,
) -> EngagementMetadata:
    """Build EngagementMetadata from form fields."""
    return EngagementMetadata(
        client_name=client_name,
        security_tester=security_tester,
        reviewed_by=reviewed_by,
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
    # Save uploaded file
    content = await file.read()
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    tmp.write(content)
    tmp.close()

    output_dir = Path(tempfile.mkdtemp())

    try:
        metadata = _build_metadata(
            client_name, client_short_name, security_tester, reviewed_by,
            device_type, scope, phase, report_type, report_number,
            assessment_start_date, assessment_finish_date,
            final_retesting_start, final_retesting_finish, released_date,
            spokesperson_name, spokesperson_designation, spokesperson_email,
            senior_name, approved_by,
        )

        # Resolve template paths
        project_root = Path(__file__).resolve().parent.parent.parent.parent
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
        ca_tj_path = ca_path.with_name(ca_path.stem + "_TextJoin" + ca_path.suffix) if ca_path else None

        # Create session and store file paths
        session_id = create_session()
        store_file(session_id, "va_normal", va_path)
        store_file(session_id, "va_textjoin", va_tj_path)
        if ca_path:
            store_file(session_id, "ca_normal", ca_path)
        if ca_tj_path:
            store_file(session_id, "ca_textjoin", ca_tj_path)

        # Build download URLs
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
        Path(tmp.name).unlink(missing_ok=True)
```

### Step 5: Create `routes/word.py`

```python
"""POST /api/word — Generate Word report."""

from __future__ import annotations
import tempfile
import logging
from pathlib import Path
from fastapi import APIRouter, Depends, File, UploadFile, Form
from fastapi.responses import JSONResponse
import pandas as pd
from openpyxl import load_workbook

from ..deps import get_current_user
from ..temp_registry import create_session, store_file
from ...metadata.engagement_metadata import EngagementMetadata
from ...pipelines.va_pipeline import run_va_pipeline
from ...pipelines.ca_pipeline import run_ca_pipeline
from ...word_writer.word_report_builder import build_word_report
from ...ingestion.raw_file_loader import load_raw_file

router = APIRouter(tags=["word"])
logger = logging.getLogger("va_ca_automation")


@router.post("/word")
async def generate_word(
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
    """Generate Word report. Returns download URL."""
    content = await file.read()
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    tmp.write(content)
    tmp.close()

    output_dir = Path(tempfile.mkdtemp())

    try:
        metadata = EngagementMetadata(
            client_name=client_name,
            security_tester=security_tester,
            reviewed_by=reviewed_by,
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

        project_root = Path(__file__).resolve().parent.parent.parent.parent
        template_path = project_root / "templates" / "va_report_template.xlsx"
        ca_template_path = project_root / "templates" / "ca_report_template.xlsx"
        word_template_path = project_root / "templates" / "Word file.docx"

        # Run VA pipeline to get Excel reports
        va_path = run_va_pipeline(
            raw_file_path=Path(tmp.name),
            template_path=template_path,
            metadata=metadata,
            output_dir=output_dir,
            generate_text_join=True,
        )

        # Run CA pipeline
        raw_df = load_raw_file(Path(tmp.name))
        ca_path = run_ca_pipeline(
            raw_df=raw_df,
            metadata=metadata,
            ca_template_path=ca_template_path,
            output_dir=output_dir,
            generate_text_join=True,
        )

        # Read VA data from generated Excel
        va_word_df = pd.DataFrame()
        va_risk_summary = {"Critical": 0, "High": 0, "Medium": 0, "Low": 0, "Grand Total": 0}
        try:
            va_wb = load_workbook(va_path, read_only=True, data_only=True)
            va_ws = va_wb["VA Report"]
            va_data = []
            for row in va_ws.iter_rows(min_row=14, max_col=9, values_only=True):
                if row[0] is not None:
                    va_data.append(row)
            if va_data:
                va_cols = ["Sr. no", "Vulnerbility Title", "Description", "Risk", "Host", "Port", "Recommendation ", "Reference", "CVE"]
                va_word_df = pd.DataFrame(va_data, columns=va_cols)
            va_wb.close()

            va_summary_wb = load_workbook(va_path, read_only=True, data_only=True)
            va_summary_ws = va_summary_wb["Summary"]
            for row in va_summary_ws.iter_rows(min_row=16, max_row=20, min_col=5, max_col=6, values_only=True):
                if row[0] and row[1] is not None:
                    label = str(row[0]).strip()
                    count = int(row[1]) if row[1] else 0
                    if label in va_risk_summary:
                        va_risk_summary[label] = count
                    elif label == "Grand Total":
                        va_risk_summary["Grand Total"] = count
            va_summary_wb.close()
        except Exception as e:
            logger.warning("Could not read VA Excel for Word report: %s", e)

        # Read CA data from generated Excel
        ca_word_df = pd.DataFrame()
        ca_risk_summary = {"FAILED": 0, "WARNING": 0, "Grand Total": 0}
        if ca_path and ca_path.exists():
            try:
                ca_wb = load_workbook(ca_path, read_only=True, data_only=True)
                ca_ws = ca_wb["CA_Report"]
                ca_data = []
                for row in ca_ws.iter_rows(min_row=14, max_col=6, values_only=True):
                    if row[0] is not None:
                        ca_data.append(row)
                if ca_data:
                    ca_cols = ["Sr.No.", "Title", "Host", "Description", "Solution", "Risk"]
                    ca_word_df = pd.DataFrame(ca_data, columns=ca_cols)
                ca_wb.close()

                ca_summary_wb = load_workbook(ca_path, read_only=True, data_only=True)
                ca_summary_ws = ca_summary_wb["Summary"]
                for row in ca_summary_ws.iter_rows(min_row=16, max_row=18, min_col=5, max_col=6, values_only=True):
                    if row[0] and row[1] is not None:
                        label = str(row[0]).strip()
                        count = int(row[1]) if row[1] else 0
                        if label in ca_risk_summary:
                            ca_risk_summary[label] = count
                        elif label == "Grand Total":
                            ca_risk_summary["Grand Total"] = count
                ca_summary_wb.close()
            except Exception as e:
                logger.warning("Could not read CA Excel for Word report: %s", e)

        # Build Word filename
        from ...naming.filename_builder import _sanitize_filename
        scope_clean = _sanitize_filename(metadata.scope_label)
        phase_clean = _sanitize_filename(metadata.phase_label)
        client_clean = _sanitize_filename(metadata.client_name.replace(" ", "_"))
        year = str(metadata.report_date.year)
        version_str = str(metadata.report_version).replace(".", "_")
        word_parts = ["VA_CA", scope_clean, phase_clean, "Audit_Report", client_clean]
        if metadata.entity_codes:
            for code in metadata.entity_codes:
                word_parts.append(_sanitize_filename(code))
        word_parts.append(year)
        word_parts.append(f"V{version_str}")
        word_filename = "_".join(word_parts) + ".docx"
        word_output_path = output_dir / word_filename

        # Generate Word report
        word_path = build_word_report(
            template_path=word_template_path,
            output_path=word_output_path,
            metadata=metadata,
            va_df=va_word_df,
            ca_df=ca_word_df,
            va_risk_summary=va_risk_summary,
            ca_risk_summary=ca_risk_summary,
        )

        # Create session and store file
        session_id = create_session()
        store_file(session_id, "word", word_path)

        return JSONResponse(content={
            "session_id": session_id,
            "files": {"word_report": f"/api/download/{session_id}/word"}
        })
    finally:
        Path(tmp.name).unlink(missing_ok=True)
```

### Step 6: Create download endpoint in `routes/__init__.py` or a separate file

```python
"""GET /api/download/{session_id}/{file_type} — Download individual report."""

from __future__ import annotations
from fastapi import APIRouter, Depends, HTTPException
from fastapi.responses import FileResponse

from ..deps import get_current_user
from ..temp_registry import get_file

router = APIRouter(tags=["download"])

FILE_TYPES = {
    "va_normal": ("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", ".xlsx"),
    "va_textjoin": ("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", ".xlsx"),
    "ca_normal": ("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", ".xlsx"),
    "ca_textjoin": ("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", ".xlsx"),
    "word": ("application/vnd.openxmlformats-officedocument.wordprocessingml.document", ".docx"),
}


@router.get("/download/{session_id}/{file_type}")
async def download_file(
    session_id: str,
    file_type: str,
    current_user: dict = Depends(get_current_user),
):
    """Download an individual report file."""
    if file_type not in FILE_TYPES:
        raise HTTPException(status_code=400, detail=f"Invalid file type: {file_type}")

    file_path = get_file(session_id, file_type)
    if not file_path:
        raise HTTPException(status_code=404, detail="File not found or expired")

    media_type, ext = FILE_TYPES[file_type]
    return FileResponse(
        path=str(file_path),
        media_type=media_type,
        filename=file_path.name,
    )
```

## Acceptance Criteria
- [ ] `/api/merge-csv` merges multiple files into one
- [ ] `/api/report` generates 4 Excel reports and returns JSON with download URLs
- [ ] `/api/word` generates Word report and returns JSON with download URL
- [ ] `/api/download/{session_id}/{file_type}` returns individual files
- [ ] All endpoints require JWT authentication
- [ ] Temp files cleaned up after 30 minutes
- [ ] File responses have correct Content-Type headers
- [ ] Downloads work individually (not bundled)
