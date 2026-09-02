"""POST /api/merge-csv — Merge multiple raw scan files."""

from __future__ import annotations

import io
import tempfile
from pathlib import Path

import pandas as pd
from fastapi import APIRouter, Depends, File, UploadFile
from fastapi.responses import StreamingResponse

from ...ingestion.merge_scan_exports import merge_scan_exports
from ..deps import get_current_user

router = APIRouter(tags=["merge"])


@router.post("/merge-csv")
async def merge_csv(
    files: list[UploadFile] = File(...),
    current_user: dict = Depends(get_current_user),
):
    """Merge multiple raw scan files into one.

    Accepts .csv and .xlsx files with the 17-column Nessus schema.
    Returns the merged file as a downloadable .xlsx.
    """
    # Save uploaded files to temp directory
    temp_dir = Path(tempfile.mkdtemp())
    saved_files = []

    try:
        for upload in files:
            content = await upload.read()
            suffix = Path(upload.filename).suffix if upload.filename else ".xlsx"
            tmp_path = temp_dir / (upload.filename or f"upload_{len(saved_files)}{suffix}")
            tmp_path.write_bytes(content)
            saved_files.append(tmp_path)

        # Use the merge function
        merged = merge_scan_exports(saved_files, dedupe=False)

        if merged is None:
            return {"error": "No valid files to merge. Check that files have the expected 17-column schema."}

        # Write to buffer for download
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            merged.to_excel(writer, index=False, sheet_name="RAW File")
        buf.seek(0)

        return StreamingResponse(
            buf,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": "attachment; filename=merged_raw.xlsx"},
        )
    finally:
        # Cleanup temp files
        import shutil
        shutil.rmtree(temp_dir, ignore_errors=True)
