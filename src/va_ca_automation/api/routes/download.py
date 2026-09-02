"""GET /api/download/{session_id}/{file_type} — Download individual report."""

from __future__ import annotations

from fastapi import APIRouter, Depends, HTTPException
from fastapi.responses import FileResponse

from ..deps import get_current_user
from ..temp_registry import get_file

router = APIRouter(tags=["download"])

FILE_TYPES = {
    "va_normal": (
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        ".xlsx",
    ),
    "va_textjoin": (
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        ".xlsx",
    ),
    "ca_normal": (
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        ".xlsx",
    ),
    "ca_textjoin": (
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        ".xlsx",
    ),
    "word": (
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        ".docx",
    ),
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

    media_type, _ext = FILE_TYPES[file_type]
    return FileResponse(
        path=str(file_path),
        media_type=media_type,
        filename=file_path.name,
    )
