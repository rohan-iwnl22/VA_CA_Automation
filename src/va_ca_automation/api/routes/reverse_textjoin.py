"""POST /api/reverse-textjoin — Build the non-TextJoin VA/CA report from any report."""

from __future__ import annotations

import logging
import re
import sys
import tempfile
from pathlib import Path

from fastapi import APIRouter, Depends, File, UploadFile
from fastapi.responses import JSONResponse

from ..deps import get_current_user
from ..temp_registry import create_session, store_file

router = APIRouter(tags=["reverse-textjoin"])
logger = logging.getLogger("va_ca_automation")


def _load_reversers():
    project_root = Path(__file__).resolve().parent.parent.parent.parent.parent
    if str(project_root) not in sys.path:
        sys.path.insert(0, str(project_root))

    from CA_TextJoin_Reverser import convert as ca_convert
    from CA_TextJoin_Reverser import detect_score as ca_detect
    from VA_TextJoin_Reverser import convert as va_convert
    from VA_TextJoin_Reverser import detect_score as va_detect

    return {"va": (va_convert, va_detect), "ca": (ca_convert, ca_detect)}


def _detect_report_type(path: Path) -> str:
    """Pick the reverser that matches the uploaded report, whichever sheet it uses."""
    reversers = _load_reversers()
    va_score = reversers["va"][1](str(path))
    ca_score = reversers["ca"][1](str(path))
    if ca_score > va_score:
        return "ca"
    if va_score > 0:
        return "va"
    raise ValueError(
        "Could not find a report table in that file. Expected a row of headers "
        "with a vulnerability title / host column (VA) or a title / host / "
        "status column (CA)."
    )


def _build_output_name(filename: str) -> str:
    """Turn 'Audit_V1.0.textjoin.xlsx' or 'Audit_V1.0_TextJoin.xlsx' into 'Audit_V1.0_Normal.xlsx'."""
    stem = Path(filename).stem if filename else "report"
    stem = re.sub(r"[ _.-]*[Tt]ext[Jj]oin$", "", stem)
    stem = re.sub(r"[ _.-]*[Nn]ormal$", "", stem)
    stem = stem.strip(" _-.") or "report"
    return f"{stem}_Normal.xlsx"


def _reverse_textjoin(input_path: Path, output_path: Path, report_type: str) -> dict:
    """Call the appropriate reverser based on report type."""
    convert, _ = _load_reversers()[report_type]
    return convert(str(input_path), str(output_path), progress_cb=lambda m: logger.info(m))


@router.post("/reverse-textjoin")
async def reverse_textjoin(
    file: UploadFile = File(...),
    current_user: dict = Depends(get_current_user),
):
    """Build the non-TextJoin (one row per host) VA/CA report from any report file."""
    temp_dir = Path(tempfile.mkdtemp())
    try:
        content = await file.read()
        suffix = Path(file.filename).suffix if file.filename else ".xlsx"
        if suffix.lower() not in (".xlsx", ".xlsm"):
            raise ValueError("Please upload an .xlsx report file.")
        input_path = temp_dir / (file.filename or f"upload{suffix}")
        input_path.write_bytes(content)

        report_type = _detect_report_type(input_path)

        output_dir = Path(tempfile.mkdtemp())
        output_path = output_dir / _build_output_name(file.filename or "report.xlsx")

        stats = _reverse_textjoin(input_path, output_path, report_type)

        session_id = create_session()
        file_type = "ca_normal_output" if report_type == "ca" else "va_normal_output"
        store_file(session_id, file_type, output_path)

        return JSONResponse(
            content={
                "session_id": session_id,
                "files": {"normal_report": f"/api/download/{session_id}/{file_type}"},
                "report_type": report_type,
                "filename": output_path.name,
                "stats": {
                    "old_count": stats.get("old_count", 0),
                    "new_count": stats.get("new_count", 0),
                    "risk_counts": stats.get("risk_counts", {}),
                    "status_counts": stats.get("status_counts", {}),
                },
            }
        )
    except ValueError as e:
        logger.warning("Reverse textjoin rejected input: %s", e)
        return JSONResponse(status_code=400, content={"detail": str(e)})
    except Exception as e:
        logger.error("Reverse textjoin failed: %s", e, exc_info=True)
        return JSONResponse(status_code=500, content={"detail": str(e)})
    finally:
        import shutil
        shutil.rmtree(temp_dir, ignore_errors=True)
