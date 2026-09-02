"""In-memory temp file registry for report downloads."""

from __future__ import annotations

import uuid
from datetime import datetime, timedelta
from pathlib import Path

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
