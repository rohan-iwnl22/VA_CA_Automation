# Task 04: Create FastAPI Application

## Objective
Create the FastAPI application with CORS, static file serving, lifespan management, and router mounting.

## Files to Create
- `src/va_ca_automation/api/main.py`
- `src/va_ca_automation/api/schemas.py`
- `src/va_ca_automation/api/deps.py`

## Files to Modify
- `pyproject.toml` — Ensure `fastapi` and `uvicorn` are in dependencies

## Implementation Steps

### Step 1: Create `src/va_ca_automation/api/schemas.py`

Pydantic models for API request/response:

```python
"""Pydantic request/response models."""

from __future__ import annotations
from pydantic import BaseModel


class LoginRequest(BaseModel):
    username: str
    password: str


class LoginResponse(BaseModel):
    access_token: str
    token_type: str = "bearer"


class CreateUserRequest(BaseModel):
    username: str
    password: str
    role: str = "user"


class UserResponse(BaseModel):
    id: int
    username: str
    role: str


class ReportRequest(BaseModel):
    client_name: str
    client_short_name: str = ""
    security_tester: str
    reviewed_by: str
    device_type: str = ""
    scope: str = "Server"
    phase: str = "First"
    report_type: str = "First"          # "First" or "Final"
    report_number: str = "1.0"
    assessment_start_date: str = ""
    assessment_finish_date: str = ""
    final_retesting_start: str = ""
    final_retesting_finish: str = ""
    released_date: str = ""
    spokesperson_name: str = ""
    spokesperson_designation: str = ""
    spokesperson_email: str = ""
    senior_name: str = ""
    approved_by: str = "Default"
```

### Step 2: Create `src/va_ca_automation/api/deps.py`

Authentication dependency:

```python
"""Authentication dependencies for FastAPI."""

from __future__ import annotations
import os
from fastapi import Depends, HTTPException, status
from fastapi.security import HTTPBearer, HTTPAuthorizationCredentials
from jose import JWTError, jwt

security = HTTPBearer()
JWT_SECRET = os.getenv("JWT_SECRET", "dev-secret-change-me")
ALGORITHM = "HS256"


def get_current_user(credentials: HTTPAuthorizationCredentials = Depends(security)):
    """Validate JWT token and return payload."""
    token = credentials.credentials
    try:
        payload = jwt.decode(token, JWT_SECRET, algorithms=[ALGORITHM])
        username: str = payload.get("sub")
        role: str = payload.get("role")
        if username is None:
            raise HTTPException(status_code=401, detail="Invalid token")
        return {"username": username, "role": role}
    except JWTError:
        raise HTTPException(status_code=401, detail="Invalid token")
```

### Step 3: Create `src/va_ca_automation/api/main.py`

```python
"""FastAPI application factory."""

from __future__ import annotations
import os
from pathlib import Path
from contextlib import asynccontextmanager
from fastapi import FastAPI
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware

from .db import init_db, close_db
from .routes import merge_csv, report, word

STATIC_DIR = Path(__file__).parent.parent.parent.parent / "static"


@asynccontextmanager
async def lifespan(app: FastAPI):
    """Startup/shutdown lifecycle."""
    await init_db()
    yield
    await close_db()


def create_app() -> FastAPI:
    """Create and configure the FastAPI application."""
    app = FastAPI(
        title="VA/CA Report Automation",
        version="0.1.0",
        lifespan=lifespan,
    )

    # CORS
    app.add_middleware(
        CORSMiddleware,
        allow_origins=["*"],
        allow_credentials=True,
        allow_methods=["*"],
        allow_headers=["*"],
    )

    # Mount static files
    if STATIC_DIR.exists():
        app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")

    # Include routers
    from .auth import router as auth_router
    app.include_router(auth_router, prefix="/api")
    app.include_router(merge_csv.router, prefix="/api")
    app.include_router(report.router, prefix="/api")
    app.include_router(word.router, prefix="/api")

    @app.get("/")
    async def root():
        """Serve the dashboard."""
        from fastapi.responses import FileResponse
        index = STATIC_DIR / "dashboard.html"
        if index.exists():
            return FileResponse(str(index))
        return {"message": "VA/CA Report Automation API"}

    @app.get("/login")
    async def login_page():
        """Serve the login page."""
        from fastapi.responses import FileResponse
        login = STATIC_DIR / "index.html"
        if login.exists():
            return FileResponse(str(login))
        return {"message": "Login page not found"}

    return app


app = create_app()
```

### Step 4: Add run script

Add to `pyproject.toml` under `[project.scripts]`:
```toml
va-ca-api = "va_ca_automation.api.main:app"
```

Or run with: `uvicorn va_ca_automation.api.main:app --reload`

## Acceptance Criteria
- [ ] FastAPI app created with CORS enabled
- [ ] Static files mounted at `/static`
- [ ] Login page served at `/login`
- [ ] Dashboard served at `/`
- [ ] All routers included
- [ ] Lifespan manages DB init/close
- [ ] `uvicorn va_ca_automation.api.main:app --reload` starts the server
