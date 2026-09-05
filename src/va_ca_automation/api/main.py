"""FastAPI application factory."""

from __future__ import annotations

from contextlib import asynccontextmanager
from pathlib import Path

from dotenv import load_dotenv

# Load .env file from project root
_project_root = Path(__file__).resolve().parent.parent.parent.parent
load_dotenv(_project_root / ".env")

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles

from .db import init_db, close_db
from .auth import router as auth_router
from .routes import download, merge_csv, report, word, reverse_textjoin

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
    app.include_router(auth_router, prefix="/api")
    app.include_router(merge_csv.router, prefix="/api")
    app.include_router(report.router, prefix="/api")
    app.include_router(word.router, prefix="/api")
    app.include_router(download.router, prefix="/api")
    app.include_router(reverse_textjoin.router, prefix="/api")

    @app.get("/")
    async def root():
        """Serve the dashboard."""
        index = STATIC_DIR / "dashboard.html"
        if index.exists():
            return FileResponse(str(index))
        return {"message": "VA/CA Report Automation API"}

    @app.get("/login")
    async def login_page():
        """Serve the login page."""
        login = STATIC_DIR / "index.html"
        if login.exists():
            return FileResponse(str(login))
        return {"message": "Login page not found"}

    return app


app = create_app()
