"""Authentication routes."""

from __future__ import annotations

from datetime import datetime, timedelta

from fastapi import APIRouter, Depends, HTTPException
from passlib.hash import bcrypt
from jose import jwt

from .db import get_db
from .deps import JWT_SECRET, ALGORITHM, get_current_user
from .schemas import CreateUserRequest, LoginRequest, LoginResponse, UserResponse

router = APIRouter(tags=["auth"])


@router.post("/login", response_model=LoginResponse)
async def login(request: LoginRequest):
    """Authenticate user and return JWT token."""
    async with get_db() as conn:
        user = await conn.fetchrow(
            "SELECT username, password_hash, role FROM users WHERE username = $1",
            request.username,
        )
        if not user or not bcrypt.verify(request.password, user["password_hash"]):
            raise HTTPException(status_code=401, detail="Invalid credentials")

        token = jwt.encode(
            {
                "sub": user["username"],
                "role": user["role"],
                "exp": datetime.utcnow() + timedelta(minutes=30),
            },
            JWT_SECRET,
            algorithm=ALGORITHM,
        )
        return LoginResponse(access_token=token)


@router.post("/admin/create-user", response_model=UserResponse)
async def create_user(
    request: CreateUserRequest,
    current_user: dict = Depends(get_current_user),
):
    """Create a new user (admin only)."""
    if current_user["role"] != "admin":
        raise HTTPException(status_code=403, detail="Admin access required")

    password_hash = bcrypt.hash(request.password)
    async with get_db() as conn:
        try:
            user = await conn.fetchrow(
                """INSERT INTO users (username, password_hash, role)
                   VALUES ($1, $2, $3)
                   RETURNING id, username, role""",
                request.username,
                password_hash,
                request.role,
            )
            return UserResponse(
                id=user["id"], username=user["username"], role=user["role"]
            )
        except Exception:
            raise HTTPException(status_code=400, detail="Username already exists")


@router.get("/admin/list-users", response_model=list[UserResponse])
async def list_users(current_user: dict = Depends(get_current_user)):
    """List all users (admin only)."""
    if current_user["role"] != "admin":
        raise HTTPException(status_code=403, detail="Admin access required")

    async with get_db() as conn:
        users = await conn.fetch("SELECT id, username, role FROM users ORDER BY id")
        return [
            UserResponse(id=u["id"], username=u["username"], role=u["role"])
            for u in users
        ]
