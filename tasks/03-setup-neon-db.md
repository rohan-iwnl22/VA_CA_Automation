# Task 03: Set Up Neon DB Connection

## Objective
Create the database connection layer for Neon DB (PostgreSQL) to store user authentication data.

## Files to Create
- `src/va_ca_automation/api/db.py`

## Files to Modify
- `pyproject.toml` — Add new dependencies

## Dependencies to Add

```
asyncpg>=0.29.0
passlib[bcrypt]>=1.7.4
python-jose[cryptography]>=3.3.0
```

## Database Schema

```sql
CREATE TABLE IF NOT EXISTS users (
    id SERIAL PRIMARY KEY,
    username VARCHAR(50) UNIQUE NOT NULL,
    password_hash VARCHAR(255) NOT NULL,
    role VARCHAR(10) NOT NULL DEFAULT 'user',
    created_at TIMESTAMP WITH TIME ZONE DEFAULT NOW()
);
```

## Implementation Steps

### Step 1: Add dependencies to `pyproject.toml`

Add to the `dependencies` list:
```toml
dependencies = [
  # ... existing deps ...
  "fastapi>=0.104.0",
  "uvicorn[standard]>=0.24.0",
  "asyncpg>=0.29.0",
  "passlib[bcrypt]>=1.7.4",
  "python-jose[cryptography]>=3.3.0",
  "python-multipart>=0.0.6",
]
```

### Step 2: Create `src/va_ca_automation/api/__init__.py`
Empty file.

### Step 3: Create `src/va_ca_automation/api/db.py`

```python
"""Neon DB connection pool and initialization."""

from __future__ import annotations

import os
from contextlib import asynccontextmanager
from typing import AsyncGenerator

import asyncpg

DATABASE_URL = os.getenv("DATABASE_URL")

_pool: asyncpg.Pool | None = None


async def get_pool() -> asyncpg.Pool:
    """Get or create the connection pool."""
    global _pool
    if _pool is None:
        _pool = await asyncpg.create_pool(DATABASE_URL, min_size=2, max_size=10)
    return _pool


async def init_db() -> None:
    """Create tables if they don't exist and seed admin user."""
    pool = await get_pool()
    async with pool.acquire() as conn:
        await conn.execute("""
            CREATE TABLE IF NOT EXISTS users (
                id SERIAL PRIMARY KEY,
                username VARCHAR(50) UNIQUE NOT NULL,
                password_hash VARCHAR(255) NOT NULL,
                role VARCHAR(10) NOT NULL DEFAULT 'user',
                created_at TIMESTAMP WITH TIME ZONE DEFAULT NOW()
            );
        """)
        # Seed admin if not exists
        from passlib.hash import bcrypt
        admin_password_hash = bcrypt.hash(os.getenv("ADMIN_PASSWORD", "admin123"))
        await conn.execute("""
            INSERT INTO users (username, password_hash, role)
            VALUES ($1, $2, 'admin')
            ON CONFLICT (username) DO NOTHING;
        """, os.getenv("ADMIN_USERNAME", "admin"), admin_password_hash)


async def close_db() -> None:
    """Close the connection pool."""
    global _pool
    if _pool:
        await _pool.close()
        _pool = None


@asynccontextmanager
async def get_db() -> AsyncGenerator[asyncpg.Connection, None]:
    """Get a database connection from the pool."""
    pool = await get_pool()
    async with pool.acquire() as conn:
        yield conn
```

### Step 4: Set environment variables

Add to `.env` or Render environment:
```
DATABASE_URL=postgresql://user:password@ep-xxx.us-east-2.aws.neon.tech/neondb?sslmode=require
ADMIN_USERNAME=admin
ADMIN_PASSWORD=<secure_password>
JWT_SECRET=<random_secret_key>
```

## Acceptance Criteria
- [ ] `db.py` created with connection pool
- [ ] `init_db()` creates users table and seeds admin
- [ ] `get_db()` context manager works
- [ ] Dependencies added to `pyproject.toml`
- [ ] Can connect to Neon DB with correct `DATABASE_URL`
