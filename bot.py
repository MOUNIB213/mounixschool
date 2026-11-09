#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
================================================================================
MOUNIX SCHOOL — Complete School & Educational Center Management Ecosystem
================================================================================
Single-File Unified Backend:
- SQLite Central Relational Database (24 Tables, Foreign Keys, Indexes, Migrations)
- FastAPI REST API (JWT Auth, Refresh Tokens, RBAC, Rate Limiting, OpenAPI)
- Aiogram 3 Telegram Bot (16-Menu Admin Dashboard, FSM Wizards, Role Menus)
- PDF & Excel RTL Arabic Reporting Engine (Receipts, Financial, Attendance, Grades)
- Automated SQLite Backup & Restore Engine with Integrity Verification
================================================================================
"""

from __future__ import annotations

import asyncio
import base64
import contextlib
import datetime
import hashlib
import hmac
import io
import json
import logging
import os
import re
import secrets
import shutil
import signal
import sqlite3
import sys
import time
import traceback
import uuid
from collections import defaultdict, deque
from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path
from typing import Any, Callable, Coroutine, Dict, Iterable, List, Optional, Sequence, Tuple, Union

import httpx
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import uvicorn
from fastapi import (
    Body,
    Depends,
    FastAPI,
    File,
    Form,
    Header,
    HTTPException,
    Path as FastAPIPath,
    Query,
    Request,
    Response,
    Security,
    UploadFile,
    status,
)
from fastapi.exceptions import RequestValidationError
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse, StreamingResponse
from fastapi.security import HTTPAuthorizationCredentials, HTTPBearer
from pydantic import BaseModel, Field

def model_to_dict(model_obj, exclude_none: bool = True) -> dict:
    if hasattr(model_obj, "model_dump"):
        return model_obj.model_dump(exclude_none=exclude_none)
    return model_obj.dict(exclude_none=exclude_none)

# ------------------------------------------------------------------------------
# Optional / Adaptive Imports (Supports both production pip env & restricted VMs)
# ------------------------------------------------------------------------------
try:
    import aiosqlite as _real_aiosqlite
    HAS_AIOSQLITE = True
except ImportError:
    _real_aiosqlite = None
    HAS_AIOSQLITE = False

try:
    import argon2
    _argon2_hasher = argon2.PasswordHasher()
    HAS_ARGON2 = True
except ImportError:
    _argon2_hasher = None
    HAS_ARGON2 = False

try:
    import bcrypt as _bcrypt
    HAS_BCRYPT = True
except ImportError:
    _bcrypt = None
    HAS_BCRYPT = False

try:
    import jwt as _pyjwt
    HAS_PYJWT = True
except ImportError:
    _pyjwt = None
    HAS_PYJWT = False

try:
    import weasyprint
    HAS_WEASYPRINT = True
except ImportError:
    weasyprint = None
    HAS_WEASYPRINT = False

try:
    from aiogram import Bot, Dispatcher, F, Router
    from aiogram.client.default import DefaultBotProperties
    from aiogram.enums import ParseMode
    from aiogram.filters import Command, CommandStart, StateFilter
    from aiogram.fsm.context import FSMContext
    from aiogram.fsm.state import State, StatesGroup
    from aiogram.fsm.storage.memory import MemoryStorage
    from aiogram.types import (
        BufferedInputFile,
        CallbackQuery,
        FSInputFile,
        InlineKeyboardButton,
        InlineKeyboardMarkup,
        Message,
        ReplyKeyboardMarkup,
        KeyboardButton,
    )
    HAS_AIOGRAM = True
except ImportError:
    HAS_AIOGRAM = False

    # Lightweight Aiogram-compatible shims so bot.py syntax, handlers, FSM states,
    # and callback routers can be imported, tested, or executed via HTTP polling fallback
    class State:
        def __init__(self, state: Optional[str] = None, group_name: Optional[str] = None):
            self._state = state
            self._group = group_name

        def __set_name__(self, owner: Any, name: str) -> None:
            self._group = owner.__name__
            self._state = f"{owner.__name__}:{name}"

        @property
        def state(self) -> str:
            return self._state or ""

        def __str__(self) -> str:
            return self.state

        def __eq__(self, other: object) -> bool:
            if isinstance(other, State):
                return self.state == other.state
            if isinstance(other, str):
                return self.state == other
            return False

        def __hash__(self) -> int:
            return hash(self.state)

    class StatesGroupMeta(type):
        def __new__(mcs, name: str, bases: tuple, namespace: dict):
            cls = super().__new__(mcs, name, bases, namespace)
            for k, v in namespace.items():
                if isinstance(v, State):
                    v.__set_name__(cls, k)
            return cls

    class StatesGroup(metaclass=StatesGroupMeta):
        pass

    class InlineKeyboardButton:
        def __init__(self, text: str, callback_data: Optional[str] = None, url: Optional[str] = None):
            self.text = text
            self.callback_data = callback_data
            self.url = url

        def to_dict(self) -> Dict[str, Any]:
            d: Dict[str, Any] = {"text": self.text}
            if self.callback_data is not None:
                d["callback_data"] = self.callback_data
            if self.url is not None:
                d["url"] = self.url
            return d

    class InlineKeyboardMarkup:
        def __init__(self, inline_keyboard: List[List[InlineKeyboardButton]]):
            self.inline_keyboard = inline_keyboard

        def to_dict(self) -> Dict[str, Any]:
            return {
                "inline_keyboard": [
                    [btn.to_dict() for btn in row] for row in self.inline_keyboard
                ]
            }

    class BufferedInputFile:
        def __init__(self, file: bytes, filename: str):
            self.data = file
            self.filename = filename

    class FSInputFile:
        def __init__(self, path: Union[str, Path], filename: Optional[str] = None):
            self.path = Path(path)
            self.filename = filename or self.path.name
            self.data = self.path.read_bytes() if self.path.exists() else b""

    class FSMContext:
        def __init__(self, storage_dict: Dict[int, Dict[str, Any]], user_id: int):
            self._storage = storage_dict
            self._user_id = user_id
            if user_id not in self._storage:
                self._storage[user_id] = {"state": None, "data": {}}

        async def set_state(self, state: Optional[Union[State, str]]) -> None:
            self._storage[self._user_id]["state"] = state.state if isinstance(state, State) else state

        async def get_state(self) -> Optional[str]:
            return self._storage[self._user_id].get("state")

        async def update_data(self, **kwargs: Any) -> Dict[str, Any]:
            self._storage[self._user_id]["data"].update(kwargs)
            return dict(self._storage[self._user_id]["data"])

        async def get_data(self) -> Dict[str, Any]:
            return dict(self._storage[self._user_id].get("data", {}))

        async def clear(self) -> None:
            self._storage[self._user_id] = {"state": None, "data": {}}


# ==============================================================================
# 1. CONFIGURATION & STRUCTURED LOGGING
# ==============================================================================

BASE_DIR = Path(os.getenv("MOUNIX_BASE_DIR", Path(__file__).resolve().parent))
DATA_DIR = Path(os.getenv("MOUNIX_DATA_DIR", BASE_DIR / "data"))
DB_PATH = Path(os.getenv("MOUNIX_DB_PATH", DATA_DIR / "mounix_school.db"))
BACKUP_DIR = Path(os.getenv("MOUNIX_BACKUP_DIR", DATA_DIR / "backups"))
UPLOAD_DIR = Path(os.getenv("MOUNIX_UPLOAD_DIR", DATA_DIR / "uploads"))
REPORTS_DIR = Path(os.getenv("MOUNIX_REPORTS_DIR", DATA_DIR / "reports"))
LOG_FILE = Path(os.getenv("MOUNIX_LOG_FILE", DATA_DIR / "mounix_school.log"))

for _d in (DATA_DIR, BACKUP_DIR, UPLOAD_DIR, REPORTS_DIR):
    _d.mkdir(parents=True, exist_ok=True)

BOT_TOKEN: str = os.getenv("BOT_TOKEN", "")
JWT_SECRET: str = os.getenv("JWT_SECRET", "mounix-school-super-secret-jwt-key-change-in-production-2026")
JWT_ALGORITHM: str = "HS256"
ACCESS_TOKEN_EXPIRE_MINUTES: int = int(os.getenv("ACCESS_TOKEN_EXPIRE_MINUTES", "60"))
REFRESH_TOKEN_EXPIRE_DAYS: int = int(os.getenv("REFRESH_TOKEN_EXPIRE_DAYS", "30"))
API_HOST: str = os.getenv("API_HOST", "0.0.0.0")
API_PORT: int = int(os.getenv("API_PORT", "8000"))
DEFAULT_ADMIN_USERNAME: str = os.getenv("DEFAULT_ADMIN_USERNAME", "admin")
DEFAULT_ADMIN_PASSWORD: str = os.getenv("DEFAULT_ADMIN_PASSWORD", "Admin@Mounix2026")
DEFAULT_ADMIN_TELEGRAM_ID: Optional[int] = (
    int(os.getenv("DEFAULT_ADMIN_TELEGRAM_ID"))
    if os.getenv("DEFAULT_ADMIN_TELEGRAM_ID", "").isdigit()
    else None
)
BACKUP_RETENTION_COUNT: int = int(os.getenv("BACKUP_RETENTION_COUNT", "10"))
MAX_UPLOAD_SIZE_BYTES: int = int(os.getenv("MAX_UPLOAD_SIZE_BYTES", str(25 * 1024 * 1024)))  # 25 MB

ALLOWED_UPLOAD_EXTENSIONS = {
    ".pdf", ".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx",
    ".txt", ".csv", ".png", ".jpg", ".jpeg", ".webp", ".zip"
}
FORBIDDEN_EXTENSIONS = {
    ".exe", ".bat", ".cmd", ".sh", ".ps1", ".vbs", ".js", ".jar",
    ".msi", ".scr", ".com", ".elf", ".bin", ".php", ".py", ".pl", ".rb", ".dll", ".so"
}


class SecretMaskingFilter(logging.Filter):
    """Prevents BOT_TOKEN, JWT_SECRET, or passwords from leaking into logs."""

    def filter(self, record: logging.LogRecord) -> bool:
        msg = record.getMessage()
        if BOT_TOKEN and len(BOT_TOKEN) > 8 and BOT_TOKEN in msg:
            msg = msg.replace(BOT_TOKEN, "[REDACTED_BOT_TOKEN]")
        if JWT_SECRET and len(JWT_SECRET) > 8 and JWT_SECRET in msg:
            msg = msg.replace(JWT_SECRET, "[REDACTED_JWT_SECRET]")
        record.msg = msg
        record.args = ()
        return True


logger = logging.getLogger("mounix_school")
logger.setLevel(logging.INFO)
if not logger.handlers:
    _fmt = logging.Formatter("%(asctime)s | %(levelname)-8s | %(name)s | %(message)s")
    _sh = logging.StreamHandler(sys.stdout)
    _sh.setFormatter(_fmt)
    _sh.addFilter(SecretMaskingFilter())
    _fh = logging.FileHandler(LOG_FILE, encoding="utf-8")
    _fh.setFormatter(_fmt)
    _fh.addFilter(SecretMaskingFilter())
    logger.addHandler(_sh)
    logger.addHandler(_fh)


# ==============================================================================
# 2. ENUMS & DOMAIN CONSTANTS
# ==============================================================================

class UserRole(str, Enum):
    ADMIN = "ADMIN"
    TEACHER = "TEACHER"
    STUDENT = "STUDENT"
    PARENT = "PARENT"


class AttendanceStatus(str, Enum):
    PRESENT = "PRESENT"
    ABSENT = "ABSENT"
    LATE = "LATE"
    EXCUSED = "EXCUSED"


class PaymentPlanType(str, Enum):
    MONTHLY = "MONTHLY"
    PER_SESSION = "PER_SESSION"
    HYBRID = "HYBRID"


class MaterialType(str, Enum):
    PDF = "PDF"
    LINK = "LINK"
    LESSON = "LESSON"
    ASSIGNMENT = "ASSIGNMENT"


class NotificationType(str, Enum):
    ABSENCE = "ABSENCE"
    SCHEDULE_CHANGE = "SCHEDULE_CHANGE"
    NEW_ASSIGNMENT = "NEW_ASSIGNMENT"
    NEW_MATERIAL = "NEW_MATERIAL"
    ANNOUNCEMENT = "ANNOUNCEMENT"
    NEW_GRADE = "NEW_GRADE"
    PAYMENT_DUE = "PAYMENT_DUE"
    PAYMENT_RECEIVED = "PAYMENT_RECEIVED"
    CLASS_REMINDER = "CLASS_REMINDER"


# ==============================================================================
# 3. SECURITY: PASSWORD HASHING & JWT ENGINE
# ==============================================================================

def hash_password(plain_password: str) -> str:
    """
    Hashes password using Argon2id if available, otherwise Bcrypt, otherwise
    PBKDF2-HMAC-SHA256 (600,000 OWASP-recommended iterations) with constant-time verification.
    """
    if HAS_ARGON2 and _argon2_hasher is not None:
        return _argon2_hasher.hash(plain_password)
    if HAS_BCRYPT and _bcrypt is not None:
        salt = _bcrypt.gensalt(rounds=12)
        return _bcrypt.hashpw(plain_password.encode("utf-8"), salt).decode("utf-8")
    salt = secrets.token_hex(16)
    iterations = 600_000
    dk = hashlib.pbkdf2_hmac("sha256", plain_password.encode("utf-8"), salt.encode("utf-8"), iterations)
    return f"pbkdf2_sha256${iterations}${salt}${dk.hex()}"


def verify_password(plain_password: str, password_hash: str) -> bool:
    """Verifies password across Argon2id, Bcrypt, or PBKDF2-HMAC-SHA256."""
    if not password_hash:
        return False
    try:
        if password_hash.startswith("$argon2") and HAS_ARGON2 and _argon2_hasher is not None:
            return _argon2_hasher.verify(password_hash, plain_password)
        if password_hash.startswith("$2") and HAS_BCRYPT and _bcrypt is not None:
            return _bcrypt.checkpw(plain_password.encode("utf-8"), password_hash.encode("utf-8"))
        if password_hash.startswith("pbkdf2_sha256$"):
            parts = password_hash.split("$")
            if len(parts) != 4:
                return False
            _, iter_str, salt, stored_hex = parts
            dk = hashlib.pbkdf2_hmac(
                "sha256",
                plain_password.encode("utf-8"),
                salt.encode("utf-8"),
                int(iter_str),
            )
            return hmac.compare_digest(dk.hex(), stored_hex)
    except Exception as exc:
        logger.warning(f"Password verification error: {exc}")
        return False
    return False


def _b64url_encode(data: bytes) -> str:
    return base64.urlsafe_b64encode(data).rstrip(b"=").decode("ascii")


def _b64url_decode(data: str) -> bytes:
    padding = "=" * (-len(data) % 4)
    return base64.urlsafe_b64decode(data + padding)


def encode_jwt(payload: Dict[str, Any], secret: str = JWT_SECRET) -> str:
    """RFC 7519 compliant HS256 JWT encoder."""
    if HAS_PYJWT and _pyjwt is not None:
        token = _pyjwt.encode(payload, secret, algorithm=JWT_ALGORITHM)
        return token if isinstance(token, str) else token.decode("utf-8")
    header = {"alg": JWT_ALGORITHM, "typ": "JWT"}
    h_b64 = _b64url_encode(json.dumps(header, separators=(",", ":")).encode("utf-8"))
    p_b64 = _b64url_encode(json.dumps(payload, separators=(",", ":")).encode("utf-8"))
    signing_input = f"{h_b64}.{p_b64}".encode("ascii")
    sig = hmac.new(secret.encode("utf-8"), signing_input, hashlib.sha256).digest()
    return f"{h_b64}.{p_b64}.{_b64url_encode(sig)}"


def decode_jwt(token: str, secret: str = JWT_SECRET) -> Dict[str, Any]:
    """RFC 7519 compliant HS256 JWT decoder with expiration validation."""
    if HAS_PYJWT and _pyjwt is not None:
        try:
            return _pyjwt.decode(token, secret, algorithms=[JWT_ALGORITHM])
        except Exception as exc:
            raise ValueError(f"Invalid or expired token: {exc}") from exc

    parts = token.split(".")
    if len(parts) != 3:
        raise ValueError("Malformed JWT token")
    h_b64, p_b64, s_b64 = parts
    signing_input = f"{h_b64}.{p_b64}".encode("ascii")
    expected_sig = hmac.new(secret.encode("utf-8"), signing_input, hashlib.sha256).digest()
    actual_sig = _b64url_decode(s_b64)
    if not hmac.compare_digest(expected_sig, actual_sig):
        raise ValueError("Invalid JWT signature")
    payload = json.loads(_b64url_decode(p_b64).decode("utf-8"))
    exp = payload.get("exp")
    if exp is not None and int(time.time()) >= int(exp):
        raise ValueError("Token has expired")
    return payload


def create_access_token(user_id: int, username: str, role: str) -> str:
    now = int(time.time())
    payload = {
        "sub": str(user_id),
        "username": username,
        "role": role,
        "type": "access",
        "jti": secrets.token_hex(12),
        "iat": now,
        "exp": now + (ACCESS_TOKEN_EXPIRE_MINUTES * 60),
    }
    return encode_jwt(payload)


def create_refresh_token(user_id: int, username: str, role: str) -> str:
    now = int(time.time())
    payload = {
        "sub": str(user_id),
        "username": username,
        "role": role,
        "type": "refresh",
        "jti": secrets.token_hex(16),
        "iat": now,
        "exp": now + (REFRESH_TOKEN_EXPIRE_DAYS * 86400),
    }
    return encode_jwt(payload)


# Revoked JTI store (in-memory + persisted in app_settings / audit_logs)
REVOKED_TOKEN_JTIS: set[str] = set()


# ==============================================================================
# 4. ASYNC SQLITE DATABASE ENGINE & 24-TABLE SCHEMA + MIGRATIONS
# ==============================================================================

class AsyncDBTransaction:
    """Provides safe async transaction context with automatic commit/rollback."""

    def __init__(self, db_path: Path, lock: asyncio.Lock):
        self.db_path = db_path
        self.lock = lock
        self.conn: Optional[sqlite3.Connection] = None

    async def __aenter__(self) -> "AsyncDBSession":
        await self.lock.acquire()
        self.conn = await asyncio.to_thread(self._open_conn)
        await asyncio.to_thread(self.conn.execute, "BEGIN IMMEDIATE")
        return AsyncDBSession(self.conn)

    def _open_conn(self) -> sqlite3.Connection:
        conn = sqlite3.connect(str(self.db_path), timeout=30.0, check_same_thread=False)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA foreign_keys = ON")
        try:
            conn.execute("PRAGMA journal_mode = WAL")
        except sqlite3.OperationalError:
            conn.execute("PRAGMA journal_mode = MEMORY")
        conn.execute("PRAGMA synchronous = NORMAL")
        conn.execute("PRAGMA busy_timeout = 10000")
        return conn

    async def __aexit__(self, exc_type: Any, exc_val: Any, exc_tb: Any) -> None:
        try:
            if self.conn is not None:
                if exc_type is None:
                    await asyncio.to_thread(self.conn.commit)
                else:
                    await asyncio.to_thread(self.conn.rollback)
                await asyncio.to_thread(self.conn.close)
        finally:
            self.lock.release()


class AsyncDBSession:
    """Async wrapper over sqlite3 connection returning dictionary rows."""

    def __init__(self, conn: sqlite3.Connection):
        self.conn = conn

    async def execute(self, sql: str, params: Sequence[Any] = ()) -> sqlite3.Cursor:
        return await asyncio.to_thread(self.conn.execute, sql, tuple(params))

    async def executemany(self, sql: str, seq_of_params: Iterable[Sequence[Any]]) -> sqlite3.Cursor:
        return await asyncio.to_thread(self.conn.executemany, sql, list(seq_of_params))

    async def fetchone(self, sql: str, params: Sequence[Any] = ()) -> Optional[Dict[str, Any]]:
        def _run() -> Optional[Dict[str, Any]]:
            cur = self.conn.execute(sql, tuple(params))
            row = cur.fetchone()
            return dict(row) if row is not None else None
        return await asyncio.to_thread(_run)

    async def fetchall(self, sql: str, params: Sequence[Any] = ()) -> List[Dict[str, Any]]:
        def _run() -> List[Dict[str, Any]]:
            cur = self.conn.execute(sql, tuple(params))
            return [dict(r) for r in cur.fetchall()]
        return await asyncio.to_thread(_run)

    async def insert(self, sql: str, params: Sequence[Any] = ()) -> int:
        def _run() -> int:
            cur = self.conn.execute(sql, tuple(params))
            return int(cur.lastrowid or 0)
        return await asyncio.to_thread(_run)


class DatabaseManager:
    """Central SQLite Database Manager supporting read queries and serialized write transactions."""

    def __init__(self, db_path: Path = DB_PATH):
        self.db_path = db_path
        self._write_lock = asyncio.Lock()

    def transaction(self) -> AsyncDBTransaction:
        return AsyncDBTransaction(self.db_path, self._write_lock)

    async def fetchone(self, sql: str, params: Sequence[Any] = ()) -> Optional[Dict[str, Any]]:
        def _run() -> Optional[Dict[str, Any]]:
            conn = sqlite3.connect(str(self.db_path), timeout=30.0, check_same_thread=False)
            conn.row_factory = sqlite3.Row
            conn.execute("PRAGMA foreign_keys = ON")
            try:
                cur = conn.execute(sql, tuple(params))
                row = cur.fetchone()
                return dict(row) if row is not None else None
            finally:
                conn.close()
        return await asyncio.to_thread(_run)

    async def fetchall(self, sql: str, params: Sequence[Any] = ()) -> List[Dict[str, Any]]:
        def _run() -> List[Dict[str, Any]]:
            conn = sqlite3.connect(str(self.db_path), timeout=30.0, check_same_thread=False)
            conn.row_factory = sqlite3.Row
            conn.execute("PRAGMA foreign_keys = ON")
            try:
                cur = conn.execute(sql, tuple(params))
                return [dict(r) for r in cur.fetchall()]
            finally:
                conn.close()
        return await asyncio.to_thread(_run)


db = DatabaseManager(DB_PATH)


SCHEMA_SQL = """
-- 1. users
CREATE TABLE IF NOT EXISTS users (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    username TEXT NOT NULL UNIQUE,
    password_hash TEXT NOT NULL,
    full_name TEXT NOT NULL,
    role TEXT NOT NULL CHECK (role IN ('ADMIN', 'TEACHER', 'STUDENT', 'PARENT')),
    phone TEXT,
    email TEXT,
    telegram_id INTEGER UNIQUE,
    language TEXT NOT NULL DEFAULT 'ar' CHECK (language IN ('ar', 'en')),
    theme_mode TEXT NOT NULL DEFAULT 'system' CHECK (theme_mode IN ('light', 'dark', 'system')),
    is_active INTEGER NOT NULL DEFAULT 1 CHECK (is_active IN (0, 1)),
    is_deleted INTEGER NOT NULL DEFAULT 0 CHECK (is_deleted IN (0, 1)),
    last_login_at TEXT,
    created_at TEXT NOT NULL DEFAULT (datetime('now')),
    updated_at TEXT NOT NULL DEFAULT (datetime('now'))
);

-- 2. levels
CREATE TABLE IF NOT EXISTS levels (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    code TEXT NOT NULL UNIQUE,
    name_ar TEXT NOT NULL,
    name_en TEXT NOT NULL,
    sort_order INTEGER NOT NULL DEFAULT 1,
    description TEXT,
    is_archived INTEGER NOT NULL DEFAULT 0 CHECK (is_archived IN (0, 1)),
    created_at TEXT NOT NULL DEFAULT (datetime('now')),
    updated_at TEXT NOT NULL DEFAULT (datetime('now'))
);

-- 3. parents
CREATE TABLE IF NOT EXISTS parents (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id INTEGER UNIQUE REFERENCES users(id) ON DELETE SET NULL,
    parent_code TEXT NOT NULL UNIQUE,
    full_name TEXT NOT NULL,
    phone TEXT NOT NULL,
    secondary_phone TEXT,
    email TEXT,
    address TEXT,
    occupation TEXT,
    notes TEXT,
    is_archived INTEGER NOT NULL DEFAULT 0 CHECK (is_archived IN (0, 1)),
    created_at TEXT NOT NULL DEFAULT (datetime('now')),
    updated_at TEXT NOT NULL DEFAULT (datetime('now'))
);

-- 4. students
CREATE TABLE IF NOT EXISTS students (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id INTEGER UNIQUE REFERENCES users(id) ON DELETE SET NULL,
    student_code TEXT NOT NULL UNIQUE,
    full_name TEXT NOT NULL,
    birth_date TEXT,
    gender TEXT DEFAULT 'M' CHECK (gender IN ('M', 'F')),
    level_id INTEGER REFERENCES levels(id) ON DELETE SET NULL,
    primary_parent_id INTEGER REFERENCES parents(id) ON DELETE SET NULL,
    phone TEXT,
    address TEXT,
    joined_date TEXT NOT NULL DEFAULT (date('now')),
    admin_notes TEXT,
    is_archived INTEGER NOT NULL DEFAULT 0 CHECK (is_archived IN (0, 1)),
    created_at TEXT NOT NULL DEFAULT (datetime('now')),
    updated_at TEXT NOT NULL DEFAULT (datetime('now'))
);

-- 5. parent_students
CREATE TABLE IF NOT EXISTS parent_students (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    parent_id INTEGER NOT NULL REFERENCES parents(id) ON DELETE CASCADE,
    student_id INTEGER NOT NULL REFERENCES students(id) ON DELETE CASCADE,
    relationship TEXT NOT NULL DEFAULT 'Guardian',
    is_primary INTEGER NOT NULL DEFAULT 1 CHECK (is_primary IN (0, 1)),
    created_at TEXT NOT NULL DEFAULT (datetime('now')),
    UNIQUE(parent_id, student_id)
);

-- 6. teachers
CREATE TABLE IF NOT EXISTS teachers (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id INTEGER UNIQUE REFERENCES users(id) ON DELETE SET NULL,
    teacher_code TEXT NOT NULL UNIQUE,
    full_name TEXT NOT NULL,
    phone TEXT NOT NULL,
    email TEXT,
    specialization TEXT NOT NULL,
    bio TEXT,
    hire_date TEXT NOT NULL DEFAULT (date('now')),
    hourly_rate REAL NOT NULL DEFAULT 0.0 CHECK (hourly_rate >= 0),
    is_archived INTEGER NOT NULL DEFAULT 0 CHECK (is_archived IN (0, 1)),
    created_at TEXT NOT NULL DEFAULT (datetime('now')),
    updated_at TEXT NOT NULL DEFAULT (datetime('now'))
);

-- 7. subjects
CREATE TABLE IF NOT EXISTS subjects (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    code TEXT NOT NULL UNIQUE,
    name_ar TEXT NOT NULL,
    name_en TEXT NOT NULL,
    description TEXT,
    level_id INTEGER REFERENCES levels(id) ON DELETE SET NULL,
    default_monthly_fee REAL NOT NULL DEFAULT 0.0 CHECK (default_monthly_fee >= 0),
    default_session_fee REAL NOT NULL DEFAULT 0.0 CHECK (default_session_fee >= 0),
    is_archived INTEGER NOT NULL DEFAULT 0 CHECK (is_archived IN (0, 1)),
    created_at TEXT NOT NULL DEFAULT (datetime('now')),
    updated_at TEXT NOT NULL DEFAULT (datetime('now'))
);

-- 8. teacher_subjects
CREATE TABLE IF NOT EXISTS teacher_subjects (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    teacher_id INTEGER NOT NULL REFERENCES teachers(id) ON DELETE CASCADE,
    subject_id INTEGER NOT NULL REFERENCES subjects(id) ON DELETE CASCADE,
    created_at TEXT NOT NULL DEFAULT (datetime('now')),
    UNIQUE(teacher_id, subject_id)
);

-- 9. classes
CREATE TABLE IF NOT EXISTS classes (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    code TEXT NOT NULL UNIQUE,
    name TEXT NOT NULL,
    level_id INTEGER NOT NULL REFERENCES levels(id) ON DELETE RESTRICT,
    subject_id INTEGER NOT NULL REFERENCES subjects(id) ON DELETE RESTRICT,
    teacher_id INTEGER NOT NULL REFERENCES teachers(id) ON DELETE RESTRICT,
    room TEXT NOT NULL,
    capacity INTEGER NOT NULL DEFAULT 25 CHECK (capacity > 0),
    billing_mode TEXT NOT NULL DEFAULT 'MONTHLY' CHECK (billing_mode IN ('MONTHLY', 'PER_SESSION', 'HYBRID')),
    monthly_fee REAL NOT NULL DEFAULT 0.0 CHECK (monthly_fee >= 0),
    session_fee REAL NOT NULL DEFAULT 0.0 CHECK (session_fee >= 0),
    is_archived INTEGER NOT NULL DEFAULT 0 CHECK (is_archived IN (0, 1)),
    created_at TEXT NOT NULL DEFAULT (datetime('now')),
    updated_at TEXT NOT NULL DEFAULT (datetime('now'))
);

-- 10. enrollments
CREATE TABLE IF NOT EXISTS enrollments (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    student_id INTEGER NOT NULL REFERENCES students(id) ON DELETE CASCADE,
    class_id INTEGER NOT NULL REFERENCES classes(id) ON DELETE CASCADE,
    enrolled_at TEXT NOT NULL DEFAULT (date('now')),
    status TEXT NOT NULL DEFAULT 'ACTIVE' CHECK (status IN ('ACTIVE', 'DROPPED', 'COMPLETED')),
    billing_override TEXT CHECK (billing_override IS NULL OR billing_override IN ('MONTHLY', 'PER_SESSION')),
    discount_percent REAL NOT NULL DEFAULT 0.0 CHECK (discount_percent >= 0 AND discount_percent <= 100),
    override_capacity_by INTEGER REFERENCES users(id) ON DELETE SET NULL,
    created_at TEXT NOT NULL DEFAULT (datetime('now')),
    updated_at TEXT NOT NULL DEFAULT (datetime('now')),
    UNIQUE(student_id, class_id)
);

-- 11. schedules
CREATE TABLE IF NOT EXISTS schedules (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    class_id INTEGER NOT NULL REFERENCES classes(id) ON DELETE CASCADE,
    subject_id INTEGER NOT NULL REFERENCES subjects(id) ON DELETE RESTRICT,
    teacher_id INTEGER NOT NULL REFERENCES teachers(id) ON DELETE RESTRICT,
    day_of_week INTEGER NOT NULL CHECK (day_of_week BETWEEN 0 AND 6), -- 0=Sunday .. 6=Saturday
    specific_date TEXT, -- Optional YYYY-MM-DD for one-off sessions
    start_time TEXT NOT NULL, -- HH:MM (24h)
    end_time TEXT NOT NULL,   -- HH:MM (24h)
    duration_minutes INTEGER NOT NULL CHECK (duration_minutes > 0),
    room TEXT NOT NULL,
    status TEXT NOT NULL DEFAULT 'SCHEDULED' CHECK (status IN ('SCHEDULED', 'CANCELLED', 'COMPLETED')),
    notes TEXT,
    created_at TEXT NOT NULL DEFAULT (datetime('now')),
    updated_at TEXT NOT NULL DEFAULT (datetime('now'))
);

-- 12. attendance
CREATE TABLE IF NOT EXISTS attendance (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    schedule_id INTEGER NOT NULL REFERENCES schedules(id) ON DELETE CASCADE,
    class_id INTEGER NOT NULL REFERENCES classes(id) ON DELETE CASCADE,
    student_id INTEGER NOT NULL REFERENCES students(id) ON DELETE CASCADE,
    session_date TEXT NOT NULL, -- YYYY-MM-DD
    status TEXT NOT NULL CHECK (status IN ('PRESENT', 'ABSENT', 'LATE', 'EXCUSED')),
    late_minutes INTEGER NOT NULL DEFAULT 0 CHECK (late_minutes >= 0),
    absence_reason TEXT,
    recorded_by INTEGER REFERENCES users(id) ON DELETE SET NULL,
    created_at TEXT NOT NULL DEFAULT (datetime('now')),
    updated_at TEXT NOT NULL DEFAULT (datetime('now')),
    UNIQUE(schedule_id, student_id, session_date)
);

-- 13. exams
CREATE TABLE IF NOT EXISTS exams (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    title TEXT NOT NULL,
    subject_id INTEGER NOT NULL REFERENCES subjects(id) ON DELETE CASCADE,
    class_id INTEGER NOT NULL REFERENCES classes(id) ON DELETE CASCADE,
    teacher_id INTEGER NOT NULL REFERENCES teachers(id) ON DELETE RESTRICT,
    exam_date TEXT NOT NULL,
    max_score REAL NOT NULL DEFAULT 20.0 CHECK (max_score > 0),
    weight REAL NOT NULL DEFAULT 1.0 CHECK (weight > 0),
    show_ranking INTEGER NOT NULL DEFAULT 1 CHECK (show_ranking IN (0, 1)),
    description TEXT,
    is_published INTEGER NOT NULL DEFAULT 1 CHECK (is_published IN (0, 1)),
    created_at TEXT NOT NULL DEFAULT (datetime('now')),
    updated_at TEXT NOT NULL DEFAULT (datetime('now'))
);

-- 14. grades
CREATE TABLE IF NOT EXISTS grades (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    exam_id INTEGER NOT NULL REFERENCES exams(id) ON DELETE CASCADE,
    student_id INTEGER NOT NULL REFERENCES students(id) ON DELETE CASCADE,
    score REAL NOT NULL CHECK (score >= 0),
    teacher_remarks TEXT,
    recorded_by INTEGER REFERENCES users(id) ON DELETE SET NULL,
    created_at TEXT NOT NULL DEFAULT (datetime('now')),
    updated_at TEXT NOT NULL DEFAULT (datetime('now')),
    UNIQUE(exam_id, student_id)
);

-- 15. payment_plans
CREATE TABLE IF NOT EXISTS payment_plans (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    plan_code TEXT NOT NULL UNIQUE,
    student_id INTEGER NOT NULL REFERENCES students(id) ON DELETE RESTRICT,
    class_id INTEGER REFERENCES classes(id) ON DELETE SET NULL,
    subject_id INTEGER REFERENCES subjects(id) ON DELETE SET NULL,
    plan_type TEXT NOT NULL CHECK (plan_type IN ('MONTHLY', 'PER_SESSION')),
    sessions_count INTEGER DEFAULT 4 CHECK (sessions_count IS NULL OR sessions_count > 0),
    total_amount REAL NOT NULL CHECK (total_amount >= 0),
    paid_amount REAL NOT NULL DEFAULT 0.0 CHECK (paid_amount >= 0),
    remaining_amount REAL NOT NULL CHECK (remaining_amount >= 0),
    due_date TEXT NOT NULL,
    period_label TEXT, -- e.g. '2026-11' or 'Pack 8 Sessions'
    status TEXT NOT NULL DEFAULT 'UNPAID' CHECK (status IN ('UNPAID', 'PARTIAL', 'PAID', 'OVERDUE', 'VOIDED')),
    notes TEXT,
    created_by INTEGER REFERENCES users(id) ON DELETE SET NULL,
    created_at TEXT NOT NULL DEFAULT (datetime('now')),
    updated_at TEXT NOT NULL DEFAULT (datetime('now'))
);

-- 16. payments
CREATE TABLE IF NOT EXISTS payments (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    receipt_number TEXT NOT NULL UNIQUE,
    student_id INTEGER NOT NULL REFERENCES students(id) ON DELETE RESTRICT,
    amount REAL NOT NULL CHECK (amount > 0),
    payment_method TEXT NOT NULL DEFAULT 'CASH' CHECK (payment_method IN ('CASH', 'BANK_TRANSFER', 'CARD', 'CHECK', 'OTHER')),
    payment_date TEXT NOT NULL DEFAULT (datetime('now')),
    description TEXT,
    recorded_by INTEGER NOT NULL REFERENCES users(id) ON DELETE RESTRICT,
    recorded_by_name TEXT NOT NULL,
    is_voided INTEGER NOT NULL DEFAULT 0 CHECK (is_voided IN (0, 1)),
    void_reason TEXT,
    voided_by INTEGER REFERENCES users(id) ON DELETE SET NULL,
    created_at TEXT NOT NULL DEFAULT (datetime('now')),
    updated_at TEXT NOT NULL DEFAULT (datetime('now'))
);

-- 17. payment_allocations
CREATE TABLE IF NOT EXISTS payment_allocations (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    payment_id INTEGER NOT NULL REFERENCES payments(id) ON DELETE RESTRICT,
    payment_plan_id INTEGER NOT NULL REFERENCES payment_plans(id) ON DELETE RESTRICT,
    allocated_amount REAL NOT NULL CHECK (allocated_amount > 0),
    created_at TEXT NOT NULL DEFAULT (datetime('now'))
);

-- 18. materials
CREATE TABLE IF NOT EXISTS materials (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    title TEXT NOT NULL,
    description TEXT,
    material_type TEXT NOT NULL DEFAULT 'PDF' CHECK (material_type IN ('PDF', 'LINK', 'LESSON')),
    subject_id INTEGER NOT NULL REFERENCES subjects(id) ON DELETE CASCADE,
    class_id INTEGER REFERENCES classes(id) ON DELETE CASCADE,
    teacher_id INTEGER NOT NULL REFERENCES teachers(id) ON DELETE CASCADE,
    file_path TEXT,
    external_url TEXT,
    publish_date TEXT NOT NULL DEFAULT (date('now')),
    is_archived INTEGER NOT NULL DEFAULT 0 CHECK (is_archived IN (0, 1)),
    created_at TEXT NOT NULL DEFAULT (datetime('now')),
    updated_at TEXT NOT NULL DEFAULT (datetime('now'))
);

-- 19. assignments
CREATE TABLE IF NOT EXISTS assignments (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    title TEXT NOT NULL,
    description TEXT,
    subject_id INTEGER NOT NULL REFERENCES subjects(id) ON DELETE CASCADE,
    class_id INTEGER NOT NULL REFERENCES classes(id) ON DELETE CASCADE,
    teacher_id INTEGER NOT NULL REFERENCES teachers(id) ON DELETE CASCADE,
    file_path TEXT,
    external_url TEXT,
    publish_date TEXT NOT NULL DEFAULT (date('now')),
    due_date TEXT NOT NULL,
    max_score REAL DEFAULT 20.0,
    is_archived INTEGER NOT NULL DEFAULT 0 CHECK (is_archived IN (0, 1)),
    created_at TEXT NOT NULL DEFAULT (datetime('now')),
    updated_at TEXT NOT NULL DEFAULT (datetime('now'))
);

-- 20. notifications
CREATE TABLE IF NOT EXISTS notifications (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    recipient_user_id INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
    sender_user_id INTEGER REFERENCES users(id) ON DELETE SET NULL,
    notification_type TEXT NOT NULL CHECK (
        notification_type IN (
            'ABSENCE', 'SCHEDULE_CHANGE', 'NEW_ASSIGNMENT', 'NEW_MATERIAL',
            'ANNOUNCEMENT', 'NEW_GRADE', 'PAYMENT_DUE', 'PAYMENT_RECEIVED', 'CLASS_REMINDER'
        )
    ),
    title_ar TEXT NOT NULL,
    title_en TEXT NOT NULL,
    body_ar TEXT NOT NULL,
    body_en TEXT NOT NULL,
    dedup_key TEXT UNIQUE,
    is_read INTEGER NOT NULL DEFAULT 0 CHECK (is_read IN (0, 1)),
    read_at TEXT,
    telegram_sent INTEGER NOT NULL DEFAULT 0 CHECK (telegram_sent IN (0, 1)),
    related_entity_type TEXT,
    related_entity_id INTEGER,
    created_at TEXT NOT NULL DEFAULT (datetime('now'))
);

-- 21. notification_preferences
CREATE TABLE IF NOT EXISTS notification_preferences (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id INTEGER NOT NULL UNIQUE REFERENCES users(id) ON DELETE CASCADE,
    enable_telegram INTEGER NOT NULL DEFAULT 1 CHECK (enable_telegram IN (0, 1)),
    enable_app INTEGER NOT NULL DEFAULT 1 CHECK (enable_app IN (0, 1)),
    notify_absence INTEGER NOT NULL DEFAULT 1 CHECK (notify_absence IN (0, 1)),
    notify_grades INTEGER NOT NULL DEFAULT 1 CHECK (notify_grades IN (0, 1)),
    notify_payments INTEGER NOT NULL DEFAULT 1 CHECK (notify_payments IN (0, 1)),
    notify_schedule INTEGER NOT NULL DEFAULT 1 CHECK (notify_schedule IN (0, 1)),
    notify_materials INTEGER NOT NULL DEFAULT 1 CHECK (notify_materials IN (0, 1)),
    notify_announcements INTEGER NOT NULL DEFAULT 1 CHECK (notify_announcements IN (0, 1)),
    updated_at TEXT NOT NULL DEFAULT (datetime('now'))
);

-- 22. app_settings
CREATE TABLE IF NOT EXISTS app_settings (
    key TEXT PRIMARY KEY,
    value TEXT NOT NULL,
    description TEXT,
    updated_by INTEGER REFERENCES users(id) ON DELETE SET NULL,
    updated_at TEXT NOT NULL DEFAULT (datetime('now'))
);

-- 23. audit_logs
CREATE TABLE IF NOT EXISTS audit_logs (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    actor_user_id INTEGER REFERENCES users(id) ON DELETE SET NULL,
    actor_username TEXT,
    actor_role TEXT,
    action TEXT NOT NULL,
    entity_type TEXT NOT NULL,
    entity_id TEXT,
    details_json TEXT,
    ip_address TEXT,
    created_at TEXT NOT NULL DEFAULT (datetime('now'))
);

-- 24. backup_logs
CREATE TABLE IF NOT EXISTS backup_logs (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    filename TEXT NOT NULL UNIQUE,
    file_path TEXT NOT NULL,
    size_bytes INTEGER NOT NULL CHECK (size_bytes >= 0),
    checksum_sha256 TEXT NOT NULL,
    integrity_status TEXT NOT NULL DEFAULT 'OK',
    created_by INTEGER REFERENCES users(id) ON DELETE SET NULL,
    restored_at TEXT,
    created_at TEXT NOT NULL DEFAULT (datetime('now'))
);

-- Indexes for performance
CREATE INDEX IF NOT EXISTS idx_users_role ON users(role);
CREATE INDEX IF NOT EXISTS idx_users_telegram_id ON users(telegram_id);
CREATE INDEX IF NOT EXISTS idx_students_code ON students(student_code);
CREATE INDEX IF NOT EXISTS idx_students_name ON students(full_name);
CREATE INDEX IF NOT EXISTS idx_students_level ON students(level_id);
CREATE INDEX IF NOT EXISTS idx_students_parent ON students(primary_parent_id);
CREATE INDEX IF NOT EXISTS idx_teachers_code ON teachers(teacher_code);
CREATE INDEX IF NOT EXISTS idx_classes_level_subj ON classes(level_id, subject_id);
CREATE INDEX IF NOT EXISTS idx_classes_teacher ON classes(teacher_id);
CREATE INDEX IF NOT EXISTS idx_enrollments_student ON enrollments(student_id);
CREATE INDEX IF NOT EXISTS idx_enrollments_class ON enrollments(class_id);
CREATE INDEX IF NOT EXISTS idx_schedules_day_time ON schedules(day_of_week, start_time, end_time);
CREATE INDEX IF NOT EXISTS idx_schedules_teacher ON schedules(teacher_id);
CREATE INDEX IF NOT EXISTS idx_schedules_class ON schedules(class_id);
CREATE INDEX IF NOT EXISTS idx_attendance_student_date ON attendance(student_id, session_date);
CREATE INDEX IF NOT EXISTS idx_attendance_class_date ON attendance(class_id, session_date);
CREATE INDEX IF NOT EXISTS idx_grades_exam ON grades(exam_id);
CREATE INDEX IF NOT EXISTS idx_grades_student ON grades(student_id);
CREATE INDEX IF NOT EXISTS idx_payment_plans_student ON payment_plans(student_id, status);
CREATE INDEX IF NOT EXISTS idx_payments_student_date ON payments(student_id, payment_date);
CREATE INDEX IF NOT EXISTS idx_notifications_recipient ON notifications(recipient_user_id, is_read);
CREATE INDEX IF NOT EXISTS idx_audit_logs_actor ON audit_logs(actor_user_id, created_at);
"""

REQUIRED_TABLES_24 = [
    "users", "students", "parents", "parent_students", "teachers", "teacher_subjects",
    "subjects", "levels", "classes", "enrollments", "schedules", "attendance",
    "exams", "grades", "payment_plans", "payments", "payment_allocations",
    "materials", "assignments", "notifications", "notification_preferences",
    "app_settings", "audit_logs", "backup_logs",
]


async def run_migrations_and_initialize_db() -> None:
    """
    Automatically creates the 24 SQLite tables, indexes, and constraints, runs
    non-destructive schema migrations, and bootstraps default settings and the initial ADMIN user.
    """
    def _sync_init() -> None:
        conn = sqlite3.connect(str(DB_PATH), timeout=30.0)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA foreign_keys = ON")
        try:
            conn.execute("PRAGMA journal_mode = WAL")
        except sqlite3.OperationalError:
            conn.execute("PRAGMA journal_mode = MEMORY")
        try:
            conn.executescript(SCHEMA_SQL)

            # Non-destructive column migration guard for future schema evolutions
            migrations: List[Tuple[str, str, str]] = [
                ("users", "theme_mode", "TEXT NOT NULL DEFAULT 'system'"),
                ("users", "language", "TEXT NOT NULL DEFAULT 'ar'"),
                ("students", "admin_notes", "TEXT"),
                ("classes", "billing_mode", "TEXT NOT NULL DEFAULT 'MONTHLY'"),
                ("exams", "show_ranking", "INTEGER NOT NULL DEFAULT 1"),
                ("payments", "is_voided", "INTEGER NOT NULL DEFAULT 0"),
                ("notifications", "dedup_key", "TEXT"),
            ]
            for table_name, col_name, col_def in migrations:
                cols = [r["name"] for r in conn.execute(f"PRAGMA table_info({table_name})").fetchall()]
                if col_name not in cols:
                    conn.execute(f"ALTER TABLE {table_name} ADD COLUMN {col_name} {col_def}")

            # Default school settings
            default_settings = {
                "school_name_ar": ("مدرسة ومؤسسة مونيكس التعليمية — MOUNIX SCHOOL", "الاسم الرسمي للمدرسة بالعربية"),
                "school_name_en": ("MOUNIX SCHOOL — Educational Center", "Official School Name in English"),
                "school_phone": ("+213 550 00 00 00", "رقم هاتف المدرسة"),
                "school_address": ("قسنطينة، الجزائر — Constantine, Algeria", "عنوان المدرسة"),
                "currency": ("DZD", "العملة المعتمدة"),
                "enable_grade_ranking": ("1", "تفعيل ترتيب التلاميذ في النتائج"),
                "backup_retention_count": (str(BACKUP_RETENTION_COUNT), "عدد النسخ الاحتياطية المحفوظة"),
                "allow_capacity_override_admin": ("1", "السماح للمدير بتجاوز الطاقة الاستيعابية للأقسام"),
                "schema_version": ("2026.11.1", "إصدار مخطط قاعدة البيانات"),
            }
            for k, (v, desc) in default_settings.items():
                conn.execute(
                    "INSERT OR IGNORE INTO app_settings (key, value, description) VALUES (?, ?, ?)",
                    (k, v, desc),
                )

            # Bootstrap initial ADMIN account if none exists
            admin_row = conn.execute("SELECT id FROM users WHERE role = 'ADMIN' AND is_deleted = 0 LIMIT 1").fetchone()
            if admin_row is None:
                pwd_hash = hash_password(DEFAULT_ADMIN_PASSWORD)
                cur = conn.execute(
                    """
                    INSERT INTO users (username, password_hash, full_name, role, phone, email, telegram_id, language, is_active)
                    VALUES (?, ?, ?, 'ADMIN', ?, ?, ?, 'ar', 1)
                    """,
                    (
                        DEFAULT_ADMIN_USERNAME,
                        pwd_hash,
                        "المدير العام — MOUNIX ADMIN",
                        "+213550000000",
                        "admin@mounixschool.dz",
                        DEFAULT_ADMIN_TELEGRAM_ID,
                    ),
                )
                admin_id = cur.lastrowid
                conn.execute(
                    "INSERT OR IGNORE INTO notification_preferences (user_id) VALUES (?)",
                    (admin_id,),
                )
                conn.execute(
                    """
                    INSERT INTO audit_logs (actor_user_id, actor_username, actor_role, action, entity_type, entity_id, details_json)
                    VALUES (?, ?, 'ADMIN', 'BOOTSTRAP_ADMIN', 'users', ?, ?)
                    """,
                    (admin_id, DEFAULT_ADMIN_USERNAME, str(admin_id), json.dumps({"note": "Initial admin created"})),
                )
            conn.commit()
        finally:
            conn.close()

    await asyncio.to_thread(_sync_init)
    logger.info("Database initialized and verified (24 relational tables active).")


# ==============================================================================
# 5. AUDIT LOGGING & UNIQUE CODE GENERATORS
# ==============================================================================

async def log_audit(
    session: Optional[AsyncDBSession],
    actor_user_id: Optional[int],
    actor_username: Optional[str],
    actor_role: Optional[str],
    action: str,
    entity_type: str,
    entity_id: Optional[Union[int, str]] = None,
    details: Optional[Dict[str, Any]] = None,
    ip_address: Optional[str] = None,
) -> None:
    sql = """
        INSERT INTO audit_logs (actor_user_id, actor_username, actor_role, action, entity_type, entity_id, details_json, ip_address)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    """
    params = (
        actor_user_id,
        actor_username or "system",
        actor_role or "SYSTEM",
        action,
        entity_type,
        str(entity_id) if entity_id is not None else None,
        json.dumps(details or {}, ensure_ascii=False),
        ip_address,
    )
    if session is not None:
        await session.execute(sql, params)
    else:
        async with db.transaction() as tx:
            await tx.execute(sql, params)


async def generate_unique_code(session: AsyncDBSession, prefix: str, table: str, column: str) -> str:
    """Generates sequential + random collision-safe unique codes like STU-2026-0001."""
    year = datetime.date.today().year
    count_row = await session.fetchone(f"SELECT COUNT(*) AS cnt FROM {table}")
    seq = (int(count_row["cnt"]) if count_row else 0) + 1
    for attempt in range(100):
        candidate = f"{prefix}-{year}-{seq + attempt:04d}"
        exists = await session.fetchone(f"SELECT 1 FROM {table} WHERE {column} = ?", (candidate,))
        if not exists:
            return candidate
    return f"{prefix}-{year}-{secrets.token_hex(3).upper()}"


# ==============================================================================
# 6. NOTIFICATION DISPATCHER (TELEGRAM + FLUTTER IN-APP)
# ==============================================================================

_telegram_send_callback: Optional[Callable[[int, str], Coroutine[Any, Any, bool]]] = None


def register_telegram_notifier(callback: Callable[[int, str], Coroutine[Any, Any, bool]]) -> None:
    global _telegram_send_callback
    _telegram_send_callback = callback


async def create_and_dispatch_notification(
    session: AsyncDBSession,
    recipient_user_id: int,
    notification_type: str,
    title_ar: str,
    title_en: str,
    body_ar: str,
    body_en: str,
    sender_user_id: Optional[int] = None,
    dedup_key: Optional[str] = None,
    related_entity_type: Optional[str] = None,
    related_entity_id: Optional[int] = None,
) -> Optional[int]:
    """
    Saves notification in database (respecting user preferences & deduplication key)
    and dispatches to Telegram if the user has a linked telegram_id and telegram notifications enabled.
    """
    if dedup_key:
        existing = await session.fetchone("SELECT id FROM notifications WHERE dedup_key = ?", (dedup_key,))
        if existing:
            return int(existing["id"])

    user = await session.fetchone(
        "SELECT id, telegram_id, language, is_active, is_deleted FROM users WHERE id = ?",
        (recipient_user_id,),
    )
    if not user or not user["is_active"] or user["is_deleted"]:
        return None

    prefs = await session.fetchone(
        "SELECT * FROM notification_preferences WHERE user_id = ?",
        (recipient_user_id,),
    )
    if not prefs:
        await session.execute(
            "INSERT OR IGNORE INTO notification_preferences (user_id) VALUES (?)",
            (recipient_user_id,),
        )
        prefs = {
            "enable_telegram": 1, "enable_app": 1,
            "notify_absence": 1, "notify_grades": 1, "notify_payments": 1,
            "notify_schedule": 1, "notify_materials": 1, "notify_announcements": 1,
        }

    type_to_pref = {
        "ABSENCE": "notify_absence",
        "NEW_GRADE": "notify_grades",
        "PAYMENT_DUE": "notify_payments",
        "PAYMENT_RECEIVED": "notify_payments",
        "SCHEDULE_CHANGE": "notify_schedule",
        "CLASS_REMINDER": "notify_schedule",
        "NEW_MATERIAL": "notify_materials",
        "NEW_ASSIGNMENT": "notify_materials",
        "ANNOUNCEMENT": "notify_announcements",
    }
    pref_col = type_to_pref.get(notification_type)
    if pref_col and not int(prefs.get(pref_col, 1)):
        return None

    notif_id = await session.insert(
        """
        INSERT INTO notifications (
            recipient_user_id, sender_user_id, notification_type,
            title_ar, title_en, body_ar, body_en, dedup_key,
            related_entity_type, related_entity_id
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            recipient_user_id, sender_user_id, notification_type,
            title_ar, title_en, body_ar, body_en, dedup_key,
            related_entity_type, related_entity_id,
        ),
    )

    if int(prefs.get("enable_telegram", 1)) and user.get("telegram_id") and _telegram_send_callback is not None:
        lang = user.get("language", "ar")
        title = title_ar if lang == "ar" else title_en
        body = body_ar if lang == "ar" else body_en
        msg_text = f"🔔 *{title}*\n\n{body}"
        try:
            sent = await _telegram_send_callback(int(user["telegram_id"]), msg_text)
            if sent:
                await session.execute("UPDATE notifications SET telegram_sent = 1 WHERE id = ?", (notif_id,))
        except Exception as exc:
            logger.warning(f"Failed sending Telegram alert to user {recipient_user_id}: {exc}")

    return notif_id


async def notify_student_and_parents(
    session: AsyncDBSession,
    student_id: int,
    notification_type: str,
    title_ar: str,
    title_en: str,
    body_ar: str,
    body_en: str,
    sender_user_id: Optional[int] = None,
    dedup_prefix: Optional[str] = None,
    related_entity_type: Optional[str] = None,
    related_entity_id: Optional[int] = None,
) -> List[int]:
    """Notifies the student user account and all linked parent user accounts."""
    recipient_user_ids: set[int] = set()
    st = await session.fetchone("SELECT user_id, primary_parent_id FROM students WHERE id = ?", (student_id,))
    if st and st.get("user_id"):
        recipient_user_ids.add(int(st["user_id"]))
    if st and st.get("primary_parent_id"):
        p_row = await session.fetchone("SELECT user_id FROM parents WHERE id = ?", (st["primary_parent_id"],))
        if p_row and p_row.get("user_id"):
            recipient_user_ids.add(int(p_row["user_id"]))

    linked_parents = await session.fetchall(
        """
        SELECT p.user_id
        FROM parent_students ps
        JOIN parents p ON p.id = ps.parent_id
        WHERE ps.student_id = ? AND p.user_id IS NOT NULL
        """,
        (student_id,),
    )
    for lp in linked_parents:
        recipient_user_ids.add(int(lp["user_id"]))

    created_ids: List[int] = []
    for uid in recipient_user_ids:
        dkey = f"{dedup_prefix}:u{uid}" if dedup_prefix else None
        nid = await create_and_dispatch_notification(
            session=session,
            recipient_user_id=uid,
            notification_type=notification_type,
            title_ar=title_ar,
            title_en=title_en,
            body_ar=body_ar,
            body_en=body_en,
            sender_user_id=sender_user_id,
            dedup_key=dkey,
            related_entity_type=related_entity_type,
            related_entity_id=related_entity_id,
        )
        if nid:
            created_ids.append(nid)
    return created_ids


# ==============================================================================
# 7. SQLITE BACKUP & RESTORE ENGINE (WITH INTEGRITY VERIFICATION)
# ==============================================================================

async def create_database_backup(actor_user_id: Optional[int] = None, actor_username: Optional[str] = None) -> Dict[str, Any]:
    """
    Uses SQLite native `.backup()` API to safely copy the live database while online,
    verifies `PRAGMA integrity_check` on the backup file, calculates SHA-256 checksum,
    logs it in `backup_logs`, and enforces retention count.
    """
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    filename = f"mounix_backup_{timestamp}.db"
    dest_path = BACKUP_DIR / filename

    def _do_backup() -> Tuple[int, str, str]:
        src_conn = sqlite3.connect(str(DB_PATH), timeout=30.0)
        dst_conn = sqlite3.connect(str(dest_path), timeout=30.0)
        try:
            src_conn.backup(dst_conn)
        finally:
            dst_conn.close()
            src_conn.close()

        # Verify integrity of backup
        chk_conn = sqlite3.connect(str(dest_path), timeout=30.0)
        try:
            row = chk_conn.execute("PRAGMA integrity_check").fetchone()
            integrity = str(row[0]) if row else "FAILED"
        finally:
            chk_conn.close()

        if integrity.lower() != "ok":
            dest_path.unlink(missing_ok=True)
            raise RuntimeError(f"Backup integrity check failed: {integrity}")

        data_bytes = dest_path.read_bytes()
        sha256_hex = hashlib.sha256(data_bytes).hexdigest()
        return len(data_bytes), sha256_hex, "OK"

    size_bytes, sha256_hex, status_str = await asyncio.to_thread(_do_backup)

    async with db.transaction() as session:
        retention_row = await session.fetchone(
            "SELECT value FROM app_settings WHERE key = 'backup_retention_count'"
        )
        retention_limit = int(retention_row["value"]) if retention_row and str(retention_row["value"]).isdigit() else BACKUP_RETENTION_COUNT

        backup_id = await session.insert(
            """
            INSERT INTO backup_logs (filename, file_path, size_bytes, checksum_sha256, integrity_status, created_by)
            VALUES (?, ?, ?, ?, ?, ?)
            """,
            (filename, str(dest_path), size_bytes, sha256_hex, status_str, actor_user_id),
        )
        await log_audit(
            session, actor_user_id, actor_username, "ADMIN",
            "CREATE_BACKUP", "backup_logs", backup_id,
            {"filename": filename, "size_bytes": size_bytes, "sha256": sha256_hex},
        )

        # Enforce retention policy
        all_backups = await session.fetchall(
            "SELECT id, filename, file_path FROM backup_logs ORDER BY id DESC"
        )
        if len(all_backups) > retention_limit:
            for old_b in all_backups[retention_limit:]:
                old_p = Path(old_b["file_path"])
                if old_p.exists():
                    with contextlib.suppress(Exception):
                        old_p.unlink()
                await session.execute("DELETE FROM backup_logs WHERE id = ?", (old_b["id"],))

        created_row = await session.fetchone("SELECT * FROM backup_logs WHERE id = ?", (backup_id,))
        return created_row or {}


async def restore_database_backup(
    backup_id_or_filename: Union[int, str],
    actor_user_id: Optional[int] = None,
    actor_username: Optional[str] = None,
) -> Dict[str, Any]:
    """
    Verifies the target backup file exists, passes SQLite `PRAGMA integrity_check`
    and has all required tables before safely restoring it into the active database
    via SQLite Backup API.
    """
    if isinstance(backup_id_or_filename, int) or str(backup_id_or_filename).isdigit():
        b_row = await db.fetchone("SELECT * FROM backup_logs WHERE id = ?", (int(backup_id_or_filename),))
    else:
        b_row = await db.fetchone("SELECT * FROM backup_logs WHERE filename = ?", (str(backup_id_or_filename),))

    if not b_row:
        raise ValueError("Backup record not found")

    src_path = Path(b_row["file_path"])
    if not src_path.exists():
        raise ValueError(f"Backup file missing on disk: {src_path.name}")

    def _verify_and_restore() -> None:
        # 1. Verify integrity & checksum before touching live DB
        chk_conn = sqlite3.connect(str(src_path), timeout=30.0)
        try:
            res = chk_conn.execute("PRAGMA integrity_check").fetchone()
            if not res or str(res[0]).lower() != "ok":
                raise ValueError("Backup file failed PRAGMA integrity_check")
            tables = {
                r[0]
                for r in chk_conn.execute("SELECT name FROM sqlite_master WHERE type='table'").fetchall()
            }
            if "users" not in tables or "students" not in tables:
                raise ValueError("Backup file is missing core MOUNIX SCHOOL tables")
        finally:
            chk_conn.close()

        # 2. Restore using SQLite backup API inside write lock
        src_conn = sqlite3.connect(str(src_path), timeout=30.0)
        dst_conn = sqlite3.connect(str(DB_PATH), timeout=30.0)
        try:
            src_conn.backup(dst_conn)
        finally:
            dst_conn.close()
            src_conn.close()

    async with db._write_lock:
        await asyncio.to_thread(_verify_and_restore)

    async with db.transaction() as session:
        await session.execute(
            "UPDATE backup_logs SET restored_at = datetime('now') WHERE id = ?",
            (b_row["id"],),
        )
        await log_audit(
            session, actor_user_id, actor_username, "ADMIN",
            "RESTORE_BACKUP", "backup_logs", b_row["id"],
            {"filename": b_row["filename"]},
        )

    return {
        "status": "RESTORED",
        "backup_id": b_row["id"],
        "filename": b_row["filename"],
        "restored_at": datetime.datetime.now().isoformat(),
    }


# ==============================================================================
# 8. ARABIC RTL PDF & EXCEL REPORTING ENGINE
# ==============================================================================

def generate_excel_workbook_bytes(
    title_ar: str,
    subtitle_ar: str,
    headers: List[str],
    rows: List[List[Any]],
    summary_pairs: Optional[List[Tuple[str, Any]]] = None,
) -> bytes:
    """
    Generates a professionally styled Right-to-Left (RTL) Arabic Excel (.xlsx) report
    using openpyxl.
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "MOUNIX REPORT"
    ws.sheet_view.rightToLeft = True

    navy_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    sub_fill = PatternFill(start_color="EFF6FF", end_color="EFF6FF", fill_type="solid")
    hdr_fill = PatternFill(start_color="2563EB", end_color="2563EB", fill_type="solid")
    zebra_fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
    border_side = Side(style="thin", color="CBD5E1")
    cell_border = Border(left=border_side, right=border_side, top=border_side, bottom=border_side)

    col_count = max(len(headers), 4)
    last_col_letter = get_column_letter(col_count)

    # Title banner
    ws.merge_cells(f"A1:{last_col_letter}1")
    ws["A1"] = f"MOUNIX SCHOOL — {title_ar}"
    ws["A1"].font = Font(name="Noto Sans Arabic", size=14, bold=True, color="FFFFFF")
    ws["A1"].fill = navy_fill
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 30

    # Subtitle banner
    ws.merge_cells(f"A2:{last_col_letter}2")
    ws["A2"] = f"{subtitle_ar} | تاريخ الإصدار: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M')}"
    ws["A2"].font = Font(name="Noto Sans Arabic", size=10, italic=False, color="1E293B")
    ws["A2"].fill = sub_fill
    ws["A2"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 22

    current_row = 4
    if summary_pairs:
        for label, val in summary_pairs:
            ws.cell(row=current_row, column=1, value=label).font = Font(bold=True, color="1E3A8A")
            ws.cell(row=current_row, column=2, value=val).font = Font(bold=True)
            current_row += 1
        current_row += 1

    # Table Headers
    ws.row_dimensions[current_row].height = 24
    for col_idx, h_text in enumerate(headers, start=1):
        c = ws.cell(row=current_row, column=col_idx, value=h_text)
        c.font = Font(name="Noto Sans Arabic", size=11, bold=True, color="FFFFFF")
        c.fill = hdr_fill
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = cell_border

    # Table Data Rows
    for r_idx, row_vals in enumerate(rows, start=current_row + 1):
        ws.row_dimensions[r_idx].height = 20
        is_even = (r_idx % 2 == 0)
        for c_idx, val in enumerate(row_vals, start=1):
            c = ws.cell(row=r_idx, column=c_idx, value=val)
            c.font = Font(name="Noto Sans Arabic", size=10, color="0F172A")
            if is_even:
                c.fill = zebra_fill
            c.alignment = Alignment(
                horizontal="right" if isinstance(val, str) else "center",
                vertical="center",
                wrap_text=True,
            )
            c.border = cell_border

    # Auto-fit columns
    for col_idx in range(1, len(headers) + 1):
        col_letter = get_column_letter(col_idx)
        max_len = len(str(headers[col_idx - 1]))
        for r in rows:
            if col_idx - 1 < len(r) and r[col_idx - 1] is not None:
                max_len = max(max_len, len(str(r[col_idx - 1])))
        ws.column_dimensions[col_letter].width = max(min(max_len + 6, 45), 14)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def generate_pdf_report_bytes(
    title_ar: str,
    subtitle_ar: str,
    headers: List[str],
    rows: List[List[Any]],
    summary_cards: Optional[List[Tuple[str, Any]]] = None,
    footer_note: str = "MOUNIX SCHOOL — نظام الإدارة المدرسية المتكامل",
) -> bytes:
    """
    Generates an Arabic RTL PDF report using WeasyPrint (with full Arabic glyph shaping
    and Noto Sans Arabic font) or ReportLab fallback if WeasyPrint is unavailable.
    """
    cards_html = ""
    if summary_cards:
        items = "".join(
            f'<div class="card"><div class="card-label">{lbl}</div><div class="card-val">{val}</div></div>'
            for lbl, val in summary_cards
        )
        cards_html = f'<div class="cards">{items}</div>'

    ths = "".join(f"<th>{h}</th>" for h in headers)
    trs = ""
    for r in rows:
        tds = "".join(f"<td>{'' if v is None else v}</td>" for v in r)
        trs += f"<tr>{tds}</tr>"
    if not rows:
        trs = f'<tr><td colspan="{len(headers)}" style="text-align:center;">لا توجد سجلات مطابقة</td></tr>'

    now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
    html_doc = f"""<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
<meta charset="utf-8">
<style>
  @page {{
    size: A4 landscape;
    margin: 14mm;
    @bottom-center {{
      content: "{footer_note} — صفحة " counter(page) " من " counter(pages);
      font-family: 'Noto Sans Arabic', 'DejaVu Sans', sans-serif;
      font-size: 9pt;
      color: #64748B;
    }}
  }}
  body {{
    font-family: 'Noto Sans Arabic', 'Noto Kufi Arabic', 'DejaVu Sans', sans-serif;
    direction: rtl;
    text-align: right;
    color: #0F172A;
    margin: 0;
    padding: 0;
  }}
  .header {{
    background: linear-gradient(135deg, #1E3A8A, #2563EB);
    color: #FFFFFF;
    padding: 16px 22px;
    border-radius: 8px;
    margin-bottom: 16px;
  }}
  .header h1 {{
    margin: 0 0 6px 0;
    font-size: 18pt;
  }}
  .header p {{
    margin: 0;
    font-size: 10pt;
    opacity: 0.9;
  }}
  .cards {{
    display: flex;
    flex-wrap: wrap;
    gap: 12px;
    margin-bottom: 16px;
  }}
  .card {{
    flex: 1;
    min-width: 130px;
    background: #F1F5F9;
    border-right: 4px solid #2563EB;
    padding: 10px 14px;
    border-radius: 6px;
  }}
  .card-label {{
    font-size: 9pt;
    color: #475569;
  }}
  .card-val {{
    font-size: 13pt;
    font-weight: bold;
    color: #1E3A8A;
    margin-top: 4px;
  }}
  table {{
    width: 100%;
    border-collapse: collapse;
    font-size: 9.5pt;
  }}
  th {{
    background-color: #1E3A8A;
    color: #FFFFFF;
    padding: 8px 10px;
    border: 1px solid #CBD5E1;
    font-weight: bold;
  }}
  td {{
    padding: 7px 10px;
    border: 1px solid #CBD5E1;
  }}
  tr:nth-child(even) td {{
    background-color: #F8FAFC;
  }}
</style>
</head>
<body>
  <div class="header">
    <h1>MOUNIX SCHOOL — {title_ar}</h1>
    <p>{subtitle_ar} | تاريخ الإصدار: {now_str}</p>
  </div>
  {cards_html}
  <table>
    <thead><tr>{ths}</tr></thead>
    <tbody>{trs}</tbody>
  </table>
</body>
</html>"""

    if HAS_WEASYPRINT and weasyprint is not None:
        return weasyprint.HTML(string=html_doc).write_pdf()

    # Fallback minimal valid PDF if WeasyPrint is not installed
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.pdfgen import canvas as rl_canvas
    buf = io.BytesIO()
    c = rl_canvas.Canvas(buf, pagesize=landscape(A4))
    c.setFont("Helvetica-Bold", 14)
    c.drawString(40, 540, f"MOUNIX SCHOOL - {title_ar}")
    c.setFont("Helvetica", 10)
    c.drawString(40, 520, f"Generated: {now_str}")
    y = 490
    c.drawString(40, y, " | ".join(str(h) for h in headers))
    y -= 20
    for r in rows[:20]:
        c.drawString(40, y, " | ".join(str(v) for v in r))
        y -= 18
    c.showPage()
    c.save()
    return buf.getvalue()


def generate_payment_receipt_pdf_bytes(receipt_data: Dict[str, Any]) -> bytes:
    """Generates an official Arabic RTL Payment Receipt PDF for a transaction."""
    alloc_rows_html = ""
    for al in receipt_data.get("allocations", []):
        alloc_rows_html += f"""
        <tr>
          <td>{al.get('plan_code', '-')}</td>
          <td>{al.get('period_label') or al.get('plan_type', '-')}</td>
          <td>{al.get('allocated_amount', 0):.2f} {receipt_data.get('currency', 'DZD')}</td>
          <td>{al.get('remaining_after', 0):.2f} {receipt_data.get('currency', 'DZD')}</td>
        </tr>
        """
    if not alloc_rows_html:
        alloc_rows_html = f"""
        <tr>
          <td colspan="4" style="text-align:center;">دفعة مالية مباشرة — {receipt_data.get('description') or 'اشتراك دراسي'}</td>
        </tr>
        """

    html_str = f"""<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
<meta charset="utf-8">
<style>
  @page {{ size: A5 landscape; margin: 12mm; }}
  body {{
    font-family: 'Noto Sans Arabic', 'DejaVu Sans', sans-serif;
    direction: rtl;
    color: #0F172A;
    margin: 0;
  }}
  .box {{
    border: 2px solid #1E3A8A;
    border-radius: 10px;
    padding: 16px;
  }}
  .top {{
    display: flex;
    justify-content: space-between;
    border-bottom: 2px solid #DBEAFE;
    padding-bottom: 10px;
    margin-bottom: 12px;
  }}
  .school-title {{ font-size: 15pt; font-weight: bold; color: #1E3A8A; }}
  .receipt-badge {{
    background: #1E3A8A;
    color: #fff;
    padding: 6px 14px;
    border-radius: 6px;
    font-weight: bold;
    font-size: 11pt;
  }}
  .grid {{
    display: table;
    width: 100%;
    margin-bottom: 12px;
    font-size: 10pt;
  }}
  .row {{ display: table-row; }}
  .cell {{ display: table-cell; padding: 5px 8px; }}
  .label {{ font-weight: bold; color: #475569; width: 22%; }}
  table.alloc {{
    width: 100%;
    border-collapse: collapse;
    font-size: 9.5pt;
    margin-top: 8px;
  }}
  table.alloc th {{
    background: #EFF6FF;
    color: #1E3A8A;
    border: 1px solid #CBD5E1;
    padding: 6px;
  }}
  table.alloc td {{
    border: 1px solid #CBD5E1;
    padding: 6px;
  }}
  .footer {{
    margin-top: 18px;
    display: flex;
    justify-content: space-between;
    font-size: 9.5pt;
    color: #334155;
  }}
</style>
</head>
<body>
  <div class="box">
    <div class="top">
      <div>
        <div class="school-title">{receipt_data.get('school_name_ar', 'MOUNIX SCHOOL')}</div>
        <div style="font-size:9pt;color:#475569;">{receipt_data.get('school_address', '')} | {receipt_data.get('school_phone', '')}</div>
      </div>
      <div class="receipt-badge">إيصال دفع رقم: {receipt_data.get('receipt_number')}</div>
    </div>

    <div class="grid">
      <div class="row">
        <div class="cell label">اسم التلميذ:</div>
        <div class="cell">{receipt_data.get('student_name')} ({receipt_data.get('student_code')})</div>
        <div class="cell label">تاريخ الدفع:</div>
        <div class="cell">{receipt_data.get('payment_date')}</div>
      </div>
      <div class="row">
        <div class="cell label">المبلغ المدفوع:</div>
        <div class="cell" style="font-weight:bold;color:#15803D;font-size:12pt;">{receipt_data.get('amount', 0):.2f} {receipt_data.get('currency', 'DZD')}</div>
        <div class="cell label">طريقة الدفع:</div>
        <div class="cell">{receipt_data.get('payment_method')}</div>
      </div>
      <div class="row">
        <div class="cell label">البيان / الوصف:</div>
        <div class="cell">{receipt_data.get('description') or 'تسديد مستحقات دراسية'}</div>
        <div class="cell label">الموظف المستلم:</div>
        <div class="cell">{receipt_data.get('recorded_by_name')}</div>
      </div>
    </div>

    <table class="alloc">
      <thead>
        <tr>
          <th>رمز الاشتراك</th>
          <th>الفترة / النوع</th>
          <th>المبلغ المخصص</th>
          <th>المتبقي من الاشتراك</th>
        </tr>
      </thead>
      <tbody>
        {alloc_rows_html}
      </tbody>
    </table>

    <div class="footer">
      <div>إجمالي الديون المتبقية على التلميذ: <strong>{receipt_data.get('total_student_debt', 0):.2f} {receipt_data.get('currency', 'DZD')}</strong></div>
      <div>توقيع الإدارة والختم: ______________________</div>
    </div>
  </div>
</body>
</html>"""

    if HAS_WEASYPRINT and weasyprint is not None:
        return weasyprint.HTML(string=html_str).write_pdf()
    return generate_pdf_report_bytes(
        title_ar=f"إيصال دفع {receipt_data.get('receipt_number')}",
        subtitle_ar=f"التلميذ: {receipt_data.get('student_name')}",
        headers=["رقم الإيصال", "التلميذ", "المبلغ", "طريقة الدفع", "التاريخ", "الموظف"],
        rows=[[
            receipt_data.get("receipt_number"),
            receipt_data.get("student_name"),
            receipt_data.get("amount"),
            receipt_data.get("payment_method"),
            receipt_data.get("payment_date"),
            receipt_data.get("recorded_by_name"),
        ]],
    )


# ==============================================================================
# 9. CORE BUSINESS LOGIC SERVICES (UNIFIED FOR BOT & REST API)
# ==============================================================================

def _time_to_minutes(time_str: str) -> int:
    """Converts HH:MM string to minutes since midnight."""
    m = re.match(r"^(\d{1,2}):(\d{2})$", time_str.strip())
    if not m:
        raise ValueError(f"Invalid time format '{time_str}', expected HH:MM (24h)")
    hh, mm = int(m.group(1)), int(m.group(2))
    if not (0 <= hh <= 23 and 0 <= mm <= 59):
        raise ValueError(f"Time out of range: {time_str}")
    return hh * 60 + mm


async def check_schedule_conflicts(
    session: AsyncDBSession,
    day_of_week: int,
    start_time: str,
    end_time: str,
    teacher_id: int,
    room: str,
    class_id: int,
    specific_date: Optional[str] = None,
    exclude_schedule_id: Optional[int] = None,
) -> Optional[str]:
    """
    Validates that no overlapping active schedule exists on the same day/date for:
    1. The same Teacher
    2. The same Room
    3. The same Class
    Returns a descriptive Arabic & English conflict message if a conflict is found, or None if clear.
    """
    start_min = _time_to_minutes(start_time)
    end_min = _time_to_minutes(end_time)
    if end_min <= start_min:
        return "وقت النهاية يجب أن يكون بعد وقت البداية (End time must be after start time)"

    params: List[Any] = [day_of_week]
    sql = """
        SELECT s.*, c.name AS class_name, t.full_name AS teacher_name
        FROM schedules s
        JOIN classes c ON c.id = s.class_id
        JOIN teachers t ON t.id = s.teacher_id
        WHERE s.status = 'SCHEDULED'
          AND s.day_of_week = ?
    """
    if exclude_schedule_id is not None:
        sql += " AND s.id != ?"
        params.append(exclude_schedule_id)

    candidates = await session.fetchall(sql, params)
    for row in candidates:
        # Check specific date compatibility
        if specific_date and row.get("specific_date") and row["specific_date"] != specific_date:
            continue
        ex_start = _time_to_minutes(row["start_time"])
        ex_end = _time_to_minutes(row["end_time"])
        # Overlap check: max(start1, start2) < min(end1, end2)
        if max(start_min, ex_start) < min(end_min, ex_end):
            if int(row["teacher_id"]) == int(teacher_id):
                return (
                    f"تعارض في جدول الأستاذ ({row['teacher_name']}) مع القسم '{row['class_name']}' "
                    f"من {row['start_time']} إلى {row['end_time']}."
                )
            if str(row["room"]).strip().lower() == str(room).strip().lower():
                return (
                    f"تعارض في القاعة ({room}) المشغولة بالقسم '{row['class_name']}' "
                    f"من {row['start_time']} إلى {row['end_time']}."
                )
            if int(row["class_id"]) == int(class_id):
                return (
                    f"تعارض في جدول القسم ({row['class_name']}) الذي لديه حصة أخرى "
                    f"من {row['start_time']} إلى {row['end_time']}."
                )
    return None


async def can_user_access_student(user: Dict[str, Any], student_id: int) -> bool:
    """
    Enforces strict backend data ownership & visibility rules:
    - ADMIN: can access all students
    - TEACHER: can only access students enrolled in classes taught by this teacher
    - STUDENT: can only access their own student record
    - PARENT: can only access students linked to their parent profile
    """
    role = user["role"]
    if role == UserRole.ADMIN.value:
        return True
    if role == UserRole.STUDENT.value:
        st = await db.fetchone("SELECT id FROM students WHERE id = ? AND user_id = ?", (student_id, user["id"]))
        return st is not None
    if role == UserRole.PARENT.value:
        parent = await db.fetchone("SELECT id FROM parents WHERE user_id = ?", (user["id"],))
        if not parent:
            return False
        pid = parent["id"]
        link = await db.fetchone(
            """
            SELECT 1 FROM students s
            LEFT JOIN parent_students ps ON ps.student_id = s.id
            WHERE s.id = ? AND (s.primary_parent_id = ? OR ps.parent_id = ?)
            """,
            (student_id, pid, pid),
        )
        return link is not None
    if role == UserRole.TEACHER.value:
        teacher = await db.fetchone("SELECT id FROM teachers WHERE user_id = ?", (user["id"],))
        if not teacher:
            return False
        enr = await db.fetchone(
            """
            SELECT 1 FROM enrollments e
            JOIN classes c ON c.id = e.class_id
            WHERE e.student_id = ? AND c.teacher_id = ?
            """,
            (student_id, teacher["id"]),
        )
        return enr is not None
    return False


async def get_student_full_dossier(student_id: int) -> Optional[Dict[str, Any]]:
    """Returns the complete 360-degree student profile required by Section 4."""
    student = await db.fetchone(
        """
        SELECT s.*, l.name_ar AS level_name_ar, l.name_en AS level_name_en,
               p.full_name AS parent_name, p.phone AS parent_phone, p.email AS parent_email,
               u.username AS username, u.telegram_id AS telegram_id
        FROM students s
        LEFT JOIN levels l ON l.id = s.level_id
        LEFT JOIN parents p ON p.id = s.primary_parent_id
        LEFT JOIN users u ON u.id = s.user_id
        WHERE s.id = ?
        """,
        (student_id,),
    )
    if not student:
        return None

    parents = await db.fetchall(
        """
        SELECT p.id, p.parent_code, p.full_name, p.phone, p.email, ps.relationship, ps.is_primary
        FROM parent_students ps
        JOIN parents p ON p.id = ps.parent_id
        WHERE ps.student_id = ?
        """,
        (student_id,),
    )

    classes = await db.fetchall(
        """
        SELECT c.id, c.code, c.name, c.room, c.billing_mode, c.monthly_fee, c.session_fee,
               sub.id AS subject_id, sub.name_ar AS subject_name_ar, sub.name_en AS subject_name_en,
               t.full_name AS teacher_name, e.status AS enrollment_status, e.enrolled_at, e.discount_percent
        FROM enrollments e
        JOIN classes c ON c.id = e.class_id
        JOIN subjects sub ON sub.id = c.subject_id
        JOIN teachers t ON t.id = c.teacher_id
        WHERE e.student_id = ?
        ORDER BY e.enrolled_at DESC
        """,
        (student_id,),
    )

    subjects_map: Dict[int, Dict[str, Any]] = {}
    for cl in classes:
        sid = cl["subject_id"]
        if sid not in subjects_map:
            subjects_map[sid] = {
                "id": sid,
                "name_ar": cl["subject_name_ar"],
                "name_en": cl["subject_name_en"],
                "teacher_name": cl["teacher_name"],
            }

    schedules = await db.fetchall(
        """
        SELECT s.*, c.name AS class_name, sub.name_ar AS subject_name_ar, sub.name_en AS subject_name_en,
               t.full_name AS teacher_name
        FROM schedules s
        JOIN enrollments e ON e.class_id = s.class_id
        JOIN classes c ON c.id = s.class_id
        JOIN subjects sub ON sub.id = s.subject_id
        JOIN teachers t ON t.id = s.teacher_id
        WHERE e.student_id = ? AND e.status = 'ACTIVE' AND s.status = 'SCHEDULED'
        ORDER BY s.day_of_week ASC, s.start_time ASC
        """,
        (student_id,),
    )

    attendance_records = await db.fetchall(
        """
        SELECT a.*, c.name AS class_name, sub.name_ar AS subject_name_ar
        FROM attendance a
        JOIN classes c ON c.id = a.class_id
        JOIN subjects sub ON sub.id = c.subject_id
        WHERE a.student_id = ?
        ORDER BY a.session_date DESC, a.id DESC
        LIMIT 100
        """,
        (student_id,),
    )
    total_att = len(attendance_records)
    present_cnt = sum(1 for r in attendance_records if r["status"] in ("PRESENT", "LATE"))
    attendance_rate = round((present_cnt / total_att) * 100.0, 1) if total_att > 0 else 100.0

    grades = await db.fetchall(
        """
        SELECT g.*, ex.title AS exam_title, ex.exam_date, ex.max_score, ex.weight,
               sub.name_ar AS subject_name_ar, sub.name_en AS subject_name_en, c.name AS class_name
        FROM grades g
        JOIN exams ex ON ex.id = g.exam_id
        JOIN subjects sub ON sub.id = ex.subject_id
        JOIN classes c ON c.id = ex.class_id
        WHERE g.student_id = ?
        ORDER BY ex.exam_date DESC
        """,
        (student_id,),
    )
    weighted_sum = 0.0
    weight_total = 0.0
    for gr in grades:
        if gr["max_score"] and gr["max_score"] > 0:
            normalized_20 = (float(gr["score"]) / float(gr["max_score"])) * 20.0
            w = float(gr.get("weight") or 1.0)
            weighted_sum += normalized_20 * w
            weight_total += w
    average_on_20 = round(weighted_sum / weight_total, 2) if weight_total > 0 else 0.0

    payment_plans = await db.fetchall(
        """
        SELECT pp.*, c.name AS class_name, sub.name_ar AS subject_name_ar
        FROM payment_plans pp
        LEFT JOIN classes c ON c.id = pp.class_id
        LEFT JOIN subjects sub ON sub.id = pp.subject_id
        WHERE pp.student_id = ? AND pp.status != 'VOIDED'
        ORDER BY pp.due_date DESC
        """,
        (student_id,),
    )

    payments = await db.fetchall(
        """
        SELECT * FROM payments
        WHERE student_id = ? AND is_voided = 0
        ORDER BY payment_date DESC
        """,
        (student_id,),
    )
    total_paid = round(sum(float(p["amount"]) for p in payments), 2)
    total_debt = round(sum(float(pp["remaining_amount"]) for pp in payment_plans), 2)

    materials = await db.fetchall(
        """
        SELECT m.*, sub.name_ar AS subject_name_ar, c.name AS class_name, t.full_name AS teacher_name
        FROM materials m
        JOIN subjects sub ON sub.id = m.subject_id
        LEFT JOIN classes c ON c.id = m.class_id
        JOIN teachers t ON t.id = m.teacher_id
        WHERE m.is_archived = 0
          AND (m.class_id IN (SELECT class_id FROM enrollments WHERE student_id = ? AND status = 'ACTIVE')
               OR m.subject_id IN (
                   SELECT c2.subject_id FROM enrollments e2 JOIN classes c2 ON c2.id = e2.class_id WHERE e2.student_id = ?
               ))
        ORDER BY m.publish_date DESC
        """,
        (student_id, student_id),
    )

    assignments = await db.fetchall(
        """
        SELECT a.*, sub.name_ar AS subject_name_ar, c.name AS class_name, t.full_name AS teacher_name
        FROM assignments a
        JOIN subjects sub ON sub.id = a.subject_id
        JOIN classes c ON c.id = a.class_id
        JOIN teachers t ON t.id = a.teacher_id
        WHERE a.is_archived = 0
          AND a.class_id IN (SELECT class_id FROM enrollments WHERE student_id = ? AND status = 'ACTIVE')
        ORDER BY a.due_date ASC
        """,
        (student_id,),
    )

    return {
        "student": student,
        "parents": parents,
        "subjects": list(subjects_map.values()),
        "classes": classes,
        "schedules": schedules,
        "attendance": {
            "summary": {
                "total_sessions": total_att,
                "present_count": sum(1 for r in attendance_records if r["status"] == "PRESENT"),
                "absent_count": sum(1 for r in attendance_records if r["status"] == "ABSENT"),
                "late_count": sum(1 for r in attendance_records if r["status"] == "LATE"),
                "excused_count": sum(1 for r in attendance_records if r["status"] == "EXCUSED"),
                "attendance_rate_percent": attendance_rate,
            },
            "records": attendance_records,
        },
        "grades": {
            "average_on_20": average_on_20,
            "total_exams": len(grades),
            "records": grades,
        },
        "finance": {
            "total_paid": total_paid,
            "total_remaining_debt": total_debt,
            "payment_plans": payment_plans,
            "payments": payments,
        },
        "materials": materials,
        "assignments": assignments,
    }


async def record_payment_transaction(
    session: AsyncDBSession,
    student_id: int,
    amount: float,
    payment_method: str,
    description: Optional[str],
    actor_user: Dict[str, Any],
    payment_plan_id: Optional[int] = None,
) -> Dict[str, Any]:
    """
    Records a payment transaction (full, partial, or multi-plan allocation).
    Strictly separates collected revenue (`payments`) from outstanding receivables (`payment_plans.remaining_amount`).
    """
    if amount <= 0:
        raise ValueError("Payment amount must be strictly greater than zero")

    st = await session.fetchone("SELECT id, full_name, student_code FROM students WHERE id = ?", (student_id,))
    if not st:
        raise ValueError("Student not found")

    receipt_number = await generate_unique_code(session, "RCP", "payments", "receipt_number")
    payment_id = await session.insert(
        """
        INSERT INTO payments (
            receipt_number, student_id, amount, payment_method, description,
            recorded_by, recorded_by_name
        ) VALUES (?, ?, ?, ?, ?, ?, ?)
        """,
        (
            receipt_number,
            student_id,
            round(amount, 2),
            payment_method,
            description or "دفعة دراسية",
            actor_user["id"],
            actor_user["full_name"],
        ),
    )

    # Allocate payment across target payment plan or oldest unpaid/partial plans
    remaining_to_allocate = round(amount, 2)
    allocations_made: List[Dict[str, Any]] = []

    if payment_plan_id is not None:
        plans = await session.fetchall(
            "SELECT * FROM payment_plans WHERE id = ? AND student_id = ? AND status != 'VOIDED'",
            (payment_plan_id, student_id),
        )
    else:
        plans = await session.fetchall(
            """
            SELECT * FROM payment_plans
            WHERE student_id = ? AND status IN ('UNPAID', 'PARTIAL', 'OVERDUE') AND remaining_amount > 0
            ORDER BY due_date ASC, id ASC
            """,
            (student_id,),
        )

    for plan in plans:
        if remaining_to_allocate <= 0:
            break
        plan_rem = round(float(plan["remaining_amount"]), 2)
        if plan_rem <= 0:
            continue
        alloc_amt = round(min(remaining_to_allocate, plan_rem), 2)
        new_paid = round(float(plan["paid_amount"]) + alloc_amt, 2)
        new_rem = round(max(0.0, float(plan["total_amount"]) - new_paid), 2)
        new_status = "PAID" if new_rem == 0 else "PARTIAL"

        await session.execute(
            """
            UPDATE payment_plans
            SET paid_amount = ?, remaining_amount = ?, status = ?, updated_at = datetime('now')
            WHERE id = ?
            """,
            (new_paid, new_rem, new_status, plan["id"]),
        )
        await session.insert(
            """
            INSERT INTO payment_allocations (payment_id, payment_plan_id, allocated_amount)
            VALUES (?, ?, ?)
            """,
            (payment_id, plan["id"], alloc_amt),
        )
        remaining_to_allocate = round(remaining_to_allocate - alloc_amt, 2)
        allocations_made.append({
            "payment_plan_id": plan["id"],
            "plan_code": plan["plan_code"],
            "plan_type": plan["plan_type"],
            "period_label": plan["period_label"],
            "allocated_amount": alloc_amt,
            "remaining_after": new_rem,
        })

    await log_audit(
        session,
        actor_user["id"],
        actor_user["username"],
        actor_user["role"],
        "RECORD_PAYMENT",
        "payments",
        payment_id,
        {
            "receipt_number": receipt_number,
            "student_id": student_id,
            "amount": amount,
            "allocations": allocations_made,
        },
    )

    # Notify student and parents
    await notify_student_and_parents(
        session=session,
        student_id=student_id,
        notification_type=NotificationType.PAYMENT_RECEIVED.value,
        title_ar="تأكيد استلام دفعة مالية",
        title_en="Payment Received Confirmation",
        body_ar=(
            f"تم تسجيل دفعة بمبلغ {amount:.2f} د.ج للتلميذ {st['full_name']} "
            f"(رقم الإيصال: {receipt_number})."
        ),
        body_en=(
            f"A payment of {amount:.2f} DZD has been recorded for {st['full_name']} "
            f"(Receipt #{receipt_number})."
        ),
        sender_user_id=actor_user["id"],
        dedup_prefix=f"pay:{payment_id}",
        related_entity_type="payments",
        related_entity_id=payment_id,
    )

    payment_row = await session.fetchone("SELECT * FROM payments WHERE id = ?", (payment_id,))
    return {
        "payment": payment_row,
        "allocations": allocations_made,
        "unallocated_credit": remaining_to_allocate,
    }


# ==============================================================================
# 10. PYDANTIC SCHEMAS & FASTAPI REST API APPLICATION
# ==============================================================================

class LoginRequest(BaseModel):
    username: str = Field(..., min_length=2, max_length=64)
    password: str = Field(..., min_length=4, max_length=128)


class RefreshTokenRequest(BaseModel):
    refresh_token: str


class UserCreateRequest(BaseModel):
    username: str = Field(..., min_length=3, max_length=64)
    password: str = Field(..., min_length=6, max_length=128)
    full_name: str = Field(..., min_length=2, max_length=120)
    role: UserRole
    phone: Optional[str] = None
    email: Optional[str] = None
    telegram_id: Optional[int] = None
    language: str = Field(default="ar")
    theme_mode: str = Field(default="system")
    is_active: bool = True


class UserUpdateRequest(BaseModel):
    full_name: Optional[str] = None
    password: Optional[str] = None
    role: Optional[UserRole] = None
    phone: Optional[str] = None
    email: Optional[str] = None
    telegram_id: Optional[int] = None
    language: Optional[str] = Field(default=None)
    theme_mode: Optional[str] = Field(default=None)
    is_active: Optional[bool] = None


class LevelCreateRequest(BaseModel):
    code: Optional[str] = None
    name_ar: str = Field(..., min_length=2)
    name_en: str = Field(..., min_length=2)
    sort_order: int = 1
    description: Optional[str] = None


class LevelUpdateRequest(BaseModel):
    name_ar: Optional[str] = None
    name_en: Optional[str] = None
    sort_order: Optional[int] = None
    description: Optional[str] = None
    is_archived: Optional[bool] = None


class SubjectCreateRequest(BaseModel):
    code: Optional[str] = None
    name_ar: str = Field(..., min_length=2)
    name_en: str = Field(..., min_length=2)
    description: Optional[str] = None
    level_id: Optional[int] = None
    default_monthly_fee: float = Field(default=0.0, ge=0)
    default_session_fee: float = Field(default=0.0, ge=0)
    teacher_ids: List[int] = Field(default_factory=list)


class SubjectUpdateRequest(BaseModel):
    name_ar: Optional[str] = None
    name_en: Optional[str] = None
    description: Optional[str] = None
    level_id: Optional[int] = None
    default_monthly_fee: Optional[float] = Field(default=None, ge=0)
    default_session_fee: Optional[float] = Field(default=None, ge=0)
    teacher_ids: Optional[List[int]] = None
    is_archived: Optional[bool] = None


class ParentCreateRequest(BaseModel):
    full_name: str = Field(..., min_length=2)
    phone: str = Field(..., min_length=6)
    secondary_phone: Optional[str] = None
    email: Optional[str] = None
    address: Optional[str] = None
    occupation: Optional[str] = None
    notes: Optional[str] = None
    username: Optional[str] = None
    password: Optional[str] = None
    telegram_id: Optional[int] = None
    student_ids: List[int] = Field(default_factory=list)


class ParentUpdateRequest(BaseModel):
    full_name: Optional[str] = None
    phone: Optional[str] = None
    secondary_phone: Optional[str] = None
    email: Optional[str] = None
    address: Optional[str] = None
    occupation: Optional[str] = None
    notes: Optional[str] = None
    is_archived: Optional[bool] = None
    student_ids: Optional[List[int]] = None


class StudentCreateRequest(BaseModel):
    full_name: str = Field(..., min_length=2)
    birth_date: Optional[str] = None
    gender: str = Field(default="M")
    level_id: Optional[int] = None
    primary_parent_id: Optional[int] = None
    phone: Optional[str] = None
    address: Optional[str] = None
    joined_date: Optional[str] = None
    admin_notes: Optional[str] = None
    username: Optional[str] = None
    password: Optional[str] = None
    telegram_id: Optional[int] = None
    class_ids: List[int] = Field(default_factory=list)


class StudentUpdateRequest(BaseModel):
    full_name: Optional[str] = None
    birth_date: Optional[str] = None
    gender: Optional[str] = Field(default=None)
    level_id: Optional[int] = None
    primary_parent_id: Optional[int] = None
    phone: Optional[str] = None
    address: Optional[str] = None
    admin_notes: Optional[str] = None
    is_archived: Optional[bool] = None
    class_ids: Optional[List[int]] = None


class TeacherCreateRequest(BaseModel):
    full_name: str = Field(..., min_length=2)
    phone: str = Field(..., min_length=6)
    email: Optional[str] = None
    specialization: str = Field(..., min_length=2)
    bio: Optional[str] = None
    hire_date: Optional[str] = None
    hourly_rate: float = Field(default=0.0, ge=0)
    username: Optional[str] = None
    password: Optional[str] = None
    telegram_id: Optional[int] = None
    subject_ids: List[int] = Field(default_factory=list)


class TeacherUpdateRequest(BaseModel):
    full_name: Optional[str] = None
    phone: Optional[str] = None
    email: Optional[str] = None
    specialization: Optional[str] = None
    bio: Optional[str] = None
    hourly_rate: Optional[float] = Field(default=None, ge=0)
    is_archived: Optional[bool] = None
    subject_ids: Optional[List[int]] = None


class ClassCreateRequest(BaseModel):
    name: str = Field(..., min_length=2)
    level_id: int
    subject_id: int
    teacher_id: int
    room: str = Field(..., min_length=1)
    capacity: int = Field(default=25, gt=0)
    billing_mode: PaymentPlanType = PaymentPlanType.MONTHLY
    monthly_fee: float = Field(default=0.0, ge=0)
    session_fee: float = Field(default=0.0, ge=0)
    student_ids: List[int] = Field(default_factory=list)
    override_capacity: bool = False


class ClassUpdateRequest(BaseModel):
    name: Optional[str] = None
    level_id: Optional[int] = None
    subject_id: Optional[int] = None
    teacher_id: Optional[int] = None
    room: Optional[str] = None
    capacity: Optional[int] = Field(default=None, gt=0)
    billing_mode: Optional[PaymentPlanType] = None
    monthly_fee: Optional[float] = Field(default=None, ge=0)
    session_fee: Optional[float] = Field(default=None, ge=0)
    is_archived: Optional[bool] = None
    student_ids: Optional[List[int]] = None
    override_capacity: bool = False


class EnrollmentCreateRequest(BaseModel):
    student_id: int
    class_id: int
    billing_override: Optional[str] = Field(default=None)
    discount_percent: float = Field(default=0.0, ge=0, le=100)
    override_capacity: bool = False


class ScheduleCreateRequest(BaseModel):
    class_id: int
    subject_id: Optional[int] = None
    teacher_id: Optional[int] = None
    day_of_week: int = Field(..., ge=0, le=6)
    specific_date: Optional[str] = None
    start_time: str = Field(...)
    end_time: str = Field(...)
    room: Optional[str] = None
    notes: Optional[str] = None


class ScheduleUpdateRequest(BaseModel):
    day_of_week: Optional[int] = Field(default=None, ge=0, le=6)
    specific_date: Optional[str] = None
    start_time: Optional[str] = Field(default=None)
    end_time: Optional[str] = Field(default=None)
    teacher_id: Optional[int] = None
    room: Optional[str] = None
    status: Optional[str] = Field(default=None)
    notes: Optional[str] = None


class AttendanceItemInput(BaseModel):
    student_id: int
    status: AttendanceStatus
    late_minutes: int = Field(default=0, ge=0)
    absence_reason: Optional[str] = None


class AttendanceCreateRequest(BaseModel):
    schedule_id: int
    session_date: str = Field(default_factory=lambda: datetime.date.today().isoformat())
    student_id: Optional[int] = None
    status: Optional[AttendanceStatus] = None
    late_minutes: int = Field(default=0, ge=0)
    absence_reason: Optional[str] = None
    records: Optional[List[AttendanceItemInput]] = None


class AttendanceUpdateRequest(BaseModel):
    status: AttendanceStatus
    late_minutes: int = Field(default=0, ge=0)
    absence_reason: Optional[str] = None


class ExamCreateRequest(BaseModel):
    title: str = Field(..., min_length=2)
    subject_id: int
    class_id: int
    teacher_id: Optional[int] = None
    exam_date: str = Field(default_factory=lambda: datetime.date.today().isoformat())
    max_score: float = Field(default=20.0, gt=0)
    weight: float = Field(default=1.0, gt=0)
    show_ranking: bool = True
    description: Optional[str] = None


class GradeItemInput(BaseModel):
    student_id: int
    score: float = Field(..., ge=0)
    teacher_remarks: Optional[str] = None


class GradeCreateRequest(BaseModel):
    exam_id: int
    student_id: Optional[int] = None
    score: Optional[float] = Field(default=None, ge=0)
    teacher_remarks: Optional[str] = None
    records: Optional[List[GradeItemInput]] = None


class GradeUpdateRequest(BaseModel):
    score: float = Field(..., ge=0)
    teacher_remarks: Optional[str] = None


class PaymentPlanCreateRequest(BaseModel):
    student_id: int
    class_id: Optional[int] = None
    subject_id: Optional[int] = None
    plan_type: str = Field(default="MONTHLY")
    sessions_count: Optional[int] = Field(default=4, gt=0)
    total_amount: float = Field(..., ge=0)
    due_date: str = Field(default_factory=lambda: datetime.date.today().isoformat())
    period_label: Optional[str] = None
    notes: Optional[str] = None


class PaymentCreateRequest(BaseModel):
    student_id: int
    amount: float = Field(..., gt=0)
    payment_method: str = Field(default="CASH")
    description: Optional[str] = None
    payment_plan_id: Optional[int] = None
    # Optional inline plan creation if creating subscription + payment together
    create_plan_type: Optional[str] = Field(default=None)
    create_plan_total_amount: Optional[float] = Field(default=None, ge=0)
    create_plan_due_date: Optional[str] = None
    create_plan_class_id: Optional[int] = None


class MaterialCreateRequest(BaseModel):
    title: str = Field(..., min_length=2)
    description: Optional[str] = None
    material_type: MaterialType = MaterialType.PDF
    subject_id: int
    class_id: Optional[int] = None
    teacher_id: Optional[int] = None
    external_url: Optional[str] = None
    publish_date: Optional[str] = None
    due_date: Optional[str] = None
    max_score: Optional[float] = 20.0


class MaterialUpdateRequest(BaseModel):
    title: Optional[str] = None
    description: Optional[str] = None
    external_url: Optional[str] = None
    due_date: Optional[str] = None
    is_archived: Optional[bool] = None


class NotificationCreateRequest(BaseModel):
    recipient_user_ids: Optional[List[int]] = None
    target_role: Optional[UserRole] = None
    target_class_id: Optional[int] = None
    notification_type: NotificationType = NotificationType.ANNOUNCEMENT
    title_ar: str = Field(..., min_length=2)
    title_en: Optional[str] = None
    body_ar: str = Field(..., min_length=2)
    body_en: Optional[str] = None


class NotificationPreferenceUpdateRequest(BaseModel):
    enable_telegram: Optional[bool] = None
    enable_app: Optional[bool] = None
    notify_absence: Optional[bool] = None
    notify_grades: Optional[bool] = None
    notify_payments: Optional[bool] = None
    notify_schedule: Optional[bool] = None
    notify_materials: Optional[bool] = None
    notify_announcements: Optional[bool] = None


class RestoreBackupRequest(BaseModel):
    backup_id: Optional[int] = None
    filename: Optional[str] = None


# ------------------------------------------------------------------------------
# FastAPI App Initialization, Unified Error Handling & Rate Limiting Middleware
# ------------------------------------------------------------------------------

app = FastAPI(
    title="MOUNIX SCHOOL — Unified Educational Ecosystem API",
    description=(
        "Complete REST API for MOUNIX SCHOOL Private School & Educational Support Center "
        "Management System, integrated with Telegram Bot & SQLite Central Database."
    ),
    version="1.0.0",
    docs_url="/docs",
    redoc_url="/redoc",
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Sliding-window rate limiter per client IP
_rate_limit_buckets: Dict[str, deque] = defaultdict(deque)
RATE_LIMIT_MAX_REQUESTS = int(os.getenv("RATE_LIMIT_MAX_REQUESTS", "240"))
RATE_LIMIT_WINDOW_SECONDS = 60


@app.middleware("http")
async def rate_limit_and_security_headers_middleware(request: Request, call_next: Callable):
    client_ip = request.client.host if request.client else "127.0.0.1"
    now = time.time()
    bucket = _rate_limit_buckets[client_ip]
    while bucket and now - bucket[0] > RATE_LIMIT_WINDOW_SECONDS:
        bucket.popleft()
    if len(bucket) >= RATE_LIMIT_MAX_REQUESTS:
        return JSONResponse(
            status_code=status.HTTP_429_TOO_MANY_REQUESTS,
            content={
                "success": False,
                "error": {
                    "code": "RATE_LIMIT_EXCEEDED",
                    "message": "تم تجاوز الحد المسموح به من الطلبات. يرجى المحاولة بعد قليل. / Too many requests.",
                },
            },
        )
    bucket.append(now)

    response: Response = await call_next(request)
    response.headers["X-Content-Type-Options"] = "nosniff"
    response.headers["X-Frame-Options"] = "DENY"
    response.headers["X-XSS-Protection"] = "1; mode=block"
    return response


@app.exception_handler(HTTPException)
async def custom_http_exception_handler(request: Request, exc: HTTPException):
    return JSONResponse(
        status_code=exc.status_code,
        content={
            "success": False,
            "error": {
                "code": f"HTTP_{exc.status_code}",
                "message": exc.detail,
            },
        },
    )


@app.exception_handler(RequestValidationError)
async def validation_exception_handler(request: Request, exc: RequestValidationError):
    return JSONResponse(
        status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
        content={
            "success": False,
            "error": {
                "code": "VALIDATION_ERROR",
                "message": "بيانات الإدخال غير صالحة / Validation error",
                "details": exc.errors(),
            },
        },
    )


@app.exception_handler(Exception)
async def unhandled_exception_handler(request: Request, exc: Exception):
    logger.error(f"Unhandled API error on {request.url.path}: {exc}\n{traceback.format_exc()}")
    return JSONResponse(
        status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
        content={
            "success": False,
            "error": {
                "code": "INTERNAL_SERVER_ERROR",
                "message": f"حدث خطأ داخلي في الخادم: {str(exc)}",
            },
        },
    )


# ------------------------------------------------------------------------------
# Auth Dependencies & RBAC Guards
# ------------------------------------------------------------------------------

bearer_scheme = HTTPBearer(auto_error=False)


async def get_current_user(
    credentials: Optional[HTTPAuthorizationCredentials] = Security(bearer_scheme),
) -> Dict[str, Any]:
    if credentials is None or not credentials.credentials:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="غير مصرح بالدخول: يرجى تقديم رمز المصادقة (Missing Bearer Token)",
        )
    token = credentials.credentials
    try:
        payload = decode_jwt(token)
    except ValueError as exc:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail=f"رمز الدخول غير صالح أو منتهي الصلاحية ({exc})",
        ) from exc

    if payload.get("type") != "access":
        raise HTTPException(status_code=status.HTTP_401_UNAUTHORIZED, detail="Invalid token type")
    jti = payload.get("jti")
    if jti and jti in REVOKED_TOKEN_JTIS:
        raise HTTPException(status_code=status.HTTP_401_UNAUTHORIZED, detail="تم تسجيل الخروج وإبطال هذا الرمز")

    user_id = int(payload["sub"])
    user = await db.fetchone(
        "SELECT id, username, full_name, role, phone, email, telegram_id, language, theme_mode, is_active, is_deleted, last_login_at, created_at FROM users WHERE id = ?",
        (user_id,),
    )
    if not user or user["is_deleted"] or not user["is_active"]:
        raise HTTPException(status_code=status.HTTP_403_FORBIDDEN, detail="الحساب غير نشط أو محذوف")
    return user


def require_roles(*allowed_roles: UserRole):
    allowed_set = {r.value for r in allowed_roles}

    async def _guard(current_user: Dict[str, Any] = Depends(get_current_user)) -> Dict[str, Any]:
        if current_user["role"] not in allowed_set:
            raise HTTPException(
                status_code=status.HTTP_403_FORBIDDEN,
                detail=f"ليس لديك صلاحية لتنفيذ هذه العملية. الأدوار المسموحة: {', '.join(allowed_set)}",
            )
        return current_user

    return _guard


def paginate_query_results(items: List[Dict[str, Any]], page: int, page_size: int) -> Dict[str, Any]:
    total = len(items)
    start = (page - 1) * page_size
    end = start + page_size
    return {
        "success": True,
        "total": total,
        "page": page,
        "page_size": page_size,
        "total_pages": max(1, (total + page_size - 1) // page_size),
        "items": items[start:end],
    }


# ------------------------------------------------------------------------------
# 10.1 Authentication & Profile Endpoints
# ------------------------------------------------------------------------------

@app.post("/auth/login", tags=["Authentication"])
async def api_login(req: LoginRequest, request: Request):
    user = await db.fetchone(
        "SELECT * FROM users WHERE username = ? AND is_deleted = 0",
        (req.username.strip(),),
    )
    if not user or not verify_password(req.password, user["password_hash"]):
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="اسم المستخدم أو كلمة المرور غير صحيحة / Invalid credentials",
        )
    if not user["is_active"]:
        raise HTTPException(status_code=status.HTTP_403_FORBIDDEN, detail="الحساب معطل من قبل الإدارة")

    access_token = create_access_token(user["id"], user["username"], user["role"])
    refresh_token = create_refresh_token(user["id"], user["username"], user["role"])

    async with db.transaction() as session:
        await session.execute(
            "UPDATE users SET last_login_at = datetime('now') WHERE id = ?",
            (user["id"],),
        )
        await log_audit(
            session, user["id"], user["username"], user["role"],
            "LOGIN", "users", user["id"],
            ip_address=request.client.host if request.client else None,
        )

    user_safe = {k: v for k, v in user.items() if k != "password_hash"}
    return {
        "success": True,
        "access_token": access_token,
        "refresh_token": refresh_token,
        "token_type": "bearer",
        "expires_in": ACCESS_TOKEN_EXPIRE_MINUTES * 60,
        "user": user_safe,
    }


@app.post("/auth/logout", tags=["Authentication"])
async def api_logout(
    credentials: Optional[HTTPAuthorizationCredentials] = Security(bearer_scheme),
    current_user: Dict[str, Any] = Depends(get_current_user),
):
    if credentials and credentials.credentials:
        with contextlib.suppress(Exception):
            payload = decode_jwt(credentials.credentials)
            if payload.get("jti"):
                REVOKED_TOKEN_JTIS.add(payload["jti"])
    await log_audit(
        None, current_user["id"], current_user["username"], current_user["role"],
        "LOGOUT", "users", current_user["id"],
    )
    return {"success": True, "message": "تم تسجيل الخروج بنجاح"}


@app.post("/auth/refresh", tags=["Authentication"])
async def api_refresh_token(req: RefreshTokenRequest):
    try:
        payload = decode_jwt(req.refresh_token)
    except ValueError as exc:
        raise HTTPException(status_code=status.HTTP_401_UNAUTHORIZED, detail=str(exc)) from exc

    if payload.get("type") != "refresh":
        raise HTTPException(status_code=status.HTTP_401_UNAUTHORIZED, detail="Invalid refresh token type")
    jti = payload.get("jti")
    if jti and jti in REVOKED_TOKEN_JTIS:
        raise HTTPException(status_code=status.HTTP_401_UNAUTHORIZED, detail="Refresh token has been revoked")

    user_id = int(payload["sub"])
    user = await db.fetchone(
        "SELECT id, username, role, is_active, is_deleted FROM users WHERE id = ?",
        (user_id,),
    )
    if not user or user["is_deleted"] or not user["is_active"]:
        raise HTTPException(status_code=status.HTTP_403_FORBIDDEN, detail="User inactive")

    # Rotate refresh token
    if jti:
        REVOKED_TOKEN_JTIS.add(jti)
    new_access = create_access_token(user["id"], user["username"], user["role"])
    new_refresh = create_refresh_token(user["id"], user["username"], user["role"])
    return {
        "success": True,
        "access_token": new_access,
        "refresh_token": new_refresh,
        "token_type": "bearer",
    }


@app.get("/me", tags=["Authentication"])
async def api_get_me(current_user: Dict[str, Any] = Depends(get_current_user)):
    role = current_user["role"]
    profile_details: Dict[str, Any] = {}
    if role == UserRole.TEACHER.value:
        profile_details["teacher"] = await db.fetchone(
            "SELECT * FROM teachers WHERE user_id = ?", (current_user["id"],)
        )
    elif role == UserRole.STUDENT.value:
        st = await db.fetchone("SELECT id FROM students WHERE user_id = ?", (current_user["id"],))
        if st:
            profile_details["student_dossier"] = await get_student_full_dossier(st["id"])
    elif role == UserRole.PARENT.value:
        parent = await db.fetchone("SELECT * FROM parents WHERE user_id = ?", (current_user["id"],))
        children = []
        if parent:
            children = await db.fetchall(
                """
                SELECT DISTINCT s.*, l.name_ar AS level_name_ar, l.name_en AS level_name_en
                FROM students s
                LEFT JOIN levels l ON l.id = s.level_id
                LEFT JOIN parent_students ps ON ps.student_id = s.id
                WHERE s.primary_parent_id = ? OR ps.parent_id = ?
                """,
                (parent["id"], parent["id"]),
            )
        profile_details["parent"] = parent
        profile_details["children"] = children

    prefs = await db.fetchone(
        "SELECT * FROM notification_preferences WHERE user_id = ?",
        (current_user["id"],),
    )
    unread_notif = await db.fetchone(
        "SELECT COUNT(*) AS cnt FROM notifications WHERE recipient_user_id = ? AND is_read = 0",
        (current_user["id"],),
    )
    return {
        "success": True,
        "user": current_user,
        "profile": profile_details,
        "notification_preferences": prefs,
        "unread_notifications_count": int(unread_notif["cnt"]) if unread_notif else 0,
    }


@app.put("/me/preferences", tags=["Authentication"])
async def api_update_my_preferences(
    language: Optional[str] = Query(default=None),
    theme_mode: Optional[str] = Query(default=None),
    prefs_req: Optional[NotificationPreferenceUpdateRequest] = Body(default=None),
    current_user: Dict[str, Any] = Depends(get_current_user),
):
    async with db.transaction() as session:
        if language:
            await session.execute(
                "UPDATE users SET language = ?, updated_at = datetime('now') WHERE id = ?",
                (language, current_user["id"]),
            )
        if theme_mode:
            await session.execute(
                "UPDATE users SET theme_mode = ?, updated_at = datetime('now') WHERE id = ?",
                (theme_mode, current_user["id"]),
            )
        if prefs_req:
            await session.execute(
                "INSERT OR IGNORE INTO notification_preferences (user_id) VALUES (?)",
                (current_user["id"],),
            )
            updates = prefs_model_to_dict(req, exclude_none=True)
            for col, val in updates.items():
                await session.execute(
                    f"UPDATE notification_preferences SET {col} = ?, updated_at = datetime('now') WHERE user_id = ?",
                    (1 if val else 0, current_user["id"]),
                )
    return {"success": True, "message": "تم تحديث التفضيلات بنجاح"}


# ------------------------------------------------------------------------------
# 10.2 Dashboard Summary Endpoint
# ------------------------------------------------------------------------------

@app.get("/dashboard", tags=["Dashboard"])
async def api_get_dashboard(current_user: Dict[str, Any] = Depends(get_current_user)):
    role = current_user["role"]
    if role == UserRole.ADMIN.value:
        students_cnt = (await db.fetchone("SELECT COUNT(*) AS c FROM students WHERE is_archived = 0"))["c"]
        teachers_cnt = (await db.fetchone("SELECT COUNT(*) AS c FROM teachers WHERE is_archived = 0"))["c"]
        parents_cnt = (await db.fetchone("SELECT COUNT(*) AS c FROM parents WHERE is_archived = 0"))["c"]
        classes_cnt = (await db.fetchone("SELECT COUNT(*) AS c FROM classes WHERE is_archived = 0"))["c"]
        subjects_cnt = (await db.fetchone("SELECT COUNT(*) AS c FROM subjects WHERE is_archived = 0"))["c"]
        rev_row = await db.fetchone("SELECT COALESCE(SUM(amount), 0) AS total FROM payments WHERE is_voided = 0")
        debt_row = await db.fetchone(
            "SELECT COALESCE(SUM(remaining_amount), 0) AS total FROM payment_plans WHERE status IN ('UNPAID', 'PARTIAL', 'OVERDUE')"
        )
        att_row = await db.fetchone(
            """
            SELECT COUNT(*) AS total,
                   SUM(CASE WHEN status IN ('PRESENT', 'LATE') THEN 1 ELSE 0 END) AS present_cnt
            FROM attendance
            """
        )
        att_rate = (
            round((float(att_row["present_cnt"] or 0) / float(att_row["total"])) * 100.0, 1)
            if att_row and att_row["total"]
            else 100.0
        )
        recent_payments = await db.fetchall(
            """
            SELECT p.*, s.full_name AS student_name, s.student_code
            FROM payments p
            JOIN students s ON s.id = p.student_id
            WHERE p.is_voided = 0
            ORDER BY p.id DESC LIMIT 5
            """
        )
        return {
            "success": True,
            "role": role,
            "stats": {
                "active_students": students_cnt,
                "active_teachers": teachers_cnt,
                "active_parents": parents_cnt,
                "active_classes": classes_cnt,
                "active_subjects": subjects_cnt,
                "collected_revenue": round(float(rev_row["total"]), 2),
                "outstanding_debt": round(float(debt_row["total"]), 2),
                "overall_attendance_rate": att_rate,
            },
            "recent_payments": recent_payments,
        }

    if role == UserRole.TEACHER.value:
        teacher = await db.fetchone("SELECT * FROM teachers WHERE user_id = ?", (current_user["id"],))
        if not teacher:
            return {"success": True, "role": role, "stats": {}}
        tid = teacher["id"]
        classes_cnt = (await db.fetchone("SELECT COUNT(*) AS c FROM classes WHERE teacher_id = ? AND is_archived = 0", (tid,)))["c"]
        students_cnt = (await db.fetchone(
            "SELECT COUNT(DISTINCT e.student_id) AS c FROM enrollments e JOIN classes c ON c.id = e.class_id WHERE c.teacher_id = ? AND e.status = 'ACTIVE'",
            (tid,),
        ))["c"]
        schedules = await db.fetchall(
            """
            SELECT s.*, c.name AS class_name, sub.name_ar AS subject_name_ar
            FROM schedules s
            JOIN classes c ON c.id = s.class_id
            JOIN subjects sub ON sub.id = s.subject_id
            WHERE s.teacher_id = ? AND s.status = 'SCHEDULED'
            ORDER BY s.day_of_week ASC, s.start_time ASC
            """,
            (tid,),
        )
        return {
            "success": True,
            "role": role,
            "stats": {
                "assigned_classes": classes_cnt,
                "total_students": students_cnt,
                "weekly_sessions": len(schedules),
            },
            "schedules": schedules,
        }

    if role == UserRole.STUDENT.value:
        st = await db.fetchone("SELECT id FROM students WHERE user_id = ?", (current_user["id"],))
        dossier = await get_student_full_dossier(st["id"]) if st else None
        return {"success": True, "role": role, "dossier": dossier}

    if role == UserRole.PARENT.value:
        parent = await db.fetchone("SELECT * FROM parents WHERE user_id = ?", (current_user["id"],))
        children_dossiers = []
        if parent:
            c_rows = await db.fetchall(
                """
                SELECT DISTINCT s.id
                FROM students s
                LEFT JOIN parent_students ps ON ps.student_id = s.id
                WHERE s.primary_parent_id = ? OR ps.parent_id = ?
                """,
                (parent["id"], parent["id"]),
            )
            for cr in c_rows:
                d = await get_student_full_dossier(cr["id"])
                if d:
                    children_dossiers.append(d)
        return {"success": True, "role": role, "children": children_dossiers}

    return {"success": True, "role": role}


# ------------------------------------------------------------------------------
# 10.3 Users Management Endpoints (/users)
# ------------------------------------------------------------------------------

@app.get("/users", tags=["Users"])
async def api_list_users(
    role: Optional[UserRole] = None,
    search: Optional[str] = None,
    page: int = Query(default=1, ge=1),
    page_size: int = Query(default=25, ge=1, le=200),
    current_user: Dict[str, Any] = Depends(require_roles(UserRole.ADMIN)),
):
    sql = "SELECT id, username, full_name, role, phone, email, telegram_id, language, theme_mode, is_active, last_login_at, created_at FROM users WHERE is_deleted = 0"
    params: List[Any] = []
    if role:
        sql += " AND role = ?"
        params.append(role.value)
    if search:
        sql += " AND (username LIKE ? OR full_name LIKE ? OR phone LIKE ?)"
        q = f"%{search.strip()}%"
        params.extend([q, q, q])
    sql += " ORDER BY id DESC"
    rows = await db.fetchall(sql, params)
    return paginate_query_results(rows, page, page_size)


@app.post("/users", status_code=status.HTTP_201_CREATED, tags=["Users"])
async def api_create_user(
    req: UserCreateRequest,
    current_user: Dict[str, Any] = Depends(require_roles(UserRole.ADMIN)),
):
    async with db.transaction() as session:
        existing = await session.fetchone("SELECT id FROM users WHERE username = ?", (req.username.strip(),))
        if existing:
            raise HTTPException(status_code=400, detail="اسم المستخدم مستخدم بالفعل / Username already exists")
        pwd_hash = hash_password(req.password)
        uid = await session.insert(
            """
            INSERT INTO users (username, password_hash, full_name, role, phone, email, telegram_id, language, theme_mode, is_active)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                req.username.strip(),
                pwd_hash,
                req.full_name.strip(),
                req.role.value,
                req.phone,
                req.email,
                req.telegram_id,
                req.language,
                req.theme_mode,
                1 if req.is_active else 0,
            ),
        )
        await session.execute("INSERT OR IGNORE INTO notification_preferences (user_id) VALUES (?)", (uid,))
        await log_audit(
            session, current_user["id"], current_user["username"], current_user["role"],
            "CREATE_USER", "users", uid, {"username": req.username, "role": req.role.value},
        )
        created = await session.fetchone(
            "SELECT id, username, full_name, role, phone, email, telegram_id, language, theme_mode, is_active, created_at FROM users WHERE id = ?",
            (uid,),
        )
    return {"success": True, "user": created}


@app.put("/users/{id}", tags=["Users"])
async def api_update_user(
    id: int,
    req: UserUpdateRequest,
    current_user: Dict[str, Any] = Depends(require_roles(UserRole.ADMIN)),
):
    async with db.transaction() as session:
        target = await session.fetchone("SELECT * FROM users WHERE id = ? AND is_deleted = 0", (id,))
        if not target:
            raise HTTPException(status_code=404, detail="المستخدم غير موجود")
        updates = model_to_dict(req, exclude_none=True)
        if "password" in updates:
            updates["password_hash"] = hash_password(updates.pop("password"))
        if "role" in updates and isinstance(updates["role"], UserRole):
            updates["role"] = updates["role"].value
        if "is_active" in updates:
            updates["is_active"] = 1 if updates["is_active"] else 0

        for col, val in updates.items():
            await session.execute(f"UPDATE users SET {col} = ?, updated_at = datetime('now') WHERE id = ?", (val, id))

        await log_audit(
            session, current_user["id"], current_user["username"], current_user["role"],
            "UPDATE_USER", "users", id, {"updated_fields": list(updates.keys())},
        )
        updated = await session.fetchone(
            "SELECT id, username, full_name, role, phone, email, telegram_id, language, theme_mode, is_active, updated_at FROM users WHERE id = ?",
            (id,),
        )
    return {"success": True, "user": updated}


@app.delete("/users/{id}", tags=["Users"])
async def api_delete_user(
    id: int,
    current_user: Dict[str, Any] = Depends(require_roles(UserRole.ADMIN)),
):
    if id == current_user["id"]:
        raise HTTPException(status_code=400, detail="لا يمكنك حذف حسابك الشخصي أثناء الجلسة النشطة")
    async with db.transaction() as session:
        await session.execute(
            "UPDATE users SET is_deleted = 1, is_active = 0, updated_at = datetime('now') WHERE id = ?",
            (id,),
        )
        await log_audit(
            session, current_user["id"], current_user["username"], current_user["role"],
            "DELETE_USER", "users", id,
        )
    return {"success": True, "message": "تم أرشفة/إلغاء تفعيل المستخدم بنجاح"}


# ------------------------------------------------------------------------------
# 10.4 Levels & Subjects Endpoints (/levels, /subjects)
# ------------------------------------------------------------------------------

@app.get("/levels", tags=["Levels & Subjects"])
async def api_list_levels(
    include_archived: bool = False,
    current_user: Dict[str, Any] = Depends(get_current_user),
):
    sql = "SELECT * FROM levels"
    if not include_archived:
        sql += " WHERE is_archived = 0"
    sql += " ORDER BY sort_order ASC, id ASC"
    rows = await db.fetchall(sql)
    return {"success": True, "items": rows}


@app.post("/levels", status_code=status.HTTP_201_CREATED, tags=["Levels & Subjects"])
async def api_create_level(
    req: LevelCreateRequest,
    current_user: Dict[str, Any] = Depends(require_roles(UserRole.ADMIN)),
):
    async with db.transaction() as session:
        code = req.code or await generate_unique_code(session, "LVL", "levels", "code")
        lid = await session.insert(
            "INSERT INTO levels (code, name_ar, name_en, sort_order, description) VALUES (?, ?, ?, ?, ?)",
            (code, req.name_ar.strip(), req.name_en.strip(), req.sort_order, req.description),
        )
        await log_audit(
            session, current_user["id"], current_user["username"], current_user["role"],
            "CREATE_LEVEL", "levels", lid, {"code": code, "name_ar": req.name_ar},
        )
        row = await session.fetchone("SELECT * FROM levels WHERE id = ?", (lid,))
    return {"success": True, "level": row}


@app.put("/levels/{id}", tags=["Levels & Subjects"])
async def api_update_level(
    id: int,
    req: LevelUpdateRequest,
    current_user: Dict[str, Any] = Depends(require_roles(UserRole.ADMIN)),
):
    async with db.transaction() as session:
        updates = model_to_dict(req, exclude_none=True)
        if "is_archived" in updates:
            updates["is_archived"] = 1 if updates["is_archived"] else 0
        for col, val in updates.items():
            await session.execute(f"UPDATE levels SET {col} = ?, updated_at = datetime('now') WHERE id = ?", (val, id))
        row = await session.fetchone("SELECT * FROM levels WHERE id = ?", (id,))
        if not row:
            raise HTTPException(status_code=404, detail="المستوى غير موجود")
    return {"success": True, "level": row}


@app.get("/subjects", tags=["Levels & Subjects"])
async def api_list_subjects(
    level_id: Optional[int] = None,
    include_archived: bool = False,
    current_user: Dict[str, Any] = Depends(get_current_user),
):
    sql = """
        SELECT s.*, l.name_ar AS level_name_ar, l.name_en AS level_name_en
        FROM subjects s
        LEFT JOIN levels l ON l.id = s.level_id
        WHERE 1=1
    """
    params: List[Any] = []
    if not include_archived:
        sql += " AND s.is_archived = 0"
    if level_id is not None:
        sql += " AND s.level_id = ?"
        params.append(level_id)
    sql += " ORDER BY s.id DESC"
    rows = await db.fetchall(sql, params)
    for r in rows:
        r["teachers"] = await db.fetchall(
            """
            SELECT t.id, t.teacher_code, t.full_name, t.specialization
            FROM teacher_subjects ts
            JOIN teachers t ON t.id = ts.teacher_id
            WHERE ts.subject_id = ? AND t.is_archived = 0
            """,
            (r["id"],),
        )
    return {"success": True, "items": rows}


@app.post("/subjects", status_code=status.HTTP_201_CREATED, tags=["Levels & Subjects"])
async def api_create_subject(
    req: SubjectCreateRequest,
    current_user: Dict[str, Any] = Depends(require_roles(UserRole.ADMIN)),
):
    async with db.transaction() as session:
        code = req.code or await generate_unique_code(session, "SUB", "subjects", "code")
        sid = await session.insert(
            """
            INSERT INTO subjects (code, name_ar, name_en, description, level_id, default_monthly_fee, default_session_fee)
            VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
            (
                code, req.name_ar.strip(), req.name_en.strip(), req.description,
                req.level_id, req.default_monthly_fee, req.default_session_fee,
            ),
        )
        for tid in req.teacher_ids:
            await session.execute(
                "INSERT OR IGNORE INTO teacher_subjects (teacher_id, subject_id) VALUES (?, ?)",
                (tid, sid),
            )
        await log_audit(
            session, current_user["id"], current_user["username"], current_user["role"],
            "CREATE_SUBJECT", "subjects", sid, {"code": code, "name_ar": req.name_ar},
        )
        row = await session.fetchone("SELECT * FROM subjects WHERE id = ?", (sid,))
    return {"success": True, "subject": row}


@app.put("/subjects/{id}", tags=["Levels & Subjects"])
async def api_update_subject(
    id: int,
    req: SubjectUpdateRequest,
    current_user: Dict[str, Any] = Depends(require_roles(UserRole.ADMIN)),
):
    async with db.transaction() as session:
        target = await session.fetchone("SELECT * FROM subjects WHERE id = ?", (id,))
        if not target:
            raise HTTPException(status_code=404, detail="المادة غير موجودة")
        updates = model_to_dict(req, exclude_none=True)
        teacher_ids = updates.pop("teacher_ids", None)
        if "is_archived" in updates:
            updates["is_archived"] = 1 if updates["is_archived"] else 0
        for col, val in updates.items():
            await session.execute(f"UPDATE subjects SET {col} = ?, updated_at = datetime('now') WHERE id = ?", (val, id))
        if teacher_ids is not None:
            await session.execute("DELETE FROM teacher_subjects WHERE subject_id = ?", (id,))
            for tid in teacher_ids:
                await session.execute(
                    "INSERT OR IGNORE INTO teacher_subjects (teacher_id, subject_id) VALUES (?, ?)",
                    (tid, id),
                )
        await log_audit(
            session, current_user["id"], current_user["username"], current_user["role"],
            "UPDATE_SUBJECT", "subjects", id,
        )
        row = await session.fetchone("SELECT * FROM subjects WHERE id = ?", (id,))
    return {"success": True, "subject": row}


@app.delete("/subjects/{id}", tags=["Levels & Subjects"])
async def api_delete_subject(
    id: int,
    current_user: Dict[str, Any] = Depends(require_roles(UserRole.ADMIN)),
):
    async with db.transaction() as session:
        await session.execute("UPDATE subjects SET is_archived = 1, updated_at = datetime('now') WHERE id = ?", (id,))
        await log_audit(
            session, current_user["id"], current_user["username"], current_user["role"],
            "ARCHIVE_SUBJECT", "subjects", id,
        )
    return {"success": True, "message": "تم أرشفة المادة بنجاح"}


# ------------------------------------------------------------------------------
# 10.5 Parents, Students, Teachers, Classes & Schedules Endpoints
# ------------------------------------------------------------------------------

@app.get("/parents", tags=["Parents"])
async def api_list_parents(
    search: Optional[str] = None,
    include_archived: bool = False,
    page: int = Query(default=1, ge=1),
    page_size: int = Query(default=25, ge=1, le=200),
    current_user: Dict[str, Any] = Depends(require_roles(UserRole.ADMIN, UserRole.TEACHER)),
):
    sql = "SELECT p.*, u.username, u.telegram_id FROM parents p LEFT JOIN users u ON u.id = p.user_id WHERE 1=1"
    params: List[Any] = []
    if not include_archived:
        sql += " AND p.is_archived = 0"
    if search:
        sql += " AND (p.full_name LIKE ? OR p.parent_code LIKE ? OR p.phone LIKE ?)"
        q = f"%{search.strip()}%"
        params.extend([q, q, q])
    sql += " ORDER BY p.id DESC"
    rows = await db.fetchall(sql, params)
    for r in rows:
        r["children"] = await db.fetchall(
            """
            SELECT DISTINCT s.id, s.student_code, s.full_name, s.level_id
            FROM students s
            LEFT JOIN parent_students ps ON ps.student_id = s.id
            WHERE s.primary_parent_id = ? OR ps.parent_id = ?
            """,
            (r["id"], r["id"]),
        )
    return paginate_query_results(rows, page, page_size)


@app.post("/parents", status_code=status.HTTP_201_CREATED, tags=["Parents"])
async def api_create_parent(
    req: ParentCreateRequest,
    current_user: Dict[str, Any] = Depends(require_roles(UserRole.ADMIN)),
):
    async with db.transaction() as session:
        parent_code = await generate_unique_code(session, "PAR", "parents", "parent_code")
        username = (req.username or parent_code.lower().replace("-", "_")).strip()
        raw_password = req.password or "Parent@12345"
        pwd_hash = hash_password(raw_password)

        uid = await session.insert(
            """
            INSERT INTO users (username, password_hash, full_name, role, phone, email, telegram_id, language, is_active)
            VALUES (?, ?, ?, 'PARENT', ?, ?, ?, 'ar', 1)
            """,
            (username, pwd_hash, req.full_name.strip(), req.phone, req.email, req.telegram_id),
        )
        await session.execute("INSERT OR IGNORE INTO notification_preferences (user_id) VALUES (?)", (uid,))

        pid = await session.insert(
            """
            INSERT INTO parents (user_id, parent_code, full_name, phone, secondary_phone, email, address, occupation, notes)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                uid, parent_code, req.full_name.strip(), req.phone,
                req.secondary_phone, req.email, req.address, req.occupation, req.notes,
            ),
        )

        for sid in req.student_ids:
            await session.execute(
                "INSERT OR IGNORE INTO parent_students (parent_id, student_id, relationship, is_primary) VALUES (?, ?, 'Guardian', 1)",
                (pid, sid),
            )
            await session.execute(
                "UPDATE students SET primary_parent_id = COALESCE(primary_parent_id, ?) WHERE id = ?",
                (pid, sid),
            )

        await log_audit(
            session, current_user["id"], current_user["username"], current_user["role"],
            "CREATE_PARENT", "parents", pid, {"parent_code": parent_code, "full_name": req.full_name},
        )
        created = await session.fetchone("SELECT * FROM parents WHERE id = ?", (pid,))
    return {"success": True, "parent": created, "credentials": {"username": username, "default_password": raw_password}}


@app.put("/parents/{id}", tags=["Parents"])
async def api_update_parent(
    id: int,
    req: ParentUpdateRequest,
    current_user: Dict[str, Any] = Depends(require_roles(UserRole.ADMIN)),
):
    async with db.transaction() as session:
        target = await session.fetchone("SELECT * FROM parents WHERE id = ?", (id,))
        if not target:
            raise HTTPException(status_code=404, detail="ولي الأمر غير موجود")
        updates = model_to_dict(req, exclude_none=True)
        student_ids = updates.pop("student_ids", None)
        if "is_archived" in updates:
            updates["is_archived"] = 1 if updates["is_archived"] else 0
        for col, val in updates.items():
            await session.execute(f"UPDATE parents SET {col} = ?, updated_at = datetime('now') WHERE id = ?", (val, id))
        if target.get("user_id") and ("full_name" in updates or "phone" in updates):
            await session.execute(
                "UPDATE users SET full_name = COALESCE(?, full_name), phone = COALESCE(?, phone), updated_at = datetime('now') WHERE id = ?",
                (updates.get("full_name"), updates.get("phone"), target["user_id"]),
            )
        if student_ids is not None:
            await session.execute("DELETE FROM parent_students WHERE parent_id = ?", (id,))
            for sid in student_ids:
                await session.execute(
                    "INSERT OR IGNORE INTO parent_students (parent_id, student_id, relationship, is_primary) VALUES (?, ?, 'Guardian', 1)",
                    (id, sid),
                )
        await log_audit(
            session, current_user["id"], current_user["username"], current_user["role"],
            "UPDATE_PARENT", "parents", id,
        )
        updated = await session.fetchone("SELECT * FROM parents WHERE id = ?", (id,))
    return {"success": True, "parent": updated}


# ------------------------------------------------------------------------------
# Students Endpoints (/students)
# ------------------------------------------------------------------------------

@app.get("/students", tags=["Students"])
async def api_list_students(
    search: Optional[str] = None,
    level_id: Optional[int] = None,
    class_id: Optional[int] = None,
    include_archived: bool = False,
    sort_by: str = Query(default="id"),
    sort_dir: str = Query(default="DESC"),
    page: int = Query(default=1, ge=1),
    page_size: int = Query(default=25, ge=1, le=200),
    current_user: Dict[str, Any] = Depends(get_current_user),
):
    role = current_user["role"]
    sql = """
        SELECT DISTINCT s.*, l.name_ar AS level_name_ar, l.name_en AS level_name_en,
               p.full_name AS parent_name, p.phone AS parent_phone
        FROM students s
        LEFT JOIN levels l ON l.id = s.level_id
        LEFT JOIN parents p ON p.id = s.primary_parent_id
        LEFT JOIN enrollments e ON e.student_id = s.id
        LEFT JOIN classes c ON c.id = e.class_id
        LEFT JOIN parent_students ps ON ps.student_id = s.id
        WHERE 1=1
    """
    params: List[Any] = []

    if role == UserRole.STUDENT.value:
        sql += " AND s.user_id = ?"
        params.append(current_user["id"])
    elif role == UserRole.PARENT.value:
        parent = await db.fetchone("SELECT id FROM parents WHERE user_id = ?", (current_user["id"],))
        if not parent:
            return paginate_query_results([], page, page_size)
        sql += " AND (s.primary_parent_id = ? OR ps.parent_id = ?)"
        params.extend([parent["id"], parent["id"]])
    elif role == UserRole.TEACHER.value:
        teacher = await db.fetchone("SELECT id FROM teachers WHERE user_id = ?", (current_user["id"],))
        if not teacher:
            return paginate_query_results([], page, page_size)
        sql += " AND c.teacher_id = ?"
        params.append(teacher["id"])

    if not include_archived:
        sql += " AND s.is_archived = 0"
    if level_id is not None:
        sql += " AND s.level_id = ?"
        params.append(level_id)
    if class_id is not None:
        sql += " AND e.class_id = ? AND e.status = 'ACTIVE'"
        params.append(class_id)
    if search:
        sql += " AND (s.full_name LIKE ? OR s.student_code LIKE ? OR s.phone LIKE ?)"
        q = f"%{search.strip()}%"
        params.extend([q, q, q])

    sql += f" ORDER BY s.{sort_by} {sort_dir}"
    rows = await db.fetchall(sql, params)
    return paginate_query_results(rows, page, page_size)


@app.post("/students", status_code=status.HTTP_201_CREATED, tags=["Students"])
async def api_create_student(
    req: StudentCreateRequest,
    current_user: Dict[str, Any] = Depends(require_roles(UserRole.ADMIN)),
):
    async with db.transaction() as session:
        student_code = await generate_unique_code(session, "STU", "students", "student_code")
        username = (req.username or student_code.lower().replace("-", "_")).strip()
        raw_password = req.password or "Student@12345"
        pwd_hash = hash_password(raw_password)

        uid = await session.insert(
            """
            INSERT INTO users (username, password_hash, full_name, role, phone, telegram_id, language, is_active)
            VALUES (?, ?, ?, 'STUDENT', ?, ?, 'ar', 1)
            """,
            (username, pwd_hash, req.full_name.strip(), req.phone, req.telegram_id),
        )
        await session.execute("INSERT OR IGNORE INTO notification_preferences (user_id) VALUES (?)", (uid,))

        joined = req.joined_date or datetime.date.today().isoformat()
        sid = await session.insert(
            """
            INSERT INTO students (
                user_id, student_code, full_name, birth_date, gender, level_id,
                primary_parent_id, phone, address, joined_date, admin_notes
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                uid, student_code, req.full_name.strip(), req.birth_date, req.gender,
                req.level_id, req.primary_parent_id, req.phone, req.address, joined, req.admin_notes,
            ),
        )

        if req.primary_parent_id:
            await session.execute(
                "INSERT OR IGNORE INTO parent_students (parent_id, student_id, relationship, is_primary) VALUES (?, ?, 'Guardian', 1)",
                (req.primary_parent_id, sid),
            )

        for cid in req.class_ids:
            cl = await session.fetchone("SELECT id, name, capacity FROM classes WHERE id = ?", (cid,))
            if cl:
                cnt_row = await session.fetchone(
                    "SELECT COUNT(*) AS cnt FROM enrollments WHERE class_id = ? AND status = 'ACTIVE'",
                    (cid,),
                )
                if int(cnt_row["cnt"]) >= int(cl["capacity"]):
                    raise HTTPException(
                        status_code=400,
                        detail=f"القسم '{cl['name']}' وصل إلى طاقته الاستيعابية القصوى ({cl['capacity']})",
                    )
                await session.execute(
                    "INSERT OR IGNORE INTO enrollments (student_id, class_id, status) VALUES (?, ?, 'ACTIVE')",
                    (sid, cid),
                )

        await log_audit(
            session, current_user["id"], current_user["username"], current_user["role"],
            "CREATE_STUDENT", "students", sid, {"student_code": student_code, "full_name": req.full_name},
        )
        created = await session.fetchone("SELECT * FROM students WHERE id = ?", (sid,))
    return {
        "success": True,
        "student": created,
        "credentials": {"username": username, "default_password": raw_password},
    }


@app.get("/students/{id}", tags=["Students"])
async def api_get_student(
    id: int,
    current_user: Dict[str, Any] = Depends(get_current_user),
):
    if not await can_user_access_student(current_user, id):
        raise HTTPException(status_code=403, detail="غير مصرح لك بالاطلاع على ملف هذا التلميذ")
    dossier = await get_student_full_dossier(id)
    if not dossier:
        raise HTTPException(status_code=404, detail="التلميذ غير موجود")
    return {"success": True, "dossier": dossier}


@app.put("/students/{id}", tags=["Students"])
async def api_update_student(
    id: int,
    req: StudentUpdateRequest,
    current_user: Dict[str, Any] = Depends(require_roles(UserRole.ADMIN)),
):
    async with db.transaction() as session:
        target = await session.fetchone("SELECT * FROM students WHERE id = ?", (id,))
        if not target:
            raise HTTPException(status_code=404, detail="التلميذ غير موجود")
        updates = model_to_dict(req, exclude_none=True)
        class_ids = updates.pop("class_ids", None)
        if "is_archived" in updates:
            updates["is_archived"] = 1 if updates["is_archived"] else 0

        for col, val in updates.items():
            await session.execute(f"UPDATE students SET {col} = ?, updated_at = datetime('now') WHERE id = ?", (val, id))

        if target.get("user_id") and ("full_name" in updates or "phone" in updates):
            await session.execute(
                "UPDATE users SET full_name = COALESCE(?, full_name), phone = COALESCE(?, phone), updated_at = datetime('now') WHERE id = ?",
                (updates.get("full_name"), updates.get("phone"), target["user_id"]),
            )

        if "primary_parent_id" in updates and updates["primary_parent_id"]:
            await session.execute(
                "INSERT OR IGNORE INTO parent_students (parent_id, student_id, relationship, is_primary) VALUES (?, ?, 'Guardian', 1)",
                (updates["primary_parent_id"], id),
            )

        if class_ids is not None:
            await session.execute("DELETE FROM enrollments WHERE student_id = ?", (id,))
            for cid in class_ids:
                await session.execute(
                    "INSERT OR IGNORE INTO enrollments (student_id, class_id, status) VALUES (?, ?, 'ACTIVE')",
                    (id, cid),
                )

        await log_audit(
            session, current_user["id"], current_user["username"], current_user["role"],
            "UPDATE_STUDENT", "students", id, {"updated_fields": list(updates.keys())},
        )
        updated = await session.fetchone("SELECT * FROM students WHERE id = ?", (id,))
    return {"success": True, "student": updated}


@app.delete("/students/{id}", tags=["Students"])
async def api_delete_student(
    id: int,
    current_user: Dict[str, Any] = Depends(require_roles(UserRole.ADMIN)),
):
    async with db.transaction() as session:
        target = await session.fetchone("SELECT * FROM students WHERE id = ?", (id,))
        if not target:
            raise HTTPException(status_code=404, detail="التلميذ غير موجود")
        await session.execute("UPDATE students SET is_archived = 1, updated_at = datetime('now') WHERE id = ?", (id,))
        if target.get("user_id"):
            await session.execute("UPDATE users SET is_active = 0, updated_at = datetime('now') WHERE id = ?", (target["user_id"],))
        await log_audit(
            session, current_user["id"], current_user["username"], current_user["role"],
            "ARCHIVE_STUDENT", "students", id,
        )
    return {"success": True, "message": "تم أرشفة التلميذ بنجاح"}


# ------------------------------------------------------------------------------
# Teachers Endpoints (/teachers)
# ------------------------------------------------------------------------------

@app.get("/teachers", tags=["Teachers"])
async def api_list_teachers(
    search: Optional[str] = None,
    include_archived: bool = False,
    page: int = Query(default=1, ge=1),
    page_size: int = Query(default=25, ge=1, le=200),
    current_user: Dict[str, Any] = Depends(get_current_user),
):
    sql = "SELECT t.*, u.username, u.telegram_id FROM teachers t LEFT JOIN users u ON u.id = t.user_id WHERE 1=1"
    params: List[Any] = []
    if not include_archived:
        sql += " AND t.is_archived = 0"
    if search:
        sql += " AND (t.full_name LIKE ? OR t.teacher_code LIKE ? OR t.specialization LIKE ? OR t.phone LIKE ?)"
        q = f"%{search.strip()}%"
        params.extend([q, q, q, q])
    sql += " ORDER BY t.id DESC"
    rows = await db.fetchall(sql, params)
    for r in rows:
        r["subjects"] = await db.fetchall(
            """
            SELECT s.id, s.code, s.name_ar, s.name_en
            FROM teacher_subjects ts
            JOIN subjects s ON s.id = ts.subject_id
            WHERE ts.teacher_id = ?
            """,
            (r["id"],),
        )
        r["classes"] = await db.fetchall(
            "SELECT id, code, name, room FROM classes WHERE teacher_id = ? AND is_archived = 0",
            (r["id"],),
        )
    return paginate_query_results(rows, page, page_size)


@app.post("/teachers", status_code=status.HTTP_201_CREATED, tags=["Teachers"])
async def api_create_teacher(
    req: TeacherCreateRequest,
    current_user: Dict[str, Any] = Depends(require_roles(UserRole.ADMIN)),
):
    async with db.transaction() as session:
        teacher_code = await generate_unique_code(session, "TCH", "teachers", "teacher_code")
        username = (req.username or teacher_code.lower().replace("-", "_")).strip()
        raw_password = req.password or "Teacher@12345"
        pwd_hash = hash_password(raw_password)

        uid = await session.insert(
            """
            INSERT INTO users (username, password_hash, full_name, role, phone, email, telegram_id, language, is_active)
            VALUES (?, ?, ?, 'TEACHER', ?, ?, ?, 'ar', 1)
            """,
            (username, pwd_hash, req.full_name.strip(), req.phone, req.email, req.telegram_id),
        )
        await session.execute("INSERT OR IGNORE INTO notification_preferences (user_id) VALUES (?)", (uid,))

        hire_date = req.hire_date or datetime.date.today().isoformat()
        tid = await session.insert(
            """
            INSERT INTO teachers (user_id, teacher_code, full_name, phone, email, specialization, bio, hire_date, hourly_rate)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                uid, teacher_code, req.full_name.strip(), req.phone, req.email,
                req.specialization.strip(), req.bio, hire_date, req.hourly_rate,
            ),
        )

        for sid in req.subject_ids:
            await session.execute(
                "INSERT OR IGNORE INTO teacher_subjects (teacher_id, subject_id) VALUES (?, ?)",
                (tid, sid),
            )

        await log_audit(
            session, current_user["id"], current_user["username"], current_user["role"],
            "CREATE_TEACHER", "teachers", tid, {"teacher_code": teacher_code, "full_name": req.full_name},
        )
        created = await session.fetchone("SELECT * FROM teachers WHERE id = ?", (tid,))
    return {
        "success": True,
        "teacher": created,
        "credentials": {"username": username, "default_password": raw_password},
    }


@app.get("/teachers/{id}", tags=["Teachers"])
async def api_get_teacher(
    id: int,
    current_user: Dict[str, Any] = Depends(get_current_user),
):
    teacher = await db.fetchone(
        "SELECT t.*, u.username, u.telegram_id FROM teachers t LEFT JOIN users u ON u.id = t.user_id WHERE t.id = ?",
        (id,),
    )
    if not teacher:
        raise HTTPException(status_code=404, detail="الأستاذ غير موجود")

    subjects = await db.fetchall(
        "SELECT s.* FROM teacher_subjects ts JOIN subjects s ON s.id = ts.subject_id WHERE ts.teacher_id = ?",
        (id,),
    )
    classes = await db.fetchall(
        "SELECT c.*, l.name_ar AS level_name_ar, s.name_ar AS subject_name_ar FROM classes c JOIN levels l ON l.id = c.level_id JOIN subjects s ON s.id = c.subject_id WHERE c.teacher_id = ?",
        (id,),
    )
    schedules = await db.fetchall(
        "SELECT sch.*, c.name AS class_name, s.name_ar AS subject_name_ar FROM schedules sch JOIN classes c ON c.id = sch.class_id JOIN subjects s ON s.id = sch.subject_id WHERE sch.teacher_id = ? AND sch.status = 'SCHEDULED' ORDER BY sch.day_of_week ASC, sch.start_time ASC",
        (id,),
    )
    exams_cnt = (await db.fetchone("SELECT COUNT(*) AS c FROM exams WHERE teacher_id = ?", (id,)))["c"]
    materials_cnt = (await db.fetchone("SELECT COUNT(*) AS c FROM materials WHERE teacher_id = ?", (id,)))["c"]
    audit_history = await db.fetchall(
        "SELECT * FROM audit_logs WHERE entity_type = 'teachers' AND entity_id = ? ORDER BY id DESC LIMIT 20",
        (str(id),),
    )

    return {
        "success": True,
        "teacher": teacher,
        "subjects": subjects,
        "classes": classes,
        "schedules": schedules,
        "activity_stats": {
            "assigned_classes_count": len(classes),
            "weekly_sessions_count": len(schedules),
            "exams_created_count": exams_cnt,
            "materials_uploaded_count": materials_cnt,
        },
        "modification_history": audit_history,
    }


@app.put("/teachers/{id}", tags=["Teachers"])
async def api_update_teacher(
    id: int,
    req: TeacherUpdateRequest,
    current_user: Dict[str, Any] = Depends(require_roles(UserRole.ADMIN)),
):
    async with db.transaction() as session:
        target = await session.fetchone("SELECT * FROM teachers WHERE id = ?", (id,))
        if not target:
            raise HTTPException(status_code=404, detail="الأستاذ غير موجود")
        updates = model_to_dict(req, exclude_none=True)
        subject_ids = updates.pop("subject_ids", None)
        if "is_archived" in updates:
            updates["is_archived"] = 1 if updates["is_archived"] else 0

        for col, val in updates.items():
            await session.execute(f"UPDATE teachers SET {col} = ?, updated_at = datetime('now') WHERE id = ?", (val, id))

        if target.get("user_id") and ("full_name" in updates or "phone" in updates):
            await session.execute(
                "UPDATE users SET full_name = COALESCE(?, full_name), phone = COALESCE(?, phone), updated_at = datetime('now') WHERE id = ?",
                (updates.get("full_name"), updates.get("phone"), target["user_id"]),
            )

        if subject_ids is not None:
            await session.execute("DELETE FROM teacher_subjects WHERE teacher_id = ?", (id,))
            for sid in subject_ids:
                await session.execute(
                    "INSERT OR IGNORE INTO teacher_subjects (teacher_id, subject_id) VALUES (?, ?)",
                    (id, sid),
                )

        await log_audit(
            session, current_user["id"], current_user["username"], current_user["role"],
            "UPDATE_TEACHER", "teachers", id, {"updated_fields": list(updates.keys())},
        )
        updated = await session.fetchone("SELECT * FROM teachers WHERE id = ?", (id,))
    return {"success": True, "teacher": updated}


@app.delete("/teachers/{id}", tags=["Teachers"])
async def api_delete_teacher(
    id: int,
    current_user: Dict[str, Any] = Depends(require_roles(UserRole.ADMIN)),
):
    async with db.transaction() as session:
        target = await session.fetchone("SELECT * FROM teachers WHERE id = ?", (id,))
        if not target:
            raise HTTPException(status_code=404, detail="الأستاذ غير موجود")
        await session.execute("UPDATE teachers SET is_archived = 1, updated_at = datetime('now') WHERE id = ?", (id,))
        if target.get("user_id"):
            await session.execute("UPDATE users SET is_active = 0, updated_at = datetime('now') WHERE id = ?", (target["user_id"],))
        await log_audit(
            session, current_user["id"], current_user["username"], current_user["role"],
            "ARCHIVE_TEACHER", "teachers", id,
        )
    return {"success": True, "message": "تم أرشفة الأستاذ بنجاح"}


# ------------------------------------------------------------------------------
# Classes & Enrollments Endpoints (/classes)
# ------------------------------------------------------------------------------

@app.get("/classes", tags=["Classes"])
async def api_list_classes(
    level_id: Optional[int] = None,
    subject_id: Optional[int] = None,
    teacher_id: Optional[int] = None,
    include_archived: bool = False,
    current_user: Dict[str, Any] = Depends(get_current_user),
):
    sql = """
        SELECT c.*, l.name_ar AS level_name_ar, l.name_en AS level_name_en,
               s.name_ar AS subject_name_ar, s.name_en AS subject_name_en,
               t.full_name AS teacher_name,
               (SELECT COUNT(*) FROM enrollments e WHERE e.class_id = c.id AND e.status = 'ACTIVE') AS enrolled_count
        FROM classes c
        JOIN levels l ON l.id = c.level_id
        JOIN subjects s ON s.id = c.subject_id
        JOIN teachers t ON t.id = c.teacher_id
        WHERE 1=1
    """
    params: List[Any] = []
    if current_user["role"] == UserRole.TEACHER.value:
        t_row = await db.fetchone("SELECT id FROM teachers WHERE user_id = ?", (current_user["id"],))
        if t_row:
            sql += " AND c.teacher_id = ?"
            params.append(t_row["id"])
    elif current_user["role"] == UserRole.STUDENT.value:
        st_row = await db.fetchone("SELECT id FROM students WHERE user_id = ?", (current_user["id"],))
        if st_row:
            sql += " AND c.id IN (SELECT class_id FROM enrollments WHERE student_id = ? AND status = 'ACTIVE')"
            params.append(st_row["id"])

    if not include_archived:
        sql += " AND c.is_archived = 0"
    if level_id is not None:
        sql += " AND c.level_id = ?"
        params.append(level_id)
    if subject_id is not None:
        sql += " AND c.subject_id = ?"
        params.append(subject_id)
    if teacher_id is not None:
        sql += " AND c.teacher_id = ?"
        params.append(teacher_id)

    sql += " ORDER BY c.id DESC"
    rows = await db.fetchall(sql, params)
    return {"success": True, "items": rows}


@app.post("/classes", status_code=status.HTTP_201_CREATED, tags=["Classes"])
async def api_create_class(
    req: ClassCreateRequest,
    current_user: Dict[str, Any] = Depends(require_roles(UserRole.ADMIN)),
):
    if len(req.student_ids) > req.capacity and not req.override_capacity:
        raise HTTPException(
            status_code=400,
            detail=f"عدد التلاميذ المسجلين ({len(req.student_ids)}) يتجاوز الطاقة الاستيعابية للقسم ({req.capacity}).",
        )
    async with db.transaction() as session:
        code = await generate_unique_code(session, "CLS", "classes", "code")
        cid = await session.insert(
            """
            INSERT INTO classes (
                code, name, level_id, subject_id, teacher_id, room,
                capacity, billing_mode, monthly_fee, session_fee
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                code, req.name.strip(), req.level_id, req.subject_id, req.teacher_id,
                req.room.strip(), req.capacity, req.billing_mode.value, req.monthly_fee, req.session_fee,
            ),
        )
        for sid in req.student_ids:
            await session.execute(
                """
                INSERT OR IGNORE INTO enrollments (student_id, class_id, status, override_capacity_by)
                VALUES (?, ?, 'ACTIVE', ?)
                """,
                (sid, cid, current_user["id"] if req.override_capacity else None),
            )
        await log_audit(
            session, current_user["id"], current_user["username"], current_user["role"],
            "CREATE_CLASS", "classes", cid, {"code": code, "name": req.name},
        )
        created = await session.fetchone("SELECT * FROM classes WHERE id = ?", (cid,))
    return {"success": True, "class": created}


@app.put("/classes/{id}", tags=["Classes"])
async def api_update_class(
    id: int,
    req: ClassUpdateRequest,
    current_user: Dict[str, Any] = Depends(require_roles(UserRole.ADMIN)),
):
    async with db.transaction() as session:
        target = await session.fetchone("SELECT * FROM classes WHERE id = ?", (id,))
        if not target:
            raise HTTPException(status_code=404, detail="القسم غير موجود")
        updates = model_to_dict(req, exclude_none=True)
        student_ids = updates.pop("student_ids", None)
        override_capacity = updates.pop("override_capacity", False)
        if "billing_mode" in updates and isinstance(updates["billing_mode"], PaymentPlanType):
            updates["billing_mode"] = updates["billing_mode"].value
        if "is_archived" in updates:
            updates["is_archived"] = 1 if updates["is_archived"] else 0

        new_cap = int(updates.get("capacity", target["capacity"]))
        if student_ids is not None and len(student_ids) > new_cap and not override_capacity:
            raise HTTPException(
                status_code=400,
                detail=f"عدد التلاميذ ({len(student_ids)}) يتجاوز الطاقة الاستيعابية ({new_cap}).",
            )

        for col, val in updates.items():
            await session.execute(f"UPDATE classes SET {col} = ?, updated_at = datetime('now') WHERE id = ?", (val, id))

        if student_ids is not None:
            await session.execute("DELETE FROM enrollments WHERE class_id = ?", (id,))
            for sid in student_ids:
                await session.execute(
                    "INSERT OR IGNORE INTO enrollments (student_id, class_id, status, override_capacity_by) VALUES (?, ?, 'ACTIVE', ?)",
                    (sid, id, current_user["id"] if override_capacity else None),
                )

        await log_audit(
            session, current_user["id"], current_user["username"], current_user["role"],
            "UPDATE_CLASS", "classes", id,
        )
        updated = await session.fetchone("SELECT * FROM classes WHERE id = ?", (id,))
    return {"success": True, "class": updated}


@app.post("/classes/{id}/enroll", tags=["Classes"])
async def api_enroll_student_in_class(
    id: int,
    req: EnrollmentCreateRequest,
    current_user: Dict[str, Any] = Depends(require_roles(UserRole.ADMIN)),
):
    async with db.transaction() as session:
        cl = await session.fetchone("SELECT * FROM classes WHERE id = ? AND is_archived = 0", (id,))
        if not cl:
            raise HTTPException(status_code=404, detail="القسم غير موجود")
        cnt_row = await session.fetchone(
            "SELECT COUNT(*) AS cnt FROM enrollments WHERE class_id = ? AND status = 'ACTIVE'",
            (id,),
        )
        current_enrolled = int(cnt_row["cnt"])
        if current_enrolled >= int(cl["capacity"]) and not req.override_capacity:
            raise HTTPException(
                status_code=400,
                detail=f"لا يمكن التسجيل: القسم وصل إلى طاقته الاستيعابية القصوى ({cl['capacity']}). يتطلب تفعيل صلاحية التجاوز الإداري.",
            )
        await session.execute(
            """
            INSERT OR REPLACE INTO enrollments (student_id, class_id, status, billing_override, discount_percent, override_capacity_by)
            VALUES (?, ?, 'ACTIVE', ?, ?, ?)
            """,
            (
                req.student_id, id, req.billing_override, req.discount_percent,
                current_user["id"] if req.override_capacity else None,
            ),
        )
        await log_audit(
            session, current_user["id"], current_user["username"], current_user["role"],
            "ENROLL_STUDENT", "enrollments", f"{id}:{req.student_id}",
            {"override_capacity": req.override_capacity},
        )
    return {"success": True, "message": "تم تسجيل التلميذ في القسم بنجاح"}


@app.delete("/classes/{id}", tags=["Classes"])
async def api_delete_class(
    id: int,
    current_user: Dict[str, Any] = Depends(require_roles(UserRole.ADMIN)),
):
    async with db.transaction() as session:
        await session.execute("UPDATE classes SET is_archived = 1, updated_at = datetime('now') WHERE id = ?", (id,))
        await log_audit(
            session, current_user["id"], current_user["username"], current_user["role"],
            "ARCHIVE_CLASS", "classes", id,
        )
    return {"success": True, "message": "تم أرشفة القسم بنجاح"}


# ------------------------------------------------------------------------------
# Schedules Endpoints (/schedules) with Conflict Detection & Notifications
# ------------------------------------------------------------------------------

@app.get("/schedules", tags=["Schedules"])
async def api_list_schedules(
    day_of_week: Optional[int] = Query(default=None, ge=0, le=6),
    class_id: Optional[int] = None,
    teacher_id: Optional[int] = None,
    status_filter: Optional[str] = None,
    current_user: Dict[str, Any] = Depends(get_current_user),
):
    sql = """
        SELECT s.*, c.name AS class_name, c.code AS class_code,
               sub.name_ar AS subject_name_ar, sub.name_en AS subject_name_en,
               t.full_name AS teacher_name
        FROM schedules s
        JOIN classes c ON c.id = s.class_id
        JOIN subjects sub ON sub.id = s.subject_id
        JOIN teachers t ON t.id = s.teacher_id
        WHERE 1=1
    """
    params: List[Any] = []
    role = current_user["role"]
    if role == UserRole.TEACHER.value:
        tr = await db.fetchone("SELECT id FROM teachers WHERE user_id = ?", (current_user["id"],))
        if tr:
            sql += " AND s.teacher_id = ?"
            params.append(tr["id"])
    elif role == UserRole.STUDENT.value:
        sr = await db.fetchone("SELECT id FROM students WHERE user_id = ?", (current_user["id"],))
        if sr:
            sql += " AND s.class_id IN (SELECT class_id FROM enrollments WHERE student_id = ? AND status = 'ACTIVE')"
            params.append(sr["id"])
    elif role == UserRole.PARENT.value:
        pr = await db.fetchone("SELECT id FROM parents WHERE user_id = ?", (current_user["id"],))
        if pr:
            sql += """
                AND s.class_id IN (
                    SELECT e.class_id FROM enrollments e
                    JOIN students st ON st.id = e.student_id
                    LEFT JOIN parent_students ps ON ps.student_id = st.id
                    WHERE e.status = 'ACTIVE' AND (st.primary_parent_id = ? OR ps.parent_id = ?)
                )
            """
            params.extend([pr["id"], pr["id"]])

    if day_of_week is not None:
        sql += " AND s.day_of_week = ?"
        params.append(day_of_week)
    if class_id is not None:
        sql += " AND s.class_id = ?"
        params.append(class_id)
    if teacher_id is not None:
        sql += " AND s.teacher_id = ?"
        params.append(teacher_id)
    if status_filter:
        sql += " AND s.status = ?"
        params.append(status_filter)

    sql += " ORDER BY s.day_of_week ASC, s.start_time ASC"
    rows = await db.fetchall(sql, params)
    return {"success": True, "items": rows}


@app.post("/schedules", status_code=status.HTTP_201_CREATED, tags=["Schedules"])
async def api_create_schedule(
    req: ScheduleCreateRequest,
    current_user: Dict[str, Any] = Depends(require_roles(UserRole.ADMIN)),
):
    async with db.transaction() as session:
        cl = await session.fetchone("SELECT * FROM classes WHERE id = ?", (req.class_id,))
        if not cl:
            raise HTTPException(status_code=404, detail="القسم غير موجود")
        subject_id = req.subject_id or cl["subject_id"]
        teacher_id = req.teacher_id or cl["teacher_id"]
        room = (req.room or cl["room"]).strip()

        conflict_msg = await check_schedule_conflicts(
            session=session,
            day_of_week=req.day_of_week,
            start_time=req.start_time,
            end_time=req.end_time,
            teacher_id=teacher_id,
            room=room,
            class_id=req.class_id,
            specific_date=req.specific_date,
        )
        if conflict_msg:
            raise HTTPException(status_code=409, detail=conflict_msg)

        duration = _time_to_minutes(req.end_time) - _time_to_minutes(req.start_time)
        sch_id = await session.insert(
            """
            INSERT INTO schedules (
                class_id, subject_id, teacher_id, day_of_week, specific_date,
                start_time, end_time, duration_minutes, room, status, notes
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 'SCHEDULED', ?)
            """,
            (
                req.class_id, subject_id, teacher_id, req.day_of_week, req.specific_date,
                req.start_time, req.end_time, duration, room, req.notes,
            ),
        )
        await log_audit(
            session, current_user["id"], current_user["username"], current_user["role"],
            "CREATE_SCHEDULE", "schedules", sch_id,
            {"class_id": req.class_id, "day": req.day_of_week, "time": f"{req.start_time}-{req.end_time}"},
        )
        created = await session.fetchone("SELECT * FROM schedules WHERE id = ?", (sch_id,))
    return {"success": True, "schedule": created}


@app.put("/schedules/{id}", tags=["Schedules"])
async def api_update_schedule(
    id: int,
    req: ScheduleUpdateRequest,
    current_user: Dict[str, Any] = Depends(require_roles(UserRole.ADMIN)),
):
    async with db.transaction() as session:
        target = await session.fetchone("SELECT * FROM schedules WHERE id = ?", (id,))
        if not target:
            raise HTTPException(status_code=404, detail="الحصة غير موجودة")
        updates = model_to_dict(req, exclude_none=True)

        new_day = int(updates.get("day_of_week", target["day_of_week"]))
        new_start = str(updates.get("start_time", target["start_time"]))
        new_end = str(updates.get("end_time", target["end_time"]))
        new_teacher = int(updates.get("teacher_id", target["teacher_id"]))
        new_room = str(updates.get("room", target["room"]))
        new_status = str(updates.get("status", target["status"]))

        if new_status == "SCHEDULED":
            conflict_msg = await check_schedule_conflicts(
                session=session,
                day_of_week=new_day,
                start_time=new_start,
                end_time=new_end,
                teacher_id=new_teacher,
                room=new_room,
                class_id=int(target["class_id"]),
                specific_date=updates.get("specific_date", target["specific_date"]),
                exclude_schedule_id=id,
            )
            if conflict_msg:
                raise HTTPException(status_code=409, detail=conflict_msg)

        updates["duration_minutes"] = _time_to_minutes(new_end) - _time_to_minutes(new_start)
        for col, val in updates.items():
            await session.execute(f"UPDATE schedules SET {col} = ?, updated_at = datetime('now') WHERE id = ?", (val, id))

        # Notify enrolled students & parents of schedule change
        enrolled = await session.fetchall(
            "SELECT student_id FROM enrollments WHERE class_id = ? AND status = 'ACTIVE'",
            (target["class_id"],),
        )
        cl = await session.fetchone("SELECT name FROM classes WHERE id = ?", (target["class_id"],))
        c_name = cl["name"] if cl else f"#{target['class_id']}"
        for enr in enrolled:
            await notify_student_and_parents(
                session=session,
                student_id=enr["student_id"],
                notification_type=NotificationType.SCHEDULE_CHANGE.value,
                title_ar="تحديث في جدول الحصص",
                title_en="Schedule Updated",
                body_ar=f"تم تعديل موعد حصة القسم '{c_name}' إلى {new_start}-{new_end} (القاعة {new_room}) — الحالة: {new_status}.",
                body_en=f"Class '{c_name}' schedule updated to {new_start}-{new_end} (Room {new_room}) — Status: {new_status}.",
                sender_user_id=current_user["id"],
                related_entity_type="schedules",
                related_entity_id=id,
            )

        await log_audit(
            session, current_user["id"], current_user["username"], current_user["role"],
            "UPDATE_SCHEDULE", "schedules", id, updates,
        )
        updated = await session.fetchone("SELECT * FROM schedules WHERE id = ?", (id,))
    return {"success": True, "schedule": updated}


@app.delete("/schedules/{id}", tags=["Schedules"])
async def api_delete_schedule(
    id: int,
    current_user: Dict[str, Any] = Depends(require_roles(UserRole.ADMIN)),
):
    async with db.transaction() as session:
        target = await session.fetchone("SELECT * FROM schedules WHERE id = ?", (id,))
        if not target:
            raise HTTPException(status_code=404, detail="الحصة غير موجودة")
        await session.execute("UPDATE schedules SET status = 'CANCELLED', updated_at = datetime('now') WHERE id = ?", (id,))
        enrolled = await session.fetchall(
            "SELECT student_id FROM enrollments WHERE class_id = ? AND status = 'ACTIVE'",
            (target["class_id"],),
        )
        for enr in enrolled:
            await notify_student_and_parents(
                session=session,
                student_id=enr["student_id"],
                notification_type=NotificationType.SCHEDULE_CHANGE.value,
                title_ar="إلغاء حصة دراسية",
                title_en="Class Session Cancelled",
                body_ar=f"تم إلغاء الحصة المبرمجة من {target['start_time']} إلى {target['end_time']}.",
                body_en=f"Scheduled session ({target['start_time']} - {target['end_time']}) has been cancelled.",
                sender_user_id=current_user["id"],
                related_entity_type="schedules",
                related_entity_id=id,
            )
        await log_audit(
            session, current_user["id"], current_user["username"], current_user["role"],
            "CANCEL_SCHEDULE", "schedules", id,
        )
    return {"success": True, "message": "تم إلغاء الحصة وإشعار المعنيين بنجاح"}


# ------------------------------------------------------------------------------
# 10.6 Attendance, Exams & Grades, Finance, Materials, Notifications, Reports & Backups
# ------------------------------------------------------------------------------

@app.get("/attendance", tags=["Attendance"])
async def api_list_attendance(
    session_date: Optional[str] = None,
    student_id: Optional[int] = None,
    class_id: Optional[int] = None,
    schedule_id: Optional[int] = None,
    status_filter: Optional[AttendanceStatus] = None,
    page: int = Query(default=1, ge=1),
    page_size: int = Query(default=50, ge=1, le=200),
    current_user: Dict[str, Any] = Depends(get_current_user),
):
    sql = """
        SELECT a.*, s.full_name AS student_name, s.student_code,
               c.name AS class_name, sub.name_ar AS subject_name_ar, sub.name_en AS subject_name_en,
               sch.start_time, sch.end_time
        FROM attendance a
        JOIN students s ON s.id = a.student_id
        JOIN classes c ON c.id = a.class_id
        JOIN subjects sub ON sub.id = c.subject_id
        JOIN schedules sch ON sch.id = a.schedule_id
        WHERE 1=1
    """
    params: List[Any] = []
    role = current_user["role"]

    if role == UserRole.STUDENT.value:
        st_row = await db.fetchone("SELECT id FROM students WHERE user_id = ?", (current_user["id"],))
        if not st_row:
            return paginate_query_results([], page, page_size)
        sql += " AND a.student_id = ?"
        params.append(st_row["id"])
    elif role == UserRole.PARENT.value:
        p_row = await db.fetchone("SELECT id FROM parents WHERE user_id = ?", (current_user["id"],))
        if not p_row:
            return paginate_query_results([], page, page_size)
        sql += """
            AND a.student_id IN (
                SELECT st.id FROM students st
                LEFT JOIN parent_students ps ON ps.student_id = st.id
                WHERE st.primary_parent_id = ? OR ps.parent_id = ?
            )
        """
        params.extend([p_row["id"], p_row["id"]])
    elif role == UserRole.TEACHER.value:
        t_row = await db.fetchone("SELECT id FROM teachers WHERE user_id = ?", (current_user["id"],))
        if not t_row:
            return paginate_query_results([], page, page_size)
        sql += " AND c.teacher_id = ?"
        params.append(t_row["id"])

    if session_date:
        sql += " AND a.session_date = ?"
        params.append(session_date)
    if student_id is not None:
        if not await can_user_access_student(current_user, student_id):
            raise HTTPException(status_code=403, detail="غير مصرح لك بالاطلاع على سجل حضور هذا التلميذ")
        sql += " AND a.student_id = ?"
        params.append(student_id)
    if class_id is not None:
        sql += " AND a.class_id = ?"
        params.append(class_id)
    if schedule_id is not None:
        sql += " AND a.schedule_id = ?"
        params.append(schedule_id)
    if status_filter is not None:
        sql += " AND a.status = ?"
        params.append(status_filter.value)

    sql += " ORDER BY a.session_date DESC, a.id DESC"
    rows = await db.fetchall(sql, params)
    return paginate_query_results(rows, page, page_size)


@app.post("/attendance", status_code=status.HTTP_201_CREATED, tags=["Attendance"])
async def api_record_attendance(
    req: AttendanceCreateRequest,
    current_user: Dict[str, Any] = Depends(require_roles(UserRole.ADMIN, UserRole.TEACHER)),
):
    async with db.transaction() as session:
        sch = await session.fetchone(
            """
            SELECT s.*, c.name AS class_name, sub.name_ar AS subject_name_ar, sub.name_en AS subject_name_en
            FROM schedules s
            JOIN classes c ON c.id = s.class_id
            JOIN subjects sub ON sub.id = s.subject_id
            WHERE s.id = ?
            """,
            (req.schedule_id,),
        )
        if not sch:
            raise HTTPException(status_code=404, detail="الحصة غير موجودة")

        if current_user["role"] == UserRole.TEACHER.value:
            t_row = await session.fetchone("SELECT id FROM teachers WHERE user_id = ?", (current_user["id"],))
            if not t_row or int(sch["teacher_id"]) != int(t_row["id"]):
                raise HTTPException(status_code=403, detail="لا يمكنك تسجيل الحضور لحصة غير مسندة إليك")

        items_to_record: List[AttendanceItemInput] = []
        if req.records:
            items_to_record.extend(req.records)
        elif req.student_id is not None and req.status is not None:
            items_to_record.append(
                AttendanceItemInput(
                    student_id=req.student_id,
                    status=req.status,
                    late_minutes=req.late_minutes,
                    absence_reason=req.absence_reason,
                )
            )
        else:
            raise HTTPException(status_code=400, detail="يرجى تحديد التلميذ وحالة الحضور أو قائمة الحضور الجماعي")

        created_records: List[Dict[str, Any]] = []
        for item in items_to_record:
            dup = await session.fetchone(
                "SELECT id FROM attendance WHERE schedule_id = ? AND student_id = ? AND session_date = ?",
                (req.schedule_id, item.student_id, req.session_date),
            )
            if dup:
                raise HTTPException(
                    status_code=409,
                    detail=f"تم تسجيل الحضور مسبقاً للتلميذ #{item.student_id} في هذه الحصة بتاريخ {req.session_date}. استخدم مسار التعديل PUT /attendance/{{id}}.",
                )

            att_id = await session.insert(
                """
                INSERT INTO attendance (
                    schedule_id, class_id, student_id, session_date, status,
                    late_minutes, absence_reason, recorded_by
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    req.schedule_id,
                    sch["class_id"],
                    item.student_id,
                    req.session_date,
                    item.status.value,
                    item.late_minutes,
                    item.absence_reason,
                    current_user["id"],
                ),
            )

            if item.status in (AttendanceStatus.ABSENT, AttendanceStatus.LATE):
                st = await session.fetchone("SELECT full_name FROM students WHERE id = ?", (item.student_id,))
                st_name = st["full_name"] if st else f"#{item.student_id}"
                if item.status == AttendanceStatus.ABSENT:
                    await notify_student_and_parents(
                        session=session,
                        student_id=item.student_id,
                        notification_type=NotificationType.ABSENCE.value,
                        title_ar="إشعار غياب تلميذ",
                        title_en="Student Absence Alert",
                        body_ar=(
                            f"نود إعلامكم بتسجيل غياب التلميذ {st_name} في حصة {sch['subject_name_ar']} "
                            f"({sch['class_name']}) بتاريخ {req.session_date}."
                        ),
                        body_en=(
                            f"Student {st_name} was marked ABSENT in {sch['subject_name_en']} "
                            f"({sch['class_name']}) on {req.session_date}."
                        ),
                        sender_user_id=current_user["id"],
                        dedup_prefix=f"att:{att_id}",
                        related_entity_type="attendance",
                        related_entity_id=att_id,
                    )
                else:
                    await notify_student_and_parents(
                        session=session,
                        student_id=item.student_id,
                        notification_type=NotificationType.ABSENCE.value,
                        title_ar="إشعار تأخر تلميذ",
                        title_en="Student Late Arrival Alert",
                        body_ar=(
                            f"تم تسجيل تأخر التلميذ {st_name} بـ ({item.late_minutes} دقيقة) في حصة "
                            f"{sch['subject_name_ar']} بتاريخ {req.session_date}."
                        ),
                        body_en=(
                            f"Student {st_name} arrived LATE ({item.late_minutes} min) to "
                            f"{sch['subject_name_en']} on {req.session_date}."
                        ),
                        sender_user_id=current_user["id"],
                        dedup_prefix=f"att:{att_id}",
                        related_entity_type="attendance",
                        related_entity_id=att_id,
                    )

            row = await session.fetchone("SELECT * FROM attendance WHERE id = ?", (att_id,))
            if row:
                created_records.append(row)

        await log_audit(
            session, current_user["id"], current_user["username"], current_user["role"],
            "RECORD_ATTENDANCE", "attendance", req.schedule_id,
            {"session_date": req.session_date, "count": len(created_records)},
        )
    return {"success": True, "count": len(created_records), "items": created_records}


@app.put("/attendance/{id}", tags=["Attendance"])
async def api_update_attendance(
    id: int,
    req: AttendanceUpdateRequest,
    current_user: Dict[str, Any] = Depends(require_roles(UserRole.ADMIN, UserRole.TEACHER)),
):
    async with db.transaction() as session:
        target = await session.fetchone("SELECT * FROM attendance WHERE id = ?", (id,))
        if not target:
            raise HTTPException(status_code=404, detail="سجل الحضور غير موجود")

        if current_user["role"] == UserRole.TEACHER.value:
            t_row = await session.fetchone("SELECT id FROM teachers WHERE user_id = ?", (current_user["id"],))
            cl_row = await session.fetchone("SELECT teacher_id FROM classes WHERE id = ?", (target["class_id"],))
            if not t_row or not cl_row or int(cl_row["teacher_id"]) != int(t_row["id"]):
                raise HTTPException(status_code=403, detail="غير مصرح لك بتعديل سجل هذا القسم")

        await session.execute(
            """
            UPDATE attendance
            SET status = ?, late_minutes = ?, absence_reason = ?, recorded_by = ?, updated_at = datetime('now')
            WHERE id = ?
            """,
            (req.status.value, req.late_minutes, req.absence_reason, current_user["id"], id),
        )
        await log_audit(
            session, current_user["id"], current_user["username"], current_user["role"],
            "UPDATE_ATTENDANCE", "attendance", id,
            {"status": req.status.value, "late_minutes": req.late_minutes},
        )
        updated = await session.fetchone("SELECT * FROM attendance WHERE id = ?", (id,))
    return {"success": True, "attendance": updated}


# ------------------------------------------------------------------------------
# Exams & Grades Endpoints (/exams, /grades)
# ------------------------------------------------------------------------------

@app.get("/exams", tags=["Exams & Grades"])
async def api_list_exams(
    class_id: Optional[int] = None,
    subject_id: Optional[int] = None,
    current_user: Dict[str, Any] = Depends(get_current_user),
):
    sql = """
        SELECT ex.*, sub.name_ar AS subject_name_ar, sub.name_en AS subject_name_en,
               c.name AS class_name, t.full_name AS teacher_name,
               (SELECT COUNT(*) FROM grades g WHERE g.exam_id = ex.id) AS graded_count
        FROM exams ex
        JOIN subjects sub ON sub.id = ex.subject_id
        JOIN classes c ON c.id = ex.class_id
        JOIN teachers t ON t.id = ex.teacher_id
        WHERE 1=1
    """
    params: List[Any] = []
    role = current_user["role"]
    if role == UserRole.TEACHER.value:
        tr = await db.fetchone("SELECT id FROM teachers WHERE user_id = ?", (current_user["id"],))
        if tr:
            sql += " AND ex.teacher_id = ?"
            params.append(tr["id"])
    elif role == UserRole.STUDENT.value:
        sr = await db.fetchone("SELECT id FROM students WHERE user_id = ?", (current_user["id"],))
        if sr:
            sql += " AND ex.class_id IN (SELECT class_id FROM enrollments WHERE student_id = ? AND status = 'ACTIVE')"
            params.append(sr["id"])
    elif role == UserRole.PARENT.value:
        pr = await db.fetchone("SELECT id FROM parents WHERE user_id = ?", (current_user["id"],))
        if pr:
            sql += """
                AND ex.class_id IN (
                    SELECT e.class_id FROM enrollments e
                    JOIN students st ON st.id = e.student_id
                    LEFT JOIN parent_students ps ON ps.student_id = st.id
                    WHERE st.primary_parent_id = ? OR ps.parent_id = ?
                )
            """
            params.extend([pr["id"], pr["id"]])

    if class_id is not None:
        sql += " AND ex.class_id = ?"
        params.append(class_id)
    if subject_id is not None:
        sql += " AND ex.subject_id = ?"
        params.append(subject_id)

    sql += " ORDER BY ex.exam_date DESC, ex.id DESC"
    rows = await db.fetchall(sql, params)
    return {"success": True, "items": rows}


@app.post("/exams", status_code=status.HTTP_201_CREATED, tags=["Exams & Grades"])
async def api_create_exam(
    req: ExamCreateRequest,
    current_user: Dict[str, Any] = Depends(require_roles(UserRole.ADMIN, UserRole.TEACHER)),
):
    async with db.transaction() as session:
        cl = await session.fetchone("SELECT * FROM classes WHERE id = ?", (req.class_id,))
        if not cl:
            raise HTTPException(status_code=404, detail="القسم غير موجود")
        teacher_id = req.teacher_id or cl["teacher_id"]
        if current_user["role"] == UserRole.TEACHER.value:
            tr = await session.fetchone("SELECT id FROM teachers WHERE user_id = ?", (current_user["id"],))
            if not tr or int(cl["teacher_id"]) != int(tr["id"]):
                raise HTTPException(status_code=403, detail="لا يمكنك إنشاء اختبار لقسم غير مسند إليك")
            teacher_id = tr["id"]

        ex_id = await session.insert(
            """
            INSERT INTO exams (title, subject_id, class_id, teacher_id, exam_date, max_score, weight, show_ranking, description)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                req.title.strip(), req.subject_id, req.class_id, teacher_id,
                req.exam_date, req.max_score, req.weight, 1 if req.show_ranking else 0, req.description,
            ),
        )
        await log_audit(
            session, current_user["id"], current_user["username"], current_user["role"],
            "CREATE_EXAM", "exams", ex_id, {"title": req.title, "max_score": req.max_score},
        )
        created = await session.fetchone("SELECT * FROM exams WHERE id = ?", (ex_id,))
    return {"success": True, "exam": created}


@app.get("/grades", tags=["Exams & Grades"])
async def api_list_grades(
    exam_id: Optional[int] = None,
    student_id: Optional[int] = None,
    class_id: Optional[int] = None,
    current_user: Dict[str, Any] = Depends(get_current_user),
):
    sql = """
        SELECT g.*, s.full_name AS student_name, s.student_code,
               ex.title AS exam_title, ex.exam_date, ex.max_score, ex.weight, ex.show_ranking,
               sub.name_ar AS subject_name_ar, sub.name_en AS subject_name_en,
               c.name AS class_name
        FROM grades g
        JOIN students s ON s.id = g.student_id
        JOIN exams ex ON ex.id = g.exam_id
        JOIN subjects sub ON sub.id = ex.subject_id
        JOIN classes c ON c.id = ex.class_id
        WHERE 1=1
    """
    params: List[Any] = []
    role = current_user["role"]
    if role == UserRole.STUDENT.value:
        sr = await db.fetchone("SELECT id FROM students WHERE user_id = ?", (current_user["id"],))
        if not sr:
            return {"success": True, "items": []}
        sql += " AND g.student_id = ?"
        params.append(sr["id"])
    elif role == UserRole.PARENT.value:
        pr = await db.fetchone("SELECT id FROM parents WHERE user_id = ?", (current_user["id"],))
        if not pr:
            return {"success": True, "items": []}
        sql += """
            AND g.student_id IN (
                SELECT st.id FROM students st
                LEFT JOIN parent_students ps ON ps.student_id = st.id
                WHERE st.primary_parent_id = ? OR ps.parent_id = ?
            )
        """
        params.extend([pr["id"], pr["id"]])
    elif role == UserRole.TEACHER.value:
        tr = await db.fetchone("SELECT id FROM teachers WHERE user_id = ?", (current_user["id"],))
        if tr:
            sql += " AND ex.teacher_id = ?"
            params.append(tr["id"])

    if exam_id is not None:
        sql += " AND g.exam_id = ?"
        params.append(exam_id)
    if student_id is not None:
        if not await can_user_access_student(current_user, student_id):
            raise HTTPException(status_code=403, detail="غير مصرح لك بالاطلاع على نتائج هذا التلميذ")
        sql += " AND g.student_id = ?"
        params.append(student_id)
    if class_id is not None:
        sql += " AND ex.class_id = ?"
        params.append(class_id)

    sql += " ORDER BY g.score DESC, ex.exam_date DESC"
    rows = await db.fetchall(sql, params)

    # Calculate ranking when exam_id is specified and show_ranking is enabled
    global_rank_setting = await db.fetchone("SELECT value FROM app_settings WHERE key = 'enable_grade_ranking'")
    ranking_enabled = (global_rank_setting and global_rank_setting["value"] == "1")
    if exam_id is not None and ranking_enabled:
        all_exam_grades = await db.fetchall(
            "SELECT student_id, score FROM grades WHERE exam_id = ? ORDER BY score DESC",
            (exam_id,),
        )
        rank_map = {r["student_id"]: idx + 1 for idx, r in enumerate(all_exam_grades)}
        for r in rows:
            if r.get("show_ranking"):
                r["rank"] = rank_map.get(r["student_id"])
    return {"success": True, "items": rows}


@app.post("/grades", status_code=status.HTTP_201_CREATED, tags=["Exams & Grades"])
async def api_record_grades(
    req: GradeCreateRequest,
    current_user: Dict[str, Any] = Depends(require_roles(UserRole.ADMIN, UserRole.TEACHER)),
):
    async with db.transaction() as session:
        exam = await session.fetchone(
            """
            SELECT ex.*, sub.name_ar AS subject_name_ar, sub.name_en AS subject_name_en, c.name AS class_name
            FROM exams ex
            JOIN subjects sub ON sub.id = ex.subject_id
            JOIN classes c ON c.id = ex.class_id
            WHERE ex.id = ?
            """,
            (req.exam_id,),
        )
        if not exam:
            raise HTTPException(status_code=404, detail="الاختبار غير موجود")

        if current_user["role"] == UserRole.TEACHER.value:
            tr = await session.fetchone("SELECT id FROM teachers WHERE user_id = ?", (current_user["id"],))
            if not tr or int(exam["teacher_id"]) != int(tr["id"]):
                raise HTTPException(status_code=403, detail="لا يمكنك رصد علامات لاختبار غير مسند إليك")

        max_score = float(exam["max_score"])
        items: List[GradeItemInput] = []
        if req.records:
            items.extend(req.records)
        elif req.student_id is not None and req.score is not None:
            items.append(
                GradeItemInput(
                    student_id=req.student_id,
                    score=req.score,
                    teacher_remarks=req.teacher_remarks,
                )
            )
        else:
            raise HTTPException(status_code=400, detail="يرجى إدخال علامة التلميذ أو قائمة العلامات")

        saved_rows: List[Dict[str, Any]] = []
        for item in items:
            if item.score < 0 or item.score > max_score:
                raise HTTPException(
                    status_code=400,
                    detail=f"العلامة ({item.score}) غير منطقية: يجب أن تكون بين 0 و العلامة القصوى ({max_score}).",
                )
            existing = await session.fetchone(
                "SELECT id FROM grades WHERE exam_id = ? AND student_id = ?",
                (req.exam_id, item.student_id),
            )
            if existing:
                gid = existing["id"]
                await session.execute(
                    """
                    UPDATE grades
                    SET score = ?, teacher_remarks = ?, recorded_by = ?, updated_at = datetime('now')
                    WHERE id = ?
                    """,
                    (item.score, item.teacher_remarks, current_user["id"], gid),
                )
            else:
                gid = await session.insert(
                    """
                    INSERT INTO grades (exam_id, student_id, score, teacher_remarks, recorded_by)
                    VALUES (?, ?, ?, ?, ?)
                    """,
                    (req.exam_id, item.student_id, item.score, item.teacher_remarks, current_user["id"]),
                )

            st = await session.fetchone("SELECT full_name FROM students WHERE id = ?", (item.student_id,))
            st_name = st["full_name"] if st else f"#{item.student_id}"
            await notify_student_and_parents(
                session=session,
                student_id=item.student_id,
                notification_type=NotificationType.NEW_GRADE.value,
                title_ar="نشر نتيجة اختبار جديدة",
                title_en="New Exam Grade Published",
                body_ar=(
                    f"تحصل التلميذ {st_name} على العلامة {item.score}/{max_score} في اختبار "
                    f"'{exam['title']}' ({exam['subject_name_ar']})."
                ),
                body_en=(
                    f"Student {st_name} scored {item.score}/{max_score} in '{exam['title']}' "
                    f"({exam['subject_name_en']})."
                ),
                sender_user_id=current_user["id"],
                dedup_prefix=f"grade:{gid}:{item.score}",
                related_entity_type="grades",
                related_entity_id=gid,
            )

            row = await session.fetchone("SELECT * FROM grades WHERE id = ?", (gid,))
            if row:
                saved_rows.append(row)

        await log_audit(
            session, current_user["id"], current_user["username"], current_user["role"],
            "RECORD_GRADES", "grades", req.exam_id, {"count": len(saved_rows)},
        )
    return {"success": True, "count": len(saved_rows), "items": saved_rows}


@app.put("/grades/{id}", tags=["Exams & Grades"])
async def api_update_grade(
    id: int,
    req: GradeUpdateRequest,
    current_user: Dict[str, Any] = Depends(require_roles(UserRole.ADMIN, UserRole.TEACHER)),
):
    async with db.transaction() as session:
        grade = await session.fetchone("SELECT * FROM grades WHERE id = ?", (id,))
        if not grade:
            raise HTTPException(status_code=404, detail="النتيجة غير موجودة")
        exam = await session.fetchone("SELECT * FROM exams WHERE id = ?", (grade["exam_id"],))
        if not exam:
            raise HTTPException(status_code=404, detail="الاختبار غير موجود")

        if current_user["role"] == UserRole.TEACHER.value:
            tr = await session.fetchone("SELECT id FROM teachers WHERE user_id = ?", (current_user["id"],))
            if not tr or int(exam["teacher_id"]) != int(tr["id"]):
                raise HTTPException(status_code=403, detail="غير مصرح لك بتعديل هذه النتيجة")

        max_score = float(exam["max_score"])
        if req.score < 0 or req.score > max_score:
            raise HTTPException(
                status_code=400,
                detail=f"العلامة ({req.score}) تتجاوز العلامة القصوى المسموحة ({max_score}).",
            )

        await session.execute(
            """
            UPDATE grades
            SET score = ?, teacher_remarks = ?, recorded_by = ?, updated_at = datetime('now')
            WHERE id = ?
            """,
            (req.score, req.teacher_remarks, current_user["id"], id),
        )
        await log_audit(
            session, current_user["id"], current_user["username"], current_user["role"],
            "UPDATE_GRADE", "grades", id, {"old_score": grade["score"], "new_score": req.score},
        )
        updated = await session.fetchone("SELECT * FROM grades WHERE id = ?", (id,))
    return {"success": True, "grade": updated}


# ------------------------------------------------------------------------------
# Finance & Payments Endpoints (/payment-plans, /payments, /reports/finance)
# ------------------------------------------------------------------------------

@app.get("/payment-plans", tags=["Payments & Finance"])
async def api_list_payment_plans(
    student_id: Optional[int] = None,
    status_filter: Optional[str] = None,
    current_user: Dict[str, Any] = Depends(get_current_user),
):
    sql = """
        SELECT pp.*, s.full_name AS student_name, s.student_code,
               c.name AS class_name, sub.name_ar AS subject_name_ar
        FROM payment_plans pp
        JOIN students s ON s.id = pp.student_id
        LEFT JOIN classes c ON c.id = pp.class_id
        LEFT JOIN subjects sub ON sub.id = pp.subject_id
        WHERE pp.status != 'VOIDED'
    """
    params: List[Any] = []
    role = current_user["role"]
    if role == UserRole.STUDENT.value:
        sr = await db.fetchone("SELECT id FROM students WHERE user_id = ?", (current_user["id"],))
        if not sr:
            return {"success": True, "items": []}
        sql += " AND pp.student_id = ?"
        params.append(sr["id"])
    elif role == UserRole.PARENT.value:
        pr = await db.fetchone("SELECT id FROM parents WHERE user_id = ?", (current_user["id"],))
        if not pr:
            return {"success": True, "items": []}
        sql += """
            AND pp.student_id IN (
                SELECT st.id FROM students st
                LEFT JOIN parent_students ps ON ps.student_id = st.id
                WHERE st.primary_parent_id = ? OR ps.parent_id = ?
            )
        """
        params.extend([pr["id"], pr["id"]])
    elif role == UserRole.TEACHER.value:
        raise HTTPException(status_code=403, detail="البيانات المالية مخصصة للإدارة وأولياء الأمور والتلاميذ المعنيين فقط")

    if student_id is not None:
        if not await can_user_access_student(current_user, student_id):
            raise HTTPException(status_code=403, detail="غير مصرح لك بالاطلاع على اشتراكات هذا التلميذ")
        sql += " AND pp.student_id = ?"
        params.append(student_id)
    if status_filter:
        sql += " AND pp.status = ?"
        params.append(status_filter)

    sql += " ORDER BY pp.due_date ASC, pp.id DESC"
    rows = await db.fetchall(sql, params)
    return {"success": True, "items": rows}


@app.post("/payment-plans", status_code=status.HTTP_201_CREATED, tags=["Payments & Finance"])
async def api_create_payment_plan(
    req: PaymentPlanCreateRequest,
    current_user: Dict[str, Any] = Depends(require_roles(UserRole.ADMIN)),
):
    async with db.transaction() as session:
        plan_code = await generate_unique_code(session, "PLN", "payment_plans", "plan_code")
        pid = await session.insert(
            """
            INSERT INTO payment_plans (
                plan_code, student_id, class_id, subject_id, plan_type,
                sessions_count, total_amount, paid_amount, remaining_amount,
                due_date, period_label, status, notes, created_by
            ) VALUES (?, ?, ?, ?, ?, ?, ?, 0.0, ?, ?, ?, 'UNPAID', ?, ?)
            """,
            (
                plan_code, req.student_id, req.class_id, req.subject_id, req.plan_type,
                req.sessions_count, round(req.total_amount, 2), round(req.total_amount, 2),
                req.due_date, req.period_label, req.notes, current_user["id"],
            ),
        )
        await log_audit(
            session, current_user["id"], current_user["username"], current_user["role"],
            "CREATE_PAYMENT_PLAN", "payment_plans", pid,
            {"plan_code": plan_code, "student_id": req.student_id, "total_amount": req.total_amount},
        )
        row = await session.fetchone("SELECT * FROM payment_plans WHERE id = ?", (pid,))
    return {"success": True, "payment_plan": row}


@app.get("/payments", tags=["Payments & Finance"])
async def api_list_payments(
    student_id: Optional[int] = None,
    search: Optional[str] = None,
    payment_method: Optional[str] = None,
    page: int = Query(default=1, ge=1),
    page_size: int = Query(default=25, ge=1, le=200),
    current_user: Dict[str, Any] = Depends(get_current_user),
):
    sql = """
        SELECT p.*, s.full_name AS student_name, s.student_code
        FROM payments p
        JOIN students s ON s.id = p.student_id
        WHERE p.is_voided = 0
    """
    params: List[Any] = []
    role = current_user["role"]
    if role == UserRole.STUDENT.value:
        sr = await db.fetchone("SELECT id FROM students WHERE user_id = ?", (current_user["id"],))
        if not sr:
            return paginate_query_results([], page, page_size)
        sql += " AND p.student_id = ?"
        params.append(sr["id"])
    elif role == UserRole.PARENT.value:
        pr = await db.fetchone("SELECT id FROM parents WHERE user_id = ?", (current_user["id"],))
        if not pr:
            return paginate_query_results([], page, page_size)
        sql += """
            AND p.student_id IN (
                SELECT st.id FROM students st
                LEFT JOIN parent_students ps ON ps.student_id = st.id
                WHERE st.primary_parent_id = ? OR ps.parent_id = ?
            )
        """
        params.extend([pr["id"], pr["id"]])
    elif role == UserRole.TEACHER.value:
        raise HTTPException(status_code=403, detail="غير مصرح للأستاذ بالاطلاع على المعاملات المالية")

    if student_id is not None:
        if not await can_user_access_student(current_user, student_id):
            raise HTTPException(status_code=403, detail="غير مصرح لك بالاطلاع على مدفوعات هذا التلميذ")
        sql += " AND p.student_id = ?"
        params.append(student_id)
    if payment_method:
        sql += " AND p.payment_method = ?"
        params.append(payment_method)
    if search:
        sql += " AND (p.receipt_number LIKE ? OR s.full_name LIKE ? OR s.student_code LIKE ? OR p.description LIKE ?)"
        q = f"%{search.strip()}%"
        params.extend([q, q, q, q])

    sql += " ORDER BY p.id DESC"
    rows = await db.fetchall(sql, params)
    return paginate_query_results(rows, page, page_size)


@app.post("/payments", status_code=status.HTTP_201_CREATED, tags=["Payments & Finance"])
async def api_create_payment(
    req: PaymentCreateRequest,
    current_user: Dict[str, Any] = Depends(require_roles(UserRole.ADMIN)),
):
    async with db.transaction() as session:
        target_plan_id = req.payment_plan_id
        if req.create_plan_type and req.create_plan_total_amount is not None:
            plan_code = await generate_unique_code(session, "PLN", "payment_plans", "plan_code")
            due = req.create_plan_due_date or datetime.date.today().isoformat()
            target_plan_id = await session.insert(
                """
                INSERT INTO payment_plans (
                    plan_code, student_id, class_id, plan_type, total_amount,
                    paid_amount, remaining_amount, due_date, status, created_by
                ) VALUES (?, ?, ?, ?, ?, 0.0, ?, ?, 'UNPAID', ?)
                """,
                (
                    plan_code, req.student_id, req.create_plan_class_id, req.create_plan_type,
                    round(req.create_plan_total_amount, 2), round(req.create_plan_total_amount, 2),
                    due, current_user["id"],
                ),
            )

        res = await record_payment_transaction(
            session=session,
            student_id=req.student_id,
            amount=req.amount,
            payment_method=req.payment_method,
            description=req.description,
            actor_user=current_user,
            payment_plan_id=target_plan_id,
        )
    return {"success": True, **res}


@app.get("/payments/{id}", tags=["Payments & Finance"])
async def api_get_payment(
    id: int,
    format: Optional[str] = Query(default=None),
    current_user: Dict[str, Any] = Depends(get_current_user),
):
    payment = await db.fetchone(
        """
        SELECT p.*, s.full_name AS student_name, s.student_code
        FROM payments p
        JOIN students s ON s.id = p.student_id
        WHERE p.id = ?
        """,
        (id,),
    )
    if not payment:
        raise HTTPException(status_code=404, detail="عملية الدفع غير موجودة")
    if not await can_user_access_student(current_user, payment["student_id"]):
        raise HTTPException(status_code=403, detail="غير مصرح لك بالاطلاع على هذا الإيصال")

    allocations = await db.fetchall(
        """
        SELECT pa.*, pp.plan_code, pp.plan_type, pp.period_label, pp.remaining_amount AS remaining_after
        FROM payment_allocations pa
        JOIN payment_plans pp ON pp.id = pa.payment_plan_id
        WHERE pa.payment_id = ?
        """,
        (id,),
    )
    debt_row = await db.fetchone(
        "SELECT COALESCE(SUM(remaining_amount), 0) AS total_debt FROM payment_plans WHERE student_id = ? AND status != 'VOIDED'",
        (payment["student_id"],),
    )
    settings_rows = await db.fetchall("SELECT key, value FROM app_settings")
    settings_map = {r["key"]: r["value"] for r in settings_rows}

    receipt_payload = {
        **payment,
        "allocations": allocations,
        "total_student_debt": round(float(debt_row["total_debt"] or 0.0), 2),
        "school_name_ar": settings_map.get("school_name_ar", "MOUNIX SCHOOL"),
        "school_address": settings_map.get("school_address", ""),
        "school_phone": settings_map.get("school_phone", ""),
        "currency": settings_map.get("currency", "DZD"),
    }

    if format == "pdf":
        pdf_bytes = await asyncio.to_thread(generate_payment_receipt_pdf_bytes, receipt_payload)
        return Response(
            content=pdf_bytes,
            media_type="application/pdf",
            headers={"Content-Disposition": f'attachment; filename="receipt_{payment["receipt_number"]}.pdf"'},
        )

    return {"success": True, "receipt": receipt_payload}


# ------------------------------------------------------------------------------
# 10.7 Materials & Assignments Endpoints (/materials)
# ------------------------------------------------------------------------------

def validate_safe_filename(filename: str) -> str:
    ext = Path(filename).suffix.lower()
    if ext in FORBIDDEN_EXTENSIONS or ext not in ALLOWED_UPLOAD_EXTENSIONS:
        raise HTTPException(
            status_code=400,
            detail=f"نوع الملف ({ext}) غير مسموح به لأسباب أمنية. الأنواع المسموحة: {', '.join(sorted(ALLOWED_UPLOAD_EXTENSIONS))}",
        )
    safe_stem = re.sub(r"[^a-zA-Z0-9_\-\u0600-\u06FF]", "_", Path(filename).stem)[:60]
    return f"{datetime.datetime.now().strftime('%Y%m%d%H%M%S')}_{secrets.token_hex(4)}_{safe_stem}{ext}"


@app.get("/materials", tags=["Materials & Assignments"])
async def api_list_materials(
    subject_id: Optional[int] = None,
    class_id: Optional[int] = None,
    material_type: Optional[MaterialType] = None,
    include_archived: bool = False,
    current_user: Dict[str, Any] = Depends(get_current_user),
):
    role = current_user["role"]
    params_mat: List[Any] = []
    sql_mat = """
        SELECT m.id, m.title, m.description, m.material_type, m.subject_id, m.class_id,
               m.teacher_id, m.file_path, m.external_url, m.publish_date, NULL AS due_date,
               m.is_archived, m.created_at,
               sub.name_ar AS subject_name_ar, sub.name_en AS subject_name_en,
               c.name AS class_name, t.full_name AS teacher_name
        FROM materials m
        JOIN subjects sub ON sub.id = m.subject_id
        LEFT JOIN classes c ON c.id = m.class_id
        JOIN teachers t ON t.id = m.teacher_id
        WHERE 1=1
    """
    params_asg: List[Any] = []
    sql_asg = """
        SELECT a.id, a.title, a.description, 'ASSIGNMENT' AS material_type, a.subject_id, a.class_id,
               a.teacher_id, a.file_path, a.external_url, a.publish_date, a.due_date,
               a.is_archived, a.created_at,
               sub.name_ar AS subject_name_ar, sub.name_en AS subject_name_en,
               c.name AS class_name, t.full_name AS teacher_name
        FROM assignments a
        JOIN subjects sub ON sub.id = a.subject_id
        JOIN classes c ON c.id = a.class_id
        JOIN teachers t ON t.id = a.teacher_id
        WHERE 1=1
    """

    if not include_archived:
        sql_mat += " AND m.is_archived = 0"
        sql_asg += " AND a.is_archived = 0"

    if role == UserRole.STUDENT.value:
        sr = await db.fetchone("SELECT id FROM students WHERE user_id = ?", (current_user["id"],))
        if not sr:
            return {"success": True, "items": []}
        sql_mat += " AND (m.class_id IS NULL OR m.class_id IN (SELECT class_id FROM enrollments WHERE student_id = ? AND status = 'ACTIVE'))"
        params_mat.append(sr["id"])
        sql_asg += " AND a.class_id IN (SELECT class_id FROM enrollments WHERE student_id = ? AND status = 'ACTIVE')"
        params_asg.append(sr["id"])
    elif role == UserRole.PARENT.value:
        pr = await db.fetchone("SELECT id FROM parents WHERE user_id = ?", (current_user["id"],))
        if not pr:
            return {"success": True, "items": []}
        sub_q = """
            SELECT e.class_id FROM enrollments e
            JOIN students st ON st.id = e.student_id
            LEFT JOIN parent_students ps ON ps.student_id = st.id
            WHERE st.primary_parent_id = ? OR ps.parent_id = ?
        """
        sql_mat += f" AND (m.class_id IS NULL OR m.class_id IN ({sub_q}))"
        params_mat.extend([pr["id"], pr["id"]])
        sql_asg += f" AND a.class_id IN ({sub_q})"
        params_asg.extend([pr["id"], pr["id"]])

    if subject_id is not None:
        sql_mat += " AND m.subject_id = ?"
        params_mat.append(subject_id)
        sql_asg += " AND a.subject_id = ?"
        params_asg.append(subject_id)
    if class_id is not None:
        sql_mat += " AND m.class_id = ?"
        params_mat.append(class_id)
        sql_asg += " AND a.class_id = ?"
        params_asg.append(class_id)

    items: List[Dict[str, Any]] = []
    if material_type is None or material_type != MaterialType.ASSIGNMENT:
        mat_rows = await db.fetchall(sql_mat + " ORDER BY m.id DESC", params_mat)
        if material_type is not None:
            mat_rows = [r for r in mat_rows if r["material_type"] == material_type.value]
        items.extend(mat_rows)
    if material_type is None or material_type == MaterialType.ASSIGNMENT:
        asg_rows = await db.fetchall(sql_asg + " ORDER BY a.id DESC", params_asg)
        items.extend(asg_rows)

    return {"success": True, "items": items}


@app.post("/materials", status_code=status.HTTP_201_CREATED, tags=["Materials & Assignments"])
async def api_create_material(
    req: MaterialCreateRequest,
    current_user: Dict[str, Any] = Depends(require_roles(UserRole.ADMIN, UserRole.TEACHER)),
):
    async with db.transaction() as session:
        teacher_id = req.teacher_id
        if current_user["role"] == UserRole.TEACHER.value:
            tr = await session.fetchone("SELECT id FROM teachers WHERE user_id = ?", (current_user["id"],))
            if not tr:
                raise HTTPException(status_code=403, detail="لم يتم العثور على ملف الأستاذ")
            teacher_id = tr["id"]
        elif teacher_id is None:
            if req.class_id:
                cl = await session.fetchone("SELECT teacher_id FROM classes WHERE id = ?", (req.class_id,))
                teacher_id = cl["teacher_id"] if cl else None
            if teacher_id is None:
                first_t = await session.fetchone("SELECT id FROM teachers LIMIT 1")
                teacher_id = first_t["id"] if first_t else 1

        pub_date = req.publish_date or datetime.date.today().isoformat()
        if req.material_type == MaterialType.ASSIGNMENT:
            if not req.class_id or not req.due_date:
                raise HTTPException(status_code=400, detail="الواجبات تتطلب تحديد القسم (class_id) وتاريخ التسليم (due_date)")
            asg_id = await session.insert(
                """
                INSERT INTO assignments (
                    title, description, subject_id, class_id, teacher_id,
                    external_url, publish_date, due_date, max_score
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    req.title.strip(), req.description, req.subject_id, req.class_id,
                    teacher_id, req.external_url, pub_date, req.due_date, req.max_score or 20.0,
                ),
            )
            enrolled = await session.fetchall(
                "SELECT student_id FROM enrollments WHERE class_id = ? AND status = 'ACTIVE'",
                (req.class_id,),
            )
            for enr in enrolled:
                await notify_student_and_parents(
                    session=session,
                    student_id=enr["student_id"],
                    notification_type=NotificationType.NEW_ASSIGNMENT.value,
                    title_ar="واجب دراسي جديد",
                    title_en="New Assignment Published",
                    body_ar=f"تم نشر واجب جديد: '{req.title}' — آخر أجل للتسليم: {req.due_date}.",
                    body_en=f"New assignment posted: '{req.title}' — Due date: {req.due_date}.",
                    sender_user_id=current_user["id"],
                    dedup_prefix=f"asg:{asg_id}",
                    related_entity_type="assignments",
                    related_entity_id=asg_id,
                )
            await log_audit(
                session, current_user["id"], current_user["username"], current_user["role"],
                "CREATE_ASSIGNMENT", "assignments", asg_id, {"title": req.title},
            )
            created = await session.fetchone("SELECT * FROM assignments WHERE id = ?", (asg_id,))
            return {"success": True, "item_category": "ASSIGNMENT", "material": created}
        else:
            mat_id = await session.insert(
                """
                INSERT INTO materials (
                    title, description, material_type, subject_id, class_id,
                    teacher_id, external_url, publish_date
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    req.title.strip(), req.description, req.material_type.value,
                    req.subject_id, req.class_id, teacher_id, req.external_url, pub_date,
                ),
            )
            if req.class_id:
                enrolled = await session.fetchall(
                    "SELECT student_id FROM enrollments WHERE class_id = ? AND status = 'ACTIVE'",
                    (req.class_id,),
                )
                for enr in enrolled:
                    await notify_student_and_parents(
                        session=session,
                        student_id=enr["student_id"],
                        notification_type=NotificationType.NEW_MATERIAL.value,
                        title_ar="ملف تعليمي جديد",
                        title_en="New Educational Material",
                        body_ar=f"تم إضافة محتوى تعليمي جديد: '{req.title}'.",
                        body_en=f"New learning material published: '{req.title}'.",
                        sender_user_id=current_user["id"],
                        dedup_prefix=f"mat:{mat_id}",
                        related_entity_type="materials",
                        related_entity_id=mat_id,
                    )
            await log_audit(
                session, current_user["id"], current_user["username"], current_user["role"],
                "CREATE_MATERIAL", "materials", mat_id, {"title": req.title},
            )
            created = await session.fetchone("SELECT * FROM materials WHERE id = ?", (mat_id,))
            return {"success": True, "item_category": "MATERIAL", "material": created}


@app.post("/materials/upload", status_code=status.HTTP_201_CREATED, tags=["Materials & Assignments"])
async def api_upload_material_file(
    title: str = Form(...),
    subject_id: int = Form(...),
    class_id: Optional[int] = Form(default=None),
    description: Optional[str] = Form(default=None),
    file: UploadFile = File(...),
    current_user: Dict[str, Any] = Depends(require_roles(UserRole.ADMIN, UserRole.TEACHER)),
):
    if not file.filename:
        raise HTTPException(status_code=400, detail="اسم الملف مفقود")
    safe_name = validate_safe_filename(file.filename)
    content = await file.read()
    if len(content) > MAX_UPLOAD_SIZE_BYTES:
        raise HTTPException(status_code=400, detail="حجم الملف يتجاوز الحد الأقصى المسموح (25MB)")
    # Block executable binary headers (MZ / ELF)
    if content.startswith(b"MZ") or content.startswith(b"\x7fELF"):
        raise HTTPException(status_code=400, detail="ممنوع رفع ملفات تنفيذية")

    dest = UPLOAD_DIR / safe_name
    await asyncio.to_thread(dest.write_bytes, content)

    async with db.transaction() as session:
        tr = await session.fetchone("SELECT id FROM teachers WHERE user_id = ?", (current_user["id"],))
        teacher_id = tr["id"] if tr else 1
        mat_id = await session.insert(
            """
            INSERT INTO materials (title, description, material_type, subject_id, class_id, teacher_id, file_path)
            VALUES (?, ?, 'PDF', ?, ?, ?, ?)
            """,
            (title.strip(), description, subject_id, class_id, teacher_id, str(dest)),
        )
        await log_audit(
            session, current_user["id"], current_user["username"], current_user["role"],
            "UPLOAD_MATERIAL_FILE", "materials", mat_id, {"filename": safe_name},
        )
        created = await session.fetchone("SELECT * FROM materials WHERE id = ?", (mat_id,))
    return {"success": True, "material": created}


@app.put("/materials/{id}", tags=["Materials & Assignments"])
async def api_update_material(
    id: int,
    req: MaterialUpdateRequest,
    is_assignment: bool = Query(default=False),
    current_user: Dict[str, Any] = Depends(require_roles(UserRole.ADMIN, UserRole.TEACHER)),
):
    table = "assignments" if is_assignment else "materials"
    async with db.transaction() as session:
        target = await session.fetchone(f"SELECT * FROM {table} WHERE id = ?", (id,))
        if not target:
            raise HTTPException(status_code=404, detail="المحتوى التعليمي غير موجود")
        updates = model_to_dict(req, exclude_none=True)
        if not is_assignment:
            updates.pop("due_date", None)
        if "is_archived" in updates:
            updates["is_archived"] = 1 if updates["is_archived"] else 0
        for col, val in updates.items():
            await session.execute(f"UPDATE {table} SET {col} = ?, updated_at = datetime('now') WHERE id = ?", (val, id))
        updated = await session.fetchone(f"SELECT * FROM {table} WHERE id = ?", (id,))
    return {"success": True, "material": updated}


# ------------------------------------------------------------------------------
# 10.8 Notifications Endpoints (/notifications)
# ------------------------------------------------------------------------------

@app.get("/notifications", tags=["Notifications"])
async def api_list_notifications(
    unread_only: bool = False,
    page: int = Query(default=1, ge=1),
    page_size: int = Query(default=30, ge=1, le=100),
    current_user: Dict[str, Any] = Depends(get_current_user),
):
    sql = "SELECT * FROM notifications WHERE recipient_user_id = ?"
    params: List[Any] = [current_user["id"]]
    if unread_only:
        sql += " AND is_read = 0"
    sql += " ORDER BY id DESC"
    rows = await db.fetchall(sql, params)
    return paginate_query_results(rows, page, page_size)


@app.post("/notifications", status_code=status.HTTP_201_CREATED, tags=["Notifications"])
async def api_create_notification(
    req: NotificationCreateRequest,
    current_user: Dict[str, Any] = Depends(require_roles(UserRole.ADMIN, UserRole.TEACHER)),
):
    async with db.transaction() as session:
        target_uids: set[int] = set(req.recipient_user_ids or [])
        if req.target_role:
            r_users = await session.fetchall(
                "SELECT id FROM users WHERE role = ? AND is_active = 1 AND is_deleted = 0",
                (req.target_role.value,),
            )
            for u in r_users:
                target_uids.add(int(u["id"]))
        if req.target_class_id:
            if current_user["role"] == UserRole.TEACHER.value:
                tr = await session.fetchone("SELECT id FROM teachers WHERE user_id = ?", (current_user["id"],))
                cl = await session.fetchone("SELECT teacher_id FROM classes WHERE id = ?", (req.target_class_id,))
                if not tr or not cl or int(cl["teacher_id"]) != int(tr["id"]):
                    raise HTTPException(status_code=403, detail="لا يمكنك إرسال إعلانات لقسم غير مسند إليك")
            enr_rows = await session.fetchall(
                """
                SELECT s.user_id AS student_uid, p.user_id AS parent_uid
                FROM enrollments e
                JOIN students s ON s.id = e.student_id
                LEFT JOIN parents p ON p.id = s.primary_parent_id
                WHERE e.class_id = ? AND e.status = 'ACTIVE'
                """,
                (req.target_class_id,),
            )
            for er in enr_rows:
                if er.get("student_uid"):
                    target_uids.add(int(er["student_uid"]))
                if er.get("parent_uid"):
                    target_uids.add(int(er["parent_uid"]))

        if not target_uids and current_user["role"] == UserRole.ADMIN.value:
            all_u = await session.fetchall("SELECT id FROM users WHERE is_active = 1 AND is_deleted = 0")
            for u in all_u:
                target_uids.add(int(u["id"]))

        created_ids = []
        title_en = req.title_en or req.title_ar
        body_en = req.body_en or req.body_ar
        for uid in target_uids:
            nid = await create_and_dispatch_notification(
                session=session,
                recipient_user_id=uid,
                notification_type=req.notification_type.value,
                title_ar=req.title_ar,
                title_en=title_en,
                body_ar=req.body_ar,
                body_en=body_en,
                sender_user_id=current_user["id"],
            )
            if nid:
                created_ids.append(nid)

        await log_audit(
            session, current_user["id"], current_user["username"], current_user["role"],
            "SEND_NOTIFICATIONS", "notifications", None,
            {"count": len(created_ids), "type": req.notification_type.value},
        )
    return {"success": True, "sent_count": len(created_ids), "notification_ids": created_ids}


@app.put("/notifications/{id}/read", tags=["Notifications"])
async def api_mark_notification_read(
    id: int,
    current_user: Dict[str, Any] = Depends(get_current_user),
):
    async with db.transaction() as session:
        notif = await session.fetchone(
            "SELECT * FROM notifications WHERE id = ? AND recipient_user_id = ?",
            (id, current_user["id"]),
        )
        if not notif:
            raise HTTPException(status_code=404, detail="الإشعار غير موجود")
        await session.execute(
            "UPDATE notifications SET is_read = 1, read_at = datetime('now') WHERE id = ?",
            (id,),
        )
    return {"success": True, "message": "تم تحديد الإشعار كمقروء"}


# ------------------------------------------------------------------------------
# 10.9 Reports & Export Endpoints (/reports/*) — JSON, PDF, and Excel (.xlsx)
# ------------------------------------------------------------------------------

def _build_export_response(
    export_format: str,
    filename_stem: str,
    title_ar: str,
    subtitle_ar: str,
    headers: List[str],
    rows: List[List[Any]],
    summary_pairs: List[Tuple[str, Any]],
    json_payload: Dict[str, Any],
):
    if export_format == "excel":
        xlsx_bytes = generate_excel_workbook_bytes(title_ar, subtitle_ar, headers, rows, summary_pairs)
        return Response(
            content=xlsx_bytes,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f'attachment; filename="{filename_stem}.xlsx"'},
        )
    if export_format == "pdf":
        pdf_bytes = generate_pdf_report_bytes(title_ar, subtitle_ar, headers, rows, summary_pairs)
        return Response(
            content=pdf_bytes,
            media_type="application/pdf",
            headers={"Content-Disposition": f'attachment; filename="{filename_stem}.pdf"'},
        )
    return {"success": True, **json_payload}


@app.get("/reports/students", tags=["Reports"])
async def api_report_students(
    level_id: Optional[int] = None,
    export: str = Query(default="json"),
    current_user: Dict[str, Any] = Depends(require_roles(UserRole.ADMIN, UserRole.TEACHER)),
):
    sql = """
        SELECT s.student_code, s.full_name, s.gender, s.birth_date,
               COALESCE(l.name_ar, '-') AS level_name_ar,
               COALESCE(p.full_name, '-') AS parent_name,
               COALESCE(s.phone, p.phone, '-') AS contact_phone,
               s.joined_date
        FROM students s
        LEFT JOIN levels l ON l.id = s.level_id
        LEFT JOIN parents p ON p.id = s.primary_parent_id
        WHERE s.is_archived = 0
    """
    params: List[Any] = []
    if level_id is not None:
        sql += " AND s.level_id = ?"
        params.append(level_id)
    sql += " ORDER BY s.id DESC"
    items = await db.fetchall(sql, params)
    headers = ["الرقم التعريفي", "الاسم الكامل", "الجنس", "تاريخ الميلاد", "المستوى", "ولي الأمر", "الهاتف", "تاريخ الانضمام"]
    rows = [
        [r["student_code"], r["full_name"], r["gender"], r["birth_date"] or "-", r["level_name_ar"], r["parent_name"], r["contact_phone"], r["joined_date"]]
        for r in items
    ]
    summary = [("إجمالي التلاميذ النشطين", len(items))]
    return _build_export_response(
        export, "mounix_students_report", "تقرير قائمة التلاميذ", "قائمة التلاميذ المسجلين",
        headers, rows, summary, {"total": len(items), "items": items},
    )


@app.get("/reports/attendance", tags=["Reports"])
async def api_report_attendance(
    class_id: Optional[int] = None,
    from_date: Optional[str] = None,
    to_date: Optional[str] = None,
    export: str = Query(default="json"),
    current_user: Dict[str, Any] = Depends(require_roles(UserRole.ADMIN, UserRole.TEACHER)),
):
    sql = """
        SELECT a.session_date, s.student_code, s.full_name AS student_name,
               c.name AS class_name, sub.name_ar AS subject_name_ar,
               a.status, a.late_minutes, COALESCE(a.absence_reason, '-') AS absence_reason
        FROM attendance a
        JOIN students s ON s.id = a.student_id
        JOIN classes c ON c.id = a.class_id
        JOIN subjects sub ON sub.id = c.subject_id
        WHERE 1=1
    """
    params: List[Any] = []
    if class_id is not None:
        sql += " AND a.class_id = ?"
        params.append(class_id)
    if from_date:
        sql += " AND a.session_date >= ?"
        params.append(from_date)
    if to_date:
        sql += " AND a.session_date <= ?"
        params.append(to_date)
    sql += " ORDER BY a.session_date DESC, a.id DESC"
    items = await db.fetchall(sql, params)

    total = len(items)
    present_cnt = sum(1 for r in items if r["status"] == "PRESENT")
    absent_cnt = sum(1 for r in items if r["status"] == "ABSENT")
    late_cnt = sum(1 for r in items if r["status"] == "LATE")
    excused_cnt = sum(1 for r in items if r["status"] == "EXCUSED")
    rate = round(((present_cnt + late_cnt) / total) * 100.0, 1) if total > 0 else 100.0

    headers = ["التاريخ", "رمز التلميذ", "التلميذ", "القسم", "المادة", "الحالة", "التأخر (د)", "السبب / الملاحظة"]
    rows = [
        [r["session_date"], r["student_code"], r["student_name"], r["class_name"], r["subject_name_ar"], r["status"], r["late_minutes"], r["absence_reason"]]
        for r in items
    ]
    summary = [
        ("إجمالي السجلات", total),
        ("الحضور", present_cnt),
        ("الغياب", absent_cnt),
        ("التأخر", late_cnt),
        ("نسبة الحضور", f"{rate}%"),
    ]
    return _build_export_response(
        export, "mounix_attendance_report", "تقرير الحضور والغياب", "إحصائيات وسجلات المواظبة",
        headers, rows, summary,
        {
            "summary": {
                "total_records": total,
                "present": present_cnt,
                "absent": absent_cnt,
                "late": late_cnt,
                "excused": excused_cnt,
                "attendance_rate_percent": rate,
            },
            "items": items,
        },
    )


@app.get("/reports/grades", tags=["Reports"])
async def api_report_grades(
    class_id: Optional[int] = None,
    exam_id: Optional[int] = None,
    export: str = Query(default="json"),
    current_user: Dict[str, Any] = Depends(require_roles(UserRole.ADMIN, UserRole.TEACHER)),
):
    sql = """
        SELECT s.student_code, s.full_name AS student_name, c.name AS class_name,
               sub.name_ar AS subject_name_ar, ex.title AS exam_title, ex.exam_date,
               g.score, ex.max_score, COALESCE(g.teacher_remarks, '-') AS remarks
        FROM grades g
        JOIN students s ON s.id = g.student_id
        JOIN exams ex ON ex.id = g.exam_id
        JOIN classes c ON c.id = ex.class_id
        JOIN subjects sub ON sub.id = ex.subject_id
        WHERE 1=1
    """
    params: List[Any] = []
    if class_id is not None:
        sql += " AND ex.class_id = ?"
        params.append(class_id)
    if exam_id is not None:
        sql += " AND g.exam_id = ?"
        params.append(exam_id)
    sql += " ORDER BY g.score DESC"
    items = await db.fetchall(sql, params)

    avg_score = round(sum(float(r["score"]) for r in items) / len(items), 2) if items else 0.0
    headers = ["الرتبة", "رمز التلميذ", "التلميذ", "القسم", "المادة", "الاختبار", "العلامة", "العلامة القصوى", "ملاحظات الأستاذ"]
    rows = [
        [idx + 1, r["student_code"], r["student_name"], r["class_name"], r["subject_name_ar"], r["exam_title"], r["score"], r["max_score"], r["remarks"]]
        for idx, r in enumerate(items)
    ]
    summary = [("عدد النتائج", len(items)), ("متوسط العلامات", avg_score)]
    return _build_export_response(
        export, "mounix_grades_report", "تقرير الاختبارات والنتائج", "ترتيب ومعدلات التلاميذ",
        headers, rows, summary,
        {"summary": {"count": len(items), "average_score": avg_score}, "items": items},
    )


@app.get("/reports/finance", tags=["Reports", "Payments & Finance"])
async def api_report_finance(
    from_date: Optional[str] = None,
    to_date: Optional[str] = None,
    export: str = Query(default="json"),
    current_user: Dict[str, Any] = Depends(require_roles(UserRole.ADMIN)),
):
    today_iso = datetime.date.today().isoformat()
    pay_sql = """
        SELECT p.receipt_number, s.student_code, s.full_name AS student_name,
               p.amount, p.payment_method, p.payment_date, p.recorded_by_name,
               COALESCE(p.description, '-') AS description
        FROM payments p
        JOIN students s ON s.id = p.student_id
        WHERE p.is_voided = 0
    """
    pay_params: List[Any] = []
    if from_date:
        pay_sql += " AND date(p.payment_date) >= ?"
        pay_params.append(from_date)
    if to_date:
        pay_sql += " AND date(p.payment_date) <= ?"
        pay_params.append(to_date)
    pay_sql += " ORDER BY p.payment_date DESC"
    payments = await db.fetchall(pay_sql, pay_params)

    plans = await db.fetchall(
        """
        SELECT pp.*, s.full_name AS student_name, s.student_code
        FROM payment_plans pp
        JOIN students s ON s.id = pp.student_id
        WHERE pp.status != 'VOIDED'
        ORDER BY pp.due_date ASC
        """
    )

    # Strictly only actually recorded payments count as collected revenue
    collected_revenue = round(sum(float(p["amount"]) for p in payments), 2)
    total_billed = round(sum(float(pl["total_amount"]) for pl in plans), 2)
    outstanding_receivables = round(sum(float(pl["remaining_amount"]) for pl in plans if pl["remaining_amount"] > 0), 2)
    overdue_plans = [
        pl for pl in plans
        if float(pl["remaining_amount"]) > 0 and str(pl["due_date"]) < today_iso
    ]
    overdue_amount = round(sum(float(pl["remaining_amount"]) for pl in overdue_plans), 2)

    headers = ["رقم الإيصال", "رمز التلميذ", "اسم التلميذ", "المبلغ المحصل", "طريقة الدفع", "تاريخ الدفع", "الموظف المستلم", "البيان"]
    rows = [
        [p["receipt_number"], p["student_code"], p["student_name"], p["amount"], p["payment_method"], p["payment_date"], p["recorded_by_name"], p["description"]]
        for p in payments
    ]
    summary = [
        ("إجمالي المداخيل المحصلة فعلياً", f"{collected_revenue:.2f} DZD"),
        ("إجمالي الاشتراكات المفوترة", f"{total_billed:.2f} DZD"),
        ("إجمالي الديون المستحقة (غير محصلة)", f"{outstanding_receivables:.2f} DZD"),
        ("المدفوعات المتأخرة", f"{overdue_amount:.2f} DZD ({len(overdue_plans)} اشتراك)"),
    ]
    return _build_export_response(
        export, "mounix_finance_report", "التقرير المالي الشامل", "المداخيل المحصلة، المستحقات، والديون المتأخرة",
        headers, rows, summary,
        {
            "summary": {
                "collected_revenue": collected_revenue,
                "total_billed": total_billed,
                "outstanding_receivables": outstanding_receivables,
                "overdue_amount": overdue_amount,
                "overdue_plans_count": len(overdue_plans),
                "transactions_count": len(payments),
            },
            "payments": payments,
            "overdue_plans": overdue_plans,
        },
    )


# ------------------------------------------------------------------------------
# 10.10 Admin Settings, Audit Logs & Backup Endpoints (/admin/*)
# ------------------------------------------------------------------------------

@app.get("/admin/settings", tags=["System & Backup"])
async def api_get_settings(current_user: Dict[str, Any] = Depends(get_current_user)):
    rows = await db.fetchall("SELECT * FROM app_settings ORDER BY key ASC")
    return {"success": True, "settings": {r["key"]: r["value"] for r in rows}, "items": rows}


@app.put("/admin/settings", tags=["System & Backup"])
async def api_update_settings(
    updates: Dict[str, str] = Body(...),
    current_user: Dict[str, Any] = Depends(require_roles(UserRole.ADMIN)),
):
    async with db.transaction() as session:
        for k, v in updates.items():
            await session.execute(
                """
                INSERT INTO app_settings (key, value, updated_by, updated_at)
                VALUES (?, ?, ?, datetime('now'))
                ON CONFLICT(key) DO UPDATE SET value = excluded.value, updated_by = excluded.updated_by, updated_at = datetime('now')
                """,
                (k, str(v), current_user["id"]),
            )
        await log_audit(
            session, current_user["id"], current_user["username"], current_user["role"],
            "UPDATE_SETTINGS", "app_settings", None, updates,
        )
    return {"success": True, "message": "تم حفظ إعدادات النظام بنجاح"}


@app.get("/admin/audit-logs", tags=["System & Backup"])
async def api_list_audit_logs(
    limit: int = Query(default=100, ge=1, le=500),
    current_user: Dict[str, Any] = Depends(require_roles(UserRole.ADMIN)),
):
    rows = await db.fetchall("SELECT * FROM audit_logs ORDER BY id DESC LIMIT ?", (limit,))
    return {"success": True, "items": rows}


@app.post("/admin/backup", status_code=status.HTTP_201_CREATED, tags=["System & Backup"])
async def api_create_backup(current_user: Dict[str, Any] = Depends(require_roles(UserRole.ADMIN))):
    b_row = await create_database_backup(current_user["id"], current_user["username"])
    return {"success": True, "backup": b_row}


@app.get("/admin/backups", tags=["System & Backup"])
async def api_list_backups(current_user: Dict[str, Any] = Depends(require_roles(UserRole.ADMIN))):
    rows = await db.fetchall("SELECT * FROM backup_logs ORDER BY id DESC")
    return {"success": True, "items": rows}


@app.post("/admin/restore", tags=["System & Backup"])
async def api_restore_backup(
    req: RestoreBackupRequest,
    current_user: Dict[str, Any] = Depends(require_roles(UserRole.ADMIN)),
):
    target = req.backup_id if req.backup_id is not None else req.filename
    if not target:
        raise HTTPException(status_code=400, detail="يرجى تحديد معرف النسخة الاحتياطية (backup_id) أو اسم الملف (filename)")
    try:
        res = await restore_database_backup(target, current_user["id"], current_user["username"])
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    return {"success": True, **res}


# ==============================================================================
# 11. TELEGRAM BOT (AIOGRAM 3) — 16-MENU ADMIN PANEL, FSM WIZARDS & ROLE PORTALS
# ==============================================================================

class BotFSM(StatesGroup):
    # Auth / Account linking
    waiting_login_username = State()
    waiting_login_password = State()

    # Search
    waiting_student_search = State()
    waiting_teacher_search = State()

    # Add Student Wizard
    add_student_name = State()
    add_student_phone = State()
    add_student_birth = State()
    add_student_level = State()

    # Add Teacher Wizard
    add_teacher_name = State()
    add_teacher_phone = State()
    add_teacher_spec = State()

    # Add Parent Wizard
    add_parent_name = State()
    add_parent_phone = State()
    add_parent_student_id = State()

    # Add Level / Subject Wizard
    add_level_name_ar = State()
    add_level_name_en = State()
    add_subject_name_ar = State()
    add_subject_name_en = State()
    add_subject_fee = State()

    # Add Class Wizard
    add_class_name = State()
    add_class_level = State()
    add_class_subject = State()
    add_class_teacher = State()
    add_class_room = State()
    add_class_capacity = State()

    # Add Schedule Wizard
    add_schedule_class = State()
    add_schedule_day = State()
    add_schedule_start = State()
    add_schedule_end = State()

    # Record Attendance Wizard
    att_select_schedule = State()
    att_select_student = State()
    att_select_status = State()

    # Create Exam & Grade Wizard
    add_exam_title = State()
    add_exam_class = State()
    add_exam_max_score = State()
    add_grade_exam_id = State()
    add_grade_student_id = State()
    add_grade_score = State()

    # Record Payment & Plan Wizard
    add_plan_student_id = State()
    add_plan_type = State()
    add_plan_amount = State()
    add_payment_student_id = State()
    add_payment_amount = State()
    add_payment_method = State()

    # Add Material / Assignment Wizard
    add_material_title = State()
    add_material_class = State()
    add_material_url = State()

    # Broadcast Notification Wizard
    broadcast_title = State()
    broadcast_body = State()

    # Settings Update Wizard
    setting_key = State()
    setting_val = State()

    # Add User Wizard
    add_user_username = State()
    add_user_password = State()
    add_user_fullname = State()
    add_user_role = State()


DAYS_AR = ["الأحد", "الإثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت"]


def build_admin_main_keyboard() -> InlineKeyboardMarkup:
    """
    Constructs the exact 16-section Telegram Admin Dashboard required by Section 14.
    """
    rows = [
        [
            InlineKeyboardButton(text="1. 📊 الإحصائيات", callback_data="menu:stats"),
            InlineKeyboardButton(text="2. 👨‍🎓 إدارة التلاميذ", callback_data="menu:students:1"),
        ],
        [
            InlineKeyboardButton(text="3. 👨‍🏫 إدارة الأساتذة", callback_data="menu:teachers:1"),
            InlineKeyboardButton(text="4. 👨‍👩‍👧 أولياء الأمور", callback_data="menu:parents:1"),
        ],
        [
            InlineKeyboardButton(text="5. 📚 المواد والمستويات", callback_data="menu:subjects"),
            InlineKeyboardButton(text="6. 🏫 الأقسام والمجموعات", callback_data="menu:classes:1"),
        ],
        [
            InlineKeyboardButton(text="7. 📅 الجداول والحصص", callback_data="menu:schedules:1"),
            InlineKeyboardButton(text="8. 📝 الحضور والغياب", callback_data="menu:attendance"),
        ],
        [
            InlineKeyboardButton(text="9. 📈 الاختبارات والنتائج", callback_data="menu:grades"),
            InlineKeyboardButton(text="10. 💰 الإدارة المالية", callback_data="menu:finance"),
        ],
        [
            InlineKeyboardButton(text="11. 📂 الملفات التعليمية", callback_data="menu:materials:1"),
            InlineKeyboardButton(text="12. 📢 الإشعارات والإعلانات", callback_data="menu:notifications"),
        ],
        [
            InlineKeyboardButton(text="13. 📄 التقارير (PDF/Excel)", callback_data="menu:reports"),
            InlineKeyboardButton(text="14. ⚙️ إعدادات النظام", callback_data="menu:settings"),
        ],
        [
            InlineKeyboardButton(text="15. 💾 النسخ الاحتياطي", callback_data="menu:backups"),
            InlineKeyboardButton(text="16. 🛡️ المستخدمون والصلاحيات", callback_data="menu:users:1"),
        ],
    ]
    return InlineKeyboardMarkup(inline_keyboard=rows)


def build_role_keyboard(role: str) -> InlineKeyboardMarkup:
    if role == UserRole.ADMIN.value:
        return build_admin_main_keyboard()
    if role == UserRole.TEACHER.value:
        return InlineKeyboardMarkup(
            inline_keyboard=[
                [
                    InlineKeyboardButton(text="👤 ملفي الشخصي", callback_data="role:profile"),
                    InlineKeyboardButton(text="🏫 أقسامي وموادي", callback_data="menu:classes:1"),
                ],
                [
                    InlineKeyboardButton(text="📅 جدول حصصي", callback_data="menu:schedules:1"),
                    InlineKeyboardButton(text="📝 تسجيل الحضور", callback_data="menu:attendance"),
                ],
                [
                    InlineKeyboardButton(text="📈 الاختبارات والعلامات", callback_data="menu:grades"),
                    InlineKeyboardButton(text="📂 الملفات والواجبات", callback_data="menu:materials:1"),
                ],
                [
                    InlineKeyboardButton(text="📢 إرسال إعلان للقسم", callback_data="menu:notifications"),
                ],
            ]
        )
    if role == UserRole.STUDENT.value:
        return InlineKeyboardMarkup(
            inline_keyboard=[
                [
                    InlineKeyboardButton(text="🎓 ملفي الدراسي الشامل", callback_data="role:student_dossier"),
                    InlineKeyboardButton(text="📅 جدولي الدراسي", callback_data="menu:schedules:1"),
                ],
                [
                    InlineKeyboardButton(text="📝 سجل حضوري", callback_data="role:my_attendance"),
                    InlineKeyboardButton(text="📈 نتائجي ومعدلي", callback_data="role:my_grades"),
                ],
                [
                    InlineKeyboardButton(text="📂 الدروس والواجبات", callback_data="menu:materials:1"),
                    InlineKeyboardButton(text="🔔 إشعاراتي", callback_data="role:my_notifications"),
                ],
            ]
        )
    # PARENT
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [
                InlineKeyboardButton(text="👨‍👧‍👦 ملفات أبنائي", callback_data="role:parent_children"),
                InlineKeyboardButton(text="📅 جداول الأبناء", callback_data="menu:schedules:1"),
            ],
            [
                InlineKeyboardButton(text="📝 متابعة الحضور والغياب", callback_data="role:my_attendance"),
                InlineKeyboardButton(text="📈 نتائج الأبناء", callback_data="role:my_grades"),
            ],
            [
                InlineKeyboardButton(text="💰 الاشتراكات والمدفوعات", callback_data="role:my_finance"),
                InlineKeyboardButton(text="🔔 التنبيهات والإعلانات", callback_data="role:my_notifications"),
            ],
        ]
    )


def nav_footer_buttons(back_cb: str = "menu:home") -> List[InlineKeyboardButton]:
    return [
        InlineKeyboardButton(text="🔙 رجوع", callback_data=back_cb),
        InlineKeyboardButton(text="🏠 الرئيسية", callback_data="menu:home"),
        InlineKeyboardButton(text="❌ إلغاء", callback_data="action:cancel"),
    ]


async def get_user_by_telegram_id(telegram_id: int) -> Optional[Dict[str, Any]]:
    return await db.fetchone(
        "SELECT * FROM users WHERE telegram_id = ? AND is_deleted = 0 AND is_active = 1",
        (telegram_id,),
    )


async def process_telegram_callback_logic(
    telegram_id: int,
    data: str,
    fsm: FSMContext,
) -> Tuple[str, InlineKeyboardMarkup, Optional[Tuple[bytes, str]]]:
    """
    Core Telegram Bot callback processor used by both Aiogram 3 Router and
    direct integration tests. Returns (message_text, InlineKeyboardMarkup, optional_file_tuple).
    """
    user = await get_user_by_telegram_id(telegram_id)
    if data == "action:cancel":
        await fsm.clear()
        if not user:
            return (
                "❌ تم إلغاء العملية. أرسل /start لتسجيل الدخول.",
                InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text="🔐 تسجيل الدخول", callback_data="auth:login")]]),
                None,
            )
        return ("✅ تم إلغاء العملية والعودة إلى القائمة الرئيسية.", build_role_keyboard(user["role"]), None)

    if data == "auth:login":
        await fsm.set_state(BotFSM.waiting_login_username)
        return (
            "🔐 *تسجيل الدخول إلى نظام MOUNIX SCHOOL*\n\nيرجى إرسال *اسم المستخدم (Username)* الخاص بك:",
            InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text="❌ إلغاء", callback_data="action:cancel")]]),
            None,
        )

    if not user:
        return (
            "🔒 *غير مصرح لك بالوصول*\nيرجى ربط حسابك بتسجيل الدخول أولاً.",
            InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text="🔐 تسجيل الدخول", callback_data="auth:login")]]),
            None,
        )

    role = user["role"]

    if data == "menu:home":
        await fsm.clear()
        return (
            f"🏫 *MOUNIX SCHOOL — لوحة التحكم الذكية*\n"
            f"مرحباً بك: *{user['full_name']}* (`{role}`)\n"
            f"اختر القسم المطلوب من القائمة أدناه:",
            build_role_keyboard(role),
            None,
        )

    # Guard admin-only sections
    admin_only_prefixes = (
        "menu:stats", "menu:students", "menu:teachers", "menu:parents",
        "menu:subjects", "menu:finance", "menu:reports", "menu:settings",
        "menu:backups", "menu:users", "act:stu:", "act:tch:", "act:par:",
        "act:sub:", "act:lvl:", "act:cls:", "act:sch:", "act:fin:", "act:bak:", "act:usr:", "confirm:",
    )
    if data.startswith(admin_only_prefixes) and role != UserRole.ADMIN.value:
        return (
            "⛔ *عذراً، هذه اللوحة مخصصة لمدير النظام (ADMIN) فقط.*",
            build_role_keyboard(role),
            None,
        )

    # 1. 📊 الإحصائيات
    if data == "menu:stats":
        stu_c = (await db.fetchone("SELECT COUNT(*) AS c FROM students WHERE is_archived = 0"))["c"]
        tch_c = (await db.fetchone("SELECT COUNT(*) AS c FROM teachers WHERE is_archived = 0"))["c"]
        par_c = (await db.fetchone("SELECT COUNT(*) AS c FROM parents WHERE is_archived = 0"))["c"]
        cls_c = (await db.fetchone("SELECT COUNT(*) AS c FROM classes WHERE is_archived = 0"))["c"]
        sub_c = (await db.fetchone("SELECT COUNT(*) AS c FROM subjects WHERE is_archived = 0"))["c"]
        rev = (await db.fetchone("SELECT COALESCE(SUM(amount), 0) AS s FROM payments WHERE is_voided = 0"))["s"]
        debt = (await db.fetchone("SELECT COALESCE(SUM(remaining_amount), 0) AS s FROM payment_plans WHERE status != 'VOIDED'"))["s"]
        att = await db.fetchone("SELECT COUNT(*) AS t, SUM(CASE WHEN status IN ('PRESENT','LATE') THEN 1 ELSE 0 END) AS p FROM attendance")
        rate = round((float(att["p"] or 0) / float(att["t"])) * 100.0, 1) if att and att["t"] else 100.0

        text = (
            "📊 *الإحصائيات العامة لمدرسة MOUNIX SCHOOL*\n\n"
            f"👨‍🎓 التلاميذ النشطون: *{stu_c}*\n"
            f"👨‍🏫 الأساتذة النشطون: *{tch_c}*\n"
            f"👨‍👩‍👧 أولياء الأمور: *{par_c}*\n"
            f"🏫 الأقسام والمجموعات: *{cls_c}*\n"
            f"📚 المواد الدراسية: *{sub_c}*\n"
            f"📝 نسبة الحضور العامة: *{rate}%*\n"
            f"💰 إجمالي المداخيل المحصلة: *{float(rev):.2f} DZD*\n"
            f"⚠️ إجمالي الديون المستحقة: *{float(debt):.2f} DZD*"
        )
        kb = InlineKeyboardMarkup(inline_keyboard=[nav_footer_buttons("menu:home")])
        return (text, kb, None)

    # 2. 👨‍🎓 إدارة التلاميذ (With Pagination & Search)
    if data.startswith("menu:students:"):
        page = int(data.split(":")[2])
        page_size = 5
        rows = await db.fetchall(
            """
            SELECT s.id, s.student_code, s.full_name, COALESCE(l.name_ar, '-') AS level_name
            FROM students s LEFT JOIN levels l ON l.id = s.level_id
            WHERE s.is_archived = 0 ORDER BY s.id DESC
            """
        )
        total = len(rows)
        total_pages = max(1, (total + page_size - 1) // page_size)
        page = max(1, min(page, total_pages))
        slice_rows = rows[(page - 1) * page_size : page * page_size]

        lines = [f"👨‍🎓 *إدارة التلاميذ* (الإجمالي: {total} | صفحة {page}/{total_pages})\n"]
        btns: List[List[InlineKeyboardButton]] = []
        for r in slice_rows:
            lines.append(f"• `{r['student_code']}` — *{r['full_name']}* ({r['level_name']})")
            btns.append([
                InlineKeyboardButton(text=f"📄 ملف: {r['full_name']}", callback_data=f"act:stu:view:{r['id']}"),
                InlineKeyboardButton(text="🗄 أرشفة", callback_data=f"confirm:stu_arch:{r['id']}"),
            ])

        nav_row: List[InlineKeyboardButton] = []
        if page > 1:
            nav_row.append(InlineKeyboardButton(text="⬅️ السابق", callback_data=f"menu:students:{page - 1}"))
        if page < total_pages:
            nav_row.append(InlineKeyboardButton(text="التالي ➡️", callback_data=f"menu:students:{page + 1}"))
        if nav_row:
            btns.append(nav_row)

        btns.append([
            InlineKeyboardButton(text="➕ إضافة تلميذ جديد", callback_data="act:stu:add"),
            InlineKeyboardButton(text="🔍 بحث عن تلميذ", callback_data="act:stu:search"),
        ])
        btns.append(nav_footer_buttons("menu:home"))
        return ("\n".join(lines), InlineKeyboardMarkup(inline_keyboard=btns), None)

    if data == "act:stu:add":
        await fsm.set_state(BotFSM.add_student_name)
        return (
            "➕ *إضافة تلميذ جديد (خطوة 1/4)*\nيرجى إدخال *الاسم الكامل للتلميذ*:",
            InlineKeyboardMarkup(inline_keyboard=[nav_footer_buttons("menu:students:1")]),
            None,
        )

    if data == "act:stu:search":
        await fsm.set_state(BotFSM.waiting_student_search)
        return (
            "🔍 *البحث عن تلميذ*\nأرسل اسم التلميذ أو رقمه التعريفي (`STU-...`):",
            InlineKeyboardMarkup(inline_keyboard=[nav_footer_buttons("menu:students:1")]),
            None,
        )

    if data.startswith("act:stu:view:"):
        sid = int(data.split(":")[3])
        dossier = await get_student_full_dossier(sid)
        if not dossier:
            return ("❌ التلميذ غير موجود.", build_admin_main_keyboard(), None)
        st = dossier["student"]
        att_s = dossier["attendance"]["summary"]
        fin_s = dossier["finance"]
        grd_s = dossier["grades"]
        text = (
            f"🎓 *بطاقة التلميذ الشاملة*\n\n"
            f"• الاسم: *{st['full_name']}*\n"
            f"• الرمز التعريفي: `{st['student_code']}`\n"
            f"• المستوى: {st.get('level_name_ar') or '-'}\n"
            f"• ولي الأمر: {st.get('parent_name') or '-'} ({st.get('parent_phone') or '-'})\n"
            f"• الهاتف: {st.get('phone') or '-'}\n"
            f"• تاريخ الانضمام: {st.get('joined_date')}\n"
            f"• الأقسام المسجل بها: {len(dossier['classes'])}\n"
            f"• نسبة الحضور: *{att_s['attendance_rate_percent']}%* (حضور {att_s['present_count']} / غياب {att_s['absent_count']})\n"
            f"• المعدل العام: *{grd_s['average_on_20']}/20*\n"
            f"• إجمالي المدفوعات: *{fin_s['total_paid']:.2f} DZD*\n"
            f"• الديون المتبقية: *{fin_s['total_remaining_debt']:.2f} DZD*"
        )
        kb = InlineKeyboardMarkup(
            inline_keyboard=[
                [InlineKeyboardButton(text="🗄 أرشفة التلميذ", callback_data=f"confirm:stu_arch:{sid}")],
                nav_footer_buttons("menu:students:1"),
            ]
        )
        return (text, kb, None)

    if data.startswith("confirm:stu_arch:"):
        sid = int(data.split(":")[2])
        kb = InlineKeyboardMarkup(
            inline_keyboard=[
                [
                    InlineKeyboardButton(text="✅ نعم، تأكيد الأرشفة", callback_data=f"act:stu:arch_do:{sid}"),
                    InlineKeyboardButton(text="❌ تراجع", callback_data=f"act:stu:view:{sid}"),
                ]
            ]
        )
        return (f"⚠️ *تأكيد عملية حساسة*\nهل أنت متأكد من أرشفة التلميذ رقم `#{sid}`؟", kb, None)

    if data.startswith("act:stu:arch_do:"):
        sid = int(data.split(":")[3])
        async with db.transaction() as session:
            await session.execute("UPDATE students SET is_archived = 1, updated_at = datetime('now') WHERE id = ?", (sid,))
            await log_audit(session, user["id"], user["username"], role, "BOT_ARCHIVE_STUDENT", "students", sid)
        return ("✅ تم أرشفة التلميذ بنجاح.", build_admin_main_keyboard(), None)

    # 3. 👨‍🏫 إدارة الأساتذة
    if data.startswith("menu:teachers:"):
        page = int(data.split(":")[2])
        rows = await db.fetchall("SELECT * FROM teachers WHERE is_archived = 0 ORDER BY id DESC")
        page_size = 5
        total = len(rows)
        total_pages = max(1, (total + page_size - 1) // page_size)
        page = max(1, min(page, total_pages))
        slice_rows = rows[(page - 1) * page_size : page * page_size]

        lines = [f"👨‍🏫 *إدارة الأساتذة* (الإجمالي: {total} | صفحة {page}/{total_pages})\n"]
        btns: List[List[InlineKeyboardButton]] = []
        for r in slice_rows:
            lines.append(f"• `{r['teacher_code']}` — *{r['full_name']}* ({r['specialization']})")
            btns.append([
                InlineKeyboardButton(text=f"👤 {r['full_name']}", callback_data=f"act:tch:view:{r['id']}"),
                InlineKeyboardButton(text="🗄 أرشفة", callback_data=f"act:tch:arch:{r['id']}"),
            ])
        btns.append([
            InlineKeyboardButton(text="➕ إضافة أستاذ", callback_data="act:tch:add"),
            InlineKeyboardButton(text="🔍 بحث عن أستاذ", callback_data="act:tch:search"),
        ])
        btns.append(nav_footer_buttons("menu:home"))
        return ("\n".join(lines), InlineKeyboardMarkup(inline_keyboard=btns), None)

    if data == "act:tch:add":
        await fsm.set_state(BotFSM.add_teacher_name)
        return (
            "➕ *إضافة أستاذ جديد (1/3)*\nأرسل *الاسم الكامل للأستاذ*:",
            InlineKeyboardMarkup(inline_keyboard=[nav_footer_buttons("menu:teachers:1")]),
            None,
        )

    if data == "act:tch:search":
        await fsm.set_state(BotFSM.waiting_teacher_search)
        return (
            "🔍 *البحث عن أستاذ*\nأرسل اسم الأستاذ أو التخصص أو الرمز:",
            InlineKeyboardMarkup(inline_keyboard=[nav_footer_buttons("menu:teachers:1")]),
            None,
        )

    if data.startswith("act:tch:view:"):
        tid = int(data.split(":")[3])
        t_row = await db.fetchone("SELECT * FROM teachers WHERE id = ?", (tid,))
        if not t_row:
            return ("❌ الأستاذ غير موجود.", build_admin_main_keyboard(), None)
        cls_rows = await db.fetchall("SELECT name, room FROM classes WHERE teacher_id = ? AND is_archived = 0", (tid,))
        cls_txt = "، ".join(c["name"] for c in cls_rows) if cls_rows else "لا توجد أقسام"
        text = (
            f"👨‍🏫 *ملف الأستاذ*\n\n"
            f"• الاسم: *{t_row['full_name']}*\n"
            f"• الرمز: `{t_row['teacher_code']}`\n"
            f"• التخصص: {t_row['specialization']}\n"
            f"• الهاتف: {t_row['phone']}\n"
            f"• الأقسام المسندة: {cls_txt}"
        )
        return (text, InlineKeyboardMarkup(inline_keyboard=[nav_footer_buttons("menu:teachers:1")]), None)

    if data.startswith("act:tch:arch:"):
        tid = int(data.split(":")[3])
        async with db.transaction() as session:
            await session.execute("UPDATE teachers SET is_archived = 1, updated_at = datetime('now') WHERE id = ?", (tid,))
            await log_audit(session, user["id"], user["username"], role, "BOT_ARCHIVE_TEACHER", "teachers", tid)
        return ("✅ تم أرشفة الأستاذ بنجاح.", build_admin_main_keyboard(), None)

    # 4. 👨‍👩‍👧 أولياء الأمور
    if data.startswith("menu:parents:"):
        rows = await db.fetchall("SELECT * FROM parents WHERE is_archived = 0 ORDER BY id DESC LIMIT 10")
        lines = [f"👨‍👩‍👧 *إدارة أولياء الأمور* (أحدث {len(rows)} سجل)\n"]
        for r in rows:
            lines.append(f"• `{r['parent_code']}` — *{r['full_name']}* ({r['phone']})")
        btns = [
            [InlineKeyboardButton(text="➕ إضافة ولي أمر وربطه بتلميذ", callback_data="act:par:add")],
            nav_footer_buttons("menu:home"),
        ]
        return ("\n".join(lines), InlineKeyboardMarkup(inline_keyboard=btns), None)

    if data == "act:par:add":
        await fsm.set_state(BotFSM.add_parent_name)
        return (
            "➕ *إضافة ولي أمر (1/3)*\nأرسل *الاسم الكامل لولي الأمر*:",
            InlineKeyboardMarkup(inline_keyboard=[nav_footer_buttons("menu:parents:1")]),
            None,
        )

    # 5. 📚 المواد والمستويات
    if data == "menu:subjects":
        lvls = await db.fetchall("SELECT * FROM levels WHERE is_archived = 0 ORDER BY sort_order ASC")
        subs = await db.fetchall("SELECT * FROM subjects WHERE is_archived = 0 ORDER BY id DESC")
        lines = ["📚 *المستويات والمواد الدراسية*\n", "*المستويات:*"]
        for l in lvls:
            lines.append(f"  • `#{l['id']}` {l['name_ar']} ({l['name_en']})")
        lines.append("\n*المواد:*")
        for s in subs:
            lines.append(f"  • `#{s['id']}` {s['name_ar']} — اشتراك شهري: {s['default_monthly_fee']} DZD")
        btns = [
            [
                InlineKeyboardButton(text="➕ إضافة مستوى دراسي", callback_data="act:lvl:add"),
                InlineKeyboardButton(text="➕ إضافة مادة دراسية", callback_data="act:sub:add"),
            ],
            nav_footer_buttons("menu:home"),
        ]
        return ("\n".join(lines), InlineKeyboardMarkup(inline_keyboard=btns), None)

    if data == "act:lvl:add":
        await fsm.set_state(BotFSM.add_level_name_ar)
        return (
            "➕ *إضافة مستوى دراسي (1/2)*\nأرسل *اسم المستوى بالعربية* (مثال: السنة الثالثة ثانوي):",
            InlineKeyboardMarkup(inline_keyboard=[nav_footer_buttons("menu:subjects")]),
            None,
        )

    if data == "act:sub:add":
        await fsm.set_state(BotFSM.add_subject_name_ar)
        return (
            "➕ *إضافة مادة دراسية (1/3)*\nأرسل *اسم المادة بالعربية* (مثال: الرياضيات):",
            InlineKeyboardMarkup(inline_keyboard=[nav_footer_buttons("menu:subjects")]),
            None,
        )

    # 6. 🏫 الأقسام والمجموعات
    if data.startswith("menu:classes:"):
        rows = await db.fetchall(
            """
            SELECT c.*, s.name_ar AS subject_name, t.full_name AS teacher_name,
                   (SELECT COUNT(*) FROM enrollments e WHERE e.class_id = c.id AND e.status = 'ACTIVE') AS enrolled
            FROM classes c
            JOIN subjects s ON s.id = c.subject_id
            JOIN teachers t ON t.id = c.teacher_id
            WHERE c.is_archived = 0 ORDER BY c.id DESC LIMIT 12
            """
        )
        lines = ["🏫 *الأقسام والمجموعات الدراسية*\n"]
        for r in rows:
            lines.append(
                f"• `#{r['id']}` *{r['name']}* ({r['subject_name']}) — الأستاذ: {r['teacher_name']} "
                f"| القاعة: {r['room']} | المسجلون: `{r['enrolled']}/{r['capacity']}`"
            )
        btns = []
        if role == UserRole.ADMIN.value:
            btns.append([InlineKeyboardButton(text="➕ إنشاء قسم جديد", callback_data="act:cls:add")])
        btns.append(nav_footer_buttons("menu:home"))
        return ("\n".join(lines), InlineKeyboardMarkup(inline_keyboard=btns), None)

    if data == "act:cls:add":
        await fsm.set_state(BotFSM.add_class_name)
        return (
            "➕ *إنشاء قسم جديد (1/6)*\nأرسل *اسم القسم / المجموعة* (مثال: فوج النخبة رياضيات 1):",
            InlineKeyboardMarkup(inline_keyboard=[nav_footer_buttons("menu:classes:1")]),
            None,
        )

    # 7. 📅 الجداول والحصص
    if data.startswith("menu:schedules:"):
        rows = await db.fetchall(
            """
            SELECT sch.*, c.name AS class_name, sub.name_ar AS subject_name, t.full_name AS teacher_name
            FROM schedules sch
            JOIN classes c ON c.id = sch.class_id
            JOIN subjects sub ON sub.id = sch.subject_id
            JOIN teachers t ON t.id = sch.teacher_id
            WHERE sch.status = 'SCHEDULED'
            ORDER BY sch.day_of_week ASC, sch.start_time ASC LIMIT 15
            """
        )
        lines = ["📅 *جدول الحصص الأسبوعي*\n"]
        for r in rows:
            day_name = DAYS_AR[int(r["day_of_week"]) % 7]
            lines.append(
                f"• `#{r['id']}` *{day_name}* ({r['start_time']} - {r['end_time']}) | "
                f"{r['class_name']} ({r['subject_name']}) | القاعة: {r['room']}"
            )
        btns = []
        if role == UserRole.ADMIN.value:
            btns.append([InlineKeyboardButton(text="➕ برمجة حصة جديدة", callback_data="act:sch:add")])
        btns.append(nav_footer_buttons("menu:home"))
        return ("\n".join(lines), InlineKeyboardMarkup(inline_keyboard=btns), None)

    if data == "act:sch:add":
        await fsm.set_state(BotFSM.add_schedule_class)
        return (
            "➕ *برمجة حصة دراسية (1/4)*\nأرسل *رقم القسم (Class ID)*:",
            InlineKeyboardMarkup(inline_keyboard=[nav_footer_buttons("menu:schedules:1")]),
            None,
        )

    # 8. 📝 الحضور والغياب
    if data == "menu:attendance":
        recent = await db.fetchall(
            """
            SELECT a.*, s.full_name AS student_name, c.name AS class_name
            FROM attendance a
            JOIN students s ON s.id = a.student_id
            JOIN classes c ON c.id = a.class_id
            ORDER BY a.id DESC LIMIT 10
            """
        )
        lines = ["📝 *إدارة الحضور والغياب*\n"]
        for r in recent:
            lines.append(f"• {r['session_date']} | *{r['student_name']}* ({r['class_name']}) — `{r['status']}`")
        btns = [
            [InlineKeyboardButton(text="✍️ تسجيل حضور/غياب تلميذ", callback_data="act:att:add")],
            nav_footer_buttons("menu:home"),
        ]
        return ("\n".join(lines), InlineKeyboardMarkup(inline_keyboard=btns), None)

    if data == "act:att:add":
        await fsm.set_state(BotFSM.att_select_schedule)
        return (
            "📝 *تسجيل الحضور والغياب (1/3)*\nأرسل *رقم الحصة (Schedule ID)*:",
            InlineKeyboardMarkup(inline_keyboard=[nav_footer_buttons("menu:attendance")]),
            None,
        )

    # 9. 📈 الاختبارات والنتائج
    if data == "menu:grades":
        exams = await db.fetchall(
            """
            SELECT ex.*, c.name AS class_name, sub.name_ar AS subject_name
            FROM exams ex
            JOIN classes c ON c.id = ex.class_id
            JOIN subjects sub ON sub.id = ex.subject_id
            ORDER BY ex.id DESC LIMIT 10
            """
        )
        lines = ["📈 *الاختبارات والنتائج*\n"]
        for ex in exams:
            lines.append(f"• اختبار `#{ex['id']}`: *{ex['title']}* ({ex['class_name']}) — العلامة القصوى: `{ex['max_score']}`")
        btns = [
            [
                InlineKeyboardButton(text="➕ إنشاء اختبار جديد", callback_data="act:exm:add"),
                InlineKeyboardButton(text="✍️ رصد علامة تلميذ", callback_data="act:grd:add"),
            ],
            nav_footer_buttons("menu:home"),
        ]
        return ("\n".join(lines), InlineKeyboardMarkup(inline_keyboard=btns), None)

    if data == "act:exm:add":
        await fsm.set_state(BotFSM.add_exam_title)
        return (
            "➕ *إنشاء اختبار جديد (1/3)*\nأرسل *عنوان الاختبار*:",
            InlineKeyboardMarkup(inline_keyboard=[nav_footer_buttons("menu:grades")]),
            None,
        )

    if data == "act:grd:add":
        await fsm.set_state(BotFSM.add_grade_exam_id)
        return (
            "✍️ *رصد علامة تلميذ (1/3)*\nأرسل *رقم الاختبار (Exam ID)*:",
            InlineKeyboardMarkup(inline_keyboard=[nav_footer_buttons("menu:grades")]),
            None,
        )

    # 10. 💰 الإدارة المالية
    if data == "menu:finance":
        rev = (await db.fetchone("SELECT COALESCE(SUM(amount), 0) AS s FROM payments WHERE is_voided = 0"))["s"]
        debt = (await db.fetchone("SELECT COALESCE(SUM(remaining_amount), 0) AS s FROM payment_plans WHERE status != 'VOIDED'"))["s"]
        recent_p = await db.fetchall(
            """
            SELECT p.*, s.full_name AS student_name
            FROM payments p JOIN students s ON s.id = p.student_id
            WHERE p.is_voided = 0 ORDER BY p.id DESC LIMIT 6
            """
        )
        lines = [
            "💰 *الإدارة المالية والاشتراكات*\n",
            f"✅ إجمالي المداخيل المحصلة: *{float(rev):.2f} DZD*",
            f"⚠️ إجمالي الديون المستحقة: *{float(debt):.2f} DZD*\n",
            "*أحدث المدفوعات:*",
        ]
        for p in recent_p:
            lines.append(f"• `{p['receipt_number']}` | *{p['student_name']}* — `{p['amount']:.2f} DZD` ({p['payment_method']})")
        btns = [
            [
                InlineKeyboardButton(text="➕ إنشاء اشتراك لتلميذ", callback_data="act:fin:add_plan"),
                InlineKeyboardButton(text="💵 تسجيل دفعة مالية", callback_data="act:fin:add_pay"),
            ],
            nav_footer_buttons("menu:home"),
        ]
        return ("\n".join(lines), InlineKeyboardMarkup(inline_keyboard=btns), None)

    if data == "act:fin:add_plan":
        await fsm.set_state(BotFSM.add_plan_student_id)
        return (
            "➕ *إنشاء اشتراك مالي (1/3)*\nأرسل *رقم التلميذ (Student ID)*:",
            InlineKeyboardMarkup(inline_keyboard=[nav_footer_buttons("menu:finance")]),
            None,
        )

    if data == "act:fin:add_pay":
        await fsm.set_state(BotFSM.add_payment_student_id)
        return (
            "💵 *تسجيل دفعة مالية (1/3)*\nأرسل *رقم التلميذ (Student ID)*:",
            InlineKeyboardMarkup(inline_keyboard=[nav_footer_buttons("menu:finance")]),
            None,
        )

    # 11. 📂 الملفات التعليمية
    if data.startswith("menu:materials:"):
        mats = await db.fetchall(
            """
            SELECT m.*, sub.name_ar AS subject_name
            FROM materials m JOIN subjects sub ON sub.id = m.subject_id
            WHERE m.is_archived = 0 ORDER BY m.id DESC LIMIT 10
            """
        )
        lines = ["📂 *الملفات التعليمية والواجبات*\n"]
        for m in mats:
            lines.append(f"• `#{m['id']}` *{m['title']}* ({m['subject_name']}) — النوع: `{m['material_type']}`")
        btns = []
        if role in (UserRole.ADMIN.value, UserRole.TEACHER.value):
            btns.append([InlineKeyboardButton(text="➕ إضافة ملف/رابط تعليمي", callback_data="act:mat:add")])
        btns.append(nav_footer_buttons("menu:home"))
        return ("\n".join(lines), InlineKeyboardMarkup(inline_keyboard=btns), None)

    if data == "act:mat:add":
        await fsm.set_state(BotFSM.add_material_title)
        return (
            "➕ *إضافة محتوى تعليمي (1/3)*\nأرسل *عنوان الدرس أو الملف*:",
            InlineKeyboardMarkup(inline_keyboard=[nav_footer_buttons("menu:materials:1")]),
            None,
        )

    # 12. 📢 الإشعارات والإعلانات
    if data == "menu:notifications":
        await fsm.set_state(BotFSM.broadcast_title)
        return (
            "📢 *إرسال إشعار / إعلان عام (1/2)*\nأرسل *عنوان الإعلان*:",
            InlineKeyboardMarkup(inline_keyboard=[nav_footer_buttons("menu:home")]),
            None,
        )

    # 13. 📄 التقارير (PDF & Excel)
    if data == "menu:reports":
        text = "📄 *مركز التقارير والتصدير (PDF & Excel RTL)*\nاختر التقرير والصيغة المطلوبة ليتم توليدها وإرسالها فوراً:"
        btns = [
            [
                InlineKeyboardButton(text="📗 قائمة التلاميذ (Excel)", callback_data="rep:students:excel"),
                InlineKeyboardButton(text="📕 قائمة التلاميذ (PDF)", callback_data="rep:students:pdf"),
            ],
            [
                InlineKeyboardButton(text="📗 الحضور والغياب (Excel)", callback_data="rep:attendance:excel"),
                InlineKeyboardButton(text="📕 الحضور والغياب (PDF)", callback_data="rep:attendance:pdf"),
            ],
            [
                InlineKeyboardButton(text="📗 النتائج والمعدلات (Excel)", callback_data="rep:grades:excel"),
                InlineKeyboardButton(text="📕 النتائج والمعدلات (PDF)", callback_data="rep:grades:pdf"),
            ],
            [
                InlineKeyboardButton(text="📗 التقرير المالي (Excel)", callback_data="rep:finance:excel"),
                InlineKeyboardButton(text="📕 التقرير المالي (PDF)", callback_data="rep:finance:pdf"),
            ],
            nav_footer_buttons("menu:home"),
        ]
        return (text, InlineKeyboardMarkup(inline_keyboard=btns), None)

    if data.startswith("rep:"):
        _, rep_type, fmt = data.split(":")
        if rep_type == "students":
            resp = await api_report_students(export=fmt, current_user=user)
        elif rep_type == "attendance":
            resp = await api_report_attendance(export=fmt, current_user=user)
        elif rep_type == "grades":
            resp = await api_report_grades(export=fmt, current_user=user)
        else:
            resp = await api_report_finance(export=fmt, current_user=user)

        ext = "xlsx" if fmt == "excel" else "pdf"
        fname = f"mounix_{rep_type}_{datetime.date.today().isoformat()}.{ext}"
        file_bytes = resp.body if isinstance(resp, Response) else b""
        return (
            f"✅ تم توليد تقرير *{rep_type.upper()}* بصيغة *{ext.upper()}* بنجاح.",
            InlineKeyboardMarkup(inline_keyboard=[nav_footer_buttons("menu:reports")]),
            (file_bytes, fname),
        )

    # 14. ⚙️ إعدادات النظام
    if data == "menu:settings":
        rows = await db.fetchall("SELECT * FROM app_settings ORDER BY key ASC")
        lines = ["⚙️ *إعدادات مدرسة MOUNIX SCHOOL*\n"]
        for r in rows:
            lines.append(f"• `{r['key']}`: *{r['value']}*")
        btns = [
            [InlineKeyboardButton(text="✏️ تعديل إعداد", callback_data="act:set:edit")],
            nav_footer_buttons("menu:home"),
        ]
        return ("\n".join(lines), InlineKeyboardMarkup(inline_keyboard=btns), None)

    if data == "act:set:edit":
        await fsm.set_state(BotFSM.setting_key)
        return (
            "✏️ *تعديل إعدادات النظام (1/2)*\nأرسل *مفتاح الإعداد (Key)* (مثال: `school_phone` أو `school_name_ar`):",
            InlineKeyboardMarkup(inline_keyboard=[nav_footer_buttons("menu:settings")]),
            None,
        )

    # 15. 💾 النسخ الاحتياطي
    if data == "menu:backups":
        backups = await db.fetchall("SELECT * FROM backup_logs ORDER BY id DESC LIMIT 8")
        lines = ["💾 *إدارة النسخ الاحتياطي لقاعدة البيانات المركزية*\n"]
        btns: List[List[InlineKeyboardButton]] = [
            [InlineKeyboardButton(text="➕ إنشاء نسخة احتياطية فورية", callback_data="act:bak:create")]
        ]
        for b in backups:
            lines.append(f"• `#{b['id']}` `{b['filename']}` ({b['size_bytes']} B) — السلامة: `{b['integrity_status']}`")
            btns.append([
                InlineKeyboardButton(text=f"♻️ استعادة #{b['id']} ({b['filename'][:18]}..)", callback_data=f"confirm:bak_rest:{b['id']}")
            ])
        btns.append(nav_footer_buttons("menu:home"))
        return ("\n".join(lines), InlineKeyboardMarkup(inline_keyboard=btns), None)

    if data == "act:bak:create":
        b_row = await create_database_backup(user["id"], user["username"])
        return (
            f"✅ *تم إنشاء النسخة الاحتياطية والتحقق من سلامتها بنجاح!*\n\n"
            f"• الملف: `{b_row['filename']}`\n"
            f"• الحجم: `{b_row['size_bytes']} bytes`\n"
            f"• البصمة (SHA-256): `{b_row['checksum_sha256'][:24]}...`",
            InlineKeyboardMarkup(inline_keyboard=[nav_footer_buttons("menu:backups")]),
            None,
        )

    if data.startswith("confirm:bak_rest:"):
        bid = int(data.split(":")[2])
        kb = InlineKeyboardMarkup(
            inline_keyboard=[
                [
                    InlineKeyboardButton(text="✅ تأكيد الاستعادة الآن", callback_data=f"act:bak:rest_do:{bid}"),
                    InlineKeyboardButton(text="❌ إلغاء", callback_data="menu:backups"),
                ]
            ]
        )
        return (f"⚠️ *تأكيد استعادة قاعدة البيانات*\nهل تريد استعادة النسخة رقم `#{bid}`؟ سيتم التحقق من سلامتها قبل الاستبدال.", kb, None)

    if data.startswith("act:bak:rest_do:"):
        bid = int(data.split(":")[3])
        res = await restore_database_backup(bid, user["id"], user["username"])
        return (
            f"♻️ *تمت استعادة النسخة الاحتياطية `{res['filename']}` بنجاح!*",
            InlineKeyboardMarkup(inline_keyboard=[nav_footer_buttons("menu:backups")]),
            None,
        )

    # 16. 🛡️ المستخدمون والصلاحيات
    if data.startswith("menu:users:"):
        rows = await db.fetchall(
            "SELECT id, username, full_name, role, is_active FROM users WHERE is_deleted = 0 ORDER BY id DESC LIMIT 12"
        )
        lines = ["🛡️ *إدارة المستخدمين والصلاحيات*\n"]
        for r in rows:
            st_icon = "🟢" if r["is_active"] else "🔴"
            lines.append(f"• {st_icon} `#{r['id']}` `{r['username']}` — *{r['full_name']}* (`{r['role']}`)")
        btns = [
            [InlineKeyboardButton(text="➕ إنشاء حساب مستخدم جديد", callback_data="act:usr:add")],
            nav_footer_buttons("menu:home"),
        ]
        return ("\n".join(lines), InlineKeyboardMarkup(inline_keyboard=btns), None)

    if data == "act:usr:add":
        await fsm.set_state(BotFSM.add_user_username)
        return (
            "➕ *إنشاء مستخدم جديد (1/4)*\nأرسل *اسم المستخدم (Username)*:",
            InlineKeyboardMarkup(inline_keyboard=[nav_footer_buttons("menu:users:1")]),
            None,
        )

    # Role-specific Student / Parent / Teacher callbacks
    if data == "role:student_dossier":
        sr = await db.fetchone("SELECT id FROM students WHERE user_id = ?", (user["id"],))
        if not sr:
            return ("❌ لم يتم ربط حسابك بملف تلميذ بعد.", build_role_keyboard(role), None)
        dossier = await get_student_full_dossier(sr["id"])
        st = dossier["student"]
        return (
            f"🎓 *ملفي الدراسي — {st['full_name']}*\n"
            f"• الرمز: `{st['student_code']}`\n"
            f"• نسبة الحضور: `{dossier['attendance']['summary']['attendance_rate_percent']}%`\n"
            f"• المعدل العام: `{dossier['grades']['average_on_20']}/20`\n"
            f"• المتبقي للدفع: `{dossier['finance']['total_remaining_debt']:.2f} DZD`",
            InlineKeyboardMarkup(inline_keyboard=[nav_footer_buttons("menu:home")]),
            None,
        )

    if data == "role:parent_children":
        pr = await db.fetchone("SELECT id FROM parents WHERE user_id = ?", (user["id"],))
        if not pr:
            return ("❌ لا يوجد ملف ولي أمر مرتبط.", build_role_keyboard(role), None)
        children = await db.fetchall(
            """
            SELECT DISTINCT s.* FROM students s
            LEFT JOIN parent_students ps ON ps.student_id = s.id
            WHERE s.primary_parent_id = ? OR ps.parent_id = ?
            """,
            (pr["id"], pr["id"]),
        )
        lines = ["👨‍👧‍👦 *الأبناء المسجلون:*\n"]
        for ch in children:
            dos = await get_student_full_dossier(ch["id"])
            lines.append(
                f"• *{ch['full_name']}* (`{ch['student_code']}`) — الحضور: "
                f"`{dos['attendance']['summary']['attendance_rate_percent']}%` | المعدل: `{dos['grades']['average_on_20']}/20` | "
                f"المستحقات: `{dos['finance']['total_remaining_debt']:.2f} DZD`"
            )
        return ("\n".join(lines), InlineKeyboardMarkup(inline_keyboard=[nav_footer_buttons("menu:home")]), None)

    if data == "role:my_notifications":
        notifs = await db.fetchall(
            "SELECT * FROM notifications WHERE recipient_user_id = ? ORDER BY id DESC LIMIT 10",
            (user["id"],),
        )
        lines = ["🔔 *أحدث الإشعارات والتنبيهات:*\n"]
        for n in notifs:
            lines.append(f"• *{n['title_ar']}*: {n['body_ar']} ({n['created_at'][:16]})")
        if not notifs:
            lines.append("لا توجد إشعارات حالياً.")
        return ("\n".join(lines), InlineKeyboardMarkup(inline_keyboard=[nav_footer_buttons("menu:home")]), None)

    return ("اختر من القائمة:", build_role_keyboard(role), None)


async def process_telegram_message_fsm(
    telegram_id: int,
    text: str,
    fsm: FSMContext,
) -> Tuple[str, InlineKeyboardMarkup]:
    """
    Handles multi-step FSM text inputs for both Admin & Role wizards.
    """
    current_state = await fsm.get_state()
    user = await get_user_by_telegram_id(telegram_id)

    # 1. Login FSM
    if current_state == BotFSM.waiting_login_username.state:
        await fsm.update_data(login_username=text.strip())
        await fsm.set_state(BotFSM.waiting_login_password)
        return (
            "🔑 يرجى إرسال *كلمة المرور (Password)*:",
            InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text="❌ إلغاء", callback_data="action:cancel")]]),
        )

    if current_state == BotFSM.waiting_login_password.state:
        data_map = await fsm.get_data()
        uname = data_map.get("login_username", "")
        u_row = await db.fetchone("SELECT * FROM users WHERE username = ? AND is_deleted = 0", (uname,))
        if not u_row or not verify_password(text.strip(), u_row["password_hash"]):
            await fsm.clear()
            return (
                "❌ اسم المستخدم أو كلمة المرور غير صحيحة.",
                InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text="🔐 إعادة المحاولة", callback_data="auth:login")]]),
            )
        async with db.transaction() as session:
            await session.execute("UPDATE users SET telegram_id = NULL WHERE telegram_id = ?", (telegram_id,))
            await session.execute(
                "UPDATE users SET telegram_id = ?, last_login_at = datetime('now') WHERE id = ?",
                (telegram_id, u_row["id"]),
            )
            await log_audit(session, u_row["id"], u_row["username"], u_row["role"], "TELEGRAM_LINK_LOGIN", "users", u_row["id"])
        await fsm.clear()
        return (
            f"✅ *تم تسجيل الدخول وربط حساب Telegram بنجاح!*\nمرحباً بك: *{u_row['full_name']}* (`{u_row['role']}`)",
            build_role_keyboard(u_row["role"]),
        )

    if not user:
        return (
            "👋 مرحباً بك في بوت *MOUNIX SCHOOL*.\nيرجى تسجيل الدخول لربط حسابك:",
            InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text="🔐 تسجيل الدخول", callback_data="auth:login")]]),
        )

    role = user["role"]

    # 2. Student Search FSM
    if current_state == BotFSM.waiting_student_search.state:
        await fsm.clear()
        q = f"%{text.strip()}%"
        rows = await db.fetchall(
            "SELECT id, student_code, full_name, phone FROM students WHERE is_archived = 0 AND (full_name LIKE ? OR student_code LIKE ? OR phone LIKE ?) LIMIT 10",
            (q, q, q),
        )
        lines = [f"🔍 *نتائج البحث عن '{text.strip()}':* ({len(rows)})\n"]
        btns = []
        for r in rows:
            lines.append(f"• `{r['student_code']}` — *{r['full_name']}*")
            btns.append([InlineKeyboardButton(text=f"📄 ملف {r['full_name']}", callback_data=f"act:stu:view:{r['id']}")])
        btns.append(nav_footer_buttons("menu:students:1"))
        return ("\n".join(lines), InlineKeyboardMarkup(inline_keyboard=btns))

    # 3. Add Student Wizard (4 steps)
    if current_state == BotFSM.add_student_name.state:
        await fsm.update_data(full_name=text.strip())
        await fsm.set_state(BotFSM.add_student_phone)
        return ("📱 *(2/4)* أرسل *رقم هاتف التلميذ أو الولي*:", InlineKeyboardMarkup(inline_keyboard=[nav_footer_buttons("menu:students:1")]))

    if current_state == BotFSM.add_student_phone.state:
        await fsm.update_data(phone=text.strip())
        await fsm.set_state(BotFSM.add_student_birth)
        return ("🎂 *(3/4)* أرسل *تاريخ الميلاد (YYYY-MM-DD)*:", InlineKeyboardMarkup(inline_keyboard=[nav_footer_buttons("menu:students:1")]))

    if current_state == BotFSM.add_student_birth.state:
        await fsm.update_data(birth_date=text.strip())
        await fsm.set_state(BotFSM.add_student_level)
        lvls = await db.fetchall("SELECT id, name_ar FROM levels WHERE is_archived = 0")
        lvl_hint = " | ".join(f"{l['id']}:{l['name_ar']}" for l in lvls) or "أرسل 0 إذا لا يوجد"
        return (f"🎓 *(4/4)* أرسل *رقم المستوى الدراسي* ({lvl_hint}):", InlineKeyboardMarkup(inline_keyboard=[nav_footer_buttons("menu:students:1")]))

    if current_state == BotFSM.add_student_level.state:
        d = await fsm.get_data()
        await fsm.clear()
        lvl_id = int(text.strip()) if text.strip().isdigit() and int(text.strip()) > 0 else None
        req = StudentCreateRequest(
            full_name=d["full_name"],
            phone=d.get("phone"),
            birth_date=d.get("birth_date"),
            level_id=lvl_id,
        )
        res = await api_create_student(req, current_user=user)
        st = res["student"]
        creds = res["credentials"]
        return (
            f"✅ *تمت إضافة التلميذ بنجاح!*\n"
            f"• الاسم: *{st['full_name']}*\n"
            f"• الرمز التعريفي: `{st['student_code']}`\n"
            f"• اسم المستخدم للتطبيق: `{creds['username']}`\n"
            f"• كلمة المرور الافتراضية: `{creds['default_password']}`",
            build_admin_main_keyboard(),
        )

    # 4. Add Teacher Wizard
    if current_state == BotFSM.add_teacher_name.state:
        await fsm.update_data(full_name=text.strip())
        await fsm.set_state(BotFSM.add_teacher_phone)
        return ("📱 *(2/3)* أرسل *رقم هاتف الأستاذ*:", InlineKeyboardMarkup(inline_keyboard=[nav_footer_buttons("menu:teachers:1")]))

    if current_state == BotFSM.add_teacher_phone.state:
        await fsm.update_data(phone=text.strip())
        await fsm.set_state(BotFSM.add_teacher_spec)
        return ("🔬 *(3/3)* أرسل *تخصص الأستاذ* (مثال: فيزياء ورياضيات):", InlineKeyboardMarkup(inline_keyboard=[nav_footer_buttons("menu:teachers:1")]))

    if current_state == BotFSM.add_teacher_spec.state:
        d = await fsm.get_data()
        await fsm.clear()
        res = await api_create_teacher(
            TeacherCreateRequest(full_name=d["full_name"], phone=d["phone"], specialization=text.strip()),
            current_user=user,
        )
        t = res["teacher"]
        return (
            f"✅ *تم تسجيل الأستاذ بنجاح!*\n• الاسم: *{t['full_name']}*\n• الرمز: `{t['teacher_code']}`",
            build_admin_main_keyboard(),
        )

    # 5. Add Parent Wizard
    if current_state == BotFSM.add_parent_name.state:
        await fsm.update_data(full_name=text.strip())
        await fsm.set_state(BotFSM.add_parent_phone)
        return ("📱 *(2/3)* أرسل *رقم هاتف ولي الأمر*:", InlineKeyboardMarkup(inline_keyboard=[nav_footer_buttons("menu:parents:1")]))

    if current_state == BotFSM.add_parent_phone.state:
        await fsm.update_data(phone=text.strip())
        await fsm.set_state(BotFSM.add_parent_student_id)
        return ("👦 *(3/3)* أرسل *رقم التلميذ (Student ID)* لربطه بولي الأمر (أو 0 لتخطي الربط):", InlineKeyboardMarkup(inline_keyboard=[nav_footer_buttons("menu:parents:1")]))

    if current_state == BotFSM.add_parent_student_id.state:
        d = await fsm.get_data()
        await fsm.clear()
        sids = [int(text.strip())] if text.strip().isdigit() and int(text.strip()) > 0 else []
        res = await api_create_parent(
            ParentCreateRequest(full_name=d["full_name"], phone=d["phone"], student_ids=sids),
            current_user=user,
        )
        p = res["parent"]
        return (
            f"✅ *تم إضافة ولي الأمر بنجاح!*\n• الاسم: *{p['full_name']}*\n• الرمز: `{p['parent_code']}`",
            build_admin_main_keyboard(),
        )

    # 6. Level & Subject Wizards
    if current_state == BotFSM.add_level_name_ar.state:
        await fsm.update_data(name_ar=text.strip())
        await fsm.set_state(BotFSM.add_level_name_en)
        return ("🌐 *(2/2)* أرسل *اسم المستوى بالإنجليزية*:", InlineKeyboardMarkup(inline_keyboard=[nav_footer_buttons("menu:subjects")]))

    if current_state == BotFSM.add_level_name_en.state:
        d = await fsm.get_data()
        await fsm.clear()
        res = await api_create_level(LevelCreateRequest(name_ar=d["name_ar"], name_en=text.strip()), current_user=user)
        return (f"✅ تم إضافة المستوى *{res['level']['name_ar']}* بنجاح!", build_admin_main_keyboard())

    if current_state == BotFSM.add_subject_name_ar.state:
        await fsm.update_data(name_ar=text.strip())
        await fsm.set_state(BotFSM.add_subject_name_en)
        return ("🌐 *(2/3)* أرسل *اسم المادة بالإنجليزية*:", InlineKeyboardMarkup(inline_keyboard=[nav_footer_buttons("menu:subjects")]))

    if current_state == BotFSM.add_subject_name_en.state:
        await fsm.update_data(name_en=text.strip())
        await fsm.set_state(BotFSM.add_subject_fee)
        return ("💵 *(3/3)* أرسل *رسوم الاشتراك الشهري الافتراضية (DZD)*:", InlineKeyboardMarkup(inline_keyboard=[nav_footer_buttons("menu:subjects")]))

    if current_state == BotFSM.add_subject_fee.state:
        d = await fsm.get_data()
        await fsm.clear()
        fee = float(text.strip()) if text.strip().replace(".", "", 1).isdigit() else 0.0
        res = await api_create_subject(
            SubjectCreateRequest(name_ar=d["name_ar"], name_en=d["name_en"], default_monthly_fee=fee),
            current_user=user,
        )
        return (f"✅ تم إضافة المادة *{res['subject']['name_ar']}* بنجاح!", build_admin_main_keyboard())

    # 7. Attendance Wizard
    if current_state == BotFSM.att_select_schedule.state:
        await fsm.update_data(schedule_id=int(text.strip()) if text.strip().isdigit() else 1)
        await fsm.set_state(BotFSM.att_select_student)
        return ("👨‍🎓 *(2/3)* أرسل *رقم التلميذ (Student ID)*:", InlineKeyboardMarkup(inline_keyboard=[nav_footer_buttons("menu:attendance")]))

    if current_state == BotFSM.att_select_student.state:
        await fsm.update_data(student_id=int(text.strip()) if text.strip().isdigit() else 1)
        await fsm.set_state(BotFSM.att_select_status)
        return ("📌 *(3/3)* أرسل الحالة (`PRESENT` أو `ABSENT` أو `LATE` أو `EXCUSED`):", InlineKeyboardMarkup(inline_keyboard=[nav_footer_buttons("menu:attendance")]))

    if current_state == BotFSM.att_select_status.state:
        d = await fsm.get_data()
        await fsm.clear()
        st_val = text.strip().upper()
        if st_val not in ("PRESENT", "ABSENT", "LATE", "EXCUSED"):
            st_val = "PRESENT"
        await api_record_attendance(
            AttendanceCreateRequest(
                schedule_id=d["schedule_id"],
                student_id=d["student_id"],
                status=AttendanceStatus(st_val),
            ),
            current_user=user,
        )
        return (f"✅ تم تسجيل الحضور (`{st_val}`) بنجاح وإرسال الإشعارات اللازمة!", build_role_keyboard(role))

    # 8. Payment Wizard
    if current_state == BotFSM.add_payment_student_id.state:
        await fsm.update_data(student_id=int(text.strip()) if text.strip().isdigit() else 1)
        await fsm.set_state(BotFSM.add_payment_amount)
        return ("💵 *(2/3)* أرسل *المبلغ المدفوع (DZD)*:", InlineKeyboardMarkup(inline_keyboard=[nav_footer_buttons("menu:finance")]))

    if current_state == BotFSM.add_payment_amount.state:
        await fsm.update_data(amount=float(text.strip()) if text.strip().replace(".", "", 1).isdigit() else 1000.0)
        await fsm.set_state(BotFSM.add_payment_method)
        return ("💳 *(3/3)* أرسل طريقة الدفع (`CASH` أو `BANK_TRANSFER` أو `CARD`):", InlineKeyboardMarkup(inline_keyboard=[nav_footer_buttons("menu:finance")]))

    if current_state == BotFSM.add_payment_method.state:
        d = await fsm.get_data()
        await fsm.clear()
        m_val = text.strip().upper()
        if m_val not in ("CASH", "BANK_TRANSFER", "CARD", "CHECK", "OTHER"):
            m_val = "CASH"
        res = await api_create_payment(
            PaymentCreateRequest(student_id=d["student_id"], amount=d["amount"], payment_method=m_val),
            current_user=user,
        )
        pay = res["payment"]
        return (
            f"✅ *تم تسجيل الدفعة المالية وإصدار الإيصال رقم `{pay['receipt_number']}` بنجاح!*\n"
            f"المبلغ: `{pay['amount']:.2f} DZD`",
            build_admin_main_keyboard(),
        )

    # 9. Broadcast Notification Wizard
    if current_state == BotFSM.broadcast_title.state:
        await fsm.update_data(title_ar=text.strip())
        await fsm.set_state(BotFSM.broadcast_body)
        return ("📝 *(2/2)* أرسل *نص الإعلان الكامل*:", InlineKeyboardMarkup(inline_keyboard=[nav_footer_buttons("menu:home")]))

    if current_state == BotFSM.broadcast_body.state:
        d = await fsm.get_data()
        await fsm.clear()
        res = await api_create_notification(
            NotificationCreateRequest(title_ar=d["title_ar"], body_ar=text.strip()),
            current_user=user,
        )
        return (f"📢 تم إرسال الإعلان إلى *{res['sent_count']}* مستخدم بنجاح!", build_role_keyboard(role))

    # 10. Settings Update Wizard
    if current_state == BotFSM.setting_key.state:
        await fsm.update_data(key=text.strip())
        await fsm.set_state(BotFSM.setting_val)
        return (f"✏️ أرسل *القيمة الجديدة* للإعداد `{text.strip()}`:", InlineKeyboardMarkup(inline_keyboard=[nav_footer_buttons("menu:settings")]))

    if current_state == BotFSM.setting_val.state:
        d = await fsm.get_data()
        await fsm.clear()
        await api_update_settings({d["key"]: text.strip()}, current_user=user)
        return (f"✅ تم تحديث الإعداد `{d['key']}` بنجاح!", build_admin_main_keyboard())

    return (
        f"🏫 *MOUNIX SCHOOL*\nمرحباً *{user['full_name']}*، اختر من القائمة:",
        build_role_keyboard(role),
    )


# ------------------------------------------------------------------------------
# Aiogram 3 Dispatcher & HTTP Bot API Polling Runner
# ------------------------------------------------------------------------------

_fallback_fsm_storage: Dict[int, Dict[str, Any]] = {}


def create_aiogram_dispatcher_and_bot() -> Optional[Tuple[Any, Any]]:
    """Initializes Aiogram 3 Bot and Dispatcher when BOT_TOKEN and aiogram are available."""
    if not HAS_AIOGRAM or not BOT_TOKEN:
        return None

    bot_instance = Bot(
        token=BOT_TOKEN,
        default=DefaultBotProperties(parse_mode=ParseMode.MARKDOWN),
    )
    dp = Dispatcher(storage=MemoryStorage())
    router = Router()

    async def _send_telegram_notification(chat_id: int, message_text: str) -> bool:
        try:
            await bot_instance.send_message(chat_id=chat_id, text=message_text)
            return True
        except Exception as exc:
            logger.warning(f"Telegram send_message failed for {chat_id}: {exc}")
            return False

    register_telegram_notifier(_send_telegram_notification)

    @router.message(CommandStart())
    async def cmd_start_handler(message: Message, state: FSMContext) -> None:
        await state.clear()
        tg_id = message.from_user.id if message.from_user else 0
        user = await get_user_by_telegram_id(tg_id)
        if not user:
            await message.answer(
                "🏫 *مرحباً بك في نظام MOUNIX SCHOOL المتكامل*\n\n"
                "يرجى تسجيل الدخول لربط حسابك (مدير / أستاذ / تلميذ / ولي أمر):",
                reply_markup=InlineKeyboardMarkup(
                    inline_keyboard=[[InlineKeyboardButton(text="🔐 تسجيل الدخول", callback_data="auth:login")]]
                ),
            )
            return
        await message.answer(
            f"🏫 *MOUNIX SCHOOL — لوحة التحكم*\nمرحباً بك: *{user['full_name']}* (`{user['role']}`)",
            reply_markup=build_role_keyboard(user["role"]),
        )

    @router.callback_query()
    async def callback_query_handler(callback: CallbackQuery, state: FSMContext) -> None:
        tg_id = callback.from_user.id if callback.from_user else 0
        cb_data = callback.data or "menu:home"
        text, markup, file_tuple = await process_telegram_callback_logic(tg_id, cb_data, state)
        with contextlib.suppress(Exception):
            await callback.answer()
        if file_tuple and callback.message:
            file_bytes, fname = file_tuple
            await callback.message.answer_document(
                BufferedInputFile(file_bytes, filename=fname),
                caption=text,
                reply_markup=markup,
            )
        elif callback.message:
            try:
                await callback.message.edit_text(text=text, reply_markup=markup)
            except Exception:
                await callback.message.answer(text=text, reply_markup=markup)

    @router.message()
    async def message_fsm_handler(message: Message, state: FSMContext) -> None:
        tg_id = message.from_user.id if message.from_user else 0
        txt = message.text or ""
        reply_text, markup = await process_telegram_message_fsm(tg_id, txt, state)
        await message.answer(text=reply_text, reply_markup=markup)

    dp.include_router(router)
    return bot_instance, dp


# ==============================================================================
# 12. UNIFIED LIFECYCLE RUNNER (FASTAPI + TELEGRAM POLLING + SQLITE)
# ==============================================================================

@app.on_event("startup")
async def on_fastapi_startup() -> None:
    await run_migrations_and_initialize_db()


@app.get("/health", tags=["System & Backup"])
async def api_health_check():
    tables_rows = await db.fetchall("SELECT name FROM sqlite_master WHERE type='table'")
    existing_tables = {r["name"] for r in tables_rows}
    missing = [t for t in REQUIRED_TABLES_24 if t not in existing_tables]
    return {
        "success": True,
        "service": "MOUNIX SCHOOL Unified Ecosystem",
        "status": "HEALTHY" if not missing else "DEGRADED",
        "database_tables_count": len(existing_tables),
        "required_24_tables_verified": len(missing) == 0,
        "missing_tables": missing,
        "telegram_bot_configured": bool(BOT_TOKEN),
        "timestamp": datetime.datetime.now().isoformat(),
    }


async def run_unified_ecosystem() -> None:
    """
    Runs SQLite initialization, FastAPI Uvicorn server, and Aiogram 3 Telegram Bot
    polling concurrently inside a single structured asyncio event loop with graceful shutdown.
    """
    await run_migrations_and_initialize_db()

    config = uvicorn.Config(
        app=app,
        host=API_HOST,
        port=API_PORT,
        log_level="info",
        access_log=True,
    )
    server = uvicorn.Server(config)
    tasks: List[asyncio.Task] = [asyncio.create_task(server.serve(), name="fastapi_uvicorn")]

    aiogram_pair = create_aiogram_dispatcher_and_bot()
    if aiogram_pair is not None:
        bot_instance, dp = aiogram_pair
        logger.info("Starting Aiogram 3 Telegram Bot polling alongside FastAPI server...")
        tasks.append(
            asyncio.create_task(
                dp.start_polling(bot_instance, handle_signals=False),
                name="aiogram_polling",
            )
        )
    else:
        if not BOT_TOKEN:
            logger.info("BOT_TOKEN environment variable is not set; running FastAPI + SQLite mode (set BOT_TOKEN to enable live Telegram polling).")

    try:
        await asyncio.gather(*tasks)
    except (asyncio.CancelledError, KeyboardInterrupt):
        logger.info("Shutting down MOUNIX SCHOOL ecosystem gracefully...")
    finally:
        if aiogram_pair is not None:
            bot_instance, _ = aiogram_pair
            with contextlib.suppress(Exception):
                await bot_instance.session.close()
        logger.info("All connections closed cleanly.")


if __name__ == "__main__":
    try:
        asyncio.run(run_unified_ecosystem())
    except KeyboardInterrupt:
        pass
