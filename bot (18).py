
import asyncio
import copy
import hashlib
import textwrap
import html
import io
import json
import logging
import math
import os
import random
import re
import sqlite3
import tempfile
import time
import uuid
import zipfile
from datetime import datetime, timedelta, timezone
from pathlib import Path
from xml.sax.saxutils import escape
from urllib.parse import quote, urlparse, parse_qs
from zoneinfo import ZoneInfo
 
import httpx
from aiohttp import web
from aiogram import Bot, Dispatcher, F
from aiogram.enums import ParseMode
from aiogram.exceptions import TelegramRetryAfter
from aiogram.filters import Command, CommandStart
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.types import (
    Message, CallbackQuery, FSInputFile, InlineKeyboardButton,
    InlineKeyboardMarkup, ReplyKeyboardMarkup, KeyboardButton,
    BotCommand, LabeledPrice,
)
from aiogram.client.default import DefaultBotProperties
from aiogram.webhook.aiohttp_server import SimpleRequestHandler, setup_application
 
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.xmlchemy import OxmlElement
from pptx.util import Inches, Pt
 
from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT, TA_CENTER
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from reportlab.platypus import Paragraph
 
# ============================================================
# 1. CONFIGURATION
# ============================================================
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
)
log = logging.getLogger("prezent-ai")


async def notify_admins_once(event_key: str, text: str, cooldown: float = 600.0):
    """Send rate-limited operational alerts without creating alert storms."""
    if not ADMIN_IDS:
        return
    now = time.monotonic()
    if now - _ADMIN_ALERT_COOLDOWNS.get(event_key, 0.0) < cooldown:
        return
    _ADMIN_ALERT_COOLDOWNS[event_key] = now
    for admin_id in ADMIN_IDS:
        try:
            await bot.send_message(admin_id, text[:3900], parse_mode=ParseMode.HTML)
        except Exception as exc:
            log.warning("Could not send admin alert %s to %s: %s", event_key, admin_id, exc)
 
BOT_TOKEN = os.getenv("BOT_TOKEN", "").strip()
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "").strip()
GEMINI_MODEL = os.getenv("GEMINI_MODEL", "gemini-3.5-flash-lite").strip()
GEMINI_FALLBACK_MODEL = os.getenv("GEMINI_FALLBACK_MODEL", "gemini-3.1-flash-lite").strip()
# Dedicated low-cost model for the Google Search fact-checking pass.
GEMINI_RESEARCH_MODEL = os.getenv("GEMINI_RESEARCH_MODEL", "gemini-3.5-flash-lite").strip()
ADMIN_IDS_RAW = os.getenv("ADMIN_IDS", os.getenv("ADMIN_ID", ""))
ADMIN_IDS = {int(x.strip()) for x in ADMIN_IDS_RAW.split(",") if x.strip().isdigit()}
DATABASE_PATH = os.getenv("DATABASE_PATH", "prezentai.sqlite3").strip()
RENDER_IS_HOSTED = bool(os.getenv("RENDER_EXTERNAL_URL", "").strip())
RENDER_DISK_MOUNT_PATH = os.getenv("RENDER_DISK_MOUNT_PATH", "/data").strip() or "/data"
TURSO_DATABASE_URL = os.getenv("TURSO_DATABASE_URL", "").strip()
TURSO_AUTH_TOKEN = os.getenv("TURSO_AUTH_TOKEN", "").strip()
PEXELS_API_KEY = os.getenv("PEXELS_API_KEY", "").strip()
WIKIMEDIA_CONTACT_EMAIL = os.getenv("WIKIMEDIA_CONTACT_EMAIL", "").strip()
MAX_DECK_IMAGES = 5
GEMINI_OUTPUT_TOKEN_LIMIT = 32768
GEMINI_RETRY_OUTPUT_TOKEN_LIMIT = 65536
TURSO_ENABLED = bool(TURSO_DATABASE_URL and TURSO_AUTH_TOKEN)
if bool(TURSO_DATABASE_URL) != bool(TURSO_AUTH_TOKEN):
    raise RuntimeError("Set BOTH TURSO_DATABASE_URL and TURSO_AUTH_TOKEN, or remove both.")


def _is_render_disk_path(path: str, mount_path: str | None = None) -> bool:
    """Return True only when the database is inside the configured mounted disk."""
    try:
        resolved = Path(path).expanduser().resolve()
        mount = Path(mount_path or RENDER_DISK_MOUNT_PATH).expanduser().resolve()
        if not mount.exists() or not mount.is_dir():
            return False
        return os.path.ismount(str(mount)) and (resolved == mount or resolved.is_relative_to(mount))
    except (OSError, RuntimeError, ValueError):
        return False


DURABLE_DB_ENABLED = TURSO_ENABLED or _is_render_disk_path(DATABASE_PATH, RENDER_DISK_MOUNT_PATH)
MAX_TOPIC_LENGTH = 500
GENERATION_LEASE_SECONDS = 600
GENERATION_LEASE_HEARTBEAT_SECONDS = 60
GENERATION_QUEUE_LIMIT = 6
GENERATION_TOTAL_TIMEOUT_SECONDS = 360
GENERATION_DETAILED_TIMEOUT_SECONDS = 480
MAX_IMAGE_DOWNLOAD_BYTES = 8 * 1024 * 1024
MAX_IMAGE_PIXELS = 28_000_000
MAX_OUTPUT_FILE_BYTES = 45 * 1024 * 1024
MAX_RESTORE_BACKUP_BYTES = 45 * 1024 * 1024
FREE_DAILY_LIMIT = 3
PRO_DAILY_LIMIT = 10
CREDIT_PACK_SIZE = 5
CREDIT_PACK_PRICE = 15
PRO_PRICE = 99
PRO_DAYS = 30
KYIV_TZ = ZoneInfo("Europe/Kyiv")
GEN_SEMAPHORE = asyncio.Semaphore(2)
GENERATION_QUEUE_WAITERS = 0
WORKER_ID = uuid.uuid4().hex
_ADMIN_ALERT_COOLDOWNS: dict[str, float] = {}
 
if not BOT_TOKEN:
    raise RuntimeError("BOT_TOKEN is missing in Render Environment.")
if not GEMINI_API_KEY:
    raise RuntimeError("GEMINI_API_KEY is missing in Render Environment.")
 
Path(DATABASE_PATH).expanduser().resolve().parent.mkdir(parents=True, exist_ok=True)
bot = Bot(token=BOT_TOKEN, default=DefaultBotProperties(parse_mode=ParseMode.HTML))
dp = Dispatcher(storage=MemoryStorage())
ACTIVE_GENERATIONS = {}
USER_LANGUAGE_CACHE: dict[int, str | None] = {}
LAST_GENERATION_PARAMS: dict[int, dict] = {}
 
# ============================================================
# 2. DATABASE
# ============================================================
class CompatRow:
    """Small sqlite3.Row-compatible mapping for remote DB-API tuple rows."""
    __slots__ = ("_columns", "_values", "_index")

    def __init__(self, values, columns):
        self._columns = tuple(columns or ())
        self._values = tuple(values)
        self._index = {name: i for i, name in enumerate(self._columns)}

    def __getitem__(self, key):
        if isinstance(key, str):
            return self._values[self._index[key]]
        return self._values[key]

    def keys(self):
        return list(self._columns)

    def __iter__(self):
        return iter(self._values)

    def __len__(self):
        return len(self._values)


class CompatCursor:
    def __init__(self, cursor):
        self._cursor = cursor

    @property
    def description(self):
        return getattr(self._cursor, "description", None)

    @property
    def rowcount(self):
        return getattr(self._cursor, "rowcount", -1)

    def _convert(self, row):
        if row is None or hasattr(row, "keys"):
            return row
        description = self.description or ()
        columns = [column[0] for column in description]
        return CompatRow(row, columns)

    def fetchone(self):
        return self._convert(self._cursor.fetchone())

    def fetchall(self):
        return [self._convert(row) for row in self._cursor.fetchall()]

    def __iter__(self):
        for row in self._cursor:
            yield self._convert(row)


class DBConnection:
    """Explicitly closed DB-API connection with sqlite-compatible row access."""
    def __init__(self, raw_connection, remote=False):
        self.raw = raw_connection
        self.remote = remote

    def execute(self, sql, params=()):
        # Both sqlite3 and turso_serverless implement DB-API execute(sql, params).
        # Never retry without params: doing so could turn a failed parameterized
        # query into an unsafe or invalid statement.
        cursor = self.raw.execute(sql, params)
        return CompatCursor(cursor)

    def executescript(self, script):
        if not self.remote and hasattr(self.raw, "executescript"):
            self.raw.executescript(script)
            return
        # The schema contains ordinary DDL statements without semicolons in literals.
        for statement in script.split(";"):
            statement = statement.strip()
            if statement:
                self.execute(statement)

    def commit(self):
        return self.raw.commit()

    def rollback(self):
        return self.raw.rollback()

    def close(self):
        return self.raw.close()

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, tb):
        try:
            if exc_type is None:
                self.commit()
            else:
                self.rollback()
        finally:
            self.close()
        return False


DB_SCHEMA_SQL = """
CREATE TABLE IF NOT EXISTS users (
    user_id BIGINT PRIMARY KEY,
    username TEXT DEFAULT '',
    first_name TEXT DEFAULT '',
    interface_language TEXT DEFAULT NULL,
    credits INTEGER NOT NULL DEFAULT 0,
    subscription_until TEXT DEFAULT NULL,
    created_at TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS daily_usage (
    user_id BIGINT NOT NULL,
    day TEXT NOT NULL,
    used INTEGER NOT NULL DEFAULT 0,
    PRIMARY KEY (user_id, day),
    FOREIGN KEY (user_id) REFERENCES users(user_id)
);
CREATE TABLE IF NOT EXISTS orders (
    payload TEXT PRIMARY KEY,
    user_id BIGINT NOT NULL,
    product TEXT NOT NULL,
    amount INTEGER NOT NULL,
    status TEXT NOT NULL DEFAULT 'pending',
    created_at TEXT NOT NULL,
    charge_id TEXT UNIQUE
);
CREATE TABLE IF NOT EXISTS generation_history (
    generation_id TEXT PRIMARY KEY,
    user_id BIGINT NOT NULL,
    topic TEXT NOT NULL,
    design TEXT NOT NULL,
    slides INTEGER NOT NULL,
    quota_mode TEXT NOT NULL,
    status TEXT NOT NULL,
    created_at TEXT NOT NULL,
    source_chat_id BIGINT,
    source_message_id BIGINT,
    detail TEXT NOT NULL DEFAULT 'medium',
    language TEXT NOT NULL DEFAULT 'uk',
    quota_day TEXT,
    photos INTEGER NOT NULL DEFAULT 1,
    lease_until TEXT,
    worker_id TEXT,
    current_stage TEXT NOT NULL DEFAULT 'queued',
    delivery_resolution TEXT,
    resolved_by BIGINT,
    resolved_at TEXT
);
CREATE TABLE IF NOT EXISTS generation_locks (
    user_id BIGINT PRIMARY KEY,
    generation_id TEXT NOT NULL UNIQUE,
    worker_id TEXT NOT NULL,
    lease_until TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS reports (
    report_id TEXT PRIMARY KEY,
    generation_id TEXT NOT NULL,
    user_id BIGINT NOT NULL,
    reason TEXT NOT NULL,
    status TEXT NOT NULL DEFAULT 'pending',
    created_at TEXT NOT NULL,
    handled_by BIGINT,
    handled_at TEXT
);
CREATE INDEX IF NOT EXISTS idx_history_user ON generation_history(user_id, created_at);
CREATE INDEX IF NOT EXISTS idx_reports_status ON reports(status);
"""


def db_connect():
    """Open a durable Turso DB when configured; otherwise local SQLite."""
    if TURSO_ENABLED:
        try:
            import turso_serverless
        except ImportError as exc:
            raise RuntimeError(
                "Turso database is configured but turso_serverless is missing from requirements.txt."
            ) from exc
        raw = turso_serverless.connect(TURSO_DATABASE_URL, auth_token=TURSO_AUTH_TOKEN)
        return DBConnection(raw, remote=True)

    conn = sqlite3.connect(DATABASE_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA foreign_keys=ON")
    conn.execute("PRAGMA busy_timeout=30000")
    return DBConnection(conn, remote=False)


def init_db():
    with db_connect() as con:
        con.executescript(DB_SCHEMA_SQL)
        # Safe migration for users created by earlier bot versions.
        columns = {row["name"] for row in con.execute("PRAGMA table_info(users)").fetchall()}
        if "interface_language" not in columns:
            con.execute("ALTER TABLE users ADD COLUMN interface_language TEXT DEFAULT NULL")
        history_columns = {row["name"] for row in con.execute("PRAGMA table_info(generation_history)").fetchall()}
        if "detail" not in history_columns:
            con.execute("ALTER TABLE generation_history ADD COLUMN detail TEXT NOT NULL DEFAULT 'medium'")
        if "language" not in history_columns:
            con.execute("ALTER TABLE generation_history ADD COLUMN language TEXT NOT NULL DEFAULT 'uk'")
        if "quota_day" not in history_columns:
            con.execute("ALTER TABLE generation_history ADD COLUMN quota_day TEXT")
        if "photos" not in history_columns:
            con.execute("ALTER TABLE generation_history ADD COLUMN photos INTEGER NOT NULL DEFAULT 1")
        if "lease_until" not in history_columns:
            con.execute("ALTER TABLE generation_history ADD COLUMN lease_until TEXT")
        if "worker_id" not in history_columns:
            con.execute("ALTER TABLE generation_history ADD COLUMN worker_id TEXT")
        if "current_stage" not in history_columns:
            con.execute("ALTER TABLE generation_history ADD COLUMN current_stage TEXT NOT NULL DEFAULT 'queued'")
        if "delivery_resolution" not in history_columns:
            con.execute("ALTER TABLE generation_history ADD COLUMN delivery_resolution TEXT")
        if "resolved_by" not in history_columns:
            con.execute("ALTER TABLE generation_history ADD COLUMN resolved_by BIGINT")
        if "resolved_at" not in history_columns:
            con.execute("ALTER TABLE generation_history ADD COLUMN resolved_at TEXT")
        con.execute("CREATE TABLE IF NOT EXISTS generation_locks (user_id BIGINT PRIMARY KEY, generation_id TEXT NOT NULL UNIQUE, worker_id TEXT NOT NULL, lease_until TEXT NOT NULL)")

def now_utc():
    return datetime.now(timezone.utc)
 
 
def now_iso():
    return now_utc().isoformat()
 
 
def today_key():
    return datetime.now(KYIV_TZ).date().isoformat()
 
 
def is_admin(user_id: int) -> bool:
    return user_id in ADMIN_IDS
 
 
def register_user(user):
    with db_connect() as con:
        con.execute(
            "INSERT INTO users(user_id, username, first_name, created_at) VALUES(?,?,?,?) "
            "ON CONFLICT(user_id) DO UPDATE SET username=excluded.username, first_name=excluded.first_name",
            (user.id, user.username or "", user.first_name or "", now_iso()),
        )
 
 
def get_user(user_id: int):
    with db_connect() as con:
        return con.execute("SELECT * FROM users WHERE user_id=?", (user_id,)).fetchone()


def load_user_language_cache():
    """Warm the tiny language cache once at startup; async handlers then do no DB I/O for language lookup."""
    USER_LANGUAGE_CACHE.clear()
    with db_connect() as con:
        rows = con.execute("SELECT user_id, interface_language FROM users").fetchall()
    for row in rows:
        value = row["interface_language"]
        USER_LANGUAGE_CACHE[int(row["user_id"])] = value if value in ("uk", "ru", "en") else None


def get_user_language(user_id: int) -> str | None:
    # Cache misses are intentionally treated as an unconfigured language; startup warms all known users.
    return USER_LANGUAGE_CACHE.get(int(user_id))


def set_user_language(user_id: int, language: str):
    if language not in ("uk", "ru", "en"):
        raise ValueError("Unsupported interface language")
    with db_connect() as con:
        con.execute(
            "INSERT INTO users(user_id, interface_language, created_at) VALUES(?,?,?) "
            "ON CONFLICT(user_id) DO UPDATE SET interface_language=excluded.interface_language",
            (user_id, language, now_iso()),
        )
    USER_LANGUAGE_CACHE[int(user_id)] = language


def begin_write_transaction(con: DBConnection):
    """Local SQLite acquires a writer lock immediately; Turso uses its remote BEGIN semantics."""
    con.execute("BEGIN" if con.remote else "BEGIN IMMEDIATE")


def get_admin_stats_db():
    with db_connect() as con:
        users = con.execute("SELECT COUNT(*) AS n FROM users").fetchone()["n"]
        gens = con.execute("SELECT COUNT(*) AS n FROM generation_history WHERE status='completed'").fetchone()["n"]
        pending = con.execute("SELECT COUNT(*) AS n FROM reports WHERE status='pending'").fetchone()["n"]
        revenue = con.execute("SELECT COALESCE(SUM(amount),0) AS n FROM orders WHERE status='paid'").fetchone()["n"]
    return int(users), int(gens), int(pending), int(revenue)


def grant_credits_db(user_id: int, amount: int):
    if amount <= 0:
        raise ValueError("Кількість кредитів має бути додатною.")
    with db_connect() as con:
        con.execute("INSERT INTO users(user_id, created_at) VALUES(?,?) ON CONFLICT(user_id) DO NOTHING", (user_id, now_iso()))
        con.execute("UPDATE users SET credits=MAX(0, credits + ?) WHERE user_id=?", (amount, user_id))


def get_database_summary():
    with db_connect() as con:
        users_row = con.execute("SELECT COUNT(*) AS n FROM users").fetchone()
        users_count = int(users_row["n"] if users_row else 0)
        credits_row = con.execute("SELECT COALESCE(SUM(credits),0) AS n FROM users").fetchone()
        total_credits = int(credits_row["n"] if credits_row else 0)
    return users_count, total_credits


def get_generation_owner(generation_id: str):
    with db_connect() as con:
        return con.execute("SELECT user_id FROM generation_history WHERE generation_id=?", (generation_id,)).fetchone()


def get_user_count_db():
    with db_connect() as con:
        row = con.execute("SELECT COUNT(*) AS n FROM users").fetchone()
        return int(row["n"] if row else 0)
 
 
def get_usage(user_id: int) -> int:
    with db_connect() as con:
        row = con.execute(
            "SELECT used FROM daily_usage WHERE user_id=? AND day=?",
            (user_id, today_key()),
        ).fetchone()
        return int(row["used"]) if row else 0
 
 
def subscription_active(row) -> bool:
    if not row or not row["subscription_until"]:
        return False
    try:
        return datetime.fromisoformat(row["subscription_until"]) > now_utc()
    except ValueError:
        return False
 
 
def _refund_quota_in_transaction(con: DBConnection, user_id: int, mode: str, quota_day: str | None):
    """Refund one reservation. The caller must hold the transaction."""
    if mode == "credit":
        con.execute("UPDATE users SET credits=credits+1 WHERE user_id=?", (user_id,))
    elif mode in ("free", "pro"):
        con.execute(
            "UPDATE daily_usage SET used=CASE WHEN used>0 THEN used-1 ELSE 0 END WHERE user_id=? AND day=?",
            (user_id, quota_day or today_key()),
        )


def _insert_generation_history(con: DBConnection, generation: dict, lease_until: str, worker_id: str):
    con.execute(
        "INSERT INTO generation_history(generation_id,user_id,topic,design,slides,quota_mode,status,created_at,detail,language,quota_day,photos,lease_until,worker_id,current_stage) "
        "VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
        (
            generation["generation_id"], generation["user_id"], generation["topic"],
            generation["design"], generation["slides"], generation["quota_mode"], "pending",
            now_iso(), generation.get("detail", "medium"), generation.get("language", "uk"),
            generation.get("quota_day") or today_key(), 1 if generation.get("photos", True) else 0,
            lease_until, worker_id, "queued",
        ),
    )


def reserve_generation(user_id: int, generation: dict | None = None, worker_id: str | None = None):
    """Atomically reserve quota, create history, and acquire a cross-worker user lock."""
    day = today_key()
    now = now_utc()
    now_text = now.isoformat()
    worker_id = worker_id or WORKER_ID
    lease_until = (now + timedelta(seconds=GENERATION_LEASE_SECONDS)).isoformat()
    mode = "admin" if is_admin(user_id) else "none"

    with db_connect() as con:
        begin_write_transaction(con)
        con.execute(
            "INSERT INTO users(user_id,created_at) VALUES(?,?) ON CONFLICT(user_id) DO NOTHING",
            (user_id, now_text),
        )

        existing_lock = con.execute(
            "SELECT generation_id,lease_until FROM generation_locks WHERE user_id=?", (user_id,)
        ).fetchone()
        if existing_lock:
            try:
                existing_expiry = datetime.fromisoformat(existing_lock["lease_until"])
            except (TypeError, ValueError):
                existing_expiry = now - timedelta(seconds=1)
            if existing_expiry > now:
                return {"allowed": False, "mode": "busy", "used": 0, "limit": 0, "credits": 0, "day": day}

            # Reconcile an expired lock before it can be reused.
            old = con.execute(
                "SELECT user_id,quota_mode,quota_day,current_stage,status FROM generation_history WHERE generation_id=?",
                (existing_lock["generation_id"],),
            ).fetchone()
            if old and old["status"] == "pending":
                uncertain = old["current_stage"] == "sending"
                new_status = "delivery_unknown" if uncertain else "recovered"
                transitioned = con.execute(
                    "UPDATE generation_history SET status=?,lease_until=NULL,worker_id=NULL "
                    "WHERE generation_id=? AND status='pending' RETURNING generation_id",
                    (new_status, existing_lock["generation_id"]),
                ).fetchone()
                if transitioned and not uncertain:
                    _refund_quota_in_transaction(
                        con, int(old["user_id"]), str(old["quota_mode"] or "none"), old["quota_day"]
                    )
                if transitioned and uncertain:
                    log.error(
                        "Expired generation %s was in sending stage; quota was NOT refunded automatically",
                        existing_lock["generation_id"],
                    )
            con.execute("DELETE FROM generation_locks WHERE user_id=?", (user_id,))

        queued_row = con.execute(
            "SELECT COUNT(*) AS n FROM generation_history "
            "WHERE status='pending' AND current_stage='queued' AND lease_until>?",
            (now_text,),
        ).fetchone()
        queued_before = int(queued_row["n"] if queued_row else 0)
        if queued_before >= GENERATION_QUEUE_LIMIT:
            return {
                "allowed": False, "mode": "queue_full", "used": 0,
                "limit": 0, "credits": 0, "day": day,
            }

        row = con.execute("SELECT * FROM users WHERE user_id=?", (user_id,)).fetchone()
        if is_admin(user_id):
            used, limit, credits = 0, "∞", int(row["credits"])
            mode = "admin"
        else:
            is_pro = subscription_active(row)
            limit = PRO_DAILY_LIMIT if is_pro else FREE_DAILY_LIMIT
            usage = con.execute(
                "SELECT used FROM daily_usage WHERE user_id=? AND day=?", (user_id, day)
            ).fetchone()
            used = int(usage["used"]) if usage else 0
            credits = int(row["credits"])
            if used < limit:
                con.execute(
                    "INSERT INTO daily_usage(user_id,day,used) VALUES(?,?,1) "
                    "ON CONFLICT(user_id,day) DO UPDATE SET used=used+1", (user_id, day)
                )
                mode = "pro" if is_pro else "free"
                used += 1
            elif credits > 0:
                con.execute("UPDATE users SET credits=credits-1 WHERE user_id=?", (user_id,))
                credits -= 1
                mode = "credit"
            else:
                return {"allowed": False, "mode": "none", "used": used, "limit": limit, "credits": credits, "day": day}

        result = {
            "allowed": True, "mode": mode, "used": used, "limit": limit,
            "credits": credits, "day": day, "queue_position": queued_before + 1,
        }
        if generation:
            record = dict(generation)
            record.update({"user_id": user_id, "quota_mode": mode, "quota_day": day})
            _insert_generation_history(con, record, lease_until, worker_id)
            con.execute(
                "INSERT INTO generation_locks(user_id,generation_id,worker_id,lease_until) VALUES(?,?,?,?)",
                (user_id, record["generation_id"], worker_id, lease_until),
            )
        return result


def renew_generation_lease(generation_id: str, user_id: int, worker_id: str, lease_seconds: int = GENERATION_LEASE_SECONDS) -> bool:
    expiry = (now_utc() + timedelta(seconds=max(120, int(lease_seconds)))).isoformat()
    with db_connect() as con:
        begin_write_transaction(con)
        con.execute(
            "UPDATE generation_locks SET lease_until=? WHERE user_id=? AND generation_id=? AND worker_id=?",
            (expiry, user_id, generation_id, worker_id),
        )
        lock = con.execute(
            "SELECT 1 FROM generation_locks WHERE user_id=? AND generation_id=? AND worker_id=?",
            (user_id, generation_id, worker_id),
        ).fetchone()
        if not lock:
            return False
        con.execute(
            "UPDATE generation_history SET lease_until=? WHERE generation_id=? AND user_id=? AND status='pending'",
            (expiry, generation_id, user_id),
        )
        return True


def set_generation_stage(generation_id: str, stage: str):
    allowed = {"queued", "research", "writing", "photos", "pptx", "pdf", "sending"}
    if stage not in allowed:
        return False
    with db_connect() as con:
        con.execute(
            "UPDATE generation_history SET current_stage=? WHERE generation_id=? AND status='pending'",
            (stage, generation_id),
        )
        row = con.execute(
            "SELECT current_stage,status FROM generation_history WHERE generation_id=?", (generation_id,)
        ).fetchone()
        return bool(row and row["status"] == "pending" and row["current_stage"] == stage)


def finalize_generation(generation_id: str, success: bool, chat_id: int | None = None, message_id: int | None = None):
    """Atomically complete/fail a generation and refund only once.

    A successful Telegram send is authoritative delivery evidence, so it may
    upgrade an ambiguous `delivery_unknown` row to `completed` without refund.
    RETURNING makes the pending->terminal transition the exact-once refund guard.
    """
    refunded_credit = False
    result = "missing"
    with db_connect() as con:
        begin_write_transaction(con)
        row = con.execute(
            "SELECT user_id,quota_mode,quota_day,current_stage,status FROM generation_history WHERE generation_id=?",
            (generation_id,),
        ).fetchone()
        if not row:
            return "missing"
        current_status = str(row["status"] or "")
        if current_status == "completed":
            con.execute("DELETE FROM generation_locks WHERE generation_id=?", (generation_id,))
            return "completed"
        user_id = int(row["user_id"])
        if success:
            if current_status not in ("pending", "delivery_unknown"):
                con.execute("DELETE FROM generation_locks WHERE generation_id=?", (generation_id,))
                return current_status or "missing"
            changed = con.execute(
                "UPDATE generation_history SET status='completed',source_chat_id=COALESCE(?,source_chat_id),"
                "source_message_id=COALESCE(?,source_message_id),lease_until=NULL,worker_id=NULL "
                "WHERE generation_id=? AND status IN ('pending','delivery_unknown') RETURNING generation_id",
                (chat_id, message_id, generation_id),
            ).fetchone()
            if not changed:
                latest = con.execute(
                    "SELECT status FROM generation_history WHERE generation_id=?", (generation_id,)
                ).fetchone()
                return str(latest["status"]) if latest else "missing"
            result = "completed"
        elif current_status != "pending":
            con.execute("DELETE FROM generation_locks WHERE generation_id=?", (generation_id,))
            return current_status
        elif row["current_stage"] == "sending":
            # A network failure during upload cannot prove Telegram did not deliver the file.
            # Do not automatically refund an ambiguous delivery; route it to admin review.
            changed = con.execute(
                "UPDATE generation_history SET status='delivery_unknown',lease_until=NULL,worker_id=NULL "
                "WHERE generation_id=? AND status='pending' RETURNING generation_id",
                (generation_id,),
            ).fetchone()
            if not changed:
                latest = con.execute(
                    "SELECT status FROM generation_history WHERE generation_id=?", (generation_id,)
                ).fetchone()
                return str(latest["status"]) if latest else "missing"
            result = "delivery_unknown"
        else:
            changed = con.execute(
                "UPDATE generation_history SET status='failed',lease_until=NULL,worker_id=NULL "
                "WHERE generation_id=? AND status='pending' RETURNING generation_id",
                (generation_id,),
            ).fetchone()
            if not changed:
                latest = con.execute(
                    "SELECT status FROM generation_history WHERE generation_id=?", (generation_id,)
                ).fetchone()
                return str(latest["status"]) if latest else "missing"
            _refund_quota_in_transaction(con, user_id, str(row["quota_mode"] or "none"), row["quota_day"])
            refunded_credit = row["quota_mode"] == "credit"
            result = "failed"
        con.execute("DELETE FROM generation_locks WHERE generation_id=?", (generation_id,))
        confirmed = con.execute(
            "SELECT status FROM generation_history WHERE generation_id=?", (generation_id,)
        ).fetchone()
        if not confirmed or confirmed["status"] != result:
            raise RuntimeError(f"Generation finalization could not be confirmed: {generation_id}")
    if refunded_credit:
        schedule_balance_backup("повернення витраченого кредиту")
    return result


async def generation_lease_heartbeat(generation_id: str, user_id: int, worker_id: str):
    while True:
        await asyncio.sleep(GENERATION_LEASE_HEARTBEAT_SECONDS)
        try:
            renewed = await asyncio.to_thread(renew_generation_lease, generation_id, user_id, worker_id)
            if not renewed:
                log.error("Generation lease was lost: generation=%s user=%s", generation_id, user_id)
                await notify_admins_once(
                    f"lease-lost-{generation_id}",
                    f"⚠️ <b>Втрачено блокування генерації.</b>\nGeneration: <code>{html.escape(generation_id)}</code>\nUser: <code>{user_id}</code>. Старий worker буде зупинено.",
                    cooldown=60,
                )
                active_task = ACTIVE_GENERATIONS.get(user_id)
                if active_task and active_task is not asyncio.current_task() and not active_task.done():
                    active_task.cancel()
                return
        except asyncio.CancelledError:
            raise
        except Exception:
            log.exception("Generation lease heartbeat failed for %s", generation_id)
def add_history(
    generation_id, user_id, topic, design, slides, quota_mode, status="pending",
    detail="medium", language="uk", quota_day=None, photos=True,
):
    """Persist the quota reservation before generation so a restart can safely recover it."""
    with db_connect() as con:
        con.execute(
            "INSERT INTO generation_history(generation_id,user_id,topic,design,slides,quota_mode,status,created_at,detail,language,quota_day,photos) "
            "VALUES(?,?,?,?,?,?,?,?,?,?,?,?)",
            (generation_id, user_id, topic, design, slides, quota_mode, status, now_iso(),
             detail if detail in DETAIL_PROFILES else "medium",
             language if language in ("uk", "ru", "en") else "uk",
             quota_day or today_key(), 1 if photos else 0),
        )


def mark_history_completed(generation_id: str):
    with db_connect() as con:
        con.execute(
            "UPDATE generation_history SET status='completed' WHERE generation_id=? AND status='pending'",
            (generation_id,),
        )


def get_latest_completed_generation(user_id: int):
    with db_connect() as con:
        return con.execute(
            "SELECT topic,design,slides,detail,language,photos FROM generation_history "
            "WHERE user_id=? AND status='completed' ORDER BY created_at DESC LIMIT 1",
            (user_id,),
        ).fetchone()


def recover_stale_pending_generations(max_age_seconds: int = 240):
    """Recover expired generation leases atomically; never auto-refund ambiguous sends."""
    cutoff = (now_utc() - timedelta(seconds=max(60, int(max_age_seconds)))).isoformat()
    now_text = now_iso()
    active_users = set(ACTIVE_GENERATIONS.keys())
    recovered = 0
    uncertain = 0
    with db_connect() as con:
        begin_write_transaction(con)
        rows = con.execute(
            "SELECT generation_id,user_id,quota_mode,quota_day,created_at,current_stage,lease_until "
            "FROM generation_history WHERE status='pending' AND "
            "((lease_until IS NOT NULL AND lease_until < ?) OR (lease_until IS NULL AND created_at < ?))",
            (now_text, cutoff),
        ).fetchall()
        for row in rows:
            user_id = int(row["user_id"])
            if user_id in active_users:
                continue
            is_uncertain = row["current_stage"] == "sending"
            new_status = "delivery_unknown" if is_uncertain else "recovered"
            transitioned = con.execute(
                "UPDATE generation_history SET status=?,lease_until=NULL,worker_id=NULL "
                "WHERE generation_id=? AND status='pending' RETURNING generation_id",
                (new_status, row["generation_id"]),
            ).fetchone()
            con.execute("DELETE FROM generation_locks WHERE generation_id=?", (row["generation_id"],))
            if not transitioned:
                continue
            if is_uncertain:
                uncertain += 1
                continue
            _refund_quota_in_transaction(con, user_id, str(row["quota_mode"] or "none"), row["quota_day"])
            recovered += 1
    if recovered or uncertain:
        log.warning("Generation recovery: refunded=%s delivery_unknown=%s", recovered, uncertain)
    return recovered
def set_history_message(generation_id: str, chat_id: int, message_id: int):
    with db_connect() as con:
        con.execute(
            "UPDATE generation_history SET source_chat_id=?, source_message_id=? WHERE generation_id=?",
            (chat_id, message_id, generation_id),
        )


def mark_history_failed(generation_id: str):
    with db_connect() as con:
        con.execute(
            "UPDATE generation_history SET status='failed' WHERE generation_id=? AND status='pending'",
            (generation_id,),
        )


def create_order(user_id: int, product: str, amount: int) -> str:
    payload = f"pzt:{product}:{user_id}:{uuid.uuid4().hex[:16]}"
    with db_connect() as con:
        con.execute(
            "INSERT INTO orders(payload,user_id,product,amount,status,created_at) VALUES(?,?,?,?,?,?)",
            (payload, user_id, product, amount, "pending", now_iso()),
        )
    return payload
 
 
def find_order(payload: str):
    with db_connect() as con:
        return con.execute("SELECT * FROM orders WHERE payload=?", (payload,)).fetchone()
 
 
def find_order_by_charge(charge_id: str):
    with db_connect() as con:
        return con.execute("SELECT * FROM orders WHERE charge_id=?", (charge_id,)).fetchone()


def process_payment(payload: str, user_id: int, amount: int, currency: str, charge_id: str):
    """Idempotently apply one verified Telegram Stars charge in a single DB transaction."""
    if currency != "XTR":
        return False, "Неправильна валюта платежу."
    if not charge_id:
        return False, "Порожній Telegram charge_id."
    with db_connect() as con:
        begin_write_transaction(con)
        order = con.execute("SELECT * FROM orders WHERE payload=?", (payload,)).fetchone()
        if not order or int(order["user_id"]) != user_id:
            return False, "Замовлення не знайдено або воно належить іншому користувачу."
        if int(order["amount"]) != amount:
            return False, "Сума платежу не збігається."
        if order["status"] == "paid":
            if order["charge_id"] == charge_id:
                return True, "already_paid"
            return False, "Замовлення вже оплачено іншим charge_id."
        if order["status"] != "pending":
            return False, "Замовлення не перебуває в очікуванні оплати."

        duplicate = con.execute("SELECT payload,user_id FROM orders WHERE charge_id=?", (charge_id,)).fetchone()
        if duplicate:
            if duplicate["payload"] == payload and int(duplicate["user_id"]) == user_id:
                return True, "already_paid"
            return False, "Цей Telegram charge_id уже прив'язаний до іншого замовлення."

        # Ensure the account row exists before posting the credit/subscription.
        con.execute(
            "INSERT INTO users(user_id,created_at) VALUES(?,?) ON CONFLICT(user_id) DO NOTHING",
            (user_id, now_iso()),
        )
        if order["product"] == "credits":
            con.execute("UPDATE users SET credits=credits+? WHERE user_id=?", (CREDIT_PACK_SIZE, user_id))
        elif order["product"] == "pro":
            row = con.execute("SELECT subscription_until FROM users WHERE user_id=?", (user_id,)).fetchone()
            current = None
            if row and row["subscription_until"]:
                try:
                    current = datetime.fromisoformat(row["subscription_until"])
                except (TypeError, ValueError):
                    current = None
            start_at = current if current and current > now_utc() else now_utc()
            until = start_at + timedelta(days=PRO_DAYS)
            con.execute("UPDATE users SET subscription_until=? WHERE user_id=?", (until.isoformat(), user_id))
        else:
            return False, "Невідомий тип товару."

        con.execute(
            "UPDATE orders SET status='paid', charge_id=? WHERE payload=? AND status='pending'",
            (charge_id, payload),
        )
        confirmed = con.execute(
            "SELECT status,charge_id FROM orders WHERE payload=?", (payload,)
        ).fetchone()
        if not confirmed or confirmed["status"] != "paid" or confirmed["charge_id"] != charge_id:
            raise RuntimeError("Order state could not be confirmed after applying payment")
        return True, order["product"]


def list_delivery_unknown_generations(limit: int = 10):
    safe_limit = max(1, min(20, int(limit)))
    with db_connect() as con:
        return con.execute(
            "SELECT generation_id,user_id,topic,slides,quota_mode,created_at "
            "FROM generation_history WHERE status='delivery_unknown' "
            "ORDER BY created_at ASC LIMIT ?",
            (safe_limit,),
        ).fetchall()


def resolve_delivery_unknown(generation_id: str, admin_id: int, resolution: str):
    """Resolve ambiguous Telegram delivery exactly once, with an auditable admin action."""
    if resolution not in ("delivered", "refund"):
        raise ValueError("resolution must be 'delivered' or 'refund'")
    refunded_credit = False
    with db_connect() as con:
        begin_write_transaction(con)
        row = con.execute(
            "SELECT user_id,quota_mode,quota_day,status FROM generation_history WHERE generation_id=?",
            (generation_id,),
        ).fetchone()
        if not row or row["status"] != "delivery_unknown":
            return None
        new_status = "completed" if resolution == "delivered" else "failed"
        changed = con.execute(
            "UPDATE generation_history SET status=?,delivery_resolution=?,resolved_by=?,resolved_at=?,"
            "lease_until=NULL,worker_id=NULL WHERE generation_id=? AND status='delivery_unknown' RETURNING generation_id",
            (new_status, resolution, admin_id, now_iso(), generation_id),
        ).fetchone()
        if not changed:
            return None
        con.execute("DELETE FROM generation_locks WHERE generation_id=?", (generation_id,))
        if resolution == "refund":
            _refund_quota_in_transaction(
                con, int(row["user_id"]), str(row["quota_mode"] or "none"), row["quota_day"]
            )
            refunded_credit = row["quota_mode"] == "credit"
        verified = con.execute(
            "SELECT status,delivery_resolution FROM generation_history WHERE generation_id=?",
            (generation_id,),
        ).fetchone()
        if not verified or verified["status"] != new_status or verified["delivery_resolution"] != resolution:
            raise RuntimeError("Administrator delivery resolution could not be confirmed")
        user_id = int(row["user_id"])
    if refunded_credit:
        schedule_balance_backup("адміністратор повернув кредит після невизначеної відправки")
    return {"user_id": user_id, "status": new_status, "resolution": resolution}


def create_report(generation_id: str, user_id: int, reason: str) -> str | None:
    with db_connect() as con:
        gen = con.execute(
            "SELECT * FROM generation_history WHERE generation_id=? AND user_id=?",
            (generation_id, user_id),
        ).fetchone()
        if not gen:
            return None
        existing = con.execute(
            "SELECT report_id FROM reports WHERE generation_id=? AND user_id=? AND status='pending'",
            (generation_id, user_id),
        ).fetchone()
        if existing:
            return existing["report_id"]
        report_id = uuid.uuid4().hex[:12]
        con.execute(
            "INSERT INTO reports(report_id,generation_id,user_id,reason,created_at) VALUES(?,?,?,?,?)",
            (report_id, generation_id, user_id, reason[:1000], now_iso()),
        )
        return report_id
 
 
def get_report(report_id: str):
    with db_connect() as con:
        return con.execute(
            "SELECT r.*, g.topic, g.design, g.slides, g.detail, g.language, g.source_chat_id, g.source_message_id "
            "FROM reports r LEFT JOIN generation_history g ON g.generation_id=r.generation_id "
            "WHERE r.report_id=?", (report_id,)
        ).fetchone()
 
 
def handle_report(report_id: str, admin_id: int, action: str):
    with db_connect() as con:
        begin_write_transaction(con)
        report = con.execute("SELECT * FROM reports WHERE report_id=?", (report_id,)).fetchone()
        if not report or report["status"] != "pending":
            return None
        if action == "refund":
            con.execute("UPDATE users SET credits=credits+1 WHERE user_id=?", (report["user_id"],))
            status = "refunded"
        else:
            status = "rejected"
        con.execute(
            "UPDATE reports SET status=?, handled_by=?, handled_at=? WHERE report_id=?",
            (status, admin_id, now_iso(), report_id),
        )
        return int(report["user_id"]), status
 
# ============================================================
# 2b. BACKUPS (SQLite без зовнішньої бази)
# ============================================================
BACKUP_TABLES = ("users", "daily_usage", "orders", "generation_history", "reports")
BALANCE_BACKUP_DEBOUNCE_SECONDS = 120
_balance_backup_task = None
_pending_balance_backup_reason = None
_APP_EVENT_LOOP = None


def make_db_snapshot():
    """Create a portable SQLite backup, including a remote Turso database if enabled."""
    snap = Path(tempfile.gettempdir()) / f"prezent_backup_{int(time.time())}_{uuid.uuid4().hex[:6]}.sqlite3"
    if not TURSO_ENABLED:
        src = sqlite3.connect(DATABASE_PATH, timeout=30)
        dst = sqlite3.connect(str(snap))
        try:
            src.backup(dst)
        finally:
            dst.close()
            src.close()
        return str(snap)

    dst = sqlite3.connect(str(snap), timeout=30)
    try:
        dst.executescript(DB_SCHEMA_SQL)
        with db_connect() as src:
            for table in BACKUP_TABLES:
                cursor = src.execute(f"SELECT * FROM {table}")
                description = cursor.description or ()
                columns = [item[0] for item in description]
                rows = cursor.fetchall()
                if not columns or not rows:
                    continue
                quoted_columns = ",".join('"' + name.replace('"', '""') + '"' for name in columns)
                placeholders = ",".join("?" for _ in columns)
                sql = f"INSERT OR REPLACE INTO {table} ({quoted_columns}) VALUES ({placeholders})"
                dst.executemany(sql, [tuple(row[name] for name in columns) for row in rows])
        dst.commit()
    except Exception:
        dst.close()
        snap.unlink(missing_ok=True)
        raise
    finally:
        try:
            dst.close()
        except Exception:
            pass
    return str(snap)


async def backup_to_admins(reason="auto"):
    if not ADMIN_IDS:
        log.warning("Database backup skipped: ADMIN_IDS is empty")
        return
    try:
        path = await asyncio.to_thread(make_db_snapshot)
    except Exception as exc:
        log.exception("Backup snapshot failed")
        await notify_admins_once(
            "database-backup-snapshot-failed",
            f"🚨 <b>Не вдалося створити резервну копію Prezent AI.</b>\n{html.escape(str(exc)[:500])}",
        )
        return
    if not path:
        return
    try:
        for admin_id in ADMIN_IDS:
            try:
                await bot.send_document(
                    admin_id,
                    FSInputFile(path, filename="prezentai_backup.sqlite3"),
                    caption=(
                        f"💾 Резервна копія бази ({reason}).\n"
                        "Для відновлення надішли файл боту з підписом /restore."
                    ),
                    disable_notification=True,
                )
            except Exception as exc:
                log.warning("Could not send backup to admin %s: %s", admin_id, exc)
                await notify_admins_once(
                    f"backup-upload-failed-{admin_id}",
                    f"⚠️ <b>Не вдалося надіслати резервну копію адміністратору.</b> ID: <code>{admin_id}</code>\n{html.escape(str(exc)[:300])}",
                )
    finally:
        Path(path).unlink(missing_ok=True)


async def _run_debounced_balance_backup():
    global _pending_balance_backup_reason, _balance_backup_task
    try:
        while _pending_balance_backup_reason:
            await asyncio.sleep(BALANCE_BACKUP_DEBOUNCE_SECONDS)
            reason = _pending_balance_backup_reason
            _pending_balance_backup_reason = None
            await backup_to_admins(reason or "зміна балансу")
    finally:
        _balance_backup_task = None
        # Catch a balance change that arrived during the previous upload.
        if _pending_balance_backup_reason:
            try:
                loop = asyncio.get_running_loop()
                _balance_backup_task = loop.create_task(_run_debounced_balance_backup())
            except RuntimeError:
                pass


def _queue_balance_backup(reason: str):
    global _pending_balance_backup_reason, _balance_backup_task
    _pending_balance_backup_reason = str(reason)[:120]
    try:
        loop = asyncio.get_running_loop()
    except RuntimeError:
        log.warning("Balance backup queue callback ran without an event loop")
        return
    if _balance_backup_task is None or _balance_backup_task.done():
        _balance_backup_task = loop.create_task(_run_debounced_balance_backup())


def schedule_balance_backup(reason="зміна балансу"):
    """Schedule backup on the application loop, including calls made from DB worker threads."""
    if DURABLE_DB_ENABLED:
        return
    try:
        loop = asyncio.get_running_loop()
    except RuntimeError:
        loop = _APP_EVENT_LOOP
    if not loop or not loop.is_running():
        log.warning("Balance changed but no application loop is available for automatic backup")
        return
    loop.call_soon_threadsafe(_queue_balance_backup, str(reason)[:120])


async def backup_loop():
    interval = 24 * 3600 if DURABLE_DB_ENABLED else 6 * 3600
    while True:
        await asyncio.sleep(interval)
        try:
            await backup_to_admins("авто, щоденно" if DURABLE_DB_ENABLED else "авто, кожні 6 год")
        except asyncio.CancelledError:
            raise
        except Exception as exc:
            log.exception("Scheduled database backup loop failed")
            await notify_admins_once(
                "database-backup-loop-failed",
                f"🚨 <b>Помилка автоматичного резервного копіювання.</b>\n{html.escape(str(exc)[:500])}",
            )


def validate_backup_file(path: str) -> dict:
    """Validate a backup before any live database is modified."""
    target = Path(path)
    if not target.is_file() or target.stat().st_size <= 0:
        raise ValueError("Резервна копія порожня або відсутня.")
    if target.stat().st_size > MAX_RESTORE_BACKUP_BYTES:
        raise ValueError("Резервна копія перевищує дозволений розмір 45 МБ.")
    uri = f"file:{target.resolve().as_posix()}?mode=ro"
    src = sqlite3.connect(uri, uri=True, timeout=15)
    try:
        check_rows = src.execute("PRAGMA quick_check").fetchall()
        if not check_rows or any(str(row[0]).lower() != "ok" for row in check_rows):
            raise ValueError("SQLite quick_check виявив пошкодження резервної копії.")
        table_names = {
            row[0] for row in src.execute("SELECT name FROM sqlite_master WHERE type='table'").fetchall()
        }
        required_columns = {
            "users": {"user_id", "credits", "subscription_until"},
            "daily_usage": {"user_id", "day", "used"},
            "orders": {"payload", "user_id", "product", "amount", "status", "charge_id"},
            "generation_history": {"generation_id", "user_id", "quota_mode", "status", "created_at"},
            "reports": {"report_id", "generation_id", "user_id", "status"},
        }
        missing = set(required_columns) - table_names
        if missing:
            raise ValueError("У резервній копії відсутні таблиці: " + ", ".join(sorted(missing)))
        for table, expected in required_columns.items():
            actual = {row[1] for row in src.execute(f'PRAGMA table_info("{table}")').fetchall()}
            absent = expected - actual
            if absent:
                raise ValueError(f"У таблиці {table} відсутні колонки: {', '.join(sorted(absent))}")
        counts = {
            table: int(src.execute(f'SELECT COUNT(*) FROM "{table}"').fetchone()[0])
            for table in required_columns
        }
        counts["credits"] = int(src.execute("SELECT COALESCE(SUM(credits),0) FROM users").fetchone()[0])
        return counts
    finally:
        src.close()


def _apply_backup_restore(path: str):
    """Apply an already validated snapshot to the configured storage."""
    if not TURSO_ENABLED:
        src = sqlite3.connect(path, timeout=30)
        try:
            dst = sqlite3.connect(DATABASE_PATH, timeout=30)
            try:
                src.backup(dst)
                dst.commit()
            finally:
                dst.close()
        finally:
            src.close()
        # Older backups may not contain lease columns/table introduced by this patch.
        # Run idempotent migrations first, then discard any historical worker locks.
        init_db()
        with db_connect() as con:
            begin_write_transaction(con)
            con.execute("DELETE FROM generation_locks")
            con.execute(
                "UPDATE generation_history SET lease_until=NULL,worker_id=NULL "
                "WHERE status='pending'"
            )
        return

    src = sqlite3.connect(path, timeout=30)
    src.row_factory = sqlite3.Row
    try:
        with db_connect() as dst:
            dst.execute("BEGIN")
            # Locks are ephemeral worker leases and must never be restored from a backup.
            dst.execute("DELETE FROM generation_locks")
            for table in reversed(BACKUP_TABLES):
                dst.execute(f"DELETE FROM {table}")
            source_tables = {
                row[0] for row in src.execute("SELECT name FROM sqlite_master WHERE type='table'").fetchall()
            }
            for table in BACKUP_TABLES:
                if table not in source_tables:
                    continue
                cursor = src.execute(f'SELECT * FROM "{table}"')
                columns = [item[0] for item in (cursor.description or ())]
                rows = cursor.fetchall()
                if not columns:
                    continue
                quoted_columns = ",".join('"' + name.replace('"', '""') + '"' for name in columns)
                placeholders = ",".join("?" for _ in columns)
                sql = f'INSERT INTO "{table}" ({quoted_columns}) VALUES ({placeholders})'
                for row in rows:
                    dst.execute(sql, tuple(row[name] for name in columns))
            dst.execute(
                "UPDATE generation_history SET lease_until=NULL,worker_id=NULL WHERE status='pending'"
            )
    finally:
        src.close()


def _database_restore_counts() -> dict:
    with db_connect() as con:
        result = {
            table: int(con.execute(f'SELECT COUNT(*) AS n FROM "{table}"').fetchone()["n"])
            for table in ("users", "daily_usage", "orders", "generation_history", "reports")
        }
        result["credits"] = int(
            con.execute("SELECT COALESCE(SUM(credits),0) AS n FROM users").fetchone()["n"]
        )
    return result


def has_active_generation_leases() -> bool:
    """Check cross-worker generation locks before destructive maintenance operations."""
    with db_connect() as con:
        row = con.execute(
            "SELECT 1 FROM generation_locks WHERE lease_until>? LIMIT 1", (now_iso(),)
        ).fetchone()
    return bool(row)


def restore_db_from(path: str):
    """Validate, snapshot, restore, verify, and roll back automatically on failure."""
    incoming = validate_backup_file(path)
    rollback_path = make_db_snapshot()
    try:
        _apply_backup_restore(path)
        actual = _database_restore_counts()
        expected = {key: incoming[key] for key in actual}
        if actual != expected:
            raise RuntimeError(
                "Перевірка після відновлення не пройдена: кількість записів у таблицях не збігається. "
                f"Очікувалося {expected}, отримано {actual}."
            )
        log.info("Database restore verified: %s", actual)
    except Exception:
        log.exception("Database restore failed; rolling back to the pre-restore snapshot")
        try:
            _apply_backup_restore(rollback_path)
        except Exception:
            log.critical("CRITICAL: automatic database restore rollback also failed", exc_info=True)
            raise RuntimeError("Відновлення не вдалося, а автоматичний відкат також завершився помилкою. Перевір базу вручну.")
        raise
    finally:
        try:
            Path(rollback_path).unlink(missing_ok=True)
        except OSError:
            log.warning("Could not remove temporary pre-restore snapshot %s", rollback_path)

# ============================================================
# 3. KEYBOARDS AND THEMES
# ============================================================
# Conversation language is selected once on the first /start and can be changed with /language.
BUTTONS = {
    "start": {"uk": "✨ Старт", "ru": "✨ Старт", "en": "✨ Start"},
    "restart": {"uk": "↻ Перезапуск", "ru": "↻ Перезапуск", "en": "↻ Restart"},
    "create": {"uk": "🎨 Створити презентацію", "ru": "🎨 Создать презентацию", "en": "🎨 Create presentation"},
    "home": {"uk": "⌂ Головне меню", "ru": "⌂ Главное меню", "en": "⌂ Main menu"},
    "cancel": {"uk": "✕ Скасувати", "ru": "✕ Отменить", "en": "✕ Cancel"},
    "features": {"uk": "💎 Можливості", "ru": "💎 Возможности", "en": "💎 Features"},
    "help": {"uk": "🆘 Допомога", "ru": "🆘 Помощь", "en": "🆘 Help"},
    "profile": {"uk": "👤 Профіль", "ru": "👤 Профиль", "en": "👤 Profile"},
    "shop": {"uk": "⭐ Магазин", "ru": "⭐ Магазин", "en": "⭐ Shop"},
    "sub": {"uk": "👑 PRO-підписка", "ru": "👑 PRO-подписка", "en": "👑 PRO subscription"},
}
BUTTON_TEXTS = {key: tuple(values.values()) for key, values in BUTTONS.items()}

UI_TEXT = {
    "choose_language": {
        "uk": "👋 <b>Вітаю у Prezent AI!</b>\n\nСпершу обери мову, якою ми будемо спілкуватися. Її можна змінити будь-коли командою /language.",
        "ru": "👋 <b>Добро пожаловать в Prezent AI!</b>\n\nСначала выбери язык нашего общения. Его можно изменить в любой момент командой /language.",
        "en": "👋 <b>Welcome to Prezent AI!</b>\n\nFirst, choose the language we should use to communicate. You can change it anytime with /language.",
    },
    "language_saved": {
        "uk": "✅ Мову спілкування збережено: <b>Українська</b>.",
        "ru": "✅ Язык общения сохранён: <b>Русский</b>.",
        "en": "✅ Conversation language saved: <b>English</b>.",
    },
    "welcome": {
        "uk": "✦ <b>PREZENT AI</b>\n<i>Ваш персональний простір для презентацій</i>\n\nВітаю, <b>{name}</b>! 👋\n\nПеретворюю вашу ідею на продуману презентацію: з логічною історією, змістовними тезами та візуальним стилем, який відповідає темі.\n\n<b>МОЖЛИВОСТІ</b>\n🎨 10 окремих візуальних стилів\n📊 6, 8, 10 або 12 слайдів\n🧠 AI-структура та унікальний контент\n📄 PowerPoint + PDF\n🎞️ Плавні переходи у PowerPoint\n\n🎁 <b>3 безкоштовні презентації щодня</b>\n⭐ Додаткові можливості — у магазині\n\nНатисніть <b>«🎨 Створити презентацію»</b>, щоб почати.",
        "ru": "✦ <b>PREZENT AI</b>\n<i>Ваше персональное пространство для презентаций</i>\n\nПривет, <b>{name}</b>! 👋\n\nЯ превращаю вашу идею в продуманную презентацию: с логичной историей, содержательными тезисами и визуальным стилем, подходящим теме.\n\n<b>ВОЗМОЖНОСТИ</b>\n🎨 10 самостоятельных визуальных стилей\n📊 6, 8, 10 или 12 слайдов\n🧠 AI-структура и уникальный контент\n📄 PowerPoint + PDF\n🎞️ Плавные переходы в PowerPoint\n\n🎁 <b>3 бесплатные презентации ежедневно</b>\n⭐ Дополнительные возможности — в магазине\n\nНажми <b>«🎨 Создать презентацию»</b>, чтобы начать.",
        "en": "✦ <b>PREZENT AI</b>\n<i>Your personal presentation studio</i>\n\nHello, <b>{name}</b>! 👋\n\nI turn your idea into a well-crafted presentation with a clear narrative, meaningful content, and a visual style that fits the subject.\n\n<b>FEATURES</b>\n🎨 10 distinct visual styles\n📊 6, 8, 10, or 12 slides\n🧠 AI-powered structure and original content\n📄 PowerPoint + PDF\n🎞️ Smooth PowerPoint transitions\n\n🎁 <b>3 free presentations every day</b>\n⭐ More options are available in the shop\n\nPress <b>«🎨 Create presentation»</b> to begin.",
    },
    "topic_prompt": {
        "uk": "📝 <b>Про що буде презентація?</b>\n\nОпиши тему одним або кількома реченнями. Додай аудиторію чи мету, якщо це важливо.\n\n<i>Наприклад: «Як штучний інтелект змінює навчання у школах».</i>\nМаксимум 500 символів.",
        "ru": "📝 <b>О чём будет презентация?</b>\n\nОпиши тему одним или несколькими предложениями. При необходимости укажи аудиторию или цель.\n\n<i>Например: «Как искусственный интеллект меняет обучение в школах».</i>\nМаксимум 500 символов.",
        "en": "📝 <b>What is your presentation about?</b>\n\nDescribe the topic in one or more sentences. Add the audience or goal if relevant.\n\n<i>Example: “How artificial intelligence is changing education in schools”.</i>\nMaximum 500 characters.",
    },
    "topic_short": {"uk": "Тема закоротка. Напиши щонайменше 3 символи.", "ru": "Тема слишком короткая. Напиши хотя бы 3 символа.", "en": "The topic is too short. Please enter at least 3 characters."},
    "topic_long": {"uk": "Тема задовга. Максимум 500 символів.", "ru": "Тема слишком длинная. Максимум 500 символов.", "en": "The topic is too long. Maximum 500 characters."},
    "slides_prompt": {"uk": "📊 <b>Скільки слайдів потрібно?</b>\n\nОбери кількість:", "ru": "📊 <b>Сколько слайдов нужно?</b>\n\nВыбери количество:", "en": "📊 <b>How many slides do you need?</b>\n\nChoose a slide count:"},
    "slides_first": {"uk": "Спочатку введи тему презентації.", "ru": "Сначала введи тему презентации.", "en": "Please enter the presentation topic first."},
    "slides_invalid": {"uk": "Неправильна кількість слайдів.", "ru": "Неверное количество слайдов.", "en": "Invalid slide count."},
    "slides_choose": {"uk": "Обери 6, 8, 10 або 12.", "ru": "Выбери 6, 8, 10 или 12.", "en": "Choose 6, 8, 10, or 12."},
    "output_language_prompt": {"uk": "🌍 <b>Якою мовою має бути презентація?</b>", "ru": "🌍 <b>На каком языке создать презентацию?</b>", "en": "🌍 <b>What language should the presentation use?</b>"},
    "detail_prompt": {
        "uk": "📚 <b>Який рівень деталізації потрібен?</b>\n\n⚡ <b>Коротко</b> — найважливіше, 2–3 тези на слайд.\n⚖️ <b>Середньо</b> — збалансоване пояснення, 3–4 тези.\n📚 <b>Детально</b> — більше пояснень, прикладів і нотаток доповідача, 4–5 тез.",
        "ru": "📚 <b>Какой уровень детализации нужен?</b>\n\n⚡ <b>Кратко</b> — только главное, 2–3 тезиса на слайд.\n⚖️ <b>Средне</b> — сбалансированное объяснение, 3–4 тезиса.\n📚 <b>Подробно</b> — больше объяснений, примеров и заметок докладчика, 4–5 тезисов.",
        "en": "📚 <b>How much detail do you need?</b>\n\n⚡ <b>Short</b> — key points only, 2–3 bullets per slide.\n⚖️ <b>Medium</b> — balanced explanations, 3–4 bullets.\n📚 <b>Detailed</b> — fuller explanations, examples, and speaker notes, 4–5 bullets.",
    },
    "detail_invalid": {
        "uk": "Обери один із трьох рівнів: Коротко, Середньо або Детально.",
        "ru": "Выбери один из трёх уровней: Кратко, Средне или Подробно.",
        "en": "Choose one of the three levels: Short, Medium, or Detailed.",
    },
    "photos_prompt": {
        "uk": "🖼️ <b>Додати фотографії до презентації?</b>\n\n📷 <b>З фото</b> — до 5 реальних зображень з відкритих джерел, з указанням автора й ліцензії.\n🚫 <b>Без фото</b> — лише текст і дизайн; створення буде трохи швидшим.",
        "ru": "🖼️ <b>Добавить фотографии в презентацию?</b>\n\n📷 <b>С фото</b> — до 5 реальных изображений из открытых источников, с указанием автора и лицензии.\n🚫 <b>Без фото</b> — только текст и дизайн; создание будет чуть быстрее.",
        "en": "🖼️ <b>Add photos to the presentation?</b>\n\n📷 <b>With photos</b> — up to 5 real images from open sources, with author and license credits.\n🚫 <b>No photos</b> — text and design only; generation is a bit faster.",
    },
    "photos_invalid": {
        "uk": "Обери: З фото або Без фото.",
        "ru": "Выбери: С фото или Без фото.",
        "en": "Choose: With photos or No photos.",
    },
    "design_prompt": {"uk": "🎨 <b>Обери візуальний стиль.</b> Кожен стиль має власну композицію, типографіку та оформлення:", "ru": "🎨 <b>Выбери визуальный стиль.</b> У каждого стиля своя композиция, типографика и оформление:", "en": "🎨 <b>Choose a visual style.</b> Each style has its own composition, typography, and visual treatment:"},
    "creating": {"uk": "Ваша презентація створюється", "ru": "Ваша презентация создаётся", "en": "Your presentation is being created"},
    "please_wait": {"uk": "AI розробляє структуру та оформлює слайди…", "ru": "AI разрабатывает структуру и оформляет слайды…", "en": "AI is shaping the story and designing the slides…"},
    "generation_complete": {"uk": "✅ <b>Презентація готова!</b>\n\nPowerPoint і PDF надіслано 📤", "ru": "✅ <b>Презентация готова!</b>\n\nPowerPoint и PDF отправлены 📤", "en": "✅ <b>Your presentation is ready!</b>\n\nPowerPoint and PDF have been sent 📤"},
    "generation_pptx_only": {"uk": "✅ <b>PowerPoint створено й надіслано!</b>\n\n⚠️ PDF цього разу не вдалося надіслати.", "ru": "✅ <b>PowerPoint создан и отправлен!</b>\n\n⚠️ В этот раз не удалось отправить PDF.", "en": "✅ <b>PowerPoint created and sent!</b>\n\n⚠️ The PDF could not be sent this time."},
    "ready_next": {"uk": "✅ <b>Готово!</b>\n\nМожеш створити наступну презентацію кнопкою нижче.", "ru": "✅ <b>Готово!</b>\n\nМожешь создать следующую презентацию кнопкой ниже.", "en": "✅ <b>Done!</b>\n\nYou can create another presentation using the button below."},
    "repeat_last": {"uk": "🔁 Створити ще раз з тими ж параметрами", "ru": "🔁 Создать ещё раз с теми же параметрами", "en": "🔁 Create again with the same settings"},
    "cancelled": {"uk": "❌ <b>Створення зупинено.</b>\n\nЛіміт або кредит повернуто.", "ru": "❌ <b>Создание остановлено.</b>\n\nЛимит или кредит возвращён.", "en": "❌ <b>Creation cancelled.</b>\n\nYour quota or credit has been restored."},
    "already_sent": {"uk": "ℹ️ <b>Презентацію вже надіслано.</b>\n\nЛіміт використано.", "ru": "ℹ️ <b>Презентация уже отправлена.</b>\n\nЛимит использован.", "en": "ℹ️ <b>The presentation has already been sent.</b>\n\nThe quota has been used."},
    "ai_failed": {"uk": "❌ <b>AI не зміг створити презентацію.</b>\n\nПричина: {reason}\n\nКредит або денний ліміт повернуто.", "ru": "❌ <b>AI не смог создать презентацию.</b>\n\nПричина: {reason}\n\nКредит или дневной лимит возвращён.", "en": "❌ <b>AI could not create the presentation.</b>\n\nReason: {reason}\n\nYour credit or daily quota has been restored."},
    "generic_failed": {"uk": "❌ <b>Не вдалося створити презентацію.</b>\n\nКод помилки: <code>{error_id}</code>\nКредит або денний ліміт повернуто.", "ru": "❌ <b>Не удалось создать презентацию.</b>\n\nКод ошибки: <code>{error_id}</code>\nКредит или дневной лимит возвращён.", "en": "❌ <b>Could not create the presentation.</b>\n\nError code: <code>{error_id}</code>\nYour credit or daily quota has been restored."},
    "unknown": {"uk": "Не зрозумів повідомлення. Скористайся кнопками меню нижче 👇", "ru": "Не понял сообщение. Используй кнопки меню ниже 👇", "en": "I didn't understand that message. Use the menu buttons below 👇"},
    "dialog_restarted": {"uk": "🔄 <b>Діалог перезапущено!</b>\n\nМожеш почати нову презентацію.", "ru": "🔄 <b>Диалог перезапущен!</b>\n\nМожешь начать новую презентацию.", "en": "🔄 <b>Conversation restarted!</b>\n\nYou can start a new presentation."},
    "action_cancelled": {"uk": "❌ Поточну дію скасовано.", "ru": "❌ Текущее действие отменено.", "en": "❌ The current action was cancelled."},
    "create_cancelled": {"uk": "❌ Створення скасовано.", "ru": "❌ Создание отменено.", "en": "❌ Creation cancelled."},
    "features": {
        "uk": "✨ <b>Можливості Prezent AI</b>\n\n🎨 10 візуально різних стилів: мінімалізм, футуризм, бізнес, шкільний, luxury, журнал, nature, vibrant, classic та mono.\n🧩 Різні композиції: картки, таймлайн, кроки, порівняння, цитати й фокус-слайди.\n📊 6/8/10/12 слайдів • 🇺🇦 🇬🇧 🇷🇺 • PPTX + PDF.\n🧠 Кожна генерація отримує нову структуру й творчий ракурс.",
        "ru": "✨ <b>Возможности Prezent AI</b>\n\n🎨 10 визуально разных стилей: минимализм, футуризм, бизнес, школьный, luxury, журнал, nature, vibrant, classic и mono.\n🧩 Разные композиции: карточки, таймлайн, шаги, сравнение, цитаты и акцентные слайды.\n📊 6/8/10/12 слайдов • 🇺🇦 🇬🇧 🇷🇺 • PPTX + PDF.\n🧠 Каждая генерация получает новую структуру и творческий ракурс.",
        "en": "✨ <b>Prezent AI features</b>\n\n🎨 10 visually distinct styles: minimal, futuristic, business, school, luxury, magazine, nature, vibrant, classic, and monochrome.\n🧩 Different compositions: cards, timelines, steps, comparisons, quotes, and focus slides.\n📊 6/8/10/12 slides • 🇺🇦 🇬🇧 🇷🇺 • PPTX + PDF.\n🧠 Every generation gets a fresh structure and creative angle.",
    },
    "help": {
        "uk": "❓ <b>Допомога</b>\n\n🎨 Створити презентацію — запусти генератор.\n👤 Профіль — ліміти, кредити та PRO.\n⭐ Магазин — придбати генерації.\n👑 PRO — до 10 презентацій на день протягом 30 днів.\n✕ Скасувати — зупинити поточну дію.\n🌐 /language — змінити мову спілкування.\n/start — головне меню • /shop — магазин • /id — Telegram ID",
        "ru": "❓ <b>Помощь</b>\n\n🎨 Создать презентацию — запусти генератор.\n👤 Профиль — лимиты, кредиты и PRO.\n⭐ Магазин — купить генерации.\n👑 PRO — до 10 презентаций в день в течение 30 дней.\n✕ Отменить — остановить текущее действие.\n🌐 /language — изменить язык общения.\n/start — главное меню • /shop — магазин • /id — Telegram ID",
        "en": "❓ <b>Help</b>\n\n🎨 Create presentation — start the generator.\n👤 Profile — quotas, credits, and PRO status.\n⭐ Shop — purchase extra generations.\n👑 PRO — up to 10 presentations per day for 30 days.\n✕ Cancel — stop the current action.\n🌐 /language — change conversation language.\n/start — home • /shop — shop • /id — Telegram ID",
    },
    "shop": {
        "uk": "✦ <b>PREZENT AI / МАГАЗИН</b>\n<i>Оберіть формат доступу</i>\n\n🎟️ <b>ПАКЕТ КРЕДИТІВ</b>\n✨ {credits} додаткових генерацій — <b>{credit_price} ⭐</b>\n💡 Кредити не згорають наприкінці дня.\n\n👑 <b>ПІДПИСКА PRO</b>\n🚀 До {pro_limit} презентацій щодня протягом {pro_days} днів — <b>{pro_price} ⭐</b>\n\n🔒 Оплата безпосередньо через Telegram Stars.",
        "ru": "✦ <b>PREZENT AI / МАГАЗИН</b>\n<i>Выбери подходящий вариант доступа</i>\n\n🎟️ <b>ПАКЕТ КРЕДИТОВ</b>\n✨ {credits} дополнительных генераций — <b>{credit_price} ⭐</b>\n💡 Кредиты не сгорают в конце дня.\n\n👑 <b>ПОДПИСКА PRO</b>\n🚀 До {pro_limit} презентаций ежедневно в течение {pro_days} дней — <b>{pro_price} ⭐</b>\n\n🔒 Оплата напрямую через Telegram Stars.",
        "en": "✦ <b>PREZENT AI / SHOP</b>\n<i>Choose the access option that suits you</i>\n\n🎟️ <b>CREDIT PACK</b>\n✨ {credits} extra generations — <b>{credit_price} ⭐</b>\n💡 Credits do not expire at the end of the day.\n\n👑 <b>PRO SUBSCRIPTION</b>\n🚀 Up to {pro_limit} presentations daily for {pro_days} days — <b>{pro_price} ⭐</b>\n\n🔒 Pay directly with Telegram Stars.",
    },
    "subscription": {
        "uk": "👑 <b>Підписка PRO</b>\n\n✅ До {limit} презентацій на день\n✅ Діє {days} днів\n✅ Оплата один раз, без автоматичного продовження.\n\nЦіна: <b>{price} ⭐</b>.",
        "ru": "👑 <b>Подписка PRO</b>\n\n✅ До {limit} презентаций в день\n✅ Действует {days} дней\n✅ Разовая оплата, без автоматического продления.\n\nЦена: <b>{price} ⭐</b>.",
        "en": "👑 <b>PRO subscription</b>\n\n✅ Up to {limit} presentations per day\n✅ Valid for {days} days\n✅ One-time payment, no automatic renewal.\n\nPrice: <b>{price} ⭐</b>.",
    },
    "limit_reached": {
        "uk": "⛔ <b>Ліміт на сьогодні вичерпано.</b>\n\nБез PRO доступно {limit} презентації на день. Можеш придбати {credits} додаткових генерацій за {price} ⭐ або активувати PRO.",
        "ru": "⛔ <b>Дневной лимит исчерпан.</b>\n\nБез PRO доступно {limit} презентации в день. Можно купить {credits} дополнительных генераций за {price} ⭐ или активировать PRO.",
        "en": "⛔ <b>Today's quota has been used.</b>\n\nWithout PRO, you can create {limit} presentations per day. Buy {credits} extra generations for {price} ⭐ or activate PRO.",
    },
    "file_caption": {
        "uk": "📊 <b>{title}</b>\n🎨 Стиль: {style}\n📚 Деталізація: {detail}\nФормат: PowerPoint • плавні Fade-переходи\n\nЯкщо помітиш проблему, натисни кнопку нижче — адміністратор перевірить файл.",
        "ru": "📊 <b>{title}</b>\n🎨 Стиль: {style}\n📚 Детализация: {detail}\nФормат: PowerPoint • плавные Fade-переходы\n\nЕсли заметишь проблему, нажми кнопку ниже — администратор проверит файл.",
        "en": "📊 <b>{title}</b>\n🎨 Style: {style}\n📚 Detail: {detail}\nFormat: PowerPoint • smooth Fade transitions\n\nIf you notice a problem, press the button below and an administrator will review the file.",
    },
    "pdf_caption": {"uk": "📄 <b>PDF</b> • версія для перегляду та друку.", "ru": "📄 <b>PDF</b> • версия для просмотра и печати.", "en": "📄 <b>PDF</b> • view and print edition."},
    "pdf_failed": {"uk": "PPTX готовий, але PDF цього разу надіслати не вдалося.", "ru": "PPTX готов, но PDF в этот раз отправить не удалось.", "en": "The PPTX is ready, but the PDF could not be sent this time."},
    "retry_menu": {"uk": "Можеш повторити спробу через меню.", "ru": "Можешь повторить попытку через меню.", "en": "You can try again from the menu."},
    "check_logs": {"uk": "Перевір Render Logs або спробуй пізніше.", "ru": "Проверь Render Logs или попробуй позже.", "en": "Check the Render logs or try again later."},
    "payment_error": {"uk": "⚠️ Платіж отримано, але покупку не вдалося автоматично зарахувати. Напиши адміністратору.", "ru": "⚠️ Платёж получен, но покупку не удалось автоматически зачислить. Напиши администратору.", "en": "⚠️ Payment received, but the purchase could not be applied automatically. Please contact the administrator."},
    "payment_credits": {"uk": "🌟 <b>ОПЛАТА УСПІШНА!</b>\n\n🎟️ На рахунок додано <b>+{credits} генерацій</b>.\n🚀 Вони автоматично використаються після денного ліміту.\n\n{profile}", "ru": "🌟 <b>ОПЛАТА УСПЕШНА!</b>\n\n🎟️ На счёт добавлено <b>+{credits} генераций</b>.\n🚀 Они автоматически будут использованы после дневного лимита.\n\n{profile}", "en": "🌟 <b>PAYMENT SUCCESSFUL!</b>\n\n🎟️ <b>+{credits} generations</b> have been added to your account.\n🚀 They will be used automatically after your daily quota.\n\n{profile}"},
    "payment_pro": {"uk": "👑 <b>PREZENT AI PRO АКТИВОВАНО!</b>\n\n💎 Підписка діє {days} днів.\n📊 Денний ліміт: до {limit} презентацій.\n\n{profile}", "ru": "👑 <b>PREZENT AI PRO АКТИВИРОВАНО!</b>\n\n💎 Подписка действует {days} дней.\n📊 Дневной лимит: до {limit} презентаций.\n\n{profile}", "en": "👑 <b>PREZENT AI PRO ACTIVATED!</b>\n\n💎 Subscription is active for {days} days.\n📊 Daily quota: up to {limit} presentations.\n\n{profile}"},
    "payment_already": {"uk": "✅ Цю покупку вже було зараховано. Повторно кредити не додавалися.", "ru": "✅ Эта покупка уже была зачислена. Кредиты повторно не добавлялись.", "en": "✅ This purchase has already been applied. No duplicate credits were added."},
    "report_prompt": {"uk": "⚠️ <b>Повідомлення про проблему</b>\n\nОпиши, що саме не так із презентацією: помилки, зіпсовані слайди, неправильний файл або інша проблема.\n\nАдміністратор отримає презентацію та перевірить її.", "ru": "⚠️ <b>Сообщение о проблеме</b>\n\nОпиши, что именно не так с презентацией: ошибки, испорченные слайды, неправильный файл или другая проблема.\n\nАдминистратор получит презентацию и проверит её.", "en": "⚠️ <b>Report a problem</b>\n\nDescribe what is wrong with the presentation: errors, damaged slides, an incorrect file, or another issue.\n\nThe administrator will receive the presentation and review it."},
    "report_too_short": {"uk": "Опиши проблему трохи детальніше (від 4 символів).", "ru": "Опиши проблему подробнее (минимум 4 символа).", "en": "Please describe the problem in more detail (at least 4 characters)."},
    "report_missing": {"uk": "Не знайшов цю презентацію в історії. Напиши адміністратору.", "ru": "Не нашёл эту презентацию в истории. Напиши администратору.", "en": "I couldn't find this presentation in the history. Please contact the administrator."},
    "report_sent": {"uk": "✅ Скаргу та файл презентації надіслано адміністратору. Після перевірки він зможе повернути 1 кредит.", "ru": "✅ Жалоба и файл презентации отправлены администратору. После проверки он сможет вернуть 1 кредит.", "en": "✅ The report and presentation file were sent to the administrator. After review, one credit may be refunded."},
    "report_send_error": {"uk": "⚠️ Скаргу збережено, але адміністратор не отримав повідомлення. Перевір ADMIN_IDS у Render.", "ru": "⚠️ Жалоба сохранена, но администратор не получил уведомление. Проверь ADMIN_IDS в Render.", "en": "⚠️ The report was saved, but the administrator was not notified. Check ADMIN_IDS in Render."},
    "refund_notice": {"uk": "✅ Адміністратор перевірив презентацію та повернув <b>1 генераційний кредит</b>. Він доступний у твоєму профілі.", "ru": "✅ Администратор проверил презентацию и вернул <b>1 кредит генерации</b>. Он доступен в твоём профиле.", "en": "✅ The administrator reviewed the presentation and refunded <b>1 generation credit</b>. It is now available in your profile."},
    "reject_notice": {"uk": "ℹ️ Адміністратор перевірив скаргу. Повернення кредиту не було схвалено.", "ru": "ℹ️ Администратор проверил жалобу. Возврат кредита не был одобрен.", "en": "ℹ️ The administrator reviewed the report. A credit refund was not approved."},
}

LANGUAGE_NAMES = {
    "uk": {"uk": "Українська", "ru": "Украинский", "en": "Ukrainian"},
    "ru": {"uk": "Російська", "ru": "Русский", "en": "Russian"},
    "en": {"uk": "Англійська", "ru": "Английский", "en": "English"},
}


def tr(language: str | None, key: str, **kwargs) -> str:
    language = language if language in ("uk", "ru", "en") else "uk"
    value = UI_TEXT.get(key, {}).get(language) or UI_TEXT.get(key, {}).get("uk", key)
    return value.format(**kwargs) if kwargs else value


def get_button(key: str, language: str = "uk") -> str:
    language = language if language in ("uk", "ru", "en") else "uk"
    return BUTTONS[key][language]


async def dismiss_choice_message(callback: CallbackQuery):
    """Remove an inline choice prompt after a valid selection."""
    try:
        await callback.message.delete()
    except Exception as exc:
        log.debug("Could not delete the consumed choice prompt: %s", exc)


def main_keyboard(language: str = "uk"):
    language = language if language in ("uk", "ru", "en") else "uk"
    return ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text=get_button("create", language))],
            [KeyboardButton(text=get_button("profile", language)), KeyboardButton(text=get_button("shop", language))],
            [KeyboardButton(text=get_button("sub", language)), KeyboardButton(text=get_button("features", language))],
            [KeyboardButton(text=get_button("help", language)), KeyboardButton(text=get_button("home", language))],
            [KeyboardButton(text=get_button("start", language)), KeyboardButton(text=get_button("restart", language))],
            [KeyboardButton(text=get_button("cancel", language))],
        ], resize_keyboard=True, is_persistent=True,
        input_field_placeholder={"uk": "Оберіть дію…", "ru": "Выберите действие…", "en": "Choose an action…"}[language],
    )


def user_language(user_id: int) -> str:
    return get_user_language(user_id) or "uk"


MENU_ACTIONS = {
    label: key
    for key, variants in BUTTONS.items()
    for label in variants.values()
}


def interface_language_keyboard(language: str | None = None):
    labels = {
        "uk": ("🇺🇦 Українська", "🇷🇺 Русский", "🇬🇧 English"),
        "ru": ("🇺🇦 Украинский", "🇷🇺 Русский", "🇬🇧 English"),
        "en": ("🇺🇦 Ukrainian", "🇷🇺 Russian", "🇬🇧 English"),
    }[language if language in ("uk", "ru", "en") else "uk"]
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text=labels[0], callback_data="ui_lang:uk"), InlineKeyboardButton(text=labels[1], callback_data="ui_lang:ru")],
        [InlineKeyboardButton(text=labels[2], callback_data="ui_lang:en")],
    ])


# For the presentation's output language (separate from the bot's interface language).
def language_keyboard(interface_language: str = "uk"):
    labels = {
        "uk": ("🇺🇦 Українська", "🇬🇧 English", "🇷🇺 Російська"),
        "ru": ("🇺🇦 Украинский", "🇬🇧 English", "🇷🇺 Русский"),
        "en": ("🇺🇦 Ukrainian", "🇬🇧 English", "🇷🇺 Russian"),
    }[interface_language if interface_language in ("uk", "ru", "en") else "uk"]
    return InlineKeyboardMarkup(inline_keyboard=[[
        InlineKeyboardButton(text=labels[0], callback_data="lang:uk"),
        InlineKeyboardButton(text=labels[1], callback_data="lang:en")],
        [InlineKeyboardButton(text=labels[2], callback_data="lang:ru")],
    ])


def detail_keyboard(language: str = "uk"):
    labels = {
        "uk": ("⚡ Коротко", "⚖️ Середньо", "📚 Детально"),
        "ru": ("⚡ Кратко", "⚖️ Средне", "📚 Подробно"),
        "en": ("⚡ Short", "⚖️ Medium", "📚 Detailed"),
    }.get(language, ("⚡ Коротко", "⚖️ Середньо", "📚 Детально"))
    keys = ("short", "medium", "detailed")
    return InlineKeyboardMarkup(inline_keyboard=[[
        InlineKeyboardButton(text=labels[i], callback_data=f"detail:{keys[i]}")
        for i in range(3)
    ]])


def photos_keyboard(language: str = "uk"):
    labels = {
        "uk": ("📷 З фото", "🚫 Без фото"),
        "ru": ("📷 С фото", "🚫 Без фото"),
        "en": ("📷 With photos", "🚫 No photos"),
    }.get(language, ("📷 З фото", "🚫 Без фото"))
    return InlineKeyboardMarkup(inline_keyboard=[[
        InlineKeyboardButton(text=labels[0], callback_data="photos:yes"),
        InlineKeyboardButton(text=labels[1], callback_data="photos:no"),
    ]])


def slides_keyboard():
    return InlineKeyboardMarkup(inline_keyboard=[[
        InlineKeyboardButton(text=str(n), callback_data=f"slides:{n}") for n in (6, 8, 10, 12)
    ]])


def audience_keyboard(language="uk"):
    labels = {
        "uk": ["🎓 Для навчання", "💼 Для бізнесу", "🌍 Широка аудиторія", "🎤 Для виступу"],
        "ru": ["🎓 Для учёбы", "💼 Для бизнеса", "🌍 Широкая аудитория", "🎤 Для выступления"],
        "en": ["🎓 For learning", "💼 For business", "🌍 General audience", "🎤 For a speech"],
    }.get(language, ["🎓 For learning", "💼 For business", "🌍 General audience", "🎤 For a speech"])
    keys = ("students", "business", "general", "speaker")
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text=labels[i], callback_data=f"audience:{keys[i]}")]
        for i in range(4)
    ])


def outline_review_keyboard(review_id, language="uk", allow_revision=True):
    labels = {
        "uk": ("✅ Підтвердити", "✏️ Змінити слайд"),
        "ru": ("✅ Подтвердить", "✏️ Изменить слайд"),
        "en": ("✅ Approve", "✏️ Revise a slide"),
    }.get(language, ("✅ Approve", "✏️ Revise a slide"))
    rows = [[InlineKeyboardButton(text=labels[0], callback_data=f"outline:approve:{review_id}")]]
    if allow_revision:
        rows.append([InlineKeyboardButton(text=labels[1], callback_data=f"outline:revise:{review_id}")])
    return InlineKeyboardMarkup(inline_keyboard=rows)


DESIGN_LABELS = {
    "minimal": {"uk": "Мінімалізм", "ru": "Минимализм", "en": "Minimal"},
    "future": {"uk": "Футуризм", "ru": "Футуризм", "en": "Futuristic"},
    "business": {"uk": "Business Pro", "ru": "Бизнес Pro", "en": "Business Pro"},
    "school": {"uk": "Шкільний креатив", "ru": "Школьный креатив", "en": "School Creative"},
    "luxury": {"uk": "Luxury", "ru": "Luxury", "en": "Luxury"},
    "magazine": {"uk": "Журнал", "ru": "Журнал", "en": "Magazine"},
    "nature": {"uk": "Eco / Nature", "ru": "Eco / Природа", "en": "Eco / Nature"},
    "vibrant": {"uk": "Яскравий", "ru": "Яркий", "en": "Vibrant"},
    "classic": {"uk": "Класичний", "ru": "Классический", "en": "Classic"},
    "mono": {"uk": "Чорно-білий", "ru": "Чёрно-белый", "en": "Monochrome"},
}


def design_name(design: str, language: str = "uk") -> str:
    return DESIGN_LABELS.get(design, DESIGN_LABELS["minimal"]).get(language, DESIGN_LABELS["minimal"]["uk"])


def design_keyboard(language: str = "uk"):
    rows, row = [], []
    for key in DESIGNS:
        row.append(InlineKeyboardButton(text=design_name(key, language), callback_data=f"design:{key}"))
        if len(row) == 2:
            rows.append(row)
            row = []
    if row:
        rows.append(row)
    return InlineKeyboardMarkup(inline_keyboard=rows)


def shop_keyboard(language: str = "uk"):
    labels = {
        "uk": (f"🎟️ 5 генерацій — {CREDIT_PACK_PRICE} ⭐", f"👑 PRO на 30 днів — {PRO_PRICE} ⭐", "👤 Мій профіль"),
        "ru": (f"🎟️ 5 генераций — {CREDIT_PACK_PRICE} ⭐", f"👑 PRO на 30 дней — {PRO_PRICE} ⭐", "👤 Мой профиль"),
        "en": (f"🎟️ 5 generations — {CREDIT_PACK_PRICE} ⭐", f"👑 PRO for 30 days — {PRO_PRICE} ⭐", "👤 My profile"),
    }[language if language in ("uk", "ru", "en") else "uk"]
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text=labels[0], callback_data="buy:credits")],
        [InlineKeyboardButton(text=labels[1], callback_data="buy:pro")],
        [InlineKeyboardButton(text=labels[2], callback_data="profile")],
    ])


def report_admin_keyboard(report_id: str):
    return InlineKeyboardMarkup(inline_keyboard=[[
        InlineKeyboardButton(text="✅ Повернути 1 кредит", callback_data=f"reportact:{report_id}:refund"),
        InlineKeyboardButton(text="❌ Відхилити", callback_data=f"reportact:{report_id}:reject"),
    ]])


def repeat_last_keyboard(language: str = "uk"):
    language = language if language in ("uk", "ru", "en") else "uk"
    return InlineKeyboardMarkup(inline_keyboard=[[
        InlineKeyboardButton(text=UI_TEXT["repeat_last"][language], callback_data="repeat:last")
    ]])


def report_user_keyboard(generation_id: str, language: str = "uk"):
    label = {"uk": "⚠️ Повідомити про брак", "ru": "⚠️ Сообщить о проблеме", "en": "⚠️ Report a problem"}[language if language in ("uk", "ru", "en") else "uk"]
    return InlineKeyboardMarkup(inline_keyboard=[[
        InlineKeyboardButton(text=label, callback_data=f"bad:{generation_id}")
    ]])

# Compatibility defaults for legacy administrative/report flows.
MAIN_KEYBOARD = main_keyboard("uk")
WELCOME_TEXT = UI_TEXT["welcome"]["uk"]
FEATURES_TEXT = UI_TEXT["features"]["uk"]
HELP_TEXT = UI_TEXT["help"]["uk"]
BTN_START = BUTTONS["start"]["uk"]
BTN_RESTART = BUTTONS["restart"]["uk"]
BTN_CREATE = BUTTONS["create"]["uk"]
BTN_HOME = BUTTONS["home"]["uk"]
BTN_CANCEL = BUTTONS["cancel"]["uk"]
BTN_FEATURES = BUTTONS["features"]["uk"]
BTN_HELP = BUTTONS["help"]["uk"]
BTN_PROFILE = BUTTONS["profile"]["uk"]
BTN_SHOP = BUTTONS["shop"]["uk"]
BTN_SUB = BUTTONS["sub"]["uk"]

# Layout key: color palette + compositional style.# Layout key: color palette + compositional style. These are real design variations,
# not merely recolored copies of the same slide.
DESIGNS = {
    "minimal": {"name": "Мінімалізм", "bg": (248, 249, 252), "accent": (28, 82, 125), "card": (255, 255, 255), "text": (30, 42, 58), "muted": (105, 119, 137), "style": "minimal"},
    "future": {"name": "Футуризм", "bg": (10, 14, 34), "accent": (0, 225, 255), "card": (24, 31, 63), "text": (244, 249, 255), "muted": (173, 192, 220), "style": "future"},
    "business": {"name": "Business Pro", "bg": (19, 33, 51), "accent": (68, 160, 255), "card": (32, 51, 73), "text": (249, 251, 255), "muted": (190, 205, 221), "style": "business"},
    "school": {"name": "Шкільний креатив", "bg": (255, 249, 235), "accent": (242, 125, 43), "card": (255, 255, 255), "text": (54, 51, 48), "muted": (117, 109, 100), "style": "school"},
    "luxury": {"name": "Luxury", "bg": (23, 22, 25), "accent": (211, 177, 105), "card": (44, 40, 36), "text": (250, 244, 231), "muted": (201, 190, 169), "style": "luxury"},
    "magazine": {"name": "Журнал", "bg": (246, 241, 232), "accent": (183, 52, 47), "card": (255, 253, 248), "text": (42, 38, 36), "muted": (112, 100, 94), "style": "magazine"},
    "nature": {"name": "Eco / Nature", "bg": (237, 245, 236), "accent": (54, 126, 85), "card": (250, 255, 249), "text": (33, 56, 43), "muted": (93, 119, 101), "style": "nature"},
    "vibrant": {"name": "Яскравий", "bg": (28, 23, 54), "accent": (255, 91, 145), "card": (51, 43, 86), "text": (255, 255, 255), "muted": (207, 197, 233), "style": "vibrant"},
    "classic": {"name": "Класичний", "bg": (240, 244, 248), "accent": (34, 74, 116), "card": (255, 255, 255), "text": (34, 48, 65), "muted": (104, 119, 136), "style": "classic"},
    "mono": {"name": "Чорно-білий", "bg": (245, 245, 245), "accent": (25, 25, 25), "card": (255, 255, 255), "text": (28, 28, 28), "muted": (105, 105, 105), "style": "mono"},
}
LANGUAGES = {"uk": "українською", "en": "англійською", "ru": "російською"}
LAYOUTS = {"cover", "split", "cards", "timeline", "steps", "quote", "comparison", "focus", "closing", "stat", "section", "two_col", "grid3", "photo_text", "portrait_fact"}
 
# Різноманітність: щоразу випадковий ракурс, тон, початок і оформлення.
CREATIVE_ANGLES = [
    "через історію та реальні життєві ситуації", "через проблеми і способи їх вирішення",
    "через питання й відповіді", "через порівняння минулого, сьогодення й майбутнього",
    "через міфи та правду", "через практичні приклади й поради", "через цікаві аналогії та метафори",
    "через погляд скептика та погляд ентузіаста", "через мандрівку «від простого до складного»",
    "через ключові ролі, людей і їхній вплив",
]
CREATIVE_TONES = [
    "живий і захопливий", "спокійний експертний", "дружній і трохи жартівливий",
    "динамічний і енергійний", "вдумливий і філософський", "чіткий діловий",
]
CREATIVE_HOOKS = [
    "обкладинка з несподіваного запитання", "обкладинка з сильного твердження",
    "обкладинка з короткої інтригуючої історії", "обкладинка з яскравого образу чи метафори",
]

NARRATIVE_ARCHITECTURES = [
    "розслідування: загадка → докази → пояснення → наслідки",
    "історія змін: було → переломний момент → що змінилося → що далі",
    "практичний путівник: проблема → інструменти → послідовність дій → результат",
    "дебати: поширене твердження → аргументи → контраргументи → зважений висновок",
    "кейс-стаді: контекст → рішення → реалізація → уроки",
    "від малого до великого: приклад → принцип → система → ширший вплив",
    "майбутній сценарій: сигнал сьогодні → можливий розвиток → ризики → підготовка",
    "питання-відповіді: ключові запитання аудиторії у драматургічній послідовності",
]

DECK_VISUAL_CONCEPTS = [
    "редакційне досьє з великими тематичними заголовками й точковими акцентами",
    "аналітична карта з візуальними зв'язками між причинами та наслідками",
    "польовий журнал із нотатками, спостереженнями та короткими висновками",
    "стратегічний бриф із ключовими рішеннями, доказами та діями",
    "історія трансформації з виразними поворотними моментами",
    "наукова лабораторна подача з гіпотезами, поясненнями та висновками",
    "візуальний путівник із практичними кроками й прикладами",
    "журнальний спецвипуск із контрастом великих тез і коротких деталей",
    "порівняльне дослідження з чіткими критеріями й зваженим підсумком",
    "концептуальна розповідь навколо однієї доречної метафори теми",
]

STYLE_PROMPTS = {
    "minimal": "Swiss editorial minimalism: strict grid, generous whitespace, precise alignment, thin rules, restrained navy accent, oversized typography, no decorative clutter, sparse but intentional visual hierarchy.",
    "future": "Premium futuristic technology aesthetic: deep midnight background, electric cyan highlights, luminous rings, technical grid, layered interface panels, data-like visual rhythm, high contrast.",
    "business": "Executive consulting deck: confident navy and blue, crisp column grids, clear hierarchy, structured comparison panels, restrained icon-like geometry, professional boardroom quality.",
    "school": "Modern educational creative design: warm paper-like background, vivid orange accents, friendly geometric shapes, clear visual storytelling, energetic but readable layouts, not childish clip-art.",
    "luxury": "Luxury editorial art direction: charcoal and warm black surfaces, muted gold accents, elegant serif-like typographic hierarchy, generous space, fine borders, understated premium details.",
    "magazine": "Contemporary magazine layout: asymmetric editorial grid, bold red accents, strong headline hierarchy, large chapter numerals, image-window placeholders made with vector shapes, dynamic page composition.",
    "nature": "Eco and nature visual identity: sage and forest green palette, organic rounded forms, botanical-inspired geometry, calm spacious layouts, natural muted contrast.",
    "vibrant": "High-energy creative studio aesthetic: rich violet base, vivid pink and coral accents, bold oversized typography, dynamic blocks, varied asymmetric compositions, energetic but controlled.",
    "classic": "Timeless academic and corporate design: ivory and slate blue, formal grid, refined typography, subtle double rules, balanced margins, clear and dignified visual hierarchy.",
    "mono": "High-contrast monochrome editorial design: black, white and grayscale only, strong typographic scale, brutalist blocks, hard edges, no rounded colorful cards, dramatic negative space.",
}
DECOR_VARIANTS = ("side", "top", "bottom", "circles", "corner", "blocks")
COVER_SHAPES = (
    (MSO_SHAPE.OVAL, MSO_SHAPE.OVAL),
    (MSO_SHAPE.ROUNDED_RECTANGLE, MSO_SHAPE.ROUNDED_RECTANGLE),
    (MSO_SHAPE.DIAMOND, MSO_SHAPE.OVAL),
    (MSO_SHAPE.HEXAGON, MSO_SHAPE.OVAL),
    (MSO_SHAPE.ROUNDED_RECTANGLE, MSO_SHAPE.OVAL),
)

ACCENT_ALTS = {
    "minimal": [(28, 82, 125), (0, 121, 107), (112, 74, 160), (173, 91, 43)],
    "future": [(0, 225, 255), (139, 92, 246), (0, 202, 156), (255, 90, 170)],
    "business": [(68, 160, 255), (0, 137, 123), (194, 85, 70), (116, 95, 185)],
    "school": [(242, 125, 43), (0, 150, 136), (233, 30, 99), (74, 112, 190)],
    "luxury": [(211, 177, 105), (174, 140, 89), (175, 160, 194), (142, 112, 78)],
    "magazine": [(183, 52, 47), (31, 104, 142), (111, 75, 134), (197, 112, 42)],
    "nature": [(54, 126, 85), (93, 133, 91), (47, 117, 127), (129, 139, 77)],
    "vibrant": [(255, 91, 145), (0, 220, 210), (255, 164, 48), (152, 112, 255)],
    "classic": [(34, 74, 116), (94, 91, 135), (40, 112, 103), (138, 92, 61)],
    "mono": [(25, 25, 25), (65, 65, 65), (0, 0, 0), (95, 95, 95)],
}


def make_deck_recipe(design: str, seed: int, slide_count: int, visual_concept: str = "") -> dict:
    """Deterministic art direction per presentation, driven by Gemini's visual seed."""
    rng = random.Random(int(seed))
    accents = ACCENT_ALTS.get(design, [DESIGNS[design]["accent"]])
    headers = ["rule", "tag", "numbered", "sidebar"]
    rng.shuffle(headers)
    decorations = list(DECOR_VARIANTS)
    rng.shuffle(decorations)
    return {
        "accent": rng.choice(accents),
        "decorations": [decorations[i % len(decorations)] for i in range(slide_count)],
        "headers": [headers[i % len(headers)] for i in range(slide_count)],
        "cover_shapes": rng.choice(COVER_SHAPES),
        "quiet_every": rng.choice((2, 3)),
        "panel_shift": rng.randint(0, 2),
        "visual_concept": str(visual_concept or "")[:160],
    }
 
class PresentationForm(StatesGroup):
    topic = State()
    slides = State()
    language = State()
    detail = State()
    audience = State()
    photos = State()
    outline_feedback = State()
    design = State()


OUTLINE_REVIEW_WAITERS = {}
AUDIENCE_GUIDANCE = {
    "students": "Explain terms simply, build concepts from basic to advanced, and use clear student-friendly examples.",
    "business": "Use a concise professional tone and emphasize decisions, practical impact, risks, and business examples.",
    "general": "Use clear language for a broad audience and briefly explain specialized terms.",
    "speaker": "Make the story easy to present aloud, with clear transitions, audience context, and memorable examples.",
}


DETAIL_PROFILES = {
    "short": {
        "min_points": 2, "max_points": 3, "max_point_chars": 90,
        "min_note_sentences": 1, "max_note_sentences": 2, "max_note_chars": 420,
        "tokens_per_slide": 450, "request_timeout": 43.0,
        "prompt": {
            "uk": "Коротко: 2–3 тези на слайд, кожна до 90 символів; нотатки доповідача — 1–2 короткі речення.",
            "ru": "Кратко: 2–3 тезиса на слайд, каждый до 90 символов; заметки докладчика — 1–2 коротких предложения.",
            "en": "Short: 2–3 bullets per slide, each up to 90 characters; speaker notes: 1–2 concise sentences.",
        },
    },
    "medium": {
        "min_points": 4, "max_points": 4, "max_point_chars": 125,
        "min_note_sentences": 2, "max_note_sentences": 3, "max_note_chars": 750,
        "tokens_per_slide": 800, "request_timeout": 43.0,
        "prompt": {
            "uk": "Середньо: 3–4 змістовні тези на слайд, кожна до 150 символів; нотатки доповідача — 2–3 речення з поясненням.",
            "ru": "Средне: 3–4 содержательных тезиса на слайд, каждый до 150 символов; заметки докладчика — 2–3 поясняющих предложения.",
            "en": "Medium: 3–4 meaningful bullets per slide, each up to 150 characters; speaker notes: 2–3 explanatory sentences.",
        },
    },
    "detailed": {
        "min_points": 5, "max_points": 5, "max_point_chars": 155,
        "min_note_sentences": 3, "max_note_sentences": 5, "max_note_chars": 1200,
        "tokens_per_slide": 1500, "request_timeout": 75.0,
        "prompt": {
            "uk": "Детально: 4–5 розгорнутих, але читабельних тез на слайд, кожна до 200 символів; нотатки доповідача — 3–5 повних речень із контекстом та прикладами.",
            "ru": "Подробно: 4–5 развёрнутых, но читабельных тезисов на слайд, каждый до 200 символов; заметки докладчика — 3–5 полных предложений с контекстом и примерами.",
            "en": "Detailed: 4–5 substantial but readable bullets per slide, each up to 200 characters; speaker notes: 3–5 complete sentences with context and examples.",
        },
    },
}


def detail_output_tokens(detail: str, slide_count: int) -> int:
    """Scale Gemini's JSON output budget to slide count and selected detail."""
    profile = DETAIL_PROFILES.get(detail, DETAIL_PROFILES["medium"])
    try:
        count = max(1, min(12, int(slide_count)))
    except (TypeError, ValueError):
        count = 8
    return min(GEMINI_OUTPUT_TOKEN_LIMIT, 1200 + count * int(profile["tokens_per_slide"]))


def truncate_at_word(value: str, limit: int) -> str:
    """Keep a point within its character limit without cutting the last word."""
    text = re.sub(r"\s+", " ", str(value or "")).strip()
    if len(text) <= limit:
        return text
    if limit <= 2:
        return text[:limit]
    candidate = text[:limit - 1].rstrip()
    boundary = candidate.rfind(" ")
    if boundary >= max(8, int(limit * 0.55)):
        candidate = candidate[:boundary].rstrip()
    return candidate.rstrip(" ,;:-") + "…"


KNOWN_PHRASE_TRANSLATIONS = {
    "ru": {
        "Джерело": "Источник", "Походження не підтверджене": "Происхождение не подтверждено",
        "КЛЮЧОВА ДУМКА": "КЛЮЧЕВАЯ МЫСЛЬ", "ПОЯСНЕННЯ": "ПОЯСНЕНИЕ",
        "КЛЮЧОВУ ІНФОРМАЦІЮ НЕ ВДАЛОСЯ ПІДТВЕРДИТИ": "КЛЮЧЕВУЮ ИНФОРМАЦИЮ НЕ УДАЛОСЬ ПОДТВЕРДИТЬ",
        "Практичний сценарій": "Практический сценарий", "Ризики та обмеження": "Риски и ограничения",
        "Висновок і джерела": "Выводы и источники",
    },
    "en": {
        "Джерело": "Source", "Походження не підтверджене": "Origin not confirmed",
        "КЛЮЧОВА ДУМКА": "KEY IDEA", "ПОЯСНЕННЯ": "EXPLANATION",
        "КЛЮЧОВУ ІНФОРМАЦІЮ НЕ ВДАЛОСЯ ПІДТВЕРДИТИ": "THE KEY INFORMATION COULD NOT BE VERIFIED",
        "Практичний сценарій": "Practical scenario", "Ризики та обмеження": "Risks and limitations",
        "Висновок і джерела": "Conclusion and sources",
    },
}

def localize_known_phrases(value: str, language: str) -> str:
    """Translate known static Ukrainian fallback labels if a model leaks them into ru/en decks."""
    result = str(value or "")
    translations = KNOWN_PHRASE_TRANSLATIONS.get(language, {})
    for source, target in translations.items():
        result = re.sub(re.escape(source), target, result, flags=re.IGNORECASE)
    return result
 
class ReportForm(StatesGroup):
    reason = State()
 
 


# 4. TEXTS AND PROFILE
# ============================================================
def format_until(value, language="uk"):
    inactive = {"uk": "не активна", "ru": "не активна", "en": "inactive"}[language if language in ("uk", "ru", "en") else "uk"]
    if not value:
        return inactive
    try:
        dt = datetime.fromisoformat(value).astimezone(KYIV_TZ)
        return dt.strftime("%d.%m.%Y %H:%M")
    except ValueError:
        return inactive


def profile_text(user_id: int, language: str | None = None):
    language = language or get_user_language(user_id) or "uk"
    row = get_user(user_id)
    if not row:
        return {"uk": "Профіль ще не створено. Натисни /start.", "ru": "Профиль ещё не создан. Нажми /start.", "en": "Profile not created yet. Press /start."}[language]
    used = get_usage(user_id)
    pro = subscription_active(row)
    limit = "∞" if is_admin(user_id) else (PRO_DAILY_LIMIT if pro else FREE_DAILY_LIMIT)
    remaining = "∞" if is_admin(user_id) else max(0, int(limit) - used)
    with db_connect() as con:
        total = con.execute(
            "SELECT COUNT(*) AS n FROM generation_history WHERE user_id=? AND status='completed'",
            (user_id,),
        ).fetchone()["n"]
    until = format_until(row["subscription_until"], language)
    if is_admin(user_id):
        daily = {"uk": "∞ (адміністратор)", "ru": "∞ (администратор)", "en": "∞ (administrator)"}[language]
    else:
        daily = {
            "uk": f"{used}/{limit} використано • залишилось {remaining}",
            "ru": f"{used}/{limit} использовано • осталось {remaining}",
            "en": f"{used}/{limit} used • {remaining} remaining",
        }[language]
    sub = {
        "uk": f"активна до {until}" if pro else "не активна",
        "ru": f"активна до {until}" if pro else "не активна",
        "en": f"active until {until}" if pro else "inactive",
    }[language]
    templates = {
        "uk": "✦ <b>ВАШ ПРОФІЛЬ PREZENT AI</b>\n\n🆔 ID: <code>{uid}</code>\n📊 Створено презентацій: <b>{total}</b>\n🎁 Денний ліміт: <b>{daily}</b>\n🎟️ Додаткові кредити: <b>{credits}</b>\n👑 PRO: <b>{sub}</b>\n\nКредити використовуються після вичерпання денного ліміту.",
        "ru": "✦ <b>ВАШ ПРОФИЛЬ PREZENT AI</b>\n\n🆔 ID: <code>{uid}</code>\n📊 Создано презентаций: <b>{total}</b>\n🎁 Дневной лимит: <b>{daily}</b>\n🎟️ Дополнительные кредиты: <b>{credits}</b>\n👑 PRO: <b>{sub}</b>\n\nКредиты используются после исчерпания дневного лимита.",
        "en": "✦ <b>YOUR PREZENT AI PROFILE</b>\n\n🆔 ID: <code>{uid}</code>\n📊 Presentations created: <b>{total}</b>\n🎁 Daily quota: <b>{daily}</b>\n🎟️ Extra credits: <b>{credits}</b>\n👑 PRO: <b>{sub}</b>\n\nCredits are used after your daily quota is exhausted.",
    }
    return templates[language].format(uid=user_id, total=total, daily=daily, credits=row["credits"], sub=sub)


# ============================================================
# ============================================================
# 5. ANIMATION — мінімалістичні рухомі зірочки без відсотків
# ============================================================
class ProgressAnimator:
    """A localized, stage-driven progress message. No fake percentage is shown."""
    TICK_SECONDS = 1.9
    CHAT_ACTION_SECONDS = 5.0
    STEPS = ("writing", "pptx", "pdf", "sending")
    WEIGHTS = {"writing": 0.60, "pptx": 0.15, "pdf": 0.15, "sending": 0.10}
    SPINNER = ("🌑", "🌒", "🌓", "🌔", "🌕", "🌖", "🌗", "🌘")
    STAGE_ICONS = {"writing": "🧠", "pptx": "🎨", "pdf": "📄", "sending": "📤"}
    SUBSTAGES = ("research", "writing", "photos")
    SUBSTAGE_ICONS = {"research": "🔎", "writing": "✍️", "photos": "🖼️"}
    STEP_LABELS = {
        "uk": {"writing": "Пишу тексти й структуру", "pptx": "Збираю слайди PowerPoint", "pdf": "Готую PDF-версію", "sending": "Надсилаю готові файли"},
        "ru": {"writing": "Пишу тексты и структуру", "pptx": "Собираю слайды PowerPoint", "pdf": "Готовлю PDF-версию", "sending": "Отправляю готовые файлы"},
        "en": {"writing": "Writing content and structure", "pptx": "Building PowerPoint slides", "pdf": "Preparing the PDF version", "sending": "Sending the finished files"},
    }
    SUBSTAGE_LABELS = {
        "uk": {"research": "Досліджую факти й джерела", "writing": "Пишу тексти та вибудовую логіку", "photos": "Підбираю й додаю фотографії"},
        "ru": {"research": "Исследую факты и источники", "writing": "Пишу тексты и выстраиваю логику", "photos": "Подбираю и добавляю фотографии"},
        "en": {"research": "Researching facts and sources", "writing": "Writing content and shaping the story", "photos": "Finding and adding photographs"},
    }
    HINTS = {
        "uk": {
            "research": ("Звіряю твердження з надійними джерелами", "Перевіряю дати, імена та цифри", "Відокремлюю факти від припущень", "Добираю джерела для ключових тез"),
            "writing": ("Підбираю найсильніші тези", "Вибудовую логіку розповіді", "Уточнюю формулювання для аудиторії", "Готую нотатки доповідача"),
            "photos": ("Шукаю зображення, що відповідають темі", "Перевіряю автора й ліцензію фото", "Оптимізую фотографії для слайдів", "Розміщую фото поруч із потрібним змістом"),
            "pptx": ("Застосовую вибраний візуальний стиль", "Налаштовую композицію кожного слайда", "Перевіряю читабельність тексту", "Додаю зображення та візуальні акценти"),
            "pdf": ("Підбираю типографіку для PDF", "Перевіряю перенесення тексту", "Зберігаю джерела та підписи фото", "Формую PDF-документ"),
            "sending": ("Перевіряю готові файли", "Готую документи для Telegram", "Додаю кнопки відгуку", "Надсилаю результат у чат"),
        },
        "ru": {
            "research": ("Сверяю утверждения с надёжными источниками", "Проверяю даты, имена и числа", "Отделяю факты от предположений", "Подбираю источники для ключевых тезисов"),
            "writing": ("Выбираю самые сильные тезисы", "Выстраиваю логику рассказа", "Уточняю формулировки для аудитории", "Готовлю заметки докладчика"),
            "photos": ("Ищу изображения по теме", "Проверяю автора и лицензию фото", "Оптимизирую фотографии для слайдов", "Размещаю фото рядом с нужным содержанием"),
            "pptx": ("Применяю выбранный визуальный стиль", "Настраиваю композицию каждого слайда", "Проверяю читаемость текста", "Добавляю изображения и акценты"),
            "pdf": ("Подбираю типографику для PDF", "Проверяю переносы текста", "Сохраняю источники и подписи фото", "Формирую PDF-документ"),
            "sending": ("Проверяю готовые файлы", "Готовлю документы для Telegram", "Добавляю кнопки отзыва", "Отправляю результат в чат"),
        },
        "en": {
            "research": ("Cross-checking claims against reliable sources", "Checking dates, names, and figures", "Separating facts from assumptions", "Matching key points with references"),
            "writing": ("Selecting the strongest key points", "Shaping a clear narrative", "Polishing wording for the audience", "Preparing speaker notes"),
            "photos": ("Finding images relevant to the topic", "Checking image authors and licenses", "Optimizing photos for slides", "Placing images beside relevant content"),
            "pptx": ("Applying your selected visual style", "Refining each slide composition", "Checking text readability", "Adding images and visual accents"),
            "pdf": ("Preparing PDF typography", "Checking text wrapping", "Preserving references and photo credits", "Building the PDF document"),
            "sending": ("Checking the finished files", "Preparing Telegram attachments", "Adding feedback controls", "Sending the result to chat"),
        },
    }
    DETAIL_LABELS = {
        "uk": {"short": "Коротко", "medium": "Середньо", "detailed": "Детально"},
        "ru": {"short": "Кратко", "medium": "Средне", "detailed": "Подробно"},
        "en": {"short": "Short", "medium": "Medium", "detailed": "Detailed"},
    }
    ESTIMATES = {
        "uk": {"normal": "Зазвичай це займає 30–60 секунд", "detailed": "Зазвичай це займає 60–90 секунд"},
        "ru": {"normal": "Обычно это занимает 30–60 секунд", "detailed": "Обычно это занимает 60–90 секунд"},
        "en": {"normal": "This usually takes 30–60 seconds", "detailed": "This usually takes 60–90 seconds"},
    }
    EXPECTED_STAGE_SECONDS = {"pptx": 18.0, "pdf": 16.0, "sending": 10.0}
    # The bar jumps to a fixed checkpoint when the generator really enters a phase,
    # and creeps towards the next checkpoint while that phase runs.
    PHASE_STARTS_PHOTOS = {"research": 0.00, "writing": 0.15, "photos": 0.50, "pptx": 0.60, "pdf": 0.80, "sending": 0.90}
    PHASE_STARTS_NO_PHOTOS = {"research": 0.00, "writing": 0.18, "pptx": 0.60, "pdf": 0.80, "sending": 0.90}
    # Expected durations are visual pacing hints, not promised completion times.
    PHASE_EXPECTED = {"research": 20.0, "writing": 40.0, "photos": 8.0, "pptx": 18.0, "pdf": 16.0, "sending": 10.0}

    def __init__(self, language: str = "uk", message: Message | None = None,
                 style: str = "minimal", detail: str = "medium", slide_count: int = 8, photos: bool = True):
        self.substages = self.SUBSTAGES if photos else tuple(k for k in self.SUBSTAGES if k != "photos")
        self.language = language if language in ("uk", "ru", "en") else "uk"
        self.message = message
        self.style = style if style in DESIGNS else "minimal"
        self.detail = detail if detail in DETAIL_PROFILES else "medium"
        self.slide_count = int(slide_count or 8)
        self.frame = 0
        self.stage = "writing"
        self.substage = "research"
        self.stage_started_at = time.monotonic()
        self.substage_started_at = time.monotonic()
        self._shown = 0.0
        self._phase_seen = None
        self._phase_started_at = time.monotonic()
        self._edit_failures = 0
        self.task = None
        self.lock = asyncio.Lock()
        self.running = False
        self.completed = False
        self._wake = asyncio.Event()
        self._last_edit_at = 0.0
        self._last_chat_action_at = 0.0

    def set_stage(self, key: str):
        """Advance the main stage only when the generator enters it."""
        if key in self.STEPS and key != self.stage:
            self.stage = key
            self.stage_started_at = time.monotonic()
            self._wake.set()

    def set_substage(self, key: str):
        """Update one of the real research/writing/photo substages."""
        if key in self.substages and key != self.substage:
            self.substage = key
            self.substage_started_at = time.monotonic()
            self._wake.set()

    def _current_stage_expected(self) -> float:
        if self.stage == "writing":
            base = 88.0 if self.detail == "detailed" else (48.0 if self.detail == "short" else 62.0)
            return base * max(0.75, min(1.5, self.slide_count / 8.0))
        return self.EXPECTED_STAGE_SECONDS.get(self.stage, 12.0) * max(0.75, min(1.5, self.slide_count / 8.0))

    def _detail_scale(self) -> float:
        k = {"short": 0.8, "medium": 1.0, "detailed": 1.4}.get(self.detail, 1.0)
        return k * max(0.75, min(1.5, self.slide_count / 8.0))

    def _phase_key(self) -> str:
        return self.substage if self.stage == "writing" else self.stage

    def _progress_fraction(self) -> float:
        if self.completed:
            return 1.0
        starts = self.PHASE_STARTS_PHOTOS if "photos" in self.substages else self.PHASE_STARTS_NO_PHOTOS
        key = self._phase_key()
        if key not in starts:
            key = "writing"
        if key != self._phase_seen:
            self._phase_seen = key
            self._phase_started_at = time.monotonic()
        order = list(starts)
        start = starts[key]
        position = order.index(key)
        end = starts[order[position + 1]] if position + 1 < len(order) else 0.99
        scale = self._detail_scale() if key in ("research", "writing") else 1.0
        expected = max(1.0, self.PHASE_EXPECTED[key] * scale)
        elapsed = max(0.0, time.monotonic() - self._phase_started_at)
        raw = start + (end - start) * 0.9 * (1.0 - math.exp(-elapsed / expected))
        self._shown = max(self._shown, min(0.985, raw))  # the bar never moves backwards
        return self._shown

    def _progress_percent(self) -> int:
        if self.completed:
            return 100
        return min(99, int(self._progress_fraction() * 100))

    def _progress_bar(self) -> str:
        """Ten bold, wide blocks stay legible in Telegram's compact chat view."""
        total_segments = 10
        if self.completed:
            return "█" * total_segments
        filled = min(total_segments - 1, max(0, int(self._progress_fraction() * total_segments)))
        return "█" * filled + "░" * (total_segments - filled)

    def _checklist(self) -> str:
        current = self.STEPS.index(self.stage)
        spinner = self.SPINNER[self.frame % len(self.SPINNER)]
        rows = []
        for index, key in enumerate(self.STEPS):
            if self.completed or index < current:
                marker = "✅"
            elif index == current:
                marker = spinner
            else:
                marker = "▫️"
            label = html.escape(self.STEP_LABELS[self.language][key])
            icon = self.STAGE_ICONS[key]
            if not self.completed and index == current:
                rows.append(f"<b>{marker} {icon} {label}</b>")
            else:
                rows.append(f"{marker} {icon} {label}")
        return "\n".join(rows)

    def _writing_substeps(self) -> str:
        if self.completed:
            return "\n".join(f"✅ {self.SUBSTAGE_ICONS[key]} {html.escape(self.SUBSTAGE_LABELS[self.language][key])}" for key in self.substages)
        current = self.substages.index(self.substage)
        spinner = self.SPINNER[self.frame % len(self.SPINNER)]
        rows = []
        for index, key in enumerate(self.substages):
            marker = "✅" if index < current else (spinner if index == current else "▫️")
            text = f"{marker} {self.SUBSTAGE_ICONS[key]} {html.escape(self.SUBSTAGE_LABELS[self.language][key])}"
            rows.append(f"<b>{text}</b>" if index == current else text)
        return "\n".join(rows)

    def render(self) -> str:
        hint_key = self.substage if self.stage == "writing" else self.stage
        hints = self.HINTS[self.language][hint_key]
        hint = hints[self.frame % len(hints)]
        dots = "·" * ((self.frame % 3) + 1)
        style_label = html.escape(design_name(self.style, self.language))
        detail_label = html.escape(self.DETAIL_LABELS[self.language][self.detail])
        estimate_key = "detailed" if self.detail == "detailed" else "normal"
        estimate = html.escape(self.ESTIMATES[self.language][estimate_key])
        title = {"uk": "Ваша презентація створюється", "ru": "Ваша презентация создаётся", "en": "Your presentation is being created"}[self.language]
        slide_word = {"uk": "слайдів", "ru": "слайдов", "en": "slides"}[self.language]
        writing_block = ""
        if self.stage == "writing":
            writing_block = "\n" + self._writing_substeps() + "\n"
        return (
            f"✨ <b>PREZENT AI STUDIO</b> ✨\n\n"
            f"<b>{html.escape(title)}</b>\n\n"
            f"{self._checklist()}\n"
            f"{writing_block}\n"
            f"<i>{html.escape(hint + dots)}</i>\n\n"
            f"🎨 {style_label}  •  📚 {detail_label}  •  📑 {self.slide_count} {slide_word}\n"
            f"<i>{estimate}</i>\n\n"
            f"<code>{self._progress_bar()}  {self._progress_percent()}%</code>"
        )

    async def _edit_message(self, text: str, enforce_interval: bool = True) -> bool:
        if not self.message:
            return False
        if enforce_interval:
            remaining = self.TICK_SECONDS - (time.monotonic() - self._last_edit_at)
            if remaining > 0:
                await asyncio.sleep(remaining)

        last_error = None
        for attempt in range(3):
            delay = 0.0
            async with self.lock:
                self._last_edit_at = time.monotonic()
                try:
                    await asyncio.wait_for(
                        self.message.edit_text(text, parse_mode=ParseMode.HTML), timeout=12.0
                    )
                    self._edit_failures = 0
                    return True
                except TelegramRetryAfter as exc:
                    delay = min(30.0, max(1.0, float(getattr(exc, "retry_after", 1.0))))
                    last_error = exc
                    log.warning("Telegram rate-limited generation animation; retry in %.1fs", delay)
                except Exception as exc:
                    message = str(exc).lower()
                    if "message is not modified" in message:
                        self._edit_failures = 0
                        return True
                    last_error = exc
                    type_name = type(exc).__name__.lower()
                    transient = (
                        any(token in type_name for token in ("network", "server", "timeout", "connection"))
                        or any(token in message for token in (
                            "timed out", "timeout", "connection reset", "connection aborted",
                            "temporarily unavailable", "bad gateway", "gateway timeout",
                            "internal server error", "service unavailable", "network is unreachable",
                        ))
                    )
                    if not transient or attempt >= 2:
                        self._edit_failures += 1
                        if self._edit_failures <= 3:
                            log.warning(
                                "Generation animation edit failed (%s): %s", type(exc).__name__, exc
                            )
                        # BadRequest errors such as "message can't be edited" will not
                        # recover on a retry. Stop the animation instead of repeating
                        # the same failed Telegram request every tick.
                        if not transient and any(token in message for token in (
                            "message can't be edited", "message to edit not found",
                            "message identifier is not specified",
                        )):
                            self.running = False
                        break
                    delay = 1.0 * (2 ** attempt)
            # Never hold the lock while waiting for FloodWait or a transient network retry.
            if delay:
                await asyncio.sleep(delay)
                if enforce_interval:
                    remaining = self.TICK_SECONDS - (time.monotonic() - self._last_edit_at)
                    if remaining > 0:
                        await asyncio.sleep(remaining)

        if self._edit_failures >= 3:
            try:
                await notify_admins_once(
                    "telegram-animation-edit-failed",
                    "⚠️ <b>Telegram не оновлює повідомлення прогресу Prezent AI.</b> Перевір Telegram API та Runtime Logs. "
                    + html.escape(str(last_error)[:500]),
                )
            except Exception:
                log.exception("Could not send animation failure alert")
        return False

    async def safe_edit(self):
        return await self._edit_message(self.render(), enforce_interval=True)

    async def _send_chat_action_if_due(self):
        if not self.message:
            return
        now = time.monotonic()
        if now - self._last_chat_action_at < self.CHAT_ACTION_SECONDS:
            return
        self._last_chat_action_at = now
        try:
            await self.message.bot.send_chat_action(chat_id=self.message.chat.id, action="upload_document")
        except Exception as exc:
            log.debug("Could not send upload_document chat action: %s", exc)

    async def run(self):
        self.running = True
        try:
            while self.running:
                try:
                    await asyncio.wait_for(self._wake.wait(), timeout=self.TICK_SECONDS)
                except asyncio.TimeoutError:
                    pass
                self._wake.clear()
                if not self.running:
                    break
                self.frame += 1
                try:
                    await self.safe_edit()
                    await self._send_chat_action_if_due()
                except asyncio.CancelledError:
                    raise
                except Exception:
                    log.exception("Generation animation failed")
        except asyncio.CancelledError:
            raise
        finally:
            self.running = False

    def start(self):
        if self.task and not self.task.done():
            return
        # Reset the writing-stage clock when a queued request actually gets a slot.
        if self.stage == "writing" and self.frame == 0:
            now = time.monotonic()
            self.stage_started_at = now
            self.substage_started_at = now
            self._phase_started_at = now
            self._phase_seen = None
            self._shown = 0.0
        self.running = True
        self.task = asyncio.create_task(self.run(), name="prezent-ai-generation-animation")

    async def stop(self):
        self.running = False
        self._wake.set()
        task, self.task = self.task, None
        if task and not task.done():
            task.cancel()
            try:
                await task
            except asyncio.CancelledError:
                pass

    async def finish(self, final_text: str):
        """Stop motion and set a localized failure/cancellation/success status."""
        await self.stop()
        await self._edit_message(final_text, enforce_interval=True)

    async def complete(self, final_text: str):
        """Show one final fully-checked frame, then replace it with the ready message."""
        await self.stop()
        if not self.message:
            return
        self.completed = True
        await self._edit_message(self.render(), enforce_interval=True)
        # _edit_message already waits for Telegram's edit interval, which keeps the full bar visible briefly.
        await self._edit_message(final_text, enforce_interval=True)


# ============================================================
# 6. GEMINI API
# ============================================================
class GeminiAPIError(Exception):
    pass
class GeminiRetryableError(GeminiAPIError):
    pass
class GeminiModelUnavailable(GeminiAPIError):
    pass
class GeminiResponseError(GeminiAPIError):
    pass


GEMINI_RESPONSE_SCHEMA = {
    "type": "OBJECT",
    "properties": {
        "title": {"type": "STRING"},
        "photo_subject": {"type": "STRING"},
        "photo_subject_kind": {"type": "STRING", "enum": ["person", "named_subject", "generic"]},
        "slides": {
            "type": "ARRAY",
            "items": {
                "type": "OBJECT",
                "properties": {
                    "title": {"type": "STRING"},
                    "subtitle": {"type": "STRING"},
                    "layout": {"type": "STRING", "enum": ["cover", "split", "cards", "timeline", "steps", "quote", "comparison", "focus", "closing", "stat", "section", "two_col", "grid3", "photo_text", "portrait_fact"]},
                    "content": {"type": "ARRAY", "items": {"type": "STRING"}},
                    "highlight": {"type": "STRING"},
                    "left_title": {"type": "STRING"},
                    "right_title": {"type": "STRING"},
                    "kicker": {"type": "STRING"},
                    "stat_value": {"type": "STRING"},
                    "stat_label": {"type": "STRING"},
                    "notes": {"type": "STRING"},
                    "section_type": {"type": "STRING", "enum": ["cover", "definition", "origin", "mechanics", "example", "interesting_facts", "money", "steps", "extra", "conclusion_sources"]},
                    "fact": {"type": "STRING"},
                    "source": {"type": "STRING"},
                    "image_query": {"type": "STRING"},
                    "image_kind": {"type": "STRING", "enum": ["person", "named_subject", "generic", "none"]},
                },
                "required": ["title", "content", "notes"],
            },
        },
    },
    "required": ["title", "slides"],
}
 
 
def extract_json(text: str) -> dict:
    if not isinstance(text, str) or not text.strip():
        raise GeminiResponseError("Gemini повернув порожню відповідь.")
    cleaned = re.sub(r"^```(?:json)?\s*", "", text.strip(), flags=re.IGNORECASE)
    cleaned = re.sub(r"\s*```$", "", cleaned)
    start, end = cleaned.find("{"), cleaned.rfind("}")
    if start < 0 or end <= start:
        raise GeminiResponseError("Gemini повернув відповідь без JSON.")
    try:
        return json.loads(cleaned[start:end + 1])
    except json.JSONDecodeError as exc:
        raise GeminiResponseError("Gemini повернув неправильний JSON.") from exc
 
 
async def request_gemini(model: str, prompt: str, max_tokens: int = 8192, timeout: float = 43.0) -> dict:
    url = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent"
    headers = {"x-goog-api-key": GEMINI_API_KEY, "Content-Type": "application/json"}
    payload = {
        "systemInstruction": {"parts": [{"text": (
            "Ти креативний професійний редактор презентацій, дизайнер і фактчекер. "
            "Пиши змістовно, без повторів та вигаданих статистичних даних. "
            "Поверни тільки правильний JSON."
        )}]},
        "contents": [{"role": "user", "parts": [{"text": prompt}]}],
        "generationConfig": {
            "maxOutputTokens": int(max_tokens),
            "responseMimeType": "application/json",
            "responseSchema": GEMINI_RESPONSE_SCHEMA,
            "temperature": 0.9,
            "topP": 0.95,
            "seed": random.randint(1, 2_000_000_000),
        },
    }
    for attempt in range(2):
        try:
            async with httpx.AsyncClient(timeout=httpx.Timeout(float(timeout), connect=12.0)) as client:
                response = await client.post(url, headers=headers, json=payload)
        except httpx.TimeoutException as exc:
            raise GeminiRetryableError(f"Модель {model} не відповіла протягом {int(timeout)} секунд.") from exc
        except httpx.HTTPError as exc:
            raise GeminiRetryableError(f"Помилка з'єднання з моделлю {model}.") from exc
        status = response.status_code
        # Check HTTP status before parsing JSON: Google 502/503 pages may be HTML.
        if status in (408, 429, 500, 502, 503, 504):
            if status in (429, 503) and attempt == 0:
                retry_after = response.headers.get("Retry-After", "")
                try:
                    delay = min(3.0, max(1.0, float(retry_after))) if retry_after else 1.5
                except (TypeError, ValueError):
                    delay = 1.5
                log.warning("Gemini model=%s returned HTTP %s; retrying once in %.1fs", model, status, delay)
                await asyncio.sleep(delay)
                continue
            raise GeminiRetryableError(f"Модель {model} тимчасово недоступна: HTTP {status}.")
        if status == 404:
            raise GeminiModelUnavailable(f"Модель {model} не знайдена або недоступна.")
        try:
            result = response.json()
        except ValueError as exc:
            raise GeminiRetryableError(f"Gemini returned non-JSON HTTP {status}; trying fallback model.") from exc
        if status in (401, 403):
            detail = result.get("error", {}).get("message", "Перевір API-ключ.")
            raise GeminiAPIError(f"Google відхилив доступ: {str(detail)[:250]}")
        if status >= 400:
            detail = result.get("error", {}).get("message", "Невідома помилка.")
            raise GeminiAPIError(f"Gemini HTTP {status}: {str(detail)[:250]}")
        if result.get("error"):
            raise GeminiAPIError(str(result["error"].get("message", "Невідома помилка."))[:300])
        candidates = result.get("candidates", [])
        if not candidates:
            reason = result.get("promptFeedback", {}).get("blockReason", "невідома причина")
            raise GeminiResponseError(f"Gemini не повернув відповідь: {reason}.")
        candidate = candidates[0]
        finish_reason = candidate.get("finishReason", "UNKNOWN")
        parts = candidate.get("content", {}).get("parts", [])
        content = "\n".join(
            p.get("text", "").strip() for p in parts
            if isinstance(p, dict) and not p.get("thought") and isinstance(p.get("text"), str)
        ).strip()
        log.info("Gemini model=%s finish=%s chars=%s", model, finish_reason, len(content))
        if finish_reason == "SAFETY":
            raise GeminiAPIError("Gemini заблокував відповідь фільтром безпеки.")
        if finish_reason == "MAX_TOKENS":
            raise GeminiResponseError("Gemini досяг ліміту вихідних токенів.")
        if not content:
            raise GeminiRetryableError("Gemini повернув порожню відповідь.")
        return extract_json(content)
    raise GeminiRetryableError(f"Модель {model} не повернула результат після повторної спроби.")
 
 


async def request_gemini_research(model: str, prompt: str, timeout: float = 10.0) -> tuple[str, list[dict]]:
    """Google Search grounding request with an explicit timeout supplied by the shared research budget."""
    url = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent"
    headers = {"x-goog-api-key": GEMINI_API_KEY, "Content-Type": "application/json"}
    payload = {
        "systemInstruction": {"parts": [{"text": (
            "Ти фактчекер. Шукай лише надійні джерела, відокремлюй підтверджені факти від невідомого. "
            "Не вигадуй числа, дати, авторів, назви джерел або URL. Пиши стисло."
        )}]},
        "contents": [{"role": "user", "parts": [{"text": prompt}]}],
        "tools": [{"googleSearch": {}}],
        "generationConfig": {"maxOutputTokens": 3500, "temperature": 0.6, "topP": 0.9},
    }
    try:
        async with httpx.AsyncClient(timeout=httpx.Timeout(float(timeout), connect=min(8.0, float(timeout)))) as client:
            response = await client.post(url, headers=headers, json=payload)
    except httpx.TimeoutException as exc:
        raise GeminiRetryableError(f"Grounding search timeout for {model}.") from exc
    except httpx.HTTPError as exc:
        raise GeminiRetryableError(f"Grounding connection error for {model}.") from exc

    status = response.status_code
    if status in (408, 429, 500, 502, 503, 504):
        raise GeminiRetryableError(f"Grounding search unavailable: HTTP {status}.")
    if status == 404:
        raise GeminiModelUnavailable(f"Grounding model {model} not found.")
    try:
        result = response.json()
    except ValueError as exc:
        raise GeminiRetryableError(f"Grounding returned non-JSON HTTP {status}; trying another model.") from exc
    if status >= 400:
        detail = result.get("error", {}).get("message", "Unknown grounding error")
        raise GeminiAPIError(f"Grounding HTTP {status}: {str(detail)[:220]}")
    candidates = result.get("candidates") or []
    if not candidates:
        raise GeminiResponseError("Grounding returned no candidates.")
    candidate = candidates[0]
    parts = candidate.get("content", {}).get("parts", [])
    text_value = "\n".join(
        str(part.get("text", "")).strip()
        for part in parts if isinstance(part, dict) and isinstance(part.get("text"), str) and not part.get("thought")
    ).strip()
    grounding = candidate.get("groundingMetadata") or candidate.get("grounding_metadata") or {}
    chunks = grounding.get("groundingChunks") or grounding.get("grounding_chunks") or []
    sources = []
    seen = set()
    for chunk in chunks:
        web_info = chunk.get("web") or chunk.get("webSource") or {}
        uri = str(web_info.get("uri") or web_info.get("url") or "").strip()
        title = str(web_info.get("title") or "Source").strip()[:180]
        if uri.startswith("https://") and uri not in seen:
            seen.add(uri)
            sources.append({"id": f"S{len(sources) + 1}", "title": title, "url": uri})
        if len(sources) >= 8:
            break
    if not text_value:
        raise GeminiRetryableError("Grounding search returned empty research text.")
    log.info("Gemini grounding model=%s text_chars=%s sources=%s domains=%s", model, len(text_value), len(sources), [urlparse(x["url"]).netloc for x in sources])
    return text_value, sources


def _plain_meta(value: str, limit: int = 180) -> str:
    value = re.sub(r"<[^>]+>", " ", str(value or ""))
    value = html.unescape(re.sub(r"\s+", " ", value)).strip()
    return value[:limit]


def _http_user_agent() -> str:
    contact = WIKIMEDIA_CONTACT_EMAIL or "https://github.com/a88845837-coder/PrezentAI"
    return f"PrezentAI/1.0 ({contact})"


_BAD_IMAGE_TITLE_WORDS = (
    "map", "diagram", "logo", "chart", "poster", "flag", "coat of arms", "icon",
    "scheme", "drawing", "stamp", "banknote", "sketch", "screenshot", "infographic",
)


def _ensure_subject_in_query(query: str, subject: str) -> str:
    """Anchor a photo query to the topic's real subject so a banana deck never gets unrelated pictures."""
    query = " ".join(str(query or "").split())
    subject = " ".join(str(subject or "").split())
    if not query or not subject:
        return query
    lowered = query.lower()
    words = [w.lower() for w in re.findall(r"[A-Za-z0-9]{3,}", subject)]
    if not words or any(w[:5] in lowered for w in words):
        return query
    return " ".join((subject + " " + query).split()[:4])


async def _search_commons_photo(client: httpx.AsyncClient, query: str, exclude: set | None = None) -> dict | None:
    """Find a freely licensed image and machine-readable author/license details on Commons."""
    params = {
        "action": "query", "generator": "search", "gsrsearch": f"filetype:bitmap {query}",
        "gsrnamespace": 6, "gsrlimit": 15, "prop": "imageinfo",
        "iiprop": "url|size|extmetadata", "iiurlwidth": 1400, "format": "json", "formatversion": 2,
    }
    try:
        contact = WIKIMEDIA_CONTACT_EMAIL or "https://github.com/a88845837-coder/PrezentAI"
        if not WIKIMEDIA_CONTACT_EMAIL:
            log.warning("WIKIMEDIA_CONTACT_EMAIL is not set. Add a real maintainer email to the Render Environment variable WIKIMEDIA_CONTACT_EMAIL.")
        user_agent = f"PrezentAI/1.0 ({contact})"
        response = await client.get("https://commons.wikimedia.org/w/api.php", params=params,
                                   headers={"User-Agent": user_agent})
        response.raise_for_status()
        payload = response.json()
        for page in (payload.get("query", {}).get("pages") or []):
            infos = page.get("imageinfo") or []
            if not infos:
                continue
            info = infos[0]
            meta = info.get("extmetadata") or {}
            license_name = _plain_meta((meta.get("LicenseShortName") or {}).get("value", ""), 90)
            artist = _plain_meta((meta.get("Artist") or {}).get("value", ""), 130)
            url = info.get("thumburl") or info.get("url")
            if not (url and license_name):
                continue
            title = str(page.get("title") or "Wikimedia Commons image")
            title_words = {w.lower() for w in re.findall(r"[A-Za-z0-9]{3,}", title.replace("File:", " "))}
            query_words = {w.lower() for w in re.findall(r"[A-Za-z0-9]{3,}", query)}
            if query_words and not (title_words & query_words):
                continue
            if any(bad in title.lower() for bad in _BAD_IMAGE_TITLE_WORDS):
                continue
            width = int(info.get("width") or 0)
            height = int(info.get("height") or 0)
            if max(width, height) < 640 or min(width, height) < 300:
                continue
            page_url = "https://commons.wikimedia.org/wiki/" + quote(title.replace(" ", "_"), safe="_")
            if exclude and page_url in exclude:
                continue
            return {
                "image_url": url, "page_url": page_url, "provider": "Wikimedia Commons",
                "author": artist or "Author not listed", "license": license_name,
                "license_url": _plain_meta((meta.get("LicenseUrl") or {}).get("value", ""), 240),
                "credit": f"{artist or 'Author not listed'} — {license_name} — Wikimedia Commons",
            }
    except Exception as exc:
        log.warning("Wikimedia image search failed for %r: %s", query, exc)
    return None


async def _search_pexels_photo(client: httpx.AsyncClient, query: str, exclude: set | None = None) -> dict | None:
    """Optional Pexels search. Requires PEXELS_API_KEY; Commons works without a key."""
    if not PEXELS_API_KEY:
        return None
    try:
        response = await client.get(
            "https://api.pexels.com/v1/search",
            params={"query": query, "orientation": "landscape", "per_page": 10},
            headers={"Authorization": PEXELS_API_KEY},
        )
        response.raise_for_status()
        photos = response.json().get("photos") or []
        for photo in photos:
            src = photo.get("src") or {}
            image_url = src.get("large") or src.get("landscape") or src.get("medium")
            if not image_url:
                continue
            photographer = str(photo.get("photographer") or "Pexels photographer")[:120]
            photographer_url = str(photo.get("photographer_url") or "https://www.pexels.com")
            page_url = str(photo.get("url") or "https://www.pexels.com")
            if exclude and page_url in exclude:
                continue
            return {
                "image_url": image_url, "page_url": page_url, "provider": "Pexels",
                "author": photographer, "author_url": photographer_url, "license": "Pexels License",
                "credit": f"Photo by {photographer} on Pexels",
            }
    except Exception as exc:
        log.warning("Pexels image search failed for %r: %s", query, exc)
    return None


def _save_optimized_photo_file(source_path: str, output_path: str):
    """Decode a downloaded image only after format and pixel limits have been checked."""
    from PIL import Image, ImageOps

    allowed_formats = {"JPEG", "PNG", "WEBP"}
    with Image.open(source_path) as probe:
        if probe.format not in allowed_formats:
            raise ValueError(f"unsupported image format: {probe.format}")
        width, height = probe.size
        if width < 320 or height < 200 or width * height > MAX_IMAGE_PIXELS:
            raise ValueError(f"unacceptable image dimensions: {width}x{height}")
        probe.verify()

    with Image.open(source_path) as im:
        im = ImageOps.exif_transpose(im).convert("RGB")
        im.thumbnail((1600, 1100))
        im.save(output_path, "JPEG", quality=86, optimize=True, progressive=True)
    return output_path


def _trusted_photo_url(url: str) -> bool:
    try:
        parsed = urlparse(str(url or ""))
        host = (parsed.hostname or "").lower().rstrip(".")
        allowed = ("upload.wikimedia.org", "images.pexels.com")
        return parsed.scheme == "https" and any(host == domain or host.endswith("." + domain) for domain in allowed)
    except Exception:
        return False


async def download_slide_images(deck: dict, directory: str) -> list[dict]:
    """Download at most five slide photos; missing photos never leave empty placeholders."""
    slides = deck.get("slides", [])
    priorities = {"cover": 0, "definition": 1, "origin": 2, "example": 3, "mechanics": 4}
    indexed = [
        (i, slide) for i, slide in enumerate(slides)
        if re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9 -]{1,60}", str(slide.get("image_query", "")).strip())
        and resolve_render_layout(slide) in ("cover", "photo_text", "portrait_fact")
    ]
    indexed.sort(key=lambda pair: (priorities.get(pair[1].get("section_type"), 10), pair[0]))
    selected = indexed[:MAX_DECK_IMAGES]
    if not selected:
        return []
    limiter = asyncio.Semaphore(3)
    used_pages: set = set()
    credits = []
    async with httpx.AsyncClient(
        timeout=httpx.Timeout(12.0, connect=7.0), follow_redirects=True,
        headers={"User-Agent": _http_user_agent(), "Accept": "image/*,application/json;q=0.9,*/*;q=0.5"},
    ) as client:
        async def one(index: int, slide: dict):
            query = " ".join(str(slide.get("image_query", "")).split()[:4])
            kind = slide.get("image_kind", "generic")
            subject = str(deck.get("photo_subject") or "").strip()
            subject_kind = deck.get("photo_subject_kind") or "generic"

            async def find_photo(q: str, k: str):
                # People are searched only on Commons, so a portrait is never a random stock photo.
                if k == "person":
                    return await _search_commons_photo(client, q, used_pages)
                if k == "named_subject":
                    found = await _search_commons_photo(client, q, used_pages)
                    if not found and PEXELS_API_KEY:
                        found = await _search_pexels_photo(client, q, used_pages)
                    return found
                found = await _search_pexels_photo(client, q, used_pages) if PEXELS_API_KEY else None
                if not found:
                    found = await _search_commons_photo(client, q, used_pages)
                return found

            # Try the slide query, then a shorter one, then the topic's own subject.
            attempts = [(query, kind)]
            short_query = " ".join(query.split()[:2])
            if short_query and short_query != query:
                attempts.append((short_query, kind))
            if subject and all(subject.lower() != q.lower() for q, _ in attempts):
                attempts.append((subject, subject_kind))
            async def download(found: dict) -> str:
                image_url = str(found.get("image_url") or "")
                if not _trusted_photo_url(image_url):
                    raise ValueError("image URL is outside the approved provider CDN domains")
                path = str(Path(directory) / f"slide_{index + 1:02d}.jpg")
                source_path = str(Path(directory) / f"slide_{index + 1:02d}.download")
                try:
                    for attempt_no in range(2):
                        try:
                            total_bytes = 0
                            async with client.stream("GET", image_url) as response:
                                response.raise_for_status()
                                final_url = str(response.url)
                                if not _trusted_photo_url(final_url):
                                    raise ValueError("image redirect left the approved CDN domains")
                                content_type = response.headers.get("content-type", "").split(";", 1)[0].strip().lower()
                                if content_type not in ("image/jpeg", "image/png", "image/webp"):
                                    raise ValueError(f"unsupported image content type: {content_type or 'missing'}")
                                declared = response.headers.get("content-length")
                                if declared and int(declared) > MAX_IMAGE_DOWNLOAD_BYTES:
                                    raise ValueError("image exceeds the 8 MB download limit")
                                with open(source_path, "wb") as output:
                                    async for chunk in response.aiter_bytes(64 * 1024):
                                        total_bytes += len(chunk)
                                        if total_bytes > MAX_IMAGE_DOWNLOAD_BYTES:
                                            raise ValueError("image exceeded the 8 MB streaming limit")
                                        output.write(chunk)
                            if total_bytes < 1024:
                                raise ValueError("downloaded image is too small or empty")
                            photo_deadline = asyncio.get_running_loop().time() + 45.0
                            await run_blocking_stage_with_deadline(
                                _save_optimized_photo_file, (source_path, path),
                                photo_deadline, "photo decode and optimization", 45.0,
                            )
                            return path
                        except httpx.HTTPStatusError as exc:
                            if exc.response.status_code not in (429, 500, 502, 503, 504) or attempt_no:
                                raise
                            await asyncio.sleep(1.5)
                finally:
                    try:
                        Path(source_path).unlink(missing_ok=True)
                    except OSError:
                        log.warning("Could not remove temporary image download %s", source_path)

            async with limiter:
                for attempt_query, attempt_kind in attempts:
                    for _ in range(3):  # a few candidates per query if a download fails
                        found = await find_photo(attempt_query, attempt_kind)
                        if not found or found["page_url"] in used_pages:
                            break
                        used_pages.add(found["page_url"])
                        try:
                            path = await download(found)
                        except Exception as exc:
                            log.warning("Photo download failed for slide %s (%s, %s): %s",
                                        index + 1, found.get("provider"), found.get("image_url"), exc)
                            continue  # the page is marked used, so the next search returns another candidate
                        slide["image_path"] = path
                        slide["image_provider"] = found["provider"]
                        slide["image_credit"] = found["credit"]
                        slide["image_page_url"] = found["page_url"]
                        slide["image_license"] = found.get("license", "")
                        if found["provider"] == "Pexels":
                            slide["image_author_url"] = found.get("author_url", "")
                        return {
                            "slide": index + 1, "provider": found["provider"],
                            "credit": found["credit"], "page_url": found["page_url"],
                            "license": found.get("license", ""),
                            "license_url": found.get("license_url", ""),
                            "author_url": found.get("author_url", ""),
                        }
                log.warning("No usable photo for slide %s (queries tried: %s)", index + 1, [q for q, _ in attempts])
                return None
        results = await asyncio.gather(*(one(i, slide) for i, slide in selected), return_exceptions=True)
    for result in results:
        if isinstance(result, dict):
            credits.append(result)
    deck["image_credits"] = credits
    log.info("Images downloaded for %s/%s selected slide queries", len(credits), len(selected))
    return credits


def _build_reference_lines(sources: list[dict], image_credits: list[dict], language: str = "uk") -> list[str]:
    """Build readable references using localized labels; full URLs remain intact."""
    labels = {
        "uk": {"source": "Джерело", "photo": "Фото, слайд", "author": "автор", "license": "ліцензія"},
        "ru": {"source": "Источник", "photo": "Фото, слайд", "author": "автор", "license": "лицензия"},
        "en": {"source": "Source", "photo": "Photo, slide", "author": "author", "license": "license"},
    }
    lang = language if language in labels else "uk"
    label = labels[lang]
    lines = []
    for source in (sources or [])[:4]:
        lines.append(f"[{source.get('id', 'S')}] {source.get('title', label['source'])} — {source.get('url', '')}")
    for credit in (image_credits or [])[:5]:
        author_url = f" ({label['author']}: {credit.get('author_url')})" if credit.get("author_url") else ""
        license_url = f" ({label['license']}: {credit.get('license_url')})" if credit.get("license_url") else ""
        credit_text = str(credit.get("credit") or "")
        if credit_text.lower().startswith("author not listed"):
            credit_text = deck_label(lang, "author_unknown") + credit_text[len("Author not listed"):]
        lines.append(f"{label['photo']} {credit.get('slide')}: {credit_text} — {credit.get('page_url')}{author_url}{license_url}")
    return lines[:9]
def _limit_note_sentences(value: str, max_sentences: int, max_chars: int) -> str:
    """Keep speaker notes within the selected detail profile without inventing text."""
    text = re.sub(r"\s+", " ", str(value or "")).strip()
    if not text:
        return ""
    sentences = re.split(r"(?<=[.!?])\s+", text)
    text = " ".join(sentences[:max_sentences]).strip()
    return text[:max_chars].rstrip()


async def _notify_generation_stage(callback, stage: str):
    if callback is None:
        return
    result = callback(stage)
    if asyncio.iscoroutine(result):
        await result


async def generate_content(data: dict, progress_callback=None) -> dict:
    topic, count = data["topic"], int(data["slides"])
    detail_key = data.get("detail", "medium")
    detail_profile = DETAIL_PROFILES.get(detail_key, DETAIL_PROFILES["medium"])
    if detail_key == "detailed":
        content_density_guidance = (
            "On every content slide, write exactly 5 distinct, informative bullet points. "
            "Each point must be a complete sentence of about 100-155 characters, explaining a fact, "
            "how something works, why it matters, or a concrete example. Avoid vague labels and repeated ideas."
        )
    elif detail_key == "medium":
        content_density_guidance = (
            "On every content slide, write exactly 4 distinct, useful bullet points. "
            "Each point should be a clear complete sentence of about 70-125 characters. "
            "Explain the key idea briefly; avoid one-word labels and repeated ideas."
        )
    else:
        content_density_guidance = (
            "On every content slide, write 2-3 concise bullet points. Keep only the essentials."
        )
    audience_guidance = AUDIENCE_GUIDANCE.get(data.get("audience", "general"), AUDIENCE_GUIDANCE["general"])
    revision_request = str(data.get("revision_request") or "").strip()[:800]
    previous_outline = str(data.get("previous_outline") or "").strip()[:1800]
    max_output_tokens = detail_output_tokens(detail_key, count)
    request_timeout = float(detail_profile["request_timeout"])
    await _notify_generation_stage(progress_callback, "research")
    language = LANGUAGES.get(data["language"], "українською")
    design = data["design"]
    style_name = design_name(design, data.get("interface_language", "uk"))
    style_direction = STYLE_PROMPTS.get(design, STYLE_PROMPTS["minimal"])
    visual_concept = random.choice(DECK_VISUAL_CONCEPTS)
    variation_id = uuid.uuid4().hex[:12]
    visual_seed = random.randint(1, 2_000_000_000)
    photos_enabled = bool(data.get("photos", True))
    if photos_enabled:
        photo_rules = """- ГОЛОВНЕ: фото мають показувати саме те, про що тема. Заповни photo_subject (1–3 слова англійською — головний реальний об'єкт теми; для теми 'банан' це 'banana'), а кожен image_query про предмет теми має містити цей об'єкт (наприклад 'banana fruit', 'banana plantation', 'banana tree').
- Не вигадуй абстрактних запитів (ідея, успіх, команда, бізнес): лише реальні речі, місця, предмети чи люди, які справді є на слайді.
- Для кожного доречного слайда вкажи image_query англійською, 2–4 прості слова (наприклад 'solar panel rooftop', 'Ada Lovelace portrait', 'electric car factory').
- image_kind має бути person для історичної/реальної людини, named_subject для конкретного продукту/компанії/винаходу, generic для загальної теми або none, якщо фото недоречне.
- Для людини шукається реальний портрет у Wikimedia Commons. Ніколи не генеруй портрет реальної людини.
- Якщо ім'я/рік невідомі, image_kind='none' та image_query порожній для слайда походження.
- Система завантажить не більше п'яти фото на всю презентацію; не намагайся заповнити кожен слайд випадковим зображенням."""
    else:
        photo_rules = """- Користувач НЕ хоче фотографій. Для КОЖНОГО слайда image_query має бути порожнім рядком, а image_kind='none'.
- Не використовуй layout photo_text і portrait_fact; обирай макети без фото (two_col, cards, focus, grid3, steps тощо)."""

    # First request: grounded fact research. If Google Search grounding is unavailable,
    # factual fields (creator/year/money) are explicitly marked unknown in the deck request.
    research_text = ""
    research_sources = []
    research_status = "unavailable"
    research_budget_seconds = 25.0 if detail_key == "detailed" else 30.0
    research_deadline = time.monotonic() + research_budget_seconds
    research_prompt = f"""
Проведи коротке фактологічне дослідження теми: {topic}
Мова відповіді: {language}. Використовуй Google Search.
Потрібно перевірити:
1) що саме означає тема й чітке визначення;
2) хто створив/винайшов/започаткував явище, якщо це має конкретного автора, і підтверджений рік;
3) як воно працює — тільки ключові факти;
4) один реальний приклад компанії, продукту або організації;
5) реальні цифри про ціну, стартові витрати, дохід, ринок або економічні показники, якщо є надійні джерела;
6) реалістичні кроки початку та витрати, якщо вони опубліковані.

Використовуй авторитетні джерела: офіційні сайти, документи, дослідження, статистику, надійні медіа.
Не домислюй автора чи рік, якщо походження теми складне або спірне. Не вигадуй сум і не перетворюй оцінки на точні цифри.
Поверни стислий текстовий бриф з фактами та обмеженнями. Не вигадуй URL: джерела будуть окремо взяті з metadata Google Search.
""".strip()
    models = list(dict.fromkeys([m for m in (GEMINI_MODEL, GEMINI_FALLBACK_MODEL) if m]))
    research_models = list(dict.fromkeys([m for m in (GEMINI_RESEARCH_MODEL, GEMINI_MODEL, GEMINI_FALLBACK_MODEL) if m]))
    if detail_key == "detailed":
        research_models = research_models[:1]
    for research_model in research_models:
        remaining_research = research_deadline - time.monotonic()
        if remaining_research <= 1.0:
            log.warning("Grounding research budget exhausted before model=%s", research_model)
            break
        per_model_timeout = min(25.0 if detail_key == "detailed" else 10.0, remaining_research)
        try:
            research_text, research_sources = await asyncio.wait_for(
                request_gemini_research(research_model, research_prompt, timeout=per_model_timeout),
                timeout=remaining_research,
            )
            if research_sources:
                research_status = "grounded"
                break
            log.warning("Gemini grounding returned no source URLs for model=%s; facts will be treated as unverified", research_model)
            research_text = ""
        except (GeminiAPIError, GeminiResponseError, asyncio.TimeoutError) as exc:
            log.warning("Gemini grounding failed for model=%s: %s", research_model, exc)
    if research_status != "grounded":
        research_fallback = {
            "uk": "У цьому запиті не вдалося знайти підтверджені джерела. Не подавай точні імена, роки, ціни, ринкові суми чи доходи як факти. Для автора, року та грошових показників напиши «невідомо» або нейтрально поясни, що надійних даних не знайдено.",
            "ru": "В этом запросе не удалось найти подтверждённые источники. Не указывай точные имена, годы, цены, объём рынка или доходы как факты. Для автора, года и денежных показателей напиши «неизвестно» или нейтрально объясни, что надёжных данных не найдено.",
            "en": "No verified sources were available for this request. Do not present exact names, years, prices, market sizes, or revenue as facts. For the originator, year, and financial figures, write “unknown” or clearly state that reliable data was not found.",
        }
        research_text = research_fallback.get(data.get("language", "uk"), research_fallback["uk"])

    no_sources_text = {
        "uk": "Підтверджених вебджерел немає. Не створюй фальшивих посилань або номерів джерел.",
        "ru": "Подтверждённых веб-источников нет. Не создавай фиктивные ссылки или идентификаторы источников.",
        "en": "No verified web sources are available. Do not invent links or source IDs.",
    }
    sources_block = "\n".join(
        f"[{source['id']}] {source['title']} — {source['url']}" for source in research_sources
    ) or no_sources_text.get(data.get("language", "uk"), no_sources_text["uk"])

    if count == 6:
        framework = """Use this story arc for exactly 6 slides:
1. COVER: a memorable, topic-specific title and a short hook.
2. DEFINITION: explain what the topic is in plain language and why it matters.
3. PHOTO AND REAL EXAMPLE: show the real subject or a concrete use; request a relevant photo when photos are enabled.
4. ORIGIN AND CREATOR: explain who created/invented it and the person's relevant story if reliably known. If there is no single creator, explain how it emerged instead.
5. HOW IT WORKS + INTERESTING FACTS: explain the key mechanism and include several surprising, verified facts.
6. CONCLUSION AND SOURCES: a clear takeaway and verified sources."""
    elif count == 8:
        framework = """Use this story arc for exactly 8 slides:
1. COVER: a memorable, topic-specific title and hook.
2. DEFINITION: what it is and why it matters.
3. PHOTO-LED EXAMPLE: a relevant real photo and a concrete example or use.
4. ORIGIN: who invented/created it and when, only when verified; otherwise explain its origin without inventing a person.
5. THE CREATOR: a short, engaging account of the verified creator's contribution and context; if no single creator exists, explain the key people or stages in its development.
6. HOW IT WORKS: explain the core mechanism clearly.
7. INTERESTING FACTS: several surprising, useful, verified facts with a strong slide headline.
8. CONCLUSION AND SOURCES: one memorable takeaway and sources."""
    else:
        extras = count - 9
        framework = (
            "Use this exact story arc: 1 cover with a memorable hook; 2 plain-language definition and importance; "
            "3 photo-led real example (request a relevant photo when enabled); 4 verified origin/inventor, or explain that there is no single inventor; "
            "5 an engaging story about the creator or the topic's development; 6 how it works; 7 practical real-world example; "
            f"then {extras} distinct topic-specific deep dives (such as impact, case study, risks, comparison, or future); "
            f"slide {count - 1} must be an INTERESTING FACTS slide with several surprising, verified facts; "
            f"slide {count} is the conclusion and sources. Exactly {count} slides."
        )

    await _notify_generation_stage(progress_callback, "writing")
    prompt = f"""
Створи професійну, змістовну презентацію для реальної аудиторії.
ТЕМА: {topic}
СЛАЙДІВ: рівно {count}
МОВА КОНТЕНТУ: {language}
МОВНА ЦІЛІСНІСТЬ: усі видимі написи слайдів — title, subtitle, kicker, left_title, right_title, fact, stat_label, highlight і content — та нотатки мають бути ТІЛЬКИ мовою презентації. Не вставляй українські слова в російську чи англійську презентацію.
ВІЗУАЛЬНИЙ СТИЛЬ: {style_name} ({design})
АРТ-ДИРЕКШН: {style_direction}
РІВЕНЬ ДЕТАЛІЗАЦІЇ: {detail_key}
ПРАВИЛО ДЕТАЛІЗАЦІЇ: {detail_profile["prompt"].get(data.get("language", "uk"), detail_profile["prompt"]["uk"])}
CONTENT DENSITY REQUIREMENT: {content_density_guidance}
STORYTELLING: Write a connected, lively explanation for a real audience, not a stack of labels or generic bullet fragments. Give each slide one clear point, use concrete examples, and make the transition from definition to origin, example, explanation, and surprising facts feel natural. Use specific, engaging slide titles rather than generic headings. Exception: slide 3 is the photo-led example and should use 3-4 short, visual-friendly points.
AUDIENCE AND PURPOSE: {audience_guidance}
REVISION REQUEST: {revision_request or "None"}
PREVIOUS OUTLINE FOR REVISION: {previous_outline or "None"}
When revising, apply the request to the specified slide(s), preserve topic and slide count, and keep unaffected sections consistent with the previous outline.
Keep text concise enough to fit the slide, but never use fragments just to fill space. Speaker notes add further explanation and examples.
Дотримуйся цього ліміту окремо для кожного слайда; не збільшуй кількість тез понад профіль.

ОБОВ'ЯЗКОВА СТРУКТУРА:
{framework}

ФАКТОЛОГІЧНИЙ БРИФ З ОКРЕМОГО GOOGLE SEARCH ЗАПИТУ:
{research_text}

ДОЗВОЛЕНІ ДЖЕРЕЛА (використовуй тільки ці точні ідентифікатори й URL; не придумуй інші):
{sources_block}

ПРАВИЛА ФАКТІВ І ГРОШЕЙ:
- Для кожного слайда заповни поле fact коротким фактом/твердженням, якщо воно має підтвердження; інакше напиши 'невідомо' мовою презентації.
- Поле source має містити лише точний ID з переліку вище, наприклад [S1]. Якщо підтвердження немає — порожній рядок.
- Не вигадуй ім'я винахідника, автора, рік, ціну, бюджет, дохід, ринкову суму, частку ринку, цитату, назву джерела чи URL.
- Якщо походження теми не має одного загальновизнаного автора, явно напиши про це; не вигадуй портрет.
- Для грошей stat_value заповнюй тільки тоді, коли конкретне число прямо підтверджене в дослідницькому брифі та має source ID; інакше stat_value має бути порожнім і не малюй stat-блок.
- Якщо дані про витрати не знайдені, напиши 'немає надійно підтверджених даних' та не оцінюй суму навмання.

ПРАВИЛА ФОТО:
{photo_rules}

ВІЗУАЛЬНА РІЗНОМАНІТНІСТЬ:
- Ця презентація має окрему концепцію: {visual_concept}. Не повторюй однакову композицію на сусідніх слайдах.
- Фото + текст використовуй для визначення та прикладу; портрет + дата — для походження; stat — тільки для перевіреної цифри; steps — для практичних кроків.
- Кожен слайд має іншу інформаційну роль; не перетворюй усе на список.

ПОВЕРНИ ЛИШЕ JSON:
{{
  "title": "коротка сильна назва",
  "photo_subject": "1 to 3 English words: the main real-world subject of the topic, or empty",
  "photo_subject_kind": "person|named_subject|generic",
  "slides": [
    {{
      "section_type": "cover|definition|origin|mechanics|example|interesting_facts|money|steps|extra|conclusion_sources",
      "title": "конкретний заголовок",
      "subtitle": "короткий підзаголовок або порожньо",
      "layout": "cover|photo_text|portrait_fact|two_col|grid3|stat|steps|section|closing|cards|timeline|comparison|focus",
      "content": ["коротка теза 1", "коротка теза 2"],
      "fact": "короткий перевірений факт або невідомо",
      "source": "[S1] або порожній рядок",
      "image_query": "2 to 4 English words",
      "image_kind": "person|named_subject|generic|none",
      "kicker": "короткий тематичний лейбл",
      "highlight": "важлива думка або порожньо",
      "stat_value": "тільки підтверджене число або порожньо",
      "stat_label": "пояснення числа або порожньо",
      "left_title": "заголовок ліворуч",
      "right_title": "заголовок праворуч",
      "notes": "нотатки доповідача згідно з вибраним рівнем деталізації, не повторюй слайд"
    }}
  ]
}}

ЗАВЕРШЕННЯ: для останнього слайда section_type='conclusion_sources', короткий підсумок і practical takeaway. Джерела й автори фото будуть додані програмно з перевірених URL.
Не додавай пояснень поза JSON. Мова всього контенту та нотаток — {language}.
""".strip()

    last_error = None
    for index, model in enumerate(models):
        try:
            try:
                parsed = await request_gemini(
                    model, prompt,
                    max_tokens=max_output_tokens,
                    timeout=request_timeout,
                )
            except GeminiResponseError as token_exc:
                if "ліміту вихідних токенів" not in str(token_exc).lower() and "max_tokens" not in str(token_exc).lower():
                    raise
                retry_token_budget = min(GEMINI_RETRY_OUTPUT_TOKEN_LIMIT, int(max_output_tokens * 1.5))
                if retry_token_budget <= max_output_tokens:
                    raise
                log.warning("Gemini MAX_TOKENS; retrying model=%s with output budget %s", model, retry_token_budget)
                parsed = await request_gemini(
                    model, prompt,
                    max_tokens=retry_token_budget,
                    timeout=request_timeout,
                )
            raw_slides = parsed.get("slides")
            if not isinstance(raw_slides, list) or not raw_slides:
                raise GeminiResponseError("Gemini не повернув список слайдів.")
            raw_slides = raw_slides[:count]
            language_code = data.get("language", "uk")
            closing_titles = {
                "uk": "Висновок і джерела",
                "ru": "Выводы и источники",
                "en": "Conclusion and sources",
            }
            filler_titles = {
                "uk": {"cover": "Тема презентації", "definition": "Ключове визначення", "origin": "Походження теми", "mechanics": "Як це працює", "example": "Практичний приклад", "money": "Фінансовий контекст", "steps": "Практичні кроки", "extra": "Додатковий аспект"},
                "ru": {"cover": "Тема презентации", "definition": "Ключевое определение", "origin": "Происхождение темы", "mechanics": "Как это работает", "example": "Практический пример", "money": "Финансовый контекст", "steps": "Практические шаги", "extra": "Дополнительный аспект"},
                "en": {"cover": "Presentation topic", "definition": "Key definition", "origin": "Origins of the topic", "mechanics": "How it works", "example": "Practical example", "money": "Financial context", "steps": "Practical steps", "extra": "Additional perspective"},
            }
            fallback_points = {
                "uk": "Додаткові підтверджені відомості для цього розділу не знайдено.",
                "ru": "Дополнительные подтверждённые сведения для этого раздела не найдены.",
                "en": "No additional verified information was found for this section.",
            }
            filler_notes = {
                "uk": "Не додавайте неперевірених фактів. Поясніть, яких даних бракує та що варто перевірити окремо.",
                "ru": "Не добавляйте непроверенных фактов. Объясните, каких данных не хватает и что следует проверить отдельно.",
                "en": "Do not add unverified facts. Explain what information is missing and what should be checked separately.",
            }
            if len(raw_slides) != count:
                raise GeminiResponseError(f"Gemini returned {len(raw_slides)} slides; {count} required")
            normalized = []
            valid_source_ids = {item["id"] for item in research_sources}
            photo_subject = " ".join(re.findall(r"[A-Za-z0-9-]+", str(parsed.get("photo_subject") or ""))[:3]) if photos_enabled else ""
            photo_subject_kind = str(parsed.get("photo_subject_kind") or "generic").lower()
            if photo_subject_kind not in ("person", "named_subject", "generic"):
                photo_subject_kind = "generic"
            for n, slide in enumerate(raw_slides[:count], 1):
                if not isinstance(slide, dict):
                    raise GeminiResponseError(f"Слайд {n} має неправильний формат.")
                points = slide.get("content", [])
                if isinstance(points, str):
                    points = [points]
                if not isinstance(points, list):
                    points = []
                points = [
                    truncate_at_word(str(v), detail_profile["max_point_chars"])
                    for v in points if str(v).strip()
                ][:detail_profile["max_points"]]
                if n == 3 and photos_enabled:
                    points = [truncate_at_word(point, 115) for point in points[:4]]
                section_hint = str(slide.get("section_type") or "").strip().lower()
                if section_hint not in ("cover", "conclusion_sources"):
                    required_points = 3 if n == 3 and photos_enabled else detail_profile["min_points"]
                    if len(points) < required_points:
                        raise GeminiResponseError(
                            f"Slide {n} has {len(points)} points; {required_points} required for {detail_key} detail"
                        )
                    minimum_point_length = 0 if n == 3 else (95 if detail_key == "detailed" else 65 if detail_key == "medium" else 0)
                    short_point_count = sum(len(point) < minimum_point_length for point in points)
                    if short_point_count:
                        log.warning(
                            "Gemini returned %s short points: slide=%s detail=%s threshold=%s",
                            short_point_count, n, detail_key, minimum_point_length,
                        )
                if not points:
                    fallback_points = {
                        "uk": "Ключову інформацію не вдалося підтвердити.",
                        "ru": "Ключевую информацию не удалось подтвердить.",
                        "en": "The key information could not be verified.",
                    }
                    points = [fallback_points.get(data.get("language", "uk"), fallback_points["uk"])]
                section_type = str(slide.get("section_type") or "extra").strip().lower()
                allowed_sections = {"cover", "definition", "origin", "mechanics", "example", "interesting_facts", "money", "steps", "extra", "conclusion_sources"}
                if section_type not in allowed_sections:
                    section_type = "extra"
                source = str(slide.get("source") or "").strip().upper()
                if source and source not in valid_source_ids and source.strip("[]") not in valid_source_ids:
                    source = ""
                elif source and not source.startswith("["):
                    source = f"[{source}]"
                language_code = data.get("language", "uk")
                unknown_fact = {"uk": "невідомо", "ru": "неизвестно", "en": "unknown"}.get(language_code, "невідомо")
                fact_value = str(slide.get("fact") or "").strip()[:230]
                stat_value = str(slide.get("stat_value") or "").strip()[:35]
                # Creator/year and money claims are not shown as factual without a grounded source ID.
                if section_type == "origin" and not source:
                    fact_value = unknown_fact
                if section_type in ("money", "steps") and not source:
                    if section_type == "money":
                        stat_value = ""
                    points = [point for point in points if not re.search(r"(?:\d|[$€₴₽%]|\b(?:USD|EUR|UAH)\b)", point, re.IGNORECASE)]
                    if not points:
                        points = [{"uk": "Немає надійно підтверджених сум або витрат.", "ru": "Нет надёжно подтверждённых сумм или расходов.", "en": "No reliably verified prices or costs were found."}.get(language_code, "Немає надійно підтверджених сум або витрат.")]
                image_query = " ".join(re.findall(r"[A-Za-z0-9-]+", str(slide.get("image_query") or ""))[:4])
                image_kind = str(slide.get("image_kind") or "none").lower()
                if image_kind not in ("person", "named_subject", "generic", "none"):
                    image_kind = "none"
                if image_kind == "none" or not photos_enabled:
                    image_query = ""
                    image_kind = "none"
                elif photo_subject and image_kind != "person":
                    image_query = _ensure_subject_in_query(image_query, photo_subject)
                # A money statistic without a valid grounded source is always suppressed.
                if section_type == "money" and (not stat_value or not source):
                    stat_value = ""
                normalized_notes = _limit_note_sentences(
                    slide.get("notes", ""),
                    detail_profile["max_note_sentences"],
                    detail_profile["max_note_chars"],
                )
                note_sentence_count = len([part for part in re.split(r"(?<=[.!?])\s+", normalized_notes) if part.strip()])
                if note_sentence_count < detail_profile["min_note_sentences"]:
                    log.warning(
                        "Gemini returned fewer speaker-note sentences than requested: slide=%s detail=%s count=%s min=%s",
                        n, detail_key, note_sentence_count, detail_profile["min_note_sentences"],
                    )
                deck_language = data.get("language", "uk")
                localized_points = [localize_known_phrases(point, deck_language) for point in points]
                normalized.append({
                    "section_type": section_type,
                    "title": localize_known_phrases(str(slide.get("title") or ({"uk": "Слайд", "ru": "Слайд", "en": "Slide"}.get(deck_language, "Slide") + f" {n}")).strip()[:140], deck_language),
                    "subtitle": localize_known_phrases(str(slide.get("subtitle") or "").strip()[:160], deck_language),
                    "content": localized_points,
                    "highlight": localize_known_phrases(str(slide.get("highlight") or "").strip()[:210], deck_language),
                    "left_title": localize_known_phrases(str(slide.get("left_title") or "").strip()[:70], deck_language),
                    "right_title": localize_known_phrases(str(slide.get("right_title") or "").strip()[:70], deck_language),
                    "kicker": localize_known_phrases(str(slide.get("kicker") or "").strip()[:60], deck_language),
                    "stat_value": stat_value,
                    "stat_label": localize_known_phrases(str(slide.get("stat_label") or "").strip()[:100], deck_language) if stat_value else "",
                    "fact": localize_known_phrases(fact_value, deck_language),
                    "source": source,
                    "image_query": image_query,
                    "image_kind": image_kind if image_query else "none",
                    "notes": localize_known_phrases(normalized_notes, deck_language),
                    "layout": str(slide.get("layout") or "").strip().lower(),
                })
            # Enforce the fixed content architecture, while allowing the UI's 6/8/10/12 slide options.
            if normalized:
                if count == 6:
                    canonical_roles = ["cover", "definition", "example", "origin", "interesting_facts", "conclusion_sources"]
                elif count == 8:
                    canonical_roles = ["cover", "definition", "example", "origin", "origin", "mechanics", "interesting_facts", "conclusion_sources"]
                else:
                    canonical_roles = ["cover", "definition", "example", "origin", "origin", "mechanics", "example"]
                    canonical_roles += ["extra"] * max(0, count - 9)
                    canonical_roles += ["interesting_facts", "conclusion_sources"]
                for role_index, role in enumerate(canonical_roles[:len(normalized)]):
                    normalized[role_index]["section_type"] = role
                normalized[0]["layout"] = "cover"
                normalized[-1]["section_type"] = "conclusion_sources"
                normalized[-1]["layout"] = "closing"
            if photos_enabled:
                # Reuse the topic subject/query when the model omits the required visual query.
                fallback_photo_query = photo_subject or next(
                    (item.get("image_query") for item in normalized if item.get("image_query")), ""
                )
                fallback_photo_kind = photo_subject_kind if photo_subject else "generic"
                for item in normalized:
                    if item.get("section_type") in ("cover", "definition", "example") and not item.get("image_query") and fallback_photo_query:
                        item["image_query"] = fallback_photo_query
                        item["image_kind"] = fallback_photo_kind
            preferred_layouts = {
                "definition": "photo_text", "origin": "portrait_fact", "mechanics": "two_col",
                "example": "photo_text", "interesting_facts": "steps", "money": "stat", "steps": "steps",
                "extra": "grid3", "conclusion_sources": "closing",
            }
            for item in normalized[1:-1]:
                preferred = preferred_layouts.get(item["section_type"])
                if item["section_type"] == "money" and not item.get("stat_value"):
                    preferred = "focus"
                if item["image_kind"] == "none" and item["layout"] in ("photo_text", "portrait_fact"):
                    item["layout"] = preferred if preferred not in ("photo_text", "portrait_fact") else "two_col"
                elif preferred:
                    item["layout"] = preferred
                if not photos_enabled and item["layout"] in ("photo_text", "portrait_fact"):
                    item["layout"] = "two_col"
                if item["section_type"] == "origin" and (not item.get("source") or not item.get("fact") or "невідом" in item.get("fact", "").lower() or "неизвест" in item.get("fact", "").lower() or "unknown" in item.get("fact", "").lower()):
                    item["fact"] = {"uk": "Невідомо або немає єдиного підтвердженого автора.", "ru": "Неизвестно или нет единого подтверждённого автора.", "en": "Unknown or no single verified originator."}.get(data.get("language", "uk"), "Невідомо")
                    item["image_query"] = ""
                    item["image_kind"] = "none"
                    item["layout"] = "two_col"
                if item["section_type"] == "money" and (not item.get("stat_value") or not item.get("source")):
                    item["stat_value"] = ""
                    item["stat_label"] = ""
                    item["layout"] = "focus"
            # Vary extra-slide compositions per deck, then route each slide to a renderer that can show every point.
            extra_pool = ["grid3", "timeline", "cards", "quote", "focus", "section", "two_col", "comparison"]
            extra_rng = random.Random(visual_seed)
            extra_previous = None
            for item in normalized[1:-1]:
                if item.get("section_type") == "extra":
                    choices = [layout for layout in extra_pool if layout != extra_previous]
                    if len(item.get("content", [])) < 3:
                        choices = [layout for layout in choices if layout != "grid3"]
                    if not choices:
                        choices = ["two_col"]
                    item["layout"] = extra_rng.choice(choices)
                    extra_previous = item["layout"]

            for item in normalized:
                item["layout"] = resolve_render_layout(item)
                capacity = layout_point_capacity(item)
                if capacity > 0 and len(item.get("content", [])) > capacity:
                    # Never silently drop generated points. The steps renderer supports five.
                    log.warning(
                        "Rerouting slide to full-width steps to preserve all points: title=%r layout=%s points=%s capacity=%s",
                        item.get("title", ""), item["layout"], len(item.get("content", [])), capacity,
                    )
                    item["layout"] = "steps"
                    capacity = LAYOUT_POINT_CAPACITY["steps"]
                if len(item.get("content", [])) > capacity and capacity > 0:
                    # The detail profile already caps at five; keep an error log if future
                    # schemas ever exceed the renderer's hard physical maximum.
                    log.error("Slide contains more points than any renderer supports: title=%r points=%s", item.get("title", ""), len(item.get("content", [])))
            if photos_enabled and len(normalized) >= 3:
                normalized[2]["layout"] = "photo_text"
                normalized[2]["content"] = [truncate_at_word(point, 115) for point in normalized[2].get("content", [])[:4]]
            section_types = [item.get("section_type") for item in normalized]
            if len(normalized) != count or section_types[0] != "cover" or section_types[1] != "definition":
                raise GeminiResponseError("Presentation is missing the required cover or definition slide")
            if section_types[2] != "example":
                raise GeminiResponseError("Slide 3 must be the photo-led example slide")
            if "origin" not in section_types[3:-1]:
                raise GeminiResponseError("Presentation is missing the required origin/creator section")
            if "interesting_facts" not in section_types[3:-1]:
                raise GeminiResponseError("Presentation is missing the required interesting-facts slide")
            if section_types[-1] != "conclusion_sources":
                raise GeminiResponseError("The final slide must contain the conclusion and sources")

            # Keep image queries only for the five high-value slides.
            priorities = {"cover": 0, "definition": 1, "origin": 2, "example": 3, "mechanics": 4}
            with_images = [(i, item) for i, item in enumerate(normalized) if item.get("image_query")]
            with_images.sort(key=lambda pair: (priorities.get(pair[1].get("section_type"), 10), pair[0]))
            allowed_img_indices = {i for i, _ in with_images[:MAX_DECK_IMAGES]}
            for i, item in enumerate(normalized):
                if i not in allowed_img_indices:
                    item["image_query"] = ""
                    item["image_kind"] = "none"
            source_lookup = {item["id"]: item for item in research_sources}
            for item in normalized:
                src_id = item.get("source", "").strip("[]")
                if src_id and src_id in source_lookup:
                    item["source_title"] = source_lookup[src_id]["title"]
                else:
                    item["source"] = ""
                    item["source_title"] = ""
            await _notify_generation_stage(progress_callback, "photos")
            return {
                "title": str(parsed.get("title") or topic).strip()[:180],
                "slides": normalized,
                "model": model,
                "visual_seed": visual_seed,
                "variation_id": variation_id,
                "visual_concept": visual_concept,
                "research_status": research_status,
                "detail": detail_key,
                "photo_subject": photo_subject,
                "photo_subject_kind": photo_subject_kind,
                "research_sources": research_sources,
                "reference_lines": _build_reference_lines(research_sources, [], data.get("language", "uk")),
            }
        except (GeminiRetryableError, GeminiModelUnavailable, GeminiResponseError) as exc:
            last_error = exc
            log.warning("Gemini model %s failed creating structured deck: %s", model, exc)
            if index < len(models) - 1:
                continue
        except GeminiAPIError:
            raise
    raise GeminiAPIError(f"Не вдалося створити презентацію. Остання помилка: {last_error}")

# 7. SLIDE LAYOUT ENGINE — композиції PowerPoint
# ============================================================
def rgb(value):
    return RGBColor(*value)
 
 
def add_shape(slide, kind, x, y, w, h, fill, line=None, radius=False):
    shape_type = MSO_SHAPE.ROUNDED_RECTANGLE if radius else kind
    shape = slide.shapes.add_shape(shape_type, Inches(x), Inches(y), Inches(w), Inches(h))
    shape.fill.solid()
    shape.fill.fore_color.rgb = rgb(fill)
    if line:
        shape.line.color.rgb = rgb(line)
        shape.line.width = Pt(1.2)
    else:
        shape.line.fill.background()
    return shape
 
 
def _estimate_wrapped_lines(text, chars_per_line):
    lines = 0
    for paragraph in str(text).splitlines() or [""]:
        wrapped = textwrap.wrap(
            paragraph, width=max(6, chars_per_line),
            break_long_words=True, break_on_hyphens=True,
        )
        lines += max(1, len(wrapped))
    return lines


def _fit_pptx_text(text, requested_size, width_in, height_in, min_size=8):
    """Estimate text fit from box geometry; shrink first, then trim only as a last resort."""
    size = float(requested_size)
    min_size = int(math.ceil(float(min_size)))
    available_width_pt = max(12.0, width_in * 72.0 - 7.0)
    available_height_pt = max(10.0, height_in * 72.0 - 4.0)
    for candidate in range(int(round(size)), min_size - 1, -1):
        chars_per_line = max(6, int(available_width_pt / max(4.0, candidate * 0.60)))
        max_lines = max(1, int(available_height_pt / max(8.0, candidate * 1.22)))
        if _estimate_wrapped_lines(text, chars_per_line) <= max_lines:
            return str(text), candidate
    candidate = min_size
    chars_per_line = max(6, int(available_width_pt / max(4.0, candidate * 0.60)))
    max_lines = max(1, int(available_height_pt / max(8.0, candidate * 1.22)))
    lines = []
    for paragraph in str(text).splitlines() or [""]:
        lines.extend(textwrap.wrap(paragraph, width=chars_per_line, break_long_words=True) or [""])
    if len(lines) > max_lines:
        lines = lines[:max_lines]
        lines[-1] = lines[-1].rstrip(" .,:;-") + "…"
    return "\n".join(lines), candidate


def add_text(slide, text, x, y, w, h, size, color, bold=False,
             align=PP_ALIGN.LEFT, font="Arial", valign=MSO_ANCHOR.TOP, min_size=8):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.margin_left = Inches(.04)
    tf.margin_right = Inches(.04)
    tf.margin_top = Inches(.02)
    tf.margin_bottom = Inches(.02)
    tf.vertical_anchor = valign
    # Long body copy must not collapse to tiny 8–9 pt text. At 90+ chars and a
    # body-sized requested font, retain at least 11 pt and trim only at word boundaries.
    effective_min_size = max(min_size, 11) if float(size) >= 12 and len(str(text)) >= 90 else min_size
    fitted_text, fitted_size = _fit_pptx_text(text, size, w, h, min_size=effective_min_size)
    p = tf.paragraphs[0]
    p.text = fitted_text
    p.alignment = align
    p.font.name = font or "Arial"
    p.font.size = Pt(fitted_size)
    p.font.bold = bold
    p.font.color.rgb = rgb(color)
    return box
 
 
def add_bullets(slide, points, x, y, w, h, theme, size=18, gap=10):
    points = points[:5]
    if not points:
        return
    total_h = h / len(points)
    longest = max(len(str(point)) for point in points)
    text_size = size - (4 if longest > 190 else 2 if longest > 125 or len(points) >= 5 else 0)
    text_size = max(11, text_size)
    for i, point in enumerate(points):
        yy = y + i * total_h
        dot_size = .10 if total_h < .62 else .12
        add_shape(slide, MSO_SHAPE.OVAL, x, yy + max(.07, (total_h - dot_size) / 2), dot_size, dot_size, theme["accent"])
        add_text(slide, point, x + .28, yy, w - .3, total_h - .02, text_size, theme["text"], min_size=11)
 
 
def add_outline(slide, kind, x, y, w, h, color, width=1.2):
    shape = slide.shapes.add_shape(kind, Inches(x), Inches(y), Inches(w), Inches(h))
    shape.fill.background()
    shape.line.color.rgb = rgb(color)
    shape.line.width = Pt(width)
    return shape


def add_panel(slide, design, x, y, w, h, theme, variant=0):
    """Style-aware content panel: shapes and edges change with the selected art direction."""
    if design in ("minimal", "classic", "magazine", "mono"):
        kind = MSO_SHAPE.RECTANGLE
        line = theme["muted"] if design == "minimal" else theme["accent"]
        shape = add_shape(slide, kind, x, y, w, h, theme["card"], line if design != "mono" else None)
    else:
        shape = add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h, theme["card"], theme["accent"] if design in ("future", "luxury") else None, True)
    if design in ("business", "magazine", "school", "nature"):
        bar_color = theme["accent"]
        add_shape(slide, MSO_SHAPE.RECTANGLE, x, y, .09 if design != "school" else .14, h, bar_color)
    elif design == "vibrant":
        palette = [(255, 91, 145), (0, 220, 210), (255, 197, 72)]
        add_shape(slide, MSO_SHAPE.RECTANGLE, x, y, w, .08, palette[variant % len(palette)])
    elif design == "mono":
        add_shape(slide, MSO_SHAPE.RECTANGLE, x, y, .12, h, theme["accent"])
    return shape


def add_style_chrome(slide, design, theme, index, total, rng, recipe=None):
    """Give each selected design a different visual identity, not just a new palette."""
    a, c, m = theme["accent"], theme["card"], theme["muted"]
    quiet_slide = bool(recipe and index not in (0, total - 1) and (index + 1) % recipe.get("quiet_every", 3) == 0)
    secondary = {
        "minimal": (120, 145, 166), "future": (107, 91, 255),
        "business": (52, 112, 170), "school": (255, 196, 72),
        "luxury": (146, 116, 65), "magazine": (40, 40, 40),
        "nature": (167, 193, 137), "vibrant": (0, 220, 210),
        "classic": (155, 173, 190), "mono": (145, 145, 145),
    }.get(design, m)
    if design == "minimal":
        add_shape(slide, MSO_SHAPE.RECTANGLE, .42, .55, .035, 6.25, a)
        add_shape(slide, MSO_SHAPE.RECTANGLE, 11.95, .55, .75, .035, m)
        add_text(slide, f"{index + 1:02d}", 11.95, 6.55, .75, .55, 25, a, True, PP_ALIGN.RIGHT, "Arial")
    elif design == "future":
        for gx in range(8):
            add_shape(slide, MSO_SHAPE.RECTANGLE, .35 + gx * 1.8, .35, .012, 6.8, (29, 49, 83))
        for gy in range(5):
            add_shape(slide, MSO_SHAPE.RECTANGLE, .35, .45 + gy * 1.55, 12.6, .012, (29, 49, 83))
        ring_x = 10.65 + rng.uniform(-.16, .16)
        ring_y = .62 + rng.uniform(-.08, .10)
        add_outline(slide, MSO_SHAPE.OVAL, ring_x, ring_y, 1.55, 1.55, a, 1.8)
        add_outline(slide, MSO_SHAPE.OVAL, ring_x + .37, ring_y + .37, .81, .81, secondary, 1.4)
        add_shape(slide, MSO_SHAPE.OVAL, ring_x + .72, ring_y + .72, .12, .12, a)
    elif design == "business":
        add_shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, .22, 7.5, a)
        add_shape(slide, MSO_SHAPE.RECTANGLE, .55, .38, 2.2, .09, a)
        add_shape(slide, MSO_SHAPE.RECTANGLE, 10.95, 6.95, 1.65, .12, secondary)
        chart_x = 10.15 + rng.uniform(-.18, .18)
        for bar_index in range(3):
            bar_h = rng.uniform(.18, .42)
            add_shape(slide, MSO_SHAPE.RECTANGLE, chart_x + bar_index * .24, 6.65 - bar_h, .12, bar_h, a)
        # The slide header owns the kicker; omit a second fixed label here.
    elif design == "school":
        circle_x = 10.8 + rng.uniform(-.16, .16)
        circle_y = .42 + rng.uniform(-.08, .12)
        add_shape(slide, MSO_SHAPE.OVAL, circle_x, circle_y, 1.45, 1.45, (255, 225, 169))
        add_shape(slide, MSO_SHAPE.OVAL, circle_x + .85, circle_y + .73, .72, .72, secondary)
        add_shape(slide, MSO_SHAPE.OVAL, .45, 6.25, .55, .55, (255, 196, 72))
        add_shape(slide, MSO_SHAPE.RECTANGLE, 1.1, 6.65, 2.0 + rng.uniform(-.25, .25), .08, a)
    elif design == "luxury":
        add_outline(slide, MSO_SHAPE.RECTANGLE, .38, .34, 12.55, 6.82, a, 1.0)
        add_outline(slide, MSO_SHAPE.RECTANGLE, .5, .46, 12.31, 6.58, secondary, .55)
        add_shape(slide, MSO_SHAPE.RECTANGLE, 5.75, .48, 1.75, .025, a)
        add_shape(slide, MSO_SHAPE.RECTANGLE, 5.75, 6.98, 1.75, .025, a)
    elif design == "magazine":
        add_shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, .28, 7.5, a)
        masthead_w = 1.25 + rng.uniform(0, .55)
        add_shape(slide, MSO_SHAPE.RECTANGLE, 10.95, .42, masthead_w, .15, a)
        add_text(slide, f"{index + 1:02d}", 10.82, .72 + rng.uniform(-.08, .12), 1.85, 1.25, 58, (224, 214, 202), True, PP_ALIGN.RIGHT, "Georgia")
        add_shape(slide, MSO_SHAPE.RECTANGLE, .62, 6.9, 12.05, .025, a)
    elif design == "nature":
        organic_x = 10.7 + rng.uniform(-.16, .16)
        add_shape(slide, MSO_SHAPE.OVAL, organic_x, .35, 1.6, 2.25, (218, 232, 207))
        leaf = add_shape(slide, MSO_SHAPE.OVAL, organic_x + .75, .75, .55, 1.35, a)
        leaf.rotation = 34 + rng.randint(-10, 10)
        leaf2 = add_shape(slide, MSO_SHAPE.OVAL, organic_x + .25, 1.1, .5, 1.05, secondary)
        leaf2.rotation = -32 + rng.randint(-9, 9)
        add_shape(slide, MSO_SHAPE.RECTANGLE, .62, 6.86, 2.8, .06, a)
    elif design == "vibrant":
        block = add_shape(slide, MSO_SHAPE.DIAMOND, 10.7 + rng.uniform(-.12, .12), .4, 1.55, 1.55, secondary)
        block.rotation = 12 + rng.randint(-12, 12)
        block2 = add_shape(slide, MSO_SHAPE.PARALLELOGRAM, .48, 6.15, 2.4, .75, (255, 91, 145))
        block2.rotation = -5 + rng.randint(-9, 9)
        add_shape(slide, MSO_SHAPE.RECTANGLE, 11.4, 6.75, 1.25 + rng.uniform(-.15, .2), .18, (255, 197, 72))
    elif design == "classic":
        add_shape(slide, MSO_SHAPE.RECTANGLE, .55, .42, 12.2, .035, a)
        add_shape(slide, MSO_SHAPE.RECTANGLE, .55, .52, 12.2, .018, m)
        add_shape(slide, MSO_SHAPE.RECTANGLE, .55, 6.92, 12.2, .025, m)
        add_text(slide, "◆", 6.15 + rng.uniform(-.25, .25), .35, .35, .3, 10, a, True, PP_ALIGN.CENTER, "Georgia")
    else:  # monochrome editorial / brutalist
        add_shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, 1.05, 7.5, (20, 20, 20))
        top_bar_w = 1.45 + rng.uniform(-.15, .35)
        add_shape(slide, MSO_SHAPE.RECTANGLE, 10.8, .38, top_bar_w, .24, (20, 20, 20))
        add_shape(slide, MSO_SHAPE.RECTANGLE, 10.8, .72, .75 + rng.uniform(0, .25), .08, (90, 90, 90))
        add_text(slide, f"{index + 1:02d}", 11.05, 6.55, 1.35, .62, 30, a, True, PP_ALIGN.RIGHT, "Arial")

    if recipe and not quiet_slide:
        decor_list = recipe.get("decorations") or ["side"]
        decor = decor_list[index % len(decor_list)]
        # Keep each page's ornament light; the deck's visual concept is also carried by content/kickers.
        if decor == "side":
            add_shape(slide, MSO_SHAPE.RECTANGLE, 12.72, 2.55, .075, 1.35 + rng.uniform(0, .65), a)
        elif decor == "top":
            add_shape(slide, MSO_SHAPE.RECTANGLE, 9.0, .18, 2.1 + rng.uniform(0, .9), .045, a)
        elif decor == "bottom":
            add_shape(slide, MSO_SHAPE.RECTANGLE, 7.4, 7.13, 3.0 + rng.uniform(0, .8), .04, a)
        elif decor == "circles":
            for j, d in enumerate((.18, .31, .46)):
                add_outline(slide, MSO_SHAPE.OVAL, 11.72 + j * .13, 5.72 + j * .09, d, d, a, .8)
        elif decor == "corner":
            mark = add_shape(slide, MSO_SHAPE.DIAMOND, 12.0, 5.75, .38, .38, a)
            mark.rotation = rng.randint(0, 45)
        elif decor == "blocks":
            for j, width in enumerate((.22, .36, .55)):
                add_shape(slide, MSO_SHAPE.RECTANGLE, 12.02, 5.72 + j * .16, width, .055, a)


def add_header(slide, index, total, theme, title, subtitle="", design="minimal", recipe=None, kicker="", language="uk"):
    font = "Georgia" if design in ("luxury", "magazine", "classic") else "Arial"
    title_size = {"minimal": 28, "business": 28, "classic": 29, "magazine": 32,
                  "luxury": 30, "vibrant": 33, "mono": 31}.get(design, 30)
    if len(title) > 62:
        title_size = max(20, title_size - 8)
    elif len(title) > 43:
        title_size = max(22, title_size - 5)
    fallback_labels = {
        "uk": {"minimal": "ІДЕЯ / ОРИГІНАЛ", "future": "ТЕХНОЛОГІЇ / ЛАБОРАТОРІЯ", "business": "СТРАТЕГІЧНИЙ БРИФ", "school": "НАВЧАННЯ / КРЕАТИВ", "luxury": "ПРЕМІАЛЬНЕ ВИДАННЯ", "magazine": "ОГЛЯД ЗНАНЬ", "nature": "ПРИРОДА / ПОГЛЯД", "vibrant": "КРЕАТИВНА КУЛЬТУРА", "classic": "АКАДЕМІЧНЕ ВИДАННЯ", "mono": "МОНОХРОМНІ НОТАТКИ"},
        "ru": {"minimal": "ИДЕЯ / ОРИГИНАЛ", "future": "ТЕХНОЛОГИИ / ЛАБОРАТОРИЯ", "business": "СТРАТЕГИЧЕСКИЙ БРИФ", "school": "ОБУЧЕНИЕ / КРЕАТИВ", "luxury": "ПРЕМИАЛЬНОЕ ИЗДАНИЕ", "magazine": "ОБЗОР ЗНАНИЙ", "nature": "ПРИРОДА / ВЗГЛЯД", "vibrant": "КРЕАТИВНАЯ КУЛЬТУРА", "classic": "АКАДЕМИЧЕСКОЕ ИЗДАНИЕ", "mono": "МОНОХРОМНЫЕ ЗАМЕТКИ"},
        "en": {"minimal": "IDEA / ORIGINAL", "future": "FUTURE SYSTEMS / LAB", "business": "STRATEGY / BRIEFING", "school": "LEARNING / CREATIVE", "luxury": "PREMIUM EDITION", "magazine": "THE KNOWLEDGE REVIEW", "nature": "NATURE / PERSPECTIVE", "vibrant": "CREATIVE CULTURE", "classic": "ACADEMIC EDITION", "mono": "MONO / NOTES"},
    }
    lang = language if language in fallback_labels else "uk"
    label = truncate_at_word(str(kicker or "").strip(), 48) or fallback_labels[lang].get(design, "PREZENT AI")
    subtitle = str(subtitle or "").strip()
    if design == "business" and subtitle.casefold() == label.casefold():
        subtitle = ""
    modes = (recipe or {}).get("headers", ["rule"])
    mode = modes[index % len(modes)] if modes else "rule"
    left_x = 1.40 if design == "mono" else .78
    title_y = .88
    title_w = 9.75 if design == "magazine" else 10.55

    if mode == "tag":
        tag_w = min(5.5, max(1.65, len(label) * .07))
        add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, left_x, .34, tag_w, .34, theme["card"], theme["accent"], True)
        add_text(slide, label, left_x + .1, .39, tag_w - .2, .22, 8, theme["accent"], True)
    elif mode == "numbered":
        add_text(slide, f"{index + 1:02d}", .58, .39, .7, .55, 22, theme["accent"], True, PP_ALIGN.CENTER)
        add_text(slide, label, 1.42, .47, 6.6, .22, 8, theme["accent"], True)
        left_x = 1.42
    elif mode == "sidebar":
        add_shape(slide, MSO_SHAPE.RECTANGLE, .58, .4, .07, 1.55, theme["accent"])
        add_text(slide, label, left_x, .46, 6.7, .22, 8, theme["accent"], True)
    else:
        add_text(slide, label, left_x, .46, 6.8, .22, 8, theme["accent"], True)

    add_text(slide, title, left_x, title_y, title_w, .66, title_size, theme["text"], True, font=font, min_size=18)
    subtitle_y = 1.65
    subtitle_h = .28
    separator_y = subtitle_y + subtitle_h + .16 if subtitle else 1.82
    if subtitle:
        add_text(slide, subtitle, left_x + .03, subtitle_y, title_w - .2, subtitle_h, 10.0, theme["muted"], False, font=font, min_size=8.5)
    if mode == "rule":
        add_shape(slide, MSO_SHAPE.RECTANGLE, left_x, separator_y, 1.0 + ((index + 1) % 4) * .27, .04, theme["accent"])
    elif mode == "tag":
        add_shape(slide, MSO_SHAPE.RECTANGLE, left_x, separator_y, 1.8 + ((index + 1) % 3) * .25, .025, theme["accent"])
    elif mode == "numbered":
        add_shape(slide, MSO_SHAPE.RECTANGLE, left_x, separator_y, .04, .18, theme["accent"])
    elif mode == "sidebar":
        add_shape(slide, MSO_SHAPE.RECTANGLE, left_x, separator_y, 2.2, .035, theme["accent"])
    add_text(slide, f"{index + 1:02d} / {total:02d}", 11.55, 6.93, .95, .25, 8, theme["muted"], True, PP_ALIGN.RIGHT)


def add_cover_slide(slide, item, theme, design, total, rng, recipe=None, language="uk"):
    title = item["title"]
    subtitle = item.get("subtitle", "")
    accent, text_color, muted = theme["accent"], theme["text"], theme["muted"]
    if design == "minimal":
        add_text(slide, deck_label(language, "cover_original"), .9, .8, 5.8, .3, 10, accent, True)
        add_text(slide, title, .9, 1.85, 7.0, 2.55, 43, text_color, True, font="Arial")
        if subtitle: add_text(slide, subtitle, .95, 4.55, 6.3, .8, 17, muted)
        add_shape(slide, MSO_SHAPE.RECTANGLE, 9.0, 1.6, .035, 4.2, accent)
        add_text(slide, "01", 9.4, 2.05, 2.3, 1.2, 72, accent, True, font="Arial")
        add_text(slide, deck_label(language, "clear_story"), 9.45, 3.6, 2.5, .9, 12, text_color, True)
    elif design == "future":
        add_text(slide, deck_label(language, "future_lab"), .9, .82, 6.8, .3, 10, accent, True)
        add_text(slide, title, .9, 1.7, 7.2, 2.65, 40, text_color, True, font="Arial")
        if subtitle: add_text(slide, subtitle, .95, 4.5, 6.2, .75, 16, muted)
        add_outline(slide, MSO_SHAPE.OVAL, 8.55, 1.35, 3.65, 3.65, accent, 2)
        add_outline(slide, MSO_SHAPE.OVAL, 9.15, 1.95, 2.45, 2.45, (107, 91, 255), 2)
        add_shape(slide, MSO_SHAPE.OVAL, 10.1, 2.9, .55, .55, accent)
        add_text(slide, "01  /  " + deck_label(language, "signal"), 8.95, 5.35, 3.4, .35, 12, accent, True, PP_ALIGN.CENTER)
    elif design == "business":
        add_shape(slide, MSO_SHAPE.RECTANGLE, 8.7, 0, 4.63, 7.5, (13, 25, 40))
        add_shape(slide, MSO_SHAPE.RECTANGLE, 8.7, 0, .15, 7.5, accent)
        add_text(slide, deck_label(language, "executive_brief"), .9, .95, 6.5, .3, 10, accent, True)
        add_text(slide, title, .9, 1.9, 7.0, 2.3, 39, text_color, True, font="Arial")
        if subtitle: add_text(slide, subtitle, .95, 4.45, 6.6, .8, 15, muted)
        for j, label_key in enumerate(("context", "insight", "action")):
            label = deck_label(language, label_key)
            yy = 1.9 + j * 1.25
            add_shape(slide, MSO_SHAPE.RECTANGLE, 9.35, yy, 2.85, .76, theme["card"])
            add_shape(slide, MSO_SHAPE.RECTANGLE, 9.35, yy, .08, .76, accent)
            add_text(slide, f"0{j+1}   {label}", 9.62, yy + .22, 2.3, .3, 13, text_color, True)
    elif design == "school":
        add_shape(slide, MSO_SHAPE.OVAL, 8.6, .85, 3.2, 3.2, (255, 224, 169))
        add_shape(slide, MSO_SHAPE.OVAL, 9.65, 1.85, 1.65, 1.65, accent)
        add_shape(slide, MSO_SHAPE.DIAMOND, 11.25, 4.35, 1.1, 1.1, (255, 196, 72))
        add_text(slide, deck_label(language, "lets_explore"), .9, .85, 5.7, .35, 11, accent, True)
        add_text(slide, title, .9, 1.65, 7.25, 2.4, 40, text_color, True, font="Arial")
        if subtitle: add_text(slide, subtitle, .95, 4.35, 6.7, .8, 16, muted)
        add_text(slide, deck_label(language, "discover_think_create"), .95, 6.15, 6.6, .3, 11, accent, True)
    elif design == "luxury":
        add_outline(slide, MSO_SHAPE.RECTANGLE, .75, .75, 11.85, 5.95, accent, 1.4)
        add_text(slide, deck_label(language, "brand"), 3.2, 1.35, 6.9, .4, 11, accent, True, PP_ALIGN.CENTER, "Georgia")
        add_text(slide, title, 1.5, 2.2, 10.3, 1.75, 40, text_color, False, PP_ALIGN.CENTER, "Georgia", MSO_ANCHOR.MIDDLE)
        if subtitle: add_text(slide, subtitle, 2.0, 4.35, 9.3, .75, 15, muted, False, PP_ALIGN.CENTER, "Georgia")
        add_shape(slide, MSO_SHAPE.RECTANGLE, 5.6, 5.55, 2.15, .025, accent)
        add_text(slide, deck_label(language, "original_edition"), 4.3, 5.82, 4.7, .3, 9, accent, True, PP_ALIGN.CENTER, "Georgia")
    elif design == "magazine":
        add_shape(slide, MSO_SHAPE.RECTANGLE, 8.55, 0, 4.78, 7.5, (232, 222, 208))
        add_text(slide, deck_label(language, "knowledge_review"), .82, .8, 7.3, .35, 9, accent, True)
        add_text(slide, title, .82, 1.55, 7.15, 2.85, 39, text_color, True, font="Georgia")
        if subtitle: add_text(slide, subtitle, .9, 4.65, 6.6, .72, 15, muted, False, font="Georgia")
        add_text(slide, "01", 9.0, 1.0, 3.55, 2.7, 104, accent, True, PP_ALIGN.CENTER, "Georgia")
        add_text(slide, deck_label(language, "feature_story"), 9.45, 4.4, 2.7, 1.0, 24, text_color, True, PP_ALIGN.CENTER, "Georgia")
    elif design == "nature":
        add_shape(slide, MSO_SHAPE.OVAL, 8.5, .6, 3.75, 5.0, (218, 232, 207))
        leaf1 = add_shape(slide, MSO_SHAPE.OVAL, 9.25, 1.2, 1.0, 3.6, accent); leaf1.rotation = 35
        leaf2 = add_shape(slide, MSO_SHAPE.OVAL, 10.5, 1.4, .9, 3.4, (120, 157, 104)); leaf2.rotation = -35
        add_text(slide, deck_label(language, "field_notes"), .9, .85, 7.0, .3, 10, accent, True)
        add_text(slide, title, .9, 1.75, 7.1, 2.3, 40, text_color, True, font="Georgia")
        if subtitle: add_text(slide, subtitle, .95, 4.35, 6.5, .85, 15, muted)
        add_text(slide, deck_label(language, "observe_understand_restore"), .95, 6.15, 6.7, .3, 10, accent, True)
    elif design == "vibrant":
        add_shape(slide, MSO_SHAPE.RECTANGLE, 8.6, 0, 4.73, 7.5, (53, 38, 92))
        d = add_shape(slide, MSO_SHAPE.DIAMOND, 9.1, .85, 2.5, 2.5, (0, 220, 210)); d.rotation = 14
        add_shape(slide, MSO_SHAPE.OVAL, 10.4, 3.0, 1.75, 1.75, (255, 91, 145))
        add_shape(slide, MSO_SHAPE.RECTANGLE, 9.05, 5.65, 3.3, .18, (255, 197, 72))
        add_text(slide, deck_label(language, "ideas_in_motion"), .9, .85, 6.8, .35, 11, (0, 220, 210), True)
        add_text(slide, title, .9, 1.7, 7.1, 2.65, 42, text_color, True, font="Arial")
        if subtitle: add_text(slide, subtitle, .95, 4.65, 6.5, .75, 16, muted)
    elif design == "classic":
        add_outline(slide, MSO_SHAPE.RECTANGLE, 1.05, .78, 11.2, 5.85, accent, 1.3)
        add_text(slide, deck_label(language, "academic_presentation"), 2.0, 1.25, 9.3, .3, 10, accent, True, PP_ALIGN.CENTER, "Georgia")
        add_text(slide, title, 1.7, 2.2, 9.9, 1.85, 37, text_color, True, PP_ALIGN.CENTER, "Georgia", MSO_ANCHOR.MIDDLE)
        if subtitle: add_text(slide, subtitle, 2.05, 4.25, 9.2, .8, 15, muted, False, PP_ALIGN.CENTER, "Georgia")
        add_shape(slide, MSO_SHAPE.RECTANGLE, 5.8, 5.55, 1.75, .025, accent)
    else:  # hard-edge monochrome / brutalist cover
        add_shape(slide, MSO_SHAPE.RECTANGLE, .72, .7, 7.3, 5.95, (22, 22, 22))
        add_text(slide, deck_label(language, "mono_edition"), 1.1, 1.15, 6.3, .3, 10, (220, 220, 220), True)
        add_text(slide, title, 1.08, 2.0, 6.25, 2.45, 39, (255, 255, 255), True, font="Arial")
        if subtitle: add_text(slide, subtitle, 1.15, 4.85, 5.9, .8, 15, (208, 208, 208))
        add_text(slide, "01", 8.75, 1.45, 3.2, 2.4, 105, (24, 24, 24), True, PP_ALIGN.CENTER, "Arial")
        add_shape(slide, MSO_SHAPE.RECTANGLE, 8.95, 4.25, 2.8, .16, (24, 24, 24))
        add_text(slide, deck_label(language, "idea_form_impact"), 8.95, 4.65, 3.0, .35, 11, text_color, True)

    if recipe:
        shape_a, shape_b = recipe.get("cover_shapes", COVER_SHAPES[0])
        mark_x = 11.55 + rng.uniform(-.12, .12)
        mark_y = 5.58 + rng.uniform(-.12, .12)
        mark = add_shape(slide, shape_a, mark_x, mark_y, .34, .34, accent)
        mark.rotation = rng.randint(0, 35)
        add_outline(slide, shape_b, mark_x + .43, mark_y + .04, .24, .24, accent, .8)


def add_fade_transition(slide):
    transition = OxmlElement("p:transition")
    transition.set("spd", "med")
    transition.append(OxmlElement("p:fade"))
    slide._element.insert(2, transition)



def add_picture_cropped(slide, image_path: str, x: float, y: float, w: float, h: float):
    """Add a centered, aspect-preserving cropped picture to a PowerPoint slide."""
    from PIL import Image
    with Image.open(image_path) as im:
        iw, ih = im.size
    picture = slide.shapes.add_picture(image_path, Inches(x), Inches(y), width=Inches(w), height=Inches(h))
    source_ratio = iw / max(1, ih)
    target_ratio = w / max(.001, h)
    if source_ratio > target_ratio:
        visible = target_ratio / source_ratio
        crop = max(0.0, (1.0 - visible) / 2.0)
        picture.crop_left = crop
        picture.crop_right = crop
    elif source_ratio < target_ratio:
        visible = source_ratio / target_ratio
        crop = max(0.0, (1.0 - visible) / 2.0)
        picture.crop_top = crop
        picture.crop_bottom = crop
    return picture


DECK_LABELS = {
    "uk": {
        "source": "Джерело", "unknown_origin": "Походження не підтверджене", "key_idea": "КЛЮЧОВА ДУМКА", "explanation": "ПОЯСНЕННЯ", "takeaway": "ГОЛОВНИЙ ВИСНОВОК", "section": "РОЗДІЛ", "feature": "ОГЛЯД", "fact": "Факт", "sources": "ДЖЕРЕЛА", "insight": "ВИСНОВОК", "chapter": "РОЗДІЛ", "author_unknown": "Автора не зазначено", "photo_credit": "Фото: авторство", "feature_story": "ГОЛОВНА ТЕМА", "context": "КОНТЕКСТ", "action": "ДІЯ",
        "cover_original": "PREZENT AI / ОРИГІНАЛЬНА ПРЕЗЕНТАЦІЯ", "clear_story": "ЧІТКА ІДЕЯ.\nСИЛЬНА РОЗПОВІДЬ.", "future_lab": "PREZENT AI / ЛАБОРАТОРІЯ МАЙБУТНЬОГО", "signal": "СИГНАЛ", "executive_brief": "PREZENT AI / ДІЛОВИЙ БРИФ", "lets_explore": "ДОСЛІДЖУЙМО", "discover_think_create": "ВІДКРИВАЙ • ДУМАЙ • СТВОРЮЙ", "brand": "P R E Z E N T   A I", "original_edition": "ОРИГІНАЛЬНЕ ВИДАННЯ", "knowledge_review": "ОГЛЯД ЗНАНЬ / СПЕЦІАЛЬНИЙ ВИПУСК", "field_notes": "ПОЛЬОВІ НОТАТКИ / ЕКОПОГЛЯД", "observe_understand_restore": "СПОСТЕРІГАЙ • РОЗУМІЙ • ВІДНОВЛЮЙ", "ideas_in_motion": "ІДЕЇ В РУСІ", "academic_presentation": "PREZENT AI • АКАДЕМІЧНА ПРЕЗЕНТАЦІЯ", "mono_edition": "PREZENT AI / МОНОХРОМНЕ ВИДАННЯ", "idea_form_impact": "ІДЕЯ / ФОРМА / ВПЛИВ",
    },
    "ru": {
        "source": "Источник", "unknown_origin": "Происхождение не подтверждено", "key_idea": "КЛЮЧЕВАЯ МЫСЛЬ", "explanation": "ПОЯСНЕНИЕ", "takeaway": "ГЛАВНЫЙ ВЫВОД", "section": "РАЗДЕЛ", "feature": "ОБЗОР", "fact": "Факт", "sources": "ИСТОЧНИКИ", "insight": "ВЫВОД", "chapter": "ГЛАВА", "author_unknown": "Автор не указан", "photo_credit": "Фото: авторство", "feature_story": "ГЛАВНАЯ ТЕМА", "context": "КОНТЕКСТ", "action": "ДЕЙСТВИЕ",
        "cover_original": "PREZENT AI / ОРИГИНАЛЬНАЯ ПРЕЗЕНТАЦИЯ", "clear_story": "ЧЁТКАЯ ИДЕЯ.\nСИЛЬНЫЙ РАССКАЗ.", "future_lab": "PREZENT AI / ЛАБОРАТОРИЯ БУДУЩЕГО", "signal": "СИГНАЛ", "executive_brief": "PREZENT AI / ДЕЛОВОЙ БРИФ", "lets_explore": "ДАВАЙТЕ ИССЛЕДОВАТЬ", "discover_think_create": "ОТКРЫВАЙ • ДУМАЙ • СОЗДАВАЙ", "brand": "P R E Z E N T   A I", "original_edition": "ОРИГИНАЛЬНОЕ ИЗДАНИЕ", "knowledge_review": "ОБЗОР ЗНАНИЙ / СПЕЦИАЛЬНЫЙ ВЫПУСК", "field_notes": "ПОЛЕВЫЕ ЗАМЕТКИ / ЭКОПЕРСПЕКТИВА", "observe_understand_restore": "НАБЛЮДАЙ • ПОНИМАЙ • ВОССТАНАВЛИВАЙ", "ideas_in_motion": "ИДЕИ В ДВИЖЕНИИ", "academic_presentation": "PREZENT AI • АКАДЕМИЧЕСКАЯ ПРЕЗЕНТАЦИЯ", "mono_edition": "PREZENT AI / МОНОХРОМНОЕ ИЗДАНИЕ", "idea_form_impact": "ИДЕЯ / ФОРМА / ВЛИЯНИЕ",
    },
    "en": {
        "source": "Source", "unknown_origin": "Origin not confirmed", "key_idea": "KEY IDEA", "explanation": "EXPLANATION", "takeaway": "MAIN TAKEAWAY", "section": "SECTION", "feature": "OVERVIEW", "fact": "Fact", "sources": "SOURCES", "insight": "INSIGHT", "chapter": "CHAPTER", "author_unknown": "Author not listed", "photo_credit": "Photo credit", "feature_story": "FEATURE STORY", "context": "CONTEXT", "action": "ACTION",
        "cover_original": "PREZENT AI / ORIGINAL PRESENTATION", "clear_story": "A CLEAR IDEA.\nA STRONG STORY.", "future_lab": "PREZENT AI / FUTURE LAB", "signal": "SIGNAL", "executive_brief": "PREZENT AI / EXECUTIVE BRIEF", "lets_explore": "LET'S EXPLORE", "discover_think_create": "DISCOVER • THINK • CREATE", "brand": "P R E Z E N T   A I", "original_edition": "AN ORIGINAL EDITION", "knowledge_review": "THE KNOWLEDGE REVIEW / SPECIAL ISSUE", "field_notes": "FIELD NOTES / ECO PERSPECTIVE", "observe_understand_restore": "OBSERVE • UNDERSTAND • RESTORE", "ideas_in_motion": "IDEAS IN MOTION", "academic_presentation": "PREZENT AI • ACADEMIC PRESENTATION", "mono_edition": "PREZENT AI / MONO EDITION", "idea_form_impact": "IDEA / FORM / IMPACT",
    },
}


def deck_label(language: str, key: str) -> str:
    lang = language if language in DECK_LABELS else "uk"
    return DECK_LABELS[lang].get(key, key)


def detail_label_for(language: str, detail: str) -> str:
    labels = {"uk": {"short": "Коротко", "medium": "Середньо", "detailed": "Детально"},
              "ru": {"short": "Кратко", "medium": "Средне", "detailed": "Подробно"},
              "en": {"short": "Short", "medium": "Medium", "detailed": "Detailed"}}
    lang = language if language in labels else "uk"
    return labels[lang].get(detail, labels[lang]["medium"])


def compact_photo_credit(value: str, author_limit: int = 34, language: str = "uk") -> str:
    """Shorten only the author name while preserving the complete license and provider."""
    text = re.sub(r"\s+", " ", str(value or "")).strip()
    parts = [part.strip() for part in re.split(r"\s+[—–]\s+", text) if part.strip()]
    if len(parts) >= 3:
        author, license_name = parts[0], parts[1]
        if author.lower() == "author not listed":
            author = deck_label(language, "author_unknown")
        if len(author) > author_limit:
            author = truncate_at_word(author, author_limit)
        return f"{author} — {license_name} — {' — '.join(parts[2:])}"
    return truncate_at_word(text, 130)


def compact_reference_line(line: str) -> str:
    raw = str(line or "").strip()
    match = re.match(r"^(\[[^]]+\])\s*(.*?)\s+[—–]\s+(https?://\S+)$", raw)
    if match:
        source_id, title, url = match.groups()
        host = urlparse(url).netloc.removeprefix("www.")
        return f"{source_id} {truncate_at_word(title, 54)} ({host})"
    # Photo references are different: preserve photographer, full license name and provider;
    # shorten only the photographer name, then show the Commons/Pexels page host.
    photo_match = re.match(r"^(.+?\s+\d+:)\s*(.*?)\s+[—–]\s+(https?://\S+)(.*)$", raw)
    if photo_match:
        prefix, credit, url, suffix = photo_match.groups()
        parts = [part.strip() for part in re.split(r"\s+[—–]\s+", credit) if part.strip()]
        if len(parts) >= 3:
            author = truncate_at_word(parts[0], 26) if len(parts[0]) > 26 else parts[0]
            credit = f"{author} — {parts[1]} — {' — '.join(parts[2:])}"
        host = urlparse(url).netloc.removeprefix("www.")
        license_url = re.search(r"\((?:license|ліцензія|лицензия):\s*(https?://[^)]+)\)", suffix, re.IGNORECASE)
        suffix_text = f" · {license_url.group(1)}" if license_url else ""
        return f"{prefix} {credit} ({host}){suffix_text}"
    return truncate_at_word(raw, 120)


LAYOUT_POINT_CAPACITY = {
    # Capacities match the number of distinct points visibly rendered by BOTH exporters.
    # Any slide with five ordinary points is routed to the full-width five-row steps layout.
    "cover": 0, "section": 0,
    "stat": 5, "two_col": 4, "steps": 5, "comparison": 5,
    "photo_text": 4, "portrait_fact": 4, "cards": 4, "timeline": 4,
    "quote": 4, "focus": 4, "closing": 4, "split": 4,
    "grid3": 3,
}


def resolve_render_layout(item: dict) -> str:
    """Choose a renderer that can show every point without tiny text or blank cards."""
    layout = str(item.get("layout") or "two_col").strip().lower()
    points = item.get("content") or []
    if isinstance(points, str):
        points = [points]
    points = [str(point) for point in points if str(point).strip()]
    if layout not in LAYOUT_POINT_CAPACITY:
        layout = "two_col"

    # Cover/section layouts are visual-only; keep their factual points in a real content layout.
    if layout in ("cover", "section") and points:
        return "two_col" if len(points) <= 4 else "steps"

    # grid3 is only used when exactly three cards are available. Two points use a split;
    # four/five points use steps instead of losing content or drawing an empty third card.
    if layout == "grid3":
        if len(points) == 3:
            return "grid3"
        if len(points) <= 2:
            return "two_col"
        return "steps"

    # Dense detailed bullets need the wide, full-height five-row renderer; do not shrink
    # four or five long points into narrow cards/columns where they become unreadable.
    dense_layouts = {
        "grid3", "two_col", "photo_text", "portrait_fact", "cards",
        "timeline", "comparison", "quote", "focus", "closing", "split",
    }
    # Five points need the full-width row layout even if each point is short.
    # Long points also move out of cards/columns before their text starts to clip.
    if len(points) >= 5 and layout != "steps":
        return "steps"
    if layout in dense_layouts and len(points) >= 3 and any(len(point) > 120 for point in points):
        return "steps"

    # A two-column slide has one lead point + up to three readable supporting points.
    if layout == "two_col" and len(points) > LAYOUT_POINT_CAPACITY["two_col"]:
        return "steps"

    # Closing slide has a dedicated source area and therefore a four-point content capacity.
    # Route excess points to steps so none of the user's content disappears. Sources are
    # retained in slide notes and per-slide source footers.
    if layout == "closing" and len(points) > LAYOUT_POINT_CAPACITY["closing"]:
        return "steps"
    if layout == "closing":
        return "closing"

    capacity = LAYOUT_POINT_CAPACITY[layout]
    if capacity and len(points) > capacity:
        return "steps"
    return layout


def layout_point_capacity(item: dict) -> int:
    """Return the actual display capacity after resolving the slide's renderer."""
    resolved = resolve_render_layout(item)
    return int(LAYOUT_POINT_CAPACITY.get(resolved, 5))


def _supporting_points(item: dict, maximum: int = 4) -> list[str]:
    """Return all visible support points without duplicating the separate highlight."""
    points = [str(value) for value in (item.get("content") or []) if str(value).strip()]
    highlight = str(item.get("highlight") or "").strip()
    if highlight and points and highlight == points[0]:
        points = points[1:]
    return points[:maximum]

def _append_pptx_fact_source(slide, item, design, theme):
    fact = str(item.get("fact") or "").strip()
    source = str(item.get("source") or "").strip()
    if not (fact or source):
        return
    lang = item.get("language", "uk")
    label = fact[:125]
    if source:
        label = (label + "  ·  " if label else "") + f"{deck_label(lang, 'source')} {source}"
    add_text(slide, label, .88, 6.56, 10.55, .38, 10.0, theme["muted"], False, font="Arial", min_size=9.0)


def create_pptx(data: dict, design: str, output_path: str):
    theme = copy.deepcopy(DESIGNS[design])
    visual_seed = int(data.get("visual_seed") or random.randint(1, 2_000_000_000))
    recipe = make_deck_recipe(design, visual_seed, len(data["slides"]), data.get("visual_concept", ""))
    theme["accent"] = recipe["accent"]
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    total = len(data["slides"])
    rng = random.Random(int(data.get("visual_seed") or random.randint(1, 2_000_000_000)))
    for index, item in enumerate(data["slides"]):
        item.setdefault("language", data.get("language", "uk"))
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = rgb(theme["bg"])
        layout = resolve_render_layout(item)
        add_style_chrome(slide, design, theme, index, total, rng, recipe)
        if layout == "cover":
            add_cover_slide(slide, item, theme, design, total, rng, recipe, data.get("language", "uk"))
            if item.get("image_path") and Path(item["image_path"]).is_file():
                add_panel(slide, design, 8.05, 1.38, 4.22, 4.88, theme, index)
                add_picture_cropped(slide, item["image_path"], 8.20, 1.53, 3.92, 4.58)
                if item.get("image_credit"):
                    add_text(slide, compact_photo_credit(item["image_credit"], author_limit=22, language=data.get("language", "uk")), 8.22, 6.10, 3.9, .34, 8.0, theme["muted"], False, min_size=7.5)
        elif layout == "closing":
            if design in ("luxury", "classic"):
                add_outline(slide, MSO_SHAPE.RECTANGLE, .8, .75, 11.75, 5.95, theme["accent"], 1.2)
                add_text(slide, deck_label(data.get("language", "uk"), "takeaway"), 2.0, 1.15, 9.3, .3, 10, theme["accent"], True, PP_ALIGN.CENTER, "Georgia")
                add_text(slide, item["title"], 1.45, 1.85, 10.45, 1.0, 32, theme["text"], True, PP_ALIGN.CENTER, "Georgia")
                add_text(slide, "\n".join(item["content"][:4]), 2.0, 3.15, 9.3, 1.72 if data.get("reference_lines") else 2.0, 16 if data.get("reference_lines") else 17, theme["text"], False, PP_ALIGN.CENTER, "Georgia")
                if item.get("highlight") and not data.get("reference_lines"):
                    add_text(slide, item["highlight"], 2.0, 5.55, 9.3, .55, 15, theme["accent"], True, PP_ALIGN.CENTER, "Georgia")
            else:
                add_header(slide, index, total, theme, item["title"], item.get("subtitle", ""), design, recipe, item.get("kicker", ""), data.get("language", "uk"))
                add_panel(slide, design, .82, 2.35, 11.55, 3.75, theme, index)
                refs = data.get("reference_lines") or []
                add_bullets(slide, item["content"][:4], 1.25, 2.72, 10.65, 1.82 if refs else 2.8, theme, 15)
                if item.get("highlight") and not refs:
                    add_text(slide, item["highlight"], 1.1, 6.2, 10.9, .45, 14, theme["accent"], True, PP_ALIGN.CENTER)
            refs = data.get("reference_lines") or []
            if refs:
                display_refs = [compact_reference_line(line) for line in refs[:8]]
                add_text(slide, deck_label(data.get("language", "uk"), "sources"), 1.15, 4.55, 10.9, .24, 10, theme["accent"], True)
                split_at = (len(display_refs) + 1) // 2
                add_text(slide, "\n".join(display_refs[:split_at]), 1.18, 4.82, 5.12, 1.42, 11.0, theme["muted"], False, min_size=10.5)
                if display_refs[split_at:]:
                    add_text(slide, "\n".join(display_refs[split_at:]), 6.55, 4.82, 5.25, 1.42, 11.0, theme["muted"], False, min_size=10.5)
        elif layout == "section":
            # Chapter divider: intentionally no standard header/list panel.
            add_shape(slide, MSO_SHAPE.RECTANGLE, .78, .85, .13, 5.8, theme["accent"])
            add_text(slide, item.get("kicker") or deck_label(data.get("language", "uk"), "section"), 1.25, 1.35, 7.4, .32, 10, theme["accent"], True)
            add_text(slide, item["title"], 1.22, 2.0, 10.6, 1.7, 37, theme["text"], True, font="Georgia" if design in ("luxury", "magazine", "classic") else "Arial", valign=MSO_ANCHOR.MIDDLE)
            if item.get("subtitle"):
                add_text(slide, item["subtitle"], 1.28, 4.0, 9.2, .72, 17, theme["muted"], False)
            add_shape(slide, MSO_SHAPE.RECTANGLE, 1.25, 5.45, 2.0 + (index % 3) * .65, .055, theme["accent"])
        elif layout == "stat":
            add_header(slide, index, total, theme, item["title"], item.get("subtitle", ""), design, recipe, item.get("kicker", ""), data.get("language", "uk"))
            value = item.get("stat_value") or item.get("highlight") or (item["content"][0] if item.get("content") else "")
            label = item.get("stat_label") or (item["content"][1] if len(item.get("content", [])) > 1 else "")
            add_panel(slide, design, .95, 2.35, 11.35, 3.85, theme, index + recipe.get("panel_shift", 0))
            add_text(slide, value, 1.35, 2.72, 5.05, 2.35, 43 if len(value) < 18 else 30, theme["accent"], True, PP_ALIGN.CENTER, "Georgia" if design in ("luxury", "magazine", "classic") else "Arial", MSO_ANCHOR.MIDDLE)
            if label:
                add_text(slide, label, 6.7, 2.85, 4.85, 1.1, 20, theme["text"], True, valign=MSO_ANCHOR.MIDDLE)
            extra_stat_points = item.get("content", [])[2:5]
            if extra_stat_points:
                add_bullets(slide, extra_stat_points, 6.72, 4.12, 4.8, 1.55, theme, 13)
            elif item.get("highlight") and value != item.get("highlight"):
                add_text(slide, item["highlight"], 6.72, 4.18, 4.8, .9, 13, theme["muted"], False)
        elif layout == "photo_text":
            add_header(slide, index, total, theme, item["title"], item.get("subtitle", ""), design, recipe, item.get("kicker", ""), data.get("language", "uk"))
            points = item.get("content", [])[:4]
            if item.get("image_path") and Path(item["image_path"]).is_file():
                add_panel(slide, design, 7.25, 2.25, 5.05, 3.95, theme, index)
                add_picture_cropped(slide, item["image_path"], 7.42, 2.42, 4.72, 3.58)
                if item.get("image_credit"):
                    add_text(slide, compact_photo_credit(item["image_credit"], language=data.get("language", "uk")), 7.45, 6.05, 4.65, .22, 7, theme["muted"], False)
                add_bullets(slide, points, .95, 2.35, 5.8, 3.85, theme, 16)
            else:
                add_panel(slide, design, .88, 2.32, 11.55, 3.9, theme, index)
                add_bullets(slide, points, 1.25, 2.68, 10.7, 3.15, theme, 17)
        elif layout == "portrait_fact":
            add_header(slide, index, total, theme, item["title"], item.get("subtitle", ""), design, recipe, item.get("kicker", ""), data.get("language", "uk"))
            points = item.get("content", [])[:4]
            if item.get("image_path") and Path(item["image_path"]).is_file():
                add_panel(slide, design, 8.05, 2.18, 4.18, 4.15, theme, index)
                add_picture_cropped(slide, item["image_path"], 8.22, 2.35, 3.84, 3.45)
                if item.get("image_credit"):
                    add_text(slide, compact_photo_credit(item["image_credit"], language=data.get("language", "uk")), 8.24, 5.86, 3.78, .22, 6.8, theme["muted"], False)
                add_text(slide, item.get("fact") or deck_label(data.get("language", "uk"), "unknown_origin"), 1.05, 2.42, 6.45, 1.45, 25, theme["accent"], True, font="Georgia" if design in ("luxury", "magazine", "classic") else "Arial")
                add_bullets(slide, points, 1.08, 4.12, 6.35, 1.95, theme, 15)
            else:
                add_text(slide, item.get("fact") or deck_label(data.get("language", "uk"), "unknown_origin"), 1.0, 2.35, 10.9, 1.0, 26, theme["accent"], True, PP_ALIGN.CENTER)
                add_bullets(slide, points, 1.45, 3.65, 10.3, 2.0, theme, 15)
        elif layout == "two_col":
            add_header(slide, index, total, theme, item["title"], item.get("subtitle", ""), design, recipe, item.get("kicker", ""), data.get("language", "uk"))
            pts = item.get("content", [])[:4]
            left_lead = pts[0] if pts else item.get("highlight", "")
            right_points = pts[1:] or ([item.get("highlight")] if item.get("highlight") else [])
            add_shape(slide, MSO_SHAPE.RECTANGLE, .9, 2.35, 5.05, 3.95, theme["card"], theme["accent"] if design in ("future", "luxury") else None)
            add_text(slide, item.get("left_title") or deck_label(data.get("language", "uk"), "key_idea"), 1.22, 2.65, 4.4, .3, 10, theme["accent"], True)
            add_text(slide, left_lead, 1.22, 3.15, 4.35, 2.65, 21, theme["text"], True, font="Georgia" if design in ("luxury", "magazine", "classic") else "Arial", valign=MSO_ANCHOR.MIDDLE, min_size=11)
            add_shape(slide, MSO_SHAPE.RECTANGLE, 6.45, 2.45, .045, 3.65, theme["accent"])
            add_text(slide, item.get("right_title") or deck_label(data.get("language", "uk"), "explanation"), 6.82, 2.55, 5.15, .3, 10, theme["accent"], True)
            add_bullets(slide, right_points, 6.82, 3.05, 5.05, 3.0, theme, 16)
        elif layout == "grid3":
            add_header(slide, index, total, theme, item["title"], item.get("subtitle", ""), design, recipe, item.get("kicker", ""), data.get("language", "uk"))
            pts = item.get("content", [])[:3]
            if len(pts) != 3:
                raise ValueError("grid3 renderer requires exactly three content points; resolve_render_layout should route otherwise")
            widths = [3.55, 3.55, 3.55]
            xs = [.88, 4.88, 8.88]
            for j, point in enumerate(pts):
                x, w = xs[j], widths[j]
                panel = add_panel(slide, design, x, 2.55, w, 3.45, theme, index + j + recipe.get("panel_shift", 0))
                add_text(slide, f"0{j + 1}", x + .24, 2.82, .72, .62, 26, theme["accent"], True, PP_ALIGN.LEFT, "Georgia" if design in ("luxury", "magazine", "classic") else "Arial")
                add_shape(slide, MSO_SHAPE.RECTANGLE, x + .25, 3.62, w - .5, .04, theme["accent"])
                add_text(slide, point, x + .26, 3.9, w - .52, 1.72, 15, theme["text"], True, valign=MSO_ANCHOR.MIDDLE)
        elif layout == "cards":
            add_header(slide, index, total, theme, item["title"], item.get("subtitle", ""), design, recipe, item.get("kicker", ""), data.get("language", "uk"))
            pts = item["content"][:4]
            serif = design in ("luxury", "magazine", "classic")
            if design == "minimal":
                # Swiss editorial: open numbered rows, rules and generous whitespace.
                row_h = 3.85 / max(1, len(pts))
                for j, point in enumerate(pts):
                    y = 2.28 + j * row_h
                    add_text(slide, f"{j + 1:02d}", .92, y + .05, .85, .55, 25, theme["accent"], True, font="Arial")
                    add_shape(slide, MSO_SHAPE.RECTANGLE, 1.95, y + .02, .025, row_h - .16, theme["muted"])
                    add_text(slide, point, 2.25, y + .04, 9.1, row_h - .14, 17, theme["text"], False)
                    if j < len(pts) - 1:
                        add_shape(slide, MSO_SHAPE.RECTANGLE, 2.25, y + row_h - .08, 9.7, .012, theme["muted"])
            elif design == "magazine":
                # Magazine feature: one dominant story block and a narrow editorial rail.
                if pts:
                    add_panel(slide, design, .82, 2.24, 6.5, 4.05, theme, index)
                    add_text(slide, f"{deck_label(data.get('language', 'uk'), 'feature')}  /  01", 1.2, 2.55, 4.8, .28, 10, theme["accent"], True, font="Georgia")
                    add_text(slide, pts[0], 1.2, 3.05, 5.6, 2.75, 23, theme["text"], True, font="Georgia")
                side = pts[1:] or []
                slot_h = 3.85 / max(1, len(side))
                for j, point in enumerate(side):
                    y = 2.28 + j * slot_h
                    add_text(slide, f"0{j + 2}", 7.72, y + .05, .72, .5, 22, theme["accent"], True, PP_ALIGN.CENTER, "Georgia")
                    add_shape(slide, MSO_SHAPE.RECTANGLE, 8.55, y + .02, .045, slot_h - .12, theme["accent"])
                    add_text(slide, point, 8.85, y + .06, 3.45, slot_h - .16, 15, theme["text"], False, font="Georgia")
            elif design == "mono":
                # Brutalist mono: hard black number blocks and stark horizontal bands.
                row_h = 3.82 / max(1, len(pts))
                for j, point in enumerate(pts):
                    y = 2.3 + j * row_h
                    add_shape(slide, MSO_SHAPE.RECTANGLE, .95, y, 1.05, row_h - .1, (22, 22, 22))
                    add_text(slide, f"{j + 1:02d}", 1.08, y + .14, .78, .45, 22, (255, 255, 255), True, PP_ALIGN.CENTER, "Arial")
                    add_shape(slide, MSO_SHAPE.RECTANGLE, 2.18, y, 9.75, row_h - .1, (255, 255, 255), theme["accent"])
                    add_text(slide, point, 2.48, y + .12, 9.15, row_h - .28, 16, theme["text"], True, font="Arial")
            elif design == "nature":
                # Nature notes: leaf markers on a calm organic vertical rhythm.
                row_h = 3.78 / max(1, len(pts))
                for j, point in enumerate(pts):
                    y = 2.32 + j * row_h
                    leaf = add_shape(slide, MSO_SHAPE.OVAL, 1.0 + (j % 2) * .12, y + .14, .35, .72, theme["accent"] if j % 2 == 0 else (120, 157, 104))
                    leaf.rotation = 25 if j % 2 == 0 else -25
                    add_panel(slide, design, 1.68, y, 10.2, row_h - .13, theme, j + index)
                    add_text(slide, point, 2.02, y + .12, 9.45, row_h - .35, 16, theme["text"], False)
            elif design == "business":
                # Executive dashboard: labelled insight tiles in a strict grid.
                positions = [(.88, 2.35), (6.78, 2.35), (.88, 4.42), (6.78, 4.42)]
                for j, point in enumerate(pts):
                    x, y = positions[j]
                    add_panel(slide, design, x, y, 5.55, 1.72, theme, j + index)
                    add_text(slide, f"{deck_label(data.get('language', 'uk'), 'insight')}  0{j + 1}", x + .28, y + .18, 4.7, .25, 9, theme["accent"], True)
                    add_text(slide, point, x + .3, y + .57, 4.88, 1.0, 15, theme["text"], True)
            elif design == "school":
                # Learning board: friendly staggered note cards and playful markers.
                positions = [(.95, 2.35), (6.75, 2.52), (1.25, 4.45), (6.45, 4.3)]
                for j, point in enumerate(pts):
                    x, y = positions[j]
                    add_panel(slide, design, x, y, 5.35, 1.62, theme, j + index)
                    add_shape(slide, MSO_SHAPE.OVAL, x + .22, y + .2, .32, .32, (255, 196, 72))
                    add_text(slide, f"{j + 1}", x + .25, y + .25, .25, .2, 9, theme["text"], True, PP_ALIGN.CENTER)
                    add_text(slide, point, x + .72, y + .22, 4.35, 1.15, 15, theme["text"], True)
            elif design == "luxury":
                # Luxury editorial: generous two-column panels with serif numerals.
                positions = [(.98, 2.42), (6.78, 2.42), (.98, 4.45), (6.78, 4.45)]
                for j, point in enumerate(pts):
                    x, y = positions[j]
                    add_outline(slide, MSO_SHAPE.RECTANGLE, x, y, 5.35, 1.62, theme["accent"], .9)
                    add_text(slide, f"0{j + 1}", x + .25, y + .18, .75, .55, 23, theme["accent"], False, PP_ALIGN.CENTER, "Georgia")
                    add_text(slide, point, x + 1.12, y + .23, 3.95, 1.15, 15, theme["text"], False, font="Georgia")
            elif design == "vibrant":
                # Creative studio: offset blocks, alternating width and bright edge markers.
                positions = [(.82, 2.32, 5.2), (6.85, 2.5, 5.45), (1.2, 4.42, 5.4), (6.45, 4.25, 5.8)]
                for j, point in enumerate(pts):
                    x, y, w = positions[j]
                    add_panel(slide, design, x, y, w, 1.58, theme, j + index)
                    diamond = add_shape(slide, MSO_SHAPE.DIAMOND, x + .2, y + .2, .38, .38, (0, 220, 210) if j % 2 == 0 else (255, 197, 72))
                    diamond.rotation = 12 * (j % 2)
                    add_text(slide, point, x + .78, y + .22, w - 1.0, 1.15, 15, theme["text"], True, font="Arial")
            else:
                # Future/classic: structured 2x2 modules, with style-specific panel treatment.
                positions = [(.9, 2.34), (6.75, 2.34), (.9, 4.42), (6.75, 4.42)]
                for j, point in enumerate(pts):
                    x, y = positions[j]
                    add_panel(slide, design, x, y, 5.55, 1.72, theme, j + index)
                    if design == "future":
                        add_outline(slide, MSO_SHAPE.OVAL, x + .22, y + .2, .48, .48, theme["accent"], 1.5)
                    else:
                        add_text(slide, f"{j + 1:02d}  /", x + .25, y + .18, .85, .35, 12, theme["accent"], True, font="Georgia")
                    add_text(slide, point, x + 1.0, y + .22, 4.22, 1.22, 15, theme["text"], True, font="Georgia" if design == "classic" else "Arial")
        elif layout == "timeline":
            add_header(slide, index, total, theme, item["title"], item.get("subtitle", ""), design, recipe, item.get("kicker", ""), data.get("language", "uk"))
            pts = item["content"][:4]
            if design in ("magazine", "business", "mono"):
                # Editorial / executive timeline: vertical spine with chapter cards.
                spine_x = 1.62 if design != "mono" else 1.92
                add_shape(slide, MSO_SHAPE.RECTANGLE, spine_x, 2.3, .045, 3.85, theme["accent"])
                row_h = 3.78 / max(1, len(pts))
                for j, point in enumerate(pts):
                    y = 2.32 + j * row_h
                    add_shape(slide, MSO_SHAPE.OVAL if design != "mono" else MSO_SHAPE.RECTANGLE, spine_x - .13, y + .12, .32, .32, theme["accent"])
                    if design == "magazine":
                        add_text(slide, f"{deck_label(data.get('language', 'uk'), 'chapter')} 0{j + 1}", 2.15, y + .02, 2.0, .24, 9, theme["accent"], True, font="Georgia")
                        add_text(slide, point, 4.25, y + .02, 7.55, row_h - .12, 15, theme["text"], False, font="Georgia")
                        add_shape(slide, MSO_SHAPE.RECTANGLE, 4.0, y + .02, .025, row_h - .16, theme["accent"])
                    else:
                        add_panel(slide, design, 2.35, y, 9.5, row_h - .12, theme, j + index)
                        add_text(slide, point, 2.68, y + .12, 8.95, row_h - .34, 15, theme["text"], True)
            elif design in ("luxury", "classic"):
                # Formal academic chronology: open horizontal axis with elegant annotations.
                yline = 4.05
                add_shape(slide, MSO_SHAPE.RECTANGLE, 1.05, yline, 11.15, .025, theme["accent"])
                for j, point in enumerate(pts):
                    cx = 1.55 + j * (10.15 / max(1, len(pts) - 1))
                    add_outline(slide, MSO_SHAPE.OVAL, cx - .2, yline - .2, .4, .4, theme["accent"], 1.2)
                    add_text(slide, f"0{j + 1}", cx - .3, 2.55, .6, .35, 17, theme["accent"], False, PP_ALIGN.CENTER, "Georgia")
                    add_text(slide, point, cx - 1.15, 4.62, 2.3, 1.3, 12, theme["text"], False, PP_ALIGN.CENTER, "Georgia")
            else:
                # Dynamic horizontal timeline for school, nature, minimal, future and vibrant.
                yline = 4.0
                add_shape(slide, MSO_SHAPE.RECTANGLE, 1.1, yline, 10.95, .055, theme["accent"])
                for j, point in enumerate(pts):
                    center_x = 1.6 + j * (10.1 / max(1, len(pts) - 1))
                    node_size = .48 if design != "minimal" else .32
                    add_shape(slide, MSO_SHAPE.OVAL, center_x - node_size / 2, yline - .2, node_size, node_size, theme["accent"])
                    add_text(slide, f"0{j + 1}", center_x - .28, 2.25, .56, .35, 16, theme["accent"], True, PP_ALIGN.CENTER, "Arial")
                    card_x = max(.55, min(10.65, center_x - 1.05))
                    add_panel(slide, design, card_x, 4.65, 2.1, 1.45, theme, j + index)
                    add_text(slide, point, card_x + .14, 4.82, 1.82, 1.10, 11, theme["text"], True, PP_ALIGN.CENTER)
        elif layout == "steps":
            add_header(slide, index, total, theme, item["title"], item.get("subtitle", ""), design, recipe, item.get("kicker", ""), data.get("language", "uk"))
            pts = item["content"][:5]
            if design in ("vibrant", "school"):
                # Energetic workshop board: alternating wide and compact tiles.
                for j, point in enumerate(pts):
                    col = j % 2
                    row = j // 2
                    x = .92 + col * 5.9 + (0.18 if (j + index) % 2 else 0)
                    y = 2.32 + row * 1.38
                    add_panel(slide, design, x, y, 5.35, 1.08, theme, j + index)
                    add_shape(slide, MSO_SHAPE.OVAL if design == "school" else MSO_SHAPE.DIAMOND, x + .18, y + .22, .48, .48, theme["accent"])
                    add_text(slide, f"{j + 1:02d}", x + .22, y + .31, .4, .22, 10, theme["bg"], True, PP_ALIGN.CENTER)
                    add_text(slide, point, x + .82, y + .15, 4.28, .78, 14, theme["text"], True)
            elif design in ("magazine", "luxury", "classic"):
                # Chapter sequence: large serif index and editorial text column.
                row_h = 3.85 / max(1, len(pts))
                for j, point in enumerate(pts):
                    y = 2.27 + j * row_h
                    add_text(slide, f"{j + 1:02d}", .98, y + .02, 1.15, .64, 29, theme["accent"], False, PP_ALIGN.CENTER, "Georgia")
                    add_shape(slide, MSO_SHAPE.RECTANGLE, 2.32, y + .04, .035, row_h - .17, theme["accent"])
                    add_text(slide, point, 2.68, y + .06, 9.1, row_h - .22, 15, theme["text"], False, font="Georgia")
            else:
                row_gap = .88 if len(pts) >= 5 else 1.0
                for j, point in enumerate(pts):
                    y = 2.2 + j * row_gap
                    size = .55
                    add_shape(slide, MSO_SHAPE.OVAL if design != "mono" else MSO_SHAPE.RECTANGLE, .92, y, size, size, theme["accent"])
                    number_color = theme["bg"] if design != "mono" else (255, 255, 255)
                    add_text(slide, str(j + 1), .99, y + .09, .4, .3, 13, number_color, True, PP_ALIGN.CENTER)
                    add_panel(slide, design, 1.72, y - .03, 10.55, .68, theme, j + index)
                    add_text(slide, point, 2.0, y + .08, 9.95, .48, 15, theme["text"], True)
        elif layout == "quote":
            add_header(slide, index, total, theme, item["title"], item.get("subtitle", ""), design, recipe, item.get("kicker", ""), data.get("language", "uk"))
            content_points = item.get("content", [])
            explicit_highlight = str(item.get("highlight") or "").strip()
            quote = explicit_highlight or (content_points[0] if content_points else item["title"])
            add_text(slide, "“", .92, 2.05, 1.0, 1.2, 76, theme["accent"], True, font="Georgia")
            add_panel(slide, design, 1.82, 2.45, 10.25, 1.72, theme, index)
            add_text(slide, quote, 2.25, 2.73, 9.35, 1.15, 23, theme["text"], True, PP_ALIGN.CENTER, "Georgia" if design in ("luxury", "magazine", "classic") else "Arial", MSO_ANCHOR.MIDDLE)
            rest = _supporting_points(item, 4)
            if rest:
                add_bullets(slide, rest, 1.45, 4.65, 10.45, 1.45, theme, 15)
        elif layout == "comparison":
            add_header(slide, index, total, theme, item["title"], item.get("subtitle", ""), design, recipe, item.get("kicker", ""), data.get("language", "uk"))
            pts = item["content"]
            left, right = pts[::2], pts[1::2]
            add_panel(slide, design, .82, 2.25, 5.55, 3.95, theme, index)
            add_panel(slide, design, 6.95, 2.25, 5.55, 3.95, theme, index + 1)
            add_text(slide, item.get("left_title") or "A", 1.18, 2.58, 4.8, .5, 19, theme["accent"], True, PP_ALIGN.CENTER)
            add_text(slide, item.get("right_title") or "B", 7.3, 2.58, 4.8, .5, 19, theme["accent"], True, PP_ALIGN.CENTER)
            add_bullets(slide, left or pts[:1], 1.2, 3.3, 4.85, 2.5, theme, 15)
            add_bullets(slide, right or pts[-1:], 7.35, 3.3, 4.8, 2.5, theme, 15)
        elif layout == "focus":
            add_header(slide, index, total, theme, item["title"], item.get("subtitle", ""), design, recipe, item.get("kicker", ""), data.get("language", "uk"))
            content_points = item.get("content", [])
            focus = item.get("highlight") or (content_points[0] if content_points else item["title"])
            add_text(slide, focus, .95, 2.15, 11.2, 1.5, 29, theme["accent"], True, PP_ALIGN.CENTER, "Georgia" if design in ("luxury", "magazine", "classic") else "Arial", MSO_ANCHOR.MIDDLE)
            add_shape(slide, MSO_SHAPE.RECTANGLE, 2.1, 3.9, 9.1, .04, theme["accent"])
            add_bullets(slide, _supporting_points(item, 4), 1.7, 4.35, 9.9, 1.55, theme, 15)
        else:  # split
            add_header(slide, index, total, theme, item["title"], item.get("subtitle", ""), design, recipe, item.get("kicker", ""), data.get("language", "uk"))
            add_panel(slide, design, .82, 2.25, 7.05, 4.05, theme, index)
            add_bullets(slide, item["content"][:4], 1.22, 2.65, 6.3, 3.35, theme, 17)
            takeaway = item.get("highlight") or item.get("kicker") or deck_label(data.get("language", "uk"), "key_idea")
            if design in ("future", "vibrant"):
                add_outline(slide, MSO_SHAPE.OVAL, 9.0, 2.65, 2.25, 2.25, theme["accent"], 2)
                add_outline(slide, MSO_SHAPE.OVAL, 9.42, 3.07, 1.4, 1.4, theme["accent"], 1.1)
                add_text(slide, takeaway, 8.52, 5.05, 3.2, .72, 13, theme["accent"], True, PP_ALIGN.CENTER)
            elif design == "magazine":
                add_text(slide, takeaway, 8.45, 2.55, 3.45, 1.65, 25, theme["accent"], True, PP_ALIGN.LEFT, "Georgia")
                add_shape(slide, MSO_SHAPE.RECTANGLE, 8.5, 4.45, 2.8, .075, theme["accent"])
            elif design == "nature":
                leaf = add_shape(slide, MSO_SHAPE.OVAL, 9.35, 2.65, 1.25, 2.55, theme["accent"])
                leaf.rotation = 32
                add_text(slide, takeaway, 8.35, 5.35, 3.4, .55, 12, theme["accent"], True, PP_ALIGN.CENTER)
            elif design == "mono":
                add_shape(slide, MSO_SHAPE.RECTANGLE, 8.75, 2.7, 2.75, 2.3, theme["accent"])
                add_text(slide, takeaway, 9.0, 3.22, 2.25, 1.12, 17, (255, 255, 255), True, PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)
            else:
                add_shape(slide, MSO_SHAPE.RECTANGLE, 8.8, 2.85, 2.55, .07, theme["accent"])
                add_text(slide, takeaway, 8.8, 3.2, 3.15, 1.55, 20, theme["accent"], True, PP_ALIGN.LEFT, "Georgia" if design in ("luxury", "classic") else "Arial", MSO_ANCHOR.MIDDLE)
            if item.get("highlight"):
                add_text(slide, item["highlight"], 8.2, 5.65, 3.8, .6, 13, theme["accent"], True, PP_ALIGN.CENTER)
        _append_pptx_fact_source(slide, item, design, theme)
        photo_credit = str(item.get("image_credit") or "").strip()
        source_note = f"\n{deck_label(data.get("language", "uk"), "photo_credit")}: {photo_credit}" if photo_credit else ""
        fact_note = f"\n{deck_label(data.get('language', 'uk'), 'fact')}: {item.get('fact')} / {deck_label(data.get('language', 'uk'), 'source')}: {item.get('source')}" if item.get("fact") or item.get("source") else ""
        try:
            references_note = ""
            if index == total - 1 and data.get("reference_lines"):
                references_note = "\n\n" + deck_label(data.get("language", "uk"), "sources") + ":\n" + "\n".join(str(line) for line in data.get("reference_lines", []))
            slide.notes_slide.notes_text_frame.text = str(item.get("notes", "")).strip() + fact_note + source_note + references_note
        except Exception:
            log.exception("Could not write speaker notes for slide %s", index + 1)
        add_fade_transition(slide)
    prs.save(output_path)

# ============================================================
# 8. PDF EXPORT
# ============================================================
async def ensure_unicode_font() -> str:
    """Find a bundled/system Unicode font. Never download fonts during a live request."""
    base_dir = Path(__file__).resolve().parent
    candidates = [
        base_dir / "fonts" / "DejaVuSans.ttf",
        base_dir / "fonts" / "LiberationSans-Regular.ttf",
        Path("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"),
        Path("/usr/share/fonts/dejavu/DejaVuSans.ttf"),
        Path("/usr/share/fonts/truetype/liberation2/LiberationSans-Regular.ttf"),
        Path("/usr/share/fonts/truetype/freefont/FreeSans.ttf"),
    ]
    for candidate in candidates:
        if candidate.is_file() and candidate.stat().st_size > 50_000:
            return str(candidate)
    raise RuntimeError(
        "Unicode TTF font is missing. Add fonts/DejaVuSans.ttf to the repository "
        "or install DejaVu Sans on the Render image; PDF export is unavailable until then."
    )

def create_pdf(data: dict, design: str, output_path: str, font_path: str):
    """Export a style-led PDF with a seeded per-deck visual recipe."""
    theme = copy.deepcopy(DESIGNS[design])
    seed = int(data.get("visual_seed") or random.randint(1, 2_000_000_000))
    recipe = make_deck_recipe(design, seed, len(data["slides"]), data.get("visual_concept", ""))
    theme["accent"] = recipe["accent"]
    bg = colors.HexColor("#%02x%02x%02x" % theme["bg"])
    accent = colors.HexColor("#%02x%02x%02x" % theme["accent"])
    card = colors.HexColor("#%02x%02x%02x" % theme["card"])
    text_color = colors.HexColor("#%02x%02x%02x" % theme["text"])
    muted = colors.HexColor("#%02x%02x%02x" % theme["muted"])
    if "PrezentSans" not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont("PrezentSans", font_path))

    page_w, page_h = landscape(A4)
    pdf = canvas.Canvas(output_path, pagesize=(page_w, page_h))
    rng = random.Random(seed)
    total = len(data["slides"])

    title_size = {"minimal": 27, "future": 25, "business": 25, "school": 27,
                  "luxury": 29, "magazine": 30, "nature": 27, "vibrant": 30,
                  "classic": 27, "mono": 28}.get(design, 26)
    title_align = TA_CENTER if design in ("luxury", "classic") else TA_LEFT
    title_style = ParagraphStyle(
        "prezent-title", fontName="PrezentSans", fontSize=title_size,
        leading=title_size * 1.17, textColor=text_color, alignment=title_align,
        spaceAfter=0,
    )
    body_style = ParagraphStyle(
        "prezent-body", fontName="PrezentSans", fontSize=11.5, leading=15.5,
        textColor=text_color, alignment=TA_LEFT,
    )
    small_style = ParagraphStyle(
        "prezent-small", fontName="PrezentSans", fontSize=9.5, leading=12.5,
        textColor=muted,
    )
    card_title_style = ParagraphStyle(
        "prezent-card-title", fontName="PrezentSans", fontSize=12.5,
        leading=15.5, textColor=accent,
    )
    focus_style = ParagraphStyle(
        "prezent-focus", fontName="PrezentSans", fontSize=19,
        leading=24, textColor=accent, alignment=TA_CENTER,
    )

    def para(text, x, top, width, style, max_height=130, min_font_size=6.5):
        raw_text = str(text)
        base_size = float(getattr(style, "fontSize", 11.0))
        style_name = str(getattr(style, "name", "")).lower()
        content_paragraph = not any(token in style_name for token in ("small", "credit", "source", "footer", "label"))
        if content_paragraph and len(raw_text) >= 90:
            base_size = max(base_size, 11.0)
            min_font_size = max(min_font_size, 11.0)
        base_leading = float(getattr(style, "leading", base_size * 1.25))
        leading_ratio = base_leading / max(1.0, base_size)

        def make_paragraph(value, font_size):
            fitted_style = ParagraphStyle(
                f"{style.name}-fit-{int(font_size * 10)}",
                parent=style,
                fontSize=font_size,
                leading=max(font_size * 1.08, font_size * leading_ratio),
            )
            safe = escape(str(value)).replace("\n", "<br/>")
            paragraph = Paragraph(safe, fitted_style)
            _, height = paragraph.wrap(width, max_height)
            return paragraph, height

        chosen = None
        minimum_size = max(6.5, float(min_font_size))
        for font_size in [base_size - step for step in range(0, max(1, int(base_size - minimum_size) + 1))]:
            if font_size < minimum_size:
                continue
            candidate, height = make_paragraph(raw_text, font_size)
            if height <= max_height:
                chosen = (candidate, height)
                break
        if chosen is None:
            # At the minimum readable size, trim by words until the paragraph fits.
            words = raw_text.split()
            lo, hi = 0, len(words)
            best = (Paragraph("", style), 0)
            while lo <= hi:
                mid = (lo + hi) // 2
                sample = " ".join(words[:mid]) + ("…" if mid < len(words) else "")
                candidate, height = make_paragraph(sample, minimum_size)
                if height <= max_height:
                    best = (candidate, height)
                    lo = mid + 1
                else:
                    hi = mid - 1
            chosen = best
        item, height = chosen
        item.drawOn(pdf, x, top - height)
        return top - height

    def draw_chrome(index):
        pdf.setFillColor(bg)
        pdf.rect(0, 0, page_w, page_h, fill=1, stroke=0)
        quiet_slide = bool(index not in (0, total - 1) and (index + 1) % recipe.get("quiet_every", 3) == 0)
        # Core design identity remains on quiet slides; optional decoration is simply skipped.
        # Every art direction gets its own frame, geometry and page furniture.
        if design == "minimal":
            pdf.setFillColor(accent)
            pdf.rect(18, 34, 3, page_h - 68, fill=1, stroke=0)
            pdf.setStrokeColor(muted)
            pdf.setLineWidth(.5)
            pdf.line(42, 40, page_w - 38, 40)
        elif design == "future":
            pdf.setStrokeColor(colors.HexColor("#1D3153"))
            pdf.setLineWidth(.35)
            for gx in range(1, 12):
                x = gx * page_w / 12
                pdf.line(x, 0, x, page_h)
            for gy in range(1, 8):
                y = gy * page_h / 8
                pdf.line(0, y, page_w, y)
            pdf.setStrokeColor(accent)
            pdf.setLineWidth(1.4)
            ring_x = page_w - 70 + rng.uniform(-8, 8)
            pdf.circle(ring_x, page_h - 70, 28 + rng.uniform(-3, 4), fill=0, stroke=1)
            pdf.setStrokeColor(colors.HexColor("#6B5BFF"))
            pdf.circle(ring_x, page_h - 70, 15, fill=0, stroke=1)
        elif design == "business":
            pdf.setFillColor(colors.HexColor("#0D1928"))
            pdf.rect(0, 0, 24, page_h, fill=1, stroke=0)
            pdf.setFillColor(accent)
            pdf.rect(24, 0, 5, page_h, fill=1, stroke=0)
            for j in range(4):
                h = 9 + rng.uniform(3, 17)
                pdf.rect(page_w - 36 - j * 12, 48, 6, h, fill=1, stroke=0)
        elif design == "school":
            pdf.setFillColor(colors.HexColor("#FFE0A9"))
            pdf.circle(page_w - 60, page_h - 55, 33, fill=1, stroke=0)
            pdf.setFillColor(accent)
            pdf.circle(page_w - 60, page_h - 55, 14, fill=1, stroke=0)
            pdf.setFillColor(colors.HexColor("#FFC447"))
            pdf.circle(38, 44, 14, fill=1, stroke=0)
        elif design == "luxury":
            pdf.setStrokeColor(accent)
            pdf.setLineWidth(1.0)
            pdf.rect(18, 18, page_w - 36, page_h - 36, fill=0, stroke=1)
            pdf.setStrokeColor(colors.HexColor("#927441"))
            pdf.setLineWidth(.45)
            pdf.rect(24, 24, page_w - 48, page_h - 48, fill=0, stroke=1)
            pdf.setFillColor(accent)
            pdf.rect(page_w / 2 - 28, page_h - 27, 56, 2, fill=1, stroke=0)
        elif design == "magazine":
            pdf.setFillColor(accent)
            pdf.rect(0, 0, 18, page_h, fill=1, stroke=0)
            pdf.setFillColor(colors.HexColor("#E8DED0"))
            pdf.rect(page_w - 82, 0, 82, page_h, fill=1, stroke=0)
            pdf.setFillColor(text_color)
            pdf.setFont("PrezentSans", 35)
            pdf.drawCentredString(page_w - 41, page_h - 78, f"{index + 1:02d}")
        elif design == "nature":
            pdf.setFillColor(colors.HexColor("#DAE8CF"))
            pdf.ellipse(page_w - 112, page_h - 164, page_w - 28, page_h - 35, fill=1, stroke=0)
            pdf.saveState()
            pdf.translate(page_w - 68, page_h - 102)
            pdf.rotate(28 + rng.randint(-9, 9))
            pdf.setFillColor(accent)
            pdf.ellipse(-13, -48, 13, 48, fill=1, stroke=0)
            pdf.restoreState()
        elif design == "vibrant":
            pdf.setFillColor(colors.HexColor("#35265C"))
            pdf.rect(page_w - 88, 0, 88, page_h, fill=1, stroke=0)
            pdf.saveState()
            pdf.translate(page_w - 48, page_h - 70)
            pdf.rotate(35 + rng.randint(-12, 12))
            pdf.setFillColor(colors.HexColor("#00DCD2"))
            pdf.rect(-20, -20, 40, 40, fill=1, stroke=0)
            pdf.restoreState()
            pdf.setFillColor(colors.HexColor("#FF5B91"))
            pdf.circle(page_w - 45, 78, 20, fill=1, stroke=0)
        elif design == "classic":
            pdf.setStrokeColor(accent)
            pdf.setLineWidth(1)
            pdf.line(32, page_h - 34, page_w - 32, page_h - 34)
            pdf.setStrokeColor(muted)
            pdf.setLineWidth(.45)
            pdf.line(32, page_h - 39, page_w - 32, page_h - 39)
            pdf.line(32, 34, page_w - 32, 34)
        else:  # mono / brutalist
            pdf.setFillColor(colors.HexColor("#161616"))
            pdf.rect(0, 0, 44, page_h, fill=1, stroke=0)
            pdf.rect(page_w - 46, page_h - 52, 26, 8, fill=1, stroke=0)
            pdf.setFillColor(colors.HexColor("#777777"))
            pdf.rect(page_w - 46, page_h - 66, 14, 3, fill=1, stroke=0)

    def draw_panel(x, y, width, height, variant=0):
        pdf.saveState()
        if design in ("minimal", "classic", "magazine", "mono"):
            pdf.setFillColor(card)
            pdf.setStrokeColor(accent if design != "minimal" else muted)
            pdf.setLineWidth(.7 if design != "mono" else 0)
            pdf.rect(x, y, width, height, fill=1, stroke=1 if design != "mono" else 0)
        else:
            pdf.setFillColor(card)
            pdf.setStrokeColor(accent if design in ("future", "luxury") else card)
            pdf.setLineWidth(1 if design in ("future", "luxury") else .2)
            pdf.roundRect(x, y, width, height, 13 if design != "future" else 4, fill=1, stroke=1 if design in ("future", "luxury") else 0)
        if design in ("business", "school", "magazine", "nature"):
            pdf.setFillColor(accent)
            pdf.rect(x, y, 5 if design != "school" else 8, height, fill=1, stroke=0)
        elif design == "vibrant":
            palette = [colors.HexColor("#FF5B91"), colors.HexColor("#00DCD2"), colors.HexColor("#FFC547")]
            pdf.setFillColor(palette[variant % len(palette)])
            pdf.rect(x, y + height - 5, width, 5, fill=1, stroke=0)
        elif design == "mono":
            pdf.setFillColor(text_color)
            pdf.rect(x, y, 7, height, fill=1, stroke=0)
        pdf.restoreState()

    def draw_bullets(points, x, top, width, max_bottom=72, size=11.5, accent_color=accent):
        current_top = top
        style = ParagraphStyle(
            f"bullet-{size}", parent=body_style, fontSize=size,
            leading=size * 1.32, textColor=text_color,
        )
        for n, point in enumerate(points[:5]):
            pdf.setFillColor(accent_color)
            pdf.circle(x + 3, current_top - 7, 2.2, fill=1, stroke=0)
            after = para(point, x + 14, current_top, width - 17, style, 78, min_font_size=max(11.0, min(size, 12.0)))
            current_top = after - 9
            if current_top < max_bottom:
                break
        return current_top

    def draw_header(index, item):
        language_code = data.get("language", "uk")
        fallback_by_language = {
            "uk": {"minimal": "ІДЕЯ / ОРИГІНАЛ", "future": "ТЕХНОЛОГІЇ / ЛАБОРАТОРІЯ", "business": "СТРАТЕГІЧНИЙ БРИФ", "school": "НАВЧАННЯ / КРЕАТИВ", "luxury": "ПРЕМІАЛЬНЕ ВИДАННЯ", "magazine": "ОГЛЯД ЗНАНЬ", "nature": "ПРИРОДА / ПОГЛЯД", "vibrant": "КРЕАТИВНА КУЛЬТУРА", "classic": "АКАДЕМІЧНЕ ВИДАННЯ", "mono": "МОНОХРОМНІ НОТАТКИ"},
            "ru": {"minimal": "ИДЕЯ / ОРИГИНАЛ", "future": "ТЕХНОЛОГИИ / ЛАБОРАТОРИЯ", "business": "СТРАТЕГИЧЕСКИЙ БРИФ", "school": "ОБУЧЕНИЕ / КРЕАТИВ", "luxury": "ПРЕМИАЛЬНОЕ ИЗДАНИЕ", "magazine": "ОБЗОР ЗНАНИЙ", "nature": "ПРИРОДА / ВЗГЛЯД", "vibrant": "КРЕАТИВНАЯ КУЛЬТУРА", "classic": "АКАДЕМИЧЕСКОЕ ИЗДАНИЕ", "mono": "МОНОХРОМНЫЕ ЗАМЕТКИ"},
            "en": {"minimal": "IDEA / ORIGINAL", "future": "TECHNOLOGY / LAB", "business": "EXECUTIVE BRIEF", "school": "LEARNING / CREATIVE", "luxury": "PREMIUM EDITION", "magazine": "KNOWLEDGE REVIEW", "nature": "NATURE / PERSPECTIVE", "vibrant": "CREATIVE CULTURE", "classic": "ACADEMIC EDITION", "mono": "MONOCHROME NOTES"},
        }
        lang = language_code if language_code in fallback_by_language else "uk"
        fallback = fallback_by_language[lang].get(design, "PREZENT AI")
        label = str(item.get("kicker") or "").strip()[:70] or fallback
        mode = recipe.get("headers", ["rule"])[index % len(recipe.get("headers", ["rule"]))]
        pdf.setFillColor(accent)
        pdf.setFont("PrezentSans", 7.5)
        if mode == "tag":
            pdf.roundRect(44, page_h - 39, min(250, max(85, len(label) * 4.6)), 17, 5, fill=0, stroke=1)
            pdf.drawString(51, page_h - 33, label)
        elif mode == "numbered":
            pdf.setFont("PrezentSans", 17)
            pdf.drawString(42, page_h - 38, f"{index + 1:02d}")
            pdf.setFont("PrezentSans", 7.5)
            pdf.drawString(82, page_h - 31, label)
        elif mode == "sidebar":
            pdf.setFillColor(accent); pdf.rect(39, page_h - 54, 3, 30, fill=1, stroke=0)
            pdf.drawString(50, page_h - 31, label)
        else:
            pdf.drawString(48 if design != "mono" else 60, page_h - 28, label)
        pdf.setFillColor(muted)
        pdf.setFont("PrezentSans", 8)
        pdf.drawRightString(page_w - 30, 25, f"{index + 1:02d} / {total:02d}")
        x = 60 if design == "mono" else 44
        width = page_w - x - (105 if design in ("magazine", "vibrant") else 82)
        top = page_h - (66 if mode == "sidebar" else 60)
        bottom = para(item["title"], x, top, width, title_style, 82)
        subtitle = item.get("subtitle", "")
        if subtitle:
            bottom = para(subtitle, x + 1, bottom - 10, width - 4, small_style, 30)
        if mode == "rule":
            pdf.setStrokeColor(accent); pdf.setLineWidth(1.1 + (index % 3) * .3)
            pdf.line(x, bottom - 17, x + 80 + (index % 4) * 22, bottom - 17)
        elif mode == "numbered":
            pdf.setFillColor(accent); pdf.rect(x, bottom - 17, 3, 25, fill=1, stroke=0)
        return x, width

    def draw_cover(index, item):
        title = item["title"]
        subtitle = item.get("subtitle", "")
        if design == "luxury":
            pdf.setStrokeColor(accent); pdf.setLineWidth(1.1)
            pdf.rect(44, 48, page_w - 88, page_h - 96, fill=0, stroke=1)
            pdf.setFillColor(accent); pdf.setFont("PrezentSans", 10)
            pdf.drawCentredString(page_w / 2, page_h - 100, deck_label(data.get("language", "uk"), "brand"))
            style = ParagraphStyle("lux-cover", parent=title_style, fontSize=31, leading=38, alignment=TA_CENTER, textColor=text_color)
            para(title, 110, page_h - 150, page_w - 220, style, 150)
            if subtitle:
                para(subtitle, 150, page_h - 325, page_w - 300, small_style, 60)
            pdf.setFillColor(accent); pdf.rect(page_w / 2 - 35, 108, 70, 1.5, fill=1, stroke=0)
        elif design == "classic":
            pdf.setStrokeColor(accent); pdf.setLineWidth(1.3)
            pdf.rect(55, 55, page_w - 110, page_h - 110, fill=0, stroke=1)
            pdf.setStrokeColor(muted); pdf.setLineWidth(.55)
            pdf.rect(63, 63, page_w - 126, page_h - 126, fill=0, stroke=1)
            pdf.setFillColor(accent); pdf.setFont("PrezentSans", 9)
            pdf.drawCentredString(page_w / 2, page_h - 130, deck_label(data.get("language", "uk"), "academic_presentation"))
            style = ParagraphStyle("classic-cover", parent=title_style, fontSize=29, leading=36, alignment=TA_CENTER)
            para(title, 115, page_h - 175, page_w - 230, style, 145)
            if subtitle: para(subtitle, 145, page_h - 345, page_w - 290, small_style, 65)
        elif design == "mono":
            pdf.setFillColor(colors.HexColor("#161616")); pdf.rect(55, 65, 430, 465, fill=1, stroke=0)
            pdf.setFillColor(colors.white); pdf.setFont("PrezentSans", 8)
            pdf.drawString(78, 490, deck_label(data.get("language", "uk"), "mono_edition"))
            mono_style = ParagraphStyle("mono-cover", parent=title_style, fontSize=30, leading=36, textColor=colors.white)
            para(title, 78, 445, 370, mono_style, 170)
            if subtitle:
                mono_sub = ParagraphStyle("mono-sub", parent=small_style, textColor=colors.HexColor("#D0D0D0"))
                para(subtitle, 82, 245, 355, mono_sub, 70)
            pdf.setFillColor(colors.HexColor("#181818")); pdf.setFont("PrezentSans", 105)
            pdf.drawCentredString(650, 300, "01")
            pdf.setFont("PrezentSans", 10); pdf.drawCentredString(650, 210, deck_label(data.get("language", "uk"), "idea_form_impact"))
        else:
            x = 52
            title_w = 430 if design in ("business", "magazine", "vibrant") else 470
            if design == "magazine":
                pdf.setFillColor(colors.HexColor("#E8DED0")); pdf.rect(page_w - 230, 0, 230, page_h, fill=1, stroke=0)
                pdf.setFillColor(accent); pdf.setFont("PrezentSans", 105); pdf.drawCentredString(page_w - 115, 295, "01")
                pdf.setFillColor(text_color); pdf.setFont("PrezentSans", 17); pdf.drawCentredString(page_w - 115, 215, deck_label(data.get("language", "uk"), "feature_story"))
            elif design == "business":
                pdf.setFillColor(colors.HexColor("#0D1928")); pdf.rect(page_w - 220, 0, 220, page_h, fill=1, stroke=0)
                for j, tag_key in enumerate(("context", "insight", "action")):
                    tag = deck_label(data.get("language", "uk"), tag_key)
                    yy = 390 - j * 100
                    pdf.setFillColor(card); pdf.rect(page_w - 190, yy - 25, 145, 58, fill=1, stroke=0)
                    pdf.setFillColor(accent); pdf.rect(page_w - 190, yy - 25, 5, 58, fill=1, stroke=0)
                    pdf.setFillColor(text_color); pdf.setFont("PrezentSans", 11); pdf.drawString(page_w - 172, yy - 3, tag)
            elif design == "future":
                pdf.setStrokeColor(accent); pdf.setLineWidth(2)
                ring_x = page_w - 120 + rng.uniform(-12, 12)
                pdf.circle(ring_x, 310, 86, fill=0, stroke=1)
                pdf.setStrokeColor(colors.HexColor("#6B5BFF")); pdf.circle(ring_x, 310, 55, fill=0, stroke=1)
                pdf.setFillColor(accent); pdf.circle(ring_x, 310, 13, fill=1, stroke=0)
            elif design == "school":
                pdf.setFillColor(colors.HexColor("#FFE0A9")); pdf.circle(page_w - 125, 385, 78, fill=1, stroke=0)
                pdf.setFillColor(accent); pdf.circle(page_w - 125, 385, 38, fill=1, stroke=0)
                pdf.setFillColor(colors.HexColor("#FFC447")); pdf.saveState(); pdf.translate(page_w - 85, 190); pdf.rotate(45); pdf.rect(-24, -24, 48, 48, fill=1, stroke=0); pdf.restoreState()
            elif design == "nature":
                pdf.setFillColor(colors.HexColor("#DAE8CF")); pdf.ellipse(page_w - 215, 130, page_w - 35, 470, fill=1, stroke=0)
                pdf.saveState(); pdf.translate(page_w - 125, 300); pdf.rotate(30 + rng.randint(-10, 10)); pdf.setFillColor(accent); pdf.ellipse(-24, -110, 24, 110, fill=1, stroke=0); pdf.restoreState()
            elif design == "vibrant":
                pdf.setFillColor(colors.HexColor("#35265C")); pdf.rect(page_w - 220, 0, 220, page_h, fill=1, stroke=0)
                pdf.saveState(); pdf.translate(page_w - 120, 380); pdf.rotate(45 + rng.randint(-12, 12)); pdf.setFillColor(colors.HexColor("#00DCD2")); pdf.rect(-48, -48, 96, 96, fill=1, stroke=0); pdf.restoreState()
                pdf.setFillColor(colors.HexColor("#FF5B91")); pdf.circle(page_w - 100, 165, 40, fill=1, stroke=0)
            elif design == "minimal":
                pdf.setFillColor(accent); pdf.rect(520, 120, 3, 340, fill=1, stroke=0)
                pdf.setFillColor(accent); pdf.setFont("PrezentSans", 88); pdf.drawString(560, 320, "01")
                pdf.setFillColor(text_color); pdf.setFont("PrezentSans", 9); pdf.drawString(565, 265, deck_label(data.get("language", "uk"), "clear_story").replace("\n", " / "))
            pdf.setFillColor(accent); pdf.setFont("PrezentSans", 9)
            pdf.drawString(x, page_h - 105, deck_label(data.get("language", "uk"), "cover_original"))
            cover_style = ParagraphStyle("cover-main", parent=title_style, fontSize=33 if design in ("school", "vibrant", "magazine") else 30, leading=39, textColor=text_color)
            para(title, x, page_h - 145, title_w, cover_style, 180)
            if subtitle:
                para(subtitle, x + 2, page_h - 340, title_w - 10, small_style, 80)
        pdf.setFillColor(muted); pdf.setFont("PrezentSans", 7.5)
        pdf.drawString(50, 24, "PREZENT AI")

    def draw_photo_credit(value, x, top, width, language="uk"):
        credit_text = compact_photo_credit(value, author_limit=28, language=language)
        credit_style = ParagraphStyle("photo-credit", parent=body_style, fontSize=7.0, leading=8.2, textColor=muted)
        return para(credit_text, x, top, width, credit_style, 19, min_font_size=7.0)

    def draw_photo(item, x, y, w, h):
        path = item.get("image_path")
        if not path or not Path(path).is_file():
            return False
        try:
            pdf.drawImage(ImageReader(path), x, y, width=w, height=h, preserveAspectRatio=True, anchor="c", mask="auto")
            return True
        except Exception as exc:
            log.debug("PDF photo draw failed: %s", exc)
            return False

    for i, item in enumerate(data["slides"]):
        draw_chrome(i)
        layout = resolve_render_layout(item)
        if layout == "cover":
            draw_cover(i, item)
            if item.get("image_path"):
                pdf.setStrokeColor(accent); pdf.setLineWidth(1.0)
                if draw_photo(item, page_w - 245, 76, 205, page_h - 155):
                    pdf.rect(page_w - 245, 76, 205, page_h - 155, fill=0, stroke=1)
                    if item.get("image_credit"):
                        draw_photo_credit(item["image_credit"], page_w - 244, 73, 205, data.get("language", "uk"))
            pdf.showPage()
            continue

        if layout == "section":
            pdf.setFillColor(accent); pdf.rect(44, 58, 7, page_h - 116, fill=1, stroke=0)
            pdf.setFillColor(accent); pdf.setFont("PrezentSans", 9)
            pdf.drawString(75, page_h - 115, (item.get("kicker") or deck_label(data.get("language", "uk"), "section"))[:60])
            section_style = ParagraphStyle("section-title", parent=title_style, fontSize=34, leading=41, alignment=TA_LEFT)
            para(item["title"], 75, page_h - 145, page_w - 150, section_style, 145)
            if item.get("subtitle"):
                para(item["subtitle"], 78, page_h - 315, page_w - 170, body_style, 70)
            pdf.showPage()
            continue

        x0, title_w = draw_header(i, item)
        points = item.get("content", [])[:5]
        if layout == "photo_text":
            left_w = 360
            draw_panel(42, 82, left_w, 300, i)
            draw_bullets(points[:4], 60, 350, left_w - 36, 235, 11.2)
            if draw_photo(item, 435, 90, page_w - 485, 300):
                pdf.setStrokeColor(accent); pdf.setLineWidth(1.0)
                pdf.rect(435, 90, page_w - 485, 300, fill=0, stroke=1)
                if item.get("image_credit"):
                    draw_photo_credit(item["image_credit"], 438, 79, page_w - 490, data.get("language", "uk"))
            else:
                draw_panel(42, 82, page_w - 84, 300, i + 1)
                draw_bullets(points[:4], 62, 350, page_w - 124, 230, 11.2)
        elif layout == "portrait_fact":
            if draw_photo(item, page_w - 300, 86, 235, 310):
                pdf.setStrokeColor(accent); pdf.setLineWidth(1.0)
                pdf.rect(page_w - 300, 86, 235, 310, fill=0, stroke=1)
                if item.get("image_credit"):
                    draw_photo_credit(item["image_credit"], page_w - 299, 75, 235, data.get("language", "uk"))
            para(item.get("fact") or deck_label(data.get("language", "uk"), "unknown_origin"), 55, 365, page_w - 390, focus_style, 120)
            draw_bullets(points[:4], 64, 235, page_w - 410, 125, 11.0)
        elif layout == "stat":
            value = item.get("stat_value") or item.get("highlight") or (points[0] if points else "")
            label = item.get("stat_label") or (points[1] if len(points) > 1 else "")
            draw_panel(52, 85, page_w - 104, 300, i)
            value_style = ParagraphStyle("stat-value", parent=focus_style, fontSize=45 if len(value) < 12 else 30, leading=50, alignment=TA_CENTER)
            para(value, 80, 355, 330, value_style, 170)
            if label:
                para(label, 430, 325, page_w - 500, card_title_style, 110)
            extra_stat_points = item.get("content", [])[2:5]
            if extra_stat_points:
                draw_bullets(extra_stat_points, 430, 205, page_w - 500, 82, 11.0)
            elif item.get("highlight") and value != item.get("highlight"):
                para(item["highlight"], 430, 190, page_w - 500, body_style, 100)
        elif layout == "two_col":
            left_lead = points[0] if points else item.get("highlight", "")
            right_points = points[1:] or ([item.get("highlight")] if item.get("highlight") else [])
            draw_panel(42, 82, 330, 305, i)
            para(item.get("left_title") or deck_label(data.get("language", "uk"), "key_idea"), 60, 360, 290, card_title_style, 35)
            lead_style = ParagraphStyle("two-col-lead", parent=body_style, fontSize=17, leading=22)
            para(left_lead, 62, 315, 285, lead_style, 210)
            pdf.setStrokeColor(accent); pdf.setLineWidth(1.5); pdf.line(395, 90, 395, 390)
            para(item.get("right_title") or deck_label(data.get("language", "uk"), "explanation"), 420, 365, page_w - 475, card_title_style, 35)
            draw_bullets(right_points[:3], 420, 325, page_w - 475, 85, 11.0)
        elif layout == "grid3":
            points3 = points[:3]
            if len(points3) != 3:
                raise ValueError("grid3 renderer requires exactly three content points; resolve_render_layout should route otherwise")
            gap = 14
            card_w = (page_w - 108 - 2 * gap) / 3
            for j, point in enumerate(points3):
                px = 42 + j * (card_w + gap)
                draw_panel(px, 90, card_w, 292, i + j)
                pdf.setFillColor(accent); pdf.setFont("PrezentSans", 26)
                pdf.drawString(px + 16, 330, f"0{j + 1}")
                pdf.setStrokeColor(accent); pdf.setLineWidth(.8)
                pdf.line(px + 16, 305, px + card_w - 16, 305)
                para(point, px + 18, 280, card_w - 36, body_style, 165)
        elif layout == "timeline":
            line_y = 245
            pdf.setStrokeColor(accent); pdf.setLineWidth(2.2)
            pdf.line(65, line_y, page_w - 65, line_y)
            pts = points[:4]
            for j, point in enumerate(pts):
                center = 70 + j * ((page_w - 140) / max(1, len(pts) - 1))
                pdf.setFillColor(accent); pdf.circle(center, line_y, 6, fill=1, stroke=0)
                pdf.setFillColor(muted); pdf.setFont("PrezentSans", 9)
                pdf.drawCentredString(center, line_y + 32, f"0{j + 1}")
                card_x = max(38, min(page_w - 190, center - 76))
                draw_panel(card_x, 85, 152, 115, j + i)
                para(point, card_x + 10, 180, 132, body_style, 90)
        elif layout == "comparison":
            col_gap = 18
            col_w = (page_w - 105 - col_gap) / 2
            left_x, right_x = 42, 42 + col_w + col_gap
            draw_panel(left_x, 85, col_w, 315, i)
            draw_panel(right_x, 85, col_w, 315, i + 1)
            para(item.get("left_title") or "A", left_x + 18, 374, col_w - 36, card_title_style, 40)
            para(item.get("right_title") or "B", right_x + 18, 374, col_w - 36, card_title_style, 40)
            draw_bullets(points[::2] or points[:1], left_x + 16, 330, col_w - 32, 105, 10.5)
            draw_bullets(points[1::2] or points[-1:], right_x + 16, 330, col_w - 32, 105, 10.5)
        elif layout == "cards":
            pts = points[:4]
            if design == "magazine" and pts:
                # Editorial feature spread: dominant lead story + right-hand notes.
                draw_panel(42, 78, 430, 315, i)
                pdf.setFillColor(accent); pdf.setFont("PrezentSans", 10)
                pdf.drawString(60, 365, f"{deck_label(data.get('language', 'uk'), 'feature')} / 01")
                feature_style = ParagraphStyle("mag-feature", parent=body_style, fontSize=18, leading=23, textColor=text_color)
                para(pts[0], 62, 340, 385, feature_style, 245)
                rest = pts[1:]
                slot_h = 315 / max(1, len(rest))
                for j, point in enumerate(rest):
                    y = 78 + 315 - (j + 1) * slot_h
                    pdf.setFillColor(accent); pdf.setFont("PrezentSans", 20)
                    pdf.drawString(500, y + slot_h - 30, f"0{j + 2}")
                    pdf.setStrokeColor(accent); pdf.setLineWidth(2)
                    pdf.line(540, y + 8, 540, y + slot_h - 8)
                    para(point, 555, y + slot_h - 18, page_w - 590, body_style, slot_h - 30)
            elif design == "mono":
                row_h = 315 / max(1, len(pts))
                for j, point in enumerate(pts):
                    y = 78 + 315 - (j + 1) * row_h
                    pdf.setFillColor(colors.HexColor("#161616")); pdf.rect(42, y + 6, 58, row_h - 8, fill=1, stroke=0)
                    pdf.setFillColor(colors.white); pdf.setFont("PrezentSans", 17)
                    pdf.drawCentredString(71, y + row_h / 2, f"{j + 1:02d}")
                    draw_panel(114, y + 6, page_w - 156, row_h - 8, j + i)
                    para(point, 130, y + row_h - 13, page_w - 192, body_style, row_h - 24)
            elif design == "nature":
                row_h = 315 / max(1, len(pts))
                for j, point in enumerate(pts):
                    y = 78 + 315 - (j + 1) * row_h
                    pdf.setFillColor(accent); pdf.saveState(); pdf.translate(68, y + row_h / 2); pdf.rotate(28 if j % 2 == 0 else -28); pdf.ellipse(-7, -17, 7, 17, fill=1, stroke=0); pdf.restoreState()
                    draw_panel(96, y + 7, page_w - 138, row_h - 14, j + i)
                    para(point, 114, y + row_h - 16, page_w - 180, body_style, row_h - 30)
            elif design == "luxury":
                card_w, card_h = 350, 145
                positions = [(48, 230), (page_w - 48 - card_w, 230), (48, 68), (page_w - 48 - card_w, 68)]
                for j, point in enumerate(pts):
                    px, py = positions[j]
                    draw_panel(px, py, card_w, card_h, j + i)
                    pdf.setStrokeColor(accent); pdf.setLineWidth(.8); pdf.rect(px + 7, py + 7, card_w - 14, card_h - 14, fill=0, stroke=1)
                    pdf.setFillColor(accent); pdf.setFont("PrezentSans", 19); pdf.drawString(px + 18, py + card_h - 30, f"0{j + 1}")
                    para(point, px + 58, py + card_h - 25, card_w - 78, body_style, 95)
            elif design == "minimal":
                row_h = 315 / max(1, len(pts))
                for j, point in enumerate(pts):
                    y = 78 + 315 - (j + 1) * row_h
                    pdf.setFillColor(accent); pdf.setFont("PrezentSans", 23)
                    pdf.drawString(58, y + row_h - 39, f"{j + 1:02d}")
                    pdf.setStrokeColor(muted); pdf.setLineWidth(.6)
                    pdf.line(112, y + 10, 112, y + row_h - 10)
                    para(point, 132, y + row_h - 17, page_w - 190, body_style, row_h - 28)
                    if j < len(pts) - 1:
                        pdf.line(132, y + 4, page_w - 48, y + 4)
            elif design == "business":
                card_w, card_h = 360, 145
                positions = [(42, 230), (page_w - 42 - card_w, 230), (42, 68), (page_w - 42 - card_w, 68)]
                for j, point in enumerate(pts):
                    px, py = positions[j]
                    draw_panel(px, py, card_w, card_h, j + i)
                    pdf.setFillColor(accent); pdf.rect(px, py + card_h - 7, card_w, 7, fill=1, stroke=0)
                    pdf.setFillColor(accent); pdf.setFont("PrezentSans", 9)
                    pdf.drawString(px + 18, py + card_h - 30, f"{deck_label(data.get('language', 'uk'), 'insight')}  0{j + 1}")
                    para(point, px + 18, py + card_h - 48, card_w - 36, body_style, 88)
            elif design == "future":
                card_w, card_h = 360, 145
                positions = [(42, 230), (page_w - 42 - card_w, 230), (42, 68), (page_w - 42 - card_w, 68)]
                for j, point in enumerate(pts):
                    px, py = positions[j]
                    draw_panel(px, py, card_w, card_h, j + i)
                    pdf.setStrokeColor(accent); pdf.setLineWidth(1.8)
                    pdf.circle(px + 32, py + card_h - 32, 11, fill=0, stroke=1)
                    pdf.setFillColor(accent); pdf.setFont("PrezentSans", 10)
                    pdf.drawCentredString(px + 32, py + card_h - 35, f"{j + 1}")
                    para(point, px + 58, py + card_h - 24, card_w - 78, body_style, 95)
            elif design == "school":
                card_w, card_h = 350, 142
                positions = [(42, 238), (page_w - 42 - card_w, 218), (60, 64), (page_w - 60 - card_w, 82)]
                for j, point in enumerate(pts):
                    px, py = positions[j]
                    draw_panel(px, py, card_w, card_h, j + i)
                    pdf.setFillColor(colors.HexColor("#FFC447")); pdf.circle(px + 30, py + card_h - 30, 10, fill=1, stroke=0)
                    pdf.setFillColor(text_color); pdf.setFont("PrezentSans", 10)
                    pdf.drawCentredString(px + 30, py + card_h - 34, str(j + 1))
                    para(point, px + 52, py + card_h - 24, card_w - 68, body_style, 92)
            elif design == "classic":
                card_w, card_h = 350, 145
                positions = [(48, 230), (page_w - 48 - card_w, 230), (48, 68), (page_w - 48 - card_w, 68)]
                for j, point in enumerate(pts):
                    px, py = positions[j]
                    draw_panel(px, py, card_w, card_h, j + i)
                    pdf.setStrokeColor(accent); pdf.setLineWidth(.7)
                    pdf.rect(px + 8, py + 8, card_w - 16, card_h - 16, fill=0, stroke=1)
                    pdf.setFillColor(accent); pdf.setFont("PrezentSans", 17)
                    pdf.drawString(px + 20, py + card_h - 33, f"0{j + 1}  /")
                    para(point, px + 84, py + card_h - 24, card_w - 105, body_style, 94)
            else:
                card_w, card_h = 360, 145
                if design == "vibrant":
                    positions = [(32, 242), (page_w - 30 - card_w, 218), (55, 65), (page_w - 45 - card_w, 88)]
                else:
                    positions = [(42, 230), (page_w - 42 - card_w, 230), (42, 68), (page_w - 42 - card_w, 68)]
                for j, point in enumerate(pts):
                    px, py = positions[j]
                    draw_panel(px, py, card_w, card_h, j + i)
                    pdf.setFillColor(accent); pdf.setFont("PrezentSans", 15)
                    pdf.drawString(px + 18, py + card_h - 28, f"{j + 1:02d}")
                    para(point, px + 20, py + card_h - 48, card_w - 42, body_style, 95)
        elif layout == "steps":
            pts = points[:5]
            row_h = 62
            start_y = 386
            for j, point in enumerate(pts):
                y = start_y - j * 68
                pdf.setFillColor(accent); pdf.circle(58, y - 23, 15, fill=1, stroke=0)
                pdf.setFillColor(bg if design != "mono" else colors.white)
                pdf.setFont("PrezentSans", 10); pdf.drawCentredString(58, y - 27, str(j + 1))
                draw_panel(88, y - row_h, page_w - 135, row_h, j + i)
                style = ParagraphStyle("stepline", parent=body_style, fontSize=11.0, leading=13.8)
                para(point, 106, y - 10, page_w - 175, style, 46)
        elif layout == "quote":
            highlight = item.get("highlight") or (points[0] if points else item["title"])
            draw_panel(64, 205, page_w - 128, 165, i)
            para("“" + highlight + "”", 95, 330, page_w - 190, focus_style, 110)
            draw_bullets(_supporting_points(item, 4), 95, 175, page_w - 190, 75, 11.0)
        elif layout == "focus":
            highlight = item.get("highlight") or (points[0] if points else item["title"])
            para(highlight, 75, 385, page_w - 150, focus_style, 100)
            pdf.setStrokeColor(accent); pdf.setLineWidth(1.7)
            pdf.line(140, 275, page_w - 140, 275)
            draw_panel(70, 75, page_w - 140, 165, i)
            draw_bullets(_supporting_points(item, 4), 95, 215, page_w - 190, 90, 11.0)
        elif layout == "closing":
            draw_panel(58, 82, page_w - 116, 315, i)
            para(item["title"], 88, 365, page_w - 176, focus_style, 68)
            draw_bullets(points[:4], 110, 285, page_w - 220, 125, 11.0)
            refs = data.get("reference_lines") or []
            if item.get("highlight") and not refs:
                para(item["highlight"], 105, 125, page_w - 210, card_title_style, 45)
            if refs:
                display_refs = [compact_reference_line(line) for line in refs[:8]]
                refs_text = "<br/>".join(html.escape(line) for line in display_refs)
                refs_style = ParagraphStyle("refs-readable", parent=body_style, fontSize=11.0, leading=13.0)
                para(deck_label(data.get("language", "uk"), "sources") + "<br/>" + refs_text, 65, 150, page_w - 130, refs_style, 105)
        else:  # split — left editorial text column, right visual anchor
            draw_panel(42, 78, 490, 315, i)
            draw_bullets(points[:4], 64, 360, 450, 105, 11.2)
            if design in ("future", "vibrant"):
                pdf.setStrokeColor(accent); pdf.setLineWidth(2)
                pdf.circle(665, 245, 76, fill=0, stroke=1)
                pdf.setStrokeColor(colors.HexColor("#6B5BFF" if design == "future" else "#FF5B91"))
                pdf.circle(665, 245, 46, fill=0, stroke=1)
                pdf.setFillColor(accent); pdf.circle(665, 245, 10, fill=1, stroke=0)
            elif design == "magazine":
                pdf.setFillColor(accent); pdf.setFont("PrezentSans", 86); pdf.drawCentredString(665, 210, f"{i + 1:02d}")
                pdf.rect(570, 180, 190, 5, fill=1, stroke=0)
            elif design == "nature":
                pdf.saveState(); pdf.translate(670, 235); pdf.rotate(28 + rng.randint(-8, 8)); pdf.setFillColor(accent); pdf.ellipse(-35, -90, 35, 90, fill=1, stroke=0); pdf.restoreState()
            elif design == "mono":
                pdf.setFillColor(text_color); pdf.rect(585, 155, 150, 150, fill=1, stroke=0)
                pdf.setFillColor(colors.white); pdf.setFont("PrezentSans", 38); pdf.drawCentredString(660, 207, f"{i + 1:02d}")
            else:
                pdf.setFillColor(accent); pdf.rect(590, 285, 145, 4, fill=1, stroke=0)
                takeaway = item.get("highlight") or item.get("kicker") or deck_label(data.get("language", "uk"), "key_idea")
                para(takeaway, 590, 260, 155, card_title_style, 120)
            if item.get("highlight"):
                para(item["highlight"], 560, 125, 210, card_title_style, 70)
        if item.get("fact") or item.get("source"):
            source_line = str(item.get("fact") or "")[:125]
            if item.get("source"):
                source_line = (source_line + " · " if source_line else "") + f"{deck_label(data.get('language', 'uk'), 'source')} {item.get('source')}"
            pdf.setFillColor(muted)
            source_style = ParagraphStyle("fact-source-footer", parent=body_style, fontSize=10.0, leading=11.8, textColor=muted)
            para(source_line, 45, 30, page_w - 90, source_style, 20, min_font_size=9.5)
        pdf.showPage()
    pdf.save()

 
# ============================================================
# 9. USER FLOW AND MENUS
# ============================================================
async def cancel_active_generation(user_id: int):
    task = ACTIVE_GENERATIONS.get(user_id)
    current = asyncio.current_task()
    if task and not task.done() and task is not current:
        log.info("Cancellation requested for user %s", user_id)
        task.cancel()
 
 
async def show_language_picker(message: Message, language: str | None = None):
    if language in ("uk", "ru", "en"):
        prompt = tr(language, "choose_language")
    else:
        prompt = (
            "👋 <b>Вітаю у Prezent AI!</b> Обери мову спілкування.\n"
            "👋 <b>Добро пожаловать в Prezent AI!</b> Выбери язык общения.\n"
            "👋 <b>Welcome to Prezent AI!</b> Choose your conversation language."
        )
    await message.answer(prompt, reply_markup=interface_language_keyboard(language))


async def show_home(message: Message, state: FSMContext, user=None):
    actual_user = user or message.from_user
    if not actual_user:
        return
    await cancel_active_generation(actual_user.id)
    await state.clear()
    await asyncio.to_thread(register_user, actual_user)
    language = get_user_language(actual_user.id)
    if not language:
        await show_language_picker(message)
        return
    welcome = tr(language, "welcome", name=html.escape(actual_user.first_name or "friend"))
    await message.answer(welcome, reply_markup=main_keyboard(language))


async def start_creation(message: Message, state: FSMContext, user=None):
    actual_user = user or message.from_user
    if not actual_user:
        return
    await cancel_active_generation(actual_user.id)
    await state.clear()
    await asyncio.to_thread(register_user, actual_user)
    language = get_user_language(actual_user.id)
    if not language:
        await show_language_picker(message)
        return
    await state.set_state(PresentationForm.topic)
    await message.answer(
        tr(language, "topic_prompt"),
        reply_markup=main_keyboard(language),
    )


@dp.message(CommandStart())
async def command_start(message: Message, state: FSMContext):
    await show_home(message, state)


@dp.message(Command("language"))
async def command_language(message: Message, state: FSMContext):
    await asyncio.to_thread(register_user, message.from_user)
    await state.clear()
    await cancel_active_generation(message.from_user.id)
    language = get_user_language(message.from_user.id)
    await show_language_picker(message, language)


@dp.callback_query(F.data.startswith("ui_lang:"))
async def interface_language_handler(callback: CallbackQuery, state: FSMContext):
    language = callback.data.split(":", 1)[1]
    if language not in ("uk", "ru", "en"):
        await callback.answer("Unsupported language", show_alert=True)
        return
    await asyncio.to_thread(register_user, callback.from_user)
    await asyncio.to_thread(set_user_language, callback.from_user.id, language)
    await state.clear()
    await callback.answer(tr(language, "language_saved").replace("<b>", "").replace("</b>", ""))
    await dismiss_choice_message(callback)
    await callback.message.answer(
        tr(language, "welcome", name=html.escape(callback.from_user.first_name or "friend")),
        reply_markup=main_keyboard(language),
    )


@dp.message(Command("restart"))
async def command_restart(message: Message, state: FSMContext):
    """A restart always asks the user to confirm/select the conversation language again."""
    await cancel_active_generation(message.from_user.id)
    await state.clear()
    await asyncio.to_thread(register_user, message.from_user)
    previous_language = get_user_language(message.from_user.id)
    await show_language_picker(message, previous_language)


@dp.message(Command("cancel"))
async def command_cancel(message: Message, state: FSMContext):
    await cancel_active_generation(message.from_user.id)
    await state.clear()
    await asyncio.to_thread(register_user, message.from_user)
    language = get_user_language(message.from_user.id)
    if not language:
        await show_language_picker(message)
        return
    await message.answer(tr(language, "action_cancelled"), reply_markup=main_keyboard(language))


@dp.message(Command("id"))
async def command_id(message: Message):
    await asyncio.to_thread(register_user, message.from_user)
    language = get_user_language(message.from_user.id)
    if not language:
        await show_language_picker(message)
        return
    texts = {
        "uk": f"Твій Telegram ID: <code>{message.from_user.id}</code>\nДодай його в Render → Environment → ADMIN_IDS.",
        "ru": f"Твой Telegram ID: <code>{message.from_user.id}</code>\nДобавь его в Render → Environment → ADMIN_IDS.",
        "en": f"Your Telegram ID: <code>{message.from_user.id}</code>\nAdd it to Render → Environment → ADMIN_IDS.",
    }
    await message.answer(texts[language], reply_markup=main_keyboard(language))

@dp.message(Command("adminstats"))
async def command_adminstats(message: Message):
    if not is_admin(message.from_user.id):
        await message.answer("⛔ Немає доступу.")
        return
    users, gens, pending, revenue = await asyncio.to_thread(get_admin_stats_db)
    await message.answer(f"🛠 <b>Статистика</b>\nКористувачів: {users}\nГотових презентацій: {gens}\nСкарг очікує: {pending}\nОтримано Stars: {revenue} ⭐")
 
@dp.message(Command("delivery_unknown"))
async def command_delivery_unknown(message: Message):
    if not message.from_user or not is_admin(message.from_user.id):
        await message.answer("⛔ Немає доступу.")
        return
    rows = await asyncio.to_thread(list_delivery_unknown_generations, 10)
    if not rows:
        await message.answer("✅ Немає презентацій із невизначеним результатом відправлення.")
        return
    lines = ["⚠️ <b>Відправлення, які потребують ручної перевірки</b>"]
    for row in rows:
        topic = html.escape(str(row["topic"] or "")[:90])
        lines.append(
            f"\nID: <code>{html.escape(str(row['generation_id']))}</code>\n"
            f"Користувач: <code>{int(row['user_id'])}</code> • Слайдів: {int(row['slides'])}\n"
            f"Тема: {topic}\nСтворено: <code>{html.escape(str(row['created_at'])[:25])}</code>"
        )
    lines.append(
        "\nЩоб підтвердити доставку: <code>/resolve_delivery ID delivered</code>\n"
        "Щоб повернути квоту після перевірки: <code>/resolve_delivery ID refund</code>"
    )
    await message.answer("\n".join(lines))

@dp.message(Command("resolve_delivery"))
async def command_resolve_delivery(message: Message):
    if not message.from_user or not is_admin(message.from_user.id):
        await message.answer("⛔ Немає доступу.")
        return
    parts = (message.text or "").split()
    if len(parts) != 3 or not re.fullmatch(r"[a-fA-F0-9]{16}", parts[1]) or parts[2] not in ("delivered", "refund"):
        await message.answer(
            "Формат: <code>/resolve_delivery ID delivered</code> або "
            "<code>/resolve_delivery ID refund</code>. Спочатку переглянь /delivery_unknown."
        )
        return
    try:
        result = await asyncio.to_thread(
            resolve_delivery_unknown, parts[1].lower(), message.from_user.id, parts[2]
        )
    except Exception as exc:
        log.exception("Admin delivery resolution failed")
        await message.answer(f"❌ Не вдалося змінити статус: <code>{html.escape(str(exc)[:250])}</code>")
        return
    if not result:
        await message.answer("ℹ️ Цю генерацію вже вирішено або її не знайдено в delivery_unknown.")
        return
    if parts[2] == "delivered":
        await message.answer(
            f"✅ Відправлення <code>{html.escape(parts[1])}</code> підтверджено як доставлене. Квота залишається використаною."
        )
        user_text = "✅ Адміністратор перевірив операцію та підтвердив, що PowerPoint було доставлено."
    else:
        await message.answer(
            f"✅ Для <code>{html.escape(parts[1])}</code> повернуто денний ліміт або 1 кредит. Рішення записане в історію."
        )
        user_text = "✅ Адміністратор перевірив невизначене відправлення та повернув квоту за цю генерацію."
    try:
        await bot.send_message(result["user_id"], user_text)
    except Exception:
        log.warning("Could not notify user %s about delivery resolution", result["user_id"])


@dp.message(Command("grant"))
async def command_grant(message: Message):
    if not is_admin(message.from_user.id):
        await message.answer("⛔ Немає доступу.")
        return
    parts = (message.text or "").split()
    if len(parts) != 3 or not parts[1].isdigit() or not parts[2].isdigit():
        await message.answer("Формат: <code>/grant ID КІЛЬКІСТЬ</code> (кількість має бути додатною)")
        return
    uid, amount = int(parts[1]), int(parts[2])
    if amount <= 0:
        await message.answer("❌ Кількість кредитів має бути більшою за нуль.")
        return
    await asyncio.to_thread(grant_credits_db, uid, amount)
    schedule_balance_backup("адміністратор змінив баланс через /grant")
    await message.answer(f"✅ Користувачу <code>{uid}</code> додано кредитів: {amount}")
 
@dp.message(Command("backup"))
async def command_backup(message: Message):
    if not is_admin(message.from_user.id):
        await message.answer("⛔ Немає доступу.")
        return
    await backup_to_admins("вручну")
    await message.answer("💾 Копію бази надіслано адміністраторам.")
 
# A document can be restored either by putting /restore in its caption,
# or by sending the document first and then sending /restore as a separate message.
# Pending file IDs live in memory for 15 minutes; the actual database remains in Turso.
PENDING_RESTORE_DOCUMENTS = {}
RESTORE_PENDING_TTL_SECONDS = 15 * 60


def _is_restore_caption(caption: str | None) -> bool:
    if not caption:
        return False
    command = caption.strip().split()[0].split("@", 1)[0].lower()
    return command == "/restore"


async def _restore_backup_document(message: Message, document):
    if not message.from_user or not is_admin(message.from_user.id):
        await message.answer("⛔ Відновлення бази доступне лише адміністратору.")
        return

    if any(task and not task.done() for task in ACTIVE_GENERATIONS.values()):
        await message.answer("⚠️ Спочатку дочекайся завершення всіх генерацій, потім повтори /restore.")
        return
    try:
        if await asyncio.to_thread(has_active_generation_leases):
            await message.answer("⚠️ У базі ще є активна генерація на одному з процесів. Дочекайся завершення та повтори /restore.")
            return
    except Exception:
        log.exception("Could not verify active generation locks before restore")
        await message.answer("❌ Не вдалося перевірити активні генерації. Відновлення зупинено для безпеки.")
        return
    file_name = (getattr(document, "file_name", None) or "backup.sqlite3").lower()
    if not file_name.endswith((".sqlite3", ".sqlite", ".db", ".db3")):
        await message.answer("❌ Надішли резервну копію SQLite у форматі .sqlite3, .sqlite або .db.")
        return
    if getattr(document, "file_size", None) and document.file_size > MAX_RESTORE_BACKUP_BYTES:
        await message.answer("❌ Резервна копія завелика. Максимальний розмір — 45 МБ.")
        return

    await message.answer("⏳ Перевіряю цілісність резервної копії, створюю контрольну копію поточної бази та відновлюю дані.")
    try:
        with tempfile.TemporaryDirectory() as tmp:
            dest = str(Path(tmp) / "restore.sqlite3")
            await bot.download(document, destination=dest)
            if Path(dest).stat().st_size > MAX_RESTORE_BACKUP_BYTES:
                raise ValueError("Резервна копія перевищує 45 МБ.")
            await asyncio.to_thread(validate_backup_file, dest)
            await asyncio.to_thread(restore_db_from, dest)
        await asyncio.to_thread(init_db)
        await asyncio.to_thread(load_user_language_cache)
        users_count, total_credits = await asyncio.to_thread(get_database_summary)
        log.info("Database restore completed by admin=%s users=%s total_credits=%s", message.from_user.id, users_count, total_credits)
        await message.answer(
            "✅ <b>Базу відновлено.</b>\n\n"
            f"Користувачів: <b>{users_count}</b>\n"
            f"Невикористаних кредитів усього: <b>{total_credits}</b>\n\n"
            "Перевір профілі та PRO-підписки перед відновленням продажів."
        )
    except Exception as exc:
        log.exception("Restore failed")
        await message.answer(
            "❌ <b>Не вдалося відновити базу.</b>\n"
            "Перевір, що це справжній файл prezentai_backup.sqlite3, і переглянь Render Runtime Logs.\n"
            f"Причина: <code>{html.escape(str(exc)[:300])}</code>"
        )


@dp.message(F.document)
async def restore_backup_document_handler(message: Message):
    if not message.from_user or not is_admin(message.from_user.id):
        return

    if _is_restore_caption(message.caption):
        await _restore_backup_document(message, message.document)
        return

    PENDING_RESTORE_DOCUMENTS[message.from_user.id] = {
        "file_id": message.document.file_id,
        "file_name": message.document.file_name or "backup.sqlite3",
        "file_size": int(message.document.file_size or 0),
        "received_at": time.time(),
        "chat_id": message.chat.id,
    }
    await message.answer(
        "📎 Файл резервної копії отримано.\n\n"
        "Тепер надішли окремим повідомленням команду <code>/restore</code> протягом 15 хвилин.\n"
        "Або надішли файл ще раз, одразу додавши <code>/restore</code> у поле підпису (Caption)."
    )


@dp.message(Command("restore"))
async def restore_backup_command_handler(message: Message):
    if not message.from_user or not is_admin(message.from_user.id):
        await message.answer("⛔ Відновлення бази доступне лише адміністратору.")
        return

    if any(task and not task.done() for task in ACTIVE_GENERATIONS.values()):
        await message.answer("⚠️ Спочатку дочекайся завершення всіх генерацій, потім повтори /restore.")
        return
    try:
        if await asyncio.to_thread(has_active_generation_leases):
            await message.answer("⚠️ У базі ще є активна генерація на одному з процесів. Дочекайся завершення та повтори /restore.")
            return
    except Exception:
        log.exception("Could not verify active generation locks before /restore")
        await message.answer("❌ Не вдалося перевірити активні генерації. Відновлення зупинено для безпеки.")
        return
    pending = PENDING_RESTORE_DOCUMENTS.get(message.from_user.id)
    if not pending:
        await message.answer(
            "Спочатку надішли файл <code>prezentai_backup.sqlite3</code> як документ, "
            "а потім окремим повідомленням команду <code>/restore</code>.\n\n"
            "Також можна надіслати файл із підписом <code>/restore</code>."
        )
        return

    pending_name = str(pending.get("file_name") or "backup.sqlite3").lower()
    if not pending_name.endswith((".sqlite3", ".sqlite", ".db", ".db3")):
        PENDING_RESTORE_DOCUMENTS.pop(message.from_user.id, None)
        await message.answer("❌ Файл має бути резервною копією SQLite (.sqlite3, .sqlite або .db).")
        return
    if pending.get("file_size", 0) > MAX_RESTORE_BACKUP_BYTES:
        PENDING_RESTORE_DOCUMENTS.pop(message.from_user.id, None)
        await message.answer("❌ Резервна копія завелика. Максимальний розмір — 45 МБ.")
        return
    if time.time() - pending.get("received_at", 0) > RESTORE_PENDING_TTL_SECONDS:
        PENDING_RESTORE_DOCUMENTS.pop(message.from_user.id, None)
        await message.answer("⌛ Час очікування файлу минув. Надішли резервну копію ще раз.")
        return

    PENDING_RESTORE_DOCUMENTS.pop(message.from_user.id, None)
    try:
        file = await bot.get_file(pending["file_id"])
        with tempfile.TemporaryDirectory() as tmp:
            dest = str(Path(tmp) / "restore.sqlite3")
            await bot.download_file(file.file_path, destination=dest)
            if Path(dest).stat().st_size > MAX_RESTORE_BACKUP_BYTES:
                raise ValueError("Резервна копія перевищує 45 МБ.")
            await asyncio.to_thread(validate_backup_file, dest)
            await asyncio.to_thread(restore_db_from, dest)
        await asyncio.to_thread(init_db)
        await asyncio.to_thread(load_user_language_cache)
        users_count, total_credits = await asyncio.to_thread(get_database_summary)
        log.info("Database restore completed by admin=%s users=%s total_credits=%s", message.from_user.id, users_count, total_credits)
        await message.answer(
            "✅ <b>Базу відновлено.</b>\n\n"
            f"Користувачів: <b>{users_count}</b>\n"
            f"Невикористаних кредитів усього: <b>{total_credits}</b>\n\n"
            "Перевір профілі та PRO-підписки перед відновленням продажів."
        )
    except Exception as exc:
        log.exception("Restore command failed")
        await message.answer(
            "❌ <b>Не вдалося відновити базу.</b>\n"
            "Перевір формат backup-файлу та Render Runtime Logs.\n"
            f"Причина: <code>{html.escape(str(exc)[:300])}</code>"
        )
 
@dp.message(F.text.in_(tuple(MENU_ACTIONS.keys())))
async def menu_button_router(message: Message, state: FSMContext):
    action = MENU_ACTIONS.get(message.text or "")
    language = get_user_language(message.from_user.id)
    if not language:
        await asyncio.to_thread(register_user, message.from_user)
        await show_language_picker(message)
        return
    if action in ("start", "home"):
        await show_home(message, state)
    elif action == "restart":
        await command_restart(message, state)
    elif action == "create":
        await start_creation(message, state)
    elif action == "cancel":
        await command_cancel(message, state)
    elif action == "features":
        await message.answer(tr(language, "features"), reply_markup=main_keyboard(language))
    elif action == "help":
        await message.answer(tr(language, "help"), reply_markup=main_keyboard(language))
    elif action == "profile":
        await asyncio.to_thread(register_user, message.from_user)
        await message.answer(await asyncio.to_thread(profile_text, message.from_user.id, language), reply_markup=main_keyboard(language))
    elif action == "shop":
        await message.answer(
            tr(language, "shop", credits=CREDIT_PACK_SIZE, credit_price=CREDIT_PACK_PRICE,
               pro_limit=PRO_DAILY_LIMIT, pro_days=PRO_DAYS, pro_price=PRO_PRICE),
            reply_markup=shop_keyboard(language),
        )
    elif action == "sub":
        await message.answer(
            tr(language, "subscription", limit=PRO_DAILY_LIMIT, days=PRO_DAYS, price=PRO_PRICE),
            reply_markup=InlineKeyboardMarkup(inline_keyboard=[[
                InlineKeyboardButton(
                    text={"uk": f"👑 Оформити PRO — {PRO_PRICE} ⭐", "ru": f"👑 Подключить PRO — {PRO_PRICE} ⭐", "en": f"👑 Activate PRO — {PRO_PRICE} ⭐"}[language],
                    callback_data="buy:pro",
                )
            ]]),
        )


@dp.message(Command("shop"))
async def command_shop(message: Message):
    await asyncio.to_thread(register_user, message.from_user)
    language = get_user_language(message.from_user.id)
    if not language:
        await show_language_picker(message)
        return
    await message.answer(
        tr(language, "shop", credits=CREDIT_PACK_SIZE, credit_price=CREDIT_PACK_PRICE,
           pro_limit=PRO_DAILY_LIMIT, pro_days=PRO_DAYS, pro_price=PRO_PRICE),
        reply_markup=shop_keyboard(language),
    )

@dp.callback_query(F.data == "profile")
async def callback_profile(callback: CallbackQuery):
    await asyncio.to_thread(register_user, callback.from_user)
    language = get_user_language(callback.from_user.id) or "uk"
    await callback.message.answer(await asyncio.to_thread(profile_text, callback.from_user.id, language), reply_markup=main_keyboard(language))
    await callback.answer()
 
@dp.callback_query(F.data == "create")
async def inline_create(callback: CallbackQuery, state: FSMContext):
    await callback.answer()
    await start_creation(callback.message, state, user=callback.from_user)
 
# ============================================================
# 10. SHOP AND TELEGRAM STARS PAYMENTS
# ============================================================
@dp.callback_query(F.data.startswith("buy:"))
async def buy_handler(callback: CallbackQuery):
    await asyncio.to_thread(register_user, callback.from_user)
    if RENDER_IS_HOSTED and not DURABLE_DB_ENABLED:
        language = get_user_language(callback.from_user.id) or "uk"
        warning = {
            "uk": "🔒 Оплату тимчасово вимкнено, щоб не втратити куплені кредити. Спочатку адміністратор має підключити постійну базу Turso.",
            "ru": "🔒 Оплата временно отключена, чтобы не потерять купленные кредиты. Сначала администратор должен подключить постоянную базу Turso.",
            "en": "🔒 Payments are temporarily disabled to protect purchased credits. The administrator must connect durable Turso storage first.",
        }[language]
        await callback.answer(warning, show_alert=True)
        return
    language = get_user_language(callback.from_user.id)
    if not language:
        await callback.answer("Choose a language first with /start.", show_alert=True)
        return
    product = callback.data.split(":", 1)[1]
    if product == "credits":
        title = {"uk": "5 генерацій Prezent AI", "ru": "5 генераций Prezent AI", "en": "5 Prezent AI generations"}[language]
        description = {"uk": "П'ять додаткових генерацій презентацій.", "ru": "Пять дополнительных генераций презентаций.", "en": "Five extra presentation generations."}[language]
        amount = CREDIT_PACK_PRICE
    elif product == "pro":
        title = {"uk": "Prezent AI PRO на 30 днів", "ru": "Prezent AI PRO на 30 дней", "en": "Prezent AI PRO for 30 days"}[language]
        description = {"uk": "До 10 презентацій на день протягом 30 днів.", "ru": "До 10 презентаций в день в течение 30 дней.", "en": "Up to 10 presentations per day for 30 days."}[language]
        amount = PRO_PRICE
    else:
        await callback.answer({"uk": "Невідомий товар.", "ru": "Неизвестный товар.", "en": "Unknown product."}[language], show_alert=True)
        return
    payload = await asyncio.to_thread(create_order, callback.from_user.id, product, amount)
    try:
        await bot.send_invoice(
            chat_id=callback.from_user.id,
            title=title,
            description=description,
            payload=payload,
            provider_token="",
            currency="XTR",
            prices=[LabeledPrice(label=title, amount=amount)],
        )
        await callback.answer({"uk": "Рахунок надіслано нижче.", "ru": "Счёт отправлен ниже.", "en": "Invoice sent below."}[language])
    except Exception:
        log.exception("Could not send invoice")
        await callback.answer({"uk": "Не вдалося створити рахунок. Спробуй ще раз.", "ru": "Не удалось создать счёт. Попробуй ещё раз.", "en": "Could not create the invoice. Please try again."}[language], show_alert=True)


@dp.pre_checkout_query()
async def pre_checkout_handler(query):
    # Telegram requires a response to every pre-checkout query within 10 seconds.
    language = get_user_language(query.from_user.id) or "uk"
    if RENDER_IS_HOSTED and not DURABLE_DB_ENABLED:
        await query.answer(
            ok=False,
            error_message="Payments are temporarily disabled until permanent storage is connected.",
        )
        return
    errors = {
        "uk": ("Замовлення не знайдено або його параметри не збігаються. Відкрий магазин і створи новий рахунок.", "Тимчасова помилка перевірки замовлення. Спробуй створити рахунок ще раз."),
        "ru": ("Заказ не найден или его параметры не совпадают. Открой магазин и создай новый счёт.", "Временная ошибка проверки заказа. Попробуй создать счёт ещё раз."),
        "en": ("Order not found or its details do not match. Open the shop and create a new invoice.", "Temporary order validation error. Please create the invoice again."),
    }
    try:
        order = await asyncio.to_thread(find_order, query.invoice_payload)
        ok = bool(
            order
            and int(order["user_id"]) == query.from_user.id
            and int(order["amount"]) == query.total_amount
            and query.currency == "XTR"
            and order["status"] == "pending"
        )
        await query.answer(ok=ok, error_message=None if ok else errors[language][0])
    except Exception:
        log.exception("Could not validate Telegram Stars pre-checkout")
        try:
            await query.answer(ok=False, error_message=errors[language][1])
        except Exception:
            log.exception("Could not reject failed pre-checkout query")


@dp.message(F.successful_payment)
async def successful_payment_handler(message: Message):
    payment = message.successful_payment
    language = get_user_language(message.from_user.id) or "uk"

    # A pre-deploy invoice could theoretically complete just as storage is
    # being reconfigured. Never credit an ephemeral Render database.
    if RENDER_IS_HOSTED and not DURABLE_DB_ENABLED:
        refund_ok = False
        try:
            refund_ok = await bot.refund_star_payment(
                user_id=message.from_user.id,
                telegram_payment_charge_id=payment.telegram_payment_charge_id,
            )
        except Exception:
            log.exception("Could not automatically refund a Stars payment without durable storage")
        for admin_id in ADMIN_IDS:
            try:
                await bot.send_message(
                    admin_id,
                    "🚨 <b>Платіж надійшов без постійної бази.</b> "
                    f"Користувач: <code>{message.from_user.id}</code>; "
                    f"сума: {payment.total_amount} ⭐; "
                    f"автоповернення: {'успішне' if refund_ok else 'НЕ вдалося'}; "
                    f"charge: <code>{html.escape(payment.telegram_payment_charge_id)}</code>."
                )
            except Exception:
                log.exception("Could not notify admin about a payment requiring refund")
        if refund_ok:
            user_text = {
                "uk": "🔒 Постійне сховище тимчасово недоступне. Оплату автоматично повернуто в Telegram Stars. Спробуй пізніше.",
                "ru": "🔒 Постоянное хранилище временно недоступно. Оплата автоматически возвращена в Telegram Stars. Попробуй позже.",
                "en": "🔒 Durable storage is temporarily unavailable. Your payment was automatically refunded in Telegram Stars. Please try again later.",
            }[language]
        else:
            user_text = {
                "uk": "⚠️ Оплату отримано, але сховище тимчасово недоступне. Адміністратора сповіщено для ручного повернення. Не повторюй платіж.",
                "ru": "⚠️ Оплата получена, но хранилище временно недоступно. Администратор уведомлён для ручного возврата. Не повторяй платёж.",
                "en": "⚠️ Payment was received, but durable storage is unavailable. The administrator has been alerted to arrange a manual refund. Do not pay again.",
            }[language]
        await message.answer(user_text, reply_markup=main_keyboard(language))
        return

    try:
        ok, result = await asyncio.to_thread(
            process_payment, payment.invoice_payload, message.from_user.id,
            payment.total_amount, payment.currency, payment.telegram_payment_charge_id,
        )
    except Exception as storage_exc:
        # A database timeout is an unknown outcome: the remote transaction may already
        # have committed. Retry idempotently and inspect the charge before any refund.
        log.exception("Stars payment application had an uncertain storage outcome")
        ok, result = False, "storage_unavailable"
        charge_id = payment.telegram_payment_charge_id
        for attempt in range(2):
            try:
                existing = await asyncio.to_thread(find_order_by_charge, charge_id)
                if (existing and existing["status"] == "paid"
                        and existing["payload"] == payment.invoice_payload
                        and int(existing["user_id"]) == message.from_user.id):
                    ok, result = True, "already_paid"
                    break
                ok, result = await asyncio.to_thread(
                    process_payment, payment.invoice_payload, message.from_user.id,
                    payment.total_amount, payment.currency, charge_id,
                )
                if ok:
                    break
            except Exception:
                log.exception("Stars payment reconciliation attempt %s failed", attempt + 1)
                await asyncio.sleep(1.0 * (attempt + 1))

        if not ok:
            details = (
                "🚨 <b>Потрібна звірка Telegram Stars.</b> Автоматичне повернення НЕ виконувалось, "
                "оскільки результат транзакції в базі невідомий.\n"
                f"Користувач: <code>{message.from_user.id}</code>\n"
                f"Сума: <b>{payment.total_amount} ⭐</b>\n"
                f"Payload: <code>{html.escape(payment.invoice_payload)}</code>\n"
                f"Charge: <code>{html.escape(charge_id)}</code>\n"
                f"Помилка: <code>{html.escape(str(storage_exc)[:300])}</code>"
            )
            await notify_admins_once(f"payment-uncertain-{charge_id}", details, cooldown=60)
            user_text = {
                "uk": "⚠️ Telegram підтвердив оплату, але база тимчасово не відповідає. Не плати повторно. Ми перевіримо транзакцію та зарахуємо покупку або організуємо повернення.",
                "ru": "⚠️ Telegram подтвердил оплату, но база временно не отвечает. Не плати повторно. Мы проверим транзакцию и зачислим покупку или организуем возврат.",
                "en": "⚠️ Telegram confirmed the payment, but the database is temporarily unavailable. Do not pay again. We will reconcile the transaction and apply the purchase or arrange a refund.",
            }[language]
            await message.answer(user_text, reply_markup=main_keyboard(language))
            return
    if not ok:
        log.error("Payment not applied: user=%s payload=%s reason=%s", message.from_user.id, payment.invoice_payload, result)
        await message.answer(tr(language, "payment_error"), reply_markup=main_keyboard(language))
        for admin_id in ADMIN_IDS:
            try:
                await bot.send_message(admin_id, f"⚠️ Перевір платіж вручну. User: {message.from_user.id}; payload: <code>{html.escape(payment.invoice_payload)}</code>; причина: {html.escape(str(result))}")
            except Exception:
                pass
        return
    if result == "credits":
        text = tr(language, "payment_credits", credits=CREDIT_PACK_SIZE, profile=await asyncio.to_thread(profile_text, message.from_user.id, language))
    elif result == "pro":
        text = tr(language, "payment_pro", days=PRO_DAYS, limit=PRO_DAILY_LIMIT, profile=await asyncio.to_thread(profile_text, message.from_user.id, language))
    else:
        text = tr(language, "payment_already")
    await message.answer(text, reply_markup=main_keyboard(language))
    if result != "already_paid":
        for admin_id in ADMIN_IDS:
            try:
                await bot.send_message(
                    admin_id,
                    f"💰 Оплата: <code>{message.from_user.id}</code> • {payment.total_amount} ⭐ • {html.escape(str(result))}",
                )
            except Exception:
                pass
        schedule_balance_backup("після оплати")


# ============================================================
# 11. REPORTS / QUALITY REVIEW
# ============================================================
@dp.callback_query(F.data.startswith("bad:"))
async def report_start(callback: CallbackQuery, state: FSMContext):
    generation_id = callback.data.split(":", 1)[1]
    row = await asyncio.to_thread(get_generation_owner, generation_id)
    if not row or int(row["user_id"]) != callback.from_user.id:
        await callback.answer("Ця презентація не належить твоєму профілю.", show_alert=True)
        return
    await state.clear()
    await state.update_data(report_generation_id=generation_id, report_source_message_id=callback.message.message_id)
    await state.set_state(ReportForm.reason)
    language = get_user_language(callback.from_user.id) or "uk"
    await callback.message.answer(
        tr(language, "report_prompt"),
        reply_markup=main_keyboard(language),
    )
    await callback.answer({"uk": "Чекаю опис проблеми.", "ru": "Жду описание проблемы.", "en": "Waiting for your description."}[get_user_language(callback.from_user.id) or "uk"])
 
 
@dp.message(ReportForm.reason)
async def report_reason_handler(message: Message, state: FSMContext):
    language = get_user_language(message.from_user.id) or "uk"
    reason = (message.text or "").strip()
    if len(reason) < 4:
        await message.answer(tr(language, "report_too_short"))
        return
    data = await state.get_data()
    generation_id = data.get("report_generation_id")
    source_message_id = data.get("report_source_message_id")
    report_id = await asyncio.to_thread(create_report, generation_id, message.from_user.id, reason)
    await state.clear()
    if not report_id:
        await message.answer(tr(language, "report_missing"), reply_markup=main_keyboard(language))
        return
    report = await asyncio.to_thread(get_report, report_id)
    delivered = 0
    for admin_id in ADMIN_IDS:
        try:
            await bot.send_message(
                admin_id,
                "⚠️ <b>Нова скарга на презентацію</b>\n\n"
                f"👤 Користувач: <code>{message.from_user.id}</code> (@{html.escape(message.from_user.username or 'немає')})\n"
                f"🆔 Скарга: <code>{report_id}</code>\n"
                f"📌 Тема: {html.escape(report['topic'] or 'невідома')}\n"
                f"🎨 Дизайн: {html.escape(report['design'] or '')}\n"
                f"📚 Деталізація: {html.escape(detail_label_for(language, report['detail'] or 'medium'))}\n"
                f"📝 Причина: {html.escape(reason)}\n\n"
                "Переглянь файл, надісланий поруч, і вибери дію:",
                reply_markup=report_admin_keyboard(report_id),
            )
            if report["source_chat_id"] and report["source_message_id"]:
                await bot.copy_message(
                    chat_id=admin_id, from_chat_id=report["source_chat_id"],
                    message_id=report["source_message_id"],
                )
            elif source_message_id:
                await bot.copy_message(chat_id=admin_id, from_chat_id=message.chat.id, message_id=source_message_id)
            delivered += 1
        except Exception:
            log.exception("Could not notify admin about report %s", report_id)
    if delivered:
        await message.answer(tr(language, "report_sent"), reply_markup=main_keyboard(language))
    else:
        await message.answer(tr(language, "report_send_error"), reply_markup=main_keyboard(language))
 
 
@dp.callback_query(F.data.startswith("reportact:"))
async def report_action_handler(callback: CallbackQuery):
    if not is_admin(callback.from_user.id):
        await callback.answer("⛔ Ця кнопка доступна тільки адміністратору.", show_alert=True)
        return
    _, report_id, action = callback.data.split(":", 2)
    if action not in ("refund", "reject"):
        await callback.answer("Невідома дія.", show_alert=True)
        return
    result = await asyncio.to_thread(handle_report, report_id, callback.from_user.id, action)
    if not result:
        await callback.answer("Скаргу вже оброблено або не знайдено.", show_alert=True)
        return
    user_id, status = result
    user_lang = get_user_language(user_id) or "uk"
    if status == "refunded":
        schedule_balance_backup("повернення кредиту за скаргою")
        try:
            await bot.send_message(user_id, tr(user_lang, "refund_notice"), reply_markup=main_keyboard(user_lang))
        except Exception:
            log.warning("Could not notify user %s about credit refund", user_id)
        await callback.message.edit_reply_markup(reply_markup=None)
        await callback.message.answer(f"✅ Кредит повернуто користувачу <code>{user_id}</code>.")
        await callback.answer("Повернено 1 кредит.")
    else:
        try:
            await bot.send_message(user_id, tr(user_lang, "reject_notice"), reply_markup=main_keyboard(user_lang))
        except Exception:
            pass
        await callback.message.edit_reply_markup(reply_markup=None)
        await callback.message.answer(f"Скаргу користувача <code>{user_id}</code> відхилено.")
        await callback.answer("Скаргу відхилено.")
 
# ============================================================
# 12. GENERATION FLOW
# ============================================================
@dp.message(PresentationForm.topic)
async def topic_handler(message: Message, state: FSMContext):
    language = get_user_language(message.from_user.id) or "uk"
    topic = (message.text or "").strip()
    if len(topic) < 3:
        await message.answer(tr(language, "topic_short"))
        return
    if len(topic) > MAX_TOPIC_LENGTH:
        await message.answer(tr(language, "topic_long"))
        return
    await state.update_data(topic=topic)
    await state.set_state(PresentationForm.slides)
    await message.answer(tr(language, "slides_prompt"), reply_markup=slides_keyboard())


@dp.callback_query(F.data.startswith("slides:"))
async def slides_handler(callback: CallbackQuery, state: FSMContext):
    language_ui = get_user_language(callback.from_user.id) or "uk"
    if await state.get_state() != PresentationForm.slides.state:
        await callback.answer(tr(language_ui, "slides_first"), show_alert=True)
        return
    try:
        count = int(callback.data.split(":", 1)[1])
    except (ValueError, IndexError):
        await callback.answer(tr(language_ui, "slides_invalid"), show_alert=True)
        return
    if count not in (6, 8, 10, 12):
        await callback.answer(tr(language_ui, "slides_choose"), show_alert=True)
        return
    await state.update_data(slides=count)
    await state.set_state(PresentationForm.language)
    await callback.message.answer(tr(language_ui, "output_language_prompt"), reply_markup=language_keyboard(language_ui))
    await callback.answer()
    await dismiss_choice_message(callback)
 
 
@dp.callback_query(F.data.startswith("lang:"))
async def language_handler(callback: CallbackQuery, state: FSMContext):
    language_ui = get_user_language(callback.from_user.id) or "uk"
    if await state.get_state() != PresentationForm.language.state:
        await callback.answer(tr(language_ui, "slides_first"), show_alert=True)
        return
    language = callback.data.split(":", 1)[1]
    if language not in LANGUAGES:
        await callback.answer("Unsupported language" if language_ui == "en" else ("Непідтримувана мова." if language_ui == "uk" else "Неподдерживаемый язык."), show_alert=True)
        return
    await state.update_data(language=language)
    await state.set_state(PresentationForm.detail)
    await callback.message.answer(tr(language_ui, "detail_prompt"), reply_markup=detail_keyboard(language_ui))
    await callback.answer()
    await dismiss_choice_message(callback)


@dp.callback_query(F.data.startswith("detail:"))
async def detail_handler(callback: CallbackQuery, state: FSMContext):
    language_ui = get_user_language(callback.from_user.id) or "uk"
    if await state.get_state() != PresentationForm.detail.state:
        await callback.answer(tr(language_ui, "slides_first"), show_alert=True)
        return
    detail = callback.data.split(":", 1)[1]
    if detail not in DETAIL_PROFILES:
        await callback.answer(tr(language_ui, "detail_invalid"), show_alert=True)
        return
    await state.update_data(detail=detail)
    await state.set_state(PresentationForm.audience)
    prompt = {"uk": "👥 <b>Для кого готуємо презентацію?</b>", "ru": "👥 <b>Для кого готовим презентацию?</b>", "en": "👥 <b>Who is this presentation for?</b>"}[language_ui]
    await callback.message.answer(prompt, reply_markup=audience_keyboard(language_ui))
    await callback.answer()
    await dismiss_choice_message(callback)


@dp.callback_query(F.data.startswith("audience:"))
async def audience_handler(callback: CallbackQuery, state: FSMContext):
    ui = user_language(callback.from_user.id)
    if await state.get_state() != PresentationForm.audience.state:
        await callback.answer("Choose an audience in the active flow first.", show_alert=True)
        return
    audience = callback.data.split(":", 1)[1]
    if audience not in AUDIENCE_GUIDANCE:
        await callback.answer("Unknown audience.", show_alert=True)
        return
    await state.update_data(audience=audience)
    await state.set_state(PresentationForm.photos)
    await callback.message.answer(tr(ui, "photos_prompt"), reply_markup=photos_keyboard(ui))
    await callback.answer()
    await dismiss_choice_message(callback)


@dp.callback_query(F.data.startswith("photos:"))
async def photos_handler(callback: CallbackQuery, state: FSMContext):
    language_ui = get_user_language(callback.from_user.id) or "uk"
    if await state.get_state() != PresentationForm.photos.state:
        await callback.answer(tr(language_ui, "slides_first"), show_alert=True)
        return
    choice = callback.data.split(":", 1)[1]
    if choice not in ("yes", "no"):
        await callback.answer(tr(language_ui, "photos_invalid"), show_alert=True)
        return
    await state.update_data(photos=(choice == "yes"))
    await state.set_state(PresentationForm.design)
    await callback.message.answer(tr(language_ui, "design_prompt"), reply_markup=design_keyboard(language_ui))
    await callback.answer()
    await dismiss_choice_message(callback)


@dp.callback_query(F.data.startswith("outline:"))
async def outline_review_handler(callback: CallbackQuery, state: FSMContext):
    _, action, review_id = callback.data.split(":", 2)
    review = OUTLINE_REVIEW_WAITERS.get(review_id)
    if not review or review["user_id"] != callback.from_user.id or review["future"].done():
        await callback.answer("Outline review expired.", show_alert=True)
        return
    ui = user_language(callback.from_user.id)
    if action == "approve":
        review["future"].set_result(None)
        await state.clear()
        await callback.message.edit_reply_markup(reply_markup=None)
        await callback.answer("План підтверджено." if ui == "uk" else "Outline approved.")
    elif action == "revise" and review["allow_revision"]:
        await state.update_data(outline_review_id=review_id)
        await state.set_state(PresentationForm.outline_feedback)
        text = {
            "uk": "Напиши, що змінити, і вкажи номер слайда. Наприклад: «На слайді 4 додай практичний приклад». До 800 символів.",
            "ru": "Напиши, что изменить, и укажи номер слайда. Например: «На слайде 4 добавь практический пример». До 800 символов.",
            "en": "Describe the change and include a slide number, e.g. “Add a practical example to slide 4.” Up to 800 characters.",
        }[ui]
        await callback.message.answer(text)
        await callback.answer()
    else:
        await callback.answer("Action unavailable.", show_alert=True)


@dp.message(PresentationForm.outline_feedback)
async def outline_feedback_handler(message: Message, state: FSMContext):
    feedback = (message.text or "").strip()
    ui = user_language(message.from_user.id)
    if not feedback or len(feedback) > 800:
        await message.answer("Правка має містити до 800 символів." if ui == "uk" else "Keep the revision under 800 characters.")
        return
    info = await state.get_data()
    review = OUTLINE_REVIEW_WAITERS.get(info.get("outline_review_id"))
    if not review or review["user_id"] != message.from_user.id or review["future"].done():
        await state.clear()
        await message.answer("Час перегляду минув. Почни створення ще раз." if ui == "uk" else "Outline review expired; start again.", reply_markup=main_keyboard(ui))
        return
    review["future"].set_result(feedback)
    await state.clear()
    await message.answer("Правку отримано, оновлюю структуру…" if ui == "uk" else "Revision received; updating the outline…")


@dp.callback_query(F.data == "repeat:last")
async def repeat_last_handler(callback: CallbackQuery, state: FSMContext):
    user_id = callback.from_user.id
    language_ui = get_user_language(user_id) or "uk"
    params = LAST_GENERATION_PARAMS.get(user_id)
    if not params:
        row = await asyncio.to_thread(get_latest_completed_generation, user_id)
        if row:
            params = {"topic": row["topic"], "design": row["design"], "slides": int(row["slides"]),
                      "detail": row["detail"] or "medium", "language": row["language"] or "uk",
                      "photos": True if row["photos"] is None else bool(row["photos"])}
            LAST_GENERATION_PARAMS[user_id] = params
    if not params or params.get("design") not in DESIGNS:
        await callback.answer({"uk": "Попередню презентацію не знайдено.", "ru": "Предыдущая презентация не найдена.", "en": "No previous presentation was found."}[language_ui], show_alert=True)
        return
    await state.clear()
    await state.update_data(**params, interface_language=language_ui)
    await state.set_state(PresentationForm.design)
    repeated_callback = callback.model_copy(update={"data": f"design:{params['design']}:repeat"})
    await design_handler(repeated_callback, state)


@dp.callback_query(F.data.startswith("design:"))
async def design_handler(callback: CallbackQuery, state: FSMContext):
    global GENERATION_QUEUE_WAITERS
    user_id = callback.from_user.id
    current_task = asyncio.current_task()
    active_task = ACTIVE_GENERATIONS.get(user_id)
    if active_task and active_task is not current_task and not active_task.done():
        ui = get_user_language(user_id) or "uk"
        waiting = {"uk": "Твоя презентація вже створюється.", "ru": "Твоя презентация уже создаётся.", "en": "Your presentation is already being created."}[ui]
        await callback.answer(waiting, show_alert=True)
        return

    # Claim locally before the first await. The DB lock below also protects other workers.
    ACTIVE_GENERATIONS[user_id] = current_task
    ui = get_user_language(user_id) or "uk"
    quota = None
    reserved = False
    succeeded = False
    status = None
    progress = None
    lease_task = None
    queue_registered = False
    semaphore_acquired = False
    generation_id = uuid.uuid4().hex[:16]
    generation_timeout = GENERATION_TOTAL_TIMEOUT_SECONDS
    pdf_sent = False
    history_created = False

    async def change_stage(key: str):
        if progress:
            progress.set_stage(key)
        updated = await asyncio.to_thread(set_generation_stage, generation_id, key)
        if not updated:
            raise RuntimeError(f"Could not persist generation stage: {key}")

    async def settle_failed_generation():
        if not history_created:
            return "not_reserved"
        if succeeded:
            return "delivered"
        try:
            return await asyncio.to_thread(finalize_generation, generation_id, False)
        except Exception:
            log.exception("Atomic quota settlement could not be confirmed for %s", generation_id)
            await notify_admins_once(
                f"generation-settlement-{generation_id}",
                f"🚨 <b>Не вдалося підтвердити повернення квоти.</b>\nUser: <code>{user_id}</code>\nGeneration: <code>{generation_id}</code>. Потрібна ручна звірка.",
                cooldown=60,
            )
            return "settlement_failed"

    try:
        if await state.get_state() != PresentationForm.design.state:
            await callback.answer(tr(ui, "design_prompt").replace("<b>", "").replace("</b>", ""), show_alert=True)
            return

        design_parts = callback.data.split(":")
        design = design_parts[1] if len(design_parts) > 1 else ""
        is_repeat_action = len(design_parts) > 2 and design_parts[2] == "repeat"
        if design not in DESIGNS:
            await callback.answer({"uk": "Непідтримуваний стиль.", "ru": "Неподдерживаемый стиль.", "en": "Unsupported style."}[ui], show_alert=True)
            return

        data = await state.get_data()
        data.setdefault("detail", "medium")
        data.setdefault("audience", "general")
        data.setdefault("photos", True)
        photos_enabled = bool(data["photos"])
        is_detailed = data.get("detail") == "detailed"
        # One user-requested outline revision can repeat the AI content stage.
        generation_timeout = (GENERATION_DETAILED_TIMEOUT_SECONDS if is_detailed else GENERATION_TOTAL_TIMEOUT_SECONDS) + 420
        if not all(k in data for k in ("topic", "slides", "language")):
            await state.clear()
            await callback.answer({"uk": "Налаштування втрачено. Почни спочатку.", "ru": "Настройки потеряны. Начни сначала.", "en": "Settings were lost. Please start again."}[ui], show_alert=True)
            return

        # Bound the in-memory queue before reserving any quota.
        if GEN_SEMAPHORE.locked() and GENERATION_QUEUE_WAITERS >= GENERATION_QUEUE_LIMIT:
            await callback.answer({
                "uk": "⏳ Черга зараз заповнена. Спробуй ще раз за хвилину.",
                "ru": "⏳ Очередь сейчас заполнена. Попробуй ещё раз через минуту.",
                "en": "⏳ The queue is currently full. Please try again in a minute.",
            }[ui], show_alert=True)
            return

        if not is_repeat_action:
            await dismiss_choice_message(callback)
        data["design"] = design
        data["interface_language"] = ui
        generation_record = {
            "generation_id": generation_id,
            "topic": data["topic"],
            "design": design,
            "slides": int(data["slides"]),
            "detail": data.get("detail", "medium"),
            "language": data.get("language", "uk"),
            "photos": photos_enabled,
        }
        quota = await asyncio.to_thread(reserve_generation, user_id, generation_record, WORKER_ID)
        if not quota["allowed"]:
            if quota.get("mode") == "busy":
                await callback.answer({
                    "uk": "Твоя презентація вже створюється на іншому процесі.",
                    "ru": "Твоя презентация уже создаётся в другом процессе.",
                    "en": "Your presentation is already being generated by another worker.",
                }[ui], show_alert=True)
                return
            if quota.get("mode") == "queue_full":
                await callback.answer({
                    "uk": "⏳ Черга генерацій заповнена. Спробуй трохи пізніше.",
                    "ru": "⏳ Очередь генераций заполнена. Попробуй немного позже.",
                    "en": "⏳ The generation queue is full. Please try again shortly.",
                }[ui], show_alert=True)
                return
            await callback.answer({"uk": "Денний ліміт вичерпано.", "ru": "Дневной лимит исчерпан.", "en": "Daily limit reached."}[ui], show_alert=True)
            await callback.message.answer(
                tr(ui, "limit_reached", limit=FREE_DAILY_LIMIT, credits=CREDIT_PACK_SIZE, price=CREDIT_PACK_PRICE),
                reply_markup=shop_keyboard(ui),
            )
            return

        reserved = True
        history_created = True  # quota + history + DB lock are committed in one transaction
        lease_task = asyncio.create_task(
            generation_lease_heartbeat(generation_id, user_id, WORKER_ID),
            name=f"generation-lease-{generation_id}",
        )
        if quota.get("mode") == "credit":
            schedule_balance_backup("списання придбаного кредиту")

        if GEN_SEMAPHORE.locked():
            GENERATION_QUEUE_WAITERS += 1
            queue_registered = True
        await state.clear()
        await callback.answer({"uk": "Додав у чергу на створення!", "ru": "Добавил в очередь на создание!", "en": "Added to the generation queue!"}[ui])

        progress = ProgressAnimator(
            language=ui, style=design, detail=data.get("detail", "medium"),
            slide_count=data.get("slides", 8), photos=photos_enabled,
        )
        queue_position = max(1, int(quota.get("queue_position") or GENERATION_QUEUE_WAITERS))
        queue_text = {
            "uk": f"⏳ <b>Ваша презентація в черзі.</b> Орієнтовна позиція: <b>{queue_position}</b>. Щойно звільниться місце, почнеться генерація.",
            "ru": f"⏳ <b>Ваша презентация в очереди.</b> Примерная позиция: <b>{queue_position}</b>. Как только освободится место, начнётся генерация.",
            "en": f"⏳ <b>Your presentation is in the queue.</b> Approximate position: <b>{queue_position}</b>. Generation starts as soon as a slot is free.",
        }[ui]
        if queue_registered:
            status = await callback.message.answer(queue_text)

        async with GEN_SEMAPHORE:
            semaphore_acquired = True
            if queue_registered:
                GENERATION_QUEUE_WAITERS = max(0, GENERATION_QUEUE_WAITERS - 1)
                queue_registered = False
            if status is None:
                status = await callback.message.answer(progress.render())
            else:
                await status.edit_text(progress.render(), parse_mode=ParseMode.HTML)
            progress.message = status
            progress.set_stage("writing")
            progress.start()

            # One whole-job deadline includes AI, photo search, PPTX, PDF, and Telegram delivery.
            loop = asyncio.get_running_loop()
            deadline = loop.time() + generation_timeout
            content_limit = 300.0 if is_detailed else 240.0
            await change_stage("writing")
            presentation_data = await run_stage_with_deadline(
                lambda: generate_content(data, progress_callback=progress.set_substage),
                deadline, "Gemini content generation", content_limit,
            )
            if not isinstance(presentation_data, dict) or not presentation_data.get("slides"):
                raise RuntimeError("Gemini returned an empty presentation structure")
            if len(presentation_data["slides"]) != int(data["slides"]):
                raise GeminiAPIError(
                    f"AI returned {len(presentation_data['slides'])} slides; {int(data['slides'])} were requested"
                )

            for revision_round in range(2):
                await progress.stop()
                heading = {"uk": "🗂 <b>Перевір структуру презентації</b>", "ru": "🗂 <b>Проверь структуру презентации</b>", "en": "🗂 <b>Review the presentation outline</b>"}[ui]
                rows = []
                for index, slide in enumerate(presentation_data["slides"], 1):
                    title = html.escape(str(slide.get("title") or f"Slide {index}")[:100])
                    points = slide.get("content") or []
                    sample = html.escape(str(points[0])[:80]) if points else ""
                    rows.append(f"<b>{index}. {title}</b>" + (f" — {sample}" if sample else ""))
                review_text = heading + "\n\n" + "\n".join(rows)
                review_id = f"{generation_id}-{revision_round}"
                review_future = loop.create_future()
                allow_revision = revision_round == 0
                OUTLINE_REVIEW_WAITERS[review_id] = {"user_id": user_id, "future": review_future, "allow_revision": allow_revision}
                try:
                    await status.edit_text(review_text, parse_mode=ParseMode.HTML,
                                           reply_markup=outline_review_keyboard(review_id, ui, allow_revision))
                    try:
                        revision = await asyncio.wait_for(review_future, timeout=90.0)
                    except asyncio.TimeoutError:
                        revision = None
                finally:
                    OUTLINE_REVIEW_WAITERS.pop(review_id, None)
                    await state.clear()
                if not revision:
                    break
                data["previous_outline"] = "\n".join(
                    f"{i}. {str(item.get('title') or '')[:100]}: {str((item.get('content') or [''])[0])[:80]}"
                    for i, item in enumerate(presentation_data["slides"], 1)
                )
                data["revision_request"] = revision
                await change_stage("writing")
                progress.set_substage("research")
                progress.start()
                presentation_data = await run_stage_with_deadline(
                    lambda: generate_content(data, progress_callback=progress.set_substage),
                    deadline, "Gemini outline revision", content_limit,
                )
                if not isinstance(presentation_data, dict) or len(presentation_data.get("slides", [])) != int(data["slides"]):
                    raise GeminiAPIError("AI returned an invalid revised outline")

            await status.edit_reply_markup(reply_markup=None)
            progress.set_stage("writing")
            progress.set_substage("photos" if photos_enabled else "writing")
            progress.start()
            await status.edit_text(progress.render(), parse_mode=ParseMode.HTML)
            with tempfile.TemporaryDirectory(prefix="prezent_ai_") as tmp:
                image_credits = []
                remaining = deadline - loop.time()
                if photos_enabled and remaining > 1.0:
                    progress.set_substage("photos")
                    await change_stage("photos")
                    try:
                        image_credits = await run_stage_with_deadline(
                            lambda: download_slide_images(presentation_data, tmp),
                            deadline, "photo search and download", 90.0,
                        )
                    except asyncio.TimeoutError:
                        log.warning("Photo stage timed out; continuing without late photos")
                presentation_data["reference_lines"] = _build_reference_lines(
                    presentation_data.get("research_sources", []), image_credits, data.get("language", "uk")
                )
                pptx_path = str(Path(tmp) / "PrezentAI.pptx")
                pdf_path = str(Path(tmp) / "PrezentAI.pdf")

                await change_stage("pptx")
                if progress:
                    progress.set_stage("pptx")
                await run_blocking_stage_with_deadline(
                    create_pptx, (presentation_data, design, pptx_path),
                    deadline, "PowerPoint build", 120.0,
                )
                await run_blocking_stage_with_deadline(
                    validate_pptx_file, (pptx_path, len(presentation_data.get("slides", []))),
                    deadline, "PowerPoint validation", 20.0,
                )

                pdf_ready = False
                await change_stage("pdf")
                if progress:
                    progress.set_stage("pdf")
                try:
                    font_path = await run_stage_with_deadline(
                        lambda: ensure_unicode_font(), deadline, "Unicode font lookup", 20.0,
                    )
                    await run_blocking_stage_with_deadline(
                        create_pdf, (presentation_data, design, pdf_path, font_path),
                        deadline, "PDF build", 90.0,
                    )
                    await run_blocking_stage_with_deadline(
                        validate_pdf_file, (pdf_path,), deadline, "PDF validation", 20.0,
                    )
                    pdf_ready = True
                except asyncio.TimeoutError:
                    log.warning("PDF generation timed out; PPTX will still be delivered")
                except Exception:
                    log.exception("PDF generation or validation failed; PPTX will still be delivered")

                # Persist the risky stage before upload. If the connection breaks during upload,
                # stale recovery will mark delivery_unknown rather than refunding twice.
                if loop.time() >= deadline:
                    raise asyncio.TimeoutError("Overall generation deadline reached before Telegram upload")
                await change_stage("sending")
                if progress:
                    progress.set_stage("sending")
                pptx_message = await run_stage_with_deadline(
                    lambda: callback.message.answer_document(
                        FSInputFile(pptx_path, filename="PrezentAI.pptx"),
                        caption=tr(
                            ui, "file_caption", title=html.escape(presentation_data["title"]),
                            style=html.escape(design_name(design, ui)),
                            detail=html.escape(detail_label_for(ui, data.get("detail", "medium"))),
                        ),
                        reply_markup=report_user_keyboard(generation_id, ui),
                    ),
                    deadline, "Telegram PPTX upload", 120.0,
                )
                succeeded = True  # Telegram returned a Message: the PPTX is delivered.
                try:
                    final_status = await asyncio.to_thread(
                        finalize_generation, generation_id, True,
                        pptx_message.chat.id, pptx_message.message_id,
                    )
                except Exception:
                    final_status = "delivery_unknown"
                    log.exception("PPTX delivered, but DB completion could not be confirmed: %s", generation_id)
                if lease_task and not lease_task.done():
                    lease_task.cancel()
                    try:
                        await lease_task
                    except asyncio.CancelledError:
                        pass
                    lease_task = None
                if final_status != "completed":
                    log.error("PPTX was delivered but completion state is %s for %s", final_status, generation_id)
                    await notify_admins_once(
                        f"generation-finalize-{generation_id}",
                        f"🚨 <b>PPTX надіслано, але завершення не підтверджене в БД.</b>\nUser: <code>{user_id}</code>\nGeneration: <code>{generation_id}</code>",
                        cooldown=60,
                    )
                LAST_GENERATION_PARAMS[user_id] = {
                    "topic": data["topic"], "design": design, "slides": data["slides"],
                    "detail": data.get("detail", "medium"), "language": data.get("language", "uk"),
                    "audience": data.get("audience", "general"), "photos": photos_enabled,
                }

                if pdf_ready and loop.time() < deadline:
                    try:
                        await run_stage_with_deadline(
                            lambda: callback.message.answer_document(
                                FSInputFile(pdf_path, filename="PrezentAI.pdf"),
                                caption=tr(ui, "pdf_caption"), reply_markup=main_keyboard(ui),
                            ),
                            deadline, "Telegram PDF upload", 90.0,
                        )
                        pdf_sent = True
                    except Exception:
                        log.exception("PDF could not be sent; PPTX was already delivered")
                if not pdf_sent:
                    try:
                        await callback.message.answer(tr(ui, "pdf_failed"), reply_markup=main_keyboard(ui))
                    except Exception:
                        log.warning("Could not send PDF fallback notice")

        if photos_enabled and not any(sl.get("image_path") for sl in presentation_data.get("slides", [])):
            await callback.message.answer({
                "uk": "ℹ️ Фото цього разу знайти не вдалося, тому презентація без зображень. Спробуй створити ще раз або уточни тему.",
                "ru": "ℹ️ Фото в этот раз найти не удалось, поэтому презентация без изображений. Попробуй создать ещё раз или уточни тему.",
                "en": "ℹ️ No suitable photos could be found this time, so the deck has no images. Try again or make the topic more specific.",
            }[ui])
        await progress.complete(tr(ui, "generation_complete") if pdf_sent else tr(ui, "generation_pptx_only"))
        await callback.message.answer(tr(ui, "ready_next"), reply_markup=main_keyboard(ui))
        await callback.message.answer(tr(ui, "repeat_last"), reply_markup=repeat_last_keyboard(ui))

    except asyncio.CancelledError:
        log.info("Generation cancelled by user %s", user_id)
        settle_result = "skipped"
        if reserved and not succeeded:
            try:
                settle_result = await settle_failed_generation()
            except Exception:
                log.exception("Could not atomically settle cancelled generation %s", generation_id)
        if settle_result == "delivery_unknown":
            await notify_admins_once(
                f"delivery-unknown-{generation_id}",
                f"⚠️ <b>Невідомий результат відправлення презентації.</b> User: <code>{user_id}</code>; generation: <code>{generation_id}</code>. Автоповернення не виконано.",
                cooldown=60,
            )
        if progress:
            if settle_result == "failed":
                finish_text = tr(ui, "cancelled")
            elif settle_result in ("delivery_unknown", "settlement_failed"):
                finish_text = {
                    "uk": "⚠️ Не вдалося підтвердити результат скасування. Не запускай повторну генерацію, доки адміністратор не перевірить статус і баланс.",
                    "ru": "⚠️ Не удалось подтвердить результат отмены. Не запускай повторную генерацию, пока администратор не проверит статус и баланс.",
                    "en": "⚠️ The cancellation result could not be confirmed. Do not start another generation until an administrator checks the status and balance.",
                }[ui]
            elif settle_result == "delivered":
                finish_text = tr(ui, "already_sent")
            else:
                finish_text = tr(ui, "create_cancelled")
            await progress.finish(finish_text)
        raise
    except asyncio.TimeoutError as exc:
        settle_result = "skipped"
        if reserved and not succeeded:
            try:
                settle_result = await settle_failed_generation()
            except Exception:
                log.exception("Could not atomically settle timed-out generation %s", generation_id)
        if settle_result == "delivery_unknown":
            timeout_text = {
                "uk": "⚠️ Час очікування відправлення минув, але Telegram міг отримати файл. Ліміт не повернуто автоматично, щоб уникнути подвійної генерації. Адміністратора сповіщено.",
                "ru": "⚠️ Истекло время отправки, но Telegram мог получить файл. Лимит не возвращён автоматически, чтобы избежать двойной генерации. Администратор уведомлён.",
                "en": "⚠️ The upload timed out, but Telegram may have received the file. The quota was not automatically refunded to prevent duplicate free generations. An administrator was notified.",
            }[ui]
            await notify_admins_once(
                f"delivery-timeout-{generation_id}",
                f"⚠️ <b>Тайм-аут під час відправлення.</b> User: <code>{user_id}</code>; generation: <code>{generation_id}</code>.",
                cooldown=60,
            )
        elif settle_result == "failed":
            timeout_text = {
                "uk": f"⌛ Генерація перевищила загальний ліміт {int(generation_timeout)} секунд. Кредит або денний ліміт повернуто.",
                "ru": f"⌛ Генерация превысила общий лимит {int(generation_timeout)} секунд. Кредит или дневной лимит возвращён.",
                "en": f"⌛ Generation exceeded the overall {int(generation_timeout)}-second limit. Your credit or daily quota was restored.",
            }[ui]
        elif settle_result == "settlement_failed":
            timeout_text = {
                "uk": "⚠️ Час генерації вичерпано, але повернення ліміту не вдалося підтвердити через помилку бази. Не запускай нову генерацію; адміністратора сповіщено.",
                "ru": "⚠️ Время генерации истекло, но возврат лимита не удалось подтвердить из-за ошибки базы. Не запускай новую генерацию; администратор уведомлён.",
                "en": "⚠️ The generation timed out, but the quota refund could not be confirmed because of a database error. Do not start another generation; the administrator has been notified.",
            }[ui]
        elif settle_result == "delivered":
            timeout_text = {
                "uk": "✅ PowerPoint уже надіслано. Додатковий етап завершився із затримкою; денний ліміт використано.",
                "ru": "✅ PowerPoint уже отправлен. Дополнительный этап завершился с задержкой; дневной лимит использован.",
                "en": "✅ The PowerPoint was already sent. A follow-up step timed out; the daily quota was used.",
            }[ui]
        else:
            timeout_text = {
                "uk": "⌛ Час генерації вичерпано. Статус квоти потрібно перевірити в профілі або через адміністратора.",
                "ru": "⌛ Время генерации истекло. Статус лимита нужно проверить в профиле или через администратора.",
                "en": "⌛ Generation timed out. Please check your quota status in the profile or with an administrator.",
            }[ui]
        log.warning("Generation deadline exceeded for %s: %s", generation_id, exc)
        if progress:
            await progress.finish(timeout_text)
    except GeminiAPIError as exc:
        settle_result = "skipped"
        if reserved and not succeeded:
            settle_result = await settle_failed_generation()
        if settle_result == "delivery_unknown":
            await notify_admins_once(
                f"delivery-unknown-{generation_id}",
                f"⚠️ <b>Невідомий результат відправлення.</b> User: <code>{user_id}</code>; generation: <code>{generation_id}</code>.",
                cooldown=60,
            )
        log.warning("Gemini generation failed: %s", exc)
        await notify_admins_once(
            "gemini-generation-error", f"⚠️ <b>Помилка Gemini у Prezent AI.</b>\n{html.escape(str(exc)[:700])}"
        )
        if progress:
            if settle_result == "failed":
                finish_text = tr(ui, "ai_failed", reason=html.escape(str(exc)))
            elif settle_result == "delivery_unknown":
                finish_text = {
                    "uk": "⚠️ Результат відправлення невідомий. Ліміт не повертався автоматично. Не запускай нову генерацію; адміністратор перевірить операцію.",
                    "ru": "⚠️ Результат отправки неизвестен. Лимит не возвращался автоматически. Не запускай новую генерацию; администратор проверит операцию.",
                    "en": "⚠️ Delivery status is unknown. The quota was not automatically refunded. Do not start another generation; an administrator will review it.",
                }[ui]
            elif settle_result == "settlement_failed":
                finish_text = {
                    "uk": "⚠️ Сталася помилка AI, але повернення квоти не підтверджене через базу даних. Не повторюй генерацію; адміністратора сповіщено.",
                    "ru": "⚠️ Произошла ошибка AI, но возврат лимита не подтверждён из-за базы данных. Не повторяй генерацию; администратор уведомлён.",
                    "en": "⚠️ The AI request failed, but the quota refund could not be confirmed because of a database error. Do not retry yet; the administrator has been notified.",
                }[ui]
            elif settle_result == "delivered":
                finish_text = {
                    "uk": "✅ PowerPoint уже надіслано. Після цього сталася додаткова помилка; денний ліміт використано.",
                    "ru": "✅ PowerPoint уже отправлен. После этого произошла дополнительная ошибка; дневной лимит использован.",
                    "en": "✅ The PowerPoint was already sent. A follow-up step failed; the daily quota was used.",
                }[ui]
            else:
                finish_text = {
                    "uk": "❌ AI не зміг створити презентацію. Перевір статус ліміту в профілі перед повторною спробою.",
                    "ru": "❌ AI не смог создать презентацию. Проверь статус лимита в профиле перед повторной попыткой.",
                    "en": "❌ AI could not create the presentation. Check your quota status in the profile before retrying.",
                }[ui]
            await progress.finish(finish_text)
        await callback.message.answer(tr(ui, "retry_menu"), reply_markup=main_keyboard(ui))
    except Exception:
        error_id = uuid.uuid4().hex[:8].upper()
        log.exception("Presentation generation failed; error_id=%s", error_id)
        settle_result = "skipped"
        if reserved and not succeeded:
            try:
                settle_result = await settle_failed_generation()
            except Exception:
                log.exception("Could not atomically settle failed generation %s", generation_id)
        if settle_result == "delivery_unknown":
            await notify_admins_once(
                f"delivery-unknown-{generation_id}",
                f"🚨 <b>Невідомий результат відправлення презентації.</b>\nUser: <code>{user_id}</code>; generation: <code>{generation_id}</code>; error: <code>{error_id}</code>. Автоповернення не виконано.",
                cooldown=60,
            )
        await notify_admins_once(
            "generation-runtime-error",
            f"🚨 <b>Помилка генерації Prezent AI.</b> Код: <code>{error_id}</code>\nUser: <code>{user_id}</code>",
        )
        if progress:
            if settle_result == "failed":
                finish_text = tr(ui, "generic_failed", error_id=error_id)
            elif settle_result == "delivery_unknown":
                finish_text = {
                    "uk": f"⚠️ Помилка під час відправлення (код {error_id}). Telegram міг отримати файл. Ліміт не повернуто автоматично; не запускай повторну генерацію до перевірки адміністратором.",
                    "ru": f"⚠️ Ошибка при отправке (код {error_id}). Telegram мог получить файл. Лимит не возвращён автоматически; не запускай повторную генерацию до проверки администратором.",
                    "en": f"⚠️ Upload error (code {error_id}). Telegram may have received the file. The quota was not automatically refunded; do not retry until an administrator reviews it.",
                }[ui]
            elif settle_result == "settlement_failed":
                finish_text = {
                    "uk": f"⚠️ Сталася помилка (код {error_id}), а повернення квоти не підтверджено. Не повторюй генерацію; адміністратора сповіщено.",
                    "ru": f"⚠️ Произошла ошибка (код {error_id}), а возврат лимита не подтверждён. Не повторяй генерацию; администратор уведомлён.",
                    "en": f"⚠️ An error occurred (code {error_id}), and the quota refund could not be confirmed. Do not retry yet; the administrator has been notified.",
                }[ui]
            elif settle_result == "delivered":
                finish_text = {
                    "uk": f"✅ PowerPoint уже надіслано. Після відправлення сталася додаткова помилка (код {error_id}); ліміт використано.",
                    "ru": f"✅ PowerPoint уже отправлен. После отправки произошла дополнительная ошибка (код {error_id}); лимит использован.",
                    "en": f"✅ The PowerPoint was already sent. A follow-up step failed (code {error_id}); the quota was used.",
                }[ui]
            else:
                finish_text = {
                    "uk": f"❌ Не вдалося створити презентацію (код {error_id}). Перевір ліміт у профілі перед повторною спробою.",
                    "ru": f"❌ Не удалось создать презентацию (код {error_id}). Проверь лимит в профиле перед повторной попыткой.",
                    "en": f"❌ Could not create the presentation (code {error_id}). Check your quota in the profile before retrying.",
                }[ui]
            await progress.finish(finish_text)
        await callback.message.answer(tr(ui, "check_logs"), reply_markup=main_keyboard(ui))
    finally:
        if queue_registered:
            GENERATION_QUEUE_WAITERS = max(0, GENERATION_QUEUE_WAITERS - 1)
        if lease_task and not lease_task.done():
            lease_task.cancel()
            try:
                await lease_task
            except asyncio.CancelledError:
                pass
        if progress:
            await progress.stop()
        if ACTIVE_GENERATIONS.get(user_id) is current_task:
            ACTIVE_GENERATIONS.pop(user_id, None)


@dp.message(F.text)
async def unknown_text(message: Message):
    await asyncio.to_thread(register_user, message.from_user)
    language = get_user_language(message.from_user.id)
    if not language:
        await show_language_picker(message)
        return
    await message.answer(tr(language, "unknown"), reply_markup=main_keyboard(language))
 
def validate_pptx_file(path: str, expected_slides: int):
    """Check the PPTX ZIP container, required OOXML members, and parsed slide count."""
    target = Path(path)
    if not target.is_file() or not (0 < target.stat().st_size <= MAX_OUTPUT_FILE_BYTES):
        raise ValueError("PowerPoint output is missing, empty, or larger than the Telegram upload limit")
    try:
        with zipfile.ZipFile(target, "r") as archive:
            bad_member = archive.testzip()
            required = {"[Content_Types].xml", "ppt/presentation.xml", "ppt/slides/slide1.xml"}
            names = set(archive.namelist())
            if bad_member:
                raise ValueError(f"corrupt PPTX archive member: {bad_member}")
            if not required.issubset(names):
                raise ValueError("PPTX is missing required Office Open XML files")
        parsed = Presentation(str(target))
        if len(parsed.slides) != int(expected_slides):
            raise ValueError(f"PPTX slide count mismatch: expected {expected_slides}, got {len(parsed.slides)}")
    except zipfile.BadZipFile as exc:
        raise ValueError("PPTX is not a valid ZIP/OOXML file") from exc
    return True


def validate_pdf_file(path: str):
    """Validate PDF signature, end marker, and presence of at least one page object."""
    target = Path(path)
    if not target.is_file() or not (0 < target.stat().st_size <= MAX_OUTPUT_FILE_BYTES):
        raise ValueError("PDF output is missing, empty, or larger than the Telegram upload limit")
    with target.open("rb") as file:
        header = file.read(8)
        file.seek(max(0, target.stat().st_size - 4096))
        tail = file.read()
    if not header.startswith(b"%PDF-") or b"%%EOF" not in tail:
        raise ValueError("PDF signature or final EOF marker is missing")
    # ReportLab writes page dictionaries as /Type /Page; this catches empty PDFs.
    with target.open("rb") as file:
        content = file.read(MAX_OUTPUT_FILE_BYTES + 1)
    if not re.search(rb"/Type\s*/Page\b", content):
        raise ValueError("PDF does not contain a readable page object")
    return True


async def run_stage_with_deadline(factory, deadline: float, stage_name: str, stage_limit: float):
    """Run an awaitable with both a stage cap and the remaining whole-job deadline."""
    remaining = deadline - asyncio.get_running_loop().time()
    allowed = min(float(stage_limit), remaining)
    if allowed <= 0:
        raise asyncio.TimeoutError(f"Overall generation deadline exceeded before stage: {stage_name}")
    try:
        return await asyncio.wait_for(factory(), timeout=allowed)
    except asyncio.TimeoutError as exc:
        raise asyncio.TimeoutError(f"Stage timed out: {stage_name} ({allowed:.1f}s budget)") from exc


async def run_blocking_stage_with_deadline(func, args: tuple, deadline: float, stage_name: str, stage_limit: float):
    """Bound a thread-backed stage without deleting temp files while its thread still runs."""
    remaining = deadline - asyncio.get_running_loop().time()
    allowed = min(float(stage_limit), remaining)
    if allowed <= 0:
        raise asyncio.TimeoutError(f"Overall generation deadline exceeded before stage: {stage_name}")
    worker = asyncio.create_task(asyncio.to_thread(func, *args))
    try:
        return await asyncio.wait_for(asyncio.shield(worker), timeout=allowed)
    except asyncio.TimeoutError as exc:
        # asyncio cannot kill a running Python thread. Wait for file writes to stop before
        # TemporaryDirectory cleanup, then report the stage as timed out to the user.
        try:
            await asyncio.shield(worker)
        except Exception:
            log.exception("Blocking stage %s failed after its deadline", stage_name)
        raise asyncio.TimeoutError(f"Stage timed out: {stage_name} ({allowed:.1f}s budget)") from exc
    except asyncio.CancelledError:
        # Also wait for file I/O to finish when the parent job is cancelled.
        try:
            await asyncio.shield(worker)
        except Exception:
            log.exception("Blocking stage %s failed during cancellation", stage_name)
        raise


# ============================================================
# 13. RENDER HEALTH CHECK + WEBHOOK
# ============================================================
WEBHOOK_PATH = "/telegram/webhook"
RENDER_EXTERNAL_URL = os.getenv("RENDER_EXTERNAL_URL", "").rstrip("/")
WEBHOOK_SECRET = os.getenv("WEBHOOK_SECRET", "").strip()
if not WEBHOOK_SECRET:
    WEBHOOK_SECRET = hashlib.sha256(BOT_TOKEN.encode("utf-8")).hexdigest()
 
 
async def health_handler(request: web.Request):
    ready = request.app["startup_ready"].is_set()
    return web.json_response({
        "status": "ok",
        "service": "Prezent AI",
        "mode": "webhook",
        "ready": ready,
        "database": "turso" if TURSO_ENABLED else ("persistent_disk" if DURABLE_DB_ENABLED else "ephemeral_sqlite"),
        "payments_enabled": (not RENDER_IS_HOSTED) or DURABLE_DB_ENABLED,
    })
 
 
async def on_startup(bot: Bot):
    if not RENDER_EXTERNAL_URL.startswith("https://"):
        raise RuntimeError("RENDER_EXTERNAL_URL is missing. Deploy as Render Web Service.")
    webhook_url = RENDER_EXTERNAL_URL + WEBHOOK_PATH
    await bot.set_webhook(
        url=webhook_url, secret_token=WEBHOOK_SECRET,
        allowed_updates=dp.resolve_used_update_types(), drop_pending_updates=False,
        max_connections=10,
    )
    info = await bot.get_webhook_info()
    log.info("Telegram webhook registered: %s", info.url)
    log.info("Webhook pending updates: %s", info.pending_update_count)
    await bot.set_my_commands([
        BotCommand(command="start", description="Головне меню"),
        BotCommand(command="restart", description="Перезапустити діалог / Restart"),
        BotCommand(command="language", description="Мова спілкування / Language"),
        BotCommand(command="cancel", description="Скасувати створення / Cancel"),
        BotCommand(command="shop", description="Магазин ⭐"),
        BotCommand(command="id", description="Мій Telegram ID"),
    ])
 
 
@web.middleware
async def startup_gate(request: web.Request, handler):
    # Telegram can keep sending updates to the existing webhook during a deploy.
    # Hold webhook requests until the database and Telegram setup are ready.
    if request.path == WEBHOOK_PATH and not request.app["startup_ready"].is_set():
        try:
            await asyncio.wait_for(request.app["startup_ready"].wait(), timeout=30)
        except asyncio.TimeoutError:
            return web.Response(status=503, text="Prezent AI is starting")
    return await handler(request)
 
 
async def startup_admin_warnings():
    if not ADMIN_IDS:
        return
    try:
        user_count = await asyncio.to_thread(get_user_count_db)
    except Exception:
        log.exception("Could not inspect users table at startup")
        user_count = -1

    alerts = []
    if RENDER_IS_HOSTED and not DURABLE_DB_ENABLED:
        alerts.append(
            "🚨 <b>КРИТИЧНО: постійне сховище не підключене.</b> "
            "Render Free стирає локальну SQLite при redeploy/restart/spin-down. "
            "Покупки Telegram Stars автоматично вимкнені, щоб не втратити гроші користувачів. "
            "Підключи Turso і додай TURSO_DATABASE_URL та TURSO_AUTH_TOKEN у Render Environment, "
            "або встанови RENDER_DISK_MOUNT_PATH і DATABASE_PATH усередині точки монтування постійного диска."
        )
    if user_count == 0:
        alerts.append(
            "⚠️ <b>Таблиця users порожня.</b> Якщо це не перший запуск, перевір TURSO_DATABASE_URL / TURSO_AUTH_TOKEN "
            "або віднови останній backup через /restore. Не приймай платежі, доки не перевіриш баланс користувачів."
        )
    if user_count < 0:
        alerts.append("⚠️ Не вдалося перевірити таблицю users під час старту. Переглянь Runtime Logs.")

    if alerts:
        message = "\n\n".join(alerts)
        for admin_id in ADMIN_IDS:
            try:
                await bot.send_message(admin_id, message)
            except Exception:
                log.warning("Could not send startup storage warning to admin %s", admin_id)


async def pending_generation_recovery_loop():
    while True:
        await asyncio.sleep(60)
        try:
            recovered = await asyncio.to_thread(recover_stale_pending_generations, 240)
            if recovered:
                log.warning("Recovered %s interrupted generation reservations", recovered)
        except asyncio.CancelledError:
            raise
        except Exception as exc:
            log.exception("Pending generation recovery sweep failed")
            await notify_admins_once(
                "generation-recovery-loop-failed",
                f"🚨 <b>Не вдалося перевірити завислі генерації.</b>\n{html.escape(str(exc)[:500])}",
            )


async def schedule_backend_initialization(app: web.Application):
    global _APP_EVENT_LOOP
    _APP_EVENT_LOOP = asyncio.get_running_loop()

    async def initialize_backend():
        try:
            # Do database/network initialization only after aiohttp has opened PORT.
            await asyncio.to_thread(init_db)
            await asyncio.to_thread(load_user_language_cache)
            user_count = await asyncio.to_thread(get_user_count_db)
            if user_count == 0 and TURSO_ENABLED:
                for admin_id in ADMIN_IDS:
                    try:
                        await bot.send_message(admin_id, "⚠️ Turso підключений, але таблиця users порожня. Перевір, чи було відновлено резервну копію старої бази.")
                    except Exception:
                        log.exception("Could not send empty database warning to admin %s", admin_id)
            recovered = await asyncio.to_thread(recover_stale_pending_generations, 240)
            if recovered:
                log.warning("Startup recovered %s interrupted quota reservations", recovered)
            await on_startup(bot)
            app["startup_ready"].set()
            app["backup_task"] = asyncio.create_task(backup_loop())
            app["pending_recovery_task"] = asyncio.create_task(pending_generation_recovery_loop())
            log.info("Prezent AI startup completed; webhook processing is ready")
            asyncio.create_task(startup_admin_warnings())
        except asyncio.CancelledError:
            raise
        except Exception as exc:
            log.exception("Deferred Prezent AI initialization failed")
            for admin_id in ADMIN_IDS:
                try:
                    await bot.send_message(
                        admin_id,
                        "🚨 <b>Prezent AI не завершив ініціалізацію.</b> "
                        "Перевір Runtime Logs, підключення Turso та змінні Environment. "
                        f"Помилка: <code>{html.escape(str(exc)[:400])}</code>"
                    )
                except Exception:
                    log.warning("Could not alert admin %s about startup failure", admin_id)
 
    app["backend_startup_task"] = asyncio.create_task(initialize_backend())
 
 
async def cancel_backend_initialization(app: web.Application):
    backup = app.get("backup_task")
    if backup and not backup.done():
        backup.cancel()
    recovery = app.get("pending_recovery_task")
    if recovery and not recovery.done():
        recovery.cancel()
    task = app.get("backend_startup_task")
    if task and not task.done():
        task.cancel()
        try:
            await task
        except asyncio.CancelledError:
            pass
 
 
def main():
    # Bind the public port first. Database initialization and Telegram API calls
    # are deliberately deferred so Render's port scanner can detect the service.
    port = int(os.getenv("PORT", "10000"))
    app = web.Application(
        client_max_size=1024 * 1024,
        middlewares=[startup_gate],
    )
    app["startup_ready"] = asyncio.Event()
    app.router.add_get("/", health_handler)
    app.router.add_get("/healthz", health_handler)
    webhook_handler = SimpleRequestHandler(
        dispatcher=dp, bot=bot, handle_in_background=True, secret_token=WEBHOOK_SECRET
    )
    webhook_handler.register(app, path=WEBHOOK_PATH)
    setup_application(app, dp, bot=bot)
    app.on_startup.append(schedule_backend_initialization)
    app.on_cleanup.append(cancel_backend_initialization)
    log.info("Starting Prezent AI Webhook server on port %s", port)
    log.info("Binding address: 0.0.0.0:%s", port)
    log.info("Primary Gemini model: %s", GEMINI_MODEL)
    log.info("Fallback Gemini model: %s", GEMINI_FALLBACK_MODEL)
    log.info("Admin IDs configured: %s", sorted(ADMIN_IDS))
    backend = "Turso (durable remote SQLite)" if TURSO_ENABLED else ("Render persistent disk SQLite" if DURABLE_DB_ENABLED else "local/ephemeral SQLite")
    log.info("Database backend: %s", backend)
    if RENDER_IS_HOSTED and not DURABLE_DB_ENABLED:
        log.error("CRITICAL: no durable storage configured. Telegram Stars purchases are disabled to protect users.")
    web.run_app(app, host="0.0.0.0", port=port, shutdown_timeout=30)
 
 
if __name__ == "__main__":
    main()