
import asyncio
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
from datetime import datetime, timedelta, timezone
from pathlib import Path
from xml.sax.saxutils import escape
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
from reportlab.platypus import Paragraph
 
# ============================================================
# 1. CONFIGURATION
# ============================================================
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
)
log = logging.getLogger("prezent-ai")
 
BOT_TOKEN = os.getenv("BOT_TOKEN", "").strip()
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "").strip()
GEMINI_MODEL = os.getenv("GEMINI_MODEL", "gemini-3.5-flash-lite").strip()
GEMINI_FALLBACK_MODEL = os.getenv("GEMINI_FALLBACK_MODEL", "gemini-3.1-flash-lite").strip()
ADMIN_IDS_RAW = os.getenv("ADMIN_IDS", os.getenv("ADMIN_ID", ""))
ADMIN_IDS = {int(x.strip()) for x in ADMIN_IDS_RAW.split(",") if x.strip().isdigit()}
DATABASE_PATH = os.getenv("DATABASE_PATH", "prezentai.sqlite3").strip()
RENDER_IS_HOSTED = bool(os.getenv("RENDER_EXTERNAL_URL", "").strip())
TURSO_DATABASE_URL = os.getenv("TURSO_DATABASE_URL", "").strip()
TURSO_AUTH_TOKEN = os.getenv("TURSO_AUTH_TOKEN", "").strip()
TURSO_ENABLED = bool(TURSO_DATABASE_URL and TURSO_AUTH_TOKEN)
if bool(TURSO_DATABASE_URL) != bool(TURSO_AUTH_TOKEN):
    raise RuntimeError("Set BOTH TURSO_DATABASE_URL and TURSO_AUTH_TOKEN, or remove both.")


def _is_render_disk_path(path: str) -> bool:
    try:
        resolved = Path(path).expanduser().resolve()
        mount = Path("/data")
        return os.path.ismount(str(mount)) and resolved.is_relative_to(mount)
    except Exception:
        return False


DURABLE_DB_ENABLED = TURSO_ENABLED or _is_render_disk_path(DATABASE_PATH)
MAX_TOPIC_LENGTH = 500
FREE_DAILY_LIMIT = 3
PRO_DAILY_LIMIT = 10
CREDIT_PACK_SIZE = 5
CREDIT_PACK_PRICE = 15
PRO_PRICE = 99
PRO_DAYS = 30
KYIV_TZ = ZoneInfo("Europe/Kyiv")
GEN_SEMAPHORE = asyncio.Semaphore(2)
 
if not BOT_TOKEN:
    raise RuntimeError("BOT_TOKEN is missing in Render Environment.")
if not GEMINI_API_KEY:
    raise RuntimeError("GEMINI_API_KEY is missing in Render Environment.")
 
Path(DATABASE_PATH).expanduser().resolve().parent.mkdir(parents=True, exist_ok=True)
bot = Bot(token=BOT_TOKEN, default=DefaultBotProperties(parse_mode=ParseMode.HTML))
dp = Dispatcher(storage=MemoryStorage())
ACTIVE_GENERATIONS = {}
 
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
    source_message_id BIGINT
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


def get_user_language(user_id: int) -> str | None:
    row = get_user(user_id)
    if not row:
        return None
    language = row["interface_language"] if "interface_language" in row.keys() else None
    return language if language in ("uk", "ru", "en") else None


def set_user_language(user_id: int, language: str):
    if language not in ("uk", "ru", "en"):
        raise ValueError("Unsupported interface language")
    with db_connect() as con:
        con.execute(
            "INSERT INTO users(user_id, interface_language, created_at) VALUES(?,?,?) "
            "ON CONFLICT(user_id) DO UPDATE SET interface_language=excluded.interface_language",
            (user_id, language, now_iso()),
        )
 
 
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
 
 
def reserve_generation(user_id: int):
    """Reserve one daily quota slot or one purchased credit atomically."""
    if is_admin(user_id):
        return {"allowed": True, "mode": "admin", "used": 0, "limit": "∞", "credits": 0, "day": today_key()}
    day = today_key()
    with db_connect() as con:
        con.execute("BEGIN")
        row = con.execute("SELECT * FROM users WHERE user_id=?", (user_id,)).fetchone()
        if not row:
            con.execute(
                "INSERT INTO users(user_id, created_at) VALUES(?,?)", (user_id, now_iso())
            )
            row = con.execute("SELECT * FROM users WHERE user_id=?", (user_id,)).fetchone()
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
            return {"allowed": True, "mode": "pro" if is_pro else "free", "used": used + 1,
                    "limit": limit, "credits": credits, "day": day}
        if credits > 0:
            con.execute("UPDATE users SET credits=credits-1 WHERE user_id=?", (user_id,))
            return {"allowed": True, "mode": "credit", "used": used, "limit": limit,
                    "credits": credits - 1, "day": day}
        return {"allowed": False, "mode": "none", "used": used, "limit": limit,
                "credits": credits, "day": day}
 
 
def release_generation(user_id: int, quota: dict):
    """Restore the exact reserved quota slot/credit even if midnight passed."""
    mode = quota.get("mode", "none")
    if mode == "admin":
        return
    with db_connect() as con:
        if mode == "credit":
            con.execute("UPDATE users SET credits=credits+1 WHERE user_id=?", (user_id,))
        elif mode in ("free", "pro"):
            con.execute(
                "UPDATE daily_usage SET used=CASE WHEN used > 0 THEN used - 1 ELSE 0 END WHERE user_id=? AND day=?",
                (user_id, quota.get("day") or today_key()),
            )
    if mode == "credit":
        schedule_balance_backup("повернення витраченого кредиту")
 
 
def add_history(generation_id, user_id, topic, design, slides, quota_mode, status="completed"):
    with db_connect() as con:
        con.execute(
            "INSERT INTO generation_history(generation_id,user_id,topic,design,slides,quota_mode,status,created_at) "
            "VALUES(?,?,?,?,?,?,?,?)",
            (generation_id, user_id, topic, design, slides, quota_mode, status, now_iso()),
        )
 
 
def set_history_message(generation_id: str, chat_id: int, message_id: int):
    with db_connect() as con:
        con.execute(
            "UPDATE generation_history SET source_chat_id=?, source_message_id=? WHERE generation_id=?",
            (chat_id, message_id, generation_id),
        )
 
 
def mark_history_failed(generation_id: str):
    with db_connect() as con:
        con.execute(
            "UPDATE generation_history SET status='failed' WHERE generation_id=?",
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
 
 
def process_payment(payload: str, user_id: int, amount: int, currency: str, charge_id: str):
    if currency != "XTR":
        return False, "Неправильна валюта платежу."
    with db_connect() as con:
        con.execute("BEGIN")
        order = con.execute("SELECT * FROM orders WHERE payload=?", (payload,)).fetchone()
        if not order or int(order["user_id"]) != user_id:
            return False, "Замовлення не знайдено або воно належить іншому користувачу."
        if order["status"] == "paid":
            return True, "already_paid"
        if int(order["amount"]) != amount or order["status"] != "pending":
            return False, "Сума або статус платежу не збігається."
        duplicate = con.execute(
            "SELECT 1 FROM orders WHERE charge_id=?", (charge_id,)
        ).fetchone()
        if duplicate:
            return True, "already_paid"
        if order["product"] == "credits":
            con.execute("UPDATE users SET credits=credits+? WHERE user_id=?", (CREDIT_PACK_SIZE, user_id))
        elif order["product"] == "pro":
            row = con.execute("SELECT subscription_until FROM users WHERE user_id=?", (user_id,)).fetchone()
            current = None
            if row and row["subscription_until"]:
                try:
                    current = datetime.fromisoformat(row["subscription_until"])
                except ValueError:
                    current = None
            start = current if current and current > now_utc() else now_utc()
            until = start + timedelta(days=PRO_DAYS)
            con.execute("UPDATE users SET subscription_until=? WHERE user_id=?", (until.isoformat(), user_id))
        else:
            return False, "Невідомий тип товару."
        con.execute(
            "UPDATE orders SET status='paid', charge_id=? WHERE payload=?", (charge_id, payload)
        )
        return True, order["product"]
 
 
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
            "SELECT r.*, g.topic, g.design, g.slides, g.source_chat_id, g.source_message_id "
            "FROM reports r LEFT JOIN generation_history g ON g.generation_id=r.generation_id "
            "WHERE r.report_id=?", (report_id,)
        ).fetchone()
 
 
def handle_report(report_id: str, admin_id: int, action: str):
    with db_connect() as con:
        con.execute("BEGIN")
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
    except Exception:
        log.exception("Backup snapshot failed")
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
            except Exception:
                log.warning("Could not send backup to admin %s", admin_id)
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


def schedule_balance_backup(reason="зміна балансу"):
    """Coalesce many balance changes into at most one Telegram backup per 45s."""
    global _pending_balance_backup_reason, _balance_backup_task
    _pending_balance_backup_reason = str(reason)[:120]
    try:
        loop = asyncio.get_running_loop()
    except RuntimeError:
        log.warning("Balance changed outside an event loop; automatic backup was not scheduled")
        return
    if _balance_backup_task is None or _balance_backup_task.done():
        _balance_backup_task = loop.create_task(_run_debounced_balance_backup())


async def backup_loop():
    while True:
        await asyncio.sleep(6 * 3600)
        await backup_to_admins("авто, кожні 6 год")


def restore_db_from(path: str):
    """Restore a Telegram SQLite backup into local SQLite or remote Turso."""
    if not TURSO_ENABLED:
        src = sqlite3.connect(path)
        try:
            src.execute("SELECT COUNT(*) FROM users").fetchone()
            dst = sqlite3.connect(DATABASE_PATH, timeout=30)
            try:
                src.backup(dst)
            finally:
                dst.close()
        finally:
            src.close()
        return

    src = sqlite3.connect(path)
    src.row_factory = sqlite3.Row
    try:
        src.execute("SELECT COUNT(*) FROM users").fetchone()
        with db_connect() as dst:
            dst.execute("BEGIN")
            for table in reversed(BACKUP_TABLES):
                dst.execute(f"DELETE FROM {table}")
            for table in BACKUP_TABLES:
                cursor = src.execute(f"SELECT * FROM {table}")
                columns = [item[0] for item in (cursor.description or ())]
                rows = cursor.fetchall()
                if not columns:
                    continue
                quoted_columns = ",".join('"' + name.replace('"', '""') + '"' for name in columns)
                placeholders = ",".join("?" for _ in columns)
                sql = f"INSERT INTO {table} ({quoted_columns}) VALUES ({placeholders})"
                for row in rows:
                    dst.execute(sql, tuple(row[name] for name in columns))
    finally:
        src.close()

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
    "design_prompt": {"uk": "🎨 <b>Обери візуальний стиль.</b> Кожен стиль має власну композицію, типографіку та оформлення:", "ru": "🎨 <b>Выбери визуальный стиль.</b> У каждого стиля своя композиция, типографика и оформление:", "en": "🎨 <b>Choose a visual style.</b> Each style has its own composition, typography, and visual treatment:"},
    "creating": {"uk": "Ваша презентація створюється", "ru": "Ваша презентация создаётся", "en": "Your presentation is being created"},
    "please_wait": {"uk": "AI розробляє структуру та оформлює слайди…", "ru": "AI разрабатывает структуру и оформляет слайды…", "en": "AI is shaping the story and designing the slides…"},
    "generation_complete": {"uk": "✅ <b>Презентація готова!</b>\n\nPowerPoint і PDF надіслано 📤", "ru": "✅ <b>Презентация готова!</b>\n\nPowerPoint и PDF отправлены 📤", "en": "✅ <b>Your presentation is ready!</b>\n\nPowerPoint and PDF have been sent 📤"},
    "generation_pptx_only": {"uk": "✅ <b>PowerPoint створено й надіслано!</b>\n\n⚠️ PDF цього разу не вдалося надіслати.", "ru": "✅ <b>PowerPoint создан и отправлен!</b>\n\n⚠️ В этот раз не удалось отправить PDF.", "en": "✅ <b>PowerPoint created and sent!</b>\n\n⚠️ The PDF could not be sent this time."},
    "ready_next": {"uk": "✅ <b>Готово!</b>\n\nМожеш створити наступну презентацію кнопкою нижче.", "ru": "✅ <b>Готово!</b>\n\nМожешь создать следующую презентацию кнопкой ниже.", "en": "✅ <b>Done!</b>\n\nYou can create another presentation using the button below."},
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
        "uk": "📊 <b>{title}</b>\n🎨 Стиль: {style}\nФормат: PowerPoint • плавні Fade-переходи\n\nЯкщо помітиш проблему, натисни кнопку нижче — адміністратор перевірить файл.",
        "ru": "📊 <b>{title}</b>\n🎨 Стиль: {style}\nФормат: PowerPoint • плавные Fade-переходы\n\nЕсли заметишь проблему, нажми кнопку ниже — администратор проверит файл.",
        "en": "📊 <b>{title}</b>\n🎨 Style: {style}\nFormat: PowerPoint • smooth Fade transitions\n\nIf you notice a problem, press the button below and an administrator will review the file.",
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


def slides_keyboard():
    return InlineKeyboardMarkup(inline_keyboard=[[
        InlineKeyboardButton(text=str(n), callback_data=f"slides:{n}") for n in (6, 8, 10, 12)
    ]])


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
LAYOUTS = {"cover", "split", "cards", "timeline", "steps", "quote", "comparison", "focus", "closing"}
 
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
 
class PresentationForm(StatesGroup):
    topic = State()
    slides = State()
    language = State()
    design = State()
 
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
# 5. ANIMATION# ============================================================
# 5. ANIMATION — мінімалістичні рухомі зірочки без відсотків
# ============================================================
class ProgressAnimator:
    """Polished Telegram status animation. It is started only by design_handler."""
    TRACK_WIDTH = 22
    TICK_SECONDS = 1.8

    STAGES = {
        "uk": {
            "writing": "Пишу тексти й вибудовую логіку розповіді…",
            "pptx": "Збираю слайди PowerPoint…",
            "pdf": "Готую PDF-версію…",
            "sending": "Перевіряю та надсилаю готові файли…",
        },
        "ru": {
            "writing": "Пишу тексты и выстраиваю логику презентации…",
            "pptx": "Собираю слайды PowerPoint…",
            "pdf": "Готовлю PDF-версию…",
            "sending": "Проверяю и отправляю готовые файлы…",
        },
        "en": {
            "writing": "Writing the content and shaping the story…",
            "pptx": "Building the PowerPoint slides…",
            "pdf": "Preparing the PDF version…",
            "sending": "Checking and sending the finished files…",
        },
    }

    def __init__(self, language: str = "uk", message: Message | None = None):
        self.language = language if language in ("uk", "ru", "en") else "uk"
        self.message = message
        self.frame = 0
        self.stage = "writing"
        self.task = None
        self.lock = asyncio.Lock()
        self.running = False

    def set_stage(self, key: str):
        """Update the label only when the generator actually enters that stage."""
        if key in self.STAGES[self.language]:
            self.stage = key

    def _track(self) -> str:
        # A bright ✨ travels along a fine line and gently returns in a loop.
        span = self.TRACK_WIDTH * 2
        raw = (self.frame * 2) % span
        position = raw if raw <= self.TRACK_WIDTH else span - raw
        cells = ["─"] * (self.TRACK_WIDTH + 1)
        cells[position] = "✨"
        # A second, smaller sparkle follows at a different distance.
        second = (position + 10) % (self.TRACK_WIDTH + 1)
        if second != position:
            cells[second] = "✧"
        return "✦ " + "".join(cells) + " ✦"

    def _starfield(self) -> str:
        width = 27
        rows = []
        symbols = ("✧", "·", "✦", "·", "✧")
        for row in range(2):
            cells = [" "] * width
            first = (self.frame * (2 + row) + 4 + row * 7) % width
            second = (self.frame * (3 + row) + 17 - row * 5) % width
            cells[first] = symbols[(self.frame + row) % len(symbols)]
            if second != first:
                cells[second] = symbols[(self.frame + row + 2) % len(symbols)]
            rows.append("".join(cells).rstrip())
        return "\n".join(rows)

    def _phase(self) -> str:
        return self.STAGES[self.language].get(self.stage, self.STAGES[self.language]["writing"])

    def render(self) -> str:
        return (
            "╭━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━╮\n"
            "          ✨ <b>PREZENT AI STUDIO</b> ✨\n"
            "╰━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━╯\n\n"
            f"<b>{html.escape(tr(self.language, 'creating'))}</b>\n\n"
            f"<pre>{html.escape(self._starfield())}\n{html.escape(self._track())}</pre>\n"
            f"<i>{html.escape(self._phase())}</i>\n"
            f"<i>{html.escape(tr(self.language, 'please_wait'))}</i>"
        )

    async def safe_edit(self):
        if not self.message:
            return False
        rendered = self.render()
        async with self.lock:
            try:
                await self.message.edit_text(rendered, parse_mode=ParseMode.HTML)
                return True
            except TelegramRetryAfter as exc:
                delay = max(1.0, float(getattr(exc, "retry_after", 1.0)))
                log.warning("Telegram rate-limited the generation animation; retrying after %.1fs", delay)
                await asyncio.sleep(delay)
                return False
            except Exception as exc:
                if "message is not modified" in str(exc).lower():
                    return True
                log.debug("Generation animation edit failed: %s", exc)
                return False

    async def run(self):
        self.running = True
        try:
            while self.running:
                await asyncio.sleep(self.TICK_SECONDS)
                self.frame += 1
                try:
                    await self.safe_edit()
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
        self.running = True
        self.task = asyncio.create_task(self.run(), name="prezent-ai-generation-animation")

    async def stop(self):
        self.running = False
        task, self.task = self.task, None
        if task and not task.done():
            task.cancel()
            try:
                await task
            except asyncio.CancelledError:
                pass


    async def finish(self, final_text: str):
        """Stop motion and replace the animation with a final localized status."""
        await self.stop()
        if not self.message:
            return
        async with self.lock:
            try:
                await self.message.edit_text(final_text, parse_mode=ParseMode.HTML)
            except TelegramRetryAfter as exc:
                delay = max(1.0, float(getattr(exc, "retry_after", 1.0)))
                await asyncio.sleep(delay)
                try:
                    await self.message.edit_text(final_text, parse_mode=ParseMode.HTML)
                except Exception as retry_exc:
                    log.warning("Could not send final generation status after flood wait: %s", retry_exc)
            except Exception as exc:
                if "message is not modified" not in str(exc).lower():
                    log.debug("Could not set final generation status: %s", exc)

    async def complete(self, final_text: str):
        """Backward-compatible success finalizer used by the generation pipeline."""
        await self.finish(final_text)

# ============================================================
# 6. GEMINI API# ============================================================
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
        "slides": {
            "type": "ARRAY",
            "items": {
                "type": "OBJECT",
                "properties": {
                    "title": {"type": "STRING"},
                    "subtitle": {"type": "STRING"},
                    "layout": {"type": "STRING", "enum": ["cover", "split", "cards", "timeline", "steps", "quote", "comparison", "focus", "closing"]},
                    "content": {"type": "ARRAY", "items": {"type": "STRING"}},
                    "highlight": {"type": "STRING"},
                    "left_title": {"type": "STRING"},
                    "right_title": {"type": "STRING"},
                    "notes": {"type": "STRING"},
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
 
 
async def request_gemini(model: str, prompt: str) -> dict:
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
            "maxOutputTokens": 8192,
            "responseMimeType": "application/json",
            "responseSchema": GEMINI_RESPONSE_SCHEMA,
            "temperature": 0.9,
            "topP": 0.95,
            "seed": random.randint(1, 2_000_000_000),
        },
    }
    for attempt in range(1, 2):
        try:
            async with httpx.AsyncClient(timeout=httpx.Timeout(65.0, connect=15.0)) as client:
                response = await client.post(url, headers=headers, json=payload)
        except httpx.TimeoutException as exc:
            raise GeminiRetryableError("Gemini не відповів протягом 65 секунд; перевірено резервну модель.") from exc
        except httpx.HTTPError as exc:
            raise GeminiRetryableError("Помилка з'єднання з Google Gemini.") from exc
        try:
            result = response.json()
        except ValueError as exc:
            raise GeminiAPIError(f"Gemini повернув неправильну відповідь. HTTP {response.status_code}.") from exc
        status = response.status_code
        if status in (408, 429, 500, 502, 503, 504):
            raise GeminiRetryableError(f"Gemini тимчасово недоступний: HTTP {status}.")
        if status == 404:
            raise GeminiModelUnavailable(f"Модель {model} не знайдена або недоступна.")
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
    raise GeminiRetryableError(f"Модель {model} не повернула результат.")
 
 
async def generate_content(data: dict) -> dict:
    topic, count = data["topic"], data["slides"]
    language = LANGUAGES.get(data["language"], "українською")
    design = data["design"]
    style_name = design_name(design, data.get("interface_language", "uk"))
    style_direction = STYLE_PROMPTS.get(design, STYLE_PROMPTS["minimal"])
    angle = random.choice(CREATIVE_ANGLES)
    tone = random.choice(CREATIVE_TONES)
    hook = random.choice(CREATIVE_HOOKS)
    architecture = random.choice(NARRATIVE_ARCHITECTURES)
    visual_concept = random.choice(DECK_VISUAL_CONCEPTS)
    variation_id = uuid.uuid4().hex[:12]
    visual_seed = random.randint(1, 2_000_000_000)
    prompt = f"""
Створи авторську, змістовну й професійну презентацію — не використовуй універсальний шаблон.
ТЕМА: {topic}
СЛАЙДІВ: рівно {count}
МОВА КОНТЕНТУ: {language}
ВИБРАНИЙ ВІЗУАЛЬНИЙ СТИЛЬ: {style_name} ({design})
АРТ-ДИРЕКШН СТИЛЮ: {style_direction}

ОБОВ'ЯЗКОВА ВІЗУАЛЬНА ІДЕНТИЧНІСТЬ:
- Обраний стиль має бути помітним на КОЖНОМУ слайді, а не лише на обкладинці.
- Не використовуй однакову композицію «заголовок зверху + список у прямокутнику» для всіх слайдів.
- Для кожного слайда обери доречну композицію: асиметричну редакційну, двоколонкову, картки, часову шкалу, покрокову схему, порівняння, цитату або великий ключовий висновок.
- Зміст і структура мають відповідати темі; не вставляй timeline, steps чи comparison, якщо вони не мають сенсу.
- Кожна презентація має власний візуальний задум, метафору, ритм та неповторний вступ.

УНІКАЛЬНА КОНЦЕПЦІЯ ЦІЄЇ ГЕНЕРАЦІЇ (ID {variation_id}):
- Архітектура розповіді: {architecture}
- Візуальна концепція та редакційний прийом: {visual_concept}
- Ракурс: {angle}
- Тон: {tone}
- Хук обкладинки: {hook}
- Внутрішній seed варіативності: {visual_seed}

Поверни тільки JSON такої структури:
{{
  "title":"Оригінальна назва",
  "slides":[
    {{"title":"Унікальний заголовок","subtitle":"короткий підзаголовок або порожній рядок",
      "layout":"split","content":["змістовна теза 1","змістовна теза 2","змістовна теза 3"],
      "highlight":"головний висновок або порожній рядок","left_title":"ліва колонка",
      "right_title":"права колонка","notes":"Нотатки доповідача: короткий сценарій пояснення слайда"}}
  ]
}}

ДОЗВОЛЕНІ layout: cover, split, cards, timeline, steps, quote, comparison, focus, closing.
КРИТИЧНО ВАЖЛИВО:
1. Слайд 1 — обкладинка з оригінальним гачком. Останній — конкретний підсумок, рішення або заклик до дії.
2. Побудуй цілісну розповідь за вказаною архітектурою. Кожен слайд має нову функцію: контекст, доказ/приклад, механізм, наслідок, рішення тощо.
3. Не використовуй банальні заголовки «Вступ», «Що таке…», «Переваги», «Висновки», якщо вони не є справді необхідними.
4. Не повторюй одну тезу іншими словами. Додавай конкретні приклади, пояснення, сценарії та доречні порівняння саме для цієї теми.
5. Чергуй композиції. Не став однаковий layout на сусідніх слайдах. Уникай однакових списків на кожному слайді.
6. У comparison дай реальні дві сторони порівняння та назви їх у left_title/right_title. У timeline — послідовність подій/етапів. У steps — логічні дії. У quote — сильну коротку тезу. У cards — незалежні блоки.
7. Обирай візуальну логіку, що підтримує саме цю тему; не підміняй зміст випадковим декором.
8. Пиши природно, грамотно, без вигаданих статистичних даних, дат, джерел і цитат. Якщо число невідоме — не вигадуй його.
9. Рівно {count} слайдів. Заголовки мають бути різними та конкретними. Кожен content — 2–5 тез, бажано коротких.
10. Мова всього тексту слайдів і нотаток — {language}. Поза JSON нічого не додавай.
11. Для кожного слайда створи notes: 2–4 речення для доповідача з поясненням, прикладом і природним переходом до наступного слайда. Не повторюй дослівно текст на слайді й не вигадуй джерела.
""".strip()
    models = list(dict.fromkeys([m for m in (GEMINI_MODEL, GEMINI_FALLBACK_MODEL) if m]))
    last_error = None
    for index, model in enumerate(models):
        try:
            parsed = await request_gemini(model, prompt)
            raw_slides = parsed.get("slides")
            if not isinstance(raw_slides, list) or len(raw_slides) != count:
                raise GeminiResponseError(f"Gemini створив не {count} слайдів.")
            normalized = []
            for n, slide in enumerate(raw_slides, 1):
                if not isinstance(slide, dict):
                    raise GeminiResponseError(f"Слайд {n} має неправильний формат.")
                points = slide.get("content", [])
                if isinstance(points, str):
                    points = [points]
                if not isinstance(points, list):
                    points = []
                points = [str(p).strip()[:220] for p in points if str(p).strip()][:5]
                if not points:
                    points = [{"uk": "Ключова інформація за темою презентації.", "ru": "Ключевая информация по теме презентации.", "en": "Key information about the presentation topic."}[data.get("language", "uk")]]
                layout = str(slide.get("layout", "")).lower().strip()
                if layout not in LAYOUTS:
                    layout = ""
                normalized.append({
                    "title": str(slide.get("title", {"uk": f"Слайд {n}", "ru": f"Слайд {n}", "en": f"Slide {n}"}[data.get("language", "uk")])).strip()[:150] or {"uk": f"Слайд {n}", "ru": f"Слайд {n}", "en": f"Slide {n}"}[data.get("language", "uk")],
                    "subtitle": str(slide.get("subtitle", "")).strip()[:180],
                    "content": points,
                    "highlight": str(slide.get("highlight", "")).strip()[:240],
                    "left_title": str(slide.get("left_title", "")).strip()[:80],
                    "right_title": str(slide.get("right_title", "")).strip()[:80],
                    "notes": str(slide.get("notes", "")).strip()[:1400],
                    "layout": layout,
                })
            assign_layouts(normalized, visual_seed)
            return {
                "title": str(parsed.get("title", topic)).strip()[:180],
                "slides": normalized,
                "model": model,
                "visual_seed": visual_seed,
                "variation_id": variation_id,
                "architecture": architecture,
                "visual_concept": visual_concept,
            }
        except (GeminiRetryableError, GeminiModelUnavailable, GeminiResponseError) as exc:
            last_error = exc
            log.warning("Gemini model %s failed: %s", model, exc)
            if index < len(models) - 1:
                continue
        except GeminiAPIError:
            raise
    raise GeminiAPIError(f"Не вдалося отримати презентацію. Остання помилка: {last_error}")

# 7. SLIDE LAYOUT ENGINE — композиції PowerPoint
# ============================================================
def assign_layouts(slides, seed=None):
    """Create a varied narrative rhythm for each deck, avoiding template-like repetition."""
    rng = random.Random(seed if seed is not None else random.SystemRandom().randint(1, 2_000_000_000))
    middle_pool = ["cards", "split", "timeline", "steps", "quote", "comparison", "focus"]
    previous = None
    used_counts = {layout: 0 for layout in middle_pool}
    for i, item in enumerate(slides):
        if i == 0:
            chosen = "cover"
        elif i == len(slides) - 1:
            chosen = "closing"
        else:
            title_and_body = (item.get("title", "") + " " + " ".join(item.get("content", []))).casefold()
            preferred = ""
            if any(w in title_and_body for w in ("етап", "крок", "процес", "step", "stage", "шаг", "алгоритм")):
                preferred = "steps"
            elif any(w in title_and_body for w in ("істор", "розвит", "еволюц", "timeline", "хронолог", "истор", "рік", "рік")):
                preferred = "timeline"
            elif any(w in title_and_body for w in ("переваг", "недолік", "порівн", "відмін", "vs", "різниц", "сравнен")):
                preferred = "comparison"
            candidate = item.get("layout")
            # Accept the model's meaningful layout only when it does not create a run.
            if candidate in middle_pool and candidate != previous and used_counts.get(candidate, 0) < 2:
                chosen = candidate
            else:
                choices = [preferred] if preferred and preferred != previous and used_counts.get(preferred, 0) < 2 else []
                choices += [x for x in middle_pool if x != previous and used_counts.get(x, 0) < 2 and x not in choices]
                if not choices:
                    choices = [x for x in middle_pool if x != previous]
                chosen = rng.choice(choices or ["split"])
            used_counts[chosen] = used_counts.get(chosen, 0) + 1
        item["layout"] = chosen
        previous = chosen



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
    available_width_pt = max(12.0, width_in * 72.0 - 7.0)
    available_height_pt = max(10.0, height_in * 72.0 - 4.0)
    for candidate in range(int(round(size)), min_size - 1, -1):
        chars_per_line = max(6, int(available_width_pt / max(4.0, candidate * 0.52)))
        max_lines = max(1, int(available_height_pt / max(8.0, candidate * 1.22)))
        if _estimate_wrapped_lines(text, chars_per_line) <= max_lines:
            return str(text), candidate
    candidate = min_size
    chars_per_line = max(6, int(available_width_pt / max(4.0, candidate * 0.52)))
    max_lines = max(1, int(available_height_pt / max(8.0, candidate * 1.22)))
    lines = []
    for paragraph in str(text).splitlines() or [""]:
        lines.extend(textwrap.wrap(paragraph, width=chars_per_line, break_long_words=True) or [""])
    if len(lines) > max_lines:
        lines = lines[:max_lines]
        lines[-1] = lines[-1].rstrip(" .,:;-") + "…"
    return "\n".join(lines), candidate


def add_text(slide, text, x, y, w, h, size, color, bold=False,
             align=PP_ALIGN.LEFT, font="Arial", valign=MSO_ANCHOR.TOP):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.margin_left = Inches(.04)
    tf.margin_right = Inches(.04)
    tf.margin_top = Inches(.02)
    tf.margin_bottom = Inches(.02)
    tf.vertical_anchor = valign
    fitted_text, fitted_size = _fit_pptx_text(text, size, w, h)
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
        add_text(slide, point, x + .28, yy, w - .3, total_h - .02, text_size, theme["text"])
 
 
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


def add_style_chrome(slide, design, theme, index, total, rng):
    """Give each selected design a different visual identity, not just a new palette."""
    a, c, m = theme["accent"], theme["card"], theme["muted"]
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
        add_text(slide, "STRATEGY  /  INSIGHT", .62, .55, 4.5, .22, 8, m, True)
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


def add_header(slide, index, total, theme, title, subtitle="", design="minimal"):
    fonts = {
        "luxury": "Georgia", "magazine": "Georgia", "classic": "Georgia",
        "future": "Arial", "vibrant": "Arial", "mono": "Arial",
    }
    font = fonts.get(design, "Arial")
    title_size = 28 if design in ("minimal", "business", "classic") else 30
    if design == "magazine":
        title_size = 32
    elif design == "luxury":
        title_size = 30
    elif design == "vibrant":
        title_size = 33
    elif design == "mono":
        title_size = 31
    if len(title) > 62:
        title_size = max(21, title_size - 8)
    elif len(title) > 43:
        title_size = max(23, title_size - 5)
    label = {
        "minimal": "IDEA  /  EDITION", "future": "FUTURE SYSTEMS  /  LAB",
        "business": "STRATEGY  /  BRIEFING", "school": "LEARNING  /  CREATIVE",
        "luxury": "PREMIUM EDITION", "magazine": "THE KNOWLEDGE REVIEW",
        "nature": "NATURE  /  PERSPECTIVE", "vibrant": "CREATIVE CULTURE",
        "classic": "ACADEMIC EDITION", "mono": "MONOCHROME  /  NOTES",
    }.get(design, "PREZENT AI")
    left_x = 1.38 if design == "mono" else .72
    title_y = 1.03 if design in ("luxury", "classic") else .82
    title_w = 10.8 if design not in ("magazine",) else 9.8
    add_text(slide, label, left_x, .46, 6.8, .22, 8, theme["accent"], True, font="Arial")
    add_text(slide, title, left_x, title_y, title_w, .78, title_size, theme["text"], True, font=font)
    if subtitle:
        add_text(slide, subtitle, left_x + .03, title_y + .82, title_w - .2, .38, 11, theme["muted"], False, font=font)
    if design == "luxury":
        add_shape(slide, MSO_SHAPE.RECTANGLE, 5.45, 1.88, 2.25, .025, theme["accent"])
    elif design == "school":
        add_shape(slide, MSO_SHAPE.OVAL, left_x, 1.72, .16, .16, (255, 196, 72))
    elif design == "magazine":
        add_shape(slide, MSO_SHAPE.RECTANGLE, left_x, 1.72, 1.4, .075, theme["accent"])
    elif design == "future":
        add_shape(slide, MSO_SHAPE.RECTANGLE, left_x, 1.72, 2.1, .035, theme["accent"])
    elif design == "mono":
        add_shape(slide, MSO_SHAPE.RECTANGLE, left_x, 1.76, 2.3, .11, theme["accent"])
    else:
        add_shape(slide, MSO_SHAPE.RECTANGLE, left_x, 1.75, 1.15, .045, theme["accent"])
    add_text(slide, f"{index + 1:02d} / {total:02d}", 11.55, 6.93, .95, .25, 8, theme["muted"], True, PP_ALIGN.RIGHT)


def add_cover_slide(slide, item, theme, design, total, rng):
    title = item["title"]
    subtitle = item.get("subtitle", "")
    accent, text_color, muted = theme["accent"], theme["text"], theme["muted"]
    if design == "minimal":
        add_text(slide, "PREZENT AI  /  ORIGINAL DECK", .9, .8, 5.8, .3, 10, accent, True)
        add_text(slide, title, .9, 1.85, 7.0, 2.55, 43, text_color, True, font="Arial")
        if subtitle: add_text(slide, subtitle, .95, 4.55, 6.3, .8, 17, muted)
        add_shape(slide, MSO_SHAPE.RECTANGLE, 9.0, 1.6, .035, 4.2, accent)
        add_text(slide, "01", 9.4, 2.05, 2.3, 1.2, 72, accent, True, font="Arial")
        add_text(slide, "A CLEAR IDEA.\nA STRONG STORY.", 9.45, 3.6, 2.5, .9, 12, text_color, True)
    elif design == "future":
        add_text(slide, "PREZENT AI  //  FUTURE LAB", .9, .82, 6.8, .3, 10, accent, True)
        add_text(slide, title, .9, 1.7, 7.2, 2.65, 40, text_color, True, font="Arial")
        if subtitle: add_text(slide, subtitle, .95, 4.5, 6.2, .75, 16, muted)
        add_outline(slide, MSO_SHAPE.OVAL, 8.55, 1.35, 3.65, 3.65, accent, 2)
        add_outline(slide, MSO_SHAPE.OVAL, 9.15, 1.95, 2.45, 2.45, (107, 91, 255), 2)
        add_shape(slide, MSO_SHAPE.OVAL, 10.1, 2.9, .55, .55, accent)
        add_text(slide, "01  /  SIGNAL", 8.95, 5.35, 3.4, .35, 12, accent, True, PP_ALIGN.CENTER)
    elif design == "business":
        add_shape(slide, MSO_SHAPE.RECTANGLE, 8.7, 0, 4.63, 7.5, (13, 25, 40))
        add_shape(slide, MSO_SHAPE.RECTANGLE, 8.7, 0, .15, 7.5, accent)
        add_text(slide, "PREZENT AI  /  EXECUTIVE BRIEF", .9, .95, 6.5, .3, 10, accent, True)
        add_text(slide, title, .9, 1.9, 7.0, 2.3, 39, text_color, True, font="Arial")
        if subtitle: add_text(slide, subtitle, .95, 4.45, 6.6, .8, 15, muted)
        for j, label in enumerate(("CONTEXT", "INSIGHT", "ACTION")):
            yy = 1.9 + j * 1.25
            add_shape(slide, MSO_SHAPE.RECTANGLE, 9.35, yy, 2.85, .76, theme["card"])
            add_shape(slide, MSO_SHAPE.RECTANGLE, 9.35, yy, .08, .76, accent)
            add_text(slide, f"0{j+1}   {label}", 9.62, yy + .22, 2.3, .3, 13, text_color, True)
    elif design == "school":
        add_shape(slide, MSO_SHAPE.OVAL, 8.6, .85, 3.2, 3.2, (255, 224, 169))
        add_shape(slide, MSO_SHAPE.OVAL, 9.65, 1.85, 1.65, 1.65, accent)
        add_shape(slide, MSO_SHAPE.DIAMOND, 11.25, 4.35, 1.1, 1.1, (255, 196, 72))
        add_text(slide, "LET'S EXPLORE", .9, .85, 5.7, .35, 11, accent, True)
        add_text(slide, title, .9, 1.65, 7.25, 2.4, 40, text_color, True, font="Arial")
        if subtitle: add_text(slide, subtitle, .95, 4.35, 6.7, .8, 16, muted)
        add_text(slide, "DISCOVER  •  THINK  •  CREATE", .95, 6.15, 6.6, .3, 11, accent, True)
    elif design == "luxury":
        add_outline(slide, MSO_SHAPE.RECTANGLE, .75, .75, 11.85, 5.95, accent, 1.4)
        add_text(slide, "P R E Z E N T   A I", 3.2, 1.35, 6.9, .4, 11, accent, True, PP_ALIGN.CENTER, "Georgia")
        add_text(slide, title, 1.5, 2.2, 10.3, 1.75, 40, text_color, False, PP_ALIGN.CENTER, "Georgia", MSO_ANCHOR.MIDDLE)
        if subtitle: add_text(slide, subtitle, 2.0, 4.35, 9.3, .75, 15, muted, False, PP_ALIGN.CENTER, "Georgia")
        add_shape(slide, MSO_SHAPE.RECTANGLE, 5.6, 5.55, 2.15, .025, accent)
        add_text(slide, "AN ORIGINAL EDITION", 4.3, 5.82, 4.7, .3, 9, accent, True, PP_ALIGN.CENTER, "Georgia")
    elif design == "magazine":
        add_shape(slide, MSO_SHAPE.RECTANGLE, 8.55, 0, 4.78, 7.5, (232, 222, 208))
        add_text(slide, "THE KNOWLEDGE REVIEW  /  SPECIAL ISSUE", .82, .8, 7.3, .35, 9, accent, True)
        add_text(slide, title, .82, 1.55, 7.15, 2.85, 39, text_color, True, font="Georgia")
        if subtitle: add_text(slide, subtitle, .9, 4.65, 6.6, .72, 15, muted, False, font="Georgia")
        add_text(slide, "01", 9.0, 1.0, 3.55, 2.7, 104, accent, True, PP_ALIGN.CENTER, "Georgia")
        add_text(slide, "FEATURE\nSTORY", 9.45, 4.4, 2.7, 1.0, 24, text_color, True, PP_ALIGN.CENTER, "Georgia")
    elif design == "nature":
        add_shape(slide, MSO_SHAPE.OVAL, 8.5, .6, 3.75, 5.0, (218, 232, 207))
        leaf1 = add_shape(slide, MSO_SHAPE.OVAL, 9.25, 1.2, 1.0, 3.6, accent); leaf1.rotation = 35
        leaf2 = add_shape(slide, MSO_SHAPE.OVAL, 10.5, 1.4, .9, 3.4, (120, 157, 104)); leaf2.rotation = -35
        add_text(slide, "FIELD NOTES  /  ECO PERSPECTIVE", .9, .85, 7.0, .3, 10, accent, True)
        add_text(slide, title, .9, 1.75, 7.1, 2.3, 40, text_color, True, font="Georgia")
        if subtitle: add_text(slide, subtitle, .95, 4.35, 6.5, .85, 15, muted)
        add_text(slide, "OBSERVE  •  UNDERSTAND  •  RESTORE", .95, 6.15, 6.7, .3, 10, accent, True)
    elif design == "vibrant":
        add_shape(slide, MSO_SHAPE.RECTANGLE, 8.6, 0, 4.73, 7.5, (53, 38, 92))
        d = add_shape(slide, MSO_SHAPE.DIAMOND, 9.1, .85, 2.5, 2.5, (0, 220, 210)); d.rotation = 14
        add_shape(slide, MSO_SHAPE.OVAL, 10.4, 3.0, 1.75, 1.75, (255, 91, 145))
        add_shape(slide, MSO_SHAPE.RECTANGLE, 9.05, 5.65, 3.3, .18, (255, 197, 72))
        add_text(slide, "IDEAS IN MOTION", .9, .85, 6.8, .35, 11, (0, 220, 210), True)
        add_text(slide, title, .9, 1.7, 7.1, 2.65, 42, text_color, True, font="Arial")
        if subtitle: add_text(slide, subtitle, .95, 4.65, 6.5, .75, 16, muted)
    elif design == "classic":
        add_outline(slide, MSO_SHAPE.RECTANGLE, 1.05, .78, 11.2, 5.85, accent, 1.3)
        add_text(slide, "PREZENT AI  •  ACADEMIC PRESENTATION", 2.0, 1.25, 9.3, .3, 10, accent, True, PP_ALIGN.CENTER, "Georgia")
        add_text(slide, title, 1.7, 2.2, 9.9, 1.85, 37, text_color, True, PP_ALIGN.CENTER, "Georgia", MSO_ANCHOR.MIDDLE)
        if subtitle: add_text(slide, subtitle, 2.05, 4.25, 9.2, .8, 15, muted, False, PP_ALIGN.CENTER, "Georgia")
        add_shape(slide, MSO_SHAPE.RECTANGLE, 5.8, 5.55, 1.75, .025, accent)
    else:  # hard-edge monochrome / brutalist cover
        add_shape(slide, MSO_SHAPE.RECTANGLE, .72, .7, 7.3, 5.95, (22, 22, 22))
        add_text(slide, "PREZENT AI  /  MONO EDITION", 1.1, 1.15, 6.3, .3, 10, (220, 220, 220), True)
        add_text(slide, title, 1.08, 2.0, 6.25, 2.45, 39, (255, 255, 255), True, font="Arial")
        if subtitle: add_text(slide, subtitle, 1.15, 4.85, 5.9, .8, 15, (208, 208, 208))
        add_text(slide, "01", 8.75, 1.45, 3.2, 2.4, 105, (24, 24, 24), True, PP_ALIGN.CENTER, "Arial")
        add_shape(slide, MSO_SHAPE.RECTANGLE, 8.95, 4.25, 2.8, .16, (24, 24, 24))
        add_text(slide, "IDEA / FORM / IMPACT", 8.95, 4.65, 3.0, .35, 11, text_color, True)


def add_fade_transition(slide):
    transition = OxmlElement("p:transition")
    transition.set("spd", "med")
    transition.append(OxmlElement("p:fade"))
    slide._element.insert(2, transition)


def create_pptx(data: dict, design: str, output_path: str):
    theme = DESIGNS[design]
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    total = len(data["slides"])
    rng = random.Random(int(data.get("visual_seed") or random.randint(1, 2_000_000_000)))
    for index, item in enumerate(data["slides"]):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = rgb(theme["bg"])
        layout = item.get("layout", "split")
        add_style_chrome(slide, design, theme, index, total, rng)
        if layout == "cover":
            add_cover_slide(slide, item, theme, design, total, rng)
        elif layout == "closing":
            if design in ("luxury", "classic"):
                add_outline(slide, MSO_SHAPE.RECTANGLE, .8, .75, 11.75, 5.95, theme["accent"], 1.2)
                add_text(slide, "THE TAKEAWAY", 2.0, 1.15, 9.3, .3, 10, theme["accent"], True, PP_ALIGN.CENTER, "Georgia")
                add_text(slide, item["title"], 1.45, 1.85, 10.45, 1.0, 32, theme["text"], True, PP_ALIGN.CENTER, "Georgia")
                add_text(slide, "\n".join(item["content"][:4]), 2.0, 3.15, 9.3, 2.0, 17, theme["text"], False, PP_ALIGN.CENTER, "Georgia")
                if item.get("highlight"):
                    add_text(slide, item["highlight"], 2.0, 5.55, 9.3, .55, 15, theme["accent"], True, PP_ALIGN.CENTER, "Georgia")
            else:
                add_header(slide, index, total, theme, item["title"], item.get("subtitle", ""), design)
                add_panel(slide, design, .82, 2.35, 11.55, 3.75, theme, index)
                add_bullets(slide, item["content"][:4], 1.25, 2.75, 10.65, 2.8, theme, 18)
                if item.get("highlight"):
                    add_text(slide, item["highlight"], 1.1, 6.2, 10.9, .45, 14, theme["accent"], True, PP_ALIGN.CENTER)
        elif layout == "cards":
            add_header(slide, index, total, theme, item["title"], item.get("subtitle", ""), design)
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
                    add_text(slide, "FEATURE  /  01", 1.2, 2.55, 4.8, .28, 10, theme["accent"], True, font="Georgia")
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
                    add_text(slide, f"INSIGHT  0{j + 1}", x + .28, y + .18, 4.7, .25, 9, theme["accent"], True)
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
            add_header(slide, index, total, theme, item["title"], item.get("subtitle", ""), design)
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
                        add_text(slide, f"CHAPTER 0{j + 1}", 2.15, y + .02, 2.0, .24, 9, theme["accent"], True, font="Georgia")
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
            add_header(slide, index, total, theme, item["title"], item.get("subtitle", ""), design)
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
            add_header(slide, index, total, theme, item["title"], item.get("subtitle", ""), design)
            quote = item.get("highlight") or item["content"][0]
            add_text(slide, "“", .92, 2.05, 1.0, 1.2, 76, theme["accent"], True, font="Georgia")
            add_panel(slide, design, 1.82, 2.45, 10.25, 1.72, theme, index)
            add_text(slide, quote, 2.25, 2.73, 9.35, 1.15, 23, theme["text"], True, PP_ALIGN.CENTER, "Georgia" if design in ("luxury", "magazine", "classic") else "Arial", MSO_ANCHOR.MIDDLE)
            rest = item["content"][1:]
            if rest:
                add_bullets(slide, rest[:3], 1.45, 4.65, 10.45, 1.45, theme, 15)
        elif layout == "comparison":
            add_header(slide, index, total, theme, item["title"], item.get("subtitle", ""), design)
            pts = item["content"]
            left, right = pts[::2], pts[1::2]
            add_panel(slide, design, .82, 2.25, 5.55, 3.95, theme, index)
            add_panel(slide, design, 6.95, 2.25, 5.55, 3.95, theme, index + 1)
            add_text(slide, item.get("left_title") or "A", 1.18, 2.58, 4.8, .5, 19, theme["accent"], True, PP_ALIGN.CENTER)
            add_text(slide, item.get("right_title") or "B", 7.3, 2.58, 4.8, .5, 19, theme["accent"], True, PP_ALIGN.CENTER)
            add_bullets(slide, left or pts[:1], 1.2, 3.3, 4.85, 2.5, theme, 15)
            add_bullets(slide, right or pts[-1:], 7.35, 3.3, 4.8, 2.5, theme, 15)
        elif layout == "focus":
            add_header(slide, index, total, theme, item["title"], item.get("subtitle", ""), design)
            focus = item.get("highlight") or item["content"][0]
            add_text(slide, focus, .95, 2.15, 11.2, 1.5, 29, theme["accent"], True, PP_ALIGN.CENTER, "Georgia" if design in ("luxury", "magazine", "classic") else "Arial", MSO_ANCHOR.MIDDLE)
            add_shape(slide, MSO_SHAPE.RECTANGLE, 2.1, 3.9, 9.1, .04, theme["accent"])
            add_bullets(slide, item["content"][1:4] or item["content"][:3], 1.7, 4.35, 9.9, 1.55, theme, 15)
        else:  # split
            add_header(slide, index, total, theme, item["title"], item.get("subtitle", ""), design)
            add_panel(slide, design, .82, 2.25, 7.05, 4.05, theme, index)
            add_bullets(slide, item["content"][:4], 1.22, 2.65, 6.3, 3.35, theme, 17)
            if design in ("future", "vibrant"):
                add_outline(slide, MSO_SHAPE.OVAL, 8.75, 2.55, 2.75, 2.75, theme["accent"], 2)
                add_outline(slide, MSO_SHAPE.OVAL, 9.25, 3.05, 1.75, 1.75, theme["accent"], 1.2)
                add_text(slide, f"{index + 1:02d}", 9.35, 3.55, 1.6, .6, 28, theme["accent"], True, PP_ALIGN.CENTER)
            elif design == "magazine":
                add_text(slide, f"{index + 1:02d}", 8.5, 2.3, 3.3, 1.6, 82, theme["accent"], True, PP_ALIGN.CENTER, "Georgia")
                add_shape(slide, MSO_SHAPE.RECTANGLE, 8.7, 4.2, 2.8, .1, theme["accent"])
            elif design == "nature":
                leaf = add_shape(slide, MSO_SHAPE.OVAL, 9.1, 2.65, 1.8, 3.0, theme["accent"])
                leaf.rotation = 32
            elif design == "mono":
                add_shape(slide, MSO_SHAPE.RECTANGLE, 8.75, 2.7, 2.75, 2.75, theme["accent"])
                add_text(slide, f"{index + 1:02d}", 9.0, 3.45, 2.2, .75, 38, (255, 255, 255), True, PP_ALIGN.CENTER)
            else:
                add_shape(slide, MSO_SHAPE.OVAL, 8.8, 2.75, 2.55, 2.55, theme["accent"])
                add_text(slide, f"{index + 1:02d}", 9.25, 3.55, 1.65, .65, 31, theme["bg"], True, PP_ALIGN.CENTER)
            if item.get("highlight"):
                add_text(slide, item["highlight"], 8.2, 5.65, 3.8, .6, 13, theme["accent"], True, PP_ALIGN.CENTER)
        try:
            slide.notes_slide.notes_text_frame.text = str(item.get("notes", "")).strip()
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
    """Export a style-led PDF with different editorial compositions per slide layout."""
    theme = DESIGNS[design]
    bg = colors.HexColor("#%02x%02x%02x" % theme["bg"])
    accent = colors.HexColor("#%02x%02x%02x" % theme["accent"])
    card = colors.HexColor("#%02x%02x%02x" % theme["card"])
    text_color = colors.HexColor("#%02x%02x%02x" % theme["text"])
    muted = colors.HexColor("#%02x%02x%02x" % theme["muted"])
    if "PrezentSans" not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont("PrezentSans", font_path))

    page_w, page_h = landscape(A4)
    pdf = canvas.Canvas(output_path, pagesize=(page_w, page_h))
    seed = int(data.get("visual_seed") or random.randint(1, 2_000_000_000))
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

    def para(text, x, top, width, style, max_height=130):
        raw_text = str(text)
        base_size = float(getattr(style, "fontSize", 11.0))
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
        for font_size in [base_size - step for step in range(0, max(1, int(base_size - 6.5) + 1))]:
            if font_size < 6.5:
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
                candidate, height = make_paragraph(sample, 6.5)
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
            after = para(point, x + 14, current_top, width - 17, style, 78)
            current_top = after - 9
            if current_top < max_bottom:
                break
        return current_top

    def draw_header(index, item):
        label = {
            "minimal": "IDEA / ORIGINAL EDITION", "future": "FUTURE SYSTEMS / LAB",
            "business": "STRATEGY / EXECUTIVE BRIEF", "school": "LEARNING / CREATIVE",
            "luxury": "PREMIUM EDITION", "magazine": "THE KNOWLEDGE REVIEW",
            "nature": "NATURE / PERSPECTIVE", "vibrant": "CREATIVE CULTURE",
            "classic": "ACADEMIC EDITION", "mono": "MONOCHROME / NOTES",
        }.get(design, "PREZENT AI")
        pdf.setFillColor(accent)
        pdf.setFont("PrezentSans", 7.5)
        pdf.drawString(48 if design != "mono" else 60, page_h - 28, label)
        pdf.setFillColor(muted)
        pdf.setFont("PrezentSans", 8)
        pdf.drawRightString(page_w - 30, 25, f"{index + 1:02d} / {total:02d}")
        x = 60 if design == "mono" else 44
        width = page_w - x - (105 if design in ("magazine", "vibrant") else 82)
        top = page_h - 60
        bottom = para(item["title"], x, top, width, title_style, 82)
        subtitle = item.get("subtitle", "")
        if subtitle:
            para(subtitle, x + 1, bottom - 8, width - 4, small_style, 30)
        return x, width

    def draw_cover(index, item):
        title = item["title"]
        subtitle = item.get("subtitle", "")
        if design == "luxury":
            pdf.setStrokeColor(accent); pdf.setLineWidth(1.1)
            pdf.rect(44, 48, page_w - 88, page_h - 96, fill=0, stroke=1)
            pdf.setFillColor(accent); pdf.setFont("PrezentSans", 10)
            pdf.drawCentredString(page_w / 2, page_h - 100, "P R E Z E N T   A I")
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
            pdf.drawCentredString(page_w / 2, page_h - 130, "PREZENT AI  •  ACADEMIC EDITION")
            style = ParagraphStyle("classic-cover", parent=title_style, fontSize=29, leading=36, alignment=TA_CENTER)
            para(title, 115, page_h - 175, page_w - 230, style, 145)
            if subtitle: para(subtitle, 145, page_h - 345, page_w - 290, small_style, 65)
        elif design == "mono":
            pdf.setFillColor(colors.HexColor("#161616")); pdf.rect(55, 65, 430, 465, fill=1, stroke=0)
            pdf.setFillColor(colors.white); pdf.setFont("PrezentSans", 8)
            pdf.drawString(78, 490, "PREZENT AI  /  MONO EDITION")
            mono_style = ParagraphStyle("mono-cover", parent=title_style, fontSize=30, leading=36, textColor=colors.white)
            para(title, 78, 445, 370, mono_style, 170)
            if subtitle:
                mono_sub = ParagraphStyle("mono-sub", parent=small_style, textColor=colors.HexColor("#D0D0D0"))
                para(subtitle, 82, 245, 355, mono_sub, 70)
            pdf.setFillColor(colors.HexColor("#181818")); pdf.setFont("PrezentSans", 105)
            pdf.drawCentredString(650, 300, "01")
            pdf.setFont("PrezentSans", 10); pdf.drawCentredString(650, 210, "IDEA / FORM / IMPACT")
        else:
            x = 52
            title_w = 430 if design in ("business", "magazine", "vibrant") else 470
            if design == "magazine":
                pdf.setFillColor(colors.HexColor("#E8DED0")); pdf.rect(page_w - 230, 0, 230, page_h, fill=1, stroke=0)
                pdf.setFillColor(accent); pdf.setFont("PrezentSans", 105); pdf.drawCentredString(page_w - 115, 295, "01")
                pdf.setFillColor(text_color); pdf.setFont("PrezentSans", 17); pdf.drawCentredString(page_w - 115, 215, "FEATURE STORY")
            elif design == "business":
                pdf.setFillColor(colors.HexColor("#0D1928")); pdf.rect(page_w - 220, 0, 220, page_h, fill=1, stroke=0)
                for j, tag in enumerate(("CONTEXT", "INSIGHT", "ACTION")):
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
                pdf.setFillColor(text_color); pdf.setFont("PrezentSans", 9); pdf.drawString(565, 265, "A CLEAR IDEA / A STRONG STORY")
            pdf.setFillColor(accent); pdf.setFont("PrezentSans", 9)
            pdf.drawString(x, page_h - 105, "PREZENT AI  /  ORIGINAL PRESENTATION")
            cover_style = ParagraphStyle("cover-main", parent=title_style, fontSize=33 if design in ("school", "vibrant", "magazine") else 30, leading=39, textColor=text_color)
            para(title, x, page_h - 145, title_w, cover_style, 180)
            if subtitle:
                para(subtitle, x + 2, page_h - 340, title_w - 10, small_style, 80)
        pdf.setFillColor(muted); pdf.setFont("PrezentSans", 7.5)
        pdf.drawString(50, 24, "PREZENT AI STUDIO  •  ORIGINAL DECK")

    for i, item in enumerate(data["slides"]):
        draw_chrome(i)
        layout = item.get("layout", "split")
        if layout == "cover":
            draw_cover(i, item)
            pdf.showPage()
            continue

        x0, title_w = draw_header(i, item)
        points = item.get("content", [])[:5]
        if layout == "timeline":
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
                pdf.drawString(60, 365, "FEATURE / 01")
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
                    pdf.drawString(px + 18, py + card_h - 30, f"INSIGHT  0{j + 1}")
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
                style = ParagraphStyle("stepline", parent=body_style, fontSize=10.8, leading=13.5)
                para(point, 106, y - 10, page_w - 175, style, 46)
        elif layout == "quote":
            highlight = item.get("highlight") or (points[0] if points else item["title"])
            draw_panel(64, 205, page_w - 128, 165, i)
            para("“" + highlight + "”", 95, 330, page_w - 190, focus_style, 110)
            draw_bullets(points[1:4], 95, 175, page_w - 190, 75, 10.5)
        elif layout == "focus":
            highlight = item.get("highlight") or (points[0] if points else item["title"])
            para(highlight, 75, 385, page_w - 150, focus_style, 100)
            pdf.setStrokeColor(accent); pdf.setLineWidth(1.7)
            pdf.line(140, 275, page_w - 140, 275)
            draw_panel(70, 75, page_w - 140, 165, i)
            draw_bullets(points[1:4] or points[:3], 95, 215, page_w - 190, 90, 11)
        elif layout == "closing":
            draw_panel(58, 82, page_w - 116, 315, i)
            para(item["title"], 88, 365, page_w - 176, focus_style, 68)
            draw_bullets(points[:4], 110, 285, page_w - 220, 125, 11.5)
            if item.get("highlight"):
                para(item["highlight"], 105, 125, page_w - 210, card_title_style, 45)
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
                pdf.setFillColor(accent); pdf.circle(665, 245, 62, fill=1, stroke=0)
                pdf.setFillColor(bg); pdf.setFont("PrezentSans", 28); pdf.drawCentredString(665, 235, f"{i + 1:02d}")
            if item.get("highlight"):
                para(item["highlight"], 560, 125, 210, card_title_style, 70)
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
    register_user(actual_user)
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
    register_user(actual_user)
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
    register_user(message.from_user)
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
    register_user(callback.from_user)
    set_user_language(callback.from_user.id, language)
    await state.clear()
    await callback.answer(tr(language, "language_saved").replace("<b>", "").replace("</b>", ""))
    try:
        await callback.message.edit_text(tr(language, "language_saved"))
    except Exception:
        pass
    await callback.message.answer(
        tr(language, "welcome", name=html.escape(callback.from_user.first_name or "friend")),
        reply_markup=main_keyboard(language),
    )


@dp.message(Command("restart"))
async def command_restart(message: Message, state: FSMContext):
    """A restart always asks the user to confirm/select the conversation language again."""
    await cancel_active_generation(message.from_user.id)
    await state.clear()
    register_user(message.from_user)
    previous_language = get_user_language(message.from_user.id)
    await show_language_picker(message, previous_language)


@dp.message(Command("cancel"))
async def command_cancel(message: Message, state: FSMContext):
    await cancel_active_generation(message.from_user.id)
    await state.clear()
    register_user(message.from_user)
    language = get_user_language(message.from_user.id)
    if not language:
        await show_language_picker(message)
        return
    await message.answer(tr(language, "action_cancelled"), reply_markup=main_keyboard(language))


@dp.message(Command("id"))
async def command_id(message: Message):
    register_user(message.from_user)
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
    with db_connect() as con:
        users = con.execute("SELECT COUNT(*) n FROM users").fetchone()["n"]
        gens = con.execute("SELECT COUNT(*) n FROM generation_history WHERE status='completed'").fetchone()["n"]
        pending = con.execute("SELECT COUNT(*) n FROM reports WHERE status='pending'").fetchone()["n"]
        revenue = con.execute("SELECT COALESCE(SUM(amount),0) n FROM orders WHERE status='paid'").fetchone()["n"]
    await message.answer(f"🛠 <b>Статистика</b>\nКористувачів: {users}\nГотових презентацій: {gens}\nСкарг очікує: {pending}\nОтримано Stars: {revenue} ⭐")
 
@dp.message(Command("grant"))
async def command_grant(message: Message):
    if not is_admin(message.from_user.id):
        await message.answer("⛔ Немає доступу.")
        return
    parts = (message.text or "").split()
    if len(parts) != 3 or not parts[1].isdigit() or not parts[2].lstrip("-").isdigit():
        await message.answer("Формат: <code>/grant ID КІЛЬКІСТЬ</code>")
        return
    uid, amount = int(parts[1]), int(parts[2])
    with db_connect() as con:
        con.execute("INSERT INTO users(user_id, created_at) VALUES(?,?) ON CONFLICT(user_id) DO NOTHING", (uid, now_iso()))
        con.execute("UPDATE users SET credits=credits+? WHERE user_id=?", (amount, uid))
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

    file_name = (getattr(document, "file_name", None) or "backup.sqlite3").lower()
    if not file_name.endswith((".sqlite3", ".sqlite", ".db", ".db3")):
        await message.answer("❌ Надішли резервну копію SQLite у форматі .sqlite3, .sqlite або .db.")
        return

    await message.answer("⏳ Перевіряю резервну копію та відновлюю дані в Turso. Не перезапускай бота до завершення.")
    try:
        with tempfile.TemporaryDirectory() as tmp:
            dest = str(Path(tmp) / "restore.sqlite3")
            await bot.download(document, destination=dest)
            await asyncio.to_thread(restore_db_from, dest)
        await asyncio.to_thread(init_db)
        with db_connect() as con:
            users_row = con.execute("SELECT COUNT(*) AS n FROM users").fetchone()
            users_count = int(users_row["n"] if users_row else 0)
            credits_row = con.execute("SELECT COALESCE(SUM(credits),0) AS n FROM users").fetchone()
            total_credits = int(credits_row["n"] if credits_row else 0)
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

    pending = PENDING_RESTORE_DOCUMENTS.get(message.from_user.id)
    if not pending:
        await message.answer(
            "Спочатку надішли файл <code>prezentai_backup.sqlite3</code> як документ, "
            "а потім окремим повідомленням команду <code>/restore</code>.\n\n"
            "Також можна надіслати файл із підписом <code>/restore</code>."
        )
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
            await asyncio.to_thread(restore_db_from, dest)
        await asyncio.to_thread(init_db)
        with db_connect() as con:
            users_row = con.execute("SELECT COUNT(*) AS n FROM users").fetchone()
            users_count = int(users_row["n"] if users_row else 0)
            credits_row = con.execute("SELECT COALESCE(SUM(credits),0) AS n FROM users").fetchone()
            total_credits = int(credits_row["n"] if credits_row else 0)
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
        register_user(message.from_user)
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
        register_user(message.from_user)
        await message.answer(profile_text(message.from_user.id, language), reply_markup=main_keyboard(language))
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
    register_user(message.from_user)
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
    register_user(callback.from_user)
    language = get_user_language(callback.from_user.id) or "uk"
    await callback.message.answer(profile_text(callback.from_user.id, language), reply_markup=main_keyboard(language))
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
    register_user(callback.from_user)
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
    payload = create_order(callback.from_user.id, product, amount)
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
        order = find_order(query.invoice_payload)
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
        ok, result = process_payment(
            payment.invoice_payload, message.from_user.id, payment.total_amount,
            payment.currency, payment.telegram_payment_charge_id,
        )
    except Exception as storage_exc:
        log.exception("Could not apply a successful Stars payment because storage failed")
        refund_ok = False
        try:
            refund_ok = await bot.refund_star_payment(
                user_id=message.from_user.id,
                telegram_payment_charge_id=payment.telegram_payment_charge_id,
            )
        except Exception:
            log.exception("Automatic refund failed after payment database error")
        for admin_id in ADMIN_IDS:
            try:
                await bot.send_message(
                    admin_id,
                    "🚨 <b>Не вдалося записати успішний платіж у базу.</b> "
                    f"Користувач: <code>{message.from_user.id}</code>; "
                    f"{payment.total_amount} ⭐; "
                    f"автоповернення: {'успішне' if refund_ok else 'НЕ вдалося'}; "
                    f"charge: <code>{html.escape(payment.telegram_payment_charge_id)}</code>; "
                    f"помилка: {html.escape(str(storage_exc)[:300])}"
                )
            except Exception:
                log.exception("Could not notify admin about payment database failure")
        user_text = (
            {
                "uk": "⚠️ Не вдалося зберегти оплату. Telegram Stars автоматично повернуто. Спробуй ще раз пізніше.",
                "ru": "⚠️ Не удалось сохранить оплату. Telegram Stars автоматически возвращены. Попробуй позже.",
                "en": "⚠️ The payment could not be stored. Telegram Stars were automatically refunded. Please try again later.",
            }[language] if refund_ok else {
                "uk": "⚠️ Оплату отримано, але базу даних тимчасово недоступна. Адміністратора сповіщено для ручного повернення. Не повторюй платіж.",
                "ru": "⚠️ Оплата получена, но база данных временно недоступна. Администратор уведомлён для ручного возврата. Не повторяй платёж.",
                "en": "⚠️ Payment was received, but the database is temporarily unavailable. The administrator has been alerted to arrange a manual refund. Do not pay again.",
            }[language]
        )
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
        text = tr(language, "payment_credits", credits=CREDIT_PACK_SIZE, profile=profile_text(message.from_user.id, language))
    elif result == "pro":
        text = tr(language, "payment_pro", days=PRO_DAYS, limit=PRO_DAILY_LIMIT, profile=profile_text(message.from_user.id, language))
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
    with db_connect() as con:
        row = con.execute(
            "SELECT user_id FROM generation_history WHERE generation_id=?", (generation_id,)
        ).fetchone()
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
    report_id = create_report(generation_id, message.from_user.id, reason)
    await state.clear()
    if not report_id:
        await message.answer(tr(language, "report_missing"), reply_markup=main_keyboard(language))
        return
    report = get_report(report_id)
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
    result = handle_report(report_id, callback.from_user.id, action)
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
    await state.set_state(PresentationForm.design)
    await callback.message.answer(tr(language_ui, "design_prompt"), reply_markup=design_keyboard(language_ui))
    await callback.answer()
 
 
@dp.callback_query(F.data.startswith("design:"))
async def design_handler(callback: CallbackQuery, state: FSMContext):
    user_id = callback.from_user.id
    interface_lang = get_user_language(user_id) or "uk"
    ui = interface_lang
    if await state.get_state() != PresentationForm.design.state:
        await callback.answer(tr(ui, "design_prompt").replace("<b>", "").replace("</b>", ""), show_alert=True)
        return
    if user_id in ACTIVE_GENERATIONS:
        waiting = {"uk": "Твоя презентація вже створюється.", "ru": "Твоя презентация уже создаётся.", "en": "Your presentation is already being created."}[ui]
        await callback.answer(waiting, show_alert=True)
        return
    design = callback.data.split(":", 1)[1]
    if design not in DESIGNS:
        await callback.answer({"uk": "Непідтримуваний стиль.", "ru": "Неподдерживаемый стиль.", "en": "Unsupported style."}[ui], show_alert=True)
        return
    data = await state.get_data()
    if not all(k in data for k in ("topic", "slides", "language")):
        await state.clear()
        await callback.answer({"uk": "Налаштування втрачено. Почни спочатку.", "ru": "Настройки потеряны. Начни сначала.", "en": "Settings were lost. Please start again."}[ui], show_alert=True)
        return
    quota = reserve_generation(user_id)
    if not quota["allowed"]:
        await callback.answer({"uk": "Денний ліміт вичерпано.", "ru": "Дневной лимит исчерпан.", "en": "Daily limit reached."}[ui], show_alert=True)
        await callback.message.answer(
            tr(ui, "limit_reached", limit=FREE_DAILY_LIMIT, credits=CREDIT_PACK_SIZE, price=CREDIT_PACK_PRICE),
            reply_markup=shop_keyboard(ui),
        )
        return
    if quota.get("mode") == "credit":
        schedule_balance_backup("списання придбаного кредиту")

    data["design"] = design
    data["interface_language"] = ui
    await state.clear()
    await callback.answer({"uk": "Додав у чергу на створення!", "ru": "Добавил в очередь на создание!", "en": "Added to the generation queue!"}[ui])

    progress = ProgressAnimator(language=ui)
    ACTIVE_GENERATIONS[user_id] = asyncio.current_task()
    generation_id = uuid.uuid4().hex[:16]
    succeeded = False
    status = None
    pdf_sent = False
    try:
        queue_text = {
            "uk": "⏳ <b>Ваша презентація в черзі.</b> Зараз одночасно створюються дві презентації. Щойно звільниться місце, почнеться генерація.",
            "ru": "⏳ <b>Ваша презентация в очереди.</b> Сейчас одновременно создаются две презентации. Как только освободится место, начнётся генерация.",
            "en": "⏳ <b>Your presentation is in the queue.</b> Two presentations are currently being generated. Your generation will start as soon as a slot is free.",
        }[ui]
        if GEN_SEMAPHORE.locked():
            status = await callback.message.answer(queue_text, reply_markup=main_keyboard(ui))

        async with GEN_SEMAPHORE:
            if status is None:
                status = await callback.message.answer(progress.render(), reply_markup=main_keyboard(ui))
            else:
                await status.edit_text(progress.render(), parse_mode=ParseMode.HTML, reply_markup=main_keyboard(ui))
            progress.message = status
            progress.set_stage("writing")
            progress.start()

            # A hard wall-clock cap covers the primary and fallback Gemini models together.
            presentation_data = await asyncio.wait_for(generate_content(data), timeout=150)

            with tempfile.TemporaryDirectory(prefix="prezent_ai_") as tmp:
                pptx_path = str(Path(tmp) / "PrezentAI.pptx")
                pdf_path = str(Path(tmp) / "PrezentAI.pdf")

                progress.set_stage("pptx")
                await asyncio.to_thread(create_pptx, presentation_data, design, pptx_path)
                if not Path(pptx_path).exists() or Path(pptx_path).stat().st_size <= 0:
                    raise RuntimeError("PowerPoint file was not created or is empty.")

                pdf_ready, pdf_error = False, None
                progress.set_stage("pdf")
                try:
                    font_path = await ensure_unicode_font()
                    await asyncio.to_thread(create_pdf, presentation_data, design, pdf_path, font_path)
                    pdf_ready = Path(pdf_path).exists() and Path(pdf_path).stat().st_size > 0
                except Exception as exc:
                    pdf_error = str(exc)
                    log.exception("PDF generation failed")

                add_history(generation_id, user_id, data["topic"], design, data["slides"], quota["mode"], "completed")
                progress.set_stage("sending")
                pptx_message = await callback.message.answer_document(
                    FSInputFile(pptx_path, filename="PrezentAI.pptx"),
                    caption=tr(ui, "file_caption", title=html.escape(presentation_data["title"]), style=html.escape(design_name(design, ui))),
                    reply_markup=report_user_keyboard(generation_id, ui),
                )
                succeeded = True
                set_history_message(generation_id, pptx_message.chat.id, pptx_message.message_id)

                if pdf_ready:
                    try:
                        await callback.message.answer_document(
                            FSInputFile(pdf_path, filename="PrezentAI.pdf"),
                            caption=tr(ui, "pdf_caption"),
                            reply_markup=main_keyboard(ui),
                        )
                        pdf_sent = True
                    except Exception:
                        log.exception("PDF could not be sent")
                if not pdf_sent:
                    await callback.message.answer(tr(ui, "pdf_failed"), reply_markup=main_keyboard(ui))

        await progress.complete(tr(ui, "generation_complete") if pdf_sent else tr(ui, "generation_pptx_only"))
        await callback.message.answer(tr(ui, "ready_next"), reply_markup=main_keyboard(ui))
    except asyncio.CancelledError:
        log.info("Generation cancelled by user %s", user_id)
        if not succeeded:
            release_generation(user_id, quota)
            mark_history_failed(generation_id)
            await progress.finish(tr(ui, "cancelled"))
        else:
            await progress.finish(tr(ui, "already_sent"))
        raise
    except asyncio.TimeoutError:
        if not succeeded:
            release_generation(user_id, quota)
            mark_history_failed(generation_id)
            schedule_balance_backup("повернення кредиту після тайм-ауту")
        timeout_text = {
            "uk": "⌛ Генерація перевищила ліміт часу 150 секунд. Кредит або денний ліміт повернуто. Спробуй ще раз трохи пізніше.",
            "ru": "⌛ Генерация превысила лимит времени 150 секунд. Кредит или дневной лимит возвращён. Попробуй ещё раз позже.",
            "en": "⌛ Generation exceeded the 150-second limit. Your credit or daily quota has been restored. Please try again later.",
        }[ui]
        await progress.finish(timeout_text)
    except GeminiAPIError as exc:
        if not succeeded:
            release_generation(user_id, quota)
            mark_history_failed(generation_id)
            schedule_balance_backup("повернення кредиту після помилки Gemini")
        log.warning("Gemini generation failed: %s", exc)
        await progress.finish(tr(ui, "ai_failed", reason=html.escape(str(exc))))
        await callback.message.answer(tr(ui, "retry_menu"), reply_markup=main_keyboard(ui))
    except Exception:
        error_id = uuid.uuid4().hex[:8].upper()
        log.exception("Presentation generation failed; error_id=%s", error_id)
        if not succeeded:
            release_generation(user_id, quota)
            mark_history_failed(generation_id)
            schedule_balance_backup("повернення кредиту після помилки генерації")
            await progress.finish(tr(ui, "generic_failed", error_id=error_id))
            await callback.message.answer(tr(ui, "check_logs"), reply_markup=main_keyboard(ui))
        else:
            await progress.finish(
                {"uk": "⚠️ Презентацію вже надіслано, але під час завершення сталася помилка.",
                 "ru": "⚠️ Презентация уже отправлена, но при завершении возникла ошибка.",
                 "en": "⚠️ The presentation was sent, but a final step failed."}[ui]
                + f"\n<code>{error_id}</code>"
            )
    finally:
        await progress.stop()
        ACTIVE_GENERATIONS.pop(user_id, None)

@dp.message(F.text)
async def unknown_text(message: Message):
    register_user(message.from_user)
    language = get_user_language(message.from_user.id)
    if not language:
        await show_language_picker(message)
        return
    await message.answer(tr(language, "unknown"), reply_markup=main_keyboard(language))
 
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
        with db_connect() as con:
            row = con.execute("SELECT COUNT(*) AS n FROM users").fetchone()
            user_count = int(row["n"] if row else 0)
    except Exception:
        log.exception("Could not inspect users table at startup")
        user_count = -1

    alerts = []
    if RENDER_IS_HOSTED and not DURABLE_DB_ENABLED:
        alerts.append(
            "🚨 <b>КРИТИЧНО: постійне сховище не підключене.</b> "
            "Render Free стирає локальну SQLite при redeploy/restart/spin-down. "
            "Покупки Telegram Stars автоматично вимкнені, щоб не втратити гроші користувачів. "
            "Підключи Turso і додай TURSO_DATABASE_URL та TURSO_AUTH_TOKEN у Render Environment."
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


async def schedule_backend_initialization(app: web.Application):
    async def initialize_backend():
        try:
            # Do database/network initialization only after aiohttp has opened PORT.
            await asyncio.to_thread(init_db)
            await on_startup(bot)
            app["startup_ready"].set()
            app["backup_task"] = asyncio.create_task(backup_loop())
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