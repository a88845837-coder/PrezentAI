import asyncio
import hashlib
import html
import io
import json
import logging
import os
import random
import re
import sqlite3
import tempfile
import uuid
import zipfile
from datetime import datetime, timedelta, timezone
from pathlib import Path
from xml.sax.saxutils import escape
from zoneinfo import ZoneInfo

import httpx
from aiohttp import web
from aiogram import Bot, Dispatcher, F
from aiogram.enums import ParseMode
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
DATABASE_URL = os.getenv("DATABASE_URL", "").strip()
RENDER_IS_HOSTED = bool(os.getenv("RENDER_EXTERNAL_URL", "").strip())
MAX_TOPIC_LENGTH = 500
FREE_DAILY_LIMIT = 3
PRO_DAILY_LIMIT = 10
CREDIT_PACK_SIZE = 5
CREDIT_PACK_PRICE = 15
PRO_PRICE = 99
PRO_DAYS = 30
KYIV_TZ = ZoneInfo("Europe/Kyiv")

if not BOT_TOKEN:
    raise RuntimeError("BOT_TOKEN is missing in Render Environment.")
if not GEMINI_API_KEY:
    raise RuntimeError("GEMINI_API_KEY is missing in Render Environment.")

if not DATABASE_URL:
    Path(DATABASE_PATH).expanduser().resolve().parent.mkdir(parents=True, exist_ok=True)
bot = Bot(token=BOT_TOKEN, default=DefaultBotProperties(parse_mode=ParseMode.HTML))
dp = Dispatcher(storage=MemoryStorage())
ACTIVE_GENERATIONS = {}

# ============================================================
# 2. DATABASE
# ============================================================
class PostgreSQLConnection:
    """Tiny adapter so the same queries work with SQLite or hosted Postgres."""
    def __init__(self, connection):
        self.connection = connection

    def execute(self, sql, params=()):
        sql = sql.replace("BEGIN IMMEDIATE", "BEGIN")
        sql = sql.replace("?", "%s")
        return self.connection.execute(sql, params or ())

    def executescript(self, script):
        for statement in script.split(";"):
            statement = statement.strip()
            if statement:
                self.execute(statement)

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, tb):
        try:
            if exc_type:
                self.connection.rollback()
            else:
                self.connection.commit()
        finally:
            self.connection.close()
        return False


def db_connect():
    if DATABASE_URL:
        try:
            import psycopg
            from psycopg.rows import dict_row
        except ImportError as exc:
            raise RuntimeError("DATABASE_URL задано, але psycopg не встановлено. Додай psycopg[binary] у requirements.txt.") from exc
        connection = psycopg.connect(DATABASE_URL, row_factory=dict_row, connect_timeout=15, autocommit=True)
        return PostgreSQLConnection(connection)
    conn = sqlite3.connect(DATABASE_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA foreign_keys=ON")
    return conn


def init_db():
    with db_connect() as con:
        con.executescript("""
        CREATE TABLE IF NOT EXISTS users (
            user_id BIGINT PRIMARY KEY,
            username TEXT DEFAULT '',
            first_name TEXT DEFAULT '',
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
        """)


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
        con.execute("BEGIN IMMEDIATE")
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
        con.execute("BEGIN IMMEDIATE")
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
        con.execute("BEGIN IMMEDIATE")
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
# 3. KEYBOARDS AND THEMES
# ============================================================
BTN_START = "🚀 Старт"
BTN_RESTART = "🔄 Перезапустити"
BTN_CREATE = "🎨 Створити презентацію"
BTN_HOME = "🏠 Головне меню"
BTN_CANCEL = "❌ Скасувати"
BTN_FEATURES = "✨ Можливості"
BTN_HELP = "❓ Допомога"
BTN_PROFILE = "👤 Мій профіль"
BTN_SHOP = "⭐ Магазин"
BTN_SUB = "👑 Підписка PRO"

MAIN_KEYBOARD = ReplyKeyboardMarkup(
    keyboard=[
        [KeyboardButton(text=BTN_START), KeyboardButton(text=BTN_RESTART)],
        [KeyboardButton(text=BTN_CREATE)],
        [KeyboardButton(text=BTN_PROFILE), KeyboardButton(text=BTN_SHOP)],
        [KeyboardButton(text=BTN_SUB)],
        [KeyboardButton(text=BTN_HOME), KeyboardButton(text=BTN_CANCEL)],
        [KeyboardButton(text=BTN_FEATURES), KeyboardButton(text=BTN_HELP)],
    ], resize_keyboard=True, is_persistent=True, input_field_placeholder="Обери дію в меню..."
)

# Layout key: color palette + compositional style. These are real design variations,
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

class PresentationForm(StatesGroup):
    topic = State()
    slides = State()
    language = State()
    design = State()

class ReportForm(StatesGroup):
    reason = State()


def slides_keyboard():
    return InlineKeyboardMarkup(inline_keyboard=[[
        InlineKeyboardButton(text=str(n), callback_data=f"slides:{n}") for n in (6, 8, 10, 12)
    ]])


def language_keyboard():
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="🇺🇦 Українська", callback_data="lang:uk"),
         InlineKeyboardButton(text="🇬🇧 English", callback_data="lang:en")],
        [InlineKeyboardButton(text="🇷🇺 Русский", callback_data="lang:ru")],
    ])


def design_keyboard():
    rows, row = [], []
    for key, theme in DESIGNS.items():
        row.append(InlineKeyboardButton(text=theme["name"], callback_data=f"design:{key}"))
        if len(row) == 2:
            rows.append(row)
            row = []
    if row:
        rows.append(row)
    return InlineKeyboardMarkup(inline_keyboard=rows)


def shop_keyboard():
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="🎟️ 5 генерацій — 15 ⭐", callback_data="buy:credits")],
        [InlineKeyboardButton(text="👑 PRO на 30 днів — 99 ⭐", callback_data="buy:pro")],
        [InlineKeyboardButton(text="👤 Мій профіль", callback_data="profile")],
    ])


def report_admin_keyboard(report_id: str):
    return InlineKeyboardMarkup(inline_keyboard=[[
        InlineKeyboardButton(text="✅ Повернути 1 кредит", callback_data=f"reportact:{report_id}:refund"),
        InlineKeyboardButton(text="❌ Відхилити", callback_data=f"reportact:{report_id}:reject"),
    ]])


def report_user_keyboard(generation_id: str):
    return InlineKeyboardMarkup(inline_keyboard=[[
        InlineKeyboardButton(text="⚠️ Повідомити про брак", callback_data=f"bad:{generation_id}")
    ]])

# ============================================================
# 4. TEXTS AND PROFILE
# ============================================================
WELCOME_TEXT = (
    "🚀 <b>Prezent AI</b>\n\n"
    "Створюй презентації у PowerPoint та PDF з різними макетами й стилями.\n\n"
    "🎨 10 дизайнерських тем\n📊 6, 8, 10 або 12 слайдів\n"
    "🧠 Структура й текст від Google Gemini\n📄 PPTX + PDF\n"
    "🎞️ Переходи Fade між слайдами\n\n"
    "🎁 Безкоштовно: 3 презентації на день.\n"
    "⭐ Додаткові генерації можна придбати в магазині.\n"
    "Обери дію в меню нижче 👇"
)
FEATURES_TEXT = (
    "✨ <b>Можливості Prezent AI</b>\n\n"
    "🎨 10 тем: Мінімалізм, Футуризм, Business Pro, Шкільний креатив, "
    "Luxury, Журнал, Eco, Яскравий, Класичний і Чорно-білий.\n"
    "🧩 Різні макети: обкладинка, картки, таймлайн, кроки, цитата, "
    "порівняння, великий акцент і фінальний слайд.\n"
    "📊 6/8/10/12 слайдів • 🇺🇦 🇬🇧 🇷🇺 • PPTX + PDF.\n"
    "⚠️ Кнопка «Повідомити про брак» надсилає презентацію адміну на перевірку."
)
HELP_TEXT = (
    "❓ <b>Допомога</b>\n\n"
    "🚀 Старт — головне меню.\n🎨 Створити презентацію — запустити генератор.\n"
    "👤 Мій профіль — ліміти, кредити та PRO.\n⭐ Магазин — придбати генерації.\n"
    "👑 Підписка PRO — 10 презентацій на день протягом 30 днів.\n"
    "❌ Скасувати — зупинити створення.\n\n"
    "/start — меню • /id — показати твій Telegram ID"
)


def format_until(value):
    if not value:
        return "не активна"
    try:
        dt = datetime.fromisoformat(value).astimezone(KYIV_TZ)
        return dt.strftime("%d.%m.%Y о %H:%M")
    except ValueError:
        return "не активна"


def profile_text(user_id: int):
    row = get_user(user_id)
    if not row:
        return "Профіль ще не створено. Натисни /start."
    used = get_usage(user_id)
    pro = subscription_active(row)
    limit = "∞" if is_admin(user_id) else (PRO_DAILY_LIMIT if pro else FREE_DAILY_LIMIT)
    remaining = "∞" if is_admin(user_id) else max(0, int(limit) - used)
    with db_connect() as con:
        total = con.execute(
            "SELECT COUNT(*) AS n FROM generation_history WHERE user_id=? AND status='completed'",
            (user_id,),
        ).fetchone()["n"]
    if is_admin(user_id):
        daily = "∞ (адміністратор)"
    else:
        daily = f"{used}/{limit} використано • залишилось {remaining}"
    sub = f"активна до {format_until(row['subscription_until'])}" if pro else "не активна"
    return (
        "👤 <b>Твій профіль Prezent AI</b>\n\n"
        f"🆔 ID: <code>{user_id}</code>\n"
        f"📊 Презентацій створено: <b>{total}</b>\n"
        f"🎁 Денний ліміт: <b>{daily}</b>\n"
        f"🎟️ Додаткові кредити: <b>{row['credits']}</b>\n"
        f"👑 Підписка PRO: <b>{sub}</b>\n\n"
        "Кредити витрачаються після використання денного ліміту."
    )

# ============================================================
# 5. ANIMATION — збережена початкова анімація з твого файлу
# ============================================================
class ProgressAnimator:
    FRAMES = ("⏳", "⌛", "🔄", "⚙️")

    def __init__(self, message: Message):
        self.message = message
        self.stage = "Підготовка запиту"
        self.percent = 4
        self.target = 22
        self.frame = 0
        self.task = None
        self.lock = asyncio.Lock()
        self.running = False

    def set_stage(self, text: str, target: int):
        self.stage = text
        self.target = max(self.percent, min(98, target))

    def render(self):
        filled = max(0, min(10, self.percent // 10))
        bar = "🟦" * filled + "⬜" * (10 - filled)
        icon = self.FRAMES[self.frame % len(self.FRAMES)]
        return (
            f"{icon} <b>Prezent AI створює презентацію</b>\n\n"
            f"{self.stage}\n{bar} <b>{self.percent}%</b>\n\n"
            "<i>Будь ласка, зачекай — це може зайняти кілька хвилин.</i>"
        )

    async def safe_edit(self):
        try:
            async with self.lock:
                await self.message.edit_text(self.render())
        except Exception as exc:
            if "message is not modified" not in str(exc).lower():
                log.debug("Progress edit skipped: %s", exc)

    async def run(self):
        self.running = True
        while self.running:
            self.frame += 1
            if self.percent < self.target:
                self.percent = min(self.target, self.percent + 1)
            await self.safe_edit()
            await asyncio.sleep(1.6)

    def start(self):
        self.task = asyncio.create_task(self.run())

    async def stop(self):
        self.running = False
        if self.task:
            self.task.cancel()
            try:
                await self.task
            except asyncio.CancelledError:
                pass

    async def finish(self, text: str):
        await self.stop()
        try:
            await self.message.edit_text(text)
        except Exception:
            log.exception("Could not update progress message")

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
        "generationConfig": {"maxOutputTokens": 8192, "responseMimeType": "application/json"},
    }
    for attempt in range(1, 4):
        try:
            async with httpx.AsyncClient(timeout=httpx.Timeout(150.0, connect=25.0)) as client:
                response = await client.post(url, headers=headers, json=payload)
        except httpx.TimeoutException as exc:
            if attempt < 3:
                await asyncio.sleep(2 ** attempt + random.uniform(0, .8))
                continue
            raise GeminiRetryableError("Gemini не відповів вчасно після трьох спроб.") from exc
        except httpx.HTTPError as exc:
            if attempt < 3:
                await asyncio.sleep(2 ** attempt)
                continue
            raise GeminiRetryableError("Помилка з'єднання з Google Gemini.") from exc
        try:
            result = response.json()
        except ValueError as exc:
            raise GeminiAPIError(f"Gemini повернув неправильну відповідь. HTTP {response.status_code}.") from exc
        status = response.status_code
        if status in (408, 429, 500, 502, 503, 504):
            if attempt < 3:
                retry_after = response.headers.get("Retry-After")
                try:
                    delay = float(retry_after) if retry_after else 2 ** attempt
                except ValueError:
                    delay = 2 ** attempt
                await asyncio.sleep(min(delay, 12) + random.uniform(0, .8))
                continue
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
            if attempt < 3:
                await asyncio.sleep(2 ** attempt)
                continue
            raise GeminiRetryableError("Gemini тричі повернув порожню відповідь.")
        return extract_json(content)
    raise GeminiRetryableError(f"Модель {model} не повернула результат.")


async def generate_content(data: dict) -> dict:
    topic, count = data["topic"], data["slides"]
    language = LANGUAGES.get(data["language"], "українською")
    style_name = DESIGNS[data["design"]]["name"]
    prompt = f"""
Створи яскраву, професійну й НЕ шаблонну презентацію.
ТЕМА: {topic}
СЛАЙДІВ: рівно {count}
МОВА: {language}
ВІЗУАЛЬНИЙ СТИЛЬ: {style_name}

Поверни тільки JSON такої структури:
{{
  "title":"Назва презентації",
  "slides":[
    {{"title":"Заголовок","subtitle":"короткий підзаголовок або порожній рядок",
      "layout":"split","content":["теза 1","теза 2","теза 3"],
      "highlight":"головна думка слайда","left_title":"ліва колонка",
      "right_title":"права колонка"}}
  ]
}}

ДОЗВОЛЕНІ layout: cover, split, cards, timeline, steps, quote, comparison, focus, closing.
Правила дизайну й контенту:
1. Перший слайд — сильна обкладинка, останній — висновок/заклик/підсумок.
2. Чергуй макети: не використовуй один і той самий layout два слайди поспіль.
3. Обирай композицію за змістом: послідовність/історія — timeline; інструкція — steps;
   дві концепції — comparison; сильна теза — quote або focus; кілька характеристик — cards;
   пояснення — split.
4. Кожен слайд має власну композицію, не роби однакові списки на всіх слайдах.
5. 2–5 коротких змістовних тез на слайд; не перевантажуй текстом.
6. Додавай конкретні приклади, пояснення та логічні зв'язки. Не повторюй тези.
7. Не вигадуй точні статистичні дані, дати, дослідження або цитати. Якщо факт невідомий — не подавай його як факт.
8. Пиши природно, грамотно та зрозуміло. Поля highlight/ subtitle можуть бути порожніми.
9. Рівно {count} слайдів. Поза JSON не додавай нічого.
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
                points = [str(p).strip()[:280] for p in points if str(p).strip()][:5]
                if not points:
                    points = ["Основна інформація за темою презентації."]
                layout = str(slide.get("layout", "")).lower().strip()
                if layout not in LAYOUTS:
                    layout = ""
                normalized.append({
                    "title": str(slide.get("title", f"Слайд {n}")).strip()[:150] or f"Слайд {n}",
                    "subtitle": str(slide.get("subtitle", "")).strip()[:180],
                    "content": points,
                    "highlight": str(slide.get("highlight", "")).strip()[:240],
                    "left_title": str(slide.get("left_title", "")).strip()[:80],
                    "right_title": str(slide.get("right_title", "")).strip()[:80],
                    "layout": layout,
                })
            assign_layouts(normalized)
            return {"title": str(parsed.get("title", topic)).strip()[:180], "slides": normalized, "model": model}
        except (GeminiRetryableError, GeminiModelUnavailable, GeminiResponseError) as exc:
            last_error = exc
            log.warning("Gemini model %s failed: %s", model, exc)
            if index < len(models) - 1:
                continue
        except GeminiAPIError:
            raise
    raise GeminiAPIError(f"Не вдалося отримати презентацію. Остання помилка: {last_error}")

# ============================================================
# 7. SLIDE LAYOUT ENGINE — композиції PowerPoint
# ============================================================
def assign_layouts(slides):
    """Enforce cover/ending and diversify layouts when the AI repeats itself."""
    pool = ["cards", "split", "timeline", "steps", "quote", "comparison", "focus"]
    random.shuffle(pool)
    previous = None
    for i, item in enumerate(slides):
        if i == 0:
            chosen = "cover"
        elif i == len(slides) - 1:
            chosen = "closing"
        else:
            candidate = item.get("layout")
            if candidate in ("cover", "closing"):
                candidate = ""
            if candidate and candidate != previous:
                chosen = candidate
            else:
                title_and_body = (item.get("title", "") + " " + " ".join(item.get("content", []))).lower()
                if any(w in title_and_body for w in ("етап", "крок", "процес", "step", "stage")):
                    preferred = "steps"
                elif any(w in title_and_body for w in ("істор", "розвит", "еволюц", "timeline", "хронолог")):
                    preferred = "timeline"
                elif any(w in title_and_body for w in ("переваг", "недолік", "порівн", "відмін", "vs", "різниц")):
                    preferred = "comparison"
                else:
                    preferred = ""
                choices = [x for x in ([preferred] if preferred else []) + pool if x != previous]
                chosen = random.choice(choices or ["split", "cards", "focus"])
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


def add_text(slide, text, x, y, w, h, size, color, bold=False,
             align=PP_ALIGN.LEFT, font="Aptos", valign=MSO_ANCHOR.TOP):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.margin_left = Inches(.04)
    tf.margin_right = Inches(.04)
    tf.margin_top = Inches(.02)
    tf.margin_bottom = Inches(.02)
    tf.vertical_anchor = valign
    p = tf.paragraphs[0]
    p.text = str(text)
    p.alignment = align
    p.font.name = font
    p.font.size = Pt(size)
    p.font.bold = bold
    p.font.color.rgb = rgb(color)
    return box


def add_bullets(slide, points, x, y, w, h, theme, size=18, gap=10):
    points = points[:5]
    total_h = h / max(1, len(points))
    for i, point in enumerate(points):
        yy = y + i * total_h
        add_shape(slide, MSO_SHAPE.OVAL, x, yy + .12, .12, .12, theme["accent"])
        add_text(slide, point, x + .28, yy, w - .3, total_h - .02,
                 size if len(points) <= 4 else size - 2, theme["text"])


def add_header(slide, index, total, theme, title, subtitle=""):
    add_text(slide, "PREZENT AI  /  PRESENTATION", .62, .28, 5.5, .25, 9, theme["accent"], True)
    add_text(slide, f"{index + 1:02d} / {total:02d}", 11.75, .28, .9, .25, 9, theme["muted"], True, PP_ALIGN.RIGHT)
    add_text(slide, title, .68, .72, 11.8, .78, 27, theme["text"], True)
    if subtitle:
        add_text(slide, subtitle, .72, 1.48, 11.7, .45, 12, theme["muted"])


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
    for index, item in enumerate(data["slides"]):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = rgb(theme["bg"])
        layout = item.get("layout", "split")
        # Small decorative shape system varies by theme.
        if theme["style"] in ("future", "vibrant"):
            add_shape(slide, MSO_SHAPE.OVAL, 10.8, -.8, 3.2, 3.2, theme["card"])
            add_shape(slide, MSO_SHAPE.OVAL, 11.65, .05, 1.3, 1.3, theme["accent"])
        elif theme["style"] in ("magazine", "luxury"):
            add_shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, .18, 7.5, theme["accent"])
        elif theme["style"] == "school":
            add_shape(slide, MSO_SHAPE.OVAL, 11.8, .2, .7, .7, theme["accent"])
        else:
            add_shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, .12, 7.5, theme["accent"])
        if layout == "cover":
            add_text(slide, "PREZENT AI  •  SMART PRESENTATION", .85, .65, 6.8, .3, 11, theme["accent"], True)
            add_text(slide, item["title"], .85, 1.65, 7.4, 2.4, 34, theme["text"], True)
            if item.get("subtitle"):
                add_text(slide, item["subtitle"], .9, 4.15, 6.9, .8, 17, theme["muted"])
            add_shape(slide, MSO_SHAPE.OVAL, 8.65, 1.45, 3.2, 3.2, theme["card"], theme["accent"])
            add_shape(slide, MSO_SHAPE.OVAL, 9.3, 2.1, 1.9, 1.9, theme["accent"])
            add_text(slide, "IDEA\n→\nIMPACT", 9.43, 2.58, 1.6, .9, 16, theme["bg"], True, PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)
            add_text(slide, f"{index + 1:02d} / {total:02d}", 11.7, 6.95, .9, .25, 9, theme["muted"], True, PP_ALIGN.RIGHT)
        elif layout == "closing":
            add_text(slide, "ГОЛОВНЕ ЗА 1 ХВИЛИНУ", .8, .55, 6, .3, 10, theme["accent"], True)
            add_text(slide, item["title"], .8, 1.05, 11.5, .95, 32, theme["text"], True)
            add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, .85, 2.3, 11.5, 3.8, theme["card"], theme["accent"], True)
            add_bullets(slide, item["content"], 1.25, 2.65, 10.7, 2.9, theme, 19)
            if item.get("highlight"):
                add_text(slide, item["highlight"], 1.1, 6.35, 10.9, .5, 15, theme["accent"], True, PP_ALIGN.CENTER)
        elif layout == "cards":
            add_header(slide, index, total, theme, item["title"], item.get("subtitle", ""))
            pts = item["content"][:4]
            cols = 2 if len(pts) > 2 else len(pts)
            cols = max(1, cols)
            rows = (len(pts) + cols - 1) // cols
            cell_w = 5.55 if cols == 2 else 11.0
            cell_h = 1.85 if rows == 2 else 3.65
            for j, point in enumerate(pts):
                col, row = j % cols, j // cols
                x, y = .75 + col * 5.75, 2.25 + row * 2.0
                add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, cell_w, cell_h, theme["card"], theme["accent"], True)
                add_text(slide, f"0{j + 1}", x + .25, y + .18, .6, .35, 15, theme["accent"], True)
                add_text(slide, point, x + .25, y + .62, cell_w - .5, cell_h - .72, 17, theme["text"], True)
        elif layout == "timeline":
            add_header(slide, index, total, theme, item["title"], item.get("subtitle", ""))
            pts = item["content"][:4]
            yline = 4.05
            add_shape(slide, MSO_SHAPE.RECTANGLE, 1.2, yline, 10.8, .055, theme["accent"])
            for j, point in enumerate(pts):
                x = 1.1 + j * (10.7 / max(1, len(pts) - 1))
                add_shape(slide, MSO_SHAPE.OVAL, x, yline - .18, .42, .42, theme["accent"])
                add_text(slide, f"0{j + 1}", x - .05, 2.35, .55, .35, 17, theme["accent"], True, PP_ALIGN.CENTER)
                add_text(slide, point, x - .62, 4.6, 1.75, 1.55, 13, theme["text"], True, PP_ALIGN.CENTER)
        elif layout == "steps":
            add_header(slide, index, total, theme, item["title"], item.get("subtitle", ""))
            pts = item["content"][:5]
            for j, point in enumerate(pts):
                y = 2.15 + j * .86
                add_shape(slide, MSO_SHAPE.OVAL, .9, y, .55, .55, theme["accent"])
                add_text(slide, str(j + 1), .98, y + .08, .38, .3, 13, theme["bg"], True, PP_ALIGN.CENTER)
                add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, 1.7, y - .04, 10.5, .66, theme["card"], None, True)
                add_text(slide, point, 2.0, y + .07, 9.9, .45, 16, theme["text"], True)
        elif layout == "quote":
            add_header(slide, index, total, theme, item["title"], item.get("subtitle", ""))
            quote = item.get("highlight") or item["content"][0]
            add_text(slide, "“", .9, 2.0, 1.0, 1.2, 72, theme["accent"], True)
            add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, 1.8, 2.35, 10.2, 1.65, theme["card"], theme["accent"], True)
            add_text(slide, quote, 2.2, 2.65, 9.4, 1.05, 23, theme["text"], True, PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)
            rest = item["content"][1:]
            if rest:
                add_bullets(slide, rest, 1.25, 4.55, 10.8, 1.6, theme, 15)
        elif layout == "comparison":
            add_header(slide, index, total, theme, item["title"], item.get("subtitle", ""))
            pts = item["content"]
            left, right = pts[::2], pts[1::2]
            add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, .8, 2.25, 5.55, 3.9, theme["card"], theme["accent"], True)
            add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, 6.95, 2.25, 5.55, 3.9, theme["card"], theme["accent"], True)
            add_text(slide, item.get("left_title") or "Підхід A", 1.15, 2.55, 4.8, .5, 19, theme["accent"], True, PP_ALIGN.CENTER)
            add_text(slide, item.get("right_title") or "Підхід B", 7.3, 2.55, 4.8, .5, 19, theme["accent"], True, PP_ALIGN.CENTER)
            add_bullets(slide, left or pts[:1], 1.2, 3.25, 4.8, 2.45, theme, 15)
            add_bullets(slide, right or pts[-1:], 7.35, 3.25, 4.8, 2.45, theme, 15)
        elif layout == "focus":
            add_header(slide, index, total, theme, item["title"], item.get("subtitle", ""))
            focus = item.get("highlight") or item["content"][0]
            add_text(slide, focus, .95, 2.05, 11.2, 1.55, 30, theme["accent"], True, PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)
            add_shape(slide, MSO_SHAPE.RECTANGLE, 2.1, 3.85, 9.1, .04, theme["accent"])
            rest = item["content"][1:] or item["content"]
            add_bullets(slide, rest[:3], 1.7, 4.35, 9.9, 1.55, theme, 15)
        else:  # split
            add_header(slide, index, total, theme, item["title"], item.get("subtitle", ""))
            add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, .8, 2.2, 7.05, 4.2, theme["card"], None, True)
            add_bullets(slide, item["content"][:4], 1.15, 2.58, 6.35, 3.45, theme, 17)
            add_shape(slide, MSO_SHAPE.OVAL, 8.65, 2.6, 2.85, 2.85, theme["accent"])
            add_text(slide, f"{index + 1:02d}", 9.18, 3.4, 1.8, .8, 34, theme["bg"], True, PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)
            if item.get("highlight"):
                add_text(slide, item["highlight"], 8.2, 5.65, 3.8, .6, 13, theme["accent"], True, PP_ALIGN.CENTER)
        add_fade_transition(slide)
    prs.save(output_path)

# ============================================================
# 8. PDF EXPORT
# ============================================================
async def ensure_unicode_font() -> str:
    candidates = [
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/truetype/liberation2/LiberationSans-Regular.ttf",
    ]
    for candidate in candidates:
        if Path(candidate).exists():
            return candidate
    target = Path(tempfile.gettempdir()) / "PrezentAI-DejaVuSans.ttf"
    if target.exists() and target.stat().st_size > 100_000:
        return str(target)
    archive_url = "https://github.com/dejavu-fonts/dejavu-fonts/releases/download/version_2_37/dejavu-fonts-ttf-2.37.zip"
    try:
        async with httpx.AsyncClient(timeout=60.0, follow_redirects=True) as client:
            response = await client.get(archive_url)
            response.raise_for_status()
        with zipfile.ZipFile(io.BytesIO(response.content)) as archive:
            font_name = next((n for n in archive.namelist() if n.endswith("/ttf/DejaVuSans.ttf")), None)
            if not font_name:
                raise RuntimeError("DejaVuSans.ttf not found in archive.")
            font_bytes = archive.read(font_name)
        if not font_bytes.startswith(b"\x00\x01\x00\x00"):
            raise RuntimeError("Downloaded file is not a valid TTF font.")
        target.write_bytes(font_bytes)
        return str(target)
    except Exception as exc:
        log.exception("Unable to download DejaVu font")
        raise RuntimeError("Не вдалося завантажити шрифт для українського PDF.") from exc


def create_pdf(data: dict, design: str, output_path: str, font_path: str):
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
    title_style = ParagraphStyle("title", fontName="PrezentSans", fontSize=22, leading=27, textColor=text_color)
    body_style = ParagraphStyle("body", fontName="PrezentSans", fontSize=13, leading=18, textColor=text_color)
    total = len(data["slides"])
    for i, item in enumerate(data["slides"]):
        pdf.setFillColor(bg); pdf.rect(0, 0, page_w, page_h, fill=1, stroke=0)
        pdf.setFillColor(accent); pdf.rect(0, 0, 10, page_h, fill=1, stroke=0)
        pdf.setFillColor(muted); pdf.setFont("PrezentSans", 8)
        pdf.drawString(34, page_h - 25, "PREZENT AI  /  PRESENTATION")
        pdf.drawRightString(page_w - 28, page_h - 25, f"{i+1:02d} / {total:02d}")
        y_title = page_h - 75
        p = Paragraph(escape(item["title"]), title_style)
        _, ph = p.wrap(page_w - 85, 65); p.drawOn(pdf, 36, y_title - ph)
        layout = item.get("layout", "split")
        if item.get("subtitle"):
            pdf.setFillColor(muted); pdf.setFont("PrezentSans", 10)
            pdf.drawString(38, y_title - ph - 18, item["subtitle"][:140])
        if layout == "cover":
            pdf.setFillColor(card); pdf.roundRect(page_w * .64, 100, 180, 180, 24, fill=1, stroke=0)
            pdf.setFillColor(accent); pdf.circle(page_w * .64 + 90, 190, 52, fill=1, stroke=0)
            content_y = y_title - ph - 60
        elif layout == "timeline":
            y = 250; pdf.setStrokeColor(accent); pdf.setLineWidth(3); pdf.line(60, y, page_w - 60, y)
            pts = item["content"][:4]
            for j, point in enumerate(pts):
                x = 65 + j * ((page_w - 130) / max(1, len(pts)-1))
                pdf.setFillColor(accent); pdf.circle(x, y, 7, fill=1, stroke=0)
                para = Paragraph(escape(point), body_style); _, hh = para.wrap(145, 90); para.drawOn(pdf, x - 55, y - 25 - hh)
            content_y = 80
        elif layout == "comparison":
            pdf.setFillColor(card); pdf.roundRect(35, 95, page_w/2 - 50, 310, 12, fill=1, stroke=0)
            pdf.roundRect(page_w/2 + 15, 95, page_w/2 - 50, 310, 12, fill=1, stroke=0)
            content_y = 360
        elif layout in ("cards", "steps", "closing", "quote", "focus"):
            pdf.setFillColor(card); pdf.roundRect(35, 90, page_w - 70, 315, 12, fill=1, stroke=0)
            content_y = 370 if layout == "cards" else 350
        else:
            pdf.setFillColor(card); pdf.roundRect(35, 90, page_w - 70, 315, 12, fill=1, stroke=0)
            content_y = 350
        points = item["content"][:5]
        if layout == "comparison":
            left, right = points[::2], points[1::2]
            columns = [(48, left), (page_w/2 + 30, right)]
            for x, group in columns:
                yy = 340
                for point in group:
                    para = Paragraph("• " + escape(point), body_style)
                    _, hh = para.wrap(page_w/2 - 75, 80); para.drawOn(pdf, x, yy - hh); yy -= hh + 20
        else:
            if layout in ("quote", "focus") and item.get("highlight"):
                focus_para = Paragraph("<b>" + escape(item["highlight"]) + "</b>", ParagraphStyle("focus", parent=body_style, fontSize=18, leading=24, textColor=accent, alignment=TA_CENTER))
                _, hh = focus_para.wrap(page_w - 120, 80); focus_para.drawOn(pdf, 60, content_y - hh)
                content_y -= hh + 22
                if layout == "quote":
                    points = points[1:]
            for j, point in enumerate(points):
                para = Paragraph((f"<b>{j+1:02d}</b>  " if layout == "steps" else "•  ") + escape(point), body_style)
                _, hh = para.wrap(page_w - 110, 65)
                para.drawOn(pdf, 58, content_y - hh)
                content_y -= hh + 15
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


async def show_home(message: Message, state: FSMContext):
    await cancel_active_generation(message.from_user.id)
    await state.clear()
    register_user(message.from_user)
    await message.answer(WELCOME_TEXT, reply_markup=MAIN_KEYBOARD)


async def start_creation(message: Message, state: FSMContext):
    await cancel_active_generation(message.from_user.id)
    await state.clear()
    register_user(message.from_user)
    await state.set_state(PresentationForm.topic)
    await message.answer(
        "📝 <b>Створення презентації</b>\n\nНапиши тему презентації.\n"
        "Наприклад: «Штучний інтелект у сучасному світі».\n"
        "Максимум 500 символів.", reply_markup=MAIN_KEYBOARD
    )


@dp.message(CommandStart())
async def command_start(message: Message, state: FSMContext):
    await show_home(message, state)

@dp.message(Command("restart"))
async def command_restart(message: Message, state: FSMContext):
    await cancel_active_generation(message.from_user.id)
    await state.clear()
    register_user(message.from_user)
    await message.answer("🔄 <b>Діалог перезапущено!</b>\n\nМожеш почати нову презентацію.", reply_markup=MAIN_KEYBOARD)

@dp.message(Command("cancel"))
async def command_cancel(message: Message, state: FSMContext):
    await cancel_active_generation(message.from_user.id)
    await state.clear()
    await message.answer("❌ Поточну дію скасовано.", reply_markup=MAIN_KEYBOARD)

@dp.message(Command("id"))
async def command_id(message: Message):
    register_user(message.from_user)
    await message.answer(f"Твій Telegram ID: <code>{message.from_user.id}</code>\nДодай його в Render → Environment → ADMIN_IDS.", reply_markup=MAIN_KEYBOARD)

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

@dp.message(F.text == BTN_START)
async def button_start(message: Message, state: FSMContext):
    await show_home(message, state)

@dp.message(F.text == BTN_RESTART)
async def button_restart(message: Message, state: FSMContext):
    await cancel_active_generation(message.from_user.id)
    await state.clear()
    await message.answer("🔄 <b>Діалог перезапущено!</b>", reply_markup=MAIN_KEYBOARD)

@dp.message(F.text == BTN_CREATE)
async def button_create(message: Message, state: FSMContext):
    await start_creation(message, state)

@dp.message(F.text == BTN_HOME)
async def button_home(message: Message, state: FSMContext):
    await show_home(message, state)

@dp.message(F.text == BTN_CANCEL)
async def button_cancel(message: Message, state: FSMContext):
    await cancel_active_generation(message.from_user.id)
    await state.clear()
    await message.answer("❌ Створення скасовано.", reply_markup=MAIN_KEYBOARD)

@dp.message(F.text == BTN_FEATURES)
async def button_features(message: Message):
    await message.answer(FEATURES_TEXT, reply_markup=MAIN_KEYBOARD)

@dp.message(F.text == BTN_HELP)
async def button_help(message: Message):
    await message.answer(HELP_TEXT, reply_markup=MAIN_KEYBOARD)

@dp.message(F.text == BTN_PROFILE)
async def button_profile(message: Message):
    register_user(message.from_user)
    await message.answer(profile_text(message.from_user.id), reply_markup=MAIN_KEYBOARD)

@dp.message(F.text == BTN_SHOP)
async def button_shop(message: Message):
    await message.answer(
        "⭐ <b>Магазин Prezent AI</b>\n\n"
        f"🎟️ {CREDIT_PACK_SIZE} додаткових генерацій — <b>{CREDIT_PACK_PRICE} ⭐</b>\n"
        f"👑 PRO на {PRO_DAYS} днів — <b>{PRO_PRICE} ⭐</b>\n\n"
        "Оплата проходить у Telegram Stars.", reply_markup=shop_keyboard()
    )

@dp.message(F.text == BTN_SUB)
async def button_subscription(message: Message):
    await message.answer(
        "👑 <b>Підписка PRO</b>\n\n"
        f"✅ До {PRO_DAILY_LIMIT} презентацій на день\n"
        f"✅ Діє {PRO_DAYS} днів\n"
        "✅ Пріоритетний доступ не обіцяю — швидкість залежить від Gemini і Render.\n\n"
        f"Ціна: <b>{PRO_PRICE} ⭐</b>. Підписка оплачується один раз і автоматично не продовжується.",
        reply_markup=InlineKeyboardMarkup(inline_keyboard=[[
            InlineKeyboardButton(text=f"👑 Оформити PRO — {PRO_PRICE} ⭐", callback_data="buy:pro")
        ]])
    )

@dp.callback_query(F.data == "profile")
async def callback_profile(callback: CallbackQuery):
    register_user(callback.from_user)
    await callback.message.answer(profile_text(callback.from_user.id), reply_markup=MAIN_KEYBOARD)
    await callback.answer()

@dp.callback_query(F.data == "create")
async def inline_create(callback: CallbackQuery, state: FSMContext):
    await callback.answer()
    await start_creation(callback.message, state)

# ============================================================
# 10. SHOP AND TELEGRAM STARS PAYMENTS
# ============================================================
@dp.callback_query(F.data.startswith("buy:"))
async def buy_handler(callback: CallbackQuery):
    register_user(callback.from_user)
    if RENDER_IS_HOSTED and not DATABASE_URL:
        await callback.answer("Спочатку підключи постійну базу DATABASE_URL.", show_alert=True)
        await callback.message.answer(
            "⚠️ Оплата тимчасово вимкнена, щоб не втратити куплені кредити після перезапуску Render. "
            "Підключи Neon/PostgreSQL через змінну DATABASE_URL і перезапусти сервіс."
        )
        return
    product = callback.data.split(":", 1)[1]
    if product == "credits":
        title = "5 генерацій Prezent AI"
        description = "П'ять додаткових генерацій презентацій."
        amount = CREDIT_PACK_PRICE
    elif product == "pro":
        title = "Prezent AI PRO на 30 днів"
        description = "До 10 презентацій на день протягом 30 днів."
        amount = PRO_PRICE
    else:
        await callback.answer("Невідомий товар.", show_alert=True)
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
        await callback.answer("Рахунок надіслано нижче.")
    except Exception:
        log.exception("Could not send invoice")
        await callback.answer("Не вдалося створити рахунок. Спробуй ще раз.", show_alert=True)


@dp.pre_checkout_query()
async def pre_checkout_handler(query):
    order = find_order(query.invoice_payload)
    ok = bool(
        order
        and int(order["user_id"]) == query.from_user.id
        and int(order["amount"]) == query.total_amount
        and query.currency == "XTR"
        and order["status"] == "pending"
    )
    await query.answer(ok=ok, error_message=None if ok else "Замовлення не знайдено або його параметри не збігаються.")


@dp.message(F.successful_payment)
async def successful_payment_handler(message: Message):
    payment = message.successful_payment
    ok, result = process_payment(
        payment.invoice_payload, message.from_user.id, payment.total_amount,
        payment.currency, payment.telegram_payment_charge_id,
    )
    if not ok:
        log.error("Payment not applied: user=%s payload=%s reason=%s", message.from_user.id, payment.invoice_payload, result)
        await message.answer("⚠️ Платіж отримано, але автоматично застосувати покупку не вдалося. Напиши адміну.")
        for admin_id in ADMIN_IDS:
            try:
                await bot.send_message(admin_id, f"⚠️ Перевір платіж вручну. User: {message.from_user.id}; payload: <code>{html.escape(payment.invoice_payload)}</code>; причина: {html.escape(str(result))}")
            except Exception:
                pass
        return
    if result == "credits":
        text = f"✅ Оплата успішна! Додано {CREDIT_PACK_SIZE} генерацій.\n\n{profile_text(message.from_user.id)}"
    elif result == "pro":
        text = f"👑 PRO активовано на {PRO_DAYS} днів!\n\n{profile_text(message.from_user.id)}"
    else:
        text = "✅ Ця покупка вже була зарахована."
    await message.answer(text, reply_markup=MAIN_KEYBOARD)

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
    await callback.message.answer(
        "⚠️ <b>Повідомлення про брак</b>\n\nОпиши, що саме не так із презентацією: помилки, зіпсовані слайди, неправильний файл або інша проблема.\n\nАдміністратор отримає презентацію та перевірить її.",
        reply_markup=MAIN_KEYBOARD,
    )
    await callback.answer("Чекаю опис проблеми.")


@dp.message(ReportForm.reason)
async def report_reason_handler(message: Message, state: FSMContext):
    reason = (message.text or "").strip()
    if len(reason) < 4:
        await message.answer("Опиши проблему трохи детальніше (від 4 символів).")
        return
    data = await state.get_data()
    generation_id = data.get("report_generation_id")
    source_message_id = data.get("report_source_message_id")
    report_id = create_report(generation_id, message.from_user.id, reason)
    await state.clear()
    if not report_id:
        await message.answer("Не знайшов цю презентацію в історії. Напиши адміну.", reply_markup=MAIN_KEYBOARD)
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
        await message.answer("✅ Скаргу та файл презентації надіслано адміністратору. Після перевірки він зможе повернути 1 кредит.", reply_markup=MAIN_KEYBOARD)
    else:
        await message.answer("⚠️ Скаргу збережено, але адміністратор не отримав повідомлення. Перевір ADMIN_IDS у Render.", reply_markup=MAIN_KEYBOARD)


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
    if status == "refunded":
        try:
            await bot.send_message(user_id, "✅ Адміністратор перевірив презентацію та повернув <b>1 генераційний кредит</b>. Він доступний у твоєму профілі.", reply_markup=MAIN_KEYBOARD)
        except Exception:
            log.warning("Could not notify user %s about credit refund", user_id)
        await callback.message.edit_reply_markup(reply_markup=None)
        await callback.message.answer(f"✅ Кредит повернуто користувачу <code>{user_id}</code>.")
        await callback.answer("Повернено 1 кредит.")
    else:
        try:
            await bot.send_message(user_id, "ℹ️ Адміністратор перевірив скаргу. Повернення кредиту не було схвалено.", reply_markup=MAIN_KEYBOARD)
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
    topic = (message.text or "").strip()
    if len(topic) < 3:
        await message.answer("Тема занадто коротка. Напиши хоча б 3 символи.")
        return
    if len(topic) > MAX_TOPIC_LENGTH:
        await message.answer("Тема занадто довга. Максимум 500 символів.")
        return
    await state.update_data(topic=topic)
    await state.set_state(PresentationForm.slides)
    await message.answer("📊 <b>Скільки слайдів потрібно?</b>\n\nОбери кількість:", reply_markup=slides_keyboard())


@dp.callback_query(F.data.startswith("slides:"))
async def slides_handler(callback: CallbackQuery, state: FSMContext):
    if await state.get_state() != PresentationForm.slides.state:
        await callback.answer("Спочатку введи тему презентації.", show_alert=True)
        return
    try:
        count = int(callback.data.split(":", 1)[1])
    except (ValueError, IndexError):
        await callback.answer("Неправильна кількість слайдів.", show_alert=True)
        return
    if count not in (6, 8, 10, 12):
        await callback.answer("Обери 6, 8, 10 або 12.", show_alert=True)
        return
    await state.update_data(slides=count)
    await state.set_state(PresentationForm.language)
    await callback.message.answer("🌍 <b>Обери мову презентації:</b>", reply_markup=language_keyboard())
    await callback.answer()


@dp.callback_query(F.data.startswith("lang:"))
async def language_handler(callback: CallbackQuery, state: FSMContext):
    if await state.get_state() != PresentationForm.language.state:
        await callback.answer("Спочатку обери кількість слайдів.", show_alert=True)
        return
    language = callback.data.split(":", 1)[1]
    if language not in LANGUAGES:
        await callback.answer("Непідтримувана мова.", show_alert=True)
        return
    await state.update_data(language=language)
    await state.set_state(PresentationForm.design)
    await callback.message.answer("🎨 <b>Обери один із 10 стилів презентації:</b>", reply_markup=design_keyboard())
    await callback.answer()


@dp.callback_query(F.data.startswith("design:"))
async def design_handler(callback: CallbackQuery, state: FSMContext):
    user_id = callback.from_user.id
    if await state.get_state() != PresentationForm.design.state:
        await callback.answer("Спочатку пройди налаштування презентації.", show_alert=True)
        return
    if user_id in ACTIVE_GENERATIONS:
        await callback.answer("Твоя презентація вже створюється.", show_alert=True)
        return
    design = callback.data.split(":", 1)[1]
    if design not in DESIGNS:
        await callback.answer("Непідтримуваний дизайн.", show_alert=True)
        return
    data = await state.get_data()
    if not all(k in data for k in ("topic", "slides", "language")):
        await state.clear()
        await callback.answer("Налаштування втрачено. Почни спочатку.", show_alert=True)
        return
    quota = reserve_generation(user_id)
    if not quota["allowed"]:
        await callback.answer("Денний ліміт вичерпано.", show_alert=True)
        await callback.message.answer(
            f"⛔ <b>Ліміт на сьогодні вичерпано.</b>\n\n"
            f"Без PRO доступно {FREE_DAILY_LIMIT} презентації на день.\n"
            "Можеш придбати 5 додаткових генерацій за 15 ⭐ або активувати PRO.",
            reply_markup=shop_keyboard(),
        )
        return
    data["design"] = design
    await state.clear()
    await callback.answer("Починаю створення!")
    status = await callback.message.answer(
        "⏳ <b>Prezent AI запускається...</b>\n\n🟦⬜⬜⬜⬜⬜⬜⬜⬜⬜ 4%",
        reply_markup=MAIN_KEYBOARD,
    )
    progress = ProgressAnimator(status)
    current_task = asyncio.current_task()
    ACTIVE_GENERATIONS[user_id] = current_task
    progress.start()
    generation_id = uuid.uuid4().hex[:16]
    succeeded = False
    try:
        progress.set_stage("🧠 AI складає план і готує текст слайдів", 62)
        presentation_data = await generate_content(data)
        progress.set_stage("🎨 Створюю різні макети й оформлюю слайди", 82)
        with tempfile.TemporaryDirectory(prefix="prezent_ai_") as tmp:
            pptx_path = str(Path(tmp) / "PrezentAI.pptx")
            pdf_path = str(Path(tmp) / "PrezentAI.pdf")
            await asyncio.to_thread(create_pptx, presentation_data, design, pptx_path)
            if not Path(pptx_path).exists():
                raise RuntimeError("PowerPoint file was not created.")
            progress.set_stage("📄 Готую PDF-файл", 94)
            pdf_ready, pdf_error = False, None
            try:
                font_path = await ensure_unicode_font()
                await asyncio.to_thread(create_pdf, presentation_data, design, pdf_path, font_path)
                pdf_ready = Path(pdf_path).exists()
            except Exception as exc:
                pdf_error = str(exc)
                log.exception("PDF generation failed")
            await progress.finish(
                "✅ <b>Презентація готова!</b>\n\n📤 Надсилаю PowerPoint та PDF..." if pdf_ready
                else "✅ <b>PowerPoint створено!</b>\n\n⚠️ PDF не вдалося підготувати. Надсилаю PPTX."
            )
            add_history(generation_id, user_id, data["topic"], design, data["slides"], quota["mode"], "completed")
            pptx_message = await callback.message.answer_document(
                FSInputFile(pptx_path, filename="PrezentAI.pptx"),
                caption=(
                    f"📊 <b>{html.escape(presentation_data['title'])}</b>\n"
                    f"🎨 Стиль: {html.escape(DESIGNS[design]['name'])}\n"
                    "Формат: PowerPoint • переходи Fade\n\n"
                    "Є проблема? Натисни кнопку нижче — адміністратор перегляне файл."
                ),
                reply_markup=report_user_keyboard(generation_id),
            )
            succeeded = True
            set_history_message(generation_id, pptx_message.chat.id, pptx_message.message_id)
            if pdf_ready:
                await callback.message.answer_document(
                    FSInputFile(pdf_path, filename="PrezentAI.pdf"),
                    caption="📄 <b>PDF</b> • версія для перегляду та друку.",
                    reply_markup=MAIN_KEYBOARD,
                )
            else:
                await callback.message.answer(
                    "PPTX готовий. PDF цього разу не створився. Перевір Render Logs, якщо це повторюється.",
                    reply_markup=MAIN_KEYBOARD,
                )
            if pdf_error:
                log.error("PDF unavailable: %s", pdf_error)
        await callback.message.answer(
            "✅ <b>Готово!</b>\n\nМожеш створити наступну презентацію кнопкою нижче.",
            reply_markup=MAIN_KEYBOARD,
        )
    except asyncio.CancelledError:
        log.info("Generation cancelled by user %s", user_id)
        if not succeeded:
            release_generation(user_id, quota)
            mark_history_failed(generation_id)
        await progress.finish("❌ <b>Створення зупинено.</b>\n\nЛіміт/кредит повернуто.")
        return
    except GeminiAPIError as exc:
        if not succeeded:
            release_generation(user_id, quota)
            mark_history_failed(generation_id)
        log.warning("Gemini generation failed: %s", exc)
        await progress.finish(
            "❌ <b>AI не зміг створити презентацію.</b>\n\n"
            f"Причина: {html.escape(str(exc))}\n\nКредит або денний ліміт повернуто."
        )
        await callback.message.answer("Можеш повторити спробу через меню.", reply_markup=MAIN_KEYBOARD)
    except Exception:
        if not succeeded:
            release_generation(user_id, quota)
            mark_history_failed(generation_id)
        error_id = uuid.uuid4().hex[:8].upper()
        log.exception("Presentation generation failed; error_id=%s", error_id)
        await progress.finish(
            "❌ <b>Не вдалося створити презентацію.</b>\n\n"
            f"Код помилки: <code>{error_id}</code>\nКредит або денний ліміт повернуто."
        )
        await callback.message.answer("Перевір Render Logs або спробуй пізніше.", reply_markup=MAIN_KEYBOARD)
    finally:
        ACTIVE_GENERATIONS.pop(user_id, None)


@dp.message(F.text)
async def unknown_text(message: Message):
    await message.answer("Не зрозумів команду. Скористайся кнопками в меню 👇", reply_markup=MAIN_KEYBOARD)

# ============================================================
# 13. RENDER HEALTH CHECK + WEBHOOK
# ============================================================
async def health_handler(request: web.Request):
    return web.json_response({"status": "ok", "service": "Prezent AI", "mode": "webhook"})

WEBHOOK_PATH = "/telegram/webhook"
RENDER_EXTERNAL_URL = os.getenv("RENDER_EXTERNAL_URL", "").rstrip("/")
WEBHOOK_SECRET = os.getenv("WEBHOOK_SECRET", "").strip()
if not WEBHOOK_SECRET:
    WEBHOOK_SECRET = hashlib.sha256(BOT_TOKEN.encode("utf-8")).hexdigest()


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
        BotCommand(command="restart", description="Перезапустити діалог"),
        BotCommand(command="cancel", description="Скасувати створення"),
        BotCommand(command="id", description="Мій Telegram ID"),
    ])


def main():
    init_db()
    port = int(os.getenv("PORT", "10000"))
    app = web.Application(client_max_size=1024 * 1024)
    app.router.add_get("/", health_handler)
    app.router.add_get("/healthz", health_handler)
    webhook_handler = SimpleRequestHandler(
        dispatcher=dp, bot=bot, handle_in_background=True, secret_token=WEBHOOK_SECRET
    )
    webhook_handler.register(app, path=WEBHOOK_PATH)
    dp.startup.register(on_startup)
    setup_application(app, dp, bot=bot)
    log.info("Starting Prezent AI Webhook server on port %s", port)
    log.info("Primary Gemini model: %s", GEMINI_MODEL)
    log.info("Fallback Gemini model: %s", GEMINI_FALLBACK_MODEL)
    log.info("Admin IDs configured: %s", sorted(ADMIN_IDS))
    log.info("Database backend: %s", "PostgreSQL" if DATABASE_URL else "SQLite")
    if RENDER_IS_HOSTED and not DATABASE_URL:
        log.warning("Render SQLite storage is ephemeral. Telegram Stars purchases are disabled until DATABASE_URL is configured.")
    web.run_app(app, host="0.0.0.0", port=port, shutdown_timeout=30)


if __name__ == "__main__":
    main()
