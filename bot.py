
import asyncio
import hashlib
import html
import io
import json
import logging
import os
import random
import re
import tempfile
import uuid
import zipfile

from pathlib import Path
from xml.sax.saxutils import escape

import httpx
from aiohttp import web

from aiogram import Bot, Dispatcher, F
from aiogram.enums import ParseMode
from aiogram.filters import Command, CommandStart
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.types import (
    Message,
    CallbackQuery,
    FSInputFile,
    InlineKeyboardButton,
    InlineKeyboardMarkup,
    ReplyKeyboardMarkup,
    KeyboardButton,
    BotCommand,
)
from aiogram.client.default import DefaultBotProperties
from aiogram.webhook.aiohttp_server import (
    SimpleRequestHandler,
    setup_application,
)

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.oxml.xmlchemy import OxmlElement
from pptx.util import Inches, Pt

from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import Frame, Paragraph, Spacer


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

# Економна основна модель.
GEMINI_MODEL = os.getenv(
    "GEMINI_MODEL",
    "gemini-3.5-flash-lite"
).strip()

# Резервна модель.
GEMINI_FALLBACK_MODEL = os.getenv(
    "GEMINI_FALLBACK_MODEL",
    "gemini-3.1-flash-lite"
).strip()

MAX_TOPIC_LENGTH = 500

if not BOT_TOKEN:
    raise RuntimeError(
        "BOT_TOKEN is missing in Render Environment."
    )

if not GEMINI_API_KEY:
    raise RuntimeError(
        "GEMINI_API_KEY is missing in Render Environment."
    )


bot = Bot(
    token=BOT_TOKEN,
    default=DefaultBotProperties(
        parse_mode=ParseMode.HTML
    )
)

dp = Dispatcher(
    storage=MemoryStorage()
)

# Поточні завдання генерації користувачів.
ACTIVE_GENERATIONS = {}


# ============================================================
# 2. PERMANENT TELEGRAM KEYBOARD
# ============================================================

BTN_START = "🚀 Старт"
BTN_RESTART = "🔄 Перезапустити"
BTN_CREATE = "🎨 Створити презентацію"
BTN_HOME = "🏠 Головне меню"
BTN_CANCEL = "❌ Скасувати"
BTN_FEATURES = "✨ Можливості"
BTN_HELP = "❓ Допомога"

MAIN_KEYBOARD = ReplyKeyboardMarkup(
    keyboard=[
        [
            KeyboardButton(text=BTN_START),
            KeyboardButton(text=BTN_RESTART),
        ],
        [
            KeyboardButton(text=BTN_CREATE),
        ],
        [
            KeyboardButton(text=BTN_HOME),
            KeyboardButton(text=BTN_CANCEL),
        ],
        [
            KeyboardButton(text=BTN_FEATURES),
            KeyboardButton(text=BTN_HELP),
        ],
    ],
    resize_keyboard=True,
    is_persistent=True,
    input_field_placeholder="Обери дію в меню..."
)


# ============================================================
# 3. DESIGNS
# ============================================================

DESIGNS = {
    "blue": {
        "name": "Синій",
        "bg": (13, 24, 48),
        "accent": (62, 133, 255),
        "card": (24, 42, 75),
    },

    "purple": {
        "name": "Фіолетовий",
        "bg": (31, 20, 55),
        "accent": (164, 108, 255),
        "card": (52, 37, 82),
    },

    "dark": {
        "name": "Темний",
        "bg": (22, 25, 32),
        "accent": (65, 194, 205),
        "card": (42, 47, 58),
    },

    "light": {
        "name": "Світлий",
        "bg": (242, 246, 252),
        "accent": (50, 105, 195),
        "card": (226, 234, 246),
    },
}

LANGUAGES = {
    "uk": "українською",
    "en": "англійською",
    "ru": "російською",
}


# ============================================================
# 4. FSM
# ============================================================

class PresentationForm(StatesGroup):
    topic = State()
    slides = State()
    language = State()
    design = State()


# ============================================================
# 5. KEYBOARD BUILDERS
# ============================================================

def slides_keyboard():

    return InlineKeyboardMarkup(
        inline_keyboard=[
            [
                InlineKeyboardButton(
                    text="6",
                    callback_data="slides:6"
                ),
                InlineKeyboardButton(
                    text="8",
                    callback_data="slides:8"
                ),
                InlineKeyboardButton(
                    text="10",
                    callback_data="slides:10"
                ),
                InlineKeyboardButton(
                    text="12",
                    callback_data="slides:12"
                ),
            ]
        ]
    )


def language_keyboard():

    return InlineKeyboardMarkup(
        inline_keyboard=[
            [
                InlineKeyboardButton(
                    text="🇺🇦 Українська",
                    callback_data="lang:uk"
                ),
                InlineKeyboardButton(
                    text="🇬🇧 English",
                    callback_data="lang:en"
                ),
            ],
            [
                InlineKeyboardButton(
                    text="🇷🇺 Русский",
                    callback_data="lang:ru"
                ),
            ]
        ]
    )


def design_keyboard():

    return InlineKeyboardMarkup(
        inline_keyboard=[
            [
                InlineKeyboardButton(
                    text="🔵 Синій",
                    callback_data="design:blue"
                ),
                InlineKeyboardButton(
                    text="🟣 Фіолетовий",
                    callback_data="design:purple"
                ),
            ],
            [
                InlineKeyboardButton(
                    text="⚫ Темний",
                    callback_data="design:dark"
                ),
                InlineKeyboardButton(
                    text="⚪ Світлий",
                    callback_data="design:light"
                ),
            ]
        ]
    )


# ============================================================
# 6. CANCEL ACTIVE GENERATION
# ============================================================

async def cancel_active_generation(user_id: int):

    task = ACTIVE_GENERATIONS.get(user_id)

    current = asyncio.current_task()

    if (
        task
        and not task.done()
        and task is not current
    ):

        log.info(
            "Cancellation requested for user %s",
            user_id
        )

        task.cancel()


# ============================================================
# 7. HOME AND MENU
# ============================================================

WELCOME_TEXT = (
    "🚀 <b>Prezent AI</b>\n\n"
    "Твій особистий AI-генератор презентацій!\n\n"
    "📊 PowerPoint із переходами між слайдами\n"
    "📄 PDF для перегляду та друку\n"
    "🎨 Чотири варіанти дизайну\n"
    "🧠 Google Gemini AI\n"
    "⚡ Анімований процес створення\n\n"
    "Обери потрібну дію в меню внизу 👇"
)

FEATURES_TEXT = (
    "✨ <b>Можливості Prezent AI</b>\n\n"
    "🧠 AI створює текст і структуру презентації\n"
    "📊 6, 8, 10 або 12 слайдів\n"
    "🌍 Українська, англійська та російська\n"
    "🎨 Синій, фіолетовий, темний та світлий дизайн\n"
    "📂 PowerPoint PPTX\n"
    "📄 PDF із підтримкою української мови\n"
    "🎞️ Переходи Fade між слайдами\n"
    "⚡ Анімований індикатор прогресу\n\n"
    "Складні анімації окремих об'єктів поки не реалізовані."
)

HELP_TEXT = (
    "❓ <b>Допомога Prezent AI</b>\n\n"
    "🚀 Старт — головне меню.\n"
    "🔄 Перезапустити — скинути поточний діалог.\n"
    "🎨 Створити презентацію — запустити генератор.\n"
    "🏠 Головне меню — повернутися на початок.\n"
    "❌ Скасувати — зупинити поточне створення.\n"
    "✨ Можливості — список функцій.\n\n"
    "Команди:\n"
    "/start — головне меню\n"
    "/restart — перезапустити діалог\n"
    "/cancel — скасувати"
)


async def show_home(
    message: Message,
    state: FSMContext
):

    await cancel_active_generation(
        message.from_user.id
    )

    await state.clear()

    await message.answer(
        WELCOME_TEXT,
        reply_markup=MAIN_KEYBOARD
    )


async def start_creation(
    message: Message,
    state: FSMContext
):

    await cancel_active_generation(
        message.from_user.id
    )

    await state.clear()

    await state.set_state(
        PresentationForm.topic
    )

    await message.answer(
        "📝 <b>Створення презентації</b>\n\n"
        "Напиши тему презентації.\n\n"
        "Наприклад: «Штучний інтелект у сучасному світі».\n"
        "Максимум 500 символів.\n\n"
        "Можеш повернутися в меню кнопкою «🏠 Головне меню».",
        reply_markup=MAIN_KEYBOARD
    )


# ============================================================
# 8. START COMMANDS
# ============================================================

@dp.message(CommandStart())
async def command_start(
    message: Message,
    state: FSMContext
):

    await show_home(
        message,
        state
    )


@dp.message(Command("restart"))
async def command_restart(
    message: Message,
    state: FSMContext
):

    await cancel_active_generation(
        message.from_user.id
    )

    await state.clear()

    await message.answer(
        "🔄 <b>Діалог перезапущено!</b>\n\n"
        "Можеш почати нову презентацію.",
        reply_markup=MAIN_KEYBOARD
    )


@dp.message(Command("cancel"))
async def command_cancel(
    message: Message,
    state: FSMContext
):

    await cancel_active_generation(
        message.from_user.id
    )

    await state.clear()

    await message.answer(
        "❌ Поточну дію скасовано.",
        reply_markup=MAIN_KEYBOARD
    )


# ============================================================
# 9. PERMANENT MENU BUTTON HANDLERS
# ============================================================

@dp.message(F.text == BTN_START)
async def button_start(
    message: Message,
    state: FSMContext
):

    await show_home(message, state)


@dp.message(F.text == BTN_RESTART)
async def button_restart(
    message: Message,
    state: FSMContext
):

    await cancel_active_generation(
        message.from_user.id
    )

    await state.clear()

    await message.answer(
        "🔄 <b>Діалог перезапущено!</b>\n\n"
        "Усі параметри попереднього налаштування скинуто.",
        reply_markup=MAIN_KEYBOARD
    )


@dp.message(F.text == BTN_CREATE)
async def button_create(
    message: Message,
    state: FSMContext
):

    await start_creation(
        message,
        state
    )


@dp.message(F.text == BTN_HOME)
async def button_home(
    message: Message,
    state: FSMContext
):

    await show_home(message, state)


@dp.message(F.text == BTN_CANCEL)
async def button_cancel(
    message: Message,
    state: FSMContext
):

    await cancel_active_generation(
        message.from_user.id
    )

    await state.clear()

    await message.answer(
        "❌ Створення скасовано.\n\n"
        "Можеш почати знову через меню.",
        reply_markup=MAIN_KEYBOARD
    )


@dp.message(F.text == BTN_FEATURES)
async def button_features(message: Message):

    await message.answer(
        FEATURES_TEXT,
        reply_markup=MAIN_KEYBOARD
    )


@dp.message(F.text == BTN_HELP)
async def button_help(message: Message):

    await message.answer(
        HELP_TEXT,
        reply_markup=MAIN_KEYBOARD
    )


# ============================================================
# 10. INLINE CREATE BUTTON SUPPORT
# ============================================================

@dp.callback_query(F.data == "create")
async def inline_create(
    callback: CallbackQuery,
    state: FSMContext
):

    await callback.answer()

    await start_creation(
        callback.message,
        state
    )


# ============================================================
# 11. ANIMATION
# НЕ ЗМІНЮВАЮ: ТА САМА ШВИДКІСТЬ І ТА САМА ШКАЛА
# ============================================================

class ProgressAnimator:

    FRAMES = (
        "⏳",
        "⌛",
        "🔄",
        "⚙️",
    )

    def __init__(self, message: Message):

        self.message = message

        self.stage = "Підготовка запиту"

        self.percent = 4
        self.target = 22

        self.frame = 0

        self.task = None
        self.lock = asyncio.Lock()

        self.running = False

    def set_stage(
        self,
        text: str,
        target: int
    ):

        self.stage = text

        self.target = max(
            self.percent,
            min(98, target)
        )

    def render(self):

        filled = max(
            0,
            min(10, self.percent // 10)
        )

        bar = (
            "🟦" * filled
            + "⬜" * (10 - filled)
        )

        icon = self.FRAMES[
            self.frame % len(self.FRAMES)
        ]

        return (
            f"{icon} <b>Prezent AI створює презентацію</b>\n\n"
            f"{self.stage}\n"
            f"{bar} <b>{self.percent}%</b>\n\n"
            "<i>Будь ласка, зачекай — це може зайняти кілька хвилин.</i>"
        )

    async def safe_edit(self):

        try:

            async with self.lock:

                await self.message.edit_text(
                    self.render()
                )

        except Exception as exc:

            if "message is not modified" not in str(exc).lower():

                log.debug(
                    "Progress edit skipped: %s",
                    exc
                )

    async def run(self):

        self.running = True

        while self.running:

            self.frame += 1

            if self.percent < self.target:

                self.percent = min(
                    self.target,
                    self.percent + 1
                )

            await self.safe_edit()

            await asyncio.sleep(1.6)

    def start(self):

        self.task = asyncio.create_task(
            self.run()
        )

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

            await self.message.edit_text(
                text
            )

        except Exception:

            log.exception(
                "Could not update progress message"
            )


# ============================================================
# 12. GEMINI ERRORS
# ============================================================

class GeminiAPIError(Exception):
    pass


class GeminiRetryableError(GeminiAPIError):
    pass


class GeminiModelUnavailable(GeminiAPIError):
    pass


class GeminiResponseError(GeminiAPIError):
    pass


# ============================================================
# 13. JSON PARSER
# ============================================================

def extract_json(text: str) -> dict:

    if not isinstance(text, str) or not text.strip():

        raise GeminiResponseError(
            "Gemini повернув порожню відповідь."
        )

    cleaned = text.strip()

    cleaned = re.sub(
        r"^```(?:json)?\s*",
        "",
        cleaned,
        flags=re.IGNORECASE
    )

    cleaned = re.sub(
        r"\s*```$",
        "",
        cleaned
    )

    start = cleaned.find("{")
    end = cleaned.rfind("}")

    if start < 0 or end <= start:

        raise GeminiResponseError(
            "Gemini повернув відповідь без JSON."
        )

    try:

        return json.loads(
            cleaned[start:end + 1]
        )

    except json.JSONDecodeError as exc:

        raise GeminiResponseError(
            "Gemini повернув неправильний JSON."
        ) from exc


# ============================================================
# 14. GEMINI REQUEST WITH RETRIES
# ============================================================

async def request_gemini(
    model: str,
    prompt: str
) -> dict:

    url = (
        "https://generativelanguage.googleapis.com/v1beta/"
        f"models/{model}:generateContent"
    )

    headers = {
        "x-goog-api-key": GEMINI_API_KEY,
        "Content-Type": "application/json",
    }

    payload = {
        "systemInstruction": {
            "parts": [
                {
                    "text": (
                        "Ти професійний автор презентацій. "
                        "Створюй точний, грамотний, логічний матеріал. "
                        "Повертай тільки правильний JSON."
                    )
                }
            ]
        },

        "contents": [
            {
                "role": "user",
                "parts": [
                    {"text": prompt}
                ]
            }
        ],

        "generationConfig": {
            "maxOutputTokens": 8192,
            "responseMimeType": "application/json",
            "thinkingConfig": {
                "thinkingLevel": "low"
            }
        }
    }

    for attempt in range(1, 4):

        try:

            async with httpx.AsyncClient(
                timeout=httpx.Timeout(
                    150.0,
                    connect=25.0
                )
            ) as client:

                response = await client.post(
                    url,
                    headers=headers,
                    json=payload
                )

        except httpx.TimeoutException as exc:

            log.warning(
                "Gemini timeout: model=%s attempt=%s/3",
                model,
                attempt
            )

            if attempt < 3:

                await asyncio.sleep(
                    2 ** attempt
                    + random.uniform(0, 0.8)
                )

                continue

            raise GeminiRetryableError(
                "Gemini не відповів вчасно після трьох спроб."
            ) from exc

        except httpx.HTTPError as exc:

            log.exception(
                "Gemini connection error: model=%s",
                model
            )

            if attempt < 3:

                await asyncio.sleep(
                    2 ** attempt
                )

                continue

            raise GeminiRetryableError(
                "Помилка з'єднання з Google Gemini."
            ) from exc

        try:

            result = response.json()

        except ValueError as exc:

            raise GeminiAPIError(
                f"Gemini повернув неправильну відповідь. "
                f"HTTP {response.status_code}."
            ) from exc

        status = response.status_code

        # Temporary overload or quota issue.
        if status in (408, 429, 500, 502, 503, 504):

            log.warning(
                "Gemini temporary HTTP %s: model=%s attempt=%s/3",
                status,
                model,
                attempt
            )

            if attempt < 3:

                retry_after = response.headers.get(
                    "Retry-After"
                )

                try:

                    delay = float(retry_after)

                except (TypeError, ValueError):

                    delay = 2 ** attempt

                delay = min(delay, 12)

                await asyncio.sleep(
                    delay + random.uniform(0, 0.8)
                )

                continue

            raise GeminiRetryableError(
                f"Gemini тимчасово недоступний: HTTP {status}."
            )

        if status == 404:

            raise GeminiModelUnavailable(
                f"Модель {model} не знайдена або недоступна."
            )

        if status in (401, 403):

            detail = result.get(
                "error",
                {}
            ).get(
                "message",
                "Перевір API-ключ."
            )

            raise GeminiAPIError(
                f"Google відхилив доступ: {str(detail)[:250]}"
            )

        if status >= 400:

            detail = result.get(
                "error",
                {}
            ).get(
                "message",
                "Невідома помилка."
            )

            raise GeminiAPIError(
                f"Gemini HTTP {status}: {str(detail)[:250]}"
            )

        if result.get("error"):

            detail = result["error"].get(
                "message",
                "Невідома помилка."
            )

            raise GeminiAPIError(
                str(detail)[:300]
            )

        candidates = result.get(
            "candidates",
            []
        )

        if not candidates:

            feedback = result.get(
                "promptFeedback",
                {}
            )

            reason = feedback.get(
                "blockReason",
                "невідома причина"
            )

            raise GeminiResponseError(
                f"Gemini не повернув кандидата відповіді: {reason}."
            )

        candidate = candidates[0]

        finish_reason = candidate.get(
            "finishReason",
            "UNKNOWN"
        )

        parts = (
            candidate.get(
                "content",
                {}
            ).get(
                "parts",
                []
            )
        )

        text_parts = []

        for part in parts:

            if not isinstance(part, dict):
                continue

            # Не додаємо внутрішні thought-частини.
            if part.get("thought"):
                continue

            part_text = part.get("text")

            if (
                isinstance(part_text, str)
                and part_text.strip()
            ):

                text_parts.append(
                    part_text.strip()
                )

        content = "\n".join(
            text_parts
        ).strip()

        usage = result.get(
            "usageMetadata",
            {}
        )

        log.info(
            "Gemini model=%s finish=%s input=%s output=%s chars=%s",
            model,
            finish_reason,
            usage.get("promptTokenCount"),
            usage.get("candidatesTokenCount"),
            len(content)
        )

        if finish_reason == "SAFETY":

            raise GeminiAPIError(
                "Gemini заблокував відповідь фільтром безпеки."
            )

        if finish_reason == "MAX_TOKENS":

            raise GeminiResponseError(
                "Gemini досяг ліміту вихідних токенів."
            )

        if not content:

            if attempt < 3:

                await asyncio.sleep(
                    2 ** attempt
                )

                continue

            raise GeminiRetryableError(
                "Gemini тричі повернув порожню відповідь."
            )

        return extract_json(content)

    raise GeminiRetryableError(
        f"Модель {model} не повернула результат."
    )


# ============================================================
# 15. GENERATE PRESENTATION CONTENT
# ============================================================

async def generate_content(data: dict) -> dict:

    topic = data["topic"]
    count = data["slides"]

    language = LANGUAGES.get(
        data["language"],
        "українською"
    )

    prompt = f"""
Створи професійну та детальну презентацію.

ТЕМА: {topic}
КІЛЬКІСТЬ СЛАЙДІВ: рівно {count}
МОВА: {language}

Поверни тільки валідний JSON без Markdown:

{{
  "title": "Назва презентації",
  "slides": [
    {{
      "title": "Заголовок слайда",
      "content": [
        "Теза 1",
        "Теза 2",
        "Теза 3",
        "Теза 4"
      ]
    }}
  ]
}}

Вимоги:
1. Рівно {count} слайдів.
2. Логічна структура: вступ, основна частина, висновки.
3. 3-5 змістовних тез на кожному слайді.
4. Не повторюй інформацію.
5. Не вигадуй статистику.
6. Використовуй грамотну мову.
7. Тези мають бути короткими та зрозумілими.
8. Враховуй, що презентацію будуть показувати на екрані.
9. Не додавай коментарів поза JSON.
""".strip()

    models = []

    for model in (
        GEMINI_MODEL,
        GEMINI_FALLBACK_MODEL
    ):

        if model and model not in models:

            models.append(model)

    last_error = None

    for model in models:

        try:

            log.info(
                "Trying Gemini model: %s",
                model
            )

            parsed = await request_gemini(
                model,
                prompt
            )

            slides = parsed.get("slides")

            if not isinstance(slides, list):

                raise GeminiResponseError(
                    "Gemini не повернув список слайдів."
                )

            if len(slides) != count:

                raise GeminiResponseError(
                    f"Gemini створив {len(slides)} "
                    f"замість {count} слайдів."
                )

            normalized = []

            for number, slide in enumerate(
                slides,
                start=1
            ):

                if not isinstance(slide, dict):

                    raise GeminiResponseError(
                        f"Слайд {number} має неправильний формат."
                    )

                title = str(
                    slide.get(
                        "title",
                        f"Слайд {number}"
                    )
                ).strip()[:180]

                points = slide.get(
                    "content",
                    []
                )

                if isinstance(points, str):

                    points = [points]

                if not isinstance(points, list):

                    points = []

                points = [
                    str(point).strip()[:350]
                    for point in points
                    if str(point).strip()
                ]

                if not points:

                    points = [
                        "Основна інформація за темою презентації."
                    ]

                normalized.append({
                    "title": title or f"Слайд {number}",
                    "content": points[:6]
                })

            if len(normalized) != count:

                raise GeminiResponseError(
                    "Не вдалося обробити всі слайди."
                )

            log.info(
                "Presentation created successfully using %s",
                model
            )

            return {
                "title": str(
                    parsed.get(
                        "title",
                        topic
                    )
                ).strip()[:180],

                "slides": normalized,

                "model": model
            }

        except (
            GeminiRetryableError,
            GeminiModelUnavailable,
            GeminiResponseError
        ) as exc:

            last_error = exc

            log.warning(
                "Model %s failed: %s",
                model,
                exc
            )

            if model != models[-1]:

                log.warning(
                    "Switching to fallback model: %s",
                    models[models.index(model) + 1]
                )

                continue

        except GeminiAPIError:

            raise

    raise GeminiAPIError(
        "Не вдалося отримати презентацію від основної "
        "або резервної моделі. "
        f"Остання помилка: {last_error}"
    )


# ============================================================
# 16. POWERPOINT FADE TRANSITION
# ============================================================

def add_fade_transition(slide):

    transition = OxmlElement(
        "p:transition"
    )

    transition.set(
        "spd",
        "med"
    )

    transition.append(
        OxmlElement("p:fade")
    )

    slide._element.insert(
        2,
        transition
    )


# ============================================================
# 17. POWERPOINT TEXT HELPER
# ============================================================

def add_text(
    slide,
    text,
    x,
    y,
    w,
    h,
    size,
    color,
    bold=False,
    align=PP_ALIGN.LEFT
):

    box = slide.shapes.add_textbox(
        Inches(x),
        Inches(y),
        Inches(w),
        Inches(h)
    )

    frame = box.text_frame

    frame.clear()
    frame.word_wrap = True

    frame.margin_left = Inches(0.03)
    frame.margin_right = Inches(0.03)
    frame.margin_top = Inches(0.02)
    frame.margin_bottom = Inches(0.02)

    paragraph = frame.paragraphs[0]

    paragraph.text = str(text)
    paragraph.alignment = align

    paragraph.font.name = "Aptos"
    paragraph.font.size = Pt(size)
    paragraph.font.bold = bold

    paragraph.font.color.rgb = RGBColor(
        *color
    )

    return box


# ============================================================
# 18. CREATE POWERPOINT
# ============================================================

def create_pptx(
    data: dict,
    design: str,
    output_path: str
):

    theme = DESIGNS[design]

    bg = theme["bg"]
    accent = theme["accent"]
    card = theme["card"]

    is_light = design == "light"

    title_color = (
        (24, 35, 55)
        if is_light
        else (255, 255, 255)
    )

    body_color = (
        (52, 64, 82)
        if is_light
        else (228, 235, 248)
    )

    prs = Presentation()

    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    for index, item in enumerate(
        data["slides"]
    ):

        slide = prs.slides.add_slide(
            prs.slide_layouts[6]
        )

        slide.background.fill.solid()

        slide.background.fill.fore_color.rgb = (
            RGBColor(*bg)
        )

        rail = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            0,
            0,
            Inches(0.16),
            Inches(7.5)
        )

        rail.fill.solid()

        rail.fill.fore_color.rgb = RGBColor(
            *accent
        )

        rail.line.fill.background()

        add_text(
            slide,
            "PREZENT AI  /  PRESENTATION",
            0.75,
            0.38,
            6.0,
            0.3,
            10,
            accent,
            True
        )

        add_text(
            slide,
            f"{index + 1:02d} / {len(data['slides']):02d}",
            11.35,
            0.38,
            1.1,
            0.3,
            10,
            accent,
            True,
            PP_ALIGN.RIGHT
        )

        add_text(
            slide,
            item["title"],
            0.78,
            0.92,
            11.7,
            1.15,
            30,
            title_color,
            True
        )

        card_shape = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(0.78),
            Inches(2.25),
            Inches(11.75),
            Inches(4.45)
        )

        card_shape.fill.solid()

        card_shape.fill.fore_color.rgb = RGBColor(
            *card
        )

        card_shape.line.color.rgb = RGBColor(
            *accent
        )

        text_box = slide.shapes.add_textbox(
            Inches(1.18),
            Inches(2.58),
            Inches(10.95),
            Inches(3.82)
        )

        tf = text_box.text_frame

        tf.clear()
        tf.word_wrap = True

        tf.margin_left = Inches(0.03)
        tf.margin_right = Inches(0.03)
        tf.margin_top = Inches(0.03)
        tf.margin_bottom = Inches(0.02)

        points = item.get(
            "content",
            []
        )[:6]

        for p_index, point in enumerate(points):

            paragraph = (
                tf.paragraphs[0]
                if p_index == 0
                else tf.add_paragraph()
            )

            paragraph.text = "•  " + str(point)

            paragraph.font.name = "Aptos"

            paragraph.font.size = Pt(
                20 if len(points) <= 4 else 17
            )

            paragraph.font.color.rgb = RGBColor(
                *body_color
            )

            paragraph.space_after = Pt(15)

            paragraph.level = 0

        add_fade_transition(slide)

    prs.save(output_path)


# ============================================================
# 19. GET UKRAINIAN FONT FOR PDF
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

    target = (
        Path(tempfile.gettempdir())
        / "PrezentAI-DejaVuSans.ttf"
    )

    if (
        target.exists()
        and target.stat().st_size > 100_000
    ):

        return str(target)

    archive_url = (
        "https://github.com/dejavu-fonts/dejavu-fonts/"
        "releases/download/version_2_37/"
        "dejavu-fonts-ttf-2.37.zip"
    )

    try:

        async with httpx.AsyncClient(
            timeout=60.0,
            follow_redirects=True
        ) as client:

            response = await client.get(
                archive_url
            )

            response.raise_for_status()

        with zipfile.ZipFile(
            io.BytesIO(response.content)
        ) as archive:

            font_name = next(
                (
                    name
                    for name in archive.namelist()
                    if name.endswith("/ttf/DejaVuSans.ttf")
                ),
                None
            )

            if not font_name:

                raise RuntimeError(
                    "DejaVuSans.ttf not found in downloaded archive."
                )

            font_bytes = archive.read(
                font_name
            )

        if not font_bytes.startswith(
            b"\x00\x01\x00\x00"
        ):

            raise RuntimeError(
                "Downloaded file is not a valid TTF font."
            )

        target.write_bytes(
            font_bytes
        )

        return str(target)

    except Exception as exc:

        log.exception(
            "Unable to download DejaVu font"
        )

        raise RuntimeError(
            "Не вдалося завантажити шрифт для українського PDF."
        ) from exc


# ============================================================
# 20. CREATE PDF
# ============================================================

def create_pdf(
    data: dict,
    design: str,
    output_path: str,
    font_path: str
):

    theme = DESIGNS[design]

    bg = colors.HexColor(
        "#%02x%02x%02x" % theme["bg"]
    )

    accent = colors.HexColor(
        "#%02x%02x%02x" % theme["accent"]
    )

    card = colors.HexColor(
        "#%02x%02x%02x" % theme["card"]
    )

    is_light = design == "light"

    title_color = (
        colors.HexColor("#182337")
        if is_light
        else colors.white
    )

    body_color = (
        colors.HexColor("#344052")
        if is_light
        else colors.HexColor("#e4ebf8")
    )

    if "PrezentSans" not in pdfmetrics.getRegisteredFontNames():

        pdfmetrics.registerFont(
            TTFont(
                "PrezentSans",
                font_path
            )
        )

    page_size = landscape(A4)

    page_w, page_h = page_size

    pdf = canvas.Canvas(
        output_path,
        pagesize=page_size
    )

    title_style = ParagraphStyle(
        "SlideTitle",
        fontName="PrezentSans",
        fontSize=23,
        leading=29,
        textColor=title_color,
        alignment=TA_LEFT,
        spaceAfter=8
    )

    body_style = ParagraphStyle(
        "SlideBody",
        fontName="PrezentSans",
        fontSize=15,
        leading=22,
        textColor=body_color,
        alignment=TA_LEFT,
        spaceAfter=13
    )

    for index, item in enumerate(
        data["slides"]
    ):

        pdf.setFillColor(bg)

        pdf.rect(
            0,
            0,
            page_w,
            page_h,
            fill=1,
            stroke=0
        )

        pdf.setFillColor(accent)

        pdf.rect(
            0,
            0,
            12,
            page_h,
            fill=1,
            stroke=0
        )

        pdf.setFillColor(accent)

        pdf.setFont(
            "PrezentSans",
            8
        )

        pdf.drawString(
            45,
            page_h - 30,
            "PREZENT AI  /  PRESENTATION"
        )

        pdf.drawRightString(
            page_w - 38,
            page_h - 30,
            f"{index + 1:02d} / {len(data['slides']):02d}"
        )

        pdf.setFillColor(card)

        pdf.roundRect(
            42,
            55,
            page_w - 84,
            page_h - 112,
            12,
            fill=1,
            stroke=0
        )

        story = [
            Paragraph(
                escape(item["title"]),
                title_style
            ),
            Spacer(1, 14)
        ]

        for point in item.get(
            "content",
            []
        )[:6]:

            story.append(
                Paragraph(
                    "•  " + escape(str(point)),
                    body_style
                )
            )

        frame = Frame(
            65,
            78,
            page_w - 130,
            page_h - 180,
            leftPadding=0,
            rightPadding=0,
            topPadding=0,
            bottomPadding=0,
            showBoundary=0
        )

        frame.addFromList(
            story,
            pdf
        )

        if story:

            log.warning(
                "Some text did not fit in PDF slide %s",
                index + 1
            )

        pdf.showPage()

    pdf.save()


# ============================================================
# 21. GENERATION HANDLER
# ============================================================

@dp.callback_query(F.data.startswith("design:"))
async def design_handler(
    callback: CallbackQuery,
    state: FSMContext
):

    user_id = callback.from_user.id

    if await state.get_state() != PresentationForm.design.state:

        await callback.answer(
            "Спочатку пройди налаштування презентації.",
            show_alert=True
        )

        return

    if user_id in ACTIVE_GENERATIONS:

        await callback.answer(
            "Твоя презентація вже створюється.",
            show_alert=True
        )

        return

    design = callback.data.split(":", 1)[1]

    if design not in DESIGNS:

        await callback.answer(
            "Непідтримуваний дизайн.",
            show_alert=True
        )

        return

    data = await state.get_data()

    if not all(
        key in data
        for key in ("topic", "slides", "language")
    ):

        await state.clear()

        await callback.answer(
            "Налаштування втрачено. Почни спочатку.",
            show_alert=True
        )

        return

    data["design"] = design

    await state.clear()

    await callback.answer(
        "Починаю створення!"
    )

    status = await callback.message.answer(
        "⏳ <b>Prezent AI запускається...</b>\n\n"
        "🟦⬜⬜⬜⬜⬜⬜⬜⬜⬜ 4%",
        reply_markup=MAIN_KEYBOARD
    )

    progress = ProgressAnimator(status)

    current_task = asyncio.current_task()

    ACTIVE_GENERATIONS[user_id] = current_task

    progress.start()

    try:

        # STAGE 1: AI TEXT

        progress.set_stage(
            "🧠 AI складає план і готує текст слайдів",
            62
        )

        presentation_data = await generate_content(data)

        # STAGE 2: POWERPOINT

        progress.set_stage(
            "🎨 Оформлюю слайди та додаю переходи",
            82
        )

        with tempfile.TemporaryDirectory(
            prefix="prezent_ai_"
        ) as tmp:

            pptx_path = str(
                Path(tmp) / "PrezentAI.pptx"
            )

            pdf_path = str(
                Path(tmp) / "PrezentAI.pdf"
            )

            await asyncio.to_thread(
                create_pptx,
                presentation_data,
                design,
                pptx_path
            )

            if not Path(pptx_path).exists():

                raise RuntimeError(
                    "PowerPoint file was not created."
                )

            # STAGE 3: PDF

            progress.set_stage(
                "📄 Готую PDF-файл",
                94
            )

            pdf_ready = False
            pdf_error = None

            try:

                font_path = await ensure_unicode_font()

                await asyncio.to_thread(
                    create_pdf,
                    presentation_data,
                    design,
                    pdf_path,
                    font_path
                )

                pdf_ready = Path(pdf_path).exists()

            except Exception as exc:

                pdf_error = str(exc)

                log.exception(
                    "PDF generation failed"
                )

            # SUCCESS OR PARTIAL SUCCESS

            if pdf_ready:

                await progress.finish(
                    "✅ <b>Презентація готова!</b>\n\n"
                    "📤 Надсилаю PowerPoint та PDF..."
                )

            else:

                await progress.finish(
                    "✅ <b>PowerPoint створено!</b>\n\n"
                    "⚠️ PDF не вдалося підготувати. "
                    "Надсилаю PPTX."
                )

                log.error(
                    "PDF unavailable: %s",
                    pdf_error
                )

            # SEND POWERPOINT FIRST

            await callback.message.answer_document(
                FSInputFile(
                    pptx_path,
                    filename="PrezentAI.pptx"
                ),
                caption=(
                    "📊 <b>PowerPoint</b>\n"
                    "Презентація з переходами Fade."
                ),
                reply_markup=MAIN_KEYBOARD
            )

            # SEND PDF IF READY

            if pdf_ready:

                await callback.message.answer_document(
                    FSInputFile(
                        pdf_path,
                        filename="PrezentAI.pdf"
                    ),
                    caption=(
                        "📄 <b>PDF</b>\n"
                        "Версія для перегляду та друку."
                    ),
                    reply_markup=MAIN_KEYBOARD
                )

            else:

                await callback.message.answer(
                    "PPTX готовий. PDF цього разу не створився. "
                    "Можеш завантажити PPTX і відкрити його в PowerPoint.",
                    reply_markup=MAIN_KEYBOARD
                )

        await callback.message.answer(
            "✅ <b>Готово!</b>\n\n"
            "Можеш створити наступну презентацію кнопкою нижче.",
            reply_markup=MAIN_KEYBOARD
        )

    except asyncio.CancelledError:

        log.info(
            "Generation cancelled by user %s",
            user_id
        )

        await progress.finish(
            "❌ <b>Створення зупинено.</b>"
        )

        # Do not propagate cancellation further.
        return

    except GeminiAPIError as exc:

        log.warning(
            "Gemini generation failed: %s",
            exc
        )

        await progress.finish(
            "❌ <b>AI не зміг створити текст презентації.</b>\n\n"
            f"Причина: {html.escape(str(exc))}\n\n"
            "Перевір доступність моделей і квоти Google AI Studio."
        )

        await callback.message.answer(
            "Можеш повторити спробу через меню.",
            reply_markup=MAIN_KEYBOARD
        )

    except Exception:

        error_id = uuid.uuid4().hex[:8].upper()

        log.exception(
            "Presentation generation failed; error_id=%s",
            error_id
        )

        await progress.finish(
            "❌ <b>Не вдалося створити презентацію.</b>\n\n"
            f"Код помилки: <code>{error_id}</code>\n"
            "Перевір Render Logs."
        )

        await callback.message.answer(
            "Повернутися в меню можна кнопкою нижче.",
            reply_markup=MAIN_KEYBOARD
        )

    finally:

        ACTIVE_GENERATIONS.pop(
            user_id,
            None
        )


# ============================================================
# 22. SELECT TOPIC
# ============================================================

@dp.message(PresentationForm.topic)
async def topic_handler(
    message: Message,
    state: FSMContext
):

    topic = (message.text or "").strip()

    if len(topic) < 3:

        await message.answer(
            "Тема занадто коротка. Напиши хоча б 3 символи.",
            reply_markup=MAIN_KEYBOARD
        )

        return

    if len(topic) > MAX_TOPIC_LENGTH:

        await message.answer(
            "Тема занадто довга. Максимум 500 символів.",
            reply_markup=MAIN_KEYBOARD
        )

        return

    await state.update_data(
        topic=topic
    )

    await state.set_state(
        PresentationForm.slides
    )

    await message.answer(
        "📊 <b>Скільки слайдів потрібно?</b>\n\n"
        "Обери кількість:",
        reply_markup=slides_keyboard()
    )


# ============================================================
# 23. SELECT SLIDES
# ============================================================

@dp.callback_query(F.data.startswith("slides:"))
async def slides_handler(
    callback: CallbackQuery,
    state: FSMContext
):

    if await state.get_state() != PresentationForm.slides.state:

        await callback.answer(
            "Спочатку введи тему презентації.",
            show_alert=True
        )

        return

    try:

        count = int(
            callback.data.split(":", 1)[1]
        )

    except (ValueError, IndexError):

        await callback.answer(
            "Неправильна кількість слайдів.",
            show_alert=True
        )

        return

    if count not in (6, 8, 10, 12):

        await callback.answer(
            "Обери 6, 8, 10 або 12.",
            show_alert=True
        )

        return

    await state.update_data(
        slides=count
    )

    await state.set_state(
        PresentationForm.language
    )

    await callback.message.answer(
        "🌍 <b>Обери мову презентації:</b>",
        reply_markup=language_keyboard()
    )

    await callback.answer()


# ============================================================
# 24. SELECT LANGUAGE
# ============================================================

@dp.callback_query(F.data.startswith("lang:"))
async def language_handler(
    callback: CallbackQuery,
    state: FSMContext
):

    if await state.get_state() != PresentationForm.language.state:

        await callback.answer(
            "Спочатку обери кількість слайдів.",
            show_alert=True
        )

        return

    language = callback.data.split(":", 1)[1]

    if language not in LANGUAGES:

        await callback.answer(
            "Непідтримувана мова.",
            show_alert=True
        )

        return

    await state.update_data(
        language=language
    )

    await state.set_state(
        PresentationForm.design
    )

    await callback.message.answer(
        "🎨 <b>Обери дизайн презентації:</b>",
        reply_markup=design_keyboard()
    )

    await callback.answer()


# ============================================================
# 25. UNKNOWN TEXT
# ============================================================

@dp.message(F.text)
async def unknown_text(message: Message):

    await message.answer(
        "Не зрозумів команду. Скористайся кнопками в меню 👇",
        reply_markup=MAIN_KEYBOARD
    )


# ============================================================
# 26. RENDER HEALTH CHECK
# ============================================================

async def health_handler(request: web.Request):

    return web.json_response({
        "status": "ok",
        "service": "Prezent AI",
        "mode": "webhook"
    })


# ============================================================
# 27. WEBHOOK CONFIGURATION
# ============================================================

WEBHOOK_PATH = "/telegram/webhook"

RENDER_EXTERNAL_URL = os.getenv(
    "RENDER_EXTERNAL_URL",
    ""
).rstrip("/")

WEBHOOK_SECRET = os.getenv(
    "WEBHOOK_SECRET",
    ""
).strip()

if not WEBHOOK_SECRET:

    WEBHOOK_SECRET = hashlib.sha256(
        BOT_TOKEN.encode("utf-8")
    ).hexdigest()


async def on_startup(bot: Bot):

    if not RENDER_EXTERNAL_URL.startswith("https://"):

        raise RuntimeError(
            "RENDER_EXTERNAL_URL is missing. "
            "Deploy this application as Render Web Service."
        )

    webhook_url = (
        RENDER_EXTERNAL_URL
        + WEBHOOK_PATH
    )

    await bot.set_webhook(
        url=webhook_url,
        secret_token=WEBHOOK_SECRET,
        allowed_updates=dp.resolve_used_update_types(),
        drop_pending_updates=False,
        max_connections=10
    )

    info = await bot.get_webhook_info()

    log.info(
        "Telegram webhook registered: %s",
        info.url
    )

    log.info(
        "Webhook pending updates: %s",
        info.pending_update_count
    )

    await bot.set_my_commands([
        BotCommand(
            command="start",
            description="Головне меню"
        ),
        BotCommand(
            command="restart",
            description="Перезапустити діалог"
        ),
        BotCommand(
            command="cancel",
            description="Скасувати створення"
        ),
    ])

    log.info(
        "Telegram commands registered"
    )


# ============================================================
# 28. MAIN WEBHOOK SERVER
# ============================================================

def main():

    port = int(
        os.getenv("PORT", "10000")
    )

    app = web.Application(
        client_max_size=1024 * 1024
    )

    app.router.add_get(
        "/",
        health_handler
    )

    app.router.add_get(
        "/healthz",
        health_handler
    )

    webhook_handler = SimpleRequestHandler(
        dispatcher=dp,
        bot=bot,
        handle_in_background=True,
        secret_token=WEBHOOK_SECRET
    )

    webhook_handler.register(
        app,
        path=WEBHOOK_PATH
    )

    dp.startup.register(
        on_startup
    )

    setup_application(
        app,
        dp,
        bot=bot
    )

    log.info(
        "Starting Prezent AI Webhook server"
    )

    log.info(
        "Listening on port %s",
        port
    )

    log.info(
        "Primary Gemini model: %s",
        GEMINI_MODEL
    )

    log.info(
        "Fallback Gemini model: %s",
        GEMINI_FALLBACK_MODEL
    )

    # Тут немає start_polling().
    # Telegram використовує тільки Webhook.

    web.run_app(
        app,
        host="0.0.0.0",
        port=port,
        shutdown_timeout=30
    )


if __name__ == "__main__":

    main()
