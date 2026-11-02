
import asyncio
import html
import json
import logging
import os
import re
import tempfile
import uuid
from pathlib import Path
from xml.sax.saxutils import escape

import httpx

from aiogram import Bot, Dispatcher, F
from aiogram.enums import ParseMode
from aiogram.filters import Command, CommandStart
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.types import (
    CallbackQuery,
    FSInputFile,
    InlineKeyboardButton,
    InlineKeyboardMarkup,
    Message,
)
from aiogram.client.default import DefaultBotProperties

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

GEMINI_MODEL = os.getenv(
    "GEMINI_MODEL",
    "gemini-3.8-flash"
).strip()

GEMINI_URL = (
    "https://generativelanguage.googleapis.com/v1beta/"
    f"models/{GEMINI_MODEL}:generateContent"
)

MAX_TOPIC_LENGTH = 500
MAX_SLIDES = 12

if not BOT_TOKEN:
    raise RuntimeError(
        "BOT_TOKEN is missing. Add it in Render Environment."
    )

if not GEMINI_API_KEY:
    raise RuntimeError(
        "GEMINI_API_KEY is missing. Add it in Render Environment."
    )


bot = Bot(
    token=BOT_TOKEN,
    default=DefaultBotProperties(
        parse_mode=ParseMode.HTML
    )
)

dp = Dispatcher(storage=MemoryStorage())


# ============================================================
# 2. DESIGN THEMES
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
# 3. MAIN MENU
# ============================================================

MAIN_MENU = InlineKeyboardMarkup(
    inline_keyboard=[
        [
            InlineKeyboardButton(
                text="🎨 Створити презентацію",
                callback_data="create"
            )
        ],
        [
            InlineKeyboardButton(
                text="✨ Можливості",
                callback_data="features"
            )
        ],
        [
            InlineKeyboardButton(
                text="❓ Допомога",
                callback_data="help"
            )
        ],
    ]
)


class PresentationForm(StatesGroup):
    topic = State()
    slides = State()
    language = State()
    design = State()


# ============================================================
# 4. PROGRESS ANIMATION
# НЕ ЗМІНЮВАВ ПОПЕРЕДНЮ ЛОГІКУ АНІМАЦІЇ
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

    def set_stage(self, text: str, target: int):

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

            await self.message.edit_text(text)

        except Exception:

            log.exception(
                "Could not update final progress message"
            )


# ============================================================
# 5. START COMMAND
# ============================================================

@dp.message(CommandStart())
async def start_handler(
    message: Message,
    state: FSMContext
):

    await state.clear()

    await message.answer(
        "🚀 <b>Prezent AI</b>\n\n"

        "Твій особистий AI-генератор презентацій!\n\n"

        "📊 PowerPoint із переходами між слайдами\n"
        "📄 PDF для перегляду та друку\n"
        "🎨 Кілька професійних дизайнів\n"
        "🧠 Google Gemini 3.8 Flash\n"
        "⚡ Анімований процес створення\n\n"

        "Натисни кнопку нижче, щоб почати.",

        reply_markup=MAIN_MENU
    )


@dp.message(Command("cancel"))
async def cancel_handler(
    message: Message,
    state: FSMContext
):

    await state.clear()

    await message.answer(
        "Створення скасовано. Можеш почати знову 👇",
        reply_markup=MAIN_MENU
    )


# ============================================================
# 6. FEATURES
# ============================================================

@dp.callback_query(F.data == "features")
async def features_handler(
    callback: CallbackQuery
):

    await callback.answer()

    await callback.message.answer(
        "✨ <b>Можливості Prezent AI</b>\n\n"

        "• Створення матеріалу через Google Gemini\n"
        "• 6, 8, 10 або 12 слайдів\n"
        "• Українська, англійська або російська\n"
        "• Синій, фіолетовий, темний або світлий дизайн\n"
        "• PowerPoint із переходами Fade\n"
        "• PDF із підтримкою українських символів\n"
        "• Анімований процес генерації\n\n"

        "Примітка: складні анімації окремих об'єктів "
        "у PowerPoint поки не реалізовані."
    )


@dp.callback_query(F.data == "help")
async def help_handler(
    callback: CallbackQuery
):

    await callback.answer()

    await callback.message.answer(
        "📝 <b>Як створити презентацію</b>\n\n"

        "1. Натисни «Створити презентацію».\n"
        "2. Напиши тему.\n"
        "3. Вибери кількість слайдів.\n"
        "4. Вибери мову.\n"
        "5. Вибери дизайн.\n"
        "6. Дочекайся готових PPTX і PDF.\n\n"

        "Для скасування використовуй /cancel."
    )


# ============================================================
# 7. CREATE PRESENTATION
# ============================================================

@dp.callback_query(F.data == "create")
async def create_handler(
    callback: CallbackQuery,
    state: FSMContext
):

    await callback.answer()

    await state.clear()

    await state.set_state(
        PresentationForm.topic
    )

    await callback.message.answer(
        "📝 <b>Яка тема презентації?</b>\n\n"

        "Наприклад: «Штучний інтелект у сучасному світі».\n"
        "Максимум 500 символів."
    )


@dp.message(PresentationForm.topic)
async def topic_handler(
    message: Message,
    state: FSMContext
):

    topic = (message.text or "").strip()

    if len(topic) < 3:

        await message.answer(
            "Тема занадто коротка. Напиши хоча б 3 символи."
        )

        return

    if len(topic) > MAX_TOPIC_LENGTH:

        await message.answer(
            "Тема занадто довга. Максимум 500 символів."
        )

        return

    await state.update_data(
        topic=topic
    )

    await state.set_state(
        PresentationForm.slides
    )

    keyboard = InlineKeyboardMarkup(
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

    await message.answer(
        "📊 <b>Скільки слайдів потрібно?</b>",
        reply_markup=keyboard
    )


# ============================================================
# 8. SELECT SLIDES
# ============================================================

@dp.callback_query(F.data.startswith("slides:"))
async def slides_handler(
    callback: CallbackQuery,
    state: FSMContext
):

    if await state.get_state() != PresentationForm.slides.state:

        await callback.answer(
            "Почни створення через меню.",
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
            "Обери варіант із кнопок.",
            show_alert=True
        )

        return

    await state.update_data(
        slides=count
    )

    await state.set_state(
        PresentationForm.language
    )

    keyboard = InlineKeyboardMarkup(
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

    await callback.message.answer(
        "🌍 <b>Обери мову презентації:</b>",
        reply_markup=keyboard
    )

    await callback.answer()


# ============================================================
# 9. SELECT LANGUAGE
# ============================================================

@dp.callback_query(F.data.startswith("lang:"))
async def language_handler(
    callback: CallbackQuery,
    state: FSMContext
):

    if await state.get_state() != PresentationForm.language.state:

        await callback.answer(
            "Почни створення через меню.",
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

    keyboard = InlineKeyboardMarkup(
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

    await callback.message.answer(
        "🎨 <b>Обери дизайн:</b>",
        reply_markup=keyboard
    )

    await callback.answer()


# ============================================================
# 10. SELECT DESIGN AND GENERATE
# ============================================================

@dp.callback_query(F.data.startswith("design:"))
async def design_handler(
    callback: CallbackQuery,
    state: FSMContext
):

    if await state.get_state() != PresentationForm.design.state:

        await callback.answer(
            "Почни створення через меню.",
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

    data["design"] = design

    await state.clear()

    await callback.answer(
        "Починаю створення!"
    )

    status = await callback.message.answer(
        "⏳ <b>Prezent AI запускається...</b>\n\n"
        "🟦⬜⬜⬜⬜⬜⬜⬜⬜⬜ 4%"
    )

    progress = ProgressAnimator(status)

    progress.start()

    try:

        # STAGE 1: GEMINI

        progress.set_stage(
            "🧠 Gemini створює план і текст слайдів",
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

            # STAGE 3: PDF

            progress.set_stage(
                "📄 Готую PDF-файл",
                94
            )

            font_path = await ensure_unicode_font()

            await asyncio.to_thread(
                create_pdf,
                presentation_data,
                design,
                pdf_path,
                font_path
            )

            # FINISH

            await progress.finish(
                "✅ <b>Презентація готова!</b>\n\n"
                "📤 Надсилаю файли..."
            )

            await callback.message.answer_document(
                FSInputFile(
                    pptx_path,
                    filename="PrezentAI.pptx"
                ),
                caption=(
                    "📊 <b>PowerPoint</b>\n"
                    "Презентація з переходами Fade."
                )
            )

            await callback.message.answer_document(
                FSInputFile(
                    pdf_path,
                    filename="PrezentAI.pdf"
                ),
                caption=(
                    "📄 <b>PDF</b>\n"
                    "Версія для перегляду та друку."
                )
            )

        await callback.message.answer(
            "Що створимо наступного разу? 👇",
            reply_markup=MAIN_MENU
        )

    except GeminiAPIError as exc:

        log.warning(
            "Gemini generation failed: %s",
            exc
        )

        await progress.finish(
            "❌ <b>Gemini не зміг створити презентацію.</b>\n\n"
            f"Причина: {html.escape(str(exc))}\n\n"
            "Перевір API-ключ, модель та квоти Google AI Studio."
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


# ============================================================
# 11. GEMINI ERROR
# ============================================================

class GeminiAPIError(Exception):
    pass


# ============================================================
# 12. JSON PARSER
# ============================================================

def extract_json(text: str) -> dict:

    if not isinstance(text, str) or not text.strip():

        raise GeminiAPIError(
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

        raise GeminiAPIError(
            "Gemini повернув відповідь без JSON."
        )

    try:

        return json.loads(
            cleaned[start:end + 1]
        )

    except json.JSONDecodeError as exc:

        raise GeminiAPIError(
            "Gemini повернув некоректний JSON."
        ) from exc


# ============================================================
# 13. GEMINI 3.8 FLASH API
# ============================================================

async def generate_content(data: dict) -> dict:

    topic = data["topic"]
    count = data["slides"]

    language = LANGUAGES.get(
        data["language"],
        "українською"
    )

    prompt = f"""
Створи професійну та змістовну презентацію.

Тема: {topic}
Кількість слайдів: рівно {count}
Мова: {language}

Поверни виключно JSON у такому форматі:

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

- Рівно {count} слайдів.
- Логічна структура: вступ, основна частина, висновки.
- 3-5 змістовних тез на кожному слайді.
- Без повторів.
- Не вигадуй статистику.
- Грамотна мова.
- Короткі тези, придатні для слайда.
- Не додавай Markdown.
- Не додавай пояснення поза JSON.
""".strip()

    headers = {
        "x-goog-api-key": GEMINI_API_KEY,
        "Content-Type": "application/json"
    }

    payload = {
        "systemInstruction": {
            "parts": [
                {
                    "text": (
                        "Ти професійний автор навчальних і ділових "
                        "презентацій. Створюй якісний, логічний, "
                        "точний і грамотний матеріал. "
                        "Повертай тільки валідний JSON."
                    )
                }
            ]
        },

        "contents": [
            {
                "role": "user",
                "parts": [
                    {
                        "text": prompt
                    }
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

    last_error = "Gemini не повернув результат."

    # Two attempts for temporary empty responses.

    for attempt in range(1, 3):

        try:

            async with httpx.AsyncClient(
                timeout=httpx.Timeout(
                    180.0,
                    connect=25.0
                )
            ) as client:

                response = await client.post(
                    GEMINI_URL,
                    headers=headers,
                    json=payload
                )

        except httpx.TimeoutException as exc:

            last_error = (
                "Час очікування Gemini вичерпано."
            )

            log.warning(
                "Gemini timeout; attempt=%s",
                attempt
            )

            if attempt == 1:

                await asyncio.sleep(3)
                continue

            raise GeminiAPIError(
                last_error
            ) from exc

        except httpx.HTTPError as exc:

            log.exception(
                "Gemini connection error"
            )

            raise GeminiAPIError(
                "Не вдалося підключитися до Google Gemini."
            ) from exc

        try:

            result = response.json()

        except ValueError as exc:

            raise GeminiAPIError(
                f"Google Gemini повернув некоректну відповідь. "
                f"HTTP {response.status_code}."
            ) from exc

        if response.status_code >= 400:

            error = result.get(
                "error",
                {}
            )

            detail = str(
                error.get(
                    "message",
                    "Невідома помилка Google Gemini"
                )
            )

            status_code = response.status_code

            log.error(
                "Gemini API error: HTTP %s, model=%s, message=%s",
                status_code,
                GEMINI_MODEL,
                detail[:500]
            )

            if status_code == 400:

                raise GeminiAPIError(
                    "Некоректний запит до Gemini (400). "
                    + detail[:250]
                )

            if status_code in (401, 403):

                raise GeminiAPIError(
                    "Google не дозволив доступ до API. "
                    "Перевір GEMINI_API_KEY, проєкт та дозволи. "
                    + detail[:220]
                )

            if status_code == 404:

                raise GeminiAPIError(
                    f"Модель {GEMINI_MODEL} не знайдена або "
                    "недоступна для цього API-проєкту. "
                    "Перевір назву моделі та доступ."
                )

            if status_code == 429:

                raise GeminiAPIError(
                    "Перевищено квоту Gemini. "
                    "Перевір ліміти Google AI Studio."
                )

            raise GeminiAPIError(
                f"Помилка Google Gemini HTTP {status_code}: "
                f"{detail[:250]}"
            )

        if result.get("error"):

            detail = result["error"].get(
                "message",
                "Помилка Google Gemini"
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

            block_reason = feedback.get(
                "blockReason",
                "невідома причина"
            )

            last_error = (
                "Gemini не повернув відповідь. "
                f"Причина: {block_reason}."
            )

            log.warning(last_error)

            if attempt == 1:

                await asyncio.sleep(3)
                continue

            raise GeminiAPIError(
                last_error
            )

        candidate = candidates[0]

        finish_reason = candidate.get(
            "finishReason",
            "UNKNOWN"
        )

        candidate_content = candidate.get(
            "content",
            {}
        )

        parts = candidate_content.get(
            "parts",
            []
        )

        text_parts = []

        for part in parts:

            if not isinstance(part, dict):
                continue

            # Не включаємо внутрішні thought-частини.

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
            "Gemini result: model=%s finish=%s "
            "prompt_tokens=%s output_tokens=%s "
            "thought_tokens=%s chars=%s attempt=%s",
            GEMINI_MODEL,
            finish_reason,
            usage.get("promptTokenCount"),
            usage.get("candidatesTokenCount"),
            usage.get("thoughtsTokenCount"),
            len(content),
            attempt
        )

        if finish_reason == "SAFETY":

            raise GeminiAPIError(
                "Gemini заблокував відповідь фільтром безпеки."
            )

        if finish_reason == "MAX_TOKENS":

            raise GeminiAPIError(
                "Gemini вичерпав ліміт вихідних токенів. "
                "Спробуй створити менше слайдів."
            )

        if not content:

            last_error = (
                "Gemini повернув порожню відповідь. "
                f"Причина завершення: {finish_reason}."
            )

            log.warning(last_error)

            if attempt == 1:

                await asyncio.sleep(3)
                continue

            raise GeminiAPIError(
                last_error
            )

        try:

            parsed = extract_json(content)

        except GeminiAPIError as exc:

            last_error = str(exc)

            log.warning(
                "Gemini JSON parsing error: %s",
                last_error
            )

            if attempt == 1:

                await asyncio.sleep(2)
                continue

            raise

        slides = parsed.get("slides")

        if not isinstance(slides, list):

            raise GeminiAPIError(
                "У відповіді Gemini немає списку слайдів."
            )

        if len(slides) != count:

            last_error = (
                f"Gemini створив {len(slides)} "
                f"замість {count} слайдів."
            )

            if attempt == 1:

                await asyncio.sleep(2)
                continue

            raise GeminiAPIError(
                last_error
            )

        normalized = []

        for index, slide in enumerate(
            slides,
            start=1
        ):

            if not isinstance(slide, dict):

                raise GeminiAPIError(
                    f"Некоректний слайд №{index}."
                )

            title = str(
                slide.get(
                    "title",
                    f"Слайд {index}"
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
                "title": title or f"Слайд {index}",
                "content": points[:6]
            })

        if len(normalized) != count:

            raise GeminiAPIError(
                "Не вдалося обробити всі слайди."
            )

        log.info(
            "Gemini presentation generated successfully: %s slides",
            count
        )

        return {
            "title": str(
                parsed.get(
                    "title",
                    topic
                )
            ).strip()[:180],

            "slides": normalized
        }

    raise GeminiAPIError(
        last_error
    )


# ============================================================
# 14. POWERPOINT TRANSITIONS
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
# 15. POWERPOINT TEXT HELPER
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
# 16. CREATE POWERPOINT
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

        # Decorative accent rail.

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

        # Header.

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
            f"{index + 1:02d}  /  {len(data['slides']):02d}",
            11.35,
            0.38,
            1.1,
            0.3,
            10,
            accent,
            True,
            PP_ALIGN.RIGHT
        )

        # Slide title.

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

        # Main content card.

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

        # Fade transition.

        add_fade_transition(slide)

    prs.save(output_path)


# ============================================================
# 17. PDF FONT
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

    font_url = (
        "https://raw.githubusercontent.com/"
        "web-fonts/ttf/master/dejavu-sans-webfont.ttf"
    )

    try:

        async with httpx.AsyncClient(
            timeout=40.0,
            follow_redirects=True
        ) as client:

            response = await client.get(
                font_url
            )

            response.raise_for_status()

        if not response.content.startswith(
            b"\x00\x01\x00\x00"
        ):

            raise RuntimeError(
                "Downloaded font is not a valid TrueType file."
            )

        if len(response.content) < 100_000:

            raise RuntimeError(
                "Downloaded font file is too small."
            )

        target.write_bytes(
            response.content
        )

        return str(target)

    except Exception as exc:

        log.exception(
            "Could not download Unicode font"
        )

        raise RuntimeError(
            "Не вдалося завантажити шрифт для українського PDF."
        ) from exc


# ============================================================
# 18. CREATE PDF
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
        spaceAfter=8,
    )

    body_style = ParagraphStyle(
        "SlideBody",
        fontName="PrezentSans",
        fontSize=15,
        leading=22,
        textColor=body_color,
        alignment=TA_LEFT,
        spaceAfter=13,
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

            Spacer(1, 14),
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
                "Some PDF text did not fit on slide %s",
                index + 1
            )

        pdf.showPage()

    pdf.save()


# ============================================================
# 19. RENDER HTTP HEALTH SERVER
# ============================================================

async def health_handler(
    reader: asyncio.StreamReader,
    writer: asyncio.StreamWriter
):

    try:

        await asyncio.wait_for(
            reader.read(2048),
            timeout=5
        )

        body = b"Prezent AI is running"

        response = (
            b"HTTP/1.1 200 OK\r\n"
            b"Content-Type: text/plain; charset=utf-8\r\n"
            + f"Content-Length: {len(body)}\r\n".encode()
            + b"Connection: close\r\n\r\n"
            + body
        )

        writer.write(response)

        await writer.drain()

    except Exception:

        log.exception(
            "Health endpoint error"
        )

    finally:

        writer.close()

        try:

            await writer.wait_closed()

        except Exception:
            pass


# ============================================================
# 20. START BOT AND HTTP SERVER
# ============================================================

async def main():

    port = int(
        os.getenv("PORT", "10000")
    )

    server = await asyncio.start_server(
        health_handler,
        "0.0.0.0",
        port
    )

    log.info(
        "HTTP health server listening on port %s",
        port
    )

    log.info(
        "Starting Prezent AI Telegram polling"
    )

    log.info(
        "Gemini model: %s",
        GEMINI_MODEL
    )

    try:

        async with server:

            await dp.start_polling(bot)

    finally:

        server.close()

        await server.wait_closed()

        await bot.session.close()


if __name__ == "__main__":

    asyncio.run(main())
