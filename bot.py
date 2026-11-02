
import os
import json
import asyncio
import logging
import tempfile
from pathlib import Path

import httpx
from aiogram import Bot, Dispatcher, F
from aiogram.filters import CommandStart
from aiogram.types import (
    Message, CallbackQuery, InlineKeyboardMarkup,
    InlineKeyboardButton, FSInputFile
)
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.context import FSMContext
from aiogram.fsm.storage.memory import MemoryStorage

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from reportlab.lib.pagesizes import A4, landscape
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

logging.basicConfig(level=logging.INFO)

BOT_TOKEN = os.getenv("BOT_TOKEN")
OPENROUTER_API_KEY = os.getenv("OPENROUTER_API_KEY")
MODEL = os.getenv("OPENROUTER_MODEL", "openrouter/free")

if not BOT_TOKEN:
    raise RuntimeError("Не задано BOT_TOKEN")

if not OPENROUTER_API_KEY:
    raise RuntimeError("Не задано OPENROUTER_API_KEY")

bot = Bot(token=BOT_TOKEN)
dp = Dispatcher(storage=MemoryStorage())

MAIN_MENU = InlineKeyboardMarkup(inline_keyboard=[
    [InlineKeyboardButton(text="🎨 Створити презентацію", callback_data="create")],
    [InlineKeyboardButton(text="📚 Можливості", callback_data="features")],
    [InlineKeyboardButton(text="ℹ️ Допомога", callback_data="help")]
])


class PresentationForm(StatesGroup):
    topic = State()
    slides = State()
    language = State()
    design = State()


def main_text():
    return (
        "✨ <b>Prezent AI</b>\n\n"
        "Твій особистий AI-генератор презентацій!\n\n"
        "🎨 Створюй красиві презентації за темою.\n"
        "📊 Отримуй PowerPoint і PDF.\n"
        "⚡ Економ свій час!"
    )


@dp.message(CommandStart())
async def start(message: Message, state: FSMContext):
    await state.clear()
    await message.answer(main_text(), reply_markup=MAIN_MENU)


@dp.callback_query(F.data == "features")
async def features(callback: CallbackQuery):
    await callback.message.answer(
        "🌟 <b>Можливості Prezent AI</b>\n\n"
        "• Створення тексту за темою\n"
        "• PowerPoint PPTX\n"
        "• PDF-файл\n"
        "• Вибір кількості слайдів\n"
        "• Українська та інші мови\n"
        "• Кілька стилів оформлення"
    )
    await callback.answer()


@dp.callback_query(F.data == "help")
async def help_handler(callback: CallbackQuery):
    await callback.message.answer(
        "Натисни «Створити презентацію», введи тему "
        "та вибери параметри. Бот створить файли для тебе."
    )
    await callback.answer()


@dp.callback_query(F.data == "create")
async def create_handler(callback: CallbackQuery, state: FSMContext):
    await state.set_state(PresentationForm.topic)
    await callback.message.answer(
        "📝 Напиши тему презентації.\n\n"
        "Наприклад: Штучний інтелект у сучасному світі"
    )
    await callback.answer()


@dp.message(PresentationForm.topic)
async def get_topic(message: Message, state: FSMContext):
    topic = (message.text or "").strip()

    if len(topic) < 3:
        await message.answer("Напиши трохи детальнішу тему.")
        return

    if len(topic) > 500:
        await message.answer("Тема занадто довга. Максимум 500 символів.")
        return

    await state.update_data(topic=topic)
    await state.set_state(PresentationForm.slides)

    keyboard = InlineKeyboardMarkup(inline_keyboard=[
        [
            InlineKeyboardButton(text="6", callback_data="slides:6"),
            InlineKeyboardButton(text="8", callback_data="slides:8"),
            InlineKeyboardButton(text="10", callback_data="slides:10"),
            InlineKeyboardButton(text="12", callback_data="slides:12")
        ]
    ])

    await message.answer("📊 Скільки слайдів потрібно?", reply_markup=keyboard)


@dp.callback_query(F.data.startswith("slides:"))
async def get_slides(callback: CallbackQuery, state: FSMContext):
    slides = int(callback.data.split(":")[1])
    await state.update_data(slides=slides)
    await state.set_state(PresentationForm.language)

    keyboard = InlineKeyboardMarkup(inline_keyboard=[
        [
            InlineKeyboardButton(text="🇺🇦 Українська", callback_data="lang:uk"),
            InlineKeyboardButton(text="🇬🇧 English", callback_data="lang:en")
        ],
        [
            InlineKeyboardButton(text="🇷🇺 Русский", callback_data="lang:ru")
        ]
    ])

    await callback.message.answer("🌍 Обери мову:", reply_markup=keyboard)
    await callback.answer()


@dp.callback_query(F.data.startswith("lang:"))
async def get_language(callback: CallbackQuery, state: FSMContext):
    language = callback.data.split(":")[1]
    await state.update_data(language=language)
    await state.set_state(PresentationForm.design)

    keyboard = InlineKeyboardMarkup(inline_keyboard=[
        [
            InlineKeyboardButton(text="🔵 Синій", callback_data="design:blue"),
            InlineKeyboardButton(text="🟣 Фіолетовий", callback_data="design:purple")
        ],
        [
            InlineKeyboardButton(text="⚫ Темний", callback_data="design:dark"),
            InlineKeyboardButton(text="⚪ Світлий", callback_data="design:light")
        ]
    ])

    await callback.message.answer("🎨 Обери дизайн:", reply_markup=keyboard)
    await callback.answer()


@dp.callback_query(F.data.startswith("design:"))
async def get_design(callback: CallbackQuery, state: FSMContext):
    design = callback.data.split(":")[1]
    data = await state.get_data()
    data["design"] = design
    await state.clear()

    await callback.answer("Починаю створення!")

    status = await callback.message.answer(
        "🚀 <b>Prezent AI запускається...</b>\n"
        "▰▱▱▱▱▱▱▱ 10%"
    )

    try:
        presentation_data = await generate_content(data, status)
        with tempfile.TemporaryDirectory() as temp_dir:
            pptx_path = str(Path(temp_dir) / "PrezentAI.pptx")
            pdf_path = str(Path(temp_dir) / "PrezentAI.pdf")

            create_pptx(presentation_data, data["design"], pptx_path)
            create_pdf(presentation_data, pdf_path)

            await status.edit_text("✨ <b>Презентація готова!</b>")

            await callback.message.answer_document(
                FSInputFile(pptx_path),
                caption="📊 Твоя презентація PowerPoint"
            )
            await callback.message.answer_document(
                FSInputFile(pdf_path),
                caption="📄 Версія PDF"
            )

    except Exception as error:
        logging.exception("Presentation generation failed")
        await status.edit_text(
            "❌ Не вдалося створити презентацію.\n"
            "Спробуй ще раз трохи пізніше."
        )


async def generate_content(data, status):
    language_names = {
        "uk": "Ukrainian",
        "en": "English",
        "ru": "Russian"
    }

    prompt = f"""
Create a detailed presentation in JSON.

Topic: {data['topic']}
Number of slides: {data['slides']}
Language: {language_names.get(data['language'], 'Ukrainian')}

Return ONLY valid JSON in this exact format:
{{
  "title": "Presentation title",
  "slides": [
    {{
      "title": "Slide title",
      "content": ["Point 1", "Point 2", "Point 3"],
      "notes": "Short explanation"
    }}
  ]
}}

Requirements:
- Exactly {data['slides']} slides.
- Educational, accurate and detailed.
- Logical structure: introduction, main content, conclusion.
- Avoid repeating information.
- Each slide should have 3-5 concise points.
- Do not use Markdown.
"""

    await status.edit_text("🧠 <b>AI створює матеріал...</b>\n▰▰▱▱▱▱▱▱ 25%")

    headers = {
        "Authorization": f"Bearer {OPENROUTER_API_KEY}",
        "Content-Type": "application/json",
        "HTTP-Referer": "https://github.com/a88845837-coder/PrezentAI",
        "X-Title": "Prezent AI"
    }

    payload = {
        "model": MODEL,
        "messages": [
            {
                "role": "system",
                "content": "You create accurate, structured presentations."
            },
            {"role": "user", "content": prompt}
        ],
        "temperature": 0.7
    }

    async with httpx.AsyncClient(timeout=180) as client:
        response = await client.post(
            "https://openrouter.ai/api/v1/chat/completions",
            headers=headers,
            json=payload
        )
        response.raise_for_status()
        result = response.json()

    raw = result["choices"][0]["message"]["content"]
    raw = raw.strip()

    if raw.startswith("```"):
        raw = raw.split("\n", 1)[1].rsplit("```", 1)[0].strip()

    parsed = json.loads(raw)

    if len(parsed.get("slides", [])) != data["slides"]:
        raise ValueError("AI returned an incorrect number of slides")

    await status.edit_text("🎨 <b>Створюю дизайн слайдів...</b>\n▰▰▰▰▱▱▱▱ 55%")

    return parsed


def create_pptx(data, design, output):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    themes = {
        "blue": ((15, 30, 65), (65, 130, 255)),
        "purple": ((35, 20, 65), (160, 100, 255)),
        "dark": ((25, 27, 32), (100, 190, 220)),
        "light": ((245, 247, 252), (55, 105, 190))
    }

    bg, accent = themes.get(design, themes["blue"])
    light = design == "light"

    for index, item in enumerate(data["slides"]):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = RGBColor(*bg)

        # Decorative accent bar
        bar = slide.shapes.add_shape(
            1, Inches(0), Inches(0), Inches(0.16), Inches(7.5)
        )
        bar.fill.solid()
        bar.fill.fore_color.rgb = RGBColor(*accent)
        bar.line.fill.background()

        title_box = slide.shapes.add_textbox(
            Inches(0.8), Inches(0.65), Inches(11.5), Inches(1.2)
        )
        tf = title_box.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = item.get("title", "Slide")
        p.font.size = Pt(30)
        p.font.bold = True
        p.font.color.rgb = RGBColor(
            (25, 35, 55) if light else (255, 255, 255)
        )

        content_box = slide.shapes.add_textbox(
            Inches(1.0), Inches(2.0), Inches(11.2), Inches(4.4)
        )
        tf = content_box.text_frame
        tf.word_wrap = True

        for i, point in enumerate(item.get("content", [])):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.text = "• " + str(point)
            p.font.size = Pt(21)
            p.font.color.rgb = RGBColor(
                (40, 50, 70) if light else (230, 235, 250)
            )
            p.space_after = Pt(18)

        footer = slide.shapes.add_textbox(
            Inches(10.7), Inches(7.0), Inches(2), Inches(0.3)
        )
        p = footer.text_frame.paragraphs[0]
        p.text = f"PREZENT AI  •  {index + 1}/{len(data['slides'])}"
        p.font.size = Pt(9)
        p.font.color.rgb = RGBColor(*accent)

    prs.save(output)


def create_pdf(data, output):
    pdf = canvas.Canvas(output, pagesize=landscape(A4))
    width, height = landscape(A4)

    font_path = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
    bold_path = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"

    if os.path.exists(font_path):
        pdfmetrics.registerFont(TTFont("DejaVu", font_path))
        pdfmetrics.registerFont(TTFont("DejaVuBold", bold_path))
        regular_font = "DejaVu"
        bold_font = "DejaVuBold"
    else:
        regular_font = "Helvetica"
        bold_font = "Helvetica-Bold"

    for index, item in enumerate(data["slides"]):
        pdf.setFillColorRGB(0.06, 0.12, 0.25)
        pdf.rect(0, 0, width, height, fill=1, stroke=0)

        pdf.setFillColorRGB(0.3, 0.55, 1)
        pdf.rect(0, 0, 12, height, fill=1, stroke=0)

        pdf.setFillColorRGB(1, 1, 1)
        pdf.setFont(bold_font, 25)
        pdf.drawString(55, height - 75, str(item.get("title", ""))[:90])

        y = height - 140
        pdf.setFont(regular_font, 16)

        for point in item.get("content", []):
            text = "• " + str(point)
            words = text.split()
            line = ""

            for word in words:
                if pdf.stringWidth(line + " " + word, regular_font, 16) > width - 120:
                    pdf.drawString(65, y, line)
                    y -= 25
                    line = word
                else:
                    line = (line + " " + word).strip()

            if line:
                pdf.drawString(65, y, line)
                y -= 35

        pdf.setFont(regular_font, 9)
        pdf.drawRightString(width - 35, 20, f"PREZENT AI • {index + 1}/{len(data['slides'])}")
        pdf.showPage()

    pdf.save()


async def main():
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())
