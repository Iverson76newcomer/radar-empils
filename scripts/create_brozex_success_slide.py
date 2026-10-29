#!/usr/bin/env python3
"""
Слайд «Опыт успеха: БРОЗЕКС ГРУПП × ЭМПИЛС» — версия с реальными данными.
Стратегия партнёра: эффективность и оптимизация вместо экстенсивного роста.
"""

import os
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

# === Брендовые константы ===
ORANGE = RGBColor(0xF4, 0x76, 0x29)
ORANGE_DARK = RGBColor(0xE5, 0x5A, 0x10)
ORANGE_LIGHT = RGBColor(0xFB, 0xC9, 0x9A)
DARK = RGBColor(0x22, 0x22, 0x22)
GRAY = RGBColor(0x6B, 0x6B, 0x6B)
LIGHT_GRAY = RGBColor(0xE8, 0xE8, 0xE8)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
BG = RGBColor(0xFA, 0xFA, 0xFA)
GREEN = RGBColor(0x4C, 0xAF, 0x50)
BLUE = RGBColor(0x21, 0x96, 0xF3)
RED_POSITIVE = RGBColor(0xD3, 0x2F, 0x2F)

BROZEX_RED = RGBColor(0xE3, 0x1E, 0x24)
BROZEX_BLUE = RGBColor(0x14, 0x4E, 0x9C)

FONT = "Calibri"

IMG_AUTO = "/home/z/my-project/upload/Brozex group - компания.jpg"
IMG_MAP = "/home/z/my-project/upload/Brozex group - карта регионов 2.jpg"

OUT_PATH = "/home/z/my-project/download/Эмпилс_Слайд_Успех_Brozex.pptx"

# === Создаём презентацию 16:9 ===
prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

slide = prs.slides.add_slide(prs.slide_layouts[6])

# === Фон ===
bg = slide.background
bg.fill.solid()
bg.fill.fore_color.rgb = BG

# === Верхняя оранжевая полоса ===
top_bar = slide.shapes.add_shape(
    MSO_SHAPE.RECTANGLE,
    Inches(0), Inches(0),
    prs.slide_width, Inches(0.25)
)
top_bar.fill.solid()
top_bar.fill.fore_color.rgb = ORANGE
top_bar.line.fill.background()
top_bar.shadow.inherit = False

# === Тег-категория ===
tag = slide.shapes.add_textbox(Inches(0.6), Inches(0.45), Inches(7), Inches(0.35))
tag_tf = tag.text_frame
tag_tf.margin_top = Pt(0); tag_tf.margin_bottom = Pt(0)
p = tag_tf.paragraphs[0]
r = p.add_run(); r.text = "КЕЙС ПАРТНЁРА  ·  ОПЫТ УСПЕШНОГО РАЗВИТИЯ"
r.font.name = FONT; r.font.size = Pt(11); r.font.bold = True
r.font.color.rgb = ORANGE

# === Главный заголовок ===
title = slide.shapes.add_textbox(Inches(0.6), Inches(0.85), Inches(11), Inches(0.85))
title_tf = title.text_frame
title_tf.margin_top = Pt(0); title_tf.margin_bottom = Pt(0)
title_tf.word_wrap = True
p = title_tf.paragraphs[0]
p.alignment = PP_ALIGN.LEFT
r1 = p.add_run(); r1.text = "БРОЗЕКС ГРУПП"
r1.font.name = FONT; r1.font.size = Pt(36); r1.font.bold = True
r1.font.color.rgb = BROZEX_RED
r2 = p.add_run(); r2.text = "  ×  ЭМПИЛС"
r2.font.name = FONT; r2.font.size = Pt(36); r2.font.bold = True
r2.font.color.rgb = ORANGE

# Подзаголовок
sub = slide.shapes.add_textbox(Inches(0.6), Inches(1.55), Inches(11), Inches(0.4))
sub_tf = sub.text_frame
sub_tf.margin_top = Pt(0); sub_tf.margin_bottom = Pt(0)
p = sub_tf.paragraphs[0]
r = p.add_run()
r.text = "2023–2026: от экстенсивного роста к операционной эффективности — качество важнее количества"
r.font.name = FONT; r.font.size = Pt(14); r.font.italic = True
r.font.color.rgb = GRAY

# === Левая колонка: 2 фотографии ===
PHOTO_LEFT = Inches(0.6)
PHOTO_WIDTH = Inches(5.7)
PHOTO1_TOP = Inches(2.2)
PHOTO1_HEIGHT = Inches(2.5)
PHOTO2_TOP = Inches(4.85)
PHOTO2_HEIGHT = Inches(2.5)

slide.shapes.add_picture(
    IMG_AUTO,
    PHOTO_LEFT, PHOTO1_TOP,
    width=PHOTO_WIDTH, height=PHOTO1_HEIGHT
)
cap1 = slide.shapes.add_textbox(PHOTO_LEFT, Inches(4.72), Inches(5.7), Inches(0.3))
cap1_tf = cap1.text_frame
cap1_tf.margin_top = Pt(0); cap1_tf.margin_bottom = Pt(0)
p = cap1_tf.paragraphs[0]
p.alignment = PP_ALIGN.CENTER
r = p.add_run(); r.text = "Собственный автопарк Brozex — магистральные тягачи MAN, Hino, Hyundai"
r.font.name = FONT; r.font.size = Pt(10); r.font.italic = True
r.font.color.rgb = GRAY

slide.shapes.add_picture(
    IMG_MAP,
    PHOTO_LEFT, PHOTO2_TOP,
    width=PHOTO_WIDTH, height=PHOTO2_HEIGHT
)
cap2 = slide.shapes.add_textbox(PHOTO_LEFT, Inches(7.37), Inches(5.7), Inches(0.3))
cap2_tf = cap2.text_frame
cap2_tf.margin_top = Pt(0); cap2_tf.margin_bottom = Pt(0)
p = cap2_tf.paragraphs[0]
p.alignment = PP_ALIGN.CENTER
r = p.add_run(); r.text = "30+ городов присутствия от Москвы до Новосибирска · хаб — Екатеринбург"
r.font.name = FONT; r.font.size = Pt(10); r.font.italic = True
r.font.color.rgb = GRAY

# === Правая колонка: 6 KPI-карточек ===
KPI_LEFT = Inches(6.6)
KPI_TOP = Inches(2.2)
KPI_WIDTH = Inches(3.15)
KPI_HEIGHT = Inches(1.5)
KPI_GAP_X = Inches(0.15)
KPI_GAP_Y = Inches(0.2)

# === Карточка 1: АКБ — 1200 стабильно ===
def add_card_akb():
    left = KPI_LEFT
    top = KPI_TOP
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, KPI_WIDTH, KPI_HEIGHT)
    card.fill.solid(); card.fill.fore_color.rgb = WHITE
    card.line.color.rgb = LIGHT_GRAY; card.line.width = Pt(0.75)
    card.shadow.inherit = False
    try: card.adjustments[0] = 0.06
    except: pass

    # Header
    head = slide.shapes.add_textbox(left+Inches(0.15), top+Inches(0.1), KPI_WIDTH-Inches(0.3), Inches(0.3))
    htf = head.text_frame; htf.margin_top = Pt(0); htf.margin_bottom = Pt(0)
    p = htf.paragraphs[0]
    r = p.add_run(); r.text = "📐  АКБ"
    r.font.name = FONT; r.font.size = Pt(12); r.font.bold = True; r.font.color.rgb = ORANGE

    # Main number
    main = slide.shapes.add_textbox(left+Inches(0.15), top+Inches(0.42), KPI_WIDTH-Inches(0.3), Inches(0.5))
    mtf = main.text_frame; mtf.margin_top = Pt(0); mtf.margin_bottom = Pt(0)
    p = mtf.paragraphs[0]
    r = p.add_run(); r.text = "1 200"
    r.font.name = FONT; r.font.size = Pt(32); r.font.bold = True; r.font.color.rgb = DARK
    r2 = p.add_run(); r2.text = "  точек"
    r2.font.name = FONT; r2.font.size = Pt(12); r2.font.color.rgb = GRAY

    # Status
    stat = slide.shapes.add_textbox(left+Inches(0.15), top+Inches(0.95), KPI_WIDTH-Inches(0.3), Inches(0.25))
    stf = stat.text_frame; stf.margin_top = Pt(0); stf.margin_bottom = Pt(0)
    p = stf.paragraphs[0]
    r = p.add_run(); r.text = "● "
    r.font.name = FONT; r.font.size = Pt(11); r.font.color.rgb = GREEN
    r2 = p.add_run(); r2.text = "стабильно · удержание"
    r2.font.name = FONT; r2.font.size = Pt(11); r2.font.bold = True; r2.font.color.rgb = GREEN

    # Desc
    desc = slide.shapes.add_textbox(left+Inches(0.15), top+Inches(1.2), KPI_WIDTH-Inches(0.3), Inches(0.25))
    dtf = desc.text_frame; dtf.margin_top = Pt(0); dtf.margin_bottom = Pt(0)
    p = dtf.paragraphs[0]
    r = p.add_run(); r.text = "стратегия удержания, не расширения"
    r.font.name = FONT; r.font.size = Pt(9); r.font.italic = True; r.font.color.rgb = GRAY

# === Карточка 2: Команда ЭТП — 6 чел, 40% ===
def add_card_etp():
    left = KPI_LEFT + (KPI_WIDTH + KPI_GAP_X)
    top = KPI_TOP
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, KPI_WIDTH, KPI_HEIGHT)
    card.fill.solid(); card.fill.fore_color.rgb = WHITE
    card.line.color.rgb = LIGHT_GRAY; card.line.width = Pt(0.75)
    card.shadow.inherit = False
    try: card.adjustments[0] = 0.06
    except: pass

    head = slide.shapes.add_textbox(left+Inches(0.15), top+Inches(0.1), KPI_WIDTH-Inches(0.3), Inches(0.3))
    htf = head.text_frame; htf.margin_top = Pt(0); htf.margin_bottom = Pt(0)
    p = htf.paragraphs[0]
    r = p.add_run(); r.text = "👥  Команда ЭТП"
    r.font.name = FONT; r.font.size = Pt(12); r.font.bold = True; r.font.color.rgb = ORANGE

    main = slide.shapes.add_textbox(left+Inches(0.15), top+Inches(0.42), KPI_WIDTH-Inches(0.3), Inches(0.5))
    mtf = main.text_frame; mtf.margin_top = Pt(0); mtf.margin_bottom = Pt(0)
    p = mtf.paragraphs[0]
    r = p.add_run(); r.text = "6"
    r.font.name = FONT; r.font.size = Pt(32); r.font.bold = True; r.font.color.rgb = DARK
    r2 = p.add_run(); r2.text = "  человек"
    r2.font.name = FONT; r2.font.size = Pt(12); r2.font.color.rgb = GRAY

    stat = slide.shapes.add_textbox(left+Inches(0.15), top+Inches(0.95), KPI_WIDTH-Inches(0.3), Inches(0.25))
    stf = stat.text_frame; stf.margin_top = Pt(0); stf.margin_bottom = Pt(0)
    p = stf.paragraphs[0]
    r = p.add_run(); r.text = "▲ "
    r.font.name = FONT; r.font.size = Pt(11); r.font.color.rgb = ORANGE
    r2 = p.add_run(); r2.text = "доля 40% в общих продажах"
    r2.font.name = FONT; r2.font.size = Pt(11); r2.font.bold = True; r2.font.color.rgb = ORANGE

    desc = slide.shapes.add_textbox(left+Inches(0.15), top+Inches(1.2), KPI_WIDTH-Inches(0.3), Inches(0.25))
    dtf = desc.text_frame; dtf.margin_top = Pt(0); dtf.margin_bottom = Pt(0)
    p = dtf.paragraphs[0]
    r = p.add_run(); r.text = "команда Эмпилс в поле у партнёра"
    r.font.name = FONT; r.font.size = Pt(9); r.font.italic = True; r.font.color.rgb = GRAY

# === Карточка 3: Логистика 5% → 4% ===
def add_card_logistics():
    left = KPI_LEFT
    top = KPI_TOP + (KPI_HEIGHT + KPI_GAP_Y)
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, KPI_WIDTH, KPI_HEIGHT)
    card.fill.solid(); card.fill.fore_color.rgb = WHITE
    card.line.color.rgb = LIGHT_GRAY; card.line.width = Pt(0.75)
    card.shadow.inherit = False
    try: card.adjustments[0] = 0.06
    except: pass

    head = slide.shapes.add_textbox(left+Inches(0.15), top+Inches(0.1), KPI_WIDTH-Inches(0.3), Inches(0.3))
    htf = head.text_frame; htf.margin_top = Pt(0); htf.margin_bottom = Pt(0)
    p = htf.paragraphs[0]
    r = p.add_run(); r.text = "🚚  Стоимость логистики"
    r.font.name = FONT; r.font.size = Pt(12); r.font.bold = True; r.font.color.rgb = ORANGE

    main = slide.shapes.add_textbox(left+Inches(0.15), top+Inches(0.42), KPI_WIDTH-Inches(0.3), Inches(0.5))
    mtf = main.text_frame; mtf.margin_top = Pt(0); mtf.margin_bottom = Pt(0)
    p = mtf.paragraphs[0]
    r1 = p.add_run(); r1.text = "5%"
    r1.font.name = FONT; r1.font.size = Pt(20); r1.font.bold = True; r1.font.color.rgb = GRAY
    r2 = p.add_run(); r2.text = "  →  "
    r2.font.name = FONT; r2.font.size = Pt(14); r2.font.color.rgb = LIGHT_GRAY
    r3 = p.add_run(); r3.text = "4%"
    r3.font.name = FONT; r3.font.size = Pt(28); r3.font.bold = True; r3.font.color.rgb = ORANGE

    stat = slide.shapes.add_textbox(left+Inches(0.15), top+Inches(0.95), KPI_WIDTH-Inches(0.3), Inches(0.25))
    stf = stat.text_frame; stf.margin_top = Pt(0); stf.margin_bottom = Pt(0)
    p = stf.paragraphs[0]
    r = p.add_run(); r.text = "▼ "
    r.font.name = FONT; r.font.size = Pt(11); r.font.color.rgb = GREEN
    r2 = p.add_run(); r2.text = "−1 п.п. за 2023→2025"
    r2.font.name = FONT; r2.font.size = Pt(11); r2.font.bold = True; r2.font.color.rgb = GREEN

    desc = slide.shapes.add_textbox(left+Inches(0.15), top+Inches(1.2), KPI_WIDTH-Inches(0.3), Inches(0.25))
    dtf = desc.text_frame; dtf.margin_top = Pt(0); dtf.margin_bottom = Pt(0)
    p = dtf.paragraphs[0]
    r = p.add_run(); r.text = "% от оборота · собственный автопарк"
    r.font.name = FONT; r.font.size = Pt(9); r.font.italic = True; r.font.color.rgb = GRAY

# === Карточка 4: Стоимость тонны +26% ===
def add_card_price():
    left = KPI_LEFT + (KPI_WIDTH + KPI_GAP_X)
    top = KPI_TOP + (KPI_HEIGHT + KPI_GAP_Y)
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, KPI_WIDTH, KPI_HEIGHT)
    card.fill.solid(); card.fill.fore_color.rgb = WHITE
    card.line.color.rgb = LIGHT_GRAY; card.line.width = Pt(0.75)
    card.shadow.inherit = False
    try: card.adjustments[0] = 0.06
    except: pass

    head = slide.shapes.add_textbox(left+Inches(0.15), top+Inches(0.1), KPI_WIDTH-Inches(0.3), Inches(0.3))
    htf = head.text_frame; htf.margin_top = Pt(0); htf.margin_bottom = Pt(0)
    p = htf.paragraphs[0]
    r = p.add_run(); r.text = "💰  Стоимость тонны"
    r.font.name = FONT; r.font.size = Pt(12); r.font.bold = True; r.font.color.rgb = ORANGE

    main = slide.shapes.add_textbox(left+Inches(0.15), top+Inches(0.42), KPI_WIDTH-Inches(0.3), Inches(0.5))
    mtf = main.text_frame; mtf.margin_top = Pt(0); mtf.margin_bottom = Pt(0)
    p = mtf.paragraphs[0]
    r = p.add_run(); r.text = "+26%"
    r.font.name = FONT; r.font.size = Pt(32); r.font.bold = True; r.font.color.rgb = ORANGE

    stat = slide.shapes.add_textbox(left+Inches(0.15), top+Inches(0.95), KPI_WIDTH-Inches(0.3), Inches(0.25))
    stf = stat.text_frame; stf.margin_top = Pt(0); stf.margin_bottom = Pt(0)
    p = stf.paragraphs[0]
    r = p.add_run(); r.text = "▲ "
    r.font.name = FONT; r.font.size = Pt(11); r.font.color.rgb = ORANGE
    r2 = p.add_run(); r2.text = "2024 → 2026 (прогноз)"
    r2.font.name = FONT; r2.font.size = Pt(11); r2.font.bold = True; r2.font.color.rgb = ORANGE

    desc = slide.shapes.add_textbox(left+Inches(0.15), top+Inches(1.2), KPI_WIDTH-Inches(0.3), Inches(0.25))
    dtf = desc.text_frame; dtf.margin_top = Pt(0); dtf.margin_bottom = Pt(0)
    p = dtf.paragraphs[0]
    r = p.add_run(); r.text = "сдвиг в средний+ и премиум сегмент"
    r.font.name = FONT; r.font.size = Pt(9); r.font.italic = True; r.font.color.rgb = GRAY

# === Карточка 5: SKU — компактно 3 строки ===
def add_card_sku():
    left = KPI_LEFT
    top = KPI_TOP + 2 * (KPI_HEIGHT + KPI_GAP_Y)
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, KPI_WIDTH, KPI_HEIGHT)
    card.fill.solid(); card.fill.fore_color.rgb = WHITE
    card.line.color.rgb = LIGHT_GRAY; card.line.width = Pt(0.75)
    card.shadow.inherit = False
    try: card.adjustments[0] = 0.06
    except: pass

    head = slide.shapes.add_textbox(left+Inches(0.15), top+Inches(0.05), KPI_WIDTH-Inches(0.3), Inches(0.3))
    htf = head.text_frame; htf.margin_top = Pt(0); htf.margin_bottom = Pt(0)
    p = htf.paragraphs[0]
    r = p.add_run(); r.text = "📦  SKU: сеть + розница"
    r.font.name = FONT; r.font.size = Pt(12); r.font.bold = True; r.font.color.rgb = ORANGE

    # 3 строки подкатегорий
    rows = slide.shapes.add_textbox(left+Inches(0.15), top+Inches(0.38), KPI_WIDTH-Inches(0.3), Inches(1.05))
    rtf = rows.text_frame; rtf.margin_top = Pt(0); rtf.margin_bottom = Pt(0)
    rtf.word_wrap = True

    # Строка 1: Дистр
    p1 = rtf.paragraphs[0]; p1.alignment = PP_ALIGN.LEFT
    p1.space_after = Pt(2)
    r = p1.add_run(); r.text = "Дистр: "
    r.font.name = FONT; r.font.size = Pt(10); r.font.color.rgb = GRAY
    r2 = p1.add_run(); r2.text = "550 SKU"
    r2.font.name = FONT; r2.font.size = Pt(11); r2.font.bold = True; r2.font.color.rgb = DARK
    r3 = p1.add_run(); r3.text = " · очищено от неликвидов"
    r3.font.name = FONT; r3.font.size = Pt(9); r3.font.color.rgb = GREEN

    # Строка 2: Сеть матрица
    p2 = rtf.add_paragraph(); p2.alignment = PP_ALIGN.LEFT
    p2.space_after = Pt(2)
    r = p2.add_run(); r.text = "Сеть (матрица): "
    r.font.name = FONT; r.font.size = Pt(10); r.font.color.rgb = GRAY
    r2 = p2.add_run(); r2.text = "150→200"
    r2.font.name = FONT; r2.font.size = Pt(11); r2.font.bold = True; r2.font.color.rgb = DARK
    r3 = p2.add_run(); r3.text = " ▲ +33%"
    r3.font.name = FONT; r3.font.size = Pt(10); r3.font.bold = True; r3.font.color.rgb = ORANGE

    # Строка 3: Прогрузка в сеть
    p3 = rtf.add_paragraph(); p3.alignment = PP_ALIGN.LEFT
    p3.space_after = Pt(2)
    r = p3.add_run(); r.text = "Прогрузка в сеть: "
    r.font.name = FONT; r.font.size = Pt(10); r.font.color.rgb = GRAY
    r2 = p3.add_run(); r2.text = "375 SKU"
    r2.font.name = FONT; r2.font.size = Pt(11); r2.font.bold = True; r2.font.color.rgb = DARK

    # Строка 4: Розница
    p4 = rtf.add_paragraph(); p4.alignment = PP_ALIGN.LEFT
    r = p4.add_run(); r.text = "Розница: "
    r.font.name = FONT; r.font.size = Pt(10); r.font.color.rgb = GRAY
    r2 = p4.add_run(); r2.text = "550 SKU"
    r2.font.name = FONT; r2.font.size = Pt(11); r2.font.bold = True; r2.font.color.rgb = DARK

# === Карточка 6: ТО 25/26 — 4 подкатегории ===
def add_card_to():
    left = KPI_LEFT + (KPI_WIDTH + KPI_GAP_X)
    top = KPI_TOP + 2 * (KPI_HEIGHT + KPI_GAP_Y)
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, KPI_WIDTH, KPI_HEIGHT)
    card.fill.solid(); card.fill.fore_color.rgb = WHITE
    card.line.color.rgb = LIGHT_GRAY; card.line.width = Pt(0.75)
    card.shadow.inherit = False
    try: card.adjustments[0] = 0.06
    except: pass

    head = slide.shapes.add_textbox(left+Inches(0.15), top+Inches(0.05), KPI_WIDTH-Inches(0.3), Inches(0.3))
    htf = head.text_frame; htf.margin_top = Pt(0); htf.margin_bottom = Pt(0)
    p = htf.paragraphs[0]
    r = p.add_run(); r.text = "📈  Рост ТО 25→26"
    r.font.name = FONT; r.font.size = Pt(12); r.font.bold = True; r.font.color.rgb = ORANGE

    rows = slide.shapes.add_textbox(left+Inches(0.15), top+Inches(0.38), KPI_WIDTH-Inches(0.3), Inches(1.05))
    rtf = rows.text_frame; rtf.margin_top = Pt(0); rtf.margin_bottom = Pt(0)
    rtf.word_wrap = True

    # Первичка
    p1 = rtf.paragraphs[0]; p1.alignment = PP_ALIGN.LEFT
    p1.space_after = Pt(2)
    r = p1.add_run(); r.text = "Первичка: "
    r.font.name = FONT; r.font.size = Pt(10); r.font.color.rgb = GRAY
    r2 = p1.add_run(); r2.text = "+15% т / +20% ₽"
    r2.font.name = FONT; r2.font.size = Pt(11); r2.font.bold = True; r2.font.color.rgb = DARK

    # Вторичка
    p2 = rtf.add_paragraph(); p2.alignment = PP_ALIGN.LEFT
    p2.space_after = Pt(2)
    r = p2.add_run(); r.text = "Вторичка: "
    r.font.name = FONT; r.font.size = Pt(10); r.font.color.rgb = GRAY
    r2 = p2.add_run(); r2.text = "+5% т / +10% ₽"
    r2.font.name = FONT; r2.font.size = Pt(11); r2.font.bold = True; r2.font.color.rgb = DARK

    # Сеть дистры
    p3 = rtf.add_paragraph(); p3.alignment = PP_ALIGN.LEFT
    p3.space_after = Pt(2)
    r = p3.add_run(); r.text = "Сеть дистры: "
    r.font.name = FONT; r.font.size = Pt(10); r.font.color.rgb = GRAY
    r2 = p3.add_run(); r2.text = "+17% ₽"
    r2.font.name = FONT; r2.font.size = Pt(11); r2.font.bold = True; r2.font.color.rgb = ORANGE

    # Розница
    p4 = rtf.add_paragraph(); p4.alignment = PP_ALIGN.LEFT
    r = p4.add_run(); r.text = "Розница: "
    r.font.name = FONT; r.font.size = Pt(10); r.font.color.rgb = GRAY
    r2 = p4.add_run(); r2.text = "+2% ₽"
    r2.font.name = FONT; r2.font.size = Pt(11); r2.font.bold = True; r2.font.color.rgb = GRAY

# Рисуем все 6 карточек
add_card_akb()
add_card_etp()
add_card_logistics()
add_card_price()
add_card_sku()
add_card_to()

# === Нижняя CTA-полоса ===
bottom_cta = slide.shapes.add_shape(
    MSO_SHAPE.RECTANGLE,
    Inches(0), Inches(7.15),
    prs.slide_width, Inches(0.35)
)
bottom_cta.fill.solid()
bottom_cta.fill.fore_color.rgb = DARK
bottom_cta.line.fill.background()
bottom_cta.shadow.inherit = False

bottom_tf = bottom_cta.text_frame
bottom_tf.margin_top = Pt(0); bottom_tf.margin_bottom = Pt(0)
bottom_tf.vertical_anchor = MSO_ANCHOR.MIDDLE
p = bottom_tf.paragraphs[0]
p.alignment = PP_ALIGN.CENTER

r1 = p.add_run(); r1.text = "Brozex Group "
r1.font.name = FONT; r1.font.size = Pt(12); r1.font.bold = True; r1.font.color.rgb = BROZEX_RED

r2 = p.add_run(); r2.text = "(г. Екатеринбург, основан 1993)  ·  "
r2.font.name = FONT; r2.font.size = Pt(11); r2.font.color.rgb = WHITE

r3 = p.add_run(); r3.text = "партнёр Эмпилс с 2023 года  ·  "
r3.font.name = FONT; r3.font.size = Pt(11); r3.font.italic = True; r3.font.color.rgb = ORANGE_LIGHT

r4 = p.add_run(); r4.text = "brozex.ru"
r4.font.name = FONT; r4.font.size = Pt(11); r4.font.bold = True; r4.font.color.rgb = WHITE

# === Сохраняем ===
os.makedirs(os.path.dirname(OUT_PATH), exist_ok=True)
prs.save(OUT_PATH)
print(f"✅ Слайд сохранён: {OUT_PATH}")
print(f"   Размер: {os.path.getsize(OUT_PATH):,} байт")
