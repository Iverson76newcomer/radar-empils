#!/usr/bin/env python3
"""
Взять файл пользователя «Слайд фежеральные сети развитие.pptx» (1 слайд, 10"×6.25")
и полностью пересоздать слайд с 5 сетями — добавить Мегаполис (Калининград).

Размер слайда: 10" × 6.25" (стандартный 4:3, как у пользователя).
5 карточек в 2 ряда: 3 сверху + 2 снизу.
"""

import os, shutil
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

SRC = "/home/z/my-project/upload/Слайд фежеральные сети развитие.pptx"
DST = "/home/z/my-project/download/Слайд_Федсети_с_Мегаполисом.pptx"
shutil.copy(SRC, DST)

prs = Presentation(DST)
slide = prs.slides[0]

# Очищаем слайд
spTree = slide.shapes._spTree
shapes_to_remove = list(slide.shapes)
for shape in shapes_to_remove:
    sp = shape._element
    spTree.remove(sp)
print(f"Удалено shape'ов: {len(shapes_to_remove)}")

# === Цвета ===
ORANGE = RGBColor(0xF4, 0x76, 0x29)
ORANGE_DARK = RGBColor(0xD5, 0x5E, 0x0A)
ORANGE_LIGHT = RGBColor(0xFB, 0xE6, 0xD5)
DARK = RGBColor(0x31, 0x31, 0x31)
GRAY = RGBColor(0x6B, 0x6B, 0x6B)
LIGHT_GRAY = RGBColor(0xF5, 0xF5, 0xF5)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
BG = RGBColor(0xFA, 0xFA, 0xFA)
GREEN = RGBColor(0x2E, 0x7D, 0x32)
GREEN_LIGHT = RGBColor(0xE8, 0xF5, 0xE9)
BLUE = RGBColor(0x15, 0x65, 0xC0)
BLUE_LIGHT = RGBColor(0xE3, 0xF2, 0xFD)

# === Данные по 5 сетям ===
CHAINS = [
    {
        "name": "ЛЕМАНА ПРО",
        "subtitle": "(бывш. Леруа Мерлен)",
        "total": "112",
        "cities": "66 городов",
        "current": "11",
        "status": "АКТИВНО",
        "status_color": GREEN,
        "status_bg": GREEN_LIGHT,
        "regions": "ЮФО: Ростовская, Краснодарский, Волгоградская, Ставропольский, Сев. Осетия, Адыгея",
        "plan": "Расширение на 101 ТТ: ЦФО, СЗФО, ПФО, УрФО, СибФО",
        "card_color": ORANGE,
    },
    {
        "name": "ПЕТРОВИЧ",
        "subtitle": "СЗФО · ЦФО · УФО",
        "total": "48",
        "cities": "СЗФО+ЦФО+УФО",
        "current": "0",
        "status": "ПЛАН ВХОДА 2027",
        "status_color": BLUE,
        "status_bg": BLUE_LIGHT,
        "regions": "СПБ, Москва, Ленинградская и Московская обл., Урал",
        "plan": "3-5 пилотов в СПБ и Москве. Класс A (СТЦ)",
        "card_color": ORANGE_DARK,
    },
    {
        "name": "МАКСИДОМ",
        "subtitle": "12 регионов РФ",
        "total": "29",
        "cities": "12 регионов",
        "current": "0",
        "status": "ПЛАН ВХОДА",
        "status_color": BLUE,
        "status_bg": BLUE_LIGHT,
        "regions": "СПБ, Н.Новгород, Казань, Екатеринбург, Самара, Уфа, Москва, МО",
        "plan": "Пилот в СПБ и Екатеринбурге — 2-3 ТТ. Гипермаркеты DIY",
        "card_color": ORANGE_DARK,
    },
    {
        "name": "OBI",
        "subtitle": "Оби · 12-14 городов",
        "total": "27",
        "cities": "12-14 городов",
        "current": "0",
        "status": "ПЛАН ВХОДА",
        "status_color": BLUE,
        "status_bg": BLUE_LIGHT,
        "regions": "Москва (8), СПБ (5), Н.Новгород (2), Казань, Самара, Краснодар",
        "plan": "2 пилота в Москве. Восстановление сети после 2022 г.",
        "card_color": ORANGE_DARK,
    },
    {
        "name": "МЕГАПОЛИС",
        "subtitle": "Калининградская обл.",
        "total": "8",
        "cities": "Калининградская обл.",
        "current": "0",
        "status": "ПЛАН ВХОДА",
        "status_color": BLUE,
        "status_bg": BLUE_LIGHT,
        "regions": "Калининград и область (8 ТЦ: пр. Мира, Камская, Энергетиков и др.)",
        "plan": "Калининград — экспортный хаб на западе РФ. 2-3 пилотных ТТ",
        "card_color": ORANGE_DARK,
    },
]

# === Вспомогательные функции ===
def add_rect(slide, left, top, width, height, fill_color, line_color=None, line_width=None):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    if line_color is None:
        shape.line.fill.background()
    else:
        shape.line.color.rgb = line_color
        if line_width:
            shape.line.width = line_width
    shape.shadow.inherit = False
    return shape

def add_round_rect(slide, left, top, width, height, fill_color, line_color=None, line_width=None):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    if line_color is None:
        shape.line.fill.background()
    else:
        shape.line.color.rgb = line_color
        if line_width:
            shape.line.width = line_width
    shape.shadow.inherit = False
    try:
        shape.adjustments[0] = 0.04
    except:
        pass
    return shape

def add_text(slide, left, top, width, height, text, font_size=10, bold=False,
             color=DARK, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, font_name="Calibri",
             italic=False, line_spacing=1.0):
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.margin_left = Pt(0); tf.margin_right = Pt(0)
    tf.margin_top = Pt(0); tf.margin_bottom = Pt(0)
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    lines = text.split('\n')
    for i, line in enumerate(lines):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.alignment = align
        if line_spacing:
            p.line_spacing = line_spacing
        r = p.add_run()
        r.text = line
        r.font.name = font_name
        r.font.size = Pt(font_size)
        r.font.bold = bold
        r.font.italic = italic
        r.font.color.rgb = color
    return tb

# === Размеры слайда 10" × 6.25" ===
SLIDE_W = Inches(10.0)
SLIDE_H = Inches(6.25)

# === 1. Фон ===
add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, BG)

# === 2. Верхняя полоса ===
add_rect(slide, 0, 0, SLIDE_W, Inches(0.10), ORANGE)

# === 3. Заголовок ===
add_text(slide, Inches(0.3), Inches(0.18), Inches(8), Inches(0.35),
         "РАЗВИТИЕ TITAN В ФЕДЕРАЛЬНЫХ СЕТЯХ",
         font_size=18, bold=True, color=DARK, anchor=MSO_ANCHOR.MIDDLE)

add_text(slide, Inches(0.3), Inches(0.55), Inches(9), Inches(0.25),
         "Сейчас: 11 ТТ Лемана Про (ЮФО) · Потенциал: 224 ТТ в 5 сетях",
         font_size=9, italic=True, color=GRAY)

add_text(slide, Inches(8.5), Inches(0.18), Inches(1.4), Inches(0.35),
         "TITAN · ФС",
         font_size=10, bold=True, color=ORANGE, align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)

# === 4. Сводные KPI ===
KPI_TOP = Inches(0.95)
KPI_HEIGHT = Inches(0.55)
KPI_WIDTH = Inches(3.10)
KPI_GAP = Inches(0.10)

KPIS = [
    ("11", "ТТ СЕЙЧАС", "Лемана Про, ЮФО", ORANGE, ORANGE_LIGHT),
    ("224", "ТТ ПОТЕНЦИАЛ", "5 федеральных сетей", ORANGE_DARK, ORANGE_LIGHT),
    ("5", "ФЕД. СЕТЕЙ", "Лемана · Петрович · Максидом · OBI · Мегаполис", DARK, LIGHT_GRAY),
]

for i, (num, label, desc, accent, bg) in enumerate(KPIS):
    left = Inches(0.3) + i * (KPI_WIDTH + KPI_GAP)
    add_round_rect(slide, left, KPI_TOP, KPI_WIDTH, KPI_HEIGHT, bg,
                   line_color=accent, line_width=Pt(0.75))
    add_text(slide, left + Inches(0.15), KPI_TOP + Inches(0.08), Inches(0.9), Inches(0.40),
             num, font_size=20, bold=True, color=accent, anchor=MSO_ANCHOR.MIDDLE)
    add_text(slide, left + Inches(1.15), KPI_TOP + Inches(0.04), KPI_WIDTH - Inches(1.3), Inches(0.22),
             label, font_size=8, bold=True, color=DARK, anchor=MSO_ANCHOR.TOP)
    add_text(slide, left + Inches(1.15), KPI_TOP + Inches(0.24), KPI_WIDTH - Inches(1.3), Inches(0.30),
             desc, font_size=7, italic=True, color=GRAY, anchor=MSO_ANCHOR.TOP, line_spacing=1.1)

# === 5. 5 карточек сетей: 3 сверху + 2 снизу ===
# Размер: 3 карточки в ряд по 3.10" ширины, 2 в ряду по 4.75"
CARD_TOP_1 = Inches(1.70)
CARD_HEIGHT = Inches(2.05)
CARD_WIDTH_TOP = Inches(3.10)  # для 3 в ряд
CARD_GAP_X = Inches(0.10)
CARD_LEFT_START = Inches(0.3)

# Верхний ряд — 3 карточки (Лемана, Петрович, Максидом)
for i in range(3):
    chain = CHAINS[i]
    left = CARD_LEFT_START + i * (CARD_WIDTH_TOP + CARD_GAP_X)
    top = CARD_TOP_1
    accent = chain["card_color"]
    
    # Фон карточки
    add_round_rect(slide, left, top, CARD_WIDTH_TOP, CARD_HEIGHT, WHITE,
                   line_color=LIGHT_GRAY, line_width=Pt(0.75))
    # Левая акцентная полоска
    add_rect(slide, left, top, Inches(0.06), CARD_HEIGHT, accent)
    
    # Название
    add_text(slide, left + Inches(0.20), top + Inches(0.10), CARD_WIDTH_TOP - Inches(0.35), Inches(0.30),
             chain["name"], font_size=13, bold=True, color=DARK, anchor=MSO_ANCHOR.TOP)
    # Подзаголовок
    add_text(slide, left + Inches(0.20), top + Inches(0.38), CARD_WIDTH_TOP - Inches(0.35), Inches(0.20),
             chain["subtitle"], font_size=8, italic=True, color=GRAY, anchor=MSO_ANCHOR.TOP)
    
    # Всего ТТ
    add_text(slide, left + Inches(0.20), top + Inches(0.62), CARD_WIDTH_TOP - Inches(0.35), Inches(0.35),
             chain["total"] + " ТТ", font_size=20, bold=True, color=accent, anchor=MSO_ANCHOR.TOP)
    # Городов
    add_text(slide, left + Inches(0.20), top + Inches(0.98), CARD_WIDTH_TOP - Inches(0.35), Inches(0.18),
             chain["cities"], font_size=8, color=GRAY, anchor=MSO_ANCHOR.TOP)
    
    # Регионы
    add_text(slide, left + Inches(0.20), top + Inches(1.18), CARD_WIDTH_TOP - Inches(0.35), Inches(0.15),
             "РЕГИОНЫ:", font_size=7, bold=True, color=accent, anchor=MSO_ANCHOR.TOP)
    add_text(slide, left + Inches(0.20), top + Inches(1.32), CARD_WIDTH_TOP - Inches(0.35), Inches(0.35),
             chain["regions"], font_size=7, color=DARK, anchor=MSO_ANCHOR.TOP, line_spacing=1.15)
    
    # Статус
    status_top = top + CARD_HEIGHT - Inches(0.32)
    add_round_rect(slide, left + Inches(0.20), status_top, Inches(1.6), Inches(0.22), chain["status_bg"])
    add_text(slide, left + Inches(0.20), status_top, Inches(1.6), Inches(0.22),
             chain["status"], font_size=8, bold=True, color=chain["status_color"],
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    
    # Охвачено (для Леманы)
    if chain["current"] != "0":
        add_text(slide, left + Inches(1.85), status_top, Inches(1.0), Inches(0.22),
                 f"ОХВАЧЕНО: {chain['current']}", font_size=7, bold=True, color=GREEN,
                 align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)

# Нижний ряд — 2 карточки (OBI, Мегаполис)
CARD_TOP_2 = CARD_TOP_1 + CARD_HEIGHT + Inches(0.15)
CARD_WIDTH_BOTTOM = Inches(4.75)

for i in range(2):
    chain = CHAINS[3 + i]
    left = CARD_LEFT_START + i * (CARD_WIDTH_BOTTOM + CARD_GAP_X)
    top = CARD_TOP_2
    accent = chain["card_color"]
    
    # Фон карточки
    add_round_rect(slide, left, top, CARD_WIDTH_BOTTOM, CARD_HEIGHT, WHITE,
                   line_color=LIGHT_GRAY, line_width=Pt(0.75))
    # Левая акцентная полоска
    add_rect(slide, left, top, Inches(0.06), CARD_HEIGHT, accent)
    
    # ЛЕВАЯ ЧАСТЬ: название + цифры
    # Название
    add_text(slide, left + Inches(0.20), top + Inches(0.10), Inches(2.0), Inches(0.30),
             chain["name"], font_size=13, bold=True, color=DARK, anchor=MSO_ANCHOR.TOP)
    # Подзаголовок
    add_text(slide, left + Inches(0.20), top + Inches(0.38), Inches(2.0), Inches(0.20),
             chain["subtitle"], font_size=8, italic=True, color=GRAY, anchor=MSO_ANCHOR.TOP)
    
    # Всего ТТ
    add_text(slide, left + Inches(0.20), top + Inches(0.62), Inches(2.0), Inches(0.35),
             chain["total"] + " ТТ", font_size=20, bold=True, color=accent, anchor=MSO_ANCHOR.TOP)
    # Городов
    add_text(slide, left + Inches(0.20), top + Inches(0.98), Inches(2.0), Inches(0.18),
             chain["cities"], font_size=8, color=GRAY, anchor=MSO_ANCHOR.TOP)
    
    # Статус
    status_top = top + CARD_HEIGHT - Inches(0.32)
    add_round_rect(slide, left + Inches(0.20), status_top, Inches(1.8), Inches(0.22), chain["status_bg"])
    add_text(slide, left + Inches(0.20), status_top, Inches(1.8), Inches(0.22),
             chain["status"], font_size=8, bold=True, color=chain["status_color"],
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    
    # ПРАВАЯ ЧАСТЬ: регионы + план
    RIGHT_LEFT = left + Inches(2.40)
    RIGHT_WIDTH = CARD_WIDTH_BOTTOM - Inches(2.6)
    
    add_text(slide, RIGHT_LEFT, top + Inches(0.10), RIGHT_WIDTH, Inches(0.15),
             "РЕГИОНЫ:", font_size=7, bold=True, color=accent, anchor=MSO_ANCHOR.TOP)
    add_text(slide, RIGHT_LEFT, top + Inches(0.25), RIGHT_WIDTH, Inches(0.50),
             chain["regions"], font_size=7, color=DARK, anchor=MSO_ANCHOR.TOP, line_spacing=1.2)
    
    add_text(slide, RIGHT_LEFT, top + Inches(0.85), RIGHT_WIDTH, Inches(0.15),
             "ПЛАН TITAN:", font_size=7, bold=True, color=accent, anchor=MSO_ANCHOR.TOP)
    add_text(slide, RIGHT_LEFT, top + Inches(1.00), RIGHT_WIDTH, Inches(0.50),
             chain["plan"], font_size=7, color=DARK, anchor=MSO_ANCHOR.TOP, line_spacing=1.2)

# === 6. Нижняя плашка-итог ===
SUMMARY_TOP = Inches(5.90)
add_round_rect(slide, Inches(0.3), SUMMARY_TOP, Inches(9.4), Inches(0.28), DARK)
add_text(slide, Inches(0.3), SUMMARY_TOP, Inches(9.4), Inches(0.28),
         "СТРАТЕГИЯ: от 11 ТТ (Лемана ЮФО) → 224 ТТ (5 сетей) · Поэтапный вход через пилоты · 2026-2027",
         font_size=9, bold=True, color=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# Сохраняем
prs.save(DST)
print(f"\n✅ Файл сохранён: {DST}")
print(f"   Размер: {os.path.getsize(DST):,} байт")
