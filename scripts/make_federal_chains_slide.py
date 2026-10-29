#!/usr/bin/env python3
"""
Создать слайд про развитие TITAN в федеральных сетях.
4 сети: Лемана Про (11/112), Петрович (0/48), Максидом (0/29), OBI (0/27).
"""

import os, shutil
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

SRC = "/home/z/my-project/upload/Titan -Дорожная_карта_fin_16.07.2026.pptx"
DST = "/home/z/my-project/download/Titan_Дорожная_карта_Федсети_v1.pptx"
shutil.copy(SRC, DST)

prs = Presentation(DST)

slide_layout = prs.slide_layouts[6]
new_slide = prs.slides.add_slide(slide_layout)

sldIdLst = prs.slides._sldIdLst
sldIds = list(sldIdLst)
new_sldId = sldIds[-1]
sldIdLst.remove(new_sldId)
sldIdLst.insert(7, new_sldId)  # позиция 8 (после B2B)

slide = new_slide

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
RED = RGBColor(0xC6, 0x28, 0x28)
RED_LIGHT = RGBColor(0xFF, 0xEB, 0xEE)
BLUE = RGBColor(0x15, 0x65, 0xC0)
BLUE_LIGHT = RGBColor(0xE3, 0xF2, 0xFD)

# === Данные по 4 сетям ===
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
        "regions": "ЮФО: Ростовская, Краснодарский, Волгоградская, Ставропольский, Сев.Осетия, Адыгея",
        "plan": "Расширение на 101 ТТ в остальных регионах: ЦФО, СЗФО, ПФО, УрФО, СибФО",
        "card_color": ORANGE,
    },
    {
        "name": "ПЕТРОВИЧ",
        "subtitle": "СЗФО · ЦФО · УФО",
        "total": "48",
        "cities": "СЗФО + ЦФО + УФО",
        "current": "0",
        "status": "ПЛАН ВХОДА",
        "status_color": BLUE,
        "status_bg": BLUE_LIGHT,
        "regions": "Санкт-Петербург, Москва, Ленинградская и Московская обл., Урал",
        "plan": "Вход через 3-5 пилотных ТТ в СПБ и Москве. Точки класса A (строительно-торговые центры)",
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
        "regions": "СПБ, Нижний Новгород, Казань, Екатеринбург, Самара, Уфа, Москва, МО",
        "plan": "Пилот в СПБ и Екатеринбурге — 2-3 ТТ. Гипермаркеты DIY, сильный B2C сегмент",
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
        "regions": "Москва (8), СПБ (5), Нижний Новгород (2), Казань, Самара, Краснодар, др.",
        "plan": "Вход через 2 пилотных ТТ в Москве. Восстановление сети после 2022 г.",
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
        shape.adjustments[0] = 0.05
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

# === Размеры слайда ===
SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)

# === 1. Фон ===
add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, BG)

# === 2. Верхняя полоса ===
add_rect(slide, 0, 0, SLIDE_W, Inches(0.12), ORANGE)

# === 3. Заголовок ===
add_text(slide, Inches(0.5), Inches(0.22), Inches(11), Inches(0.45),
         "РАЗВИТИЕ TITAN В ФЕДЕРАЛЬНЫХ СЕТЯХ",
         font_size=22, bold=True, color=DARK, anchor=MSO_ANCHOR.MIDDLE)

add_text(slide, Inches(0.5), Inches(0.70), Inches(12), Inches(0.3),
         "Сейчас: 11 ТТ Лемана Про (ЮФО) · Потенциал: 216 ТТ в 4 федеральных сетях",
         font_size=10, italic=True, color=GRAY)

add_text(slide, Inches(11.0), Inches(0.22), Inches(2.0), Inches(0.45),
         "TITAN · ФС",
         font_size=12, bold=True, color=ORANGE, align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)

# === 4. Сводные KPI сверху ===
KPI_TOP = Inches(1.20)
KPI_HEIGHT = Inches(0.70)
KPI_WIDTH = Inches(4.05)
KPI_GAP = Inches(0.10)

KPIS = [
    ("11", "ТТ СЕЙЧАС", "Лемана Про, ЮФО — активно", ORANGE, ORANGE_LIGHT),
    ("216", "ТТ ПОТЕНЦИАЛ", "4 федеральные сети РФ", ORANGE_DARK, ORANGE_LIGHT),
    ("4", "ФЕД. СЕТИ", "Лемана · Петрович · Максидом · OBI", DARK, LIGHT_GRAY),
]

for i, (num, label, desc, accent, bg) in enumerate(KPIS):
    left = Inches(0.5) + i * (KPI_WIDTH + KPI_GAP)
    add_round_rect(slide, left, KPI_TOP, KPI_WIDTH, KPI_HEIGHT, bg,
                   line_color=accent, line_width=Pt(0.75))
    add_text(slide, left + Inches(0.20), KPI_TOP + Inches(0.10), Inches(1.2), Inches(0.55),
             num, font_size=26, bold=True, color=accent, anchor=MSO_ANCHOR.MIDDLE)
    add_text(slide, left + Inches(1.50), KPI_TOP + Inches(0.05), KPI_WIDTH - Inches(1.6), Inches(0.30),
             label, font_size=10, bold=True, color=DARK, anchor=MSO_ANCHOR.TOP)
    add_text(slide, left + Inches(1.50), KPI_TOP + Inches(0.32), KPI_WIDTH - Inches(1.6), Inches(0.40),
             desc, font_size=8, italic=True, color=GRAY, anchor=MSO_ANCHOR.TOP, line_spacing=1.15)

# === 5. 4 карточки сетей (2×2) ===
CARD_TOP = Inches(2.20)
CARD_HEIGHT = Inches(2.20)
CARD_WIDTH = Inches(6.10)
CARD_GAP_X = Inches(0.15)
CARD_GAP_Y = Inches(0.15)
CARD_LEFT_START = Inches(0.5)

for i, chain in enumerate(CHAINS):
    row = i // 2
    col = i % 2
    left = CARD_LEFT_START + col * (CARD_WIDTH + CARD_GAP_X)
    top = CARD_TOP + row * (CARD_HEIGHT + CARD_GAP_Y)
    
    accent = chain["card_color"]
    
    # Фон карточки
    add_round_rect(slide, left, top, CARD_WIDTH, CARD_HEIGHT, WHITE,
                   line_color=LIGHT_GRAY, line_width=Pt(0.75))
    
    # Левая акцентная полоска
    add_rect(slide, left, top, Inches(0.08), CARD_HEIGHT, accent)
    
    # === ЛЕВАЯ ЧАСТЬ КАРТОЧКИ: название + цифры ===
    # Название сети (большим шрифтом)
    add_text(slide, left + Inches(0.25), top + Inches(0.12), Inches(2.8), Inches(0.40),
             chain["name"], font_size=18, bold=True, color=DARK, anchor=MSO_ANCHOR.TOP)
    # Подзаголовок
    add_text(slide, left + Inches(0.25), top + Inches(0.50), Inches(2.8), Inches(0.25),
             chain["subtitle"], font_size=9, italic=True, color=GRAY, anchor=MSO_ANCHOR.TOP)
    
    # Всего ТТ (крупно)
    add_text(slide, left + Inches(0.25), top + Inches(0.85), Inches(2.8), Inches(0.50),
             chain["total"] + " ТТ", font_size=28, bold=True, color=accent, anchor=MSO_ANCHOR.TOP)
    # Городов
    add_text(slide, left + Inches(0.25), top + Inches(1.40), Inches(2.8), Inches(0.25),
             chain["cities"], font_size=9, color=GRAY, anchor=MSO_ANCHOR.TOP)
    
    # Статус (плашка)
    status_top = top + Inches(1.70)
    add_round_rect(slide, left + Inches(0.25), status_top, Inches(2.2), Inches(0.35), chain["status_bg"])
    add_text(slide, left + Inches(0.25), status_top, Inches(2.2), Inches(0.35),
             chain["status"], font_size=10, bold=True, color=chain["status_color"],
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    
    # С текущим охватом (для Леманы)
    if chain["current"] != "0":
        add_text(slide, left + Inches(2.55), status_top, Inches(1.0), Inches(0.35),
                 f"ОХВАЧЕНО: {chain['current']}", font_size=9, bold=True, color=GREEN,
                 align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)
    
    # === ПРАВАЯ ЧАСТЬ КАРТОЧКИ: регионы + план ===
    RIGHT_LEFT = left + Inches(3.30)
    RIGHT_WIDTH = CARD_WIDTH - Inches(3.5)
    
    # Регионы присутствия
    add_text(slide, RIGHT_LEFT, top + Inches(0.12), RIGHT_WIDTH, Inches(0.20),
             "РЕГИОНЫ:", font_size=8, bold=True, color=accent, anchor=MSO_ANCHOR.TOP)
    add_text(slide, RIGHT_LEFT, top + Inches(0.32), RIGHT_WIDTH, Inches(0.50),
             chain["regions"], font_size=8, color=DARK, anchor=MSO_ANCHOR.TOP, line_spacing=1.25)
    
    # План развития
    plan_top = top + Inches(0.95)
    add_text(slide, RIGHT_LEFT, plan_top, RIGHT_WIDTH, Inches(0.20),
             "ПЛАН TITAN:", font_size=8, bold=True, color=accent, anchor=MSO_ANCHOR.TOP)
    add_text(slide, RIGHT_LEFT, plan_top + Inches(0.20), RIGHT_WIDTH, Inches(1.0),
             chain["plan"], font_size=8, color=DARK, anchor=MSO_ANCHOR.TOP, line_spacing=1.25)

# === 6. Нижняя плашка-итог ===
SUMMARY_TOP = Inches(6.95)
add_round_rect(slide, Inches(0.5), SUMMARY_TOP, Inches(12.33), Inches(0.35), DARK)
add_text(slide, Inches(0.5), SUMMARY_TOP, Inches(12.33), Inches(0.35),
         "СТРАТЕГИЯ: от 11 ТТ (Лемана ЮФО) → 216 ТТ (4 фед. сети) · Поэтапный вход через пилоты в каждой сети · 2026-2027",
         font_size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# === 7. Нижняя полоска ===
add_rect(slide, 0, Inches(7.30), SLIDE_W, Inches(0.20), ORANGE)

# Сохраняем
prs.save(DST)
print(f"\n✅ Файл сохранён: {DST}")
print(f"   Размер: {os.path.getsize(DST):,} байт")
print(f"   Слайд вставлен на позицию 8 (после B2B)")
print(f"   Всего слайдов: {len(prs.slides)}")
