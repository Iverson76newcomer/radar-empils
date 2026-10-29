#!/usr/bin/env python3
"""
Создать слайд о продвижении TITAN в федеральной сети «Лемана ПРО»
и вставить его в презентацию Titan_Дорожная_карта_АКБ_v2.pptx.

Источник: Prodvizhenie_Lemana_ProTitan.xlsx, лист «Лемана Про».
"""

import os, shutil
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn

SRC = "/home/z/my-project/download/Titan_Дорожная_карта_АКБ_v2.pptx"
DST = "/home/z/my-project/download/Titan_Дорожная_карта_Лемана_v1.pptx"
shutil.copy(SRC, DST)

prs = Presentation(DST)
print(f"Слайдов в исходной презентации: {len(prs.slides)}")

# Вставляем новый слайд после слайда 5 (пилоты) — позиция 5 (индекс 5)
slide_layout = prs.slide_layouts[6]  # Blank
new_slide = prs.slides.add_slide(slide_layout)

sldIdLst = prs.slides._sldIdLst
sldIds = list(sldIdLst)
new_sldId = sldIds[-1]
sldIdLst.remove(new_sldId)
sldIdLst.insert(5, new_sldId)  # позиция 6 (после пилотов)

slide = new_slide

# === Цвета фирменные ===
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

# === Данные активностей ===
ACTIVITIES = [
    # (№, Название активности, Срок, Ответственный, Статус)
    ("1", "РК ТМ TITAN на 6 топовых экранах Ростова-на-Дону",
     "с 01.07 на 1 мес", "Литвиненко Н.", "РК запущена"),
    ("2", "Таргетированная реклама в интернете. ГЕО по 11 ТТ",
     "с 01.07 по 31.08", "Борисова Д.", "РК с 07.07 по 31.07 и с 01.08 по 30.08"),
    ("3", "Промо-акция «Подарок за покупку» в трафиковые дни",
     "17-26 июля", "Литвиненко Н.", "Ростов + Краснодар"),
    ("4", "Кешбэк 5% на TITAN во всех 11 ТТ",
     "с 01.07 по 31.08", "Пустовалов С.И.", "Запущен"),
    ("5", "Мастер-классы силами Эмпилс (3 ТТ Ростова)",
     "17 и 24 июля", "Эмпилс", "Проведены"),
    ("6", "Обеспечение выкрасами 11 ТТ (EPL + TITAN)",
     "до 16.07", "Литвиненко Н.", "В работе"),
    ("7", "Веерные выкрасы B2B Лемана Про (13 шт)",
     "до 22.07", "Литвиненко Н.", "Распределение"),
]

# 11 торговых точек
STORES = [
    ("1", "Ростов-на-Дону, ул. Доватора, 158", "94"),
    ("2", "Аксай, Аксайский пр-т, 23", "7"),
    ("3", "Верхнетемерницкий, ул. Обсерваторная, 15", "14"),
    ("4", "Волгоград, Университетский пр-т, 107 (Акварель)", "24"),
    ("5", "Краснодар, ул. Покрышкина, 5", "92"),
    ("6", "Краснодар, ул. Новороссийская, 234/1", "8"),
    ("7", "Новороссийск, ул. Ленина, 7", "165"),
    ("8", "Ставропольский край, с. Татарка", "69"),
    ("9", "Владикавказ, Архонское шоссе, 1А", "150"),
    ("10", "Адыгея, Тургеневское шоссе, 27/1", "147"),
    ("11", "Волжский, пр-т Ленина, 2Ж", "70"),
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

# === 1. Фон слайда ===
add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, BG)

# === 2. Верхняя оранжевая полоса ===
add_rect(slide, 0, 0, SLIDE_W, Inches(0.12), ORANGE)

# === 3. Заголовок ===
add_text(slide, Inches(0.5), Inches(0.22), Inches(11), Inches(0.45),
         "ПРОДВИЖЕНИЕ TITAN В ФЕДЕРАЛЬНОЙ СЕТИ «ЛЕМАНА ПРО»",
         font_size=20, bold=True, color=DARK, anchor=MSO_ANCHOR.MIDDLE)

add_text(slide, Inches(0.5), Inches(0.70), Inches(12), Inches(0.3),
         "11 торговых точек · 6 регионов ЮФО · 7 активностей · Цель: показать уходимость грунт-эмалей TITAN",
         font_size=10, italic=True, color=GRAY)

# Бренд-марка справа сверху
add_text(slide, Inches(11.0), Inches(0.22), Inches(2.0), Inches(0.45),
         "TITAN · 2026",
         font_size=12, bold=True, color=ORANGE, align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)

# === 4. Три KPI-плашки сверху (быстрый обзор) ===
KPI_TOP = Inches(1.25)
KPI_HEIGHT = Inches(0.70)
KPI_WIDTH = Inches(4.05)
KPI_GAP = Inches(0.10)

KPIS = [
    ("11", "ТОРГОВЫХ ТОЧЕК", "ЮФО: Ростовская, Краснодарский, Волгоградская, Ставропольский, Сев.Осетия, Адыгея"),
    ("7", "АКТИВНОСТЕЙ ПРОДВИЖЕНИЯ", "РК на экранах · Таргет · Промо · Кешбэк · Мастер-классы · Выкрасы"),
    ("2 мес", "ДЛИТЕЛЬНОСТЬ ПРОГРАММЫ", "1 июля — 31 августа 2026 г."),
]

for i, (num, label, desc) in enumerate(KPIS):
    left = Inches(0.5) + i * (KPI_WIDTH + KPI_GAP)
    add_round_rect(slide, left, KPI_TOP, KPI_WIDTH, KPI_HEIGHT, ORANGE_LIGHT,
                   line_color=ORANGE, line_width=Pt(0.75))
    # Большая цифра
    add_text(slide, left + Inches(0.20), KPI_TOP + Inches(0.10), Inches(1.2), Inches(0.55),
             num, font_size=26, bold=True, color=ORANGE_DARK, anchor=MSO_ANCHOR.MIDDLE)
    # Подпись
    add_text(slide, left + Inches(1.50), KPI_TOP + Inches(0.05), KPI_WIDTH - Inches(1.6), Inches(0.30),
             label, font_size=10, bold=True, color=DARK, anchor=MSO_ANCHOR.TOP)
    # Описание
    add_text(slide, left + Inches(1.50), KPI_TOP + Inches(0.32), KPI_WIDTH - Inches(1.6), Inches(0.40),
             desc, font_size=8, italic=True, color=GRAY, anchor=MSO_ANCHOR.TOP, line_spacing=1.15)

# === 5. Левая колонка — 7 активностей ===
ACT_TOP = Inches(2.20)
ACT_LEFT = Inches(0.5)
ACT_WIDTH = Inches(7.50)
ACT_HEIGHT_TOTAL = Inches(5.0)

# Заголовок секции
add_text(slide, ACT_LEFT, ACT_TOP, ACT_WIDTH, Inches(0.30),
         "7 АКТИВНОСТЕЙ ПРОДВИЖЕНИЯ",
         font_size=12, bold=True, color=ORANGE_DARK, anchor=MSO_ANCHOR.MIDDLE)

# Таблица активностей
TABLE_TOP = ACT_TOP + Inches(0.35)
ROW_HEIGHT = Inches(0.55)
HEADER_HEIGHT = Inches(0.30)

# Заголовок таблицы
add_rect(slide, ACT_LEFT, TABLE_TOP, ACT_WIDTH, HEADER_HEIGHT, DARK)
add_text(slide, ACT_LEFT + Inches(0.10), TABLE_TOP, Inches(0.35), HEADER_HEIGHT,
         "№", font_size=9, bold=True, color=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
add_text(slide, ACT_LEFT + Inches(0.50), TABLE_TOP, Inches(3.50), HEADER_HEIGHT,
         "Активность", font_size=9, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
add_text(slide, ACT_LEFT + Inches(4.10), TABLE_TOP, Inches(1.70), HEADER_HEIGHT,
         "Срок", font_size=9, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
add_text(slide, ACT_LEFT + Inches(5.90), TABLE_TOP, Inches(1.50), HEADER_HEIGHT,
         "Статус", font_size=9, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)

# Данные активностей
for i, (num, name, term, resp, status) in enumerate(ACTIVITIES):
    y = TABLE_TOP + HEADER_HEIGHT + i * ROW_HEIGHT
    bg_color = WHITE if i % 2 == 0 else LIGHT_GRAY
    add_rect(slide, ACT_LEFT, y, ACT_WIDTH, ROW_HEIGHT, bg_color)
    
    # Номер
    add_text(slide, ACT_LEFT + Inches(0.10), y, Inches(0.35), ROW_HEIGHT,
             num, font_size=11, bold=True, color=ORANGE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    # Активность
    add_text(slide, ACT_LEFT + Inches(0.50), y + Inches(0.05), Inches(3.50), Inches(0.45),
             name, font_size=9, bold=True, color=DARK, anchor=MSO_ANCHOR.TOP, line_spacing=1.1)
    # Срок
    add_text(slide, ACT_LEFT + Inches(4.10), y, Inches(1.70), ROW_HEIGHT,
             term, font_size=8, color=GRAY, anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.1)
    # Статус
    add_text(slide, ACT_LEFT + Inches(5.90), y, Inches(1.50), ROW_HEIGHT,
             status, font_size=8, italic=True, color=GREEN, anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.1)

# === 6. Правая колонка — список 11 ТТ + механика акции ===
RIGHT_LEFT = Inches(8.30)
RIGHT_WIDTH = Inches(4.55)

# Заголовок правой колонки
add_text(slide, RIGHT_LEFT, ACT_TOP, RIGHT_WIDTH, Inches(0.30),
         "11 ТОРГОВЫХ ТОЧЕК — ЮФО",
         font_size=12, bold=True, color=ORANGE_DARK, anchor=MSO_ANCHOR.MIDDLE)

# Список ТТ (компактный)
STORES_TOP = ACT_TOP + Inches(0.35)
STORE_ROW_HEIGHT = Inches(0.22)

for i, (num, addr, reg) in enumerate(STORES):
    y = STORES_TOP + i * STORE_ROW_HEIGHT
    bg_color = WHITE if i % 2 == 0 else LIGHT_GRAY
    add_rect(slide, RIGHT_LEFT, y, RIGHT_WIDTH, STORE_ROW_HEIGHT, bg_color)
    
    # Номер
    add_text(slide, RIGHT_LEFT + Inches(0.05), y, Inches(0.30), STORE_ROW_HEIGHT,
             num, font_size=9, bold=True, color=ORANGE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    # Адрес
    add_text(slide, RIGHT_LEFT + Inches(0.35), y, Inches(3.50), STORE_ROW_HEIGHT,
             addr, font_size=8, color=DARK, anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.0)
    # Рег №
    add_text(slide, RIGHT_LEFT + Inches(3.95), y, Inches(0.55), STORE_ROW_HEIGHT,
             reg, font_size=8, bold=True, color=GRAY, align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)

# === 7. Блок «Механика акции» (внизу справа) ===
MECH_TOP = STORES_TOP + 11 * STORE_ROW_HEIGHT + Inches(0.15)
MECH_HEIGHT = Inches(1.85)

add_round_rect(slide, RIGHT_LEFT, MECH_TOP, RIGHT_WIDTH, MECH_HEIGHT, ORANGE_LIGHT,
               line_color=ORANGE, line_width=Pt(0.75))

add_text(slide, RIGHT_LEFT + Inches(0.15), MECH_TOP + Inches(0.10), RIGHT_WIDTH - Inches(0.3), Inches(0.30),
         "🎁 МЕХАНИКА «ПОДАРОК ЗА ПОКУПКУ»",
         font_size=11, bold=True, color=ORANGE_DARK, anchor=MSO_ANCHOR.MIDDLE)

add_text(slide, RIGHT_LEFT + Inches(0.15), MECH_TOP + Inches(0.45), RIGHT_WIDTH - Inches(0.3), Inches(1.40),
         "• от 2 000 ₽ — кисть + перчатки в подарок\n• от 4 000 ₽ — сертификат Лемана ПРО 500 ₽\n• + 5% кешбэк на TITAN во всех 11 ТТ\n• Мастер-классы: 17 и 24 июля, 3 ТТ Ростова",
         font_size=9, color=DARK, anchor=MSO_ANCHOR.TOP, line_spacing=1.4)

# === 8. Подвал ===
FOOTER_TOP = Inches(7.00)
add_rect(slide, 0, FOOTER_TOP, SLIDE_W, Inches(0.30), DARK)
add_text(slide, Inches(0.5), FOOTER_TOP, Inches(9), Inches(0.30),
         "Empils  •  TITAN  •  Продвижение в федеральных сетях  •  2026",
         font_size=10, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
add_text(slide, Inches(10.0), FOOTER_TOP, Inches(3.0), Inches(0.30),
         "6 / 22",
         font_size=10, italic=True, color=RGBColor(0x90, 0x98, 0x9F), align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)

# Нижняя оранжевая полоска
add_rect(slide, 0, Inches(7.30), SLIDE_W, Inches(0.20), ORANGE)

# Сохраняем
prs.save(DST)
print(f"\n✅ Файл сохранён: {DST}")
print(f"   Размер: {os.path.getsize(DST):,} байт")
print(f"   Слайд вставлен на позицию 6 (после пилотов)")
print(f"   Всего слайдов: {len(prs.slides)}")
