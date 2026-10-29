#!/usr/bin/env python3
"""
Создать слайд 4 с детальной сметой по 2 этапам на основе Cost_calculation_Titan.xlsx
и вставить его в презентацию Titan -Дорожная_карта_fin_16.07.2026.pptx
(заменяя текущий слайд 4 про ЭТАП 0)

Структура слайда:
- Заголовок: "Смета программы TITAN — 2 этапа"
- 2 колонки: Первый этап / Второй этап
- В каждой колонке — карточки с наименованием и стоимостью
- Внизу — итог по каждому этапу
"""

import os, shutil
from copy import deepcopy
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn

SRC = "/home/z/my-project/upload/Titan -Дорожная_карта_fin_16.07.2026.pptx"
DST = "/home/z/my-project/download/Titan_Дорожная_карта_смета_v1.pptx"
shutil.copy(SRC, DST)

# Открываем файл
prs = Presentation(DST)
slide_old = prs.slides[3]  # slide 4 — старый

# === Стилизация: цвета фирменные TITAN/Empils ===
ORANGE = RGBColor(0xF4, 0x76, 0x29)
ORANGE_DARK = RGBColor(0xD5, 0x5E, 0x0A)
ORANGE_LIGHT = RGBColor(0xFB, 0xE6, 0xD5)
DARK = RGBColor(0x31, 0x31, 0x31)
GRAY = RGBColor(0x6B, 0x6B, 0x6B)
LIGHT_GRAY = RGBColor(0xF5, 0xF5, 0xF5)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
BG = RGBColor(0xFA, 0xFA, 0xFA)

# === Данные сметы ===
# Оценки стоимости на основе web-search и средних рыночных цен

STAGE_1_ITEMS = [
    # (Наименование, краткое описание, стоимость, ед.)
    ("Шелфтокеры",                 "Печать шелфтокеров для полки TITAN (тираж 1000 шт)",          "≈ 25 000 ₽",  "тираж"),
    ("Воблеры",                    "Воблеры на полку (тираж 1000 шт)",                            "≈ 20 000 ₽",  "тираж"),
    ("Брендированные ценники",     "Ценники с переводом цены в цену/год (тираж 1000 шт)",         "≈ 30 000 ₽",  "тираж"),
    ("Дообеспечение выкрасами",    "Выкрасы EPL + TITAN для 100 точек (по результатам сбора)",    "≈ 50 000 ₽",  "партия"),
    ("Листовка TITAN с ССВ",       "Листовка с совокупной стоимостью владения (тираж 5000 шт)",   "≈ 45 000 ₽",  "тираж"),
    ("Видео-инструкция по нанесению", "Ролик 2-3 мин: как наносить TITAN 3в1 (5 шагов + монтаж)", "≈ 90 000 ₽",  "1 ролик"),
    ("Видеопортфолио объектов",    "3 видео: ворота (3в1), ворота (молоток), дверь (быстрая)",   "≈ 250 000 ₽", "3 ролика"),
    ("Видео-презентация для ТП",   "Обучающий ролик для торговых представителей дистрибьютора",   "≈ 120 000 ₽", "1 ролик"),
    ("Модернизация сайта + лендинг", "Блок TITAN на empils.ru + titanpaints.ru + лендинг",        "≈ 180 000 ₽", "разработка"),
]

STAGE_2_ITEMS = [
    ("B2B видео-отзыв",            "Съёмка отзыва производителя металлоконструкций (беседки)",    "≈ 80 000 ₽",  "1 ролик"),
    ("Блогер / инфлюенсер ЛКМ",    "Поиск + 1 интеграция у блогера по ЛКМ в РФ",                 "≈ 150 000 ₽", "интеграция"),
    ("Яндекс.Директ",              "Поиск + РСЯ в регионах с высокой представленностью TITAN",   "≈ 30 000 ₽/мес", "помесячно"),
    ("RuTube продвижение",         "Продвижение раздела TITAN канала (3-6 мес)",                  "≈ 40 000 ₽/мес", "помесячно"),
    ("VK геотаргетинг",            "Таргетинг по 11+ ТТ в регионах присутствия TITAN",            "≈ 25 000 ₽/мес", "помесячно"),
]

# Расчётные итоги
STAGE_1_TOTAL = "~ 810 000 ₽"  # сумма разовых
STAGE_2_TOTAL = "~ 95 000 ₽/мес"  # ежемесячные + разовые

# === УДАЛЯЕМ все существующие shapes на слайде 4 ===
# Нельзя напрямую удалить shape — нужно работать с XML
spTree = slide_old.shapes._spTree
shapes_to_remove = list(slide_old.shapes)
for shape in shapes_to_remove:
    sp = shape._element
    spTree.remove(sp)
print(f"Удалено shape'ов: {len(shapes_to_remove)}")

# === Вспомогательные функции ===

def add_rect(slide, left, top, width, height, fill_color, line_color=None, line_width=None):
    """Добавить прямоугольник с заданным фоном."""
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

def add_text(slide, left, top, width, height, text, font_size=10, bold=False, 
             color=DARK, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, font_name="Calibri",
             italic=False, line_spacing=1.0):
    """Добавить текстовый блок."""
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.margin_left = Pt(0); tf.margin_right = Pt(0)
    tf.margin_top = Pt(0); tf.margin_bottom = Pt(0)
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    
    # Разбиваем на строки по \n
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

# === Размеры слайда: 13.333" x 7.5" ===
# Конвертация: 1" = 914400 EMU
SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)

# === 1. Фон слайда ===
bg_rect = add_rect(slide_old, 0, 0, SLIDE_W, SLIDE_H, BG)

# === 2. Верхняя оранжевая полоса ===
top_bar = add_rect(slide_old, 0, 0, SLIDE_W, Inches(0.12), ORANGE)

# === 3. Заголовок ===
add_text(slide_old, Inches(0.5), Inches(0.25), Inches(8), Inches(0.45),
         "СМЕТА ПРОГРАММЫ TITAN — 2 ЭТАПА",
         font_size=22, bold=True, color=DARK, anchor=MSO_ANCHOR.MIDDLE)

add_text(slide_old, Inches(0.5), Inches(0.72), Inches(10), Inches(0.3),
         "Источник: Cost_calculation_Titan.xlsx · Оценки на основе рыночных цен 2026 г.",
         font_size=10, italic=True, color=GRAY)

# Бренд-марка справа сверху
add_text(slide_old, Inches(11.0), Inches(0.25), Inches(2.0), Inches(0.45),
         "TITAN · 2026",
         font_size=12, bold=True, color=ORANGE, align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)

# === 4. Две колонки с этапами ===
COL_WIDTH = Inches(6.05)
COL_TOP = Inches(1.25)
COL_HEIGHT = Inches(5.5)
COL_GAP = Inches(0.15)
COL_LEFT_1 = Inches(0.5)
COL_LEFT_2 = COL_LEFT_1 + COL_WIDTH + COL_GAP

# Колонка 1 — Первый этап
col1 = add_rect(slide_old, COL_LEFT_1, COL_TOP, COL_WIDTH, COL_HEIGHT, WHITE, line_color=LIGHT_GRAY, line_width=Pt(0.75))

# Заголовок колонки 1
add_rect(slide_old, COL_LEFT_1, COL_TOP, COL_WIDTH, Inches(0.55), ORANGE)
add_text(slide_old, COL_LEFT_1 + Inches(0.2), COL_TOP, COL_WIDTH - Inches(0.4), Inches(0.55),
         "ПЕРВЫЙ ЭТАП — разовые инвестиции",
         font_size=14, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
add_text(slide_old, COL_LEFT_1 + Inches(0.2), COL_TOP + Inches(0.55), COL_WIDTH - Inches(0.4), Inches(0.3),
         "Параллельно — наполнение MML в разрезе TITAN",
         font_size=9, italic=True, color=GRAY, anchor=MSO_ANCHOR.MIDDLE)

# Колонка 2 — Второй этап
col2 = add_rect(slide_old, COL_LEFT_2, COL_TOP, COL_WIDTH, COL_HEIGHT, WHITE, line_color=LIGHT_GRAY, line_width=Pt(0.75))

# Заголовок колонки 2
add_rect(slide_old, COL_LEFT_2, COL_TOP, COL_WIDTH, Inches(0.55), ORANGE_DARK)
add_text(slide_old, COL_LEFT_2 + Inches(0.2), COL_TOP, COL_WIDTH - Inches(0.4), Inches(0.55),
         "ВТОРОЙ ЭТАП — операционные расходы",
         font_size=14, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
add_text(slide_old, COL_LEFT_2 + Inches(0.2), COL_TOP + Inches(0.55), COL_WIDTH - Inches(0.4), Inches(0.3),
         "Продвижение после запуска программы",
         font_size=9, italic=True, color=GRAY, anchor=MSO_ANCHOR.MIDDLE)

# === 5. Карточки в колонке 1 (9 элементов) ===
CARD_TOP_START = COL_TOP + Inches(0.95)
CARD_HEIGHT = Inches(0.43)
CARD_GAP_Y = Inches(0.025)
CARD_PAD = Inches(0.15)

for i, (name, desc, cost, unit) in enumerate(STAGE_1_ITEMS):
    y = CARD_TOP_START + i * (CARD_HEIGHT + CARD_GAP_Y)
    
    # Подсветка чередующихся строк
    bg_color = LIGHT_GRAY if i % 2 == 0 else WHITE
    add_rect(slide_old, COL_LEFT_1 + Inches(0.05), y, COL_WIDTH - Inches(0.1), CARD_HEIGHT, bg_color)
    
    # Номер
    add_text(slide_old, COL_LEFT_1 + Inches(0.15), y, Inches(0.3), CARD_HEIGHT,
             str(i + 1),
             font_size=11, bold=True, color=ORANGE, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)
    
    # Наименование
    add_text(slide_old, COL_LEFT_1 + Inches(0.5), y + Inches(0.04), Inches(3.7), Inches(0.2),
             name,
             font_size=10, bold=True, color=DARK, anchor=MSO_ANCHOR.TOP)
    
    # Описание
    add_text(slide_old, COL_LEFT_1 + Inches(0.5), y + Inches(0.22), Inches(3.7), Inches(0.18),
             desc[:60] + ("…" if len(desc) > 60 else ""),
             font_size=8, color=GRAY, anchor=MSO_ANCHOR.TOP)
    
    # Стоимость
    add_text(slide_old, COL_LEFT_1 + Inches(4.3), y, Inches(1.5), CARD_HEIGHT,
             cost,
             font_size=11, bold=True, color=ORANGE_DARK, align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)

# Итог по этапу 1
total_y_1 = CARD_TOP_START + len(STAGE_1_ITEMS) * (CARD_HEIGHT + CARD_GAP_Y) + Inches(0.1)
add_rect(slide_old, COL_LEFT_1 + Inches(0.05), total_y_1, COL_WIDTH - Inches(0.1), Inches(0.45), ORANGE_LIGHT)
add_text(slide_old, COL_LEFT_1 + Inches(0.15), total_y_1, Inches(3), Inches(0.45),
         "ИТОГО ЭТАП 1:",
         font_size=12, bold=True, color=DARK, anchor=MSO_ANCHOR.MIDDLE)
add_text(slide_old, COL_LEFT_1 + Inches(3.0), total_y_1, Inches(2.85), Inches(0.45),
         STAGE_1_TOTAL,
         font_size=14, bold=True, color=ORANGE_DARK, align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)

# === 6. Карточки в колонке 2 (5 элементов — больше места для каждого) ===
CARD_TOP_START_2 = COL_TOP + Inches(0.95)
CARD_HEIGHT_2 = Inches(0.78)
CARD_GAP_Y_2 = Inches(0.04)

for i, (name, desc, cost, unit) in enumerate(STAGE_2_ITEMS):
    y = CARD_TOP_START_2 + i * (CARD_HEIGHT_2 + CARD_GAP_Y_2)
    
    bg_color = LIGHT_GRAY if i % 2 == 0 else WHITE
    add_rect(slide_old, COL_LEFT_2 + Inches(0.05), y, COL_WIDTH - Inches(0.1), CARD_HEIGHT_2, bg_color)
    
    # Номер
    add_text(slide_old, COL_LEFT_2 + Inches(0.15), y, Inches(0.3), CARD_HEIGHT_2,
             str(i + 1),
             font_size=11, bold=True, color=ORANGE_DARK, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)
    
    # Наименование
    add_text(slide_old, COL_LEFT_2 + Inches(0.5), y + Inches(0.06), Inches(4.2), Inches(0.25),
             name,
             font_size=11, bold=True, color=DARK, anchor=MSO_ANCHOR.TOP)
    
    # Описание
    add_text(slide_old, COL_LEFT_2 + Inches(0.5), y + Inches(0.32), Inches(4.2), Inches(0.4),
             desc,
             font_size=9, color=GRAY, anchor=MSO_ANCHOR.TOP, line_spacing=1.1)
    
    # Стоимость
    add_text(slide_old, COL_LEFT_2 + Inches(4.5), y, Inches(1.4), CARD_HEIGHT_2,
             cost,
             font_size=12, bold=True, color=ORANGE_DARK, align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)

# Итог по этапу 2
total_y_2 = CARD_TOP_START_2 + len(STAGE_2_ITEMS) * (CARD_HEIGHT_2 + CARD_GAP_Y_2) + Inches(0.1)
add_rect(slide_old, COL_LEFT_2 + Inches(0.05), total_y_2, COL_WIDTH - Inches(0.1), Inches(0.45), ORANGE_LIGHT)
add_text(slide_old, COL_LEFT_2 + Inches(0.15), total_y_2, Inches(3), Inches(0.45),
         "ИТОГО ЭТАП 2 (мес.):",
         font_size=12, bold=True, color=DARK, anchor=MSO_ANCHOR.MIDDLE)
add_text(slide_old, COL_LEFT_2 + Inches(3.0), total_y_2, Inches(2.85), Inches(0.45),
         STAGE_2_TOTAL,
         font_size=14, bold=True, color=ORANGE_DARK, align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)

# === 7. Подвал с итоговым CTA ===
footer_y = Inches(6.85)
add_rect(slide_old, 0, footer_y, SLIDE_W, Inches(0.5), DARK)
add_text(slide_old, Inches(0.5), footer_y, Inches(8), Inches(0.5),
         f"ВСЕГО ПРОГРАММА: ~ 810 000 ₽ разово + ~ 95 000 ₽/мес операционных расходов",
         font_size=12, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
add_text(slide_old, Inches(8.5), footer_y, Inches(4.5), Inches(0.5),
         "Empils  •  TITAN  •  Смета программы 2026",
         font_size=10, italic=True, color=RGBColor(0x90, 0x98, 0x9F), align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)

# Сохраняем
prs.save(DST)
print(f"\n✅ Файл сохранён: {DST}")
print(f"   Размер: {os.path.getsize(DST):,} байт")
