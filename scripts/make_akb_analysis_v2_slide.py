#!/usr/bin/env python3
"""
Обновить слайд 4 (АКБ-анализ) в презентации Titan_Дорожная_карта_пилоты_v1.pptx
на основе пересчитанных данных из NEW TITAN_анализ_дистрибьюторов.xlsx.

Ключевые изменения:
- АКБ: 6006 → 5255 (−12.5%)
- РТТ: 465 → 469 (+0.86%) — стабилизировались!
- Вторичка: 35575 → 32719 (−8%)
- Первичка: 104812 → 119217 (+13.7%) — сильный рост
"""

import os, shutil
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

SRC = "/home/z/my-project/download/Titan_Дорожная_карта_пилоты_v1.pptx"
DST = "/home/z/my-project/download/Titan_Дорожная_карта_АКБ_v2.pptx"
shutil.copy(SRC, DST)

prs = Presentation(DST)
slide = prs.slides[3]  # slide 4 (АКБ-анализ, индекс 3)

# === Цвета фирменные ===
ORANGE = RGBColor(0xF4, 0x76, 0x29)
ORANGE_DARK = RGBColor(0xD5, 0x5E, 0x0A)
ORANGE_LIGHT = RGBColor(0xFB, 0xE6, 0xD5)
DARK = RGBColor(0x31, 0x31, 0x31)
GRAY = RGBColor(0x6B, 0x6B, 0x6B)
LIGHT_GRAY = RGBColor(0xF5, 0xF5, 0xF5)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
BG = RGBColor(0xFA, 0xFA, 0xFA)
RED = RGBColor(0xC6, 0x28, 0x28)
RED_LIGHT = RGBColor(0xFF, 0xEB, 0xEE)
GREEN = RGBColor(0x2E, 0x7D, 0x32)
GREEN_LIGHT = RGBColor(0xE8, 0xF5, 0xE9)
AMBER = RGBColor(0xF5, 0x7C, 0x00)
AMBER_LIGHT = RGBColor(0xFF, 0xF3, 0xE0)

# === Обновлённые данные ===
METRICS = [
    # (Название, 2025, 2026, изменение %, комментарий, тип тренда)
    ("АКБ розничных клиентов", "6 006", "5 255", "−12.5%", "Розница СЖИМАЕТСЯ", "danger"),
    ("РТТ, продающие TITAN",  "465",   "469",   "+0.9%",  "Розничная база стабильна", "positive"),
    ("Кг TITAN (вторичка)",   "35 575","32 719","−8%",    "Берут реже и меньше — проблема глубины", "amber"),
    ("Кг TITAN (первичка)",   "104 812","119 217","+13.7%","Рост отгрузок", "positive"),
]

# === Удаляем все shapes на слайде 4 ===
spTree = slide.shapes._spTree
shapes_to_remove = list(slide.shapes)
for shape in shapes_to_remove:
    sp = shape._element
    spTree.remove(sp)
print(f"Удалено shape'ов: {len(shapes_to_remove)}")

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
        shape.adjustments[0] = 0.06
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
         "TITAN — Аналитика 1 полугодия: 2025 vs 2026",
         font_size=22, bold=True, color=DARK, anchor=MSO_ANCHOR.MIDDLE)

add_text(slide, Inches(0.5), Inches(0.70), Inches(12), Inches(0.3),
         "Без «СтройТерминала» и «Технологии, ресурсы строительства» (не предоставляли статистику в 2025-2026 гг). Источник: NEW TITAN_анализ_дистрибьюторов.xlsx",
         font_size=9, italic=True, color=GRAY)

# Бренд-марка справа сверху
add_text(slide, Inches(11.0), Inches(0.22), Inches(2.0), Inches(0.45),
         "TITAN · 2026",
         font_size=12, bold=True, color=ORANGE, align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)

# === 4. 4 KPI-карточки в ряд ===
CARD_TOP = Inches(1.35)
CARD_HEIGHT = Inches(2.45)
CARD_WIDTH = Inches(3.05)
CARD_GAP = Inches(0.12)
CARD_LEFT_START = Inches(0.5)

for i, (name, val_2025, val_2026, change, comment, trend) in enumerate(METRICS):
    left = CARD_LEFT_START + i * (CARD_WIDTH + CARD_GAP)
    
    # Цвет карточки по тренду
    if trend == "danger":
        accent_color = RED
        accent_bg = RED_LIGHT
        change_color = RED
    elif trend == "positive":
        accent_color = GREEN
        accent_bg = GREEN_LIGHT
        change_color = GREEN
    else:  # amber
        accent_color = AMBER
        accent_bg = AMBER_LIGHT
        change_color = AMBER
    
    # Фон карточки
    add_round_rect(slide, left, CARD_TOP, CARD_WIDTH, CARD_HEIGHT, WHITE,
                   line_color=LIGHT_GRAY, line_width=Pt(0.75))
    
    # Верхняя цветная плашка-заголовок
    add_rect(slide, left, CARD_TOP, CARD_WIDTH, Inches(0.5), accent_color)
    add_text(slide, left + Inches(0.15), CARD_TOP, CARD_WIDTH - Inches(0.3), Inches(0.5),
             name,
             font_size=11, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.0)
    
    # Значения 2025 → 2026
    val_top = CARD_TOP + Inches(0.65)
    
    # 2025
    add_text(slide, left + Inches(0.15), val_top, Inches(1.2), Inches(0.25),
             "2025", font_size=9, color=GRAY, anchor=MSO_ANCHOR.MIDDLE)
    add_text(slide, left + Inches(0.15), val_top + Inches(0.25), Inches(1.4), Inches(0.45),
             val_2025, font_size=22, bold=True, color=GRAY, anchor=MSO_ANCHOR.MIDDLE)
    
    # Стрелка
    add_text(slide, left + Inches(1.35), val_top + Inches(0.25), Inches(0.4), Inches(0.45),
             "→", font_size=24, bold=True, color=GRAY, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    
    # 2026
    add_text(slide, left + Inches(1.65), val_top, Inches(1.2), Inches(0.25),
             "2026", font_size=9, color=accent_color, anchor=MSO_ANCHOR.MIDDLE)
    add_text(slide, left + Inches(1.65), val_top + Inches(0.25), Inches(1.4), Inches(0.45),
             val_2026, font_size=22, bold=True, color=accent_color, anchor=MSO_ANCHOR.MIDDLE)
    
    # Изменение в плашке
    change_top = CARD_TOP + Inches(1.45)
    add_round_rect(slide, left + Inches(0.15), change_top, Inches(1.1), Inches(0.4), accent_bg)
    add_text(slide, left + Inches(0.15), change_top, Inches(1.1), Inches(0.4),
             change, font_size=14, bold=True, color=change_color,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    
    # Комментарий
    add_text(slide, left + Inches(1.35), change_top, CARD_WIDTH - Inches(1.5), Inches(0.5),
             comment, font_size=9, italic=True, color=DARK,
             anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.1)

# === 5. Блок с парадоксом сезона (обновлённый) ===
INSIGHT_TOP = Inches(4.10)
INSIGHT_HEIGHT = Inches(2.0)
INSIGHT_LEFT = Inches(0.5)
INSIGHT_WIDTH = Inches(12.33)

add_round_rect(slide, INSIGHT_LEFT, INSIGHT_TOP, INSIGHT_WIDTH, INSIGHT_HEIGHT, ORANGE_LIGHT,
               line_color=ORANGE, line_width=Pt(1))

add_text(slide, INSIGHT_LEFT + Inches(0.3), INSIGHT_TOP + Inches(0.15), Inches(11), Inches(0.4),
         "⚠ ПАРАДОКС СЕЗОНА",
         font_size=14, bold=True, color=ORANGE_DARK, anchor=MSO_ANCHOR.MIDDLE)

add_text(slide, INSIGHT_LEFT + Inches(0.3), INSIGHT_TOP + Inches(0.55), Inches(11.5), Inches(0.5),
         "первичка растёт (+13.7%), вторичка падает (−8%). Дистрибьюторы закупают TITAN на склад, но не продают в розницу.",
         font_size=12, bold=True, color=DARK, anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.2)

# Две колонки — тревожные + позитивные сигналы (обновлённые)
SIGNAL_TOP = INSIGHT_TOP + Inches(1.10)
SIGNAL_WIDTH = Inches(5.85)

# Левая — тревожные сигналы
add_text(slide, INSIGHT_LEFT + Inches(0.3), SIGNAL_TOP, SIGNAL_WIDTH, Inches(0.3),
         "🚨 Тревожные сигналы", font_size=11, bold=True, color=RED)

add_text(slide, INSIGHT_LEFT + Inches(0.3), SIGNAL_TOP + Inches(0.32), SIGNAL_WIDTH, Inches(0.6),
         "• АКБ −12.5% (розница сжимается)\n• Вторичка −8% — берут реже, чем закупают\n• 47 из 109 договоров (44%) — «спящие», без закупок",
         font_size=10, color=DARK, line_spacing=1.3)

# Правая — позитивные сигналы
SIGNAL_LEFT_2 = INSIGHT_LEFT + Inches(6.4)
add_text(slide, SIGNAL_LEFT_2, SIGNAL_TOP, SIGNAL_WIDTH, Inches(0.3),
         "✅ Позитивные сигналы", font_size=11, bold=True, color=GREEN)

add_text(slide, SIGNAL_LEFT_2, SIGNAL_TOP + Inches(0.32), SIGNAL_WIDTH, Inches(0.6),
         "• РТТ +0.9% (465→469) — розничная база СТАБИЛИЗИРОВАЛАСЬ\n• Первичка +13.7% (104 812→119 217 кг) — сильный рост\n• Уралстройсервис +55%, Держава Трейд (KZ) +11%",
         font_size=10, color=DARK, line_spacing=1.3)

# === 6. Подвал с итоговым выводом (обновлённый) ===
FOOTER_TOP = Inches(6.30)
add_rect(slide, 0, FOOTER_TOP, SLIDE_W, Inches(0.5), DARK)
add_text(slide, Inches(0.5), FOOTER_TOP, Inches(9.5), Inches(0.5),
         "ВЫВОД: тренд улучшается (−35% → −8% по вторичке, +1.8% → +13.7% по первичке). Нужна программа активации розницы и работы с «спящими» договорами",
         font_size=11, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
add_text(slide, Inches(10.0), FOOTER_TOP, Inches(3.0), Inches(0.5),
         "Empils  •  TITAN  •  2026",
         font_size=10, italic=True, color=RGBColor(0x90, 0x98, 0x9F), align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)

# === 7. Нижняя оранжевая полоска ===
add_rect(slide, 0, Inches(7.30), SLIDE_W, Inches(0.20), ORANGE)

# Сохраняем
prs.save(DST)
print(f"\n✅ Файл сохранён: {DST}")
print(f"   Размер: {os.path.getsize(DST):,} байт")
