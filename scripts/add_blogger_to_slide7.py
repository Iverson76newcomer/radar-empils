#!/usr/bin/env python3
"""
Добавить пункт 11 «Работа с блогером / инфлюенсером» в слайд 7
презентации Titan -Дорожная_карта_fin_16.07.2026.pptx

Текущая структура левой колонки:
- Заголовок "ЭТАП 3 — ТРЕТИЧНЫЕ ПРОДАЖИ" (top: 0.08")
- Подзаголовок "Розница → Покупатель" (top: 0.08", справа)
- "ИНСТРУМЕНТЫ МАРКЕТИНГА" (top: 0.95")
- 10 пунктов-карточек (top: 1.45" - 6.65"), каждая высотой 0.52"

Свободное место: 6.65" → 7.05" (футер) = 0.4"

Решение: добавить пункт 11 как новую карточку высотой 0.32"
сразу после пункта 10 (top: 6.68"), уменьшив высоту и отступы.
"""

import os, shutil
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

SRC = "/home/z/my-project/upload/Titan -Дорожная_карта_fin_16.07.2026.pptx"
DST = "/home/z/my-project/download/Titan_Дорожная_карта_с_блогером_v1.pptx"
shutil.copy(SRC, DST)

prs = Presentation(DST)
slide = prs.slides[6]  # slide 7 (0-indexed)

# === Цвета (такие же как на слайде) ===
ORANGE = RGBColor(0xF4, 0x76, 0x29)
ORANGE_DARK = RGBColor(0xD5, 0x5E, 0x0A)
ORANGE_LIGHT = RGBColor(0xFB, 0xE6, 0xD5)
DARK = RGBColor(0x31, 0x31, 0x31)
GRAY = RGBColor(0x6B, 0x6B, 0x6B)
LIGHT_GRAY = RGBColor(0xF5, 0xF5, 0xF5)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
BG = RGBColor(0xFA, 0xFA, 0xFA)

# Фон карточки (как у других)
CARD_BG = RGBColor(0xFA, 0xFA, 0xFA)  # лёгкий серый

# === Координаты новой карточки (после пункта 10) ===
# Пункт 10 заканчивается на top=6.65" (height=0.52")
# Футер на 7.05"
# Делаем карточку высотой 0.32" с top=6.68"

NEW_TOP = Inches(6.68)
NEW_HEIGHT = Inches(0.32)
LEFT = Inches(0.40)
WIDTH = Inches(8.00)

# === Создаём shapes для нового пункта 11 ===

# 1. Фон карточки (прямоугольник)
card = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, LEFT, NEW_TOP, WIDTH, NEW_HEIGHT)
card.fill.solid()
card.fill.fore_color.rgb = WHITE
card.line.color.rgb = LIGHT_GRAY
card.line.width = Pt(0.75)
card.shadow.inherit = False

# 2. Кружок-номер (как у других пунктов: 0.32 x 0.32 на позиции 0.50, NEW_TOP+0.05)
num_circle = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(0.50), NEW_TOP + Inches(0.05), Inches(0.22), Inches(0.22))
num_circle.fill.solid()
num_circle.fill.fore_color.rgb = ORANGE
num_circle.line.fill.background()
num_circle.shadow.inherit = False

# Текст номера в кружке
tf = num_circle.text_frame
tf.margin_left = Pt(0); tf.margin_right = Pt(0)
tf.margin_top = Pt(0); tf.margin_bottom = Pt(0)
tf.vertical_anchor = MSO_ANCHOR.MIDDLE
p = tf.paragraphs[0]
p.alignment = PP_ALIGN.CENTER
r = p.add_run()
r.text = "11"
r.font.name = "Calibri"
r.font.size = Pt(9)
r.font.bold = True
r.font.color.rgb = WHITE

# 3. Текст карточки: заголовок + описание
text_box = slide.shapes.add_textbox(Inches(0.80), NEW_TOP, Inches(7.50), NEW_HEIGHT)
tf = text_box.text_frame
tf.margin_left = Pt(0); tf.margin_right = Pt(0)
tf.margin_top = Pt(0); tf.margin_bottom = Pt(0)
tf.vertical_anchor = MSO_ANCHOR.MIDDLE
tf.word_wrap = True

# Заголовок (первый параграф)
p1 = tf.paragraphs[0]
p1.alignment = PP_ALIGN.LEFT
p1.line_spacing = 1.0
r1 = p1.add_run()
r1.text = "Работа с блогером / инфлюенсером — "
r1.font.name = "Calibri"
r1.font.size = Pt(10)
r1.font.bold = True
r1.font.color.rgb = DARK

# Описание (второй параграф)
p2 = tf.add_paragraph()
p2.alignment = PP_ALIGN.LEFT
p2.line_spacing = 1.0
r2 = p2.add_run()
r2.text = "Сотрудничество с блогером в сфере ЛКМ для продвижения TITAN"
r2.font.name = "Calibri"
r2.font.size = Pt(9)
r2.font.italic = True
r2.font.color.rgb = GRAY

# Сохраняем
prs.save(DST)
print(f"✅ Файл сохранён: {DST}")
print(f"   Размер: {os.path.getsize(DST):,} байт")
