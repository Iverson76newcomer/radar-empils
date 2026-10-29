#!/usr/bin/env python3
"""
PPT-слайд с 3 главными цифрами программы продвижения TITAN.
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Cm, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn

OUT = '/home/z/my-project/download/TITAN_3_цифры_программы.pptx'

# Цвета
COL_PRIMARY = RGBColor(0x1E, 0x3A, 0x5F)
COL_ACCENT = RGBColor(0xF4, 0x79, 0x20)
COL_RED = RGBColor(0xC6, 0x28, 0x28)
COL_GREEN = RGBColor(0x16, 0xA3, 0x4A)
COL_BODY = RGBColor(0x1A, 0x1F, 0x36)
COL_GRAY = RGBColor(0x6B, 0x72, 0x80)
COL_WHITE = RGBColor(0xFF, 0xFF, 0xFF)
COL_LIGHT_GREEN = RGBColor(0xDC, 0xFC, 0xE7)
COL_LIGHT_RED = RGBColor(0xFE, 0xE2, 0xE2)
COL_LIGHT_YELLOW = RGBColor(0xFE, 0xF9, 0xC3)

prs = Presentation()
prs.slide_width = Cm(33.867)  # 16:9 widescreen
prs.slide_height = Cm(19.05)

slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank

# Заголовок
title_box = slide.shapes.add_textbox(Cm(1.5), Cm(0.8), Cm(30), Cm(1.5))
tf = title_box.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
p.alignment = PP_ALIGN.LEFT
r = p.add_run()
r.text = 'TITAN — программа продвижения: 3 цифры, которые меняют повестку'
r.font.name = 'Arial'
r.font.size = Pt(28)
r.font.bold = True
r.font.color.rgb = COL_PRIMARY

# Подзаголовок
sub_box = slide.shapes.add_textbox(Cm(1.5), Cm(2.3), Cm(30), Cm(1.2))
tf = sub_box.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
p.alignment = PP_ALIGN.LEFT
r = p.add_run()
r.text = 'Полугодие 2025 vs 2026 (без СтройТерминала — не предоставлял статистику в 2025)'
r.font.name = 'Arial'
r.font.size = Pt(13)
r.font.color.rgb = COL_GRAY
r.font.italic = True

# === 3 КРУПНЫЕ КАРТОЧКИ С ЦИФРАМИ ===
# Карточка 1: АКБ +53% (зелёная)
card_y = Cm(4.0)
card_h = Cm(6.5)
card_w = Cm(10.0)

cards = [
    {
        'x': Cm(1.5),
        'label': 'АКБ розничных клиентов',
        'v25': '6 096',
        'v26': '5 181',
        'delta': '−15%',
        'comment': 'Розница СЖИМАЕТСЯ — рост был иллюзией СтройТерминала',
        'bg': COL_LIGHT_RED,
        'delta_color': COL_RED,
        'border_color': COL_RED,
    },
    {
        'x': Cm(12.0),
        'label': 'РТТ, продающие TITAN',
        'v25': '464',
        'v26': '436',
        'delta': '−6%',
        'comment': 'Розничная база сокращается',
        'bg': COL_LIGHT_RED,
        'delta_color': COL_RED,
        'border_color': COL_RED,
    },
    {
        'x': Cm(22.5),
        'label': 'Кг TITAN (вторичка)',
        'v25': '30 589',
        'v26': '20 046',
        'delta': '−35%',
        'comment': 'Берут реже и меньше',
        'bg': COL_LIGHT_RED,
        'delta_color': COL_RED,
        'border_color': COL_RED,
    },
]

for card in cards:
    # Карточка-плашка
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, card['x'], card_y, card_w, card_h)
    shape.fill.solid()
    shape.fill.fore_color.rgb = card['bg']
    shape.line.color.rgb = card['border_color']
    shape.line.width = Pt(2)
    shape.shadow.inherit = False
    
    # Лейбл
    lbl_box = slide.shapes.add_textbox(card['x'] + Cm(0.5), card_y + Cm(0.3), card_w - Cm(1), Cm(0.8))
    tf = lbl_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = card['label']
    r.font.name = 'Arial'
    r.font.size = Pt(13)
    r.font.bold = True
    r.font.color.rgb = COL_PRIMARY
    
    # 2025
    v25_box = slide.shapes.add_textbox(card['x'] + Cm(0.5), card_y + Cm(1.2), card_w - Cm(1), Cm(1))
    tf = v25_box.text_frame
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = card['v25']
    r.font.name = 'Arial'
    r.font.size = Pt(20)
    r.font.color.rgb = COL_GRAY
    
    # Стрелка
    arr_box = slide.shapes.add_textbox(card['x'] + Cm(0.5), card_y + Cm(2.1), card_w - Cm(1), Cm(0.6))
    tf = arr_box.text_frame
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = '→'
    r.font.name = 'Arial'
    r.font.size = Pt(16)
    r.font.color.rgb = COL_GRAY
    
    # 2026 (крупно)
    v26_box = slide.shapes.add_textbox(card['x'] + Cm(0.5), card_y + Cm(2.7), card_w - Cm(1), Cm(1.5))
    tf = v26_box.text_frame
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = card['v26']
    r.font.name = 'Arial'
    r.font.size = Pt(36)
    r.font.bold = True
    r.font.color.rgb = COL_PRIMARY
    
    # Дельта (крупно, цветная)
    delta_box = slide.shapes.add_textbox(card['x'] + Cm(0.5), card_y + Cm(4.3), card_w - Cm(1), Cm(1.2))
    tf = delta_box.text_frame
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = card['delta']
    r.font.name = 'Arial'
    r.font.size = Pt(28)
    r.font.bold = True
    r.font.color.rgb = card['delta_color']
    
    # Комментарий
    cmt_box = slide.shapes.add_textbox(card['x'] + Cm(0.5), card_y + Cm(5.5), card_w - Cm(1), Cm(0.9))
    tf = cmt_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = card['comment']
    r.font.name = 'Arial'
    r.font.size = Pt(11)
    r.font.italic = True
    r.font.color.rgb = COL_BODY

# === ВЫВОД ===
concl_y = Cm(11.3)
concl_box = slide.shapes.add_textbox(Cm(1.5), concl_y, Cm(31), Cm(1.2))
tf = concl_box.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
p.alignment = PP_ALIGN.LEFT
r = p.add_run()
r.text = 'ВЫВОД: '
r.font.name = 'Arial'
r.font.size = Pt(16)
r.font.bold = True
r.font.color.rgb = COL_RED
r2 = p.add_run()
r2.text = 'ВСЕ 3 метрики падают. Рост был иллюзией СтройТерминала. Цель программы — не «закрепить успех», а ОСТАНОВИТЬ ПАДЕНИЕ.'
r2.font.name = 'Arial'
r2.font.size = Pt(14)
r2.font.color.rgb = COL_BODY

# === 4 ПРИОРИТЕТА ===
prio_y = Cm(13.0)
prio_box = slide.shapes.add_textbox(Cm(1.5), prio_y, Cm(31), Cm(0.6))
tf = prio_box.text_frame
p = tf.paragraphs[0]
p.alignment = PP_ALIGN.LEFT
r = p.add_run()
r.text = 'ЧТО ДЕЛАТЬ — 4 ПРИОРИТЕТА ПРОГРАММЫ'
r.font.name = 'Arial'
r.font.size = Pt(14)
r.font.bold = True
r.font.color.rgb = COL_PRIMARY

prio_y2 = Cm(13.7)
prio_h = Cm(1.1)
prio_w = Cm(7.5)
prios = [
    ('1', 'Спасти Технологии РС (Оберон)', '−86% кг, 9 клиентов ушло'),
    ('2', 'Увеличить MML', '12 дистр = 98% кг, расширить с 30 до 44 SKU'),
    ('3', 'Вернуть 30 клиентов', '10 000 кг потенциал, таргет-акция'),
    ('4', 'Масштабировать АЮСС + Уралстрой', '+867% / +55% — 2 точки роста'),
]
for i, (num, title, desc) in enumerate(prios):
    x = Cm(1.5 + i * 8.0)
    # Номер
    num_shape = slide.shapes.add_shape(MSO_SHAPE.OVAL, x, prio_y2, Cm(0.9), Cm(0.9))
    num_shape.fill.solid()
    num_shape.fill.fore_color.rgb = COL_ACCENT
    num_shape.line.fill.background()
    num_shape.shadow.inherit = False
    tf = num_shape.text_frame
    tf.word_wrap = False
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = num
    r.font.name = 'Arial'
    r.font.size = Pt(16)
    r.font.bold = True
    r.font.color.rgb = COL_WHITE
    # Заголовок приоритета
    t_box = slide.shapes.add_textbox(x + Cm(1.1), prio_y2, prio_w - Cm(1.1), Cm(0.6))
    tf = t_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    r = p.add_run()
    r.text = title
    r.font.name = 'Arial'
    r.font.size = Pt(12)
    r.font.bold = True
    r.font.color.rgb = COL_PRIMARY
    # Описание
    d_box = slide.shapes.add_textbox(x + Cm(1.1), prio_y2 + Cm(0.55), prio_w - Cm(1.1), Cm(0.7))
    tf = d_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    r = p.add_run()
    r.text = desc
    r.font.name = 'Arial'
    r.font.size = Pt(9)
    r.font.color.rgb = COL_BODY

# === Подвал ===
footer_y = Cm(17.8)
footer_box = slide.shapes.add_textbox(Cm(1.5), footer_y, Cm(31), Cm(0.8))
tf = footer_box.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
p.alignment = PP_ALIGN.LEFT
r = p.add_run()
r.text = 'Источник: «Титан итого.xlsx» (3 вкладки). Исключён СтройТерминал: в 2025 не предоставлял статистику, в 2026 дал 4 161 АКБ — искажал динамику. Период: 1 полугодие 2025 vs 2026.'
r.font.name = 'Arial'
r.font.size = Pt(8)
r.font.color.rgb = COL_GRAY
r.font.italic = True

prs.save(OUT)
print(f'OK: {OUT}')
import os
print(f'Размер: {os.path.getsize(OUT) / 1024:.1f} KB')
