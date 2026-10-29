#!/usr/bin/env python3
"""
PPT-слайд: TITAN — аналитика 1 полугодия 2025 vs 2026 (по последней версии).
Источник: TITAN_анализ_дистрибьюторов_для_программы_v2.xlsx
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Cm, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
import os

OUT = '/home/z/my-project/download/TITAN_аналитика_1полугодие_2025vs2026.pptx'

# Цвета
COL_PRIMARY = RGBColor(0x1E, 0x3A, 0x5F)
COL_ACCENT = RGBColor(0xF4, 0x79, 0x20)
COL_RED = RGBColor(0xC6, 0x28, 0x28)
COL_GREEN = RGBColor(0x16, 0xA3, 0x4A)
COL_BLUE = RGBColor(0x1E, 0x40, 0xAF)
COL_ORANGE = RGBColor(0xF4, 0x79, 0x20)
COL_BODY = RGBColor(0x1A, 0x1F, 0x36)
COL_GRAY = RGBColor(0x6B, 0x72, 0x80)
COL_WHITE = RGBColor(0xFF, 0xFF, 0xFF)
COL_LIGHT_RED = RGBColor(0xFE, 0xE2, 0xE2)
COL_LIGHT_GREEN = RGBColor(0xDC, 0xFC, 0xE7)
COL_LIGHT_BLUE = RGBColor(0xDB, 0xEA, 0xFE)
COL_LIGHT_ORANGE = RGBColor(0xFF, 0xED, 0xD5)
COL_LIGHT_GRAY = RGBColor(0xF3, 0xF4, 0xF6)
COL_LIGHT_YELLOW = RGBColor(0xFE, 0xF9, 0xC3)

prs = Presentation()
prs.slide_width = Cm(33.867)
prs.slide_height = Cm(19.05)

slide = prs.slides.add_slide(prs.slide_layouts[6])

# === ЗАГОЛОВОК ===
title_box = slide.shapes.add_textbox(Cm(1.0), Cm(0.5), Cm(31), Cm(1.2))
tf = title_box.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
p.alignment = PP_ALIGN.LEFT
r = p.add_run()
r.text = 'TITAN — аналитика 1 полугодия: 2025 vs 2026'
r.font.name = 'Arial'
r.font.size = Pt(26)
r.font.bold = True
r.font.color.rgb = COL_PRIMARY

sub_box = slide.shapes.add_textbox(Cm(1.0), Cm(1.8), Cm(31), Cm(0.6))
tf = sub_box.text_frame
p = tf.paragraphs[0]
p.alignment = PP_ALIGN.LEFT
r = p.add_run()
r.text = 'Без СтройТерминала (не предоставлял статистику в 2025). Источник: TITAN_анализ_дистрибьюторов_для_программы_v2.xlsx'
r.font.name = 'Arial'
r.font.size = Pt(10)
r.font.color.rgb = COL_GRAY
r.font.italic = True

# === 4 КАРТОЧКИ С ГЛАВНЫМИ ЦИФРАМИ ===
# 3 красные (вторичка падает) + 1 зелёная (первичка растёт)
card_y = Cm(2.7)
card_h = Cm(4.8)
card_w = Cm(7.7)
gap = Cm(0.3)

cards = [
    {
        'x': Cm(1.0),
        'label': 'АКБ розничных клиентов',
        'v25': '6 096', 'v26': '5 181',
        'delta': '−15%',
        'comment': 'Розница сжимается',
        'bg': COL_LIGHT_RED, 'delta_color': COL_RED, 'border': COL_RED,
    },
    {
        'x': Cm(1.0) + card_w + gap,
        'label': 'РТТ, продающие TITAN',
        'v25': '464', 'v26': '436',
        'delta': '−6%',
        'comment': 'База сокращается',
        'bg': COL_LIGHT_RED, 'delta_color': COL_RED, 'border': COL_RED,
    },
    {
        'x': Cm(1.0) + 2 * (card_w + gap),
        'label': 'Кг TITAN (вторичка)',
        'v25': '30 589', 'v26': '20 046',
        'delta': '−35%',
        'comment': 'Берут реже и меньше',
        'bg': COL_LIGHT_RED, 'delta_color': COL_RED, 'border': COL_RED,
    },
    {
        'x': Cm(1.0) + 3 * (card_w + gap),
        'label': 'Кг TITAN (первичка)',
        'v25': '90 504', 'v26': '92 094',
        'delta': '+1.8%',
        'comment': 'Завод растёт, но склады затовариваются',
        'bg': COL_LIGHT_GREEN, 'delta_color': COL_GREEN, 'border': COL_GREEN,
    },
]

for card in cards:
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, card['x'], card_y, card_w, card_h)
    shape.fill.solid()
    shape.fill.fore_color.rgb = card['bg']
    shape.line.color.rgb = card['border']
    shape.line.width = Pt(2)
    shape.shadow.inherit = False
    
    # Label
    lbl = slide.shapes.add_textbox(card['x'] + Cm(0.3), card_y + Cm(0.2), card_w - Cm(0.6), Cm(0.7))
    tf = lbl.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = card['label']
    r.font.name = 'Arial'; r.font.size = Pt(12); r.font.bold = True; r.font.color.rgb = COL_PRIMARY
    
    # 2025
    v25 = slide.shapes.add_textbox(card['x'] + Cm(0.3), card_y + Cm(1.0), card_w - Cm(0.6), Cm(0.8))
    tf = v25.text_frame
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = card['v25']
    r.font.name = 'Arial'; r.font.size = Pt(18); r.font.color.rgb = COL_GRAY
    
    # Arrow
    arr = slide.shapes.add_textbox(card['x'] + Cm(0.3), card_y + Cm(1.8), card_w - Cm(0.6), Cm(0.4))
    tf = arr.text_frame
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = '→'
    r.font.name = 'Arial'; r.font.size = Pt(14); r.font.color.rgb = COL_GRAY
    
    # 2026
    v26 = slide.shapes.add_textbox(card['x'] + Cm(0.3), card_y + Cm(2.2), card_w - Cm(0.6), Cm(1.2))
    tf = v26.text_frame
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = card['v26']
    r.font.name = 'Arial'; r.font.size = Pt(30); r.font.bold = True; r.font.color.rgb = COL_PRIMARY
    
    # Delta
    delta = slide.shapes.add_textbox(card['x'] + Cm(0.3), card_y + Cm(3.4), card_w - Cm(0.6), Cm(0.8))
    tf = delta.text_frame
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = card['delta']
    r.font.name = 'Arial'; r.font.size = Pt(22); r.font.bold = True; r.font.color.rgb = card['delta_color']
    
    # Comment
    cmt = slide.shapes.add_textbox(card['x'] + Cm(0.3), card_y + Cm(4.2), card_w - Cm(0.6), Cm(0.5))
    tf = cmt.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = card['comment']
    r.font.name = 'Arial'; r.font.size = Pt(9); r.font.italic = True; r.font.color.rgb = COL_BODY

# === ГЛАВНЫЙ ВЫВОД — парадокс ===
concl_y = Cm(8.0)
concl_box = slide.shapes.add_textbox(Cm(1.0), concl_y, Cm(31.5), Cm(1.4))
tf = concl_box.text_frame; tf.word_wrap = True
p = tf.paragraphs[0]; p.alignment = PP_ALIGN.LEFT
r = p.add_run()
r.text = '⚠ ПАРАДОКС СЕЗОНА: '
r.font.name = 'Arial'; r.font.size = Pt(15); r.font.bold = True; r.font.color.rgb = COL_ACCENT
r2 = p.add_run()
r2.text = 'первичка растёт (+1.8%), вторичка падает (−35%). Дистрибьюторы закупают TITAN на склад, но не продают в розницу. Если до конца сезона не распродадут — Q4 потеряем ещё больше.'
r2.font.name = 'Arial'; r2.font.size = Pt(13); r2.font.color.rgb = COL_BODY

# === 4 ГРУППЫ ДИСТРИБЬЮТОРОВ (по первичке) ===
groups_y = Cm(9.8)
grp_box = slide.shapes.add_textbox(Cm(1.0), groups_y, Cm(20), Cm(0.5))
tf = grp_box.text_frame
p = tf.paragraphs[0]; p.alignment = PP_ALIGN.LEFT
r = p.add_run()
r.text = 'Группы дистрибьюторов (первичка, кг)'
r.font.name = 'Arial'; r.font.size = Pt(12); r.font.bold = True; r.font.color.rgb = COL_PRIMARY

# Таблица-плашка 4 группы
groups_data = [
    ('🟢 Стабильные', '47', '74 749', '70 449', '−6%', COL_GREEN, COL_LIGHT_GREEN),
    ('🔵 Новые 2026', '8', '3 036', '1', 'new', COL_BLUE, COL_LIGHT_BLUE),
    ('🟠 Ушедшие', '7', '820', '0', '−100%', COL_ORANGE, COL_LIGHT_ORANGE),
    ('⚪ Без данных', '47', '14 936', '18 609', '+25%', COL_GRAY, COL_LIGHT_GRAY),
]
grp_card_w = Cm(4.9)
grp_y = groups_y + Cm(0.6)
for i, (name, count, kg25, kg26, delta, color, bg) in enumerate(groups_data):
    x = Cm(1.0) + i * (grp_card_w + Cm(0.1))
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, grp_y, grp_card_w, Cm(2.4))
    shape.fill.solid(); shape.fill.fore_color.rgb = bg
    shape.line.color.rgb = color; shape.line.width = Pt(1.5)
    shape.shadow.inherit = False
    
    # Name
    nm = slide.shapes.add_textbox(x + Cm(0.2), grp_y + Cm(0.15), grp_card_w - Cm(0.4), Cm(0.5))
    tf = nm.text_frame
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = name
    r.font.name = 'Arial'; r.font.size = Pt(10); r.font.bold = True; r.font.color.rgb = color
    
    # Count
    cnt = slide.shapes.add_textbox(x + Cm(0.2), grp_y + Cm(0.65), grp_card_w - Cm(0.4), Cm(0.4))
    tf = cnt.text_frame
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = f'{count} договоров'
    r.font.name = 'Arial'; r.font.size = Pt(9); r.font.color.rgb = COL_BODY
    
    # kg
    kg = slide.shapes.add_textbox(x + Cm(0.2), grp_y + Cm(1.1), grp_card_w - Cm(0.4), Cm(0.8))
    tf = kg.text_frame
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = f'{kg25} → {kg26}'
    r.font.name = 'Arial'; r.font.size = Pt(11); r.font.bold = True; r.font.color.rgb = COL_PRIMARY
    
    # Delta
    dlt = slide.shapes.add_textbox(x + Cm(0.2), grp_y + Cm(1.85), grp_card_w - Cm(0.4), Cm(0.5))
    tf = dlt.text_frame
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = delta
    r.font.name = 'Arial'; r.font.size = Pt(14); r.font.bold = True; r.font.color.rgb = color

# === ТРЕВОЖНЫЕ СИГНАЛЫ (справа вверху от групп) ===
alarm_y = Cm(9.8)
alarm_x = Cm(21.5)
alarm_box = slide.shapes.add_textbox(alarm_x, alarm_y, Cm(11.5), Cm(0.5))
tf = alarm_box.text_frame
p = tf.paragraphs[0]; p.alignment = PP_ALIGN.LEFT
r = p.add_run()
r.text = '🚨 Тревожные сигналы'
r.font.name = 'Arial'; r.font.size = Pt(12); r.font.bold = True; r.font.color.rgb = COL_RED

alarms = [
    'Все 3 метрики вторички падают: АКБ −15%, РТТ −6%, кг −35%',
    'Севастьянов (−52%) — крупнейшее падение, Бурдинский сократил закупки вдвое',
    '47 из 109 договоров (44%) — «спящие», без закупок',
]
for i, txt in enumerate(alarms):
    y = alarm_y + Cm(0.5) + i * Cm(0.65)
    b = slide.shapes.add_textbox(alarm_x, y, Cm(11.5), Cm(0.6))
    tf = b.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.LEFT
    r = p.add_run()
    r.text = f'• {txt}'
    r.font.name = 'Arial'; r.font.size = Pt(10); r.font.color.rgb = COL_BODY

# === ПОЗИТИВНЫЕ СИГНАЛЫ (справа ниже) ===
pos_y = Cm(12.5)
pos_box = slide.shapes.add_textbox(alarm_x, pos_y, Cm(11.5), Cm(0.5))
tf = pos_box.text_frame
p = tf.paragraphs[0]; p.alignment = PP_ALIGN.LEFT
r = p.add_run()
r.text = '✅ Позитивные сигналы'
r.font.name = 'Arial'; r.font.size = Pt(12); r.font.bold = True; r.font.color.rgb = COL_GREEN

positives = [
    'Уралстройсервис +55% (327→508 кг), РТТ ×2 (29→58)',
    'АЮСС-ВСТК +2-9%, привлёк 5 новых клиентов',
    'Держава Трейд (Казахстан) +11% = 7 065 кг (35% всей вторички)',
]
for i, txt in enumerate(positives):
    y = pos_y + Cm(0.5) + i * Cm(0.65)
    b = slide.shapes.add_textbox(alarm_x, y, Cm(11.5), Cm(0.6))
    tf = b.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.LEFT
    r = p.add_run()
    r.text = f'• {txt}'
    r.font.name = 'Arial'; r.font.size = Pt(10); r.font.color.rgb = COL_BODY

# === 4 ПРИОРИТЕТА — внизу ===
prio_y = Cm(14.7)
prio_title = slide.shapes.add_textbox(Cm(1.0), prio_y, Cm(31), Cm(0.5))
tf = prio_title.text_frame
p = tf.paragraphs[0]; p.alignment = PP_ALIGN.LEFT
r = p.add_run()
r.text = 'ЧТО ДЕЛАТЬ — 4 ПРИОРИТЕТА НА Q3 2026'
r.font.name = 'Arial'; r.font.size = Pt(12); r.font.bold = True; r.font.color.rgb = COL_PRIMARY

prio_y2 = Cm(15.3)
prios = [
    ('1', 'Аудит складов', '12 стабильных — запросить остатки до 25.08'),
    ('2', 'Севастьянов', 'Встреча с Бурдинским до 31.08'),
    ('3', '47 «спящих»', 'Активировать или закрыть до 10.09'),
    ('4', '30 потерянных клиентов', 'Таргет-акция возврата, ~10 000 кг'),
]
prio_w = Cm(7.7)
for i, (num, title, desc) in enumerate(prios):
    x = Cm(1.0) + i * (prio_w + Cm(0.3))
    
    # Номер (кружок)
    num_shape = slide.shapes.add_shape(MSO_SHAPE.OVAL, x, prio_y2, Cm(0.9), Cm(0.9))
    num_shape.fill.solid(); num_shape.fill.fore_color.rgb = COL_ACCENT
    num_shape.line.fill.background()
    num_shape.shadow.inherit = False
    tf = num_shape.text_frame; tf.word_wrap = False
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = num
    r.font.name = 'Arial'; r.font.size = Pt(16); r.font.bold = True; r.font.color.rgb = COL_WHITE
    
    # Заголовок
    t_box = slide.shapes.add_textbox(x + Cm(1.1), prio_y2, prio_w - Cm(1.1), Cm(0.6))
    tf = t_box.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.LEFT
    r = p.add_run()
    r.text = title
    r.font.name = 'Arial'; r.font.size = Pt(12); r.font.bold = True; r.font.color.rgb = COL_PRIMARY
    
    # Описание
    d_box = slide.shapes.add_textbox(x + Cm(1.1), prio_y2 + Cm(0.55), prio_w - Cm(1.1), Cm(0.9))
    tf = d_box.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.LEFT
    r = p.add_run()
    r.text = desc
    r.font.name = 'Arial'; r.font.size = Pt(9); r.font.color.rgb = COL_BODY

# === Подвал ===
footer_y = Cm(18.2)
footer = slide.shapes.add_textbox(Cm(1.0), footer_y, Cm(31), Cm(0.5))
tf = footer.text_frame
p = tf.paragraphs[0]; p.alignment = PP_ALIGN.LEFT
r = p.add_run()
r.text = 'TITAN_анализ_дистрибьюторов_для_программы_v2.xlsx — 4 листа: Сводка / Дистрибьюторы по группам (109) / 30 потерянных клиентов / 16 новых клиентов 2026'
r.font.name = 'Arial'; r.font.size = Pt(8); r.font.italic = True; r.font.color.rgb = COL_GRAY

prs.save(OUT)
print(f'OK: {OUT}')
print(f'Размер: {os.path.getsize(OUT) / 1024:.1f} KB')
