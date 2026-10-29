#!/usr/bin/env python3
"""
PPTX-макет листовки А4 (односторонняя) для TITAN 3в1.
3 блока: 
1. Что это + для каких поверхностей
2. Инструкция по применению (5 шагов)
3. TCO с фото до/после
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Cm, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
import os

OUT = '/home/z/my-project/download/TITAN_листовка_А4_макет.pptx'

# Цвета
COL_PRIMARY = RGBColor(0x1E, 0x3A, 0x5F)   # тёмно-синий
COL_ACCENT = RGBColor(0xF4, 0x79, 0x20)    # оранжевый
COL_RED = RGBColor(0xC6, 0x28, 0x28)
COL_GREEN = RGBColor(0x16, 0xA3, 0x4A)
COL_BODY = RGBColor(0x1A, 0x1F, 0x36)
COL_GRAY = RGBColor(0x6B, 0x72, 0x80)
COL_WHITE = RGBColor(0xFF, 0xFF, 0xFF)
COL_LIGHT_BG = RGBColor(0xF3, 0xF4, 0xF6)
COL_LIGHT_BLUE = RGBColor(0xDB, 0xEA, 0xFE)
COL_LIGHT_GREEN = RGBColor(0xDC, 0xFC, 0xE7)
COL_LIGHT_RED = RGBColor(0xFE, 0xE2, 0xE2)
COL_LIGHT_ORANGE = RGBColor(0xFF, 0xED, 0xD5)

# А4 — 21 × 29.7 см (портрет)
prs = Presentation()
prs.slide_width = Cm(21.0)
prs.slide_height = Cm(29.7)

slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank

# === БЕЛЫЙ ФОН ===
bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
bg.fill.solid(); bg.fill.fore_color.rgb = COL_WHITE
bg.line.fill.background()
bg.shadow.inherit = False

# === ВЕРХНЯЯ ПОЛОСА (заголовок) ===
top_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, Cm(2.2))
top_bar.fill.solid(); top_bar.fill.fore_color.rgb = COL_PRIMARY
top_bar.line.fill.background()
top_bar.shadow.inherit = False

# Логотип TITAN
logo = slide.shapes.add_textbox(Cm(1.0), Cm(0.3), Cm(10), Cm(1.6))
tf = logo.text_frame; tf.word_wrap = True
p = tf.paragraphs[0]; p.alignment = PP_ALIGN.LEFT
r = p.add_run()
r.text = "TITAN"
r.font.name = 'Arial'; r.font.size = Pt(36); r.font.bold = True; r.font.color.rgb = COL_ACCENT
r2 = p.add_run()
r2.text = "  3в1 по ржавчине"
r2.font.name = 'Arial'; r2.font.size = Pt(18); r2.font.bold = False; r2.font.color.rgb = COL_WHITE

# Слоган справа
slogan = slide.shapes.add_textbox(Cm(11), Cm(0.5), Cm(9), Cm(1.2))
tf = slogan.text_frame; tf.word_wrap = True
p = tf.paragraphs[0]; p.alignment = PP_ALIGN.RIGHT
r = p.add_run()
r.text = "Красишь 1 раз — служит 10 лет"
r.font.name = 'Arial'; r.font.size = Pt(14); r.font.bold = True; r.font.color.rgb = COL_ACCENT

# ========================================================
# БЛОК 1: ЧТО ЭТО И ДЛЯ КАКИХ ПОВЕРХНОСТЕЙ (строки 2.5–8.5)
# ========================================================

# Заголовок блока 1
h1 = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Cm(1.0), Cm(2.8), Cm(19.0), Cm(0.8))
h1.fill.solid(); h1.fill.fore_color.rgb = COL_ACCENT
h1.line.fill.background()
h1.shadow.inherit = False
tf = h1.text_frame
p = tf.paragraphs[0]; p.alignment = PP_ALIGN.LEFT
tf.margin_left = Cm(0.4)
r = p.add_run()
r.text = "  ЧТО ЭТО И ДЛЯ КАКИХ ПОВЕРХНОСТЕЙ"
r.font.name = 'Arial'; r.font.size = Pt(13); r.font.bold = True; r.font.color.rgb = COL_WHITE

# Описание продукта
desc_box = slide.shapes.add_textbox(Cm(1.0), Cm(3.8), Cm(19.0), Cm(2.2))
tf = desc_box.text_frame; tf.word_wrap = True
p = tf.paragraphs[0]; p.alignment = PP_ALIGN.LEFT
r = p.add_run()
r.text = "Грунт-эмаль по ржавчине 3в1 — грунт, эмаль и преобразователь ржавчины в одной банке."
r.font.name = 'Arial'; r.font.size = Pt(11); r.font.bold = True; r.font.color.rgb = COL_PRIMARY

p2 = tf.add_paragraph(); p2.alignment = PP_ALIGN.LEFT
p2.space_before = Pt(4)
r = p2.add_run()
r.text = ("Применяется для защиты и окраски стальных и чугунных поверхностей с остатками "
          "плотно держащейся ржавчины, эксплуатируемых в атмосферных условиях и внутри помещений. "
          "Не требует грунтовки — наносится прямо на ржавчину.")
r.font.name = 'Arial'; r.font.size = Pt(10); r.font.color.rgb = COL_BODY

# Поверхности — 4 плашки
surfaces = [
    ("Металл\n(сталь, чугун)", COL_LIGHT_RED),
    ("Оцинковка\n(+ цветные металлы)", COL_LIGHT_BLUE),
    ("Нержавейка", COL_LIGHT_GREEN),
    ("Дерево", COL_LIGHT_ORANGE),
]
for i, (text, fill) in enumerate(surfaces):
    x = Cm(1.0 + i * 4.9)
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, Cm(6.2), Cm(4.5), Cm(1.8))
    shape.fill.solid(); shape.fill.fore_color.rgb = fill
    shape.line.color.rgb = COL_PRIMARY; shape.line.width = Pt(1)
    shape.shadow.inherit = False
    tf = shape.text_frame; tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = text
    r.font.name = 'Arial'; r.font.size = Pt(10); r.font.bold = True; r.font.color.rgb = COL_PRIMARY

# Ключевые характеристики — 3 плашки
specs = [
    ("Степень блеска", "Глянцевая"),
    ("Время высыхания", "«От пыли» — 1,5 ч"),
    ("Срок службы", "10 лет"),
]
for i, (label, value) in enumerate(specs):
    x = Cm(1.0 + i * 6.4)
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, Cm(8.3), Cm(6.0), Cm(1.5))
    shape.fill.solid(); shape.fill.fore_color.rgb = COL_LIGHT_BG
    shape.line.color.rgb = COL_GRAY; shape.line.width = Pt(0.5)
    shape.shadow.inherit = False
    tf = shape.text_frame; tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = label + '\n'
    r.font.name = 'Arial'; r.font.size = Pt(9); r.font.color.rgb = COL_GRAY
    r2 = p.add_run()
    r2.text = value
    r2.font.name = 'Arial'; r2.font.size = Pt(12); r2.font.bold = True; r2.font.color.rgb = COL_PRIMARY

# ========================================================
# БЛОК 2: ИНСТРУКЦИЯ ПО ПРИМЕНЕНИЮ (строки 10–16)
# ========================================================

# Заголовок блока 2
h2 = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Cm(1.0), Cm(10.2), Cm(19.0), Cm(0.8))
h2.fill.solid(); h2.fill.fore_color.rgb = COL_ACCENT
h2.line.fill.background()
h2.shadow.inherit = False
tf = h2.text_frame
p = tf.paragraphs[0]; p.alignment = PP_ALIGN.LEFT
tf.margin_left = Cm(0.4)
r = p.add_run()
r.text = "  ИНСТРУКЦИЯ ПО ПРИМЕНЕНИЮ — 5 ШАГОВ"
r.font.name = 'Arial'; r.font.size = Pt(13); r.font.bold = True; r.font.color.rgb = COL_WHITE

# 5 шагов — горизонтальные плашки с номерами
steps = [
    ("1", "Очистить", "Удалить рыхлую ржавчину\nметаллической щёткой.\nНе нужно до блеска!"),
    ("2", "Обезжирить", "Протереть растворителем\nили уайт-спиритом."),
    ("3", "Перемешать", "Тщательно перемешать\n1 минуту. Грунт уже\nвнутри — отдельный не нужен."),
    ("4", "Покрасить", "Нанести 2 слоя кистью\nили валиком. Между слоями\n— 1,5 часа на отлип."),
    ("5", "Готово", "Покрытие набирает\nпрочность за 7 дней.\nСрок службы — 10 лет."),
]

for i, (num, title, desc) in enumerate(steps):
    x = Cm(0.5 + i * 4.1)
    y = Cm(11.3)
    
    # Номер (кружок)
    num_shape = slide.shapes.add_shape(MSO_SHAPE.OVAL, x + Cm(1.3), y, Cm(1.2), Cm(1.2))
    num_shape.fill.solid(); num_shape.fill.fore_color.rgb = COL_ACCENT
    num_shape.line.fill.background()
    num_shape.shadow.inherit = False
    tf = num_shape.text_frame
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    r = p.add_run()
    r.text = num
    r.font.name = 'Arial'; r.font.size = Pt(18); r.font.bold = True; r.font.color.rgb = COL_WHITE
    
    # Карточка шага
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y + Cm(1.4), Cm(3.8), Cm(4.0))
    card.fill.solid(); card.fill.fore_color.rgb = COL_LIGHT_BG
    card.line.color.rgb = COL_GRAY; card.line.width = Pt(0.5)
    card.shadow.inherit = False
    
    tf = card.text_frame; tf.word_wrap = True
    tf.margin_top = Cm(0.3)
    tf.margin_left = Cm(0.2)
    tf.margin_right = Cm(0.2)
    
    # Заголовок шага
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = title
    r.font.name = 'Arial'; r.font.size = Pt(11); r.font.bold = True; r.font.color.rgb = COL_PRIMARY
    
    # Описание
    p2 = tf.add_paragraph(); p2.alignment = PP_ALIGN.CENTER
    p2.space_before = Pt(4)
    r = p2.add_run()
    r.text = desc
    r.font.name = 'Arial'; r.font.size = Pt(9); r.font.color.rgb = COL_BODY

# Стрелки между шагами
for i in range(4):
    x = Cm(4.2 + i * 4.1)
    arrow = slide.shapes.add_textbox(x, Cm(11.5), Cm(1.0), Cm(1.0))
    tf = arrow.text_frame
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    r = p.add_run()
    r.text = "→"
    r.font.name = 'Arial'; r.font.size = Pt(20); r.font.bold = True; r.font.color.rgb = COL_ACCENT

# ========================================================
# БЛОК 3: TCO С ПРИМЕРАМИ ДО/ПОСЛЕ (строки 16.5–28)
# ========================================================

# Заголовок блока 3
h3 = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Cm(1.0), Cm(16.2), Cm(19.0), Cm(0.8))
h3.fill.solid(); h3.fill.fore_color.rgb = COL_ACCENT
h3.line.fill.background()
h3.shadow.inherit = False
tf = h3.text_frame
p = tf.paragraphs[0]; p.alignment = PP_ALIGN.LEFT
tf.margin_left = Cm(0.4)
r = p.add_run()
r.text = "  TCO — СОВОКУПНАЯ СТОИМОСТЬ ВЛАДЕНИЯ ЗА 10 ЛЕТ"
r.font.name = 'Arial'; r.font.size = Pt(13); r.font.bold = True; r.font.color.rgb = COL_WHITE

# Подзаголовок
sub = slide.shapes.add_textbox(Cm(1.0), Cm(17.2), Cm(19.0), Cm(0.7))
tf = sub.text_frame; tf.word_wrap = True
p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
r = p.add_run()
r.text = "Красите 1 раз на 10 лет вместо 4 раз — экономия до 59%"
r.font.name = 'Arial'; r.font.size = Pt(12); r.font.bold = True; r.font.color.rgb = COL_GREEN

# === 3 объекта TCO — горизонтальные карточки ===
objects = [
    {
        'name': 'Дверь металлическая',
        'area': '4.4 м²',
        'pf_total': '4 640 ₽',
        'titan_total': '1 610 ₽',
        'saving': '3 030 ₽',
        'saving_pct': '65%',
        'img_path': '/tmp/door_before_after.png',
    },
    {
        'name': 'Гараж металлический',
        'area': '54 м²',
        'pf_total': '43 692 ₽',
        'titan_total': '17 580 ₽',
        'saving': '26 112 ₽',
        'saving_pct': '59%',
        'img_path': None,
    },
    {
        'name': 'Ворота распашные',
        'area': '16 м²',
        'pf_total': '13 860 ₽',
        'titan_total': '5 320 ₽',
        'saving': '8 540 ₽',
        'saving_pct': '61%',
        'img_path': '/tmp/gate_after.jpg',
    },
]

for i, obj in enumerate(objects):
    x = Cm(0.5 + i * 7.0)
    y = Cm(18.2)
    w = Cm(6.5)
    h = Cm(8.5)
    
    # Карточка
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    card.fill.solid(); card.fill.fore_color.rgb = COL_LIGHT_BG
    card.line.color.rgb = COL_PRIMARY; card.line.width = Pt(1)
    card.shadow.inherit = False
    
    # Заголовок объекта
    title_box = slide.shapes.add_textbox(x + Cm(0.3), y + Cm(0.2), w - Cm(0.6), Cm(1.0))
    tf = title_box.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = obj['name']
    r.font.name = 'Arial'; r.font.size = Pt(12); r.font.bold = True; r.font.color.rgb = COL_PRIMARY
    p2 = tf.add_paragraph(); p2.alignment = PP_ALIGN.CENTER
    r = p2.add_run()
    r.text = f"Площадь: {obj['area']}"
    r.font.name = 'Arial'; r.font.size = Pt(9); r.font.color.rgb = COL_GRAY
    
    # Фото до/после (если есть)
    if obj['img_path'] and os.path.exists(obj['img_path']):
        try:
            img_w = Cm(5.5)
            img_h = Cm(3.0)
            img_x = x + (w - img_w) / 2
            img_y = y + Cm(1.4)
            slide.shapes.add_picture(obj['img_path'], img_x, img_y, img_w, img_h)
        except:
            pass
    else:
        # Заглушка
        placeholder = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x + Cm(0.5), y + Cm(1.4), Cm(5.5), Cm(3.0))
        placeholder.fill.solid(); placeholder.fill.fore_color.rgb = COL_LIGHT_BLUE
        placeholder.line.color.rgb = COL_GRAY; placeholder.line.width = Pt(0.5)
        placeholder.shadow.inherit = False
        tf = placeholder.text_frame
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = "Фото до/после"
        r.font.name = 'Arial'; r.font.size = Pt(10); r.font.color.rgb = COL_GRAY
    
    # Сравнение цен
    # ПФ-115
    pf_box = slide.shapes.add_textbox(x + Cm(0.3), y + Cm(4.7), w - Cm(0.6), Cm(1.2))
    tf = pf_box.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = "ПФ-115: "
    r.font.name = 'Arial'; r.font.size = Pt(10); r.font.color.rgb = COL_GRAY
    r2 = p.add_run()
    r2.text = f"{obj['pf_total']}"
    r2.font.name = 'Arial'; r2.font.size = Pt(11); r2.font.bold = True; r2.font.color.rgb = COL_RED
    r3 = p.add_run()
    r3.text = " (4 перекраски)"
    r3.font.name = 'Arial'; r3.font.size = Pt(8); r3.font.color.rgb = COL_GRAY
    
    # TITAN
    titan_box = slide.shapes.add_textbox(x + Cm(0.3), y + Cm(5.8), w - Cm(0.6), Cm(1.2))
    tf = titan_box.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = "TITAN: "
    r.font.name = 'Arial'; r.font.size = Pt(10); r.font.color.rgb = COL_GRAY
    r2 = p.add_run()
    r2.text = f"{obj['titan_total']}"
    r2.font.name = 'Arial'; r2.font.size = Pt(11); r2.font.bold = True; r2.font.color.rgb = COL_GREEN
    r3 = p.add_run()
    r3.text = " (1 покраска)"
    r3.font.name = 'Arial'; r3.font.size = Pt(8); r3.font.color.rgb = COL_GRAY
    
    # Экономия
    save_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x + Cm(0.5), y + Cm(7.0), Cm(5.5), Cm(1.2))
    save_box.fill.solid(); save_box.fill.fore_color.rgb = COL_ACCENT
    save_box.line.fill.background()
    save_box.shadow.inherit = False
    tf = save_box.text_frame
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = f"ЭКОНОМИЯ: {obj['saving']}"
    r.font.name = 'Arial'; r.font.size = Pt(12); r.font.bold = True; r.font.color.rgb = COL_WHITE
    p2 = tf.add_paragraph(); p2.alignment = PP_ALIGN.CENTER
    r = p2.add_run()
    r.text = f"({obj['saving_pct']} за 10 лет)"
    r.font.name = 'Arial'; r.font.size = Pt(9); r.font.color.rgb = COL_WHITE

# === НИЖНЯЯ ПОЛОСА ===
footer_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, Cm(28.5), prs.slide_width, Cm(1.2))
footer_bar.fill.solid(); footer_bar.fill.fore_color.rgb = COL_PRIMARY
footer_bar.line.fill.background()
footer_bar.shadow.inherit = False

footer = slide.shapes.add_textbox(Cm(1.0), Cm(28.6), Cm(19.0), Cm(1.0))
tf = footer.text_frame; tf.word_wrap = True
tf.vertical_anchor = MSO_ANCHOR.MIDDLE
p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
r = p.add_run()
r.text = "TITAN 3в1 — грунт + эмаль + преобразователь ржавчины в одной банке  ·  "
r.font.name = 'Arial'; r.font.size = Pt(9); r.font.color.rgb = COL_WHITE
r2 = p.add_run()
r2.text = "Красишь 1 раз — служит 10 лет"
r2.font.name = 'Arial'; r2.font.size = Pt(9); r2.font.bold = True; r2.font.color.rgb = COL_ACCENT

# Доп. инфо
info = slide.shapes.add_textbox(Cm(1.0), Cm(27.3), Cm(19.0), Cm(1.0))
tf = info.text_frame; tf.word_wrap = True
p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
r = p.add_run()
r.text = ("Фасовка: 0.4 / 0.75 / 2 / 5 л  ·  Расход: 50–120 г/м²  ·  "
         "Разбавление: ксилол до 5%  ·  Температура нанесения: +10°C до +30°C  ·  "
         "titanpaints.ru")
r.font.name = 'Arial'; r.font.size = Pt(8); r.font.color.rgb = COL_GRAY

prs.save(OUT)
print(f'OK: {OUT}')
print(f'Размер: {os.path.getsize(OUT) / 1024:.1f} KB')
