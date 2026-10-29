#!/usr/bin/env python3
"""
Создание финального слайда «Призыв к действию + контакты» для презентации Эмпилс.
Поле телефона оставлено пустым — региональный менеджер вписывает вручную
перед отправкой или распечаткой.

Сохраняется как standalone .pptx (один слайд 16:9) — менеджер может
вставить его в любую презентацию через «Copy → Paste Slide» или использовать отдельно.
"""

import os
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn

# === Брендовые константы EMPILS ===
ORANGE = RGBColor(0xF4, 0x76, 0x29)         # основной фирменный оранжевый
ORANGE_LIGHT = RGBColor(0xFB, 0xC9, 0x9A)   # светлый акцент
DARK = RGBColor(0x31, 0x31, 0x31)           # основной текст
GRAY = RGBColor(0x6B, 0x6B, 0x6B)           # подписи
LIGHT_GRAY = RGBColor(0xE8, 0xE8, 0xE8)     # разделители
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
BG = RGBColor(0xFA, 0xFA, 0xFA)             # фон слайда

FONT = "Calibri"

# Логотип из исходной презентации
LOGO_PATH = "/home/z/my-project/extracted_images/partners_pres/ppt/media/image1.png"

OUT_PATH = "/home/z/my-project/download/Эмпилс_Финальный_слайд_Контакты.pptx"

# === Создаём презентацию 16:9 ===
prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

blank_layout = prs.slide_layouts[6]
slide = prs.slides.add_slide(blank_layout)

# === Фон слайда ===
bg = slide.background
bg.fill.solid()
bg.fill.fore_color.rgb = BG

# === Верхняя оранжевая полоса (бренд-маркер) ===
top_bar = slide.shapes.add_shape(
    MSO_SHAPE.RECTANGLE,
    Inches(0), Inches(0),
    prs.slide_width, Inches(0.35)
)
top_bar.fill.solid()
top_bar.fill.fore_color.rgb = ORANGE
top_bar.line.fill.background()
top_bar.shadow.inherit = False

# === Логотип EMPILS (верхний левый угол) ===
slide.shapes.add_picture(
    LOGO_PATH,
    Inches(0.6), Inches(0.6),
    width=Inches(2.2)
)

# === Бейдж года справа сверху ===
year_box = slide.shapes.add_textbox(
    Inches(11.0), Inches(0.55),
    Inches(2.0), Inches(0.5)
)
year_tf = year_box.text_frame
year_tf.margin_top = Pt(0); year_tf.margin_bottom = Pt(0)
year_tf.margin_left = Pt(0); year_tf.margin_right = Pt(0)
p = year_tf.paragraphs[0]
p.alignment = PP_ALIGN.RIGHT
r = p.add_run(); r.text = "2026"
r.font.name = FONT; r.font.size = Pt(20); r.font.bold = True
r.font.color.rgb = ORANGE

# === Главный заголовок (CTA) ===
title_box = slide.shapes.add_textbox(
    Inches(0.9), Inches(1.9),
    Inches(11.5), Inches(1.4)
)
title_tf = title_box.text_frame
title_tf.margin_top = Pt(0); title_tf.margin_bottom = Pt(0)
title_tf.word_wrap = True

p1 = title_tf.paragraphs[0]
p1.alignment = PP_ALIGN.LEFT
r1 = p1.add_run(); r1.text = "СТАНЬ ПАРТНЁРОМ ЭМПИЛС"
r1.font.name = FONT; r1.font.size = Pt(54); r1.font.bold = True
r1.font.color.rgb = DARK

# Подзаголовок
p2 = title_tf.add_paragraph()
p2.alignment = PP_ALIGN.LEFT
p2.space_before = Pt(8)
r2 = p2.add_run()
r2.text = "Развивайте бизнес вместе с надёжным производителем ЛКМ"
r2.font.name = FONT; r2.font.size = Pt(20)
r2.font.color.rgb = GRAY
r2.font.italic = True

# === Оранжевая разделительная линия ===
divider = slide.shapes.add_shape(
    MSO_SHAPE.RECTANGLE,
    Inches(0.9), Inches(3.5),
    Inches(1.8), Pt(4)
)
divider.fill.solid()
divider.fill.fore_color.rgb = ORANGE
divider.line.fill.background()
divider.shadow.inherit = False

# === Левая колонка: преимущества партнёрства ===
left_box = slide.shapes.add_textbox(
    Inches(0.9), Inches(3.8),
    Inches(5.8), Inches(3.2)
)
left_tf = left_box.text_frame
left_tf.margin_top = Pt(0); left_tf.margin_bottom = Pt(0)
left_tf.word_wrap = True

benefits = [
    ("11 брендов", "от эконома до премиума — полный спектр ЛКМ"),
    ("2% от оборота", "трейд-маркетинговая поддержка"),
    ("39,3 тыс. тонн", "производственная мощность в год"),
    (">100 дистрибьюторов", "уже работают с нами в РФ и СНГ"),
    ("Команда ЭТП", "120 РТТ на 1 представителя в поле"),
]

for i, (head, sub) in enumerate(benefits):
    if i == 0:
        p = left_tf.paragraphs[0]
    else:
        p = left_tf.add_paragraph()
        p.space_before = Pt(10)
    p.alignment = PP_ALIGN.LEFT

    r_dot = p.add_run(); r_dot.text = "▸  "
    r_dot.font.name = FONT; r_dot.font.size = Pt(16); r_dot.font.bold = True
    r_dot.font.color.rgb = ORANGE

    r_h = p.add_run(); r_h.text = head
    r_h.font.name = FONT; r_h.font.size = Pt(15); r_h.font.bold = True
    r_h.font.color.rgb = DARK

    r_s = p.add_run(); r_s.text = "  —  " + sub
    r_s.font.name = FONT; r_s.font.size = Pt(13)
    r_s.font.color.rgb = GRAY

# === Вертикальный разделитель между колонками ===
v_div = slide.shapes.add_shape(
    MSO_SHAPE.RECTANGLE,
    Inches(7.0), Inches(3.8),
    Pt(1.5), Inches(3.0)
)
v_div.fill.solid()
v_div.fill.fore_color.rgb = LIGHT_GRAY
v_div.line.fill.background()
v_div.shadow.inherit = False

# === Правая колонка: КОНТАКТЫ ===
# Заголовок блока контактов
ct_header = slide.shapes.add_textbox(
    Inches(7.4), Inches(3.8),
    Inches(5.5), Inches(0.5)
)
ct_tf = ct_header.text_frame
ct_tf.margin_top = Pt(0); ct_tf.margin_bottom = Pt(0)
p = ct_tf.paragraphs[0]
p.alignment = PP_ALIGN.LEFT
r = p.add_run(); r.text = "СВЯЖИТЕСЬ С НАМИ"
r.font.name = FONT; r.font.size = Pt(13); r.font.bold = True
r.font.color.rgb = ORANGE

# === Карточка контактов ===
card = slide.shapes.add_shape(
    MSO_SHAPE.ROUNDED_RECTANGLE,
    Inches(7.4), Inches(4.3),
    Inches(5.5), Inches(2.7)
)
card.fill.solid()
card.fill.fore_color.rgb = WHITE
card.line.color.rgb = LIGHT_GRAY
card.line.width = Pt(1)
card.shadow.inherit = False
# Скругление углов поменьше
try:
    card.adjustments[0] = 0.05
except Exception:
    pass

# Заполняем карточку
card_tf = card.text_frame
card_tf.margin_top = Pt(16); card_tf.margin_bottom = Pt(16)
card_tf.margin_left = Pt(20); card_tf.margin_right = Pt(20)
card_tf.word_wrap = True
card_tf.vertical_anchor = MSO_ANCHOR.MIDDLE

# Строка: Менеджер (заполняется менеджером)
def add_contact_line(tf, first_paragraph, label, value, label_color=GRAY, value_color=DARK,
                     value_bold=False, value_size=14, label_size=11, value_placeholder=False):
    """Add one contact row: bold orange label + value."""
    if first_paragraph:
        p = tf.paragraphs[0]
    else:
        p = tf.add_paragraph()
        p.space_before = Pt(10)
    p.alignment = PP_ALIGN.LEFT

    r_label = p.add_run(); r_label.text = label + "  "
    r_label.font.name = FONT; r_label.font.size = Pt(label_size); r_label.font.bold = True
    r_label.font.color.rgb = label_color

    r_val = p.add_run(); r_val.text = value
    r_val.font.name = FONT; r_val.font.size = Pt(value_size)
    r_val.font.color.rgb = value_color
    r_val.font.bold = value_bold

# 1. Ваш менеджер — поле для имени (тоже оставляем пустым, менеджер сам впишет)
add_contact_line(card_tf, True,
                 "ВАШ МЕНЕДЖЕР:",
                 "____________________________",
                 label_color=ORANGE, value_color=GRAY,
                 value_size=13, label_size=11)

# 2. Телефон — ГЛАВНОЕ: пустое поле для ручного заполнения
p_phone = card_tf.add_paragraph()
p_phone.space_before = Pt(14)
p_phone.alignment = PP_ALIGN.LEFT

r_ph_lbl = p_phone.add_run()
r_ph_lbl.text = "ТЕЛЕФОН:  "
r_ph_lbl.font.name = FONT; r_ph_lbl.font.size = Pt(11); r_ph_lbl.font.bold = True
r_ph_lbl.font.color.rgb = ORANGE

r_ph_val = p_phone.add_run()
r_ph_val.text = "+7 (___) ___-__-__"
r_ph_val.font.name = FONT; r_ph_val.font.size = Pt(18)
r_ph_val.font.color.rgb = LIGHT_GRAY
r_ph_val.font.bold = True

# Подсказка под телефоном
p_hint = card_tf.add_paragraph()
p_hint.space_before = Pt(2)
p_hint.alignment = PP_ALIGN.LEFT
r_hint = p_hint.add_run()
r_hint.text = "  ← впишите ваш номер перед отправкой / распечаткой"
r_hint.font.name = FONT; r_hint.font.size = Pt(9)
r_hint.font.italic = True
r_hint.font.color.rgb = ORANGE

# 3. Email
add_contact_line(card_tf, False,
                 "EMAIL:  ",
                 "info@empils.ru",
                 label_color=ORANGE, value_color=DARK,
                 value_size=14, label_size=11, value_bold=False)

# 4. Сайт
add_contact_line(card_tf, False,
                 "САЙТ:  ",
                 "www.empils.ru",
                 label_color=ORANGE, value_color=DARK,
                 value_size=14, label_size=11, value_bold=False)

# 5. Адрес
add_contact_line(card_tf, False,
                 "АДРЕС:  ",
                 "г. Ростов-на-Дону, ул. Ленина, 158",
                 label_color=ORANGE, value_color=DARK,
                 value_size=13, label_size=11)

# === Нижняя CTA-полоса (оранжевая, финальный призыв) ===
bottom_cta = slide.shapes.add_shape(
    MSO_SHAPE.RECTANGLE,
    Inches(0), Inches(7.05),
    prs.slide_width, Inches(0.45)
)
bottom_cta.fill.solid()
bottom_cta.fill.fore_color.rgb = ORANGE
bottom_cta.line.fill.background()
bottom_cta.shadow.inherit = False

bottom_tf = bottom_cta.text_frame
bottom_tf.margin_top = Pt(0); bottom_tf.margin_bottom = Pt(0)
bottom_tf.vertical_anchor = MSO_ANCHOR.MIDDLE
p = bottom_tf.paragraphs[0]
p.alignment = PP_ALIGN.CENTER
r = p.add_run()
r.text = "230 ₽/кг — зарабатываем вместе  ·  ЗАО «Эмпилс» · 2026"
r.font.name = FONT; r.font.size = Pt(13); r.font.bold = True
r.font.color.rgb = WHITE

# === Сохраняем ===
os.makedirs(os.path.dirname(OUT_PATH), exist_ok=True)
prs.save(OUT_PATH)
print(f"✅ Слайд сохранён: {OUT_PATH}")
print(f"   Размер: {os.path.getsize(OUT_PATH):,} байт")
