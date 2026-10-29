#!/usr/bin/env python3
"""
Слайд: Конкурентное окружение TITAN по 3 категориям.
Цены TITAN — расчёт от отгрузочной цены Эмпилс + 25% дистрибьютор + 25% розница.
Цены конкурентов — из слайдов 11-13 презентации TITAN (Все Инструменты.ру).
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

# === COLORS ===
RED = RGBColor(0xC6, 0x28, 0x28)
DARK_RED = RGBColor(0x8E, 0x00, 0x00)
GREEN = RGBColor(0x16, 0xA3, 0x4A)
DARK_GREEN = RGBColor(0x0B, 0x6E, 0x31)
DARK = RGBColor(0x1A, 0x1F, 0x36)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
GRAY = RGBColor(0x9C, 0xA3, 0xAF)
LIGHT_GRAY = RGBColor(0xE5, 0xE7, 0xEB)
ORANGE = RGBColor(0xF4, 0x79, 0x20)
YELLOW = RGBColor(0xFB, 0xBF, 0x24)
GOLD = RGBColor(0xB8, 0x86, 0x0B)
DARK_BLUE = RGBColor(0x1E, 0x3A, 0x5F)
LIGHT_GREEN = RGBColor(0xDC, 0xFC, 0xE7)
LIGHT_ORANGE = RGBColor(0xFF, 0xED, 0xD5)
LIGHT_RED = RGBColor(0xFE, 0xE2, 0xE2)
LIGHT_BLUE = RGBColor(0xDB, 0xEA, 0xFE)

OUTPUT = '/home/z/my-project/download/Слайд_конкурентное_окружение.pptx'

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)


def add_textbox(slide, left, top, width, height, text, font_size=18,
                color=DARK, bold=False, align=PP_ALIGN.LEFT, font_name='Arial',
                anchor=MSO_ANCHOR.TOP, italic=False):
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = Pt(2); tf.margin_right = Pt(2)
    tf.margin_top = Pt(2); tf.margin_bottom = Pt(2)
    tf.vertical_anchor = anchor
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(font_size)
    run.font.color.rgb = color
    run.font.bold = bold
    run.font.italic = italic
    run.font.name = font_name
    return tb


def add_shape(slide, shape_type, left, top, width, height, fill_color=None,
              line_color=None, line_width=None):
    shp = slide.shapes.add_shape(shape_type, left, top, width, height)
    if fill_color is None:
        shp.fill.background()
    else:
        shp.fill.solid()
        shp.fill.fore_color.rgb = fill_color
    if line_color is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line_color
        if line_width is not None:
            shp.line.width = line_width
    shp.shadow.inherit = False
    return shp


def add_multi_text(slide, left, top, width, height, lines,
                   anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER):
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = Pt(1); tf.margin_right = Pt(1)
    tf.margin_top = Pt(1); tf.margin_bottom = Pt(1)
    for i, (text, size, color, bold) in enumerate(lines):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.alignment = align
        r = p.add_run()
        r.text = text
        r.font.size = Pt(size)
        r.font.color.rgb = color
        r.font.bold = bold
        r.font.name = 'Arial'
    return tb


def main():
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H

    slide = prs.slides.add_slide(prs.slide_layouts[6])

    add_shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H,
              fill_color=WHITE, line_color=None)

    # === ШАПКА ===
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, Inches(0.65),
              fill_color=RED, line_color=None)
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0, Inches(0.65), SLIDE_W, Inches(0.04),
              fill_color=YELLOW, line_color=None)
    add_textbox(slide, Inches(0.4), Inches(0.08), SLIDE_W - Inches(0.8), Inches(0.5),
                'КОНКУРЕНТНОЕ ОКРУЖЕНИЕ — TITAN vs КОНКУРЕНТЫ',
                font_size=22, color=WHITE, bold=True,
                align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)
    add_textbox(slide, SLIDE_W - Inches(5.0), Inches(0.08),
                Inches(4.6), Inches(0.5),
                '3 категории  •  Цены Все Инструменты.ру',
                font_size=11, color=WHITE, bold=False,
                align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)

    # === МЕТОДОЛОГИЯ РАСЧЁТА ===
    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
              Inches(0.4), Inches(0.85), SLIDE_W - Inches(0.8), Inches(0.55),
              fill_color=LIGHT_BLUE, line_color=DARK_BLUE, line_width=Pt(1))
    add_textbox(slide, Inches(0.4), Inches(0.85), SLIDE_W - Inches(0.8), Inches(0.55),
                'Расчёт цены TITAN: отгрузочная цена Эмпилс +25% маржа дистрибьютора +25% маржа розницы = цена на полке. Цены конкурентов — сайт Все Инструменты.ру.',
                font_size=10, color=DARK_BLUE, bold=True,
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # === 3 КАТЕГОРИИ — по горизонтали ===

    # Категория 1: 3в1 по ржавчине 0.75 л
    cat1_x = Inches(0.4)
    cat2_x = Inches(4.65)
    cat3_x = Inches(8.90)
    cat_w = Inches(4.10)
    cat_y = Inches(1.55)
    cat_h = Inches(5.30)

    # ===== КАТЕГОРИЯ 1: 3в1 ПО РЖАВЧИНЕ =====
    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
              cat1_x, cat_y, cat_w, cat_h,
              fill_color=WHITE, line_color=RED, line_width=Pt(2))
    # Шапка
    add_shape(slide, MSO_SHAPE.RECTANGLE,
              cat1_x, cat_y, cat_w, Inches(0.70),
              fill_color=RED, line_color=None)
    add_textbox(slide, cat1_x, cat_y + Inches(0.05), cat_w, Inches(0.35),
                'КАТЕГОРИЯ 1', font_size=9, color=YELLOW, bold=True,
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_textbox(slide, cat1_x, cat_y + Inches(0.32), cat_w, Inches(0.40),
                'Грунт-эмаль 3в1 по ржавчине 0.75 л', font_size=11, color=WHITE, bold=True,
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # Расчёт цены TITAN
    add_shape(slide, MSO_SHAPE.RECTANGLE,
              cat1_x + Inches(0.15), cat_y + Inches(0.85),
              cat_w - Inches(0.30), Inches(1.20),
              fill_color=LIGHT_GREEN, line_color=GREEN, line_width=Pt(1))
    add_multi_text(slide, cat1_x + Inches(0.15), cat_y + Inches(0.88),
                   cat_w - Inches(0.30), Inches(1.15),
                   [
                       ('TITAN 3в1 чёрный RAL 9005 0.75 л', 10, DARK, True),
                       ('Отгрузка Эмпилс: 494 ₽', 9, DARK, False),
                       ('+25% дистрибьютор: 617 ₽', 9, DARK, False),
                       ('+25% розница: 771 ₽ ← цена на полке', 10, GREEN, True),
                   ])

    # Таблица конкурентов
    comp1_y = cat_y + Inches(2.20)
    add_textbox(slide, cat1_x + Inches(0.15), comp1_y,
                cat_w - Inches(0.30), Inches(0.30),
                'КОНКУРЕНТЫ (Все Инструменты.ру):', font_size=9, color=DARK_BLUE, bold=True,
                align=PP_ALIGN.LEFT)

    comp1_data = [
        ('TITAN 3в1 (Эмпилс)', '0.75 л', '771 ₽', GREEN, True),
        ('Elcon 3в1 глянцевая', '0.8 кг', '958 ₽', DARK, False),
        ('Dufa CORREX 3в1', '0.75 л', '1 020 ₽', DARK, False),
        ('Dali 3в1', '0.75 л', '1 055 ₽', DARK, False),
        ('Husky Klondike', '0.9 л', '1 153 ₽', DARK, False),
        ('Tikkurila METALLISTA', '0.8 л', '2 085 ₽', DARK, False),
    ]
    comp_row_h = Inches(0.40)
    for i, (name, vol, price, color, is_titan) in enumerate(comp1_data):
        cy = comp1_y + Inches(0.35) + comp_row_h * i
        bg = LIGHT_GREEN if is_titan else (LIGHT_GRAY if i % 2 == 0 else WHITE)
        add_shape(slide, MSO_SHAPE.RECTANGLE,
                  cat1_x + Inches(0.15), cy, cat_w - Inches(0.30), comp_row_h,
                  fill_color=bg, line_color=GRAY, line_width=Pt(0.3))
        add_textbox(slide, cat1_x + Inches(0.20), cy, Inches(2.3), comp_row_h,
                    name, font_size=8, color=color, bold=is_titan,
                    align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)
        add_textbox(slide, cat1_x + Inches(2.50), cy, Inches(0.7), comp_row_h,
                    vol, font_size=7, color=GRAY,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_textbox(slide, cat1_x + Inches(3.20), cy, Inches(0.75), comp_row_h,
                    price, font_size=9, color=color, bold=True,
                    align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)

    # Вывод
    add_shape(slide, MSO_SHAPE.RECTANGLE,
              cat1_x + Inches(0.15), cat_y + cat_h - Inches(0.55),
              cat_w - Inches(0.30), Inches(0.45),
              fill_color=GREEN, line_color=None)
    add_textbox(slide, cat1_x + Inches(0.15), cat_y + cat_h - Inches(0.55),
                cat_w - Inches(0.30), Inches(0.45),
                'TITAN дешевле всех на 19-63%', font_size=10, color=YELLOW, bold=True,
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # ===== КАТЕГОРИЯ 2: БЫСТРОСОХНУЩАЯ =====
    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
              cat2_x, cat_y, cat_w, cat_h,
              fill_color=WHITE, line_color=DARK_BLUE, line_width=Pt(2))
    add_shape(slide, MSO_SHAPE.RECTANGLE,
              cat2_x, cat_y, cat_w, Inches(0.70),
              fill_color=DARK_BLUE, line_color=None)
    add_textbox(slide, cat2_x, cat_y + Inches(0.05), cat_w, Inches(0.35),
                'КАТЕГОРИЯ 2', font_size=9, color=YELLOW, bold=True,
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_textbox(slide, cat2_x, cat_y + Inches(0.32), cat_w, Inches(0.40),
                'Быстросохнущая матовая 0.75 л', font_size=11, color=WHITE, bold=True,
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # Расчёт цены TITAN
    add_shape(slide, MSO_SHAPE.RECTANGLE,
              cat2_x + Inches(0.15), cat_y + Inches(0.85),
              cat_w - Inches(0.30), Inches(1.20),
              fill_color=LIGHT_GREEN, line_color=GREEN, line_width=Pt(1))
    add_multi_text(slide, cat2_x + Inches(0.15), cat_y + Inches(0.88),
                   cat_w - Inches(0.30), Inches(1.15),
                   [
                       ('TITAN быстросохнущая матовая 0.75 л', 10, DARK, True),
                       ('Отгрузка Эмпилс: 518 ₽', 9, DARK, False),
                       ('+25% дистрибьютор: 648 ₽', 9, DARK, False),
                       ('+25% розница: 810 ₽ ← цена на полке', 10, GREEN, True),
                   ])

    # Таблица конкурентов
    comp2_y = cat_y + Inches(2.20)
    add_textbox(slide, cat2_x + Inches(0.15), comp2_y,
                cat_w - Inches(0.30), Inches(0.30),
                'КОНКУРЕНТЫ (Все Инструменты.ру):', font_size=9, color=DARK_BLUE, bold=True,
                align=PP_ALIGN.LEFT)

    comp2_data = [
        ('Новбытхим 5в1', '0.75 л', '765 ₽', DARK, False),
        ('TITAN быстросохнущая (Эмпилс)', '0.75 л', '810 ₽', GREEN, True),
        ('Elcon 3в1 быстросохнущая', '0.8 кг', '838 ₽', DARK, False),
        ('Лакра PROF IT 4в1', '0.8 кг', '906 ₽', DARK, False),
    ]
    for i, (name, vol, price, color, is_titan) in enumerate(comp2_data):
        cy = comp2_y + Inches(0.35) + comp_row_h * i
        bg = LIGHT_GREEN if is_titan else (LIGHT_GRAY if i % 2 == 0 else WHITE)
        add_shape(slide, MSO_SHAPE.RECTANGLE,
                  cat2_x + Inches(0.15), cy, cat_w - Inches(0.30), comp_row_h,
                  fill_color=bg, line_color=GRAY, line_width=Pt(0.3))
        add_textbox(slide, cat2_x + Inches(0.20), cy, Inches(2.3), comp_row_h,
                    name, font_size=8, color=color, bold=is_titan,
                    align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)
        add_textbox(slide, cat2_x + Inches(2.50), cy, Inches(0.7), comp_row_h,
                    vol, font_size=7, color=GRAY,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_textbox(slide, cat2_x + Inches(3.20), cy, Inches(0.75), comp_row_h,
                    price, font_size=9, color=color, bold=True,
                    align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)

    # Вывод
    add_shape(slide, MSO_SHAPE.RECTANGLE,
              cat2_x + Inches(0.15), cat_y + cat_h - Inches(0.55),
              cat_w - Inches(0.30), Inches(0.45),
              fill_color=GREEN, line_color=None)
    add_textbox(slide, cat2_x + Inches(0.15), cat_y + cat_h - Inches(0.55),
                cat_w - Inches(0.30), Inches(0.45),
                'TITAN на 2-м месте — премиум-качество', font_size=10, color=YELLOW, bold=True,
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # ===== КАТЕГОРИЯ 3: МОЛОТКОВАЯ =====
    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
              cat3_x, cat_y, cat_w, cat_h,
              fill_color=WHITE, line_color=GOLD, line_width=Pt(2))
    add_shape(slide, MSO_SHAPE.RECTANGLE,
              cat3_x, cat_y, cat_w, Inches(0.70),
              fill_color=GOLD, line_color=None)
    add_textbox(slide, cat3_x, cat_y + Inches(0.05), cat_w, Inches(0.35),
                'КАТЕГОРИЯ 3', font_size=9, color=WHITE, bold=True,
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_textbox(slide, cat3_x, cat_y + Inches(0.32), cat_w, Inches(0.40),
                'Молотковый эффект 0.8 кг', font_size=11, color=WHITE, bold=True,
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # Расчёт цены TITAN
    add_shape(slide, MSO_SHAPE.RECTANGLE,
              cat3_x + Inches(0.15), cat_y + Inches(0.85),
              cat_w - Inches(0.30), Inches(1.20),
              fill_color=LIGHT_GREEN, line_color=GREEN, line_width=Pt(1))
    add_multi_text(slide, cat3_x + Inches(0.15), cat_y + Inches(0.88),
                   cat_w - Inches(0.30), Inches(1.15),
                   [
                       ('TITAN молотковая 0.8 кг', 10, DARK, True),
                       ('Отгрузка Эмпилс: 701 ₽', 9, DARK, False),
                       ('+25% дистрибьютор: 877 ₽', 9, DARK, False),
                       ('+25% розница: 1 096 ₽ ← цена на полке', 10, GREEN, True),
                   ])

    # Таблица конкурентов
    comp3_y = cat_y + Inches(2.20)
    add_textbox(slide, cat3_x + Inches(0.15), comp3_y,
                cat_w - Inches(0.30), Inches(0.30),
                'КОНКУРЕНТЫ (Все Инструменты.ру):', font_size=9, color=DARK_BLUE, bold=True,
                align=PP_ALIGN.LEFT)

    comp3_data = [
        ('Dali 3в1 молотковая', '0.75 л', '1 057 ₽', DARK, False),
        ('Elcon кузнечная', '0.8 кг', '1 097 ₽', DARK, False),
        ('TITAN молотковая (Эмпилс)', '0.8 кг', '1 096 ₽', GREEN, True),
        ('Husky Klondike молотковая', '0.9 л', '1 196 ₽', DARK, False),
        ('Лакра молотковая', '0.8 кг', '1 308 ₽', DARK, False),
        ('Dufa HAMMERLACK', '0.75 л', '1 653 ₽', DARK, False),
    ]
    for i, (name, vol, price, color, is_titan) in enumerate(comp3_data):
        cy = comp3_y + Inches(0.35) + comp_row_h * i
        bg = LIGHT_GREEN if is_titan else (LIGHT_GRAY if i % 2 == 0 else WHITE)
        add_shape(slide, MSO_SHAPE.RECTANGLE,
                  cat3_x + Inches(0.15), cy, cat_w - Inches(0.30), comp_row_h,
                  fill_color=bg, line_color=GRAY, line_width=Pt(0.3))
        add_textbox(slide, cat3_x + Inches(0.20), cy, Inches(2.3), comp_row_h,
                    name, font_size=8, color=color, bold=is_titan,
                    align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)
        add_textbox(slide, cat3_x + Inches(2.50), cy, Inches(0.7), comp_row_h,
                    vol, font_size=7, color=GRAY,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_textbox(slide, cat3_x + Inches(3.20), cy, Inches(0.75), comp_row_h,
                    price, font_size=9, color=color, bold=True,
                    align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)

    # Вывод
    add_shape(slide, MSO_SHAPE.RECTANGLE,
              cat3_x + Inches(0.15), cat_y + cat_h - Inches(0.55),
              cat_w - Inches(0.30), Inches(0.45),
              fill_color=GREEN, line_color=None)
    add_textbox(slide, cat3_x + Inches(0.15), cat_y + cat_h - Inches(0.55),
                cat_w - Inches(0.30), Inches(0.45),
                'TITAN 3-й — дешевле премиум-аналогов', font_size=10, color=YELLOW, bold=True,
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # === ФУТЕР ===
    add_textbox(slide, Inches(0.4), Inches(7.05),
                SLIDE_W - Inches(0.8), Inches(0.35),
                'Empils  •  TITAN  •  Цены конкурентов — Все Инструменты.ру  •  Цены TITAN — расчёт от отгрузочной цены Эмпилс от 10.07.2026',
                font_size=8, color=GRAY, bold=False,
                align=PP_ALIGN.CENTER, italic=True)

    prs.save(OUTPUT)
    print(f'OK: {OUTPUT}')
    print(f'Слайдов: {len(prs.slides)}')


if __name__ == '__main__':
    main()
