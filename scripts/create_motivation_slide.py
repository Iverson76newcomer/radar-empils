#!/usr/bin/env python3
"""
Мотивирующий слайд для дистрибьютора — расчёт прибыли от TITAN.
Добавляется в дорожную карту как отдельный PPTX (1 слайд).
Стиль совпадает с дорожной картой TITAN.
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
import os

# === COLORS (как в дорожной карте) ===
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
BLUE = RGBColor(0x25, 0x63, 0xEB)
DARK_BLUE = RGBColor(0x1E, 0x3A, 0x5F)
LIGHT_GREEN = RGBColor(0xDC, 0xFC, 0xE7)
LIGHT_YELLOW = RGBColor(0xFE, 0xF9, 0xC3)
LIGHT_ORANGE = RGBColor(0xFF, 0xED, 0xD5)
LIGHT_RED = RGBColor(0xFE, 0xE2, 0xE2)
LIGHT_BLUE = RGBColor(0xDB, 0xEA, 0xFE)

OUTPUT = '/home/z/my-project/download/Слайд_мотивация_расчёт_прибыли.pptx'

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


# ==============================================================
# СОЗДАНИЕ ПРЕЗЕНТАЦИИ
# ==============================================================

def main():
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H

    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # Белый фон
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H,
              fill_color=WHITE, line_color=None)

    # === ШАПКА ===
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, Inches(0.65),
              fill_color=RED, line_color=None)
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0, Inches(0.65), SLIDE_W, Inches(0.04),
              fill_color=YELLOW, line_color=None)
    add_textbox(slide, Inches(0.4), Inches(0.08), SLIDE_W - Inches(0.8), Inches(0.5),
                'СКОЛЬКО ВЫ ЗАРАБАТЫВАЕТЕ НА TITAN',
                font_size=22, color=WHITE, bold=True,
                align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)
    add_textbox(slide, SLIDE_W - Inches(5.0), Inches(0.08),
                Inches(4.6), Inches(0.5),
                'Пример: дистрибьютор со 100 точками',
                font_size=12, color=WHITE, bold=False,
                align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)

    # === ЛЕВАЯ ЧАСТЬ — исходные данные ===

    # (1) Структура точек (РЕАЛЬНЫЙ MML)
    add_textbox(slide, Inches(0.4), Inches(0.85),
                Inches(6.0), Inches(0.35),
                '1. СТРУКТУРА ВАШИХ 100 ТОЧЕК (реальный MML)',
                font_size=12, color=DARK_BLUE, bold=True, align=PP_ALIGN.LEFT)

    points = [
        ('Класс A — Федеральная DIY', '5 точек', '16 SKU × 6 шт = 480 банок', RED),
        ('Класс B — Локальная сеть', '10 точек', '12 SKU × 4 шт = 480 банок', ORANGE),
        ('Класс C — Супермаркет', '15 точек', '2 SKU × 3 шт = 90 банок', GREEN),
        ('Класс D — Хозмаг', '70 точек', '0 SKU × 2 шт = 0 банок', DARK_BLUE),
    ]
    p_y = Inches(1.20)
    p_h = Inches(0.42)
    for i, (name, count, formula, color) in enumerate(points):
        cy = p_y + p_h * i
        bg = LIGHT_GRAY if i % 2 == 0 else WHITE
        add_shape(slide, MSO_SHAPE.RECTANGLE,
                  Inches(0.4), cy, Inches(6.3), p_h,
                  fill_color=bg, line_color=GRAY, line_width=Pt(0.3))
        add_textbox(slide, Inches(0.5), cy, Inches(2.7), p_h,
                    name, font_size=10, color=DARK, bold=True,
                    align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)
        add_shape(slide, MSO_SHAPE.RECTANGLE,
                  Inches(3.2), cy, Inches(1.0), p_h,
                  fill_color=color, line_color=None)
        add_textbox(slide, Inches(3.2), cy, Inches(1.0), p_h,
                    count, font_size=10, color=WHITE, bold=True,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_textbox(slide, Inches(4.3), cy, Inches(2.3), p_h,
                    formula, font_size=9, color=GRAY, bold=False,
                    align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE, italic=True)

    # Итого банок
    add_shape(slide, MSO_SHAPE.RECTANGLE,
              Inches(0.4), Inches(2.95), Inches(6.3), Inches(0.45),
              fill_color=DARK, line_color=None)
    add_textbox(slide, Inches(0.5), Inches(2.95), Inches(4.0), Inches(0.45),
                'ИТОГО БАНОК НА ПОЛКЕ:',
                font_size=12, color=WHITE, bold=True,
                align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)
    add_textbox(slide, Inches(4.5), Inches(2.95), Inches(2.1), Inches(0.45),
                '1 050 шт',
                font_size=18, color=YELLOW, bold=True,
                align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)

    # (2) Цепочка цены
    add_textbox(slide, Inches(0.4), Inches(3.55),
                Inches(6.0), Inches(0.35),
                '2. ЦЕПОВКА ЦЕНЫ (по 16 SKU)',
                font_size=12, color=DARK_BLUE, bold=True, align=PP_ALIGN.LEFT)

    steps = [
        ('Отгрузочная цена Эмпилс', 'от 494 до 2 076 ₽', DARK_BLUE, LIGHT_GRAY),
        ('+25% маржа дистрибьютора', '617–2 595 ₽', ORANGE, LIGHT_ORANGE),
        ('+25% маржа розницы (цена на полке)', '772–3 244 ₽', GREEN, LIGHT_GREEN),
    ]
    s_y = Inches(3.90)
    s_h = Inches(0.42)
    for i, (label, price, color, bg) in enumerate(steps):
        cy = s_y + s_h * i
        add_shape(slide, MSO_SHAPE.RECTANGLE,
                  Inches(0.4), cy, Inches(6.3), s_h,
                  fill_color=bg, line_color=color, line_width=Pt(0.5))
        add_textbox(slide, Inches(0.5), cy, Inches(4.3), s_h,
                    label, font_size=10, color=DARK, bold=True,
                    align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)
        add_textbox(slide, Inches(4.8), cy, Inches(1.8), s_h,
                    price, font_size=12, color=color, bold=True,
                    align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)

    # (3) Валовая прибыль дистрибьютора (2 карточки внизу слева)
    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
              Inches(0.4), Inches(5.30), Inches(6.3), Inches(1.75),
              fill_color=LIGHT_GREEN, line_color=GREEN, line_width=Pt(2))
    add_textbox(slide, Inches(0.5), Inches(5.35),
                Inches(6.1), Inches(0.4),
                '3. ВАЛОВАЯ ПРИБЫЛЬ ДИСТРИБЬЮТОРА (за 1 оборот)',
                font_size=11, color=GREEN, bold=True, align=PP_ALIGN.LEFT)

    # 2 колонки — месяц и год
    add_shape(slide, MSO_SHAPE.RECTANGLE,
              Inches(0.6), Inches(5.80), Inches(2.9), Inches(1.15),
              fill_color=WHITE, line_color=GREEN, line_width=Pt(1.5))
    add_multi_text(slide, Inches(0.6), Inches(5.80), Inches(2.9), Inches(1.15),
                   [
                       ('В МЕСЯЦ', 10, GREEN, True),
                       ('224 747 ₽', 22, GREEN, True),
                       ('при 1 050 банок\nоборот 1 раз/мес', 8, GRAY, False),
                   ])

    add_shape(slide, MSO_SHAPE.RECTANGLE,
              Inches(3.65), Inches(5.80), Inches(2.9), Inches(1.15),
              fill_color=GREEN, line_color=None)
    add_multi_text(slide, Inches(3.65), Inches(5.80), Inches(2.9), Inches(1.15),
                   [
                       ('В ГОД', 10, YELLOW, True),
                       ('2 696 962 ₽', 22, WHITE, True),
                       ('× 12 месяцев\nваловая прибыль', 8, LIGHT_GRAY, False),
                   ])

    # === ПРАВАЯ ЧАСТЬ — главный мотиватор: прибыль дистрибьютора + рост ===

    # Заголовок правой части
    add_textbox(slide, Inches(7.0), Inches(0.85),
                Inches(6.0), Inches(0.35),
                'ВАША ПРИБЫЛЬ НА TITAN',
                font_size=12, color=DARK_BLUE, bold=True, align=PP_ALIGN.LEFT)

    # === 2 крупные карточки: МЕСЯЦ и ГОД ===
    card_y = Inches(1.30)
    card_h = Inches(1.80)
    card_w = Inches(2.90)

    # Карточка "В МЕСЯЦ"
    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
              Inches(7.0), card_y, card_w, card_h,
              fill_color=LIGHT_GREEN, line_color=GREEN, line_width=Pt(2.5))
    add_textbox(slide, Inches(7.0), card_y + Inches(0.15),
                card_w, Inches(0.4),
                'В МЕСЯЦ', font_size=14, color=GREEN, bold=True,
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_textbox(slide, Inches(7.0), card_y + Inches(0.55),
                card_w, Inches(0.8),
                '224 747 ₽', font_size=32, color=GREEN, bold=True,
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_textbox(slide, Inches(7.0), card_y + Inches(1.35),
                card_w, Inches(0.4),
                'при 1 050 банок\nоборот 1 раз/мес', font_size=10, color=GRAY,
                align=PP_ALIGN.CENTER, italic=True)

    # Карточка "В ГОД"
    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
              Inches(10.10), card_y, card_w, card_h,
              fill_color=GREEN, line_color=None)
    add_textbox(slide, Inches(10.10), card_y + Inches(0.15),
                card_w, Inches(0.4),
                'В ГОД', font_size=14, color=YELLOW, bold=True,
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_textbox(slide, Inches(10.10), card_y + Inches(0.55),
                card_w, Inches(0.8),
                '2 696 962 ₽', font_size=32, color=WHITE, bold=True,
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_textbox(slide, Inches(10.10), card_y + Inches(1.35),
                card_w, Inches(0.4),
                '× 12 месяцев\nваловая прибыль', font_size=10, color=LIGHT_GRAY,
                align=PP_ALIGN.CENTER, italic=True)

    # Стрелка вниз
    add_shape(slide, MSO_SHAPE.DOWN_ARROW,
              Inches(9.8), Inches(3.25), Inches(0.4), Inches(0.4),
              fill_color=GREEN, line_color=None)

    # === СЦЕНАРИЙ РОСТА — главная мотивация ===
    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
              Inches(7.0), Inches(3.85), Inches(6.0), Inches(3.20),
              fill_color=GREEN, line_color=None)

    # Заголовок
    add_textbox(slide, Inches(7.0), Inches(3.95),
                Inches(6.0), Inches(0.5),
                '🚀  СЦЕНАРИЙ РОСТА +30% К 12-МУ МЕСЯЦУ',
                font_size=14, color=YELLOW, bold=True,
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # Подзаголовок
    add_textbox(slide, Inches(7.0), Inches(4.45),
                Inches(6.0), Inches(0.35),
                'при выполнении программы TITAN (обучение + POS + планограмма)',
                font_size=9, color=WHITE, bold=False,
                align=PP_ALIGN.CENTER, italic=True)

    # 3 столбца сравнения (цифры реального MML +30%)
    compare = [
        ('СЕЙЧАС', '1 050', 'банок/мес', '224 747 ₽', 'прибыль/мес', WHITE),
        ('+30%', '+315', 'банок', '+67 424 ₽', 'рост/мес', YELLOW),
        ('ЧЕРЕЗ ГОД', '1 365', 'банок/мес', '292 171 ₽', 'прибыль/мес', WHITE),
    ]
    cw = Inches(1.85)
    for i, (title, val1, unit1, val2, unit2, val_color) in enumerate(compare):
        cx = Inches(7.15) + cw * i + Inches(0.10) * i
        # Карточка
        bg = RGBColor(0x0B, 0x6E, 0x31) if i == 1 else RGBColor(0x16, 0xA3, 0x4A)
        add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
                  cx, Inches(4.90), cw, Inches(1.50),
                  fill_color=bg, line_color=None)

        # Заголовок карточки
        add_textbox(slide, cx, Inches(4.95), cw, Inches(0.30),
                    title, font_size=10, color=YELLOW if i == 1 else WHITE, bold=True,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

        # Большая цифра (объём)
        add_textbox(slide, cx, Inches(5.25), cw, Inches(0.40),
                    val1, font_size=20, color=val_color, bold=True,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

        # Единицы
        add_textbox(slide, cx, Inches(5.65), cw, Inches(0.20),
                    unit1, font_size=8, color=RGBColor(0xCB, 0xD5, 0xE1),
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

        # Прибыль
        add_textbox(slide, cx, Inches(5.85), cw, Inches(0.30),
                    val2, font_size=12, color=val_color, bold=True,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

        # Единицы прибыли
        add_textbox(slide, cx, Inches(6.15), cw, Inches(0.20),
                    unit2, font_size=8, color=RGBColor(0xCB, 0xD5, 0xE1),
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # Главная мотивирующая строка (внизу)
    add_shape(slide, MSO_SHAPE.RECTANGLE,
              Inches(7.15), Inches(6.50), Inches(5.7), Inches(0.40),
              fill_color=YELLOW, line_color=None)
    add_textbox(slide, Inches(7.15), Inches(6.50), Inches(5.7), Inches(0.40),
                '💰  +809 088 ₽/год — дополнительно к текущей прибыли',
                font_size=13, color=DARK_RED, bold=True,
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # Финальная мотивирующая фраза
    add_textbox(slide, Inches(7.0), Inches(6.95),
                Inches(6.0), Inches(0.35),
                'Красишь 1 раз — служит 10 лет  •  Партнёрство, которое окупается за 4 месяца',
                font_size=10, color=GREEN, bold=True,
                align=PP_ALIGN.CENTER, italic=True)


    # === ФУТЕР ===
    add_textbox(slide, Inches(0.4), Inches(7.15),
                Inches(6.0), Inches(0.30),
                'Empils  •  TITAN  •  Расчёт по реальному MML, отгрузочные цены Эмпилс от 10.07.2026',
                font_size=8, color=GRAY, bold=False,
                align=PP_ALIGN.LEFT, italic=True)

    # ==============================================================
    # СЛАЙД 2: ДЕТАЛИЗАЦИЯ ПО 16 SKU
    # ==============================================================
    slide2 = prs.slides.add_slide(prs.slide_layouts[6])

    add_shape(slide2, MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H,
              fill_color=WHITE, line_color=None)

    # Шапка
    add_shape(slide2, MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, Inches(0.65),
              fill_color=GREEN, line_color=None)
    add_shape(slide2, MSO_SHAPE.RECTANGLE, 0, Inches(0.65), SLIDE_W, Inches(0.04),
              fill_color=YELLOW, line_color=None)
    add_textbox(slide2, Inches(0.4), Inches(0.08), SLIDE_W - Inches(0.8), Inches(0.5),
                'ДЕТАЛЬНЫЙ РАСЧЁТ ПО 16 SKU TITAN',
                font_size=22, color=WHITE, bold=True,
                align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)
    add_textbox(slide2, SLIDE_W - Inches(5.0), Inches(0.08),
                Inches(4.6), Inches(0.5),
                'Прозрачный расчёт прибыли',
                font_size=12, color=WHITE, bold=False,
                align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)

    # Данные по 16 SKU — РЕАЛЬНЫЙ MML (как сейчас)
    sku_data = [
        # (код, короткое название, категория, классы, цена, шт на полке, выручка, прибыль)
        ('87289', '3в1 чёрный 0.75 л', '3в1 по ржавчине', 'AB', 493.55, 70, 34548, 8637),
        ('87292', '3в1 коричневый 0.75 л', '3в1 по ржавчине', 'AB', 493.55, 70, 34548, 8637),
        ('87282', '3в1 серый 2 л', '3в1 по ржавчине', 'AB', 1273.35, 70, 89134, 22284),
        ('87287', '3в1 графит 2 л', '3в1 по ржавчине', 'AB', 1273.35, 70, 89134, 22284),
        ('87290', '3в1 чёрный 2 л', '3в1 по ржавчине', 'AB', 1273.35, 70, 89134, 22284),
        ('87293', '3в1 коричневый 2 л', '3в1 по ржавчине', 'AB', 1273.35, 70, 89134, 22284),
        ('87281', '3в1 серый 0.75 л', '3в1 по ржавчине', 'AB', 493.55, 70, 34548, 8637),
        ('87269', '3в1 белый 0.75 л', '3в1 по ржавчине', 'AB', 540.73, 70, 37851, 9463),
        ('87270', '3в1 белый 2 л', '3в1 по ржавчине', 'A', 1364.01, 30, 40920, 10230),
        ('87286', '3в1 графит 0.75 л', '3в1 по ржавчине', 'AB', 493.55, 70, 34548, 8637),
        ('77636', 'Молотковая чёрная 0.8 кг', 'Молотковая', 'ABC', 701.37, 115, 80657, 20164),
        ('77660', 'Молотковая коричневая 0.8 кг', 'Молотковая', 'ABC', 701.37, 115, 80657, 20164),
        ('77637', 'Молотковая чёрная 2.5 кг', 'Молотковая', 'A', 2075.61, 30, 62268, 15567),
        ('77642', 'Молотковая серебристая 0.8 кг', 'Молотковая', 'A', 701.37, 30, 21041, 5260),
        ('83777', 'Быстросохнущая 0.75 л', 'Быстросохнущая', 'AB', 518.22, 70, 36275, 9069),
        ('83779', 'Быстросохнущая 2.2 л', 'Быстросохнущая', 'A', 1486.13, 30, 44584, 11146),
    ]

    # Заголовок таблицы
    headers = ['Код', 'Продукт', 'Категория', 'Кл.', 'Цена Эмпилс', 'Шт', 'Выручка', 'Прибыль 25%']
    col_widths = [Inches(0.8), Inches(3.0), Inches(2.0), Inches(0.6), Inches(1.5), Inches(0.9), Inches(1.7), Inches(1.7)]
    table_y = Inches(0.85)
    header_h = Inches(0.35)

    cx = Inches(0.3)
    for i, h in enumerate(headers):
        add_shape(slide2, MSO_SHAPE.RECTANGLE,
                  cx, table_y, col_widths[i], header_h,
                  fill_color=DARK, line_color=WHITE, line_width=Pt(1))
        add_textbox(slide2, cx, table_y, col_widths[i], header_h,
                    h, font_size=9, color=WHITE, bold=True,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        cx += col_widths[i]

    # Данные
    row_h = Inches(0.26)
    cat_colors = {
        '3в1 по ржавчине': LIGHT_RED,
        'Молотковая': LIGHT_ORANGE,
        'Быстросохнущая': LIGHT_BLUE,
    }

    for ri, (code, name, cat, classes, price, units, revenue, profit) in enumerate(sku_data):
        cy = table_y + header_h + row_h * ri
        bg = cat_colors.get(cat, LIGHT_GRAY)
        cx = Inches(0.3)
        # Код
        add_shape(slide2, MSO_SHAPE.RECTANGLE,
                  cx, cy, col_widths[0], row_h,
                  fill_color=bg, line_color=GRAY, line_width=Pt(0.3))
        add_textbox(slide2, cx, cy, col_widths[0], row_h,
                    code, font_size=8, color=DARK, bold=True,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        cx += col_widths[0]
        # Продукт
        add_shape(slide2, MSO_SHAPE.RECTANGLE,
                  cx, cy, col_widths[1], row_h,
                  fill_color=bg, line_color=GRAY, line_width=Pt(0.3))
        add_textbox(slide2, cx + Inches(0.05), cy, col_widths[1] - Inches(0.10), row_h,
                    name, font_size=8, color=DARK, bold=False,
                    align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)
        cx += col_widths[1]
        # Категория
        add_shape(slide2, MSO_SHAPE.RECTANGLE,
                  cx, cy, col_widths[2], row_h,
                  fill_color=bg, line_color=GRAY, line_width=Pt(0.3))
        add_textbox(slide2, cx + Inches(0.05), cy, col_widths[2] - Inches(0.10), row_h,
                    cat, font_size=7, color=DARK, bold=False,
                    align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)
        cx += col_widths[2]
        # Классы
        add_shape(slide2, MSO_SHAPE.RECTANGLE,
                  cx, cy, col_widths[3], row_h,
                  fill_color=bg, line_color=GRAY, line_width=Pt(0.3))
        add_textbox(slide2, cx, cy, col_widths[3], row_h,
                    classes, font_size=9, color=DARK, bold=True,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        cx += col_widths[3]
        # Цена
        add_shape(slide2, MSO_SHAPE.RECTANGLE,
                  cx, cy, col_widths[4], row_h,
                  fill_color=bg, line_color=GRAY, line_width=Pt(0.3))
        add_textbox(slide2, cx, cy, col_widths[4], row_h,
                    f'{price:.0f} ₽', font_size=8, color=DARK, bold=False,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        cx += col_widths[4]
        # Шт на полке
        add_shape(slide2, MSO_SHAPE.RECTANGLE,
                  cx, cy, col_widths[5], row_h,
                  fill_color=bg, line_color=GRAY, line_width=Pt(0.3))
        add_textbox(slide2, cx, cy, col_widths[5], row_h,
                    str(units), font_size=8, color=DARK, bold=True,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        cx += col_widths[5]
        # Выручка
        add_shape(slide2, MSO_SHAPE.RECTANGLE,
                  cx, cy, col_widths[6], row_h,
                  fill_color=bg, line_color=GRAY, line_width=Pt(0.3))
        add_textbox(slide2, cx, cy, col_widths[6], row_h,
                    f'{revenue:,} ₽'.replace(',', ' '), font_size=8, color=DARK, bold=False,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        cx += col_widths[6]
        # Прибыль
        add_shape(slide2, MSO_SHAPE.RECTANGLE,
                  cx, cy, col_widths[7], row_h,
                  fill_color=bg, line_color=GRAY, line_width=Pt(0.3))
        add_textbox(slide2, cx, cy, col_widths[7], row_h,
                    f'{profit:,} ₽'.replace(',', ' '), font_size=8, color=GREEN, bold=True,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # ИТОГО
    total_y = table_y + header_h + row_h * len(sku_data)
    cx = Inches(0.3)
    # Объединим первые 5 колонок для надписи ИТОГО
    add_shape(slide2, MSO_SHAPE.RECTANGLE,
              cx, total_y, sum(col_widths[:5], Inches(0)), row_h,
              fill_color=DARK, line_color=WHITE, line_width=Pt(1))
    add_textbox(slide2, cx + Inches(0.5), total_y, sum(col_widths[:5], Inches(0)) - Inches(1), row_h,
                'ИТОГО',
                font_size=11, color=WHITE, bold=True,
                align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)
    cx += sum(col_widths[:5], Inches(0))
    # Шт на полке
    add_shape(slide2, MSO_SHAPE.RECTANGLE,
              cx, total_y, col_widths[5], row_h,
              fill_color=DARK, line_color=WHITE, line_width=Pt(1))
    add_textbox(slide2, cx, total_y, col_widths[5], row_h,
                '1 050', font_size=10, color=YELLOW, bold=True,
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    cx += col_widths[5]
    # Выручка
    add_shape(slide2, MSO_SHAPE.RECTANGLE,
              cx, total_y, col_widths[6], row_h,
              fill_color=DARK, line_color=WHITE, line_width=Pt(1))
    add_textbox(slide2, cx, total_y, col_widths[6], row_h,
                '898 987 ₽', font_size=9, color=WHITE, bold=True,
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    cx += col_widths[6]
    # Прибыль
    add_shape(slide2, MSO_SHAPE.RECTANGLE,
              cx, total_y, col_widths[7], row_h,
              fill_color=GREEN, line_color=WHITE, line_width=Pt(1))
    add_textbox(slide2, cx, total_y, col_widths[7], row_h,
                '224 747 ₽', font_size=10, color=YELLOW, bold=True,
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # === ИТОГИ ПО КАТЕГОРИЯМ ===
    cat_y = total_y + row_h + Inches(0.15)
    add_textbox(slide2, Inches(0.4), cat_y,
                Inches(12.5), Inches(0.35),
                'СТРУКТУРА ПРИБЫЛИ ПО КАТЕГОРИЯМ',
                font_size=11, color=DARK, bold=True, align=PP_ALIGN.LEFT)

    cat_data = [
        ('3в1 по ржавчине', '660 шт', '143 376 ₽', '63.8%', RED, LIGHT_RED),
        ('Молотковая', '290 шт', '61 156 ₽', '27.2%', ORANGE, LIGHT_ORANGE),
        ('Быстросохнущая', '100 шт', '20 215 ₽', '9.0%', BLUE, LIGHT_BLUE),
    ]
    cat_card_y = cat_y + Inches(0.35)
    cat_card_w = Inches(4.0)
    cat_card_h = Inches(1.10)

    for i, (cat_name, units, profit, share, color, bg) in enumerate(cat_data):
        cx = Inches(0.4) + (cat_card_w + Inches(0.15)) * i
        add_shape(slide2, MSO_SHAPE.ROUNDED_RECTANGLE,
                  cx, cat_card_y, cat_card_w, cat_card_h,
                  fill_color=bg, line_color=color, line_width=Pt(2))
        # Шапка
        add_shape(slide2, MSO_SHAPE.RECTANGLE,
                  cx, cat_card_y, cat_card_w, Inches(0.32),
                  fill_color=color, line_color=None)
        add_textbox(slide2, cx, cat_card_y, cat_card_w, Inches(0.32),
                    cat_name, font_size=11, color=WHITE, bold=True,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        # Штуки + прибыль + доля в одну строку
        add_multi_text(slide2, cx + Inches(0.1), cat_card_y + Inches(0.40),
                       cat_card_w - Inches(0.2), Inches(0.65),
                       [
                           (f'{units}  →  {profit}  ({share})', 12, color, True),
                           ('доля прибыли', 8, GRAY, False),
                       ])

    # Футер
    add_textbox(slide2, Inches(0.4), Inches(6.95),
                SLIDE_W - Inches(0.8), Inches(0.30),
                'Empils  •  TITAN  •  Расчёт по отгрузочным ценам Эмпилс от 10.07.2026  •  Глубина полки: A=6, B=4, C=3, D=2 шт',
                font_size=8, color=GRAY, bold=False,
                align=PP_ALIGN.CENTER, italic=True)

    # Сохранение
    prs.save(OUTPUT)
    print(f'OK: {OUTPUT}')
    print(f'Слайдов: {len(prs.slides)}')


if __name__ == '__main__':
    main()
