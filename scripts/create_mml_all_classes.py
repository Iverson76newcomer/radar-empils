#!/usr/bin/env python3
"""
Слайд: MML TITAN — Было → Стало (все классы A, B, C, D).
Сводное сравнение с итогами по каждому классу.
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

RED = RGBColor(0xC6, 0x28, 0x28)
GREEN = RGBColor(0x16, 0xA3, 0x4A)
DARK = RGBColor(0x1A, 0x1F, 0x36)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
GRAY = RGBColor(0x9C, 0xA3, 0xAF)
LIGHT_GRAY = RGBColor(0xE5, 0xE7, 0xEB)
ORANGE = RGBColor(0xF4, 0x79, 0x20)
YELLOW = RGBColor(0xFB, 0xBF, 0x24)
DARK_BLUE = RGBColor(0x1E, 0x3A, 0x5F)
LIGHT_GREEN = RGBColor(0xDC, 0xFC, 0xE7)
LIGHT_RED = RGBColor(0xFE, 0xE2, 0xE2)
LIGHT_ORANGE = RGBColor(0xFF, 0xED, 0xD5)
LIGHT_BLUE = RGBColor(0xDB, 0xEA, 0xFE)

OUTPUT = '/home/z/my-project/download/MML_все_классы_было_стало.pptx'

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
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, Inches(0.60),
              fill_color=DARK_BLUE, line_color=None)
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0, Inches(0.60), SLIDE_W, Inches(0.04),
              fill_color=YELLOW, line_color=None)
    add_textbox(slide, Inches(0.4), Inches(0.07), SLIDE_W - Inches(0.8), Inches(0.46),
                'MML TITAN — БЫЛО → СТАЛО (все классы ТТ)',
                font_size=20, color=WHITE, bold=True,
                align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)
    add_textbox(slide, SLIDE_W - Inches(5.0), Inches(0.07),
                Inches(4.6), Inches(0.46),
                '30 SKU → 46 SKU (+16) · на основе продаж',
                font_size=11, color=YELLOW, bold=False,
                align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)

    # === 4 КАРТОЧКИ КЛАССОВ (в ряд) ===
    classes = [
        ('A', 'Федеральная DIY', '16', '21', '+5', RED,
         [
             '3в1 зелёный RAL 6005 (0.75л)',
             'Быстросохн. серый (0.75л)',
             'Быстросохн. корич. RAL 8017 (0.75л)',
             'Быстросохн. белый RAL 9010 (0.75л)',
             'Молотковая медный (0.8кг)',
         ]),
        ('B', 'Локальная сеть', '12', '16', '+4', ORANGE,
         [
             '3в1 зелёный RAL 6005 (0.75л)',
             'Быстросохн. серый (0.75л)',
             'Быстросохн. корич. RAL 8017 (0.75л)',
             'Молотковая серебристый (0.8кг)',
         ]),
        ('C', 'Супермаркет', '2', '7', '+5', GREEN,
         [
             '3в1 чёрный RAL 9005 (0.75л)',
             '3в1 корич. RAL 8017 (0.75л)',
             '3в1 серый RAL 7004 (0.75л)',
             'Быстросохн. чёрный RAL 9005 (0.75л)',
             'Быстросохн. серый (0.75л)',
         ]),
        ('D', 'Хозмаг', '0', '2', '+2', DARK_BLUE,
         [
             '3в1 чёрный RAL 9005 (0.75л)',
             '3в1 корич. RAL 8017 (0.75л)',
         ]),
    ]

    card_y = Inches(0.80)
    card_w = Inches(3.10)
    card_h = Inches(3.85)
    gap = Inches(0.12)

    for i, (cls, channel, was, now, diff, color, additions) in enumerate(classes):
        cx = Inches(0.4) + (card_w + gap) * i

        # Карточка
        add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
                  cx, card_y, card_w, card_h,
                  fill_color=WHITE, line_color=color, line_width=Pt(2))

        # Шапка
        add_shape(slide, MSO_SHAPE.RECTANGLE,
                  cx, card_y, card_w, Inches(0.55),
                  fill_color=color, line_color=None)
        add_textbox(slide, cx, card_y + Inches(0.02),
                    card_w, Inches(0.30),
                    f'КЛАСС {cls}', font_size=12, color=WHITE, bold=True,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_textbox(slide, cx, card_y + Inches(0.28),
                    card_w, Inches(0.25),
                    channel, font_size=8, color=YELLOW if color != DARK_BLUE else WHITE,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

        # Было → Стало
        add_multi_text(slide, cx, card_y + Inches(0.65),
                       card_w, Inches(0.55),
                       [
                           (f'{was}  →  {now}', 24, DARK, True),
                       ])

        # Изменение
        add_shape(slide, MSO_SHAPE.RECTANGLE,
                  cx + Inches(0.50), card_y + Inches(1.20),
                  card_w - Inches(1.0), Inches(0.32),
                  fill_color=color, line_color=None)
        add_textbox(slide, cx + Inches(0.50), card_y + Inches(1.20),
                    card_w - Inches(1.0), Inches(0.32),
                    f'{diff} SKU', font_size=14, color=WHITE, bold=True,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

        # Что добавляем
        add_textbox(slide, cx + Inches(0.15), card_y + Inches(1.65),
                    card_w - Inches(0.30), Inches(0.25),
                    'ЧТО ДОБАВИТЬ:', font_size=8, color=color, bold=True,
                    align=PP_ALIGN.LEFT)

        for j, addition in enumerate(additions):
            ay = card_y + Inches(1.95) + Inches(0.30) * j
            # Маркер
            add_shape(slide, MSO_SHAPE.OVAL,
                      cx + Inches(0.20), ay + Inches(0.04),
                      Inches(0.12), Inches(0.12),
                      fill_color=GREEN, line_color=None)
            add_textbox(slide, cx + Inches(0.38), ay,
                        card_w - Inches(0.50), Inches(0.28),
                        addition, font_size=7, color=DARK,
                        align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)

    # === НИЖНЯЯ ЧАСТЬ — сводная таблица по категориям ===
    tbl_y = Inches(4.85)
    add_textbox(slide, Inches(0.4), tbl_y,
                SLIDE_W - Inches(0.8), Inches(0.30),
                'СВОДКА ПО КАТЕГОРИЯМ',
                font_size=11, color=DARK_BLUE, bold=True, align=PP_ALIGN.LEFT)

    # Таблица
    headers = ['Категория', 'Было (A)', 'Стало (A)', 'Было (B)', 'Стало (B)', 'Было (C)', 'Стало (C)', 'Было (D)', 'Стало (D)']
    col_w = Inches(2.8)
    small_w = Inches(1.18)
    th_y = Inches(5.20)
    th_h = Inches(0.38)

    cx = Inches(0.4)
    # Первая колонка шире
    add_shape(slide, MSO_SHAPE.RECTANGLE,
              cx, th_y, col_w, th_h,
              fill_color=DARK_BLUE, line_color=WHITE, line_width=Pt(1))
    add_textbox(slide, cx, th_y, col_w, th_h,
                'Категория', font_size=9, color=WHITE, bold=True,
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    cx += col_w

    for i, h in enumerate(headers[1:]):
        bg = [RED, RED, ORANGE, ORANGE, GREEN, GREEN, DARK_BLUE, DARK_BLUE][i]
        add_shape(slide, MSO_SHAPE.RECTANGLE,
                  cx, th_y, small_w, th_h,
                  fill_color=bg, line_color=WHITE, line_width=Pt(1))
        add_textbox(slide, cx, th_y, small_w, th_h,
                    h, font_size=8, color=WHITE, bold=True,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        cx += small_w

    # Данные
    data = [
        ('3в1 по ржавчине, 0.75 л', '5', '5', '5', '5', '0', '3', '0', '2'),
        ('3в1 по ржавчине, 2 л', '5', '5', '4', '4', '0', '0', '0', '0'),
        ('Быстросохнущая, 0.75 л', '1', '1', '1', '1', '0', '2', '0', '0'),
        ('Быстросохнущая, 2.2 л', '1', '1', '0', '0', '0', '0', '0', '0'),
        ('Быстросохнущая, 0.75 л (нов.)', '0', '3', '0', '2', '0', '2', '0', '0'),
        ('Молотковая, 0.8 кг', '3', '3', '2', '3', '2', '2', '0', '0'),
        ('Молотковая, 2.5 кг', '1', '1', '0', '0', '0', '0', '0', '0'),
    ]

    row_h = Inches(0.22)
    for ri, row in enumerate(data):
        ry = th_y + th_h + row_h * ri
        bg = LIGHT_GRAY if ri % 2 == 0 else WHITE
        cx = Inches(0.4)
        # Категория
        add_shape(slide, MSO_SHAPE.RECTANGLE,
                  cx, ry, col_w, row_h,
                  fill_color=bg, line_color=GRAY, line_width=Pt(0.3))
        add_textbox(slide, cx + Inches(0.05), ry,
                    col_w - Inches(0.10), row_h,
                    row[0], font_size=8, color=DARK, bold=True,
                    align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)
        cx += col_w

        for ci, val in enumerate(row[1:]):
            # Подсветка новых
            is_new = val != '0' and ci % 2 == 1 and row[(ci+1)//2*2-1] != row[(ci+1)//2*2]
            # На самом деле проверяем: если "стало" > "было"
            idx_was = ci - 1 if ci % 2 == 1 else ci
            cell_bg = LIGHT_GREEN if (ci % 2 == 1 and val != '0' and val != row[ci]) else bg
            val_color = RED if val == '0' else (GREEN if ci % 2 == 1 and val != row[ci-1 if ci > 0 else ci] else DARK)

            add_shape(slide, MSO_SHAPE.RECTANGLE,
                      cx, ry, small_w, row_h,
                      fill_color=cell_bg, line_color=GRAY, line_width=Pt(0.3))
            add_textbox(slide, cx, ry, small_w, row_h,
                        str(val), font_size=8, color=val_color,
                        bold=(ci % 2 == 1),
                        align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
            cx += small_w

    # ИТОГО
    ry = th_y + th_h + row_h * len(data)
    cx = Inches(0.4)
    add_shape(slide, MSO_SHAPE.RECTANGLE,
              cx, ry, col_w, Inches(0.30),
              fill_color=DARK_BLUE, line_color=WHITE, line_width=Pt(1))
    add_textbox(slide, cx, ry, col_w, Inches(0.30),
                'ИТОГО SKU', font_size=10, color=WHITE, bold=True,
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    cx += col_w

    totals = [('16', RED), ('21', RED), ('12', ORANGE), ('16', ORANGE), ('2', GREEN), ('7', GREEN), ('0', DARK_BLUE), ('2', DARK_BLUE)]
    for val, color in totals:
        add_shape(slide, MSO_SHAPE.RECTANGLE,
                  cx, ry, small_w, Inches(0.30),
                  fill_color=color, line_color=WHITE, line_width=Pt(1))
        add_textbox(slide, cx, ry, small_w, Inches(0.30),
                    val, font_size=12, color=WHITE, bold=True,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        cx += small_w

    # === ФИНАЛЬНАЯ СТРОКА ===
    final_y = ry + Inches(0.35)
    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
              Inches(0.4), final_y, SLIDE_W - Inches(0.8), Inches(0.45),
              fill_color=GREEN, line_color=None)
    add_textbox(slide, Inches(0.4), final_y, SLIDE_W - Inches(0.8), Inches(0.45),
                'ИТОГО: 30 SKU → 46 SKU (+16)  ·  Все добавления основаны на реальных продажах  ·  Источник: «Титан цвета», 2026',
                font_size=12, color=YELLOW, bold=True,
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # Футер
    add_textbox(slide, Inches(0.4), Inches(7.20),
                SLIDE_W - Inches(0.8), Inches(0.25),
                'Empils  •  TITAN  •  MML — Minimal Must List  •  Красишь 1 раз — служит 10 лет',
                font_size=8, color=GRAY, italic=True, align=PP_ALIGN.CENTER)

    prs.save(OUTPUT)
    print(f'OK: {OUTPUT}')
    print(f'Слайдов: {len(prs.slides)}')


if __name__ == '__main__':
    main()
