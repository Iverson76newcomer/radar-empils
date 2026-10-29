#!/usr/bin/env python3
"""
Слайд: Сравнение текущего и рекомендуемого MML.
8 рекомендаций на основе реальных продаж.
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

RED = RGBColor(0xC6, 0x28, 0x28)
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

OUTPUT = '/home/z/my-project/download/Слайд_MML_сравнение.pptx'

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
              fill_color=DARK_BLUE, line_color=None)
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0, Inches(0.65), SLIDE_W, Inches(0.04),
              fill_color=YELLOW, line_color=None)
    add_textbox(slide, Inches(0.4), Inches(0.08), SLIDE_W - Inches(0.8), Inches(0.5),
                'MML: КАК БЫЛО → КАК ПРЕДЛАГАЕТСЯ',
                font_size=22, color=WHITE, bold=True,
                align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)
    add_textbox(slide, SLIDE_W - Inches(5.0), Inches(0.08),
                Inches(4.6), Inches(0.5),
                '8 рекомендаций на основе продаж',
                font_size=11, color=YELLOW, bold=False,
                align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)

    # === ВЕРХНЯЯ СТРОКА — 4 КАРТОЧКИ КЛАССОВ ===
    classes = [
        ('A', 'Федеральная DIY', '16', '21', '+5', RED),
        ('B', 'Локальная сеть', '12', '16', '+4', ORANGE),
        ('C', 'Супермаркет', '2', '4', '+2', GREEN),
        ('D', 'Хозмаг', '0', '0', '0', DARK_BLUE),
    ]
    card_y = Inches(0.85)
    card_w = Inches(3.05)
    card_h = Inches(1.50)
    gap = Inches(0.15)

    for i, (cls, channel, was, now, diff, color) in enumerate(classes):
        cx = Inches(0.4) + (card_w + gap) * i

        # Карточка
        add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
                  cx, card_y, card_w, card_h,
                  fill_color=WHITE, line_color=color, line_width=Pt(2))
        # Шапка
        add_shape(slide, MSO_SHAPE.RECTANGLE,
                  cx, card_y, card_w, Inches(0.40),
                  fill_color=color, line_color=None)
        add_textbox(slide, cx, card_y, card_w, Inches(0.40),
                    f'КЛАСС {cls} — {channel}', font_size=10, color=WHITE, bold=True,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

        # Было → Стало
        add_multi_text(slide, cx, card_y + Inches(0.45),
                       card_w, Inches(0.55),
                       [
                           (f'{was}  →  {now}', 22, DARK, True),
                       ])
        # Изменение
        diff_color = GREEN if diff != '0' else GRAY
        add_textbox(slide, cx, card_y + Inches(1.05),
                    card_w, Inches(0.35),
                    f'{diff} SKU', font_size=14, color=diff_color, bold=True,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # === ЦЕНТРАЛЬНАЯ СТРЕЛКА ===
    add_shape(slide, MSO_SHAPE.DOWN_ARROW,
              Inches(6.4), Inches(2.45), Inches(0.5), Inches(0.35),
              fill_color=GREEN, line_color=None)

    # === ОСНОВНАЯ ТАБЛИЦА — 8 РЕКОМЕНДАЦИЙ ===
    add_textbox(slide, Inches(0.4), Inches(2.90),
                Inches(12.5), Inches(0.35),
                '8 РЕКОМЕНДАЦИЙ ПО КОРРЕКТИРОВКЕ MML — НА ОСНОВЕ РЕАЛЬНЫХ ПРОДАЖ',
                font_size=12, color=DARK_BLUE, bold=True, align=PP_ALIGN.LEFT)

    # Заголовок таблицы
    headers = ['№', 'Категория', 'Цвет / RAL', 'Объём', 'Вторичка\n(шт)', 'АКБ', 'Было\n(MML)', 'Стало\n(MML)']
    col_widths = [Inches(0.5), Inches(2.3), Inches(3.0), Inches(0.9), Inches(1.2), Inches(0.9), Inches(1.5), Inches(1.7)]
    table_y = Inches(3.30)
    header_h = Inches(0.40)

    cx = Inches(0.4)
    for i, h in enumerate(headers):
        add_shape(slide, MSO_SHAPE.RECTANGLE,
                  cx, table_y, col_widths[i], header_h,
                  fill_color=DARK, line_color=WHITE, line_width=Pt(1))
        add_textbox(slide, cx, table_y, col_widths[i], header_h,
                    h, font_size=9, color=WHITE, bold=True,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        cx += col_widths[i]

    # Данные
    recs = [
        (1, '3в1 по ржавчине', 'Чёрный RAL 9005', '0.75 л', '7 738', '393', 'A, B', 'A, B, C', LIGHT_RED),
        (2, '3в1 по ржавчине', 'Шоколад-коричневый RAL 8017', '0.75 л', '4 524', '287', 'A, B', 'A, B, C', LIGHT_RED),
        (3, '3в1 по ржавчине', 'Зелёный RAL 6005', '0.75 л', '1 893', '111', '—', 'A, B', LIGHT_RED),
        (4, 'Быстросохнущая', 'Серый', '0.75 л', '2 695', '199', '—', 'A, B', LIGHT_ORANGE),
        (5, 'Быстросохнущая', 'Шоколадно-кор. RAL 8017', '0.75 л', '2 232', '193', '—', 'A, B', LIGHT_ORANGE),
        (6, 'Быстросохнущая', 'Белый RAL 9010', '0.75 л', '1 507', '119', '—', 'A', LIGHT_ORANGE),
        (7, 'Молотковая', 'Серебристый', '0.8 кг', '2 394', '198', 'A', 'A, B', LIGHT_GRAY),
        (8, 'Молотковая', 'Медный', '0.8 кг', '1 162', '102', '—', 'A', LIGHT_GRAY),
    ]

    row_h = Inches(0.36)
    for ri, (num, cat, color_name, vol, sec_qty, akb, was, now, bg) in enumerate(recs):
        cy = table_y + header_h + row_h * ri
        cx = Inches(0.4)

        values = [num, cat, color_name, vol, sec_qty, akb, was, now]
        for ci, val in enumerate(values):
            add_shape(slide, MSO_SHAPE.RECTANGLE,
                      cx, cy, col_widths[ci], row_h,
                      fill_color=bg, line_color=GRAY, line_width=Pt(0.3))

            # Спеццвета
            if ci == 7:  # "Стало" — зелёным
                val_color = GREEN
                val_bold = True
            elif ci == 6:  # "Было" — красным если "—"
                val_color = RED if val == '—' else DARK
                val_bold = True
            elif ci == 4 and int(str(val).replace(' ', '')) > 1000:
                val_color = DARK
                val_bold = True
            else:
                val_color = DARK
                val_bold = (ci == 0 or ci == 1)

            add_textbox(slide, cx, cy, col_widths[ci], row_h,
                        str(val), font_size=9, color=val_color, bold=val_bold,
                        align=PP_ALIGN.CENTER if ci not in [1, 2] else PP_ALIGN.LEFT,
                        anchor=MSO_ANCHOR.MIDDLE)
            cx += col_widths[ci]

    # === ИТОГОВАЯ СТРОКА ===
    total_y = table_y + header_h + row_h * len(recs)
    cx = Inches(0.4)
    # Объединим первые 4 колонки
    add_shape(slide, MSO_SHAPE.RECTANGLE,
              cx, total_y, sum(col_widths[:4], Inches(0)), row_h,
              fill_color=DARK_BLUE, line_color=WHITE, line_width=Pt(1))
    add_textbox(slide, cx + Inches(0.5), total_y, sum(col_widths[:4], Inches(0)) - Inches(1), row_h,
                'ИТОГО: 8 рекомендаций на основе продаж → +11 SKU-позиций в MML',
                font_size=10, color=WHITE, bold=True,
                align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)
    cx += sum(col_widths[:4], Inches(0))

    # Вторичка (сумма)
    add_shape(slide, MSO_SHAPE.RECTANGLE,
              cx, total_y, col_widths[4], row_h,
              fill_color=DARK_BLUE, line_color=WHITE, line_width=Pt(1))
    add_textbox(slide, cx, total_y, col_widths[4], row_h,
                '22 645', font_size=9, color=YELLOW, bold=True,
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    cx += col_widths[4]

    # АКБ (сумма)
    add_shape(slide, MSO_SHAPE.RECTANGLE,
              cx, total_y, col_widths[5], row_h,
              fill_color=DARK_BLUE, line_color=WHITE, line_width=Pt(1))
    add_textbox(slide, cx, total_y, col_widths[5], row_h,
                '1 522', font_size=9, color=YELLOW, bold=True,
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    cx += col_widths[5]

    # Было
    add_shape(slide, MSO_SHAPE.RECTANGLE,
              cx, total_y, col_widths[6], row_h,
              fill_color=DARK_BLUE, line_color=WHITE, line_width=Pt(1))
    add_textbox(slide, cx, total_y, col_widths[6], row_h,
                '30 SKU', font_size=9, color=WHITE, bold=True,
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    cx += col_widths[6]

    # Стало
    add_shape(slide, MSO_SHAPE.RECTANGLE,
              cx, total_y, col_widths[7], row_h,
              fill_color=GREEN, line_color=WHITE, line_width=Pt(1))
    add_textbox(slide, cx, total_y, col_widths[7], row_h,
                '41 SKU', font_size=10, color=YELLOW, bold=True,
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # === НИЖНЯЯ ПОЛОСА — обоснование ===
    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
              Inches(0.4), Inches(6.75), SLIDE_W - Inches(0.8), Inches(0.65),
              fill_color=LIGHT_GREEN, line_color=GREEN, line_width=Pt(1))
    add_textbox(slide, Inches(0.4), Inches(6.75), SLIDE_W - Inches(0.8), Inches(0.65),
                '✓  Каждая рекомендация основана на реальных данных: вторичные продажи (дистр.→розница) + первичные (завод→дистр.) + АКБ (активные клиенты). Источник: «Аналитика Титан цвета», 2026.',
                font_size=10, color=DARK_GREEN, bold=True,
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # Футер
    add_textbox(slide, Inches(0.4), Inches(7.30),
                SLIDE_W - Inches(0.8), Inches(0.20),
                'Empils  •  TITAN  •  MML — Minimal Must List  •  Красишь 1 раз — служит 10 лет',
                font_size=8, color=GRAY, bold=False,
                align=PP_ALIGN.CENTER, italic=True)

    prs.save(OUTPUT)
    print(f'OK: {OUTPUT}')
    print(f'Слайдов: {len(prs.slides)}')


if __name__ == '__main__':
    main()
