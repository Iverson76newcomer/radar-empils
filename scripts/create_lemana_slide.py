#!/usr/bin/env python3
"""
Слайд: Продвижение TITAN в сети Лемана ПРО (ЮФО).
Комплекс мер: 7 активностей + 11 торговых точек.
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

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
DARK_BLUE = RGBColor(0x1E, 0x3A, 0x5F)
LIGHT_GREEN = RGBColor(0xDC, 0xFC, 0xE7)
LIGHT_ORANGE = RGBColor(0xFF, 0xED, 0xD5)
LIGHT_RED = RGBColor(0xFE, 0xE2, 0xE2)
LIGHT_BLUE = RGBColor(0xDB, 0xEA, 0xFE)

OUTPUT = '/home/z/my-project/download/Слайд_Лемана_ПРО_TITAN.pptx'

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
                'ПРОДВИЖЕНИЕ TITAN В СЕТИ «ЛЕМАНА ПРО» — ЮФО',
                font_size=20, color=WHITE, bold=True,
                align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)
    add_textbox(slide, SLIDE_W - Inches(5.0), Inches(0.08),
                Inches(4.6), Inches(0.5),
                '11 торговых точек  •  7 активностей',
                font_size=11, color=YELLOW, bold=False,
                align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)

    # === ВЕРХНЯЯ ПОЛОСА — цели ===
    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
              Inches(0.4), Inches(0.82), SLIDE_W - Inches(0.8), Inches(0.55),
              fill_color=LIGHT_BLUE, line_color=DARK_BLUE, line_width=Pt(1))
    add_textbox(slide, Inches(0.4), Inches(0.82), SLIDE_W - Inches(0.8), Inches(0.55),
                'Цель: 1) Показать уходимость грунт-эмалей TITAN в 11 ТТ Лемана ПРО на территории ЮФО   2) Инициировать пробную покупку',
                font_size=10, color=DARK_BLUE, bold=True,
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # === ЛЕВАЯ ЧАСТЬ — 7 активностей ===
    add_textbox(slide, Inches(0.4), Inches(1.50),
                Inches(7.0), Inches(0.35),
                '7 АКТИВНОСТЕЙ ПРОДВИЖЕНИЯ',
                font_size=12, color=RED, bold=True, align=PP_ALIGN.LEFT)

    activities = [
        ('1', 'РК на экранах Ростова-на-Дону',
         '6 топовых экранов города · с 01.07 на 1 месяц',
         '✅ Запущена', GREEN),
        ('2', 'Таргетированная реклама в интернете',
         'ГЕО по 11 ТТ · ВКонтакте · с 07.07 по 31.08',
         '✅ Запущена', GREEN),
        ('3', 'Промо-акция «Подарок за покупку»',
         'Ростов: 17-18 июля · Краснодар: 18-19 июля · от 1900₽ — кисть+перчатки, от 3900₽ — сертификат 500₽',
         '🔴 Подготовка', RED),
        ('4', 'Кэшбек 5% на TITAN',
         'Во всех 11 ТТ Лемана ПРО · с 01.07 по 31.08',
         '✅ Запущен', GREEN),
        ('5', 'Мастер-классы силами Эмпилс',
         '3 ТТ в Ростове-на-Дону · 17 и 24 июля · выкрасы с покупателем + презентация',
         '🔴 Подготовка', RED),
        ('6', 'Обеспечение выкрасами 11 ТТ',
         'Выкрасы EPL и TITAN · до 16.07',
         '🟡 В работе', ORANGE),
        ('7', 'Веерные выкрасы B2B Лемана ПРО',
         '13 веерных выкрасов · изготовление до 14.07, распределение до 22.07',
         '🟡 В работе', ORANGE),
    ]

    a_y = Inches(1.90)
    a_h = Inches(0.66)
    for i, (num, name, desc, status, color) in enumerate(activities):
        cy = a_y + a_h * i
        bg = LIGHT_GRAY if i % 2 == 0 else WHITE
        add_shape(slide, MSO_SHAPE.RECTANGLE,
                  Inches(0.4), cy, Inches(7.5), a_h - Inches(0.03),
                  fill_color=bg, line_color=GRAY, line_width=Pt(0.3))
        # Номер
        add_shape(slide, MSO_SHAPE.OVAL,
                  Inches(0.50), cy + Inches(0.12),
                  Inches(0.36), Inches(0.36),
                  fill_color=RED, line_color=None)
        add_textbox(slide, Inches(0.50), cy + Inches(0.12),
                    Inches(0.36), Inches(0.36),
                    num, font_size=12, color=WHITE, bold=True,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        # Название + описание
        add_multi_text(slide, Inches(0.95), cy + Inches(0.02),
                       Inches(5.3), a_h - Inches(0.06),
                       [
                           (name, 10, DARK, True),
                           (desc, 8, GRAY, False),
                       ], anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.LEFT)
        # Статус
        status_bg = LIGHT_GREEN if '✅' in status else (LIGHT_RED if '🔴' in status else LIGHT_ORANGE)
        add_shape(slide, MSO_SHAPE.RECTANGLE,
                  Inches(6.35), cy + Inches(0.10),
                  Inches(1.45), Inches(0.42),
                  fill_color=status_bg, line_color=color, line_width=Pt(1))
        add_textbox(slide, Inches(6.35), cy + Inches(0.10),
                    Inches(1.45), Inches(0.42),
                    status, font_size=8, color=color, bold=True,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # === ПРАВАЯ ЧАСТЬ — 11 торговых точек ===
    add_textbox(slide, Inches(8.2), Inches(1.50),
                Inches(5.0), Inches(0.35),
                '11 ТОРГОВЫХ ТОЧЕК — ЮФО',
                font_size=12, color=DARK_BLUE, bold=True, align=PP_ALIGN.LEFT)

    # Заголовок таблицы ТТ
    tt_headers = ['№', 'Город', 'Рег. №']
    tt_col_w = [Inches(0.5), Inches(3.5), Inches(1.0)]
    tt_y = Inches(1.90)
    tt_header_h = Inches(0.32)

    cx = Inches(8.2)
    for i, h in enumerate(tt_headers):
        add_shape(slide, MSO_SHAPE.RECTANGLE,
                  cx, tt_y, tt_col_w[i], tt_header_h,
                  fill_color=DARK_BLUE, line_color=WHITE, line_width=Pt(1))
        add_textbox(slide, cx, tt_y, tt_col_w[i], tt_header_h,
                    h, font_size=9, color=WHITE, bold=True,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        cx += tt_col_w[i]

    # Данные ТТ
    tt_data = [
        ('1', 'Ростов-на-Дону, ул. Доватора, 158', '94'),
        ('2', 'Аксай, Аксайский пр-т, 23', '7'),
        ('3', 'Верхнетемерницкий, ул. Обсерваторная, 15', '14'),
        ('4', 'Волгоград, Университетский пр-т, 107', '24'),
        ('5', 'Краснодар, ул. Покрышкина, 5', '92'),
        ('6', 'Краснодар, ул. Новороссийская, 234/1', '8'),
        ('7', 'Новороссийск, ул. Ленина, 7', '165'),
        ('8', 'Ставропольский край, с. Татарка', '69'),
        ('9', 'Владикавказ, Архонское шоссе, 1А', '150'),
        ('10', 'Адыгея, Тургеневское шоссе, 27/1', '147'),
        ('11', 'Волжский, пр-т Ленина, 2Ж', '70'),
    ]

    tt_row_h = Inches(0.33)
    for ri, (num, city, reg) in enumerate(tt_data):
        cy = tt_y + tt_header_h + tt_row_h * ri
        bg = LIGHT_GRAY if ri % 2 == 0 else WHITE
        cx = Inches(8.2)
        for ci, val in enumerate([num, city, reg]):
            add_shape(slide, MSO_SHAPE.RECTANGLE,
                      cx, cy, tt_col_w[ci], tt_row_h,
                      fill_color=bg, line_color=GRAY, line_width=Pt(0.3))
            add_textbox(slide, cx, cy, tt_col_w[ci], tt_row_h,
                        val, font_size=8, color=DARK, bold=(ci == 0),
                        align=PP_ALIGN.CENTER if ci != 1 else PP_ALIGN.LEFT,
                        anchor=MSO_ANCHOR.MIDDLE)
            cx += tt_col_w[ci]

    # === ПРАВАЯ ЧАСТЬ — доп. блок: география + механика ===
    geo_y = tt_y + tt_header_h + tt_row_h * len(tt_data) + Inches(0.20)

    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
              Inches(8.2), geo_y, Inches(5.0), Inches(1.45),
              fill_color=LIGHT_ORANGE, line_color=ORANGE, line_width=Pt(1.5))

    add_textbox(slide, Inches(8.3), geo_y + Inches(0.05),
                Inches(4.8), Inches(0.30),
                'ГЕОГРАФИЯ И МЕХАНИКА', font_size=10, color=ORANGE, bold=True,
                align=PP_ALIGN.LEFT)

    add_multi_text(slide, Inches(8.3), geo_y + Inches(0.35),
                   Inches(4.8), Inches(1.05),
                   [
                       ('6 регионов ЮФО: Ростовская обл., Краснодарский край, Волгоградская обл., Ставропольский край, Сев. Осетия, Адыгея', 8, DARK, False),
                       (' ', 4, WHITE, False),
                       ('Механика: от 1 900 ₽ — кисть + перчатки в подарок', 8, DARK, True),
                       ('            от 3 900 ₽ — сертификат Лемана ПРО 500 ₽', 8, DARK, True),
                       ('            + 5% кэшбек на TITAN', 8, GREEN, True),
                   ], anchor=MSO_ANCHOR.TOP, align=PP_ALIGN.LEFT)

    # === НИЖНЯЯ ПОЛОСА — слоган ===
    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
              Inches(0.4), Inches(6.95), SLIDE_W - Inches(0.8), Inches(0.45),
              fill_color=RED, line_color=None)
    add_textbox(slide, Inches(0.4), Inches(6.95), SLIDE_W - Inches(0.8), Inches(0.45),
                'Красишь 1 раз — служит 10 лет  •  Период: июль — август 2026  •  Empils  •  TITAN',
                font_size=12, color=YELLOW, bold=True,
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    prs.save(OUTPUT)
    print(f'OK: {OUTPUT}')
    print(f'Слайдов: {len(prs.slides)}')


if __name__ == '__main__':
    main()
