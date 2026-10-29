#!/usr/bin/env python3
"""
Слайд: Обучение торговых представителей по программе TITAN TCO.
Масштаб: 12 регионов, 6 директоров, 19 РМ, 109 дистрибьюторов.
Цель: дистанционно обучить каждого ТП общего прайса.
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
BLUE = RGBColor(0x25, 0x63, 0xEB)
LIGHT_GREEN = RGBColor(0xDC, 0xFC, 0xE7)
LIGHT_ORANGE = RGBColor(0xFF, 0xED, 0xD5)
LIGHT_RED = RGBColor(0xFE, 0xE2, 0xE2)
LIGHT_BLUE = RGBColor(0xDB, 0xEA, 0xFE)

OUTPUT = '/home/z/my-project/download/Слайд_обучение_ТП_TCO.pptx'

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
                '🎓  ОБУЧЕНИЕ ТОРГОВЫХ ПРЕДСТАВИТЕЛЕЙ ПО ПРОГРАММЕ TITAN TCO',
                font_size=20, color=WHITE, bold=True,
                align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)
    add_textbox(slide, SLIDE_W - Inches(5.0), Inches(0.08),
                Inches(4.6), Inches(0.5),
                'Сбор данных → Обучение → Тест',
                font_size=12, color=YELLOW, bold=False,
                align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)

    # === ВЕРХНЯЯ ПОЛОСА — масштаб ===
    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
              Inches(0.4), Inches(0.85), SLIDE_W - Inches(0.8), Inches(0.65),
              fill_color=ORANGE, line_color=None)
    add_textbox(slide, Inches(0.4), Inches(0.85), SLIDE_W - Inches(0.8), Inches(0.65),
                'Масштаб программы: 12 регионов  •  6 региональных директоров  •  19 региональных менеджеров  •  109 дистрибьюторов  •  ~500 ТП общего прайса',
                font_size=13, color=WHITE, bold=True,
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # === ЛЕВАЯ ЧАСТЬ — таблица по регионам ===

    add_textbox(slide, Inches(0.4), Inches(1.70),
                Inches(7.0), Inches(0.35),
                'СТРУКТУРА ДИСТРИБУЦИИ ПО РЕГИОНАМ',
                font_size=12, color=DARK_BLUE, bold=True, align=PP_ALIGN.LEFT)

    # Таблица регионов
    headers = ['Регион (SBU)', 'Директор', 'РМ', 'Дистр.']
    col_widths = [Inches(2.5), Inches(2.8), Inches(0.8), Inches(0.9)]
    table_y = Inches(2.10)
    header_h = Inches(0.35)

    cx = Inches(0.4)
    for i, h in enumerate(headers):
        add_shape(slide, MSO_SHAPE.RECTANGLE,
                  cx, table_y, col_widths[i], header_h,
                  fill_color=DARK_BLUE, line_color=WHITE, line_width=Pt(1))
        add_textbox(slide, cx, table_y, col_widths[i], header_h,
                    h, font_size=9, color=WHITE, bold=True,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        cx += col_widths[i]

    # Данные по регионам (из файла)
    regions = [
        ('Дальний Восток', 'Белокопытов А.Н.', '4', '21', LIGHT_GRAY),
        ('Экспорт', 'Копай-Гора Д.А.', '1', '21', WHITE),
        ('Юг', 'Меликсетьян С.З.', '2', '19', LIGHT_GRAY),
        ('Поволжье', 'Меликсетьян С.З.', '4', '9', WHITE),
        ('Урал', 'Захарова О.Г.', '4', '9', LIGHT_GRAY),
        ('Западная Сибирь', 'Захарова О.Г.', '2', '6', WHITE),
        ('Центр', 'Еремеев В.Н.', '2', '6', LIGHT_GRAY),
        ('Северо-Запад', 'Еремеев В.Н.', '1', '5', WHITE),
        ('Федеральные сети', 'Пустовалов С.И.', '1', '5', LIGHT_GRAY),
        ('Москва', 'Еремеев В.Н.', '1', '3', WHITE),
        ('Новые территории', 'Меликсетьян С.З.', '1', '3', LIGHT_GRAY),
        ('Восточная Сибирь', 'Белокопытов А.Н.', '1', '2', WHITE),
    ]

    row_h = Inches(0.30)
    for ri, (region, director, n_rm, n_dist, bg) in enumerate(regions):
        cy = table_y + header_h + row_h * ri
        cx = Inches(0.4)
        # Регион
        add_shape(slide, MSO_SHAPE.RECTANGLE,
                  cx, cy, col_widths[0], row_h,
                  fill_color=bg, line_color=GRAY, line_width=Pt(0.3))
        add_textbox(slide, cx + Inches(0.05), cy, col_widths[0] - Inches(0.10), row_h,
                    region, font_size=8, color=DARK, bold=True,
                    align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)
        cx += col_widths[0]
        # Директор
        add_shape(slide, MSO_SHAPE.RECTANGLE,
                  cx, cy, col_widths[1], row_h,
                  fill_color=bg, line_color=GRAY, line_width=Pt(0.3))
        add_textbox(slide, cx + Inches(0.05), cy, col_widths[1] - Inches(0.10), row_h,
                    director, font_size=8, color=DARK,
                    align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)
        cx += col_widths[1]
        # РМ
        add_shape(slide, MSO_SHAPE.RECTANGLE,
                  cx, cy, col_widths[2], row_h,
                  fill_color=bg, line_color=GRAY, line_width=Pt(0.3))
        add_textbox(slide, cx, cy, col_widths[2], row_h,
                    n_rm, font_size=9, color=DARK, bold=True,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        cx += col_widths[2]
        # Дистрибьюторы
        add_shape(slide, MSO_SHAPE.RECTANGLE,
                  cx, cy, col_widths[3], row_h,
                  fill_color=bg, line_color=GRAY, line_width=Pt(0.3))
        add_textbox(slide, cx, cy, col_widths[3], row_h,
                    n_dist, font_size=9, color=RED, bold=True,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # ИТОГО
    total_y = table_y + header_h + row_h * len(regions)
    cx = Inches(0.4)
    add_shape(slide, MSO_SHAPE.RECTANGLE,
              cx, total_y, col_widths[0] + col_widths[1], row_h,
              fill_color=DARK, line_color=WHITE, line_width=Pt(1))
    add_textbox(slide, cx + Inches(0.05), total_y, col_widths[0] + col_widths[1] - Inches(0.10), row_h,
                'ИТОГО',
                font_size=10, color=WHITE, bold=True,
                align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)
    cx += col_widths[0] + col_widths[1]
    # РМ
    add_shape(slide, MSO_SHAPE.RECTANGLE,
              cx, total_y, col_widths[2], row_h,
              fill_color=DARK, line_color=WHITE, line_width=Pt(1))
    add_textbox(slide, cx, total_y, col_widths[2], row_h,
                '19', font_size=10, color=YELLOW, bold=True,
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    cx += col_widths[2]
    # Дистрибьюторы
    add_shape(slide, MSO_SHAPE.RECTANGLE,
              cx, total_y, col_widths[3], row_h,
              fill_color=DARK, line_color=WHITE, line_width=Pt(1))
    add_textbox(slide, cx, total_y, col_widths[3], row_h,
                '109', font_size=10, color=YELLOW, bold=True,
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # === ПРАВАЯ ЧАСТЬ — процесс и цель ===

    # (1) Что собираем по каждому ТП
    add_textbox(slide, Inches(7.8), Inches(1.70),
                Inches(5.5), Inches(0.35),
                'ЧТО СОБИРАЕМ ПО КАЖДОМУ ТП',
                font_size=12, color=DARK_BLUE, bold=True, align=PP_ALIGN.LEFT)

    fields = [
        ('ФИО торгового представителя', 'полные данные для учёта'),
        ('Телефон', 'для связи и рассылки материалов'),
        ('Email', 'для отправки обучения и теста'),
        ('Тип прайса: общий / эксклюзивный', 'обучаем общий прайс'),
        ('Дистрибьютор и регион', 'привязка к РМ и директору'),
        ('Дата последнего обучения', 'если обучался ранее'),
        ('Готовность пройти тест', 'да / нет / запланировано'),
    ]
    f_y = Inches(2.10)
    f_h = Inches(0.38)
    for i, (field, desc) in enumerate(fields):
        cy = f_y + f_h * i
        bg = LIGHT_GRAY if i % 2 == 0 else WHITE
        add_shape(slide, MSO_SHAPE.RECTANGLE,
                  Inches(7.8), cy, Inches(5.4), f_h - Inches(0.03),
                  fill_color=bg, line_color=GRAY, line_width=Pt(0.3))
        add_shape(slide, MSO_SHAPE.OVAL,
                  Inches(7.90), cy + Inches(0.08),
                  Inches(0.22), Inches(0.22),
                  fill_color=ORANGE, line_color=None)
        add_textbox(slide, Inches(7.90), cy + Inches(0.08),
                    Inches(0.22), Inches(0.22),
                    str(i + 1), font_size=8, color=WHITE, bold=True,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_multi_text(slide, Inches(8.20), cy,
                       Inches(4.9), f_h - Inches(0.03),
                       [
                           (field + ' — ', 9, DARK, True),
                           (desc, 8, GRAY, False),
                       ], anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.LEFT)

    # (2) Процесс обучения
    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
              Inches(7.8), Inches(4.85), Inches(5.4), Inches(2.30),
              fill_color=LIGHT_GREEN, line_color=GREEN, line_width=Pt(2))
    add_textbox(slide, Inches(7.8), Inches(4.90),
                Inches(5.4), Inches(0.35),
                '🎯  ПРОЦЕСС ОБУЧЕНИЯ TCO',
                font_size=12, color=GREEN, bold=True, align=PP_ALIGN.LEFT)

    steps = [
        ('1', 'Сбор данных по ТП', 'Региональные менеджеры заполняют таблицу по своим дистрибьюторам', RED),
        ('2', 'Рассылка материалов', 'Онлайн-материалы по TITAN TCO + видеоинструкция', ORANGE),
        ('3', 'Прохождение теста', 'Каждый ТП проходит тест на знание программы TCO', GREEN),
        ('4', 'Сертификация', 'Сертификат «ТП TITAN TCO» + включение в программу мотивации', DARK_BLUE),
    ]
    st_y = Inches(5.30)
    st_h = Inches(0.42)
    for i, (num, title, desc, color) in enumerate(steps):
        cy = st_y + st_h * i
        add_shape(slide, MSO_SHAPE.OVAL,
                  Inches(7.95), cy + Inches(0.05),
                  Inches(0.28), Inches(0.28),
                  fill_color=color, line_color=None)
        add_textbox(slide, Inches(7.95), cy + Inches(0.05),
                    Inches(0.28), Inches(0.28),
                    num, font_size=9, color=WHITE, bold=True,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_multi_text(slide, Inches(8.35), cy,
                       Inches(4.7), st_h,
                       [
                           (title + ' — ', 9, color, True),
                           (desc, 8, DARK, False),
                       ], anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.LEFT)

    # === ФУТЕР ===
    add_textbox(slide, Inches(0.4), Inches(7.15),
                SLIDE_W - Inches(0.8), Inches(0.30),
                'Empils  •  TITAN  •  Программа TCO для торговых представителей общего прайса  •  Красишь 1 раз — служит 10 лет',
                font_size=8, color=GRAY, bold=False,
                align=PP_ALIGN.CENTER, italic=True)

    prs.save(OUTPUT)
    print(f'OK: {OUTPUT}')
    print(f'Слайдов: {len(prs.slides)}')


if __name__ == '__main__':
    main()
