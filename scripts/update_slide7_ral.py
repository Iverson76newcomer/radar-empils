#!/usr/bin/env python3
"""
Обновление слайда 7 (сводная таблица) — добавить RAL каждого цвета.
Берём исходный файл пользователя, заменяем только слайд 7.
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

INPUT = '/home/z/my-project/upload/Планограммы_TITAN_презентация.pptx'
OUTPUT = '/home/z/my-project/download/Планограммы_TITAN_презентация.pptx'

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


def main():
    prs = Presentation(INPUT)
    print(f'Загружено: {INPUT}')
    print(f'Слайдов: {len(prs.slides)}')

    # Удаляем старый слайд 7 (индекс 6)
    # В python-pptx удаляем через XML
    old_slide = prs.slides[6]
    xml_slides = prs.slides._sldIdLst
    slides_list = list(xml_slides)
    xml_slides.remove(slides_list[6])
    print('Старый слайд 7 удалён')

    # Создаём новый слайд 7 на его месте
    # Добавляем в конец, потом переместим
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    add_shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H,
              fill_color=WHITE, line_color=None)

    # Шапка
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, Inches(0.65),
              fill_color=DARK_BLUE, line_color=None)
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0, Inches(0.65), SLIDE_W, Inches(0.04),
              fill_color=YELLOW, line_color=None)
    add_textbox(slide, Inches(0.4), Inches(0.08), SLIDE_W - Inches(0.8), Inches(0.5),
                'СВОДНАЯ ТАБЛИЦА — MML TITAN ПО КЛАССАМ ТТ',
                font_size=22, color=WHITE, bold=True,
                align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)
    add_textbox(slide, SLIDE_W - Inches(5.0), Inches(0.08),
                Inches(4.6), Inches(0.5),
                'Сравнение по категориям · RAL указан',
                font_size=11, color=YELLOW, bold=False,
                align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)

    # Таблица — с RAL каждого цвета
    headers = ['Категория / Объём', 'A — Федер. DIY', 'B — Локал. сеть', 'C — Супермаркет', 'D — Хозмаг']
    col_widths = [Inches(3.5), Inches(2.4), Inches(2.4), Inches(2.2), Inches(2.0)]
    table_y = Inches(0.95)
    header_h = Inches(0.45)

    cx = Inches(0.3)
    for i, h in enumerate(headers):
        add_shape(slide, MSO_SHAPE.RECTANGLE,
                  cx, table_y, col_widths[i], header_h,
                  fill_color=DARK_BLUE, line_color=WHITE, line_width=Pt(1))
        add_textbox(slide, cx, table_y, col_widths[i], header_h,
                    h, font_size=10, color=WHITE, bold=True,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        cx += col_widths[i]

    # Данные с RAL для каждого цвета
    data = [
        ('3в1 по ржавчине, 0.75 л',
         '5 цветов:\nRAL 9005 (чёрный)\nRAL 8017 (шоколад)\nRAL 7004 (серый)\nRAL 7024 (графит)\nRAL 9010 (белый)',
         '5 цветов:\nRAL 9005 (чёрный)\nRAL 8017 (шоколад)\nRAL 7004 (серый)\nRAL 7024 (графит)\nRAL 9010 (белый)',
         '—', '—'),
        
        ('3в1 по ржавчине, 2 л',
         '5 цветов:\nRAL 9005 (чёрный)\nRAL 8017 (шоколад)\nRAL 7004 (серый)\nRAL 7024 (графит)\nRAL 9010 (белый)',
         '4 цвета:\nRAL 9005 (чёрный)\nRAL 8017 (шоколад)\nRAL 7004 (серый)\nRAL 7024 (графит)',
         '—', '—'),
        
        ('Быстросохнущая, 0.75 л',
         '1 цвет:\nRAL 9005 (чёрный)',
         '1 цвет:\nRAL 9005 (чёрный)',
         '—', '—'),
        
        ('Быстросохнущая, 2.2 л',
         '1 цвет:\nRAL 9005 (чёрный)',
         '—', '—', '—'),
        
        ('Молотковая, 0.8 кг',
         '3 цвета:\nЧёрный\nКоричневый\nСеребристый',
         '2 цвета:\nЧёрный\nКоричневый',
         '2 цвета:\nЧёрный\nКоричневый',
         '—'),
        
        ('Молотковая, 2.5 кг',
         '1 цвет:\nЧёрный',
         '—', '—', '—'),
    ]

    row_h = Inches(0.75)
    for ri, row in enumerate(data):
        ry = table_y + header_h + row_h * ri
        bg = LIGHT_GRAY if ri % 2 == 0 else WHITE
        cx = Inches(0.3)
        for ci, val in enumerate(row):
            add_shape(slide, MSO_SHAPE.RECTANGLE,
                      cx, ry, col_widths[ci], row_h,
                      fill_color=bg, line_color=GRAY, line_width=Pt(0.3))
            val_color = RED if val == '—' else DARK
            val_bold = (ci == 0)
            add_textbox(slide, cx + Inches(0.05), ry,
                        col_widths[ci] - Inches(0.10), row_h,
                        val, font_size=8, color=val_color, bold=val_bold,
                        align=PP_ALIGN.LEFT if ci == 0 else PP_ALIGN.CENTER,
                        anchor=MSO_ANCHOR.MIDDLE)
            cx += col_widths[ci]

    # ИТОГО
    ry = table_y + header_h + row_h * len(data)
    cx = Inches(0.3)
    totals = ['ИТОГО SKU', '16', '12', '2', '0']
    for ci, val in enumerate(totals):
        bg_color = DARK_BLUE if ci == 0 else [RED, ORANGE, GREEN, DARK_BLUE][ci-1]
        add_shape(slide, MSO_SHAPE.RECTANGLE,
                  cx, ry, col_widths[ci], Inches(0.45),
                  fill_color=bg_color, line_color=WHITE, line_width=Pt(1))
        add_textbox(slide, cx, ry, col_widths[ci], Inches(0.45),
                    val, font_size=14, color=WHITE, bold=True,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        cx += col_widths[ci]

    # === ТОП-5 цветов ===
    top_y = ry + Inches(0.55) + Inches(0.20)
    add_textbox(slide, Inches(0.3), top_y,
                Inches(12.5), Inches(0.35),
                'ТОП-5 ЦВЕТОВ ПО ВТОРИЧНЫМ ПРОДАЖАМ:',
                font_size=12, color=DARK_BLUE, bold=True, align=PP_ALIGN.LEFT)

    top_colors = [
        ('1', 'Чёрный RAL 9005', '7 738 шт', '393 АКБ', RGBColor(0x1A, 0x1A, 0x1A)),
        ('2', 'Графитовый RAL 7024', '5 172 шт', '240 АКБ', RGBColor(0x40, 0x40, 0x40)),
        ('3', 'Серый RAL 7004', '5 071 шт', '247 АКБ', RGBColor(0xB0, 0xB0, 0xB0)),
        ('4', 'Шоколад RAL 8017', '4 524 шт', '287 АКБ', RGBColor(0x3D, 0x28, 0x17)),
        ('5', 'Белый RAL 9010', '3 273 шт', '225 АКБ', RGBColor(0xF5, 0xF5, 0xDC)),
    ]
    card_w = Inches(2.45)
    card_h = Inches(1.55)
    card_y = top_y + Inches(0.40)

    for i, (num, name, qty, akb, color) in enumerate(top_colors):
        cx = Inches(0.3) + (card_w + Inches(0.10)) * i
        add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
                  cx, card_y, card_w, card_h,
                  fill_color=WHITE, line_color=color, line_width=Pt(2))
        add_shape(slide, MSO_SHAPE.OVAL,
                  cx + Inches(0.10), card_y + Inches(0.10),
                  Inches(0.45), Inches(0.45),
                  fill_color=color, line_color=DARK, line_width=Pt(1))
        brightness = 0.299 * color[0] + 0.587 * color[1] + 0.114 * color[2]
        tc = DARK if brightness > 140 else WHITE
        add_textbox(slide, cx + Inches(0.10), card_y + Inches(0.10),
                    Inches(0.45), Inches(0.45),
                    num, font_size=14, color=tc, bold=True,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_textbox(slide, cx + Inches(0.65), card_y + Inches(0.10),
                    card_w - Inches(0.75), Inches(0.45),
                    name, font_size=9, color=DARK, bold=True,
                    align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)
        add_textbox(slide, cx + Inches(0.10), card_y + Inches(0.65),
                    card_w - Inches(0.20), Inches(0.35),
                    qty, font_size=14, color=RED, bold=True,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_textbox(slide, cx + Inches(0.10), card_y + Inches(1.05),
                    card_w - Inches(0.20), Inches(0.35),
                    akb, font_size=10, color=GRAY,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # Футер
    add_textbox(slide, Inches(0.3), Inches(7.15),
                SLIDE_W - Inches(0.6), Inches(0.25),
                'Empils  •  TITAN  •  Планограммы выкладки  •  Красишь 1 раз — служит 10 лет',
                font_size=8, color=GRAY, bold=False,
                align=PP_ALIGN.CENTER, italic=True)

    # Перемещаем новый слайд на позицию 7 (индекс 6)
    # python-pptx: перемещаем через XML
    xml_slides = prs.slides._sldIdLst
    slides_list = list(xml_slides)
    # Новый слайд добавлен в конец — перемещаем его на позицию 6
    new_slide_elem = slides_list[-1]
    xml_slides.remove(new_slide_elem)
    # Вставляем перед слайдом 8 (индекс 7 в текущем списке после удаления)
    xml_slides.insert(6, new_slide_elem)

    prs.save(OUTPUT)
    print(f'Сохранено: {OUTPUT}')
    print(f'Слайдов: {len(prs.slides)}')


if __name__ == '__main__':
    main()
