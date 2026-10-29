#!/usr/bin/env python3
"""
TITAN — Дорожная карта развития продаж (v2).
Иструменты взяты СТРОГО из файла "Матрицы_инструментов_ТITAN_два_сценария.docx".
Добавлен слайд "Поиск нового партнёра" с табличкой из 6 инструментов.
В третичных продажах — визуализации POSm.
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
import os

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
BLUE = RGBColor(0x25, 0x63, 0xEB)
DARK_BLUE = RGBColor(0x1E, 0x3A, 0x5F)
LIGHT_BLUE = RGBColor(0xDB, 0xEA, 0xFE)
LIGHT_GREEN = RGBColor(0xDC, 0xFC, 0xE7)
LIGHT_ORANGE = RGBColor(0xFF, 0xED, 0xD5)
LIGHT_RED = RGBColor(0xFE, 0xE2, 0xE2)

OUTPUT = '/home/z/my-project/download/Дорожная_карта_TITAN.pptx'
POSM_DIR = '/home/z/my-project/download/titan_posm_visuals'

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


def add_bullet_list(slide, left, top, width, height, items, font_size=12,
                    color=DARK, bullet_color=ORANGE, line_spacing=1.3):
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = Pt(4); tf.margin_right = Pt(4)
    tf.margin_top = Pt(4); tf.margin_bottom = Pt(4)
    for i, item in enumerate(items):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        p.line_spacing = line_spacing
        r1 = p.add_run()
        r1.text = '•  '
        r1.font.size = Pt(font_size)
        r1.font.bold = True
        r1.font.color.rgb = bullet_color
        r1.font.name = 'Arial'
        r2 = p.add_run()
        r2.text = item
        r2.font.size = Pt(font_size)
        r2.font.color.rgb = color
        r2.font.name = 'Arial'
    return tb


def header_band(slide, title, subtitle=None, color=RED):
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, Inches(0.65),
              fill_color=color, line_color=None)
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0, Inches(0.65), SLIDE_W, Inches(0.04),
              fill_color=YELLOW, line_color=None)
    add_textbox(slide, Inches(0.4), Inches(0.08), SLIDE_W - Inches(0.8), Inches(0.5),
                title, font_size=22, color=WHITE, bold=True,
                align=PP_ALIGN.LEFT, font_name='Arial',
                anchor=MSO_ANCHOR.MIDDLE)
    if subtitle:
        add_textbox(slide, SLIDE_W - Inches(5.0), Inches(0.08),
                    Inches(4.6), Inches(0.5),
                    subtitle, font_size=12, color=WHITE, bold=False,
                    align=PP_ALIGN.RIGHT, font_name='Arial',
                    anchor=MSO_ANCHOR.MIDDLE)


def footer(slide):
    add_textbox(slide, Inches(0.4), Inches(7.05),
                SLIDE_W - Inches(0.8), Inches(0.4),
                'Empils  •  TITAN  •  Дорожная карта развития продаж  •  10 лет без перекраски — 10 лет экономии!',
                font_size=9, color=GRAY, bold=False,
                align=PP_ALIGN.CENTER, font_name='Arial', italic=True)


# ==============================================================
# СЛАЙД 1: ТИТУЛ
# ==============================================================

def slide_title(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H,
              fill_color=WHITE, line_color=None)

    add_shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, Inches(5.5), SLIDE_H,
              fill_color=RED, line_color=None)
    add_shape(slide, MSO_SHAPE.RECTANGLE, Inches(5.5), 0, Inches(0.10), SLIDE_H,
              fill_color=YELLOW, line_color=None)

    add_textbox(slide, Inches(0.4), Inches(1.0), Inches(5.0), Inches(1.2),
                'TITAN', font_size=80, color=WHITE, bold=True,
                align=PP_ALIGN.LEFT, font_name='Arial')
    add_textbox(slide, Inches(0.4), Inches(2.3), Inches(5.0), Inches(0.7),
                'ДОРОЖНАЯ КАРТА', font_size=36, color=WHITE, bold=True,
                align=PP_ALIGN.LEFT, font_name='Arial')
    add_textbox(slide, Inches(0.4), Inches(3.1), Inches(5.0), Inches(0.7),
                'РАЗВИТИЯ ПРОДАЖ', font_size=36, color=WHITE, bold=True,
                align=PP_ALIGN.LEFT, font_name='Arial')

    add_shape(slide, MSO_SHAPE.RECTANGLE, Inches(0.4), Inches(4.0),
              Inches(2.5), Inches(0.05), fill_color=YELLOW, line_color=None)

    add_textbox(slide, Inches(0.4), Inches(4.2), Inches(5.0), Inches(0.5),
                'Программа для дистрибьюторов',
                font_size=18, color=WHITE, bold=False, align=PP_ALIGN.LEFT)
    add_textbox(slide, Inches(0.4), Inches(4.75), Inches(5.0), Inches(0.5),
                '2 сценария · поиск · 3 этапа продаж',
                font_size=16, color=WHITE, bold=False, align=PP_ALIGN.LEFT, italic=True)
    add_textbox(slide, Inches(0.4), Inches(5.4), Inches(5.0), Inches(1.0),
                'Эмпилс гарантирует: при выполнении всех\nдействий продажи TITAN будут расти',
                font_size=14, color=YELLOW, bold=True, align=PP_ALIGN.LEFT)

    # Правая часть
    add_textbox(slide, Inches(6.0), Inches(1.0), Inches(7.0), Inches(1.0),
                '10 лет без перекраски',
                font_size=36, color=GREEN, bold=True,
                align=PP_ALIGN.LEFT, font_name='Arial')
    add_textbox(slide, Inches(6.0), Inches(1.9), Inches(7.0), Inches(1.0),
                '10 лет экономии!',
                font_size=36, color=GREEN, bold=True,
                align=PP_ALIGN.LEFT, font_name='Arial')

    add_shape(slide, MSO_SHAPE.RECTANGLE, Inches(6.0), Inches(2.95),
              Inches(2.0), Inches(0.05), fill_color=YELLOW, line_color=None)

    add_textbox(slide, Inches(6.0), Inches(3.15), Inches(7.0), Inches(0.5),
                'Структура программы',
                font_size=18, color=DARK, bold=False, align=PP_ALIGN.LEFT, italic=True)

    # 4 блока структуры
    blocks = [
        ('1', 'ПОИСК ПАРТНЁРА', '6 инструментов', RED),
        ('2', 'СЦЕНАРИЙ А', 'Новый партнёр', ORANGE),
        ('3', 'СЦЕНАРИЙ Б', 'Действующий партнёр', GREEN),
        ('4', '3 ЭТАПА', 'Первичные · Вторичные · Третичные', DARK_BLUE),
    ]
    block_y = Inches(3.9)
    block_w = Inches(1.7)
    block_h = Inches(2.4)
    for i, (num, name, sub, color) in enumerate(blocks):
        cx = Inches(6.0) + (block_w + Inches(0.15)) * i
        add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
                  cx, block_y, block_w, block_h,
                  fill_color=WHITE, line_color=color, line_width=Pt(2.5))
        add_shape(slide, MSO_SHAPE.RECTANGLE,
                  cx, block_y, block_w, Inches(0.5),
                  fill_color=color, line_color=None)
        add_textbox(slide, cx, block_y + Inches(0.05),
                    block_w, Inches(0.4),
                    f'БЛОК {num}', font_size=12, color=WHITE, bold=True,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_textbox(slide, cx + Inches(0.05), block_y + Inches(0.65),
                    block_w - Inches(0.10), Inches(0.6),
                    name, font_size=12, color=DARK, bold=True,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_textbox(slide, cx + Inches(0.05), block_y + Inches(1.3),
                    block_w - Inches(0.10), Inches(1.0),
                    sub, font_size=10, color=GRAY, bold=False,
                    align=PP_ALIGN.CENTER)

    add_textbox(slide, Inches(6.0), Inches(6.6), Inches(7.0), Inches(0.5),
                'Empils  •  100+ лет на рынке ЛКМ  •  ТОП-3 производителей России',
                font_size=11, color=GRAY, bold=False,
                align=PP_ALIGN.LEFT, italic=True)


# ==============================================================
# СЛАЙД 2: О КОМПАНИИ ЭМПИЛС
# ==============================================================

def slide_about_company(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H,
              fill_color=WHITE, line_color=None)

    header_band(slide, 'О КОМПАНИИ «ЭМПИЛС»',
                '100+ лет на рынке ЛКМ  •  ТОП-3 производителей России', color=DARK_BLUE)

    metrics = [
        ('100+', 'лет непрерывного\nразвития в химической\nотрасле', DARK_BLUE),
        ('ТОП-3', 'крупнейших\nпроизводителей ЛКМ\nв РФ', RED),
        ('8%', 'доля рынка в денежном\nвыражении', GREEN),
        ('39,3', 'тыс. тонн продукции\nза 2025 год', ORANGE),
    ]
    metric_y = Inches(1.0)
    metric_w = Inches(3.05)
    metric_h = Inches(2.2)
    for i, (num, label, color) in enumerate(metrics):
        cx = Inches(0.4) + (metric_w + Inches(0.15)) * i
        add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
                  cx, metric_y, metric_w, metric_h,
                  fill_color=WHITE, line_color=color, line_width=Pt(2))
        add_textbox(slide, cx, metric_y + Inches(0.15),
                    metric_w, Inches(1.0),
                    num, font_size=48, color=color, bold=True,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_textbox(slide, cx + Inches(0.1), metric_y + Inches(1.25),
                    metric_w - Inches(0.2), Inches(0.9),
                    label, font_size=11, color=DARK, bold=False,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.TOP)

    add_textbox(slide, Inches(0.4), Inches(3.5), Inches(12.5), Inches(0.5),
                'ГЕОГРАФИЯ БИЗНЕСА',
                font_size=14, color=DARK_BLUE, bold=True, align=PP_ALIGN.LEFT)

    geo_items = [
        ('100+', 'дистрибьюторов\nв России и за рубежом', RED),
        ('150', 'городов России\nс представленностью', ORANGE),
        ('9', 'стран экспорта\n(СНГ + Монголия)', GREEN),
        ('36 000 м²', 'площадь складов\nкомпании', DARK_BLUE),
    ]
    geo_y = Inches(4.1)
    for i, (num, label, color) in enumerate(geo_items):
        cx = Inches(0.4) + (metric_w + Inches(0.15)) * i
        add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
                  cx, geo_y, metric_w, Inches(1.5),
                  fill_color=LIGHT_GRAY, line_color=color, line_width=Pt(1.5))
        add_textbox(slide, cx, geo_y + Inches(0.1),
                    metric_w, Inches(0.6),
                    num, font_size=24, color=color, bold=True,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_textbox(slide, cx + Inches(0.1), geo_y + Inches(0.75),
                    metric_w - Inches(0.2), Inches(0.7),
                    label, font_size=10, color=DARK, bold=False,
                    align=PP_ALIGN.CENTER)

    add_textbox(slide, Inches(0.4), Inches(5.85), Inches(12.5), Inches(0.5),
                'ПРОИЗВОДСТВО — 4 ЦЕХА, 168 000 ТОНН В ГОД',
                font_size=14, color=DARK_BLUE, bold=True, align=PP_ALIGN.LEFT)

    workshops = [
        ('Цех лаков', '30 000 т/год', 'полуфабрикатные лаки'),
        ('Цех эмалей (малотоннажные)', '18 000 т/год', 'партии от 700 кг'),
        ('Цех водных покрытий', '30 000 т/год', 'экологичная продукция'),
        ('Цех эмалей (крупнотоннажные)', '90 000 т/год', 'до 280 т/сутки'),
    ]
    ws_y = Inches(6.35)
    ws_w = (Inches(12.5) - Inches(0.45)) / 4
    for i, (name, vol, desc) in enumerate(workshops):
        cx = Inches(0.4) + (ws_w + Inches(0.15)) * i
        add_shape(slide, MSO_SHAPE.RECTANGLE,
                  cx, ws_y, ws_w, Inches(0.65),
                  fill_color=DARK_BLUE, line_color=None)
        add_multi_text(slide, cx + Inches(0.05), ws_y + Inches(0.03),
                       ws_w - Inches(0.10), Inches(0.6),
                       [
                           (name, 9, WHITE, True),
                           (vol + ' · ' + desc, 8, RGBColor(0xCB, 0xD5, 0xE1), False),
                       ])

    footer(slide)


# ==============================================================
# СЛАЙД 3: СПЕЦИАЛИЗАЦИЯ TITAN
# ==============================================================

def slide_titan_specialization(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H,
              fill_color=WHITE, line_color=None)

    header_band(slide, 'СПЕЦИАЛИЗАЦИЯ TITAN',
                'У коррозии нет шансов  •  10 лет защиты', color=RED)

    # Левая часть — слоган
    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
              Inches(0.4), Inches(1.0), Inches(6.0), Inches(5.8),
              fill_color=RED, line_color=None)

    add_textbox(slide, Inches(0.6), Inches(1.3), Inches(5.6), Inches(0.6),
                'У КОРРОЗИИ', font_size=32, color=WHITE, bold=True,
                align=PP_ALIGN.LEFT)
    add_textbox(slide, Inches(0.6), Inches(1.95), Inches(5.6), Inches(0.6),
                'НЕТ ШАНСОВ', font_size=32, color=YELLOW, bold=True,
                align=PP_ALIGN.LEFT)

    add_shape(slide, MSO_SHAPE.RECTANGLE, Inches(0.6), Inches(2.75),
              Inches(2.0), Inches(0.05), fill_color=YELLOW, line_color=None)

    add_textbox(slide, Inches(0.6), Inches(2.95), Inches(5.6), Inches(2.5),
                'Специализированная линейка материалов по металлу TITAN сочетает в себе свойства самых последних и инновационных разработок в сфере лакокрасочных покрытий.\n\nМатериалы TITAN предотвращают появление коррозии и сохраняют целостность конструкции в течение 10 лет.\n\nПодтверждено независимым испытательным центром ООО «Рутил».',
                font_size=13, color=WHITE, bold=False, align=PP_ALIGN.LEFT)

    add_shape(slide, MSO_SHAPE.RECTANGLE,
              Inches(0.6), Inches(5.7), Inches(5.6), Inches(0.9),
              fill_color=YELLOW, line_color=None)
    add_textbox(slide, Inches(0.6), Inches(5.7), Inches(5.6), Inches(0.9),
                '10 лет без перекраски —\n10 лет экономии!',
                font_size=20, color=RED, bold=True,
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # Правая часть — 3 линейки
    add_textbox(slide, Inches(6.7), Inches(1.0), Inches(6.5), Inches(0.5),
                '3 ЛИНЕЙКИ TITAN',
                font_size=18, color=DARK, bold=True, align=PP_ALIGN.LEFT)

    lines = [
        ('3в1 ПО РЖАВЧИНЕ', 'Грунт + эмаль + преобразователь ржавчины\nв одной банке. Глянцевая. 9 цветов RAL.',
         '0.75 л · 2 л · 5 л', RED, '10 лет гарантии'),
        ('БЫСТРОСОХНУЩАЯ', 'Матовая. Сохнет 30 минут.\nДо -15 °C. Для профессионалов.',
         '0.75 л · 2.2 л', DARK, '1.5 ч на отлип'),
        ('МОЛОТКОВАЯ', 'Декоративный молотковый эффект.\nСкрывает дефекты поверхности.',
         '0.8 кг · 2.5 кг', GOLD, '10 лет защиты'),
    ]
    line_y = Inches(1.6)
    line_h = Inches(1.65)
    for i, (name, desc, vols, color, badge) in enumerate(lines):
        cy = line_y + (line_h + Inches(0.10)) * i
        add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
                  Inches(6.7), cy, Inches(6.5), line_h,
                  fill_color=WHITE, line_color=color, line_width=Pt(2))
        add_shape(slide, MSO_SHAPE.RECTANGLE,
                  Inches(6.7), cy, Inches(0.15), line_h,
                  fill_color=color, line_color=None)
        add_textbox(slide, Inches(7.0), cy + Inches(0.10),
                    Inches(4.0), Inches(0.5),
                    name, font_size=15, color=color, bold=True,
                    align=PP_ALIGN.LEFT)
        add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
                  Inches(11.0), cy + Inches(0.10),
                  Inches(2.0), Inches(0.35),
                  fill_color=color, line_color=None)
        add_textbox(slide, Inches(11.0), cy + Inches(0.10),
                    Inches(2.0), Inches(0.35),
                    badge, font_size=10, color=WHITE, bold=True,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_textbox(slide, Inches(7.0), cy + Inches(0.60),
                    Inches(6.0), Inches(0.7),
                    desc, font_size=11, color=DARK, bold=False,
                    align=PP_ALIGN.LEFT)
        add_textbox(slide, Inches(7.0), cy + Inches(1.25),
                    Inches(6.0), Inches(0.35),
                    'Фасовка: ' + vols, font_size=10, color=GRAY, bold=False,
                    align=PP_ALIGN.LEFT, italic=True)

    footer(slide)


# ==============================================================
# СЛАЙД 4: ПОИСК НОВОГО ПАРТНЁРА (НОВЫЙ СЛАЙД)
# ==============================================================

def slide_partner_search(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H,
              fill_color=WHITE, line_color=None)

    header_band(slide, 'ПОИСК НОВОГО ПАРТНЁРА',
                '6 инструментов выхода на контакт  •  Сценарий А', color=DARK_BLUE)

    # Подзаголовок
    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
              Inches(0.4), Inches(0.95), SLIDE_W - Inches(0.8), Inches(0.7),
              fill_color=LIGHT_BLUE, line_color=DARK_BLUE, line_width=Pt(1.5))
    add_textbox(slide, Inches(0.6), Inches(0.95), SLIDE_W - Inches(1.2), Inches(0.7),
                'Цель: выйти на контакт с потенциальным партнёром и презентовать программу TITAN',
                font_size=14, color=DARK_BLUE, bold=True,
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # Табличка из 6 инструментов
    tools = [
        ('1', 'Персональный визит РМ',
         'Региональный менеджер приезжает в офис потенциального партнёра',
         'Самый эффективный канал. Конверсия 40-60%'),
        ('2', 'Звонок РМ (теплый прозвон)',
         'Звонок после email или бумажного письма — договориться о встрече',
         'Конверсия в встречу 15-25%'),
        ('3', 'Отправка приглашения в МАХ',
         'Презентация программы TITAN через мессенджер МАХ',
         'Удобно для дистрибьюторов, использующих МАХ'),
        ('4', 'Email-рассылка по базе «белые пятна»',
         'Персонализированное письмо + приглашение посетить завод',
         'Масштабируемый канал. Конверсия 3-7%'),
        ('5', 'Письмо по обычной почте',
         'Напечатанное приглашение посетить завод, отправленное Почтой России',
         'Конверсия 8-15%. Бумага воспринимается как «серьёзный жест»'),
        ('6', 'Приглашение на профильной выставке',
         'Контакт в период участия партнёра в выставке (MosBuild и др.)',
         'Личный контакт на стенде Эмпилс'),
    ]

    table_y = Inches(1.85)
    row_h = Inches(0.78)

    # Заголовок таблицы
    headers = ['№', 'Инструмент', 'Описание', 'Эффективность']
    col_widths = [Inches(0.6), Inches(3.5), Inches(5.5), Inches(3.3)]
    cx = Inches(0.4)
    for i, h in enumerate(headers):
        add_shape(slide, MSO_SHAPE.RECTANGLE,
                  cx, table_y, col_widths[i], Inches(0.45),
                  fill_color=DARK_BLUE, line_color=WHITE, line_width=Pt(1))
        add_textbox(slide, cx, table_y, col_widths[i], Inches(0.45),
                    h, font_size=11, color=WHITE, bold=True,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        cx += col_widths[i]

    # Данные
    for ri, (num, name, desc, eff) in enumerate(tools):
        ry = table_y + Inches(0.45) + row_h * ri
        bg = LIGHT_GRAY if ri % 2 == 0 else WHITE
        cx = Inches(0.4)
        # №
        add_shape(slide, MSO_SHAPE.RECTANGLE,
                  cx, ry, col_widths[0], row_h,
                  fill_color=ORANGE, line_color=WHITE, line_width=Pt(1))
        add_textbox(slide, cx, ry, col_widths[0], row_h,
                    num, font_size=18, color=WHITE, bold=True,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        cx += col_widths[0]
        # Инструмент
        add_shape(slide, MSO_SHAPE.RECTANGLE,
                  cx, ry, col_widths[1], row_h,
                  fill_color=bg, line_color=GRAY, line_width=Pt(0.5))
        add_textbox(slide, cx + Inches(0.1), ry, col_widths[1] - Inches(0.2), row_h,
                    name, font_size=11, color=DARK, bold=True,
                    align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)
        cx += col_widths[1]
        # Описание
        add_shape(slide, MSO_SHAPE.RECTANGLE,
                  cx, ry, col_widths[2], row_h,
                  fill_color=bg, line_color=GRAY, line_width=Pt(0.5))
        add_textbox(slide, cx + Inches(0.1), ry, col_widths[2] - Inches(0.2), row_h,
                    desc, font_size=10, color=DARK, bold=False,
                    align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)
        cx += col_widths[2]
        # Эффективность
        add_shape(slide, MSO_SHAPE.RECTANGLE,
                  cx, ry, col_widths[3], row_h,
                  fill_color=bg, line_color=GRAY, line_width=Pt(0.5))
        add_textbox(slide, cx + Inches(0.1), ry, col_widths[3] - Inches(0.2), row_h,
                    eff, font_size=10, color=GREEN, bold=True,
                    align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)

    footer(slide)


# ==============================================================
# СЛАЙД 5: ДВА СЦЕНАРИЯ РАБОТЫ
# ==============================================================

def slide_two_scenarios(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H,
              fill_color=WHITE, line_color=None)

    header_band(slide, 'ДВА СЦЕНАРИЯ РАБОТЫ С ДИСТРИБЬЮТОРОМ',
                'От «убедить попробовать» до «увеличить продажи»', color=DARK)

    # Текст из Word файла (строго)
    scenarios = [
        ('А', 'НОВЫЙ ПАРТНЁР', '«Убедить попробовать»',
         'Дистрибьютор не хочет заводить TITAN, потому что у него уже есть аналог. Главный барьер — страх: «Закуплю — не продам, заморожу деньги, потеряю время». Инструменты этого сценария снимают страх и делают первую закупку максимально безболезненной.',
         ['Бонус за первичную закупку 5-10%',
          'Возврат неликвида в 90 дней',
          'Отсрочка 70-90 дней',
          'Выездное обучение команды ТП (1 день)',
          'Бесплатные POS-материалы + стартовый набор',
          'Стартовый набор постов для соцсетей'],
         RED, LIGHT_RED),
        ('Б', 'ДЕЙСТВУЮЩИЙ ПАРТНЁР', '«Увеличить продажи»',
         'Дистрибьютор уже продаёт TITAN 3в1, но продажи не растут или нестабильны. Он знает продукт, но не вкладывается в его продвижение. Причины: нет мотивации ТП, нет инструментов для розницы, нет обучения продавцов, нет системного подхода.',
         ['Купи Титана для 100% планограммы → бонус',
          'Мин. заказ от 1 банки. Глубина: B-5, C-3, D-2',
          'Отсрочка 14 дней для розницы',
          'Выездное обучение продавцов РТТ (45 мин)',
          'Консигнация для топ-5 точек',
          '5 баночек 0.4 для трудных РТТ'],
         GREEN, LIGHT_GREEN),
    ]

    block_y = Inches(1.0)
    block_w = Inches(6.2)
    block_h = Inches(5.7)

    for i, (letter, name, sub, problem, tools, color, bg_color) in enumerate(scenarios):
        cx = Inches(0.4) + (block_w + Inches(0.30)) * i

        add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
                  cx, block_y, block_w, block_h,
                  fill_color=bg_color, line_color=color, line_width=Pt(2.5))

        # Буква
        add_shape(slide, MSO_SHAPE.OVAL,
                  cx + Inches(0.3), block_y + Inches(0.3),
                  Inches(0.8), Inches(0.8),
                  fill_color=color, line_color=None)
        add_textbox(slide, cx + Inches(0.3), block_y + Inches(0.3),
                    Inches(0.8), Inches(0.8),
                    letter, font_size=36, color=WHITE, bold=True,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

        add_textbox(slide, cx + Inches(1.3), block_y + Inches(0.30),
                    block_w - Inches(1.5), Inches(0.5),
                    name, font_size=20, color=color, bold=True,
                    align=PP_ALIGN.LEFT)
        add_textbox(slide, cx + Inches(1.3), block_y + Inches(0.75),
                    block_w - Inches(1.5), Inches(0.4),
                    sub, font_size=13, color=DARK, bold=False,
                    align=PP_ALIGN.LEFT, italic=True)

        # Проблема (из Word)
        add_textbox(slide, cx + Inches(0.3), block_y + Inches(1.5),
                    block_w - Inches(0.6), Inches(1.5),
                    problem, font_size=10, color=DARK, bold=False,
                    align=PP_ALIGN.LEFT)

        # Инструменты (из Word)
        add_textbox(slide, cx + Inches(0.3), block_y + Inches(3.1),
                    block_w - Inches(0.6), Inches(0.4),
                    'КЛЮЧЕВЫЕ ИНСТРУМЕНТЫ:',
                    font_size=11, color=color, bold=True, align=PP_ALIGN.LEFT)

        add_bullet_list(slide, cx + Inches(0.3), block_y + Inches(3.5),
                        block_w - Inches(0.6), Inches(2.0),
                        tools, font_size=11, color=DARK, bullet_color=color)

    footer(slide)


# ==============================================================
# СЛАЙД 6: ЭТАП 1 — ПЕРВИЧНЫЕ ПРОДАЖИ (из Word файла)
# ==============================================================

def slide_stage_1(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H,
              fill_color=WHITE, line_color=None)

    header_band(slide, 'ЭТАП 1 — ПЕРВИЧНЫЕ ПРОДАЖИ',
                'Завод → Дистрибьютор', color=RED)

    add_textbox(slide, Inches(0.4), Inches(0.95),
                Inches(8.0), Inches(0.5),
                'ИНСТРУМЕНТЫ (из матрицы Сценарий А)',
                font_size=14, color=RED, bold=True, align=PP_ALIGN.LEFT)

    # Строго из Word файла
    tools = [
        ('Экскурсия на завод',
         'Показать производство → лояльность руководства дистрибьютора'),
        ('Бонус за первичную закупку 5-10% товаром',
         'За 80% покрытия АКБ точек А, Б. Мотивация первички на покрытие'),
        ('Возврат неликвида в 90 дней',
         'Снять страх «куплю и не продам» на всех уровнях'),
        ('Отсрочка 70-90 дней',
         'Снизить финансовую нагрузку на старте'),
        ('Выездное обучение команды ТП общего прайса (1 день)',
         'Команда умеет продавать с первого дня'),
        ('Бесплатные POS-материалы + стартовый набор',
         'Выкрасы, шелфтокеры, воблеры, планограммы, мини-буклет портфолио с ССВ'),
        ('Стартовый набор постов для соцсетей',
         'Картинка + текст для соцсетей дистрибьютора + фото/видео для интернет-магазина'),
    ]

    tool_y = Inches(1.45)
    tool_h = Inches(0.72)
    for i, (name, desc) in enumerate(tools):
        cy = tool_y + tool_h * i
        bg = LIGHT_GRAY if i % 2 == 0 else WHITE
        add_shape(slide, MSO_SHAPE.RECTANGLE,
                  Inches(0.4), cy, Inches(8.0), tool_h,
                  fill_color=bg, line_color=GRAY, line_width=Pt(0.3))
        add_shape(slide, MSO_SHAPE.OVAL,
                  Inches(0.5), cy + Inches(0.15),
                  Inches(0.42), Inches(0.42),
                  fill_color=RED, line_color=None)
        add_textbox(slide, Inches(0.5), cy + Inches(0.15),
                    Inches(0.42), Inches(0.42),
                    str(i + 1), font_size=14, color=WHITE, bold=True,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_multi_text(slide, Inches(1.05), cy + Inches(0.05),
                       Inches(7.2), Inches(0.65),
                       [
                           (name, 11, RED, True),
                           (desc, 9, DARK, False),
                       ], anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.LEFT)

    # ЧТО ДЕЛАЕТ ДИСТРИБЬЮТОР (справа)
    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
              Inches(8.6), Inches(0.95), Inches(4.5), Inches(5.9),
              fill_color=LIGHT_RED, line_color=RED, line_width=Pt(2))

    add_textbox(slide, Inches(8.6), Inches(1.05),
                Inches(4.5), Inches(0.5),
                'ЧТО ДЕЛАЕТ ДИСТРИБЬЮТОР',
                font_size=14, color=RED, bold=True,
                align=PP_ALIGN.CENTER)

    # Таблица «Обязательство | Срок»
    obligations = [
        ('Оформить первую закупку (мин. объём)', '0–1 месяц'),
        ('Обучить торговых представителей по TITAN', '1–2 месяца'),
        ('Обеспечить выкладку в точках А и Б согласно планограмме', '1–3 месяца'),
        ('Предоставлять отчётность по выполнению планограммы', 'Ежемесячно'),
    ]
    ob_y = Inches(1.65)
    ob_h = Inches(1.15)
    for i, (obligation, term) in enumerate(obligations):
        cy = ob_y + ob_h * i
        add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
                  Inches(8.8), cy, Inches(4.1), ob_h - Inches(0.10),
                  fill_color=WHITE, line_color=RED, line_width=Pt(1))
        # Номер
        add_shape(slide, MSO_SHAPE.OVAL,
                  Inches(8.95), cy + Inches(0.15),
                  Inches(0.35), Inches(0.35),
                  fill_color=RED, line_color=None)
        add_textbox(slide, Inches(8.95), cy + Inches(0.15),
                    Inches(0.35), Inches(0.35),
                    str(i + 1), font_size=11, color=WHITE, bold=True,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        # Обязательство
        add_textbox(slide, Inches(9.4), cy + Inches(0.10),
                    Inches(3.4), Inches(0.55),
                    obligation, font_size=9, color=DARK, bold=True,
                    align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP)
        # Срок (внизу карточки, на цветной плашке)
        add_shape(slide, MSO_SHAPE.RECTANGLE,
                  Inches(9.4), cy + Inches(0.70),
                  Inches(3.35), Inches(0.28),
                  fill_color=RED, line_color=None)
        add_textbox(slide, Inches(9.4), cy + Inches(0.70),
                    Inches(3.35), Inches(0.28),
                    '⏱ ' + term, font_size=9, color=WHITE, bold=True,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    footer(slide)


# ==============================================================
# СЛАЙД 7: ЭТАП 2 — ВТОРИЧНЫЕ ПРОДАЖИ (из Word файла)
# ==============================================================

def slide_stage_2(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H,
              fill_color=WHITE, line_color=None)

    header_band(slide, 'ЭТАП 2 — ВТОРИЧНЫЕ ПРОДАЖИ',
                'Дистрибьютор → Розница', color=ORANGE)

    add_textbox(slide, Inches(0.4), Inches(0.95),
                Inches(8.0), Inches(0.5),
                'ИНСТРУМЕНТЫ (из матрицы Сценарий А)',
                font_size=14, color=ORANGE, bold=True, align=PP_ALIGN.LEFT)

    # Строго из Word файла
    tools = [
        ('Купи Титана для 100% заполнения планограммы → бонус',
         'Мотивация вторички на необходимый объём по выполнению планограммы'),
        ('Минимальный заказ от 1 банки для розницы',
         'Глубина остатка на полке: B — от 5, C — 3, D — 2. Ширина — штуки'),
        ('Отсрочка 14 дней для розницы',
         'Снизить финансовую нагрузку на розницу'),
        ('Выездное обучение продавцов ключевых РТТ (45 мин)',
         'РМ / ЭТП — команда умеет продавать с первого дня'),
        ('Бесплатные POS-материалы + стартовый набор',
         'Дополнительно 5 баночек 0.4 каждому РМ для работы с трудными РТТ'),
        ('Стартовый набор постов для соцсетей',
         'Картинка + текст + фото/видео для интернет-магазина'),
        ('Консигнация для топ-5 точек дистрибьютора',
         'На усмотрение РМ. Товар стоит на полке без оплаты — риск = 0'),
    ]

    tool_y = Inches(1.45)
    tool_h = Inches(0.72)
    for i, (name, desc) in enumerate(tools):
        cy = tool_y + tool_h * i
        bg = LIGHT_GRAY if i % 2 == 0 else WHITE
        add_shape(slide, MSO_SHAPE.RECTANGLE,
                  Inches(0.4), cy, Inches(8.0), tool_h,
                  fill_color=bg, line_color=GRAY, line_width=Pt(0.3))
        add_shape(slide, MSO_SHAPE.OVAL,
                  Inches(0.5), cy + Inches(0.15),
                  Inches(0.42), Inches(0.42),
                  fill_color=ORANGE, line_color=None)
        add_textbox(slide, Inches(0.5), cy + Inches(0.15),
                    Inches(0.42), Inches(0.42),
                    str(i + 1), font_size=14, color=WHITE, bold=True,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_multi_text(slide, Inches(1.05), cy + Inches(0.05),
                       Inches(7.2), Inches(0.65),
                       [
                           (name, 11, ORANGE, True),
                           (desc, 9, DARK, False),
                       ], anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.LEFT)

    # ЧТО ДЕЛАЕТ ДИСТРИБЬЮТОР (справа)
    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
              Inches(8.6), Inches(0.95), Inches(4.5), Inches(5.9),
              fill_color=LIGHT_ORANGE, line_color=ORANGE, line_width=Pt(2))

    add_textbox(slide, Inches(8.6), Inches(1.05),
                Inches(4.5), Inches(0.5),
                'ЧТО ДЕЛАЕТ ДИСТРИБЬЮТОР',
                font_size=14, color=ORANGE, bold=True,
                align=PP_ALIGN.CENTER)

    obligations = [
        ('Обучить 100% ТП по TITAN (методика ТСО)', 'К 3-му месяцу'),
        ('Открыть новые точки (+25% к базе)', 'Ежемесячно'),
        ('Обеспечить выкладку в точках В и С', '3–6 месяцев'),
        ('Предоставлять фотоотчёты по выкладке', 'Ежемесячно'),
    ]
    ob_y = Inches(1.65)
    ob_h = Inches(1.15)
    for i, (obligation, term) in enumerate(obligations):
        cy = ob_y + ob_h * i
        add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
                  Inches(8.8), cy, Inches(4.1), ob_h - Inches(0.10),
                  fill_color=WHITE, line_color=ORANGE, line_width=Pt(1))
        add_shape(slide, MSO_SHAPE.OVAL,
                  Inches(8.95), cy + Inches(0.15),
                  Inches(0.35), Inches(0.35),
                  fill_color=ORANGE, line_color=None)
        add_textbox(slide, Inches(8.95), cy + Inches(0.15),
                    Inches(0.35), Inches(0.35),
                    str(i + 1), font_size=11, color=WHITE, bold=True,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_textbox(slide, Inches(9.4), cy + Inches(0.10),
                    Inches(3.4), Inches(0.55),
                    obligation, font_size=9, color=DARK, bold=True,
                    align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP)
        add_shape(slide, MSO_SHAPE.RECTANGLE,
                  Inches(9.4), cy + Inches(0.70),
                  Inches(3.35), Inches(0.28),
                  fill_color=ORANGE, line_color=None)
        add_textbox(slide, Inches(9.4), cy + Inches(0.70),
                    Inches(3.35), Inches(0.28),
                    '⏱ ' + term, font_size=9, color=WHITE, bold=True,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    footer(slide)


# ==============================================================
# СЛАЙД 8: ЭТАП 3 — ТРЕТИЧНЫЕ ПРОДАЖИ (в стиле слайдов 6/7)
# ==============================================================

def slide_stage_3_main(prs):
    """Слайд: Этап 3 — Третичные продажи.
    Слева — инструменты (из Word), справа — что делает дистрибьютор.
    Визуально аналогично слайдам 6 (Этап 1) и 7 (Этап 2)."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H,
              fill_color=WHITE, line_color=None)

    header_band(slide, 'ЭТАП 3 — ТРЕТИЧНЫЕ ПРОДАЖИ',
                'Розница → Покупатель', color=GREEN)

    add_textbox(slide, Inches(0.4), Inches(0.95),
                Inches(8.0), Inches(0.5),
                'ИНСТРУМЕНТЫ МАРКЕТИНГА (из матрицы)',
                font_size=14, color=GREEN, bold=True, align=PP_ALIGN.LEFT)

    # Инструменты строго из Word файла
    tools = [
        ('Шелфтокер на полку / со стопером',
         'Слоган «10 лет без перекраски — 10 лет экономии!» + QR на titanpaints.ru'),
        ('Воблер на полку',
         'Слоган «Красишь 1 раз — служит 10 лет». Розница видит выгоду'),
        ('Брендированный ценник',
         'Переводит цену банки в «цену за год» (760 руб. → 76 руб./год)'),
        ('Портфолио объектов с ССВ',
         'Экономические параметры совокупной стоимости владения'),
        ('MML + Планограмма выкладки Титан',
         'С шелфтокерами и воблерами. 3 полки · 22 SKU · 3 линейки'),
        ('Инструкция по нанесению (мини + видео)',
         'Открытие по QR-коду. 5 шагов + видео инструкция'),
        ('Фото/видео POSm на облаке',
         'Материалы для розничных точек на облаке'),
        ('Контент для соцсетей дистрибьюторов',
         'Библиотека готовых постов (картинка + текст)'),
        ('Видеоролики для ТП и продаж',
         'УТП по Титан + имитация идеальной продажи клиенту'),
        ('QR на банке + иконография ССВ',
         'Видеоинструкция + калькулятор ССВ + «Дороже банка — дешевле покраска!»'),
    ]

    tool_y = Inches(1.45)
    tool_h = Inches(0.52)
    for i, (name, desc) in enumerate(tools):
        cy = tool_y + tool_h * i
        bg = LIGHT_GRAY if i % 2 == 0 else WHITE
        add_shape(slide, MSO_SHAPE.RECTANGLE,
                  Inches(0.4), cy, Inches(8.0), tool_h,
                  fill_color=bg, line_color=GRAY, line_width=Pt(0.3))
        add_shape(slide, MSO_SHAPE.OVAL,
                  Inches(0.5), cy + Inches(0.10),
                  Inches(0.32), Inches(0.32),
                  fill_color=GREEN, line_color=None)
        add_textbox(slide, Inches(0.5), cy + Inches(0.10),
                    Inches(0.32), Inches(0.32),
                    str(i + 1), font_size=11, color=WHITE, bold=True,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_multi_text(slide, Inches(0.90), cy + Inches(0.02),
                       Inches(7.4), Inches(0.48),
                       [
                           (name + ' — ', 10, GREEN, True),
                           (desc, 9, DARK, False),
                       ], anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.LEFT)

    # ЧТО ДЕЛАЕТ ДИСТРИБЬЮТОР (справа) — аналогично слайдам 6/7
    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
              Inches(8.6), Inches(0.95), Inches(4.5), Inches(5.9),
              fill_color=LIGHT_GREEN, line_color=GREEN, line_width=Pt(2))

    add_textbox(slide, Inches(8.6), Inches(1.05),
                Inches(4.5), Inches(0.5),
                'ЧТО ДЕЛАЕТ ДИСТРИБЬЮТОР',
                font_size=14, color=GREEN, bold=True,
                align=PP_ALIGN.CENTER)

    obligations = [
        ('Обеспечить наличие POS-материалов в 80% точек с TITAN', 'К 6-му месяцу'),
        ('Обучить продавцов розницы (сертификация ТСО)', '3–6 месяцев'),
        ('Предоставлять фото выкладок', 'Ежемесячно'),
        ('Собирать обратную связь от покупателей', 'Постоянно'),
    ]
    ob_y = Inches(1.65)
    ob_h = Inches(1.15)
    for i, (obligation, term) in enumerate(obligations):
        cy = ob_y + ob_h * i
        add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
                  Inches(8.8), cy, Inches(4.1), ob_h - Inches(0.10),
                  fill_color=WHITE, line_color=GREEN, line_width=Pt(1))
        add_shape(slide, MSO_SHAPE.OVAL,
                  Inches(8.95), cy + Inches(0.15),
                  Inches(0.35), Inches(0.35),
                  fill_color=GREEN, line_color=None)
        add_textbox(slide, Inches(8.95), cy + Inches(0.15),
                    Inches(0.35), Inches(0.35),
                    str(i + 1), font_size=11, color=WHITE, bold=True,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_textbox(slide, Inches(9.4), cy + Inches(0.10),
                    Inches(3.4), Inches(0.55),
                    obligation, font_size=9, color=DARK, bold=True,
                    align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP)
        add_shape(slide, MSO_SHAPE.RECTANGLE,
                  Inches(9.4), cy + Inches(0.70),
                  Inches(3.35), Inches(0.28),
                  fill_color=GREEN, line_color=None)
        add_textbox(slide, Inches(9.4), cy + Inches(0.70),
                    Inches(3.35), Inches(0.28),
                    '⏱ ' + term, font_size=9, color=WHITE, bold=True,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    footer(slide)


# ==============================================================
# СЛАЙДЫ 9-11: ЭТАП 3 — ТРЕТИЧНЫЕ ПРОДАЖИ (МАРКЕТИНГ) с визуализациями
# ==============================================================

def slide_stage_3_intro(prs):
    """Слайд 9: Введение в третичные продажи — маркетинг"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H,
              fill_color=WHITE, line_color=None)

    header_band(slide, 'ЭТАП 3 — ТРЕТИЧНЫЕ ПРОДАЖИ (МАРКЕТИНГ)',
                'Розница → Покупатель  •  9 инструментов с визуализацией', color=GREEN)

    # 9 инструментов в сетке 3×3 с мини-визуализациями
    add_textbox(slide, Inches(0.4), Inches(0.95),
                Inches(12.5), Inches(0.5),
                '9 ИНСТРУМЕНТОВ МАРКЕТИНГА — С ВИЗУАЛИЗАЦИЕЙ POSM',
                font_size=14, color=GREEN, bold=True, align=PP_ALIGN.LEFT)

    # 9 карточек с визуализациями
    tools = [
        ('1', 'Шелфтокер на полку', '01_shelftalker.png',
         'Слоган «10 лет без перекраски — 10 лет экономии!» + QR на titanpaints.ru'),
        ('2', 'Воблер', '02_wobbler.png',
         'Слоган «Красишь 1 раз — служит 10 лет»'),
        ('3', 'Брендированный ценник', '03_price_tag.png',
         'Переводит цену банки в «цену за год»'),
        ('4', 'Портфолио объектов с ССВ', '04_portfolio.png',
         'Экономические параметры совокупной стоимости владения'),
        ('5', 'MML + Планограмма', '05_planogram.png',
         'Выкладка Титан с шелфтокерами и воблерами'),
        ('6', 'Инструкция по нанесению', '06_instruction_qr.png',
         'Мини-инструкция + видео по QR-коду'),
        ('7', 'Фото/видео POSm на облаке', '07_cloud_posm.png',
         'Материалы для розничных точек на облаке'),
        ('8', 'Контент для соцсетей', '08_social_content.png',
         'Библиотека готовых постов (картинка + текст)'),
        ('9', 'Видеоролики', '09_videos.png',
         'Для ТП общего прайса + имитация идеальной продажи'),
    ]

    card_y = Inches(1.5)
    card_w = Inches(4.1)
    card_h = Inches(1.85)
    gap_x = Inches(0.15)
    gap_y = Inches(0.15)

    for i, (num, name, img, desc) in enumerate(tools):
        row = i // 3
        col = i % 3
        cx = Inches(0.4) + (card_w + gap_x) * col
        cy = card_y + (card_h + gap_y) * row

        # Карточка
        add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
                  cx, cy, card_w, card_h,
                  fill_color=WHITE, line_color=GREEN, line_width=Pt(1.5))

        # Номер
        add_shape(slide, MSO_SHAPE.OVAL,
                  cx + Inches(0.10), cy + Inches(0.10),
                  Inches(0.35), Inches(0.35),
                  fill_color=GREEN, line_color=None)
        add_textbox(slide, cx + Inches(0.10), cy + Inches(0.10),
                    Inches(0.35), Inches(0.35),
                    num, font_size=12, color=WHITE, bold=True,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

        # Название
        add_textbox(slide, cx + Inches(0.55), cy + Inches(0.10),
                    card_w - Inches(0.65), Inches(0.35),
                    name, font_size=11, color=DARK, bold=True,
                    align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)

        # Изображение
        img_path = os.path.join(POSM_DIR, img)
        if os.path.exists(img_path):
            slide.shapes.add_picture(img_path,
                                     cx + Inches(0.15), cy + Inches(0.55),
                                     width=card_w - Inches(0.30),
                                     height=Inches(0.85))

        # Описание (внизу карточки)
        add_textbox(slide, cx + Inches(0.10), cy + Inches(1.45),
                    card_w - Inches(0.20), Inches(0.35),
                    desc, font_size=8, color=GRAY, bold=False,
                    align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)

    footer(slide)


def slide_stage_3_detail_1(prs):
    """Слайд 9: Детализация инструментов 1-5 (с большими визуализациями)"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H,
              fill_color=WHITE, line_color=None)

    header_band(slide, 'ЭТАП 3 — POS-МАТЕРИАЛЫ НА ПОЛКЕ',
                'Инструменты 1-5  •  Шелфтокер · Воблер · Ценник · Портфолио · Планограмма',
                color=GREEN)

    # 5 карточек с большими визуализациями
    tools = [
        ('1', 'ШЕЛФТОКЕР', '01_shelftalker.png',
         'На полку с выкладкой Титан / со стопером.\nСлоган «10 лет без перекраски — 10 лет экономии!»\n+ QR-код на titanpaints.ru'),
        ('2', 'ВОБЛЕР', '02_wobbler.png',
         'На полку. Слоган «Красишь 1 раз — служит 10 лет».\nРозница видит выгоду и берёт TITAN на пробу.'),
        ('3', 'БРЕНДИРОВАННЫЙ ЦЕННИК', '03_price_tag.png',
         'Переводит цену банки в «цену за год».\n760 руб. за банку → 76 руб./год (760 ÷ 10 лет).'),
        ('4', 'ПОРТФОЛИО ОБЪЕКТОВ', '04_portfolio.png',
         'Наиболее распространённые объекты\nс экономическими параметрами ССВ.\nВорота, забор, гараж, решётки.'),
        ('5', 'MML + ПЛАНОГРАММА', '05_planogram.png',
         'Выкладка Титан с шелфтокерами и воблерами.\n3 полки · 22 SKU · 3 линейки.'),
    ]

    card_y = Inches(1.0)
    card_w = Inches(2.5)
    card_h = Inches(5.8)
    gap_x = Inches(0.10)

    for i, (num, name, img, desc) in enumerate(tools):
        cx = Inches(0.4) + (card_w + gap_x) * i

        # Карточка
        add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
                  cx, card_y, card_w, card_h,
                  fill_color=WHITE, line_color=GREEN, line_width=Pt(2))

        # Шапка с номером
        add_shape(slide, MSO_SHAPE.RECTANGLE,
                  cx, card_y, card_w, Inches(0.7),
                  fill_color=GREEN, line_color=None)
        add_textbox(slide, cx, card_y + Inches(0.05),
                    card_w, Inches(0.3),
                    f'ИНСТРУМЕНТ {num}', font_size=9, color=YELLOW, bold=True,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_textbox(slide, cx, card_y + Inches(0.30),
                    card_w, Inches(0.4),
                    name, font_size=11, color=WHITE, bold=True,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

        # Изображение (большое)
        img_path = os.path.join(POSM_DIR, img)
        if os.path.exists(img_path):
            slide.shapes.add_picture(img_path,
                                     cx + Inches(0.15), card_y + Inches(0.85),
                                     width=card_w - Inches(0.30),
                                     height=Inches(3.0))

        # Описание
        add_textbox(slide, cx + Inches(0.10), card_y + Inches(4.0),
                    card_w - Inches(0.20), Inches(1.7),
                    desc, font_size=9, color=DARK, bold=False,
                    align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP)

    footer(slide)


def slide_stage_3_detail_2(prs):
    """Слайд 10: Детализация инструментов 6-9 + доп. (QR, иконография)"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H,
              fill_color=WHITE, line_color=None)

    header_band(slide, 'ЭТАП 3 — КОНТЕНТ И ОБУЧЕНИЕ ПОКУПАТЕЛЯ',
                'Инструменты 6-9 + QR на банке + иконография ССВ',
                color=GREEN)

    # 4 карточки + 2 дополнительные
    tools = [
        ('6', 'ИНСТРУКЦИЯ ПО НАНЕСЕНИЮ', '06_instruction_qr.png',
         'Открытие по QR-коду.\nМини-инструкция (5 шагов)\n+ видео инструкция.'),
        ('7', 'ФОТО/ВИДЕО POSm НА ОБЛАКЕ', '07_cloud_posm.png',
         'Материалы для розничных точек\nна облаке. Фото выкладки,\nвидео, шаблоны.'),
        ('8', 'КОНТЕНТ ДЛЯ СОЦСЕТЕЙ', '08_social_content.png',
         'Папка на облаке с библиотекой\nготовых постов про Титан\n(картинка + текст).'),
        ('9', 'ВИДЕОРОЛИКИ', '09_videos.png',
         'Для ТП общего прайса:\nУТП, свойства, выгода по Титан.\n+ ролик идеальной продажи.'),
    ]

    card_y = Inches(1.0)
    card_w = Inches(3.05)
    card_h = Inches(3.5)
    gap_x = Inches(0.15)

    for i, (num, name, img, desc) in enumerate(tools):
        cx = Inches(0.4) + (card_w + gap_x) * i

        # Карточка
        add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
                  cx, card_y, card_w, card_h,
                  fill_color=WHITE, line_color=GREEN, line_width=Pt(2))

        # Шапка
        add_shape(slide, MSO_SHAPE.RECTANGLE,
                  cx, card_y, card_w, Inches(0.65),
                  fill_color=GREEN, line_color=None)
        add_textbox(slide, cx, card_y + Inches(0.05),
                    card_w, Inches(0.3),
                    f'ИНСТРУМЕНТ {num}', font_size=9, color=YELLOW, bold=True,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_textbox(slide, cx, card_y + Inches(0.30),
                    card_w, Inches(0.35),
                    name, font_size=10, color=WHITE, bold=True,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

        # Изображение
        img_path = os.path.join(POSM_DIR, img)
        if os.path.exists(img_path):
            slide.shapes.add_picture(img_path,
                                     cx + Inches(0.15), card_y + Inches(0.80),
                                     width=card_w - Inches(0.30),
                                     height=Inches(1.7))

        # Описание
        add_textbox(slide, cx + Inches(0.10), card_y + Inches(2.60),
                    card_w - Inches(0.20), Inches(0.85),
                    desc, font_size=9, color=DARK, bold=False,
                    align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP)

    # Дополнительные инструменты (из таблицы Word — QR на банке + иконография)
    add_y = Inches(4.7)
    add_h = Inches(2.1)

    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
              Inches(0.4), add_y, Inches(6.2), add_h,
              fill_color=LIGHT_GREEN, line_color=GREEN, line_width=Pt(1.5))
    add_textbox(slide, Inches(0.6), add_y + Inches(0.10),
                Inches(5.8), Inches(0.4),
                'ДОП: QR-КОД НА БАНКЕ', font_size=12, color=GREEN, bold=True)
    add_textbox(slide, Inches(0.6), add_y + Inches(0.50),
                Inches(5.8), Inches(1.5),
                'QR-код на банке → видеоинструкция использования + калькулятор ССВ (совокупной стоимости владения).\n\nПокупатель сам узнаёт о TITAN, без продавца.',
                font_size=11, color=DARK, bold=False, align=PP_ALIGN.LEFT)

    # Иконография ССВ
    img_path = os.path.join(POSM_DIR, '10_tco_iconography.png')
    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
              Inches(6.8), add_y, Inches(6.3), add_h,
              fill_color=LIGHT_GREEN, line_color=GREEN, line_width=Pt(1.5))
    add_textbox(slide, Inches(7.0), add_y + Inches(0.10),
                Inches(5.8), Inches(0.4),
                'ДОП: ИКОНОГРАФИКА ССВ НА БАНКЕ', font_size=12, color=GREEN, bold=True)
    if os.path.exists(img_path):
        slide.shapes.add_picture(img_path,
                                 Inches(7.0), add_y + Inches(0.50),
                                 width=Inches(1.8), height=Inches(1.5))
    add_textbox(slide, Inches(9.0), add_y + Inches(0.50),
                Inches(3.9), Inches(1.5),
                'Шелфтокер-иконографика ССВ.\nСлоган: «Дороже банка —\nдешевле покраска!\nЭкономия 60% за 10 лет!»\n\nИконография этапов нанесения\nдля клиента DIY (если нет QR).',
                font_size=9, color=DARK, bold=False, align=PP_ALIGN.LEFT)

    footer(slide)


# ==============================================================
# СЛАЙД 11: ГАРАНТИЯ РЕЗУЛЬТАТА
# ==============================================================

def slide_guarantee(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H,
              fill_color=WHITE, line_color=None)

    header_band(slide, 'ГАРАНТИЯ РЕЗУЛЬТАТА',
                'Что получает дистрибьютор при выполнении программы', color=DARK_BLUE)

    # Главный слоган
    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
              Inches(0.4), Inches(1.0), SLIDE_W - Inches(0.8), Inches(1.3),
              fill_color=DARK_BLUE, line_color=None)
    add_textbox(slide, Inches(0.4), Inches(1.05), SLIDE_W - Inches(0.8), Inches(0.6),
                'Эмпилс гарантирует: при выполнении всех действий продажи TITAN будут расти',
                font_size=20, color=WHITE, bold=True,
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_textbox(slide, Inches(0.4), Inches(1.65), SLIDE_W - Inches(0.8), Inches(0.5),
                'Высокомаржинальный продукт TITAN — двигатель роста прибыли дистрибьютора',
                font_size=14, color=YELLOW, bold=True,
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # 4 карточки — что получает дистрибьютор
    add_textbox(slide, Inches(0.4), Inches(2.6), Inches(12.5), Inches(0.5),
                'ЧТО ПОЛУЧАЕТ ДИСТРИБЬЮТОР',
                font_size=14, color=DARK_BLUE, bold=True, align=PP_ALIGN.LEFT)

    benefits = [
        ('💰', 'Финансовая выгода',
         'Бонус 5-10% товаром, отсрочка 70-90 дней, возврат неликвида в 90 дней. Снятие страха «куплю — не продам».',
         GREEN),
        ('🎓', 'Обученная команда',
         'Выездное обучение ТП (1 день) + обучение продавцов РТТ (45 мин). Команда умеет продавать с первого дня.',
         BLUE),
        ('📊', 'POS-материалы бесплатно',
         'Стартовый набор: выкрасы, шелфтокеры, воблеры, планограммы, мини-буклет портфолио с ССВ. Полка готова без затрат.',
         ORANGE),
        ('📱', 'Контент и аналитика',
         'Посты для соцсетей, фото/видео для интернет-магазина, материалы на облаке. Готовый контент для продвижения.',
         RED),
    ]
    ben_y = Inches(3.2)
    ben_w = Inches(3.05)
    ben_h = Inches(3.2)
    for i, (emoji, title, desc, color) in enumerate(benefits):
        cx = Inches(0.4) + (ben_w + Inches(0.15)) * i
        add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
                  cx, ben_y, ben_w, ben_h,
                  fill_color=WHITE, line_color=color, line_width=Pt(2))
        add_shape(slide, MSO_SHAPE.RECTANGLE,
                  cx, ben_y, ben_w, Inches(0.6),
                  fill_color=color, line_color=None)
        # Эмодзи
        tb = slide.shapes.add_textbox(cx + Inches(0.15), ben_y + Inches(0.05),
                                      Inches(0.6), Inches(0.5))
        tf = tb.text_frame
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = emoji
        r.font.size = Pt(24)
        r.font.name = 'Segoe UI Emoji'

        add_textbox(slide, cx + Inches(0.75), ben_y + Inches(0.05),
                    ben_w - Inches(0.85), Inches(0.5),
                    title, font_size=12, color=WHITE, bold=True,
                    anchor=MSO_ANCHOR.MIDDLE)
        add_textbox(slide, cx + Inches(0.15), ben_y + Inches(0.75),
                    ben_w - Inches(0.30), ben_h - Inches(0.85),
                    desc, font_size=11, color=DARK, bold=False,
                    align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP)

    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
              Inches(0.4), Inches(6.55), SLIDE_W - Inches(0.8), Inches(0.5),
              fill_color=YELLOW, line_color=None)
    add_textbox(slide, Inches(0.4), Inches(6.55), SLIDE_W - Inches(0.8), Inches(0.5),
                'Эмпилс берёт на себя обучение, POS-материалы и контент. Дистрибьютор — закупку и выкладку.',
                font_size=13, color=DARK, bold=True,
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)


# ==============================================================
# СЛАЙД 11: ОБЯЗАТЕЛЬСТВА СТОРОН
# ==============================================================

def slide_obligations(prs):
    """Слайд: Обязательства сторон — кто за что отвечает."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H,
              fill_color=WHITE, line_color=None)

    header_band(slide, 'ОБЯЗАТЕЛЬСТВА СТОРОН',
                'Кто за что отвечает  •  Эмпилс vs Дистрибьютор', color=DARK_BLUE)

    # Левая колонка — что гарантирует Эмпилс
    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
              Inches(0.4), Inches(1.0), Inches(6.2), Inches(5.9),
              fill_color=LIGHT_BLUE, line_color=DARK_BLUE, line_width=Pt(2))

    # Шапка
    add_shape(slide, MSO_SHAPE.RECTANGLE,
              Inches(0.4), Inches(1.0), Inches(6.2), Inches(0.6),
              fill_color=DARK_BLUE, line_color=None)
    add_textbox(slide, Inches(0.4), Inches(1.0), Inches(6.2), Inches(0.6),
                '✓  ЧТО ГАРАНТИРУЕМ МЫ (ЭМПИЛС)',
                font_size=14, color=WHITE, bold=True,
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # 7 гарантий
    empils_guarantees = [
        ('1', 'Качество продукта', 'Срок службы 10 лет при правильном нанесении'),
        ('2', 'Обучение', 'Выездное обучение ТП + материалы + сертификация'),
        ('3', 'POS-поддержка', 'Шелфтокеры, воблеры, выкрасы, ценники — бесплатно'),
        ('4', 'Маркетинг', 'Готовый контент для соцсетей, фото/видео'),
        ('5', 'Возврат неликвида', 'В течение 90 дней (в пределах срока годности)'),
        ('6', 'Техподдержка', 'Прямой доступ к техническому специалисту'),
        ('7', 'Финансовые стимулы', 'Бонусы, отсрочка, субсидирование маркетинга'),
    ]
    eg_y = Inches(1.75)
    eg_h = Inches(0.72)
    for i, (num, title, desc) in enumerate(empils_guarantees):
        cy = eg_y + eg_h * i
        bg = WHITE if i % 2 == 0 else LIGHT_GRAY
        add_shape(slide, MSO_SHAPE.RECTANGLE,
                  Inches(0.55), cy, Inches(5.9), eg_h - Inches(0.05),
                  fill_color=bg, line_color=GRAY, line_width=Pt(0.3))
        # Номер
        add_shape(slide, MSO_SHAPE.OVAL,
                  Inches(0.65), cy + Inches(0.13),
                  Inches(0.32), Inches(0.32),
                  fill_color=DARK_BLUE, line_color=None)
        add_textbox(slide, Inches(0.65), cy + Inches(0.13),
                    Inches(0.32), Inches(0.32),
                    num, font_size=10, color=WHITE, bold=True,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        # Заголовок + описание
        add_multi_text(slide, Inches(1.10), cy + Inches(0.03),
                       Inches(5.3), eg_h - Inches(0.10),
                       [
                           (title, 10, DARK_BLUE, True),
                           (desc, 8, DARK, False),
                       ], anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.LEFT)

    # Правая колонка — что делает дистрибьютор
    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
              Inches(6.8), Inches(1.0), Inches(6.2), Inches(5.9),
              fill_color=LIGHT_ORANGE, line_color=ORANGE, line_width=Pt(2))

    # Шапка
    add_shape(slide, MSO_SHAPE.RECTANGLE,
              Inches(6.8), Inches(1.0), Inches(6.2), Inches(0.6),
              fill_color=ORANGE, line_color=None)
    add_textbox(slide, Inches(6.8), Inches(1.0), Inches(6.2), Inches(0.6),
                '📋  ЧТО ДЕЛАЕТ ДИСТРИБЬЮТОР',
                font_size=14, color=WHITE, bold=True,
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # 6 обязательств дистрибьютора
    distributor_obligations = [
        ('1', 'Минимальный заказ', 'Согласно согласованному объёму'),
        ('2', 'Выкладка', 'Согласно планограмме (точки А, Б, В)'),
        ('3', 'Обучение', 'Обучение ТП и продавцов розницы'),
        ('4', 'Отчётность', 'Фото выкладок, отчёты по продажам'),
        ('5', 'Продвижение', 'Использование POS-материалов и контента'),
        ('6', 'Обратная связь', 'Ежемесячные отчёты о продажах и проблемах'),
    ]
    do_y = Inches(1.75)
    do_h = Inches(0.84)
    for i, (num, title, desc) in enumerate(distributor_obligations):
        cy = do_y + do_h * i
        bg = WHITE if i % 2 == 0 else LIGHT_GRAY
        add_shape(slide, MSO_SHAPE.RECTANGLE,
                  Inches(6.95), cy, Inches(5.9), do_h - Inches(0.05),
                  fill_color=bg, line_color=GRAY, line_width=Pt(0.3))
        add_shape(slide, MSO_SHAPE.OVAL,
                  Inches(7.05), cy + Inches(0.15),
                  Inches(0.32), Inches(0.32),
                  fill_color=ORANGE, line_color=None)
        add_textbox(slide, Inches(7.05), cy + Inches(0.15),
                    Inches(0.32), Inches(0.32),
                    num, font_size=10, color=WHITE, bold=True,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_multi_text(slide, Inches(7.50), cy + Inches(0.03),
                       Inches(5.3), do_h - Inches(0.10),
                       [
                           (title, 10, ORANGE, True),
                           (desc, 8, DARK, False),
                       ], anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.LEFT)

    footer(slide)


# ==============================================================
# СЛАЙД 12: KPI И КОНТРОЛЬ
# ==============================================================

def slide_kpi(prs):
    """Слайд: KPI и контроль — как измеряем успех."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H,
              fill_color=WHITE, line_color=None)

    header_band(slide, 'KPI И КОНТРОЛЬ',
                'Как измеряем успех  •  10 метрик по 3 этапам', color=DARK_BLUE)

    # Подзаголовок
    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
              Inches(0.4), Inches(0.95), SLIDE_W - Inches(0.8), Inches(0.55),
              fill_color=LIGHT_BLUE, line_color=DARK_BLUE, line_width=Pt(1))
    add_textbox(slide, Inches(0.4), Inches(0.95), SLIDE_W - Inches(0.8), Inches(0.55),
                'Без цифр программа — просто список пожеланий. Вот наши измеримые цели по этапам.',
                font_size=12, color=DARK_BLUE, bold=True,
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # Таблица KPI
    table_y = Inches(1.7)
    # Заголовки
    headers = ['Этап', 'Метрика', 'Целевое значение', 'Периодичность']
    col_widths = [Inches(0.8), Inches(5.0), Inches(4.0), Inches(2.7)]
    cx = Inches(0.4)
    for i, h in enumerate(headers):
        add_shape(slide, MSO_SHAPE.RECTANGLE,
                  cx, table_y, col_widths[i], Inches(0.45),
                  fill_color=DARK_BLUE, line_color=WHITE, line_width=Pt(1))
        add_textbox(slide, cx, table_y, col_widths[i], Inches(0.45),
                    h, font_size=11, color=WHITE, bold=True,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        cx += col_widths[i]

    # Данные
    kpis = [
        ('1', 'Объём закупок TITAN', '+30% к базовому периоду', 'Ежемесячно', RED),
        ('1', 'Доля TITAN в ассортименте', 'Не менее 5 цветов RAL', 'Ежеквартально', RED),
        ('2', 'Количество активных РТТ с TITAN', '+25% к базовому периоду', 'Ежемесячно', ORANGE),
        ('2', 'Охват ТП обучением ТСО', '100% ТП пилотного дистрибьютора', 'К 3-му месяцу', ORANGE),
        ('2', 'Доля новых точек от TITAN', 'Не менее 20%', 'Ежемесячно', ORANGE),
        ('3', 'Точки с POS-материалами', '80% точек с TITAN на полке', 'К 6-му месяцу', GREEN),
        ('3', 'Сертифицированные продавцы ТСО', 'Не менее 20 на пилота', 'К 6-му месяцу', GREEN),
        ('3', 'Конверсия в покупку TITAN', 'Не менее 30%', 'Ежеквартально', GREEN),
        ('3', 'Сканирования QR-кода', '500+ сканов/мес', 'Ежемесячно', GREEN),
        ('3', 'Снижение возвратов', '−30% к базовому периоду', 'Ежеквартально', GREEN),
    ]

    row_h = Inches(0.50)
    for ri, (stage, metric, target, period, color) in enumerate(kpis):
        ry = table_y + Inches(0.45) + row_h * ri
        bg = LIGHT_GRAY if ri % 2 == 0 else WHITE
        cx = Inches(0.4)
        # Этап (цветной кружок)
        add_shape(slide, MSO_SHAPE.RECTANGLE,
                  cx, ry, col_widths[0], row_h,
                  fill_color=color, line_color=WHITE, line_width=Pt(1))
        add_textbox(slide, cx, ry, col_widths[0], row_h,
                    stage, font_size=14, color=WHITE, bold=True,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        cx += col_widths[0]
        # Метрика
        add_shape(slide, MSO_SHAPE.RECTANGLE,
                  cx, ry, col_widths[1], row_h,
                  fill_color=bg, line_color=GRAY, line_width=Pt(0.3))
        add_textbox(slide, cx + Inches(0.1), ry, col_widths[1] - Inches(0.2), row_h,
                    metric, font_size=10, color=DARK, bold=True,
                    align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)
        cx += col_widths[1]
        # Целевое значение
        add_shape(slide, MSO_SHAPE.RECTANGLE,
                  cx, ry, col_widths[2], row_h,
                  fill_color=bg, line_color=GRAY, line_width=Pt(0.3))
        add_textbox(slide, cx + Inches(0.1), ry, col_widths[2] - Inches(0.2), row_h,
                    target, font_size=10, color=color, bold=True,
                    align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)
        cx += col_widths[2]
        # Периодичность
        add_shape(slide, MSO_SHAPE.RECTANGLE,
                  cx, ry, col_widths[3], row_h,
                  fill_color=bg, line_color=GRAY, line_width=Pt(0.3))
        add_textbox(slide, cx + Inches(0.1), ry, col_widths[3] - Inches(0.2), row_h,
                    period, font_size=9, color=GRAY, bold=False,
                    align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE, italic=True)

    # Итоговая полоса
    summary_y = table_y + Inches(0.45) + row_h * len(kpis) + Inches(0.15)
    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
              Inches(0.4), summary_y, SLIDE_W - Inches(0.8), Inches(0.5),
              fill_color=DARK_BLUE, line_color=None)
    add_textbox(slide, Inches(0.4), summary_y, SLIDE_W - Inches(0.8), Inches(0.5),
                'Итог: 2 KPI на этапе 1 · 3 KPI на этапе 2 · 5 KPI на этапе 3 — все метрики соблюдаются ежемесячно',
                font_size=11, color=WHITE, bold=True,
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    footer(slide)


# ==============================================================
# СЛАЙД 13: ФИНАЛ
# ==============================================================

def slide_final(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H,
              fill_color=RED, line_color=None)

    add_shape(slide, MSO_SHAPE.RECTANGLE, 0, Inches(2.5), SLIDE_W, Inches(0.10),
              fill_color=YELLOW, line_color=None)

    add_textbox(slide, Inches(0.5), Inches(0.8), SLIDE_W - Inches(1.0), Inches(1.5),
                'TITAN', font_size=96, color=WHITE, bold=True,
                align=PP_ALIGN.CENTER, font_name='Arial')

    add_textbox(slide, Inches(0.5), Inches(2.8), SLIDE_W - Inches(1.0), Inches(1.0),
                '10 лет без перекраски — 10 лет экономии!',
                font_size=36, color=YELLOW, bold=True,
                align=PP_ALIGN.CENTER, font_name='Arial')

    add_textbox(slide, Inches(0.5), Inches(4.0), SLIDE_W - Inches(1.0), Inches(0.8),
                'Начнём работу вместе?',
                font_size=28, color=WHITE, bold=True,
                align=PP_ALIGN.CENTER, font_name='Arial')

    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
              Inches(2.5), Inches(5.0), Inches(8.3), Inches(1.6),
              fill_color=WHITE, line_color=YELLOW, line_width=Pt(3))

    add_textbox(slide, Inches(2.5), Inches(5.1), Inches(8.3), Inches(0.5),
                'Empils  •  TITAN',
                font_size=20, color=RED, bold=True,
                align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(2.5), Inches(5.6), Inches(8.3), Inches(0.5),
                'www.titanpaints.ru  •  www.empils.ru',
                font_size=16, color=DARK, bold=True,
                align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(2.5), Inches(6.05), Inches(8.3), Inches(0.5),
                '+7 (863) 203-71-32  •  Ростов-на-Дону, пер. 1-й Машиностроительный, 21',
                font_size=11, color=GRAY, bold=False,
                align=PP_ALIGN.CENTER, italic=True)

    add_textbox(slide, Inches(0.5), Inches(6.95), SLIDE_W - Inches(1.0), Inches(0.5),
                'Спасибо за внимание  •  2026',
                font_size=12, color=WHITE, bold=False,
                align=PP_ALIGN.CENTER, italic=True)


# ==============================================================
# СБОРКА
# ==============================================================

def main():
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H

    slide_title(prs)                    # 1
    slide_about_company(prs)            # 2
    slide_titan_specialization(prs)     # 3
    slide_partner_search(prs)           # 4 — Поиск нового партнёра
    slide_two_scenarios(prs)            # 5
    slide_stage_1(prs)                  # 6 — Первичные + что делает дистрибьютор
    slide_stage_2(prs)                  # 7 — Вторичные + что делает дистрибьютор
    slide_stage_3_main(prs)             # 8 — Третичные + что делает дистрибьютор (НОВЫЙ)
    slide_stage_3_intro(prs)            # 9 — 9 инструментов с визуализацией POSM
    slide_stage_3_detail_1(prs)         # 10 — POS-материалы (1-5)
    slide_stage_3_detail_2(prs)         # 11 — Контент (6-9 + доп)
    slide_guarantee(prs)                # 12 — Гарантия результата
    slide_obligations(prs)              # 13 — Обязательства сторон
    slide_kpi(prs)                      # 14 — KPI и контроль
    slide_final(prs)                    # 15 — Финал

    prs.save(OUTPUT)
    print(f'OK: {OUTPUT}')
    print(f'Слайдов: {len(prs.slides)}')


if __name__ == '__main__':
    main()
