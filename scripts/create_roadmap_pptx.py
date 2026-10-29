#!/usr/bin/env python3
"""
TITAN — Дорожная карта развития продаж.
B2B-презентация для дистрибьюторов.
Цикл: О компании → TITAN специализация → Сводная матрица →
       3 этапа (первичные/вторичные/третичные) → 2 сценария →
       Гарантия результата → Финал.
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
SHELF_BG = RGBColor(0xF3, 0xF4, 0xF6)
SHELF_BROWN = RGBColor(0x8B, 0x6F, 0x47)
ORANGE = RGBColor(0xF4, 0x79, 0x20)  # фирменный оранжевый Эмпилс
YELLOW = RGBColor(0xFB, 0xBF, 0x24)
GOLD = RGBColor(0xB8, 0x86, 0x0B)
BLUE = RGBColor(0x25, 0x63, 0xEB)
DARK_BLUE = RGBColor(0x1E, 0x3A, 0x5F)
LIGHT_BLUE = RGBColor(0xDB, 0xEA, 0xFE)
LIGHT_GREEN = RGBColor(0xDC, 0xFC, 0xE7)
LIGHT_ORANGE = RGBColor(0xFF, 0xED, 0xD5)
LIGHT_RED = RGBColor(0xFE, 0xE2, 0xE2)

OUTPUT = '/home/z/my-project/download/Дорожная_карта_TITAN.pptx'

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
    """Список с буллетами."""
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
        # Буллет
        r1 = p.add_run()
        r1.text = '•  '
        r1.font.size = Pt(font_size)
        r1.font.bold = True
        r1.font.color.rgb = bullet_color
        r1.font.name = 'Arial'
        # Текст
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

    # Левая красная плашка
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, Inches(5.5), SLIDE_H,
              fill_color=RED, line_color=None)
    add_shape(slide, MSO_SHAPE.RECTANGLE, Inches(5.5), 0, Inches(0.10), SLIDE_H,
              fill_color=YELLOW, line_color=None)

    # TITAN
    add_textbox(slide, Inches(0.4), Inches(1.0), Inches(5.0), Inches(1.2),
                'TITAN', font_size=80, color=WHITE, bold=True,
                align=PP_ALIGN.LEFT, font_name='Arial')

    # Подзаголовок
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
                '3 этапа · 2 сценария · 30+ инструментов',
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

    # 3 этапа — карточки
    add_textbox(slide, Inches(6.0), Inches(3.15), Inches(7.0), Inches(0.5),
                '3 этапа развития продаж',
                font_size=18, color=DARK, bold=False, align=PP_ALIGN.LEFT, italic=True)

    stages = [
        ('1', 'ПЕРВИЧНЫЕ', 'Завод → Дистрибьютор', RED),
        ('2', 'ВТОРИЧНЫЕ', 'Дистрибьютор → Розница', ORANGE),
        ('3', 'ТРЕТИЧНЫЕ', 'Розница → Покупатель', GREEN),
    ]
    card_y = Inches(3.9)
    card_w = Inches(2.3)
    card_h = Inches(2.4)
    for i, (num, name, sub, color) in enumerate(stages):
        cx = Inches(6.0) + (card_w + Inches(0.15)) * i
        add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
                  cx, card_y, card_w, card_h,
                  fill_color=WHITE, line_color=color, line_width=Pt(2.5))
        add_shape(slide, MSO_SHAPE.RECTANGLE,
                  cx, card_y, card_w, Inches(0.5),
                  fill_color=color, line_color=None)
        add_textbox(slide, cx, card_y + Inches(0.05),
                    card_w, Inches(0.4),
                    f'ЭТАП {num}', font_size=14, color=WHITE, bold=True,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_textbox(slide, cx + Inches(0.1), card_y + Inches(0.7),
                    card_w - Inches(0.2), Inches(0.5),
                    name, font_size=14, color=DARK, bold=True,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_textbox(slide, cx + Inches(0.1), card_y + Inches(1.25),
                    card_w - Inches(0.2), Inches(0.5),
                    sub, font_size=10, color=GRAY, bold=False,
                    align=PP_ALIGN.CENTER)
        # Большая цифра
        add_textbox(slide, cx + Inches(0.1), card_y + Inches(1.7),
                    card_w - Inches(0.2), Inches(0.6),
                    ['10+', '15+', '12+'][i] + ' инструментов',
                    font_size=11, color=color, bold=True,
                    align=PP_ALIGN.CENTER)

    # Подвал
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

    # 4 крупные метрики
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
        # Большая цифра
        add_textbox(slide, cx, metric_y + Inches(0.15),
                    metric_w, Inches(1.0),
                    num, font_size=48, color=color, bold=True,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        # Подпись
        add_textbox(slide, cx + Inches(0.1), metric_y + Inches(1.25),
                    metric_w - Inches(0.2), Inches(0.9),
                    label, font_size=11, color=DARK, bold=False,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.TOP)

    # География бизнеса
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

    # Производство
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

    # Левая часть — слоган + описание
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

    # Слоган внизу
    add_shape(slide, MSO_SHAPE.RECTANGLE,
              Inches(0.6), Inches(5.7), Inches(5.6), Inches(0.9),
              fill_color=YELLOW, line_color=None)
    add_textbox(slide, Inches(0.6), Inches(5.7), Inches(5.6), Inches(0.9),
                '10 лет без перекраски —\n10 лет экономии!',
                font_size=20, color=RED, bold=True,
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # Правая часть — 3 линейки TITAN
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
        # Карточка
        add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
                  Inches(6.7), cy, Inches(6.5), line_h,
                  fill_color=WHITE, line_color=color, line_width=Pt(2))
        # Цветная полоска слева
        add_shape(slide, MSO_SHAPE.RECTANGLE,
                  Inches(6.7), cy, Inches(0.15), line_h,
                  fill_color=color, line_color=None)
        # Название
        add_textbox(slide, Inches(7.0), cy + Inches(0.10),
                    Inches(4.0), Inches(0.5),
                    name, font_size=15, color=color, bold=True,
                    align=PP_ALIGN.LEFT)
        # Бейдж
        add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
                  Inches(11.0), cy + Inches(0.10),
                  Inches(2.0), Inches(0.35),
                  fill_color=color, line_color=None)
        add_textbox(slide, Inches(11.0), cy + Inches(0.10),
                    Inches(2.0), Inches(0.35),
                    badge, font_size=10, color=WHITE, bold=True,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        # Описание
        add_textbox(slide, Inches(7.0), cy + Inches(0.60),
                    Inches(6.0), Inches(0.7),
                    desc, font_size=11, color=DARK, bold=False,
                    align=PP_ALIGN.LEFT)
        # Объёмы
        add_textbox(slide, Inches(7.0), cy + Inches(1.25),
                    Inches(6.0), Inches(0.35),
                    'Фасовка: ' + vols, font_size=10, color=GRAY, bold=False,
                    align=PP_ALIGN.LEFT, italic=True)

    footer(slide)


# ==============================================================
# СЛАЙД 4: СВОДНАЯ МАТРИЦА TITAN ПО КЛАССАМ ТТ
# ==============================================================

def slide_mml_matrix(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H,
              fill_color=WHITE, line_color=None)

    header_band(slide, 'СВОДНАЯ МАТРИЦА TITAN ПО КЛАССАМ ТТ',
                'MML — Minimal Must List  •  4 канала продаж', color=DARK)

    # Таблица
    table_x = Inches(0.4)
    table_y = Inches(1.0)
    table_w = SLIDE_W - Inches(0.8)

    # Заголовок
    headers = ['Линейка / Объём', 'A — Федеральная DIY', 'B — Локальная сеть', 'C — Супермаркет', 'D — Хозмаг']
    col_widths = [Inches(3.0), Inches(2.6), Inches(2.6), Inches(2.4), Inches(2.4)]
    row_h = Inches(0.55)

    # Заголовки
    for i, h in enumerate(headers):
        cx = table_x + sum(col_widths[:i], Inches(0))
        bg_color = [DARK, RED, DARK, ORANGE, GREEN][i]
        add_shape(slide, MSO_SHAPE.RECTANGLE,
                  cx, table_y, col_widths[i], row_h,
                  fill_color=bg_color, line_color=WHITE, line_width=Pt(1))
        add_textbox(slide, cx, table_y, col_widths[i], row_h,
                    h, font_size=11, color=WHITE, bold=True,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # Данные
    data = [
        ('3в1 по ржавчине, 0.75 л', '5 цветов (9005, 7004, 8017, 7024, 9010)', '5 цветов (9005, 7004, 8017, 7024, 9010)', '—', '—'),
        ('3в1 по ржавчине, 2 л', '5 цветов (9005, 7004, 8017, 7024, 9010)', '4 цвета (9005, 7004, 8017, 7024)', '—', '—'),
        ('Быстросохнущая, 0.75 л', '1 цвет (9005)', '1 цвет (9005)', '—', '—'),
        ('Быстросохнущая, 2.2 л', '1 цвет (9005)', '—', '—', '—'),
        ('Молотковая, 0.8 кг', '3 цвета (чёрный, коричневый, серебристый)', '2 цвета (чёрный, коричневый)', '2 цвета (чёрный, коричневый)', '—'),
        ('Молотковая, 2.5 кг', '1 цвет (чёрный)', '—', '—', '—'),
    ]
    for ri, row in enumerate(data):
        ry = table_y + row_h * (ri + 1)
        bg = LIGHT_GRAY if ri % 2 == 0 else WHITE
        for ci, val in enumerate(row):
            cx = table_x + sum(col_widths[:ci], Inches(0))
            add_shape(slide, MSO_SHAPE.RECTANGLE,
                      cx, ry, col_widths[ci], row_h,
                      fill_color=bg, line_color=GRAY, line_width=Pt(0.5))
            # Цвет для пустых значений
            val_color = RED if val == '—' else DARK
            val_bold = (ci == 0)
            add_textbox(slide, cx + Inches(0.05), ry,
                        col_widths[ci] - Inches(0.10), row_h,
                        val, font_size=10, color=val_color, bold=val_bold,
                        align=PP_ALIGN.LEFT if ci == 0 else PP_ALIGN.CENTER,
                        anchor=MSO_ANCHOR.MIDDLE)

    # ИТОГО
    ry = table_y + row_h * (len(data) + 1)
    totals = ['ИТОГО SKU TITAN', '16', '12', '2', '0']
    for ci, val in enumerate(totals):
        cx = table_x + sum(col_widths[:ci], Inches(0))
        bg = DARK if ci == 0 else [RED, DARK, ORANGE, GREEN][ci-1]
        add_shape(slide, MSO_SHAPE.RECTANGLE,
                  cx, ry, col_widths[ci], row_h,
                  fill_color=bg, line_color=WHITE, line_width=Pt(1))
        val_color = WHITE
        add_textbox(slide, cx, ry, col_widths[ci], row_h,
                    val, font_size=14, color=val_color, bold=True,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # Проблема
    problem_y = ry + row_h + Inches(0.3)
    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
              Inches(0.4), problem_y, SLIDE_W - Inches(0.8), Inches(1.4),
              fill_color=LIGHT_RED, line_color=RED, line_width=Pt(2))

    add_textbox(slide, Inches(0.6), problem_y + Inches(0.10),
                Inches(12.0), Inches(0.4),
                '⚠  ГЛАВНАЯ ПРОБЛЕМА: TITAN не доходит до маленьких точек',
                font_size=14, color=RED, bold=True, align=PP_ALIGN.LEFT)
    add_textbox(slide, Inches(0.6), problem_y + Inches(0.55),
                Inches(12.0), Inches(0.8),
                'В классах C и D (233 точки — 80% розничного канала) TITAN 3в1 отсутствует. Покупатель в супермаркете или хозмаге не может купить продукт, на котором построена вся ТСО-методология. Рекомендуем добавить 5 SKU в класс C и 3 SKU в класс D.',
                font_size=11, color=DARK, bold=False, align=PP_ALIGN.LEFT)

    footer(slide)


# ==============================================================
# СЛАЙД 5: ДОРОЖНАЯ КАРТА — 3 ЭТАПА
# ==============================================================

def slide_roadmap(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H,
              fill_color=WHITE, line_color=None)

    header_band(slide, 'ДОРОЖНАЯ КАРТА — 3 ЭТАПА РАЗВИТИЯ ПРОДАЖ',
                'От завода до конечного покупателя', color=DARK_BLUE)

    # 3 больших блока — этапы
    stages = [
        ('1', 'ПЕРВИЧНЫЕ\nПРОДАЖИ', 'Завод → Дистрибьютор',
         'Мотивация дистрибьютора закупить и держать остатки TITAN',
         ['Ретробонус 8% + 3%', 'Отсрочка 70-90 дней', 'Выездное обучение',
          'Экскурсия на завод', 'Ценовая защита 3-6 мес.'],
         RED),
        ('2', 'ВТОРИЧНЫЕ\nПРОДАЖИ', 'Дистрибьютор → Розница',
         'Мотивация торговых представителей и розничных точек',
         ['Бонус за новую точку', 'Квартальный конкурс ТП',
          'Калькулятор ТСО для ТП', 'Набор выкрасов',
          'Бесплатные POS-материалы'],
         ORANGE),
        ('3', 'ТРЕТИЧНЫЕ\nПРОДАЖИ', 'Розница → Покупатель',
         'Мотивация покупателя выбрать TITAN вместо ПФ-115',
         ['Шпаргалка продавца', 'Конкурс «Продавец ТСО»',
          'QR-код на банке', 'Выкрасы «до/после»',
          'Гарантия 10 лет на банке'],
         GREEN),
    ]

    block_y = Inches(1.0)
    block_w = Inches(4.1)
    block_h = Inches(5.6)

    for i, (num, name, sub, goal, tools, color) in enumerate(stages):
        cx = Inches(0.4) + (block_w + Inches(0.20)) * i

        # Карточка
        add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
                  cx, block_y, block_w, block_h,
                  fill_color=WHITE, line_color=color, line_width=Pt(2.5))

        # Цветная шапка
        header_h = Inches(1.3)
        add_shape(slide, MSO_SHAPE.RECTANGLE,
                  cx, block_y, block_w, header_h,
                  fill_color=color, line_color=None)

        # Номер этапа
        add_textbox(slide, cx + Inches(0.2), block_y + Inches(0.10),
                    Inches(1.0), Inches(0.5),
                    f'ЭТАП {num}', font_size=12, color=WHITE, bold=True,
                    align=PP_ALIGN.LEFT)

        # Название
        add_textbox(slide, cx + Inches(0.2), block_y + Inches(0.45),
                    block_w - Inches(0.4), Inches(0.9),
                    name, font_size=18, color=WHITE, bold=True,
                    align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)

        # Подзаголовок (Завод → Дистрибьютор)
        add_textbox(slide, cx + Inches(0.2), block_y + header_h + Inches(0.10),
                    block_w - Inches(0.4), Inches(0.4),
                    sub, font_size=14, color=color, bold=True,
                    align=PP_ALIGN.CENTER)

        # Цель
        add_textbox(slide, cx + Inches(0.2), block_y + header_h + Inches(0.55),
                    block_w - Inches(0.4), Inches(0.7),
                    goal, font_size=11, color=DARK, bold=False,
                    align=PP_ALIGN.CENTER, italic=True)

        # Разделитель
        add_shape(slide, MSO_SHAPE.RECTANGLE,
                  cx + Inches(0.5), block_y + header_h + Inches(1.4),
                  block_w - Inches(1.0), Inches(0.03),
                  fill_color=color, line_color=None)

        # Инструменты
        add_textbox(slide, cx + Inches(0.2), block_y + header_h + Inches(1.55),
                    block_w - Inches(0.4), Inches(0.4),
                    'КЛЮЧЕВЫЕ ИНСТРУМЕНТЫ:',
                    font_size=10, color=color, bold=True, align=PP_ALIGN.LEFT)

        add_bullet_list(slide, cx + Inches(0.2), block_y + header_h + Inches(2.0),
                        block_w - Inches(0.4), Inches(2.5),
                        tools, font_size=11, color=DARK, bullet_color=color)

    # Стрелки между этапами
    for i in range(2):
        arrow_x = Inches(0.4) + block_w * (i + 1) + Inches(0.20) * i + Inches(0.05)
        add_shape(slide, MSO_SHAPE.RIGHT_ARROW,
                  arrow_x, Inches(3.5), Inches(0.10), Inches(0.5),
                  fill_color=GRAY, line_color=None)

    # Итоговая полоса
    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
              Inches(0.4), Inches(6.75), SLIDE_W - Inches(0.8), Inches(0.3),
              fill_color=DARK_BLUE, line_color=None)
    add_textbox(slide, Inches(0.4), Inches(6.75), SLIDE_W - Inches(0.8), Inches(0.3),
                'Гарантия Эмпилс: выполнение всех 3 этапов → рост продаж TITAN на 25-40% за 12 месяцев',
                font_size=12, color=WHITE, bold=True,
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)


# ==============================================================
# СЛАЙД 6: 2 СЦЕНАРИЯ РАБОТЫ
# ==============================================================

def slide_two_scenarios(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H,
              fill_color=WHITE, line_color=None)

    header_band(slide, 'ДВА СЦЕНАРИЯ РАБОТЫ С ДИСТРИБЬЮТОРОМ',
                'От «убедить попробовать» до «увеличить продажи»', color=DARK)

    # 2 крупных блока
    scenarios = [
        ('А', 'НОВЫЙ ПАРТНЁР', '«Убедить попробовать»',
         'Дистрибьютор не хочет заводить TITAN — у него уже есть аналог. Главный барьер — страх: «Закуплю — не продам, заморожу деньги, потеряю время».',
         'Снять страх и сделать первую закупку максимально безболезненной',
         ['Бонус за первичную закупку 5-10%', 'Возврат неликвида в 90 дней',
          'Отсрочка 70-90 дней', 'Бесплатные POS-материалы',
          'Выездное обучение команды ТП', 'Экскурсия на завод'],
         RED, LIGHT_RED),
        ('Б', 'ДЕЙСТВУЮЩИЙ ПАРТНЁР', '«Увеличить продажи»',
         'Дистрибьютор уже продаёт TITAN 3в1, но продажи не растут или нестабильны. Он знает продукт, но не вкладывается в его продвижение.',
         'Мотивировать ТП, дать инструменты рознице, обучить продавцов',
         ['Ретробонус 8% + 3%', 'Прогрессивная скидка 3/5/7%',
          'Co-op маркетинг 50%', 'Сертификация ТП по ТСО',
          'Квартальный конкурс ТП', 'CRM-интеграция + аналитика'],
         GREEN, LIGHT_GREEN),
    ]

    block_y = Inches(1.0)
    block_w = Inches(6.2)
    block_h = Inches(5.7)

    for i, (letter, name, sub, problem, goal, tools, color, bg_color) in enumerate(scenarios):
        cx = Inches(0.4) + (block_w + Inches(0.30)) * i

        # Карточка
        add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
                  cx, block_y, block_w, block_h,
                  fill_color=bg_color, line_color=color, line_width=Pt(2.5))

        # Буква сценария
        add_shape(slide, MSO_SHAPE.OVAL,
                  cx + Inches(0.3), block_y + Inches(0.3),
                  Inches(0.8), Inches(0.8),
                  fill_color=color, line_color=None)
        add_textbox(slide, cx + Inches(0.3), block_y + Inches(0.3),
                    Inches(0.8), Inches(0.8),
                    letter, font_size=36, color=WHITE, bold=True,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

        # Название
        add_textbox(slide, cx + Inches(1.3), block_y + Inches(0.30),
                    block_w - Inches(1.5), Inches(0.5),
                    name, font_size=20, color=color, bold=True,
                    align=PP_ALIGN.LEFT)
        add_textbox(slide, cx + Inches(1.3), block_y + Inches(0.75),
                    block_w - Inches(1.5), Inches(0.4),
                    sub, font_size=13, color=DARK, bold=False,
                    align=PP_ALIGN.LEFT, italic=True)

        # Проблема
        add_textbox(slide, cx + Inches(0.3), block_y + Inches(1.5),
                    block_w - Inches(0.6), Inches(1.3),
                    problem, font_size=11, color=DARK, bold=False,
                    align=PP_ALIGN.LEFT)

        # Цель
        goal_y = block_y + Inches(2.7)
        add_shape(slide, MSO_SHAPE.RECTANGLE,
                  cx + Inches(0.3), goal_y,
                  block_w - Inches(0.6), Inches(0.5),
                  fill_color=color, line_color=None)
        add_textbox(slide, cx + Inches(0.3), goal_y,
                    block_w - Inches(0.6), Inches(0.5),
                    'ЦЕЛЬ: ' + goal, font_size=11, color=WHITE, bold=True,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

        # Инструменты
        add_textbox(slide, cx + Inches(0.3), block_y + Inches(3.4),
                    block_w - Inches(0.6), Inches(0.4),
                    'ИНСТРУМЕНТЫ:',
                    font_size=11, color=color, bold=True, align=PP_ALIGN.LEFT)

        add_bullet_list(slide, cx + Inches(0.3), block_y + Inches(3.8),
                        block_w - Inches(0.6), Inches(1.8),
                        tools, font_size=11, color=DARK, bullet_color=color)


# ==============================================================
# СЛАЙД 7: ЭТАП 1 — ПЕРВИЧНЫЕ ПРОДАЖИ
# ==============================================================

def slide_stage_1(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H,
              fill_color=WHITE, line_color=None)

    header_band(slide, 'ЭТАП 1 — ПЕРВИЧНЫЕ ПРОДАЖИ',
                'Завод → Дистрибьютор  •  10+ инструментов', color=RED)

    # Левая часть — инструменты
    add_textbox(slide, Inches(0.4), Inches(0.95),
                Inches(8.0), Inches(0.5),
                'ИНСТРУМЕНТЫ МОТИВАЦИИ ДИСТРИБЬЮТОРА',
                font_size=14, color=RED, bold=True, align=PP_ALIGN.LEFT)

    tools = [
        ('Ретробонус 8% + 3%', 'Ежемесячно + полугодовой бонус. Регулярная мотивация закупать больше.'),
        ('Прогрессивная скидка 3/5/7%', 'Чем больше закупает — тем выгоднее. Объёмная шкала.'),
        ('Отсрочка 70-90 дней', 'Снижение финансовой нагрузки на старте.'),
        ('Ценовая защита 3-6 мес.', 'Гарантия маржи — дистрибьютор планирует рост.'),
        ('Возврат неликвида в 90 дней', 'Снять страх «куплю и не продам».'),
        ('Co-op маркетинг 50%', 'Совместное промо — инвестиции в рост продаж.'),
        ('Выездное обучение команды ТП', 'Команда умеет продавать с первого дня.'),
        ('Экскурсия на завод', 'Показать производство → лояльность руководства.'),
    ]

    tool_y = Inches(1.45)
    tool_h = Inches(0.62)
    for i, (name, desc) in enumerate(tools):
        cy = tool_y + tool_h * i
        # Цвет чередуется
        bg = LIGHT_GRAY if i % 2 == 0 else WHITE
        add_shape(slide, MSO_SHAPE.RECTANGLE,
                  Inches(0.4), cy, Inches(8.0), tool_h,
                  fill_color=bg, line_color=GRAY, line_width=Pt(0.3))
        # Номер
        add_shape(slide, MSO_SHAPE.OVAL,
                  Inches(0.5), cy + Inches(0.10),
                  Inches(0.40), Inches(0.40),
                  fill_color=RED, line_color=None)
        add_textbox(slide, Inches(0.5), cy + Inches(0.10),
                    Inches(0.40), Inches(0.40),
                    str(i + 1), font_size=12, color=WHITE, bold=True,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        # Название + описание
        add_multi_text(slide, Inches(1.05), cy + Inches(0.05),
                       Inches(7.2), Inches(0.55),
                       [
                           (name, 11, RED, True),
                           (desc, 9, DARK, False),
                       ], anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.LEFT)

    # Правая часть — KPI этапа
    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
              Inches(8.6), Inches(0.95), Inches(4.5), Inches(5.9),
              fill_color=LIGHT_RED, line_color=RED, line_width=Pt(2))

    add_textbox(slide, Inches(8.6), Inches(1.05),
                Inches(4.5), Inches(0.5),
                'KPI ЭТАПА 1',
                font_size=16, color=RED, bold=True,
                align=PP_ALIGN.CENTER)

    kpis = [
        ('+30%', 'рост закупок TITAN 3в1\nдистрибьютором', 'ежемесячно'),
        ('5+', 'цветов RAL в наличии\nна складе дистрибьютора', 'ежеквартально'),
        ('0', 'просрочек платежа\n> 60 дней', 'ежемесячно'),
        ('100%', 'команда ТП прошла\nобучение ТСО', 'к мес. 3'),
    ]
    kpi_y = Inches(1.65)
    kpi_h = Inches(1.25)
    for i, (val, label, period) in enumerate(kpis):
        cy = kpi_y + kpi_h * i
        # Карточка KPI
        add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
                  Inches(8.8), cy, Inches(4.1), kpi_h - Inches(0.10),
                  fill_color=WHITE, line_color=RED, line_width=Pt(1))
        # Значение
        add_textbox(slide, Inches(8.8), cy + Inches(0.05),
                    Inches(1.5), Inches(0.7),
                    val, font_size=24, color=RED, bold=True,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        # Подпись
        add_multi_text(slide, Inches(10.3), cy + Inches(0.05),
                       Inches(2.5), kpi_h - Inches(0.20),
                       [
                           (label, 10, DARK, True),
                           (period, 9, GRAY, False),
                       ], anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.LEFT)

    footer(slide)


# ==============================================================
# СЛАЙД 8: ЭТАП 2 — ВТОРИЧНЫЕ ПРОДАЖИ
# ==============================================================

def slide_stage_2(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H,
              fill_color=WHITE, line_color=None)

    header_band(slide, 'ЭТАП 2 — ВТОРИЧНЫЕ ПРОДАЖИ',
                'Дистрибьютор → Розница  •  15+ инструментов', color=ORANGE)

    add_textbox(slide, Inches(0.4), Inches(0.95),
                Inches(8.0), Inches(0.5),
                'ИНСТРУМЕНТЫ МОТИВАЦИИ ТОРГОВЫХ ПРЕДСТАВИТЕЛЕЙ И РОЗНИЦЫ',
                font_size=14, color=ORANGE, bold=True, align=PP_ALIGN.LEFT)

    tools = [
        ('Бонус за новую точку 500-1000 ₽', 'ТП активно расширяют сеть точек с TITAN.'),
        ('Ретробонус 2-3% для розницы', 'Финансовая мотивация розницы за оборот.'),
        ('Квартальный конкурс ТП (топ-3: 10/7/5 тыс.)', 'Соревнование → эмоции → рост продаж.'),
        ('Калькулятор ТСО для ТП (мобильный)', 'ТСО-инструмент в поле — на телефоне ТП.'),
        ('Набор выкрасов для ТП (портативный)', 'Наглядность = конверсия в магазине.'),
        ('Сравнительная таблица TITAN vs конкуренты (А4)', 'Стандарт качества работы с полкой.'),
        ('Бесплатные POS-материалы + стартовый набор', 'Полка готова к продажам без затрат партнёра.'),
        ('Сертификация ТП по ТСО (бонус 3000 ₽)', 'Превратить знание в навык.'),
        ('Программа для розницы: скидка за 3+ цвета RAL', 'Стимулирование ассортимента.'),
        ('CRM-интеграция + аналитика продаж', 'Данные для принятия решений.'),
    ]

    tool_y = Inches(1.45)
    tool_h = Inches(0.50)
    for i, (name, desc) in enumerate(tools):
        cy = tool_y + tool_h * i
        bg = LIGHT_GRAY if i % 2 == 0 else WHITE
        add_shape(slide, MSO_SHAPE.RECTANGLE,
                  Inches(0.4), cy, Inches(8.0), tool_h,
                  fill_color=bg, line_color=GRAY, line_width=Pt(0.3))
        add_shape(slide, MSO_SHAPE.OVAL,
                  Inches(0.5), cy + Inches(0.08),
                  Inches(0.34), Inches(0.34),
                  fill_color=ORANGE, line_color=None)
        add_textbox(slide, Inches(0.5), cy + Inches(0.08),
                    Inches(0.34), Inches(0.34),
                    str(i + 1), font_size=11, color=WHITE, bold=True,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_multi_text(slide, Inches(0.95), cy + Inches(0.03),
                       Inches(7.3), Inches(0.45),
                       [
                           (name, 11, ORANGE, True),
                           (desc, 9, DARK, False),
                       ], anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.LEFT)

    # KPI
    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
              Inches(8.6), Inches(0.95), Inches(4.5), Inches(5.9),
              fill_color=LIGHT_ORANGE, line_color=ORANGE, line_width=Pt(2))

    add_textbox(slide, Inches(8.6), Inches(1.05),
                Inches(4.5), Inches(0.5),
                'KPI ЭТАПА 2',
                font_size=16, color=ORANGE, bold=True,
                align=PP_ALIGN.CENTER)

    kpis = [
        ('+25%', 'рост активных розничных\nточек с TITAN 3в1', 'ежемесячно'),
        ('100%', 'ТП прошли обучение\nТСО', 'к мес. 3'),
        ('20%+', 'доля новых точек\nот TITAN', 'ежемесячно'),
        ('80%', 'точек с POS-\nматериалами', 'к мес. 6'),
    ]
    kpi_y = Inches(1.65)
    kpi_h = Inches(1.25)
    for i, (val, label, period) in enumerate(kpis):
        cy = kpi_y + kpi_h * i
        add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
                  Inches(8.8), cy, Inches(4.1), kpi_h - Inches(0.10),
                  fill_color=WHITE, line_color=ORANGE, line_width=Pt(1))
        add_textbox(slide, Inches(8.8), cy + Inches(0.05),
                    Inches(1.5), Inches(0.7),
                    val, font_size=24, color=ORANGE, bold=True,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_multi_text(slide, Inches(10.3), cy + Inches(0.05),
                       Inches(2.5), kpi_h - Inches(0.20),
                       [
                           (label, 10, DARK, True),
                           (period, 9, GRAY, False),
                       ], anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.LEFT)

    footer(slide)


# ==============================================================
# СЛАЙД 9: ЭТАП 3 — ТРЕТИЧНЫЕ ПРОДАЖИ
# ==============================================================

def slide_stage_3(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H,
              fill_color=WHITE, line_color=None)

    header_band(slide, 'ЭТАП 3 — ТРЕТИЧНЫЕ ПРОДАЖИ',
                'Розница → Покупатель  •  12+ инструментов', color=GREEN)

    add_textbox(slide, Inches(0.4), Inches(0.95),
                Inches(8.0), Inches(0.5),
                'ИНСТРУМЕНТЫ МОТИВАЦИИ ПОКУПАТЕЛЯ',
                font_size=14, color=GREEN, bold=True, align=PP_ALIGN.LEFT)

    tools = [
        ('Шпаргалка продавца А6 (ламинированная)', 'Скрипт «2 минуты ССВ» — продавец закрывает сделку.'),
        ('Конкурс «Продавец месяца» (5/3/2 тыс. + бейдж)', 'Мотивация продавца рекомендовать TITAN.'),
        ('Бонус продавцу 50-100 ₽ за банку', 'Прямая финансовая мотивация за каждую продажу.'),
        ('Сертификация «Продавец ССВ» (+3000/−3000 ₽)', 'Подтверждение квалификации — доверие покупателя.'),
        ('QR-код на банке → видеоинструкция + калькулятор', 'Покупатель сам узнаёт о TITAN, без продавца.'),
        ('Выкрасы «до/после» + «через год» на полке', 'Наглядная демонстрация качества.'),
        ('Калькулятор ТСО для покупателя (онлайн + QR)', 'Доказательство: TITAN дешевле на 61% за 10 лет.'),
        ('Гарантия 10 лет + сертификаты на банке', 'Доверие → конверсия → повторные покупки.'),
        ('Подарок при покупке (кисть/валик)', 'Стимулирование покупки.'),
        ('Реферальная программа «Приведи соседа»', 'Сарафанное радио → новые покупатели бесплатно.'),
        ('Воблер на полку «10 лет без перекраски»', 'Розница видит выгоду и берёт TITAN на пробу.'),
        ('Иконографика ССВ на банке', 'Покупатель сам видит экономию — без помощи продавца.'),
    ]

    tool_y = Inches(1.45)
    tool_h = Inches(0.42)
    for i, (name, desc) in enumerate(tools):
        cy = tool_y + tool_h * i
        bg = LIGHT_GRAY if i % 2 == 0 else WHITE
        add_shape(slide, MSO_SHAPE.RECTANGLE,
                  Inches(0.4), cy, Inches(8.0), tool_h,
                  fill_color=bg, line_color=GRAY, line_width=Pt(0.3))
        add_shape(slide, MSO_SHAPE.OVAL,
                  Inches(0.5), cy + Inches(0.06),
                  Inches(0.30), Inches(0.30),
                  fill_color=GREEN, line_color=None)
        add_textbox(slide, Inches(0.5), cy + Inches(0.06),
                    Inches(0.30), Inches(0.30),
                    str(i + 1), font_size=10, color=WHITE, bold=True,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_multi_text(slide, Inches(0.90), cy + Inches(0.02),
                       Inches(7.4), Inches(0.38),
                       [
                           (name + ' — ', 10, GREEN, True),
                           (desc, 9, DARK, False),
                       ], anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.LEFT)

    # KPI
    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
              Inches(8.6), Inches(0.95), Inches(4.5), Inches(5.9),
              fill_color=LIGHT_GREEN, line_color=GREEN, line_width=Pt(2))

    add_textbox(slide, Inches(8.6), Inches(1.05),
                Inches(4.5), Inches(0.5),
                'KPI ЭТАПА 3',
                font_size=16, color=GREEN, bold=True,
                align=PP_ALIGN.CENTER)

    kpis = [
        ('30%+', 'конверсия в покупку\nTITAN', 'ежеквартально'),
        ('500+', 'сканирований QR\nна банке/мес', 'ежемесячно'),
        ('20+', 'сертифицированных\nпродавцов ТСО', 'к мес. 6'),
        ('-30%', 'снижение возвратов\nTITAN 3в1', 'ежеквартально'),
    ]
    kpi_y = Inches(1.65)
    kpi_h = Inches(1.25)
    for i, (val, label, period) in enumerate(kpis):
        cy = kpi_y + kpi_h * i
        add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
                  Inches(8.8), cy, Inches(4.1), kpi_h - Inches(0.10),
                  fill_color=WHITE, line_color=GREEN, line_width=Pt(1))
        add_textbox(slide, Inches(8.8), cy + Inches(0.05),
                    Inches(1.5), Inches(0.7),
                    val, font_size=24, color=GREEN, bold=True,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_multi_text(slide, Inches(10.3), cy + Inches(0.05),
                       Inches(2.5), kpi_h - Inches(0.20),
                       [
                           (label, 10, DARK, True),
                           (period, 9, GRAY, False),
                       ], anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.LEFT)

    footer(slide)


# ==============================================================
# СЛАЙД 10: ГАРАНТИЯ РЕЗУЛЬТАТА
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
                'При выполнении всех 3 этапов — рост продаж TITAN на 25-40% за 12 месяцев',
                font_size=22, color=WHITE, bold=True,
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_textbox(slide, Inches(0.4), Inches(1.65), SLIDE_W - Inches(0.8), Inches(0.5),
                'ROI программы: окупаемость за 6-8 месяцев при минимальном сценарии',
                font_size=14, color=YELLOW, bold=True,
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # 4 карточки — что получает дистрибьютор
    add_textbox(slide, Inches(0.4), Inches(2.6), Inches(12.5), Inches(0.5),
                'ЧТО ПОЛУЧАЕТ ДИСТРИБЬЮТОР',
                font_size=14, color=DARK_BLUE, bold=True, align=PP_ALIGN.LEFT)

    benefits = [
        ('💰', 'Финансовая выгода',
         'Ретробонусы 8-11%, отсрочка 70-90 дней, ценовая защита. Доп. выручка при росте 25% окупает вложения за 6-8 мес.',
         GREEN),
        ('🎓', 'Обученная команда',
         '100% ТП сертифицированы по ТСО. Продавцы розницы знают аргументы vs ПФ-115. Конверсия в продажу TITAN — 30%+.',
         BLUE),
        ('📊', 'Аналитика и контроль',
         'CRM-интеграция, отчётность по QR-сканированиям, аудит полок. Данные для принятия решений по каждой точке.',
         ORANGE),
        ('🏗', 'Маркетинговая поддержка',
         'POS-материалы, выкрасы, воблеры, планограммы, контент для соцсетей. Полка готова к продажам без затрат партнёра.',
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
        # Цветная шапка
        add_shape(slide, MSO_SHAPE.RECTANGLE,
                  cx, ben_y, ben_w, Inches(0.6),
                  fill_color=color, line_color=None)
        # Эмодзи + название в шапке
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
        # Описание
        add_textbox(slide, cx + Inches(0.15), ben_y + Inches(0.75),
                    ben_w - Inches(0.30), ben_h - Inches(0.85),
                    desc, font_size=11, color=DARK, bold=False,
                    align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP)

    # Финальная полоса
    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
              Inches(0.4), Inches(6.55), SLIDE_W - Inches(0.8), Inches(0.5),
              fill_color=YELLOW, line_color=None)
    add_textbox(slide, Inches(0.4), Inches(6.55), SLIDE_W - Inches(0.8), Inches(0.5),
                'Empils берёт на себя обучение, POS-материалы и аналитику. Дистрибьютор — закупку и выкладку.',
                font_size=13, color=DARK, bold=True,
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)


# ==============================================================
# СЛАЙД 11: СЛЕДУЮЩИЕ ШАГИ
# ==============================================================

def slide_next_steps(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H,
              fill_color=WHITE, line_color=None)

    header_band(slide, 'СЛЕДУЮЩИЕ ШАГИ',
                'Таймлайн внедрения — 12 месяцев', color=DARK_BLUE)

    # Таймлайн
    timeline = [
        ('Мес. 0-2', 'ПОДГОТОВКА',
         'Создание всех материалов: калькулятор ТСО, шпаргалки, выкрасы, QR-коды, обучающие видео, планограммы',
         DARK_BLUE),
        ('Мес. 1-4', 'ПИЛОТ',
         'Запуск на 1-2 дистрибьюторах. Выездное обучение торговых команд. Первичная закупка по планограмме',
         RED),
        ('Мес. 2-4', 'ОБУЧЕНИЕ',
         'Выездное обучение торговых команд пилотных дистрибьюторов. Сертификация ТП по ТСО',
         ORANGE),
        ('Мес. 3-4', 'POS + КАЛЬКУЛЯТОР',
         'Запуск POS-материалов и онлайн-калькулятора ТСО. QR-коды на банках. Выкрасы на полках',
         GREEN),
        ('Мес. 5-8', 'МАСШТАБИРОВАНИЕ',
         'Расширение на 3-5 дистрибьюторов. Программа мотивации продавцов розницы. Сертификация «Продавец ТСО»',
         DARK_BLUE),
        ('Мес. 9-12', 'АНАЛИЗ',
         'Анализ результатов, корректировка, принятие решения о полном масштабировании на все 100+ дистрибьюторов',
         RED),
    ]

    tl_y = Inches(1.1)
    tl_h = Inches(0.85)
    for i, (period, name, desc, color) in enumerate(timeline):
        cy = tl_y + tl_h * i
        # Цветная плашка с периодом
        add_shape(slide, MSO_SHAPE.RECTANGLE,
                  Inches(0.4), cy, Inches(2.0), tl_h - Inches(0.10),
                  fill_color=color, line_color=None)
        add_multi_text(slide, Inches(0.4), cy + Inches(0.05),
                       Inches(2.0), tl_h - Inches(0.20),
                       [
                           (period, 14, WHITE, True),
                           (name, 11, YELLOW, True),
                       ], anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)
        # Описание
        add_shape(slide, MSO_SHAPE.RECTANGLE,
                  Inches(2.5), cy, Inches(10.4), tl_h - Inches(0.10),
                  fill_color=LIGHT_GRAY, line_color=color, line_width=Pt(1))
        add_textbox(slide, Inches(2.7), cy + Inches(0.05),
                    Inches(10.0), tl_h - Inches(0.20),
                    desc, font_size=11, color=DARK, bold=False,
                    align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)

    # Итоговая плашка
    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
              Inches(0.4), Inches(6.5), SLIDE_W - Inches(0.8), Inches(0.5),
              fill_color=GREEN, line_color=None)
    add_textbox(slide, Inches(0.4), Inches(6.5), SLIDE_W - Inches(0.8), Inches(0.5),
                'Результат 12 месяцев: +25-40% к продажам TITAN 3в1 · 100% ТП обучены · 80% точек с POS',
                font_size=13, color=WHITE, bold=True,
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)


# ==============================================================
# СЛАЙД 12: ФИНАЛ
# ==============================================================

def slide_final(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H,
              fill_color=RED, line_color=None)

    # Жёлтая полоса
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0, Inches(2.5), SLIDE_W, Inches(0.10),
              fill_color=YELLOW, line_color=None)

    # TITAN
    add_textbox(slide, Inches(0.5), Inches(0.8), SLIDE_W - Inches(1.0), Inches(1.5),
                'TITAN', font_size=96, color=WHITE, bold=True,
                align=PP_ALIGN.CENTER, font_name='Arial')

    # Слоган
    add_textbox(slide, Inches(0.5), Inches(2.8), SLIDE_W - Inches(1.0), Inches(1.0),
                '10 лет без перекраски — 10 лет экономии!',
                font_size=36, color=YELLOW, bold=True,
                align=PP_ALIGN.CENTER, font_name='Arial')

    # Призыв
    add_textbox(slide, Inches(0.5), Inches(4.0), SLIDE_W - Inches(1.0), Inches(0.8),
                'Начнём работу вместе?',
                font_size=28, color=WHITE, bold=True,
                align=PP_ALIGN.CENTER, font_name='Arial')

    # Контакты
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

    # Подвал
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
    slide_mml_matrix(prs)               # 4
    slide_roadmap(prs)                  # 5
    slide_two_scenarios(prs)            # 6
    slide_stage_1(prs)                  # 7
    slide_stage_2(prs)                  # 8
    slide_stage_3(prs)                  # 9
    slide_guarantee(prs)                # 10
    slide_next_steps(prs)               # 11
    slide_final(prs)                    # 12

    prs.save(OUTPUT)
    print(f'OK: {OUTPUT}')
    print(f'Слайдов: {len(prs.slides)}')


if __name__ == '__main__':
    main()
