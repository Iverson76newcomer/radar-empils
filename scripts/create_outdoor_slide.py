#!/usr/bin/env python3
"""
Слайд: Комплексное продвижение TITAN — Outdoor + Экстремальный тест.
3 инструмента: Экстремальный тест 365 дней, Яндекс.Директ, RuTube.
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
BLUE = RGBColor(0x25, 0x63, 0xEB)
LIGHT_GREEN = RGBColor(0xDC, 0xFC, 0xE7)
LIGHT_ORANGE = RGBColor(0xFF, 0xED, 0xD5)
LIGHT_RED = RGBColor(0xFE, 0xE2, 0xE2)
LIGHT_BLUE = RGBColor(0xDB, 0xEA, 0xFE)

OUTPUT = '/home/z/my-project/download/Слайд_outdoor_продвижение.pptx'

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
                'ПРОДВИЖЕНИЕ TITAN — OUTDOOR + ЭКСТРЕМАЛЬНЫЙ ТЕСТ',
                font_size=20, color=WHITE, bold=True,
                align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)
    add_textbox(slide, SLIDE_W - Inches(5.0), Inches(0.08),
                Inches(4.6), Inches(0.5),
                '3 инструмента · за пределами полки',
                font_size=11, color=YELLOW, bold=False,
                align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)

    # === 3 КАРТОЧКИ — по горизонтали ===
    card_w = Inches(4.10)
    card_h = Inches(5.90)
    card_y = Inches(0.85)
    gap = Inches(0.15)

    # ===== КАРТОЧКА 1: ЭКСТРЕМАЛЬНЫЙ ТЕСТ 365 ДНЕЙ =====
    cx1 = Inches(0.4)
    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
              cx1, card_y, card_w, card_h,
              fill_color=WHITE, line_color=RED, line_width=Pt(2.5))
    # Шапка
    add_shape(slide, MSO_SHAPE.RECTANGLE,
              cx1, card_y, card_w, Inches(0.80),
              fill_color=RED, line_color=None)
    add_textbox(slide, cx1, card_y + Inches(0.05), card_w, Inches(0.35),
                '🔥  ЭКСТРЕМАЛЬНЫЙ ТЕСТ', font_size=9, color=YELLOW, bold=True,
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_textbox(slide, cx1, card_y + Inches(0.35), card_w, Inches(0.45),
                'TITAN vs 365 дней\nпод открытым небом', font_size=13, color=WHITE, bold=True,
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # Описание
    desc1_y = card_y + Inches(0.95)
    items_1 = [
        ('Что:', '10 металлических пластин красим TITAN 3в1 и выставляем на крышу здания в Ростове-на-Дону', RED),
        ('Условия:', 'Солнце, дождь, снег, град, перепады от -25°C до +40°C — 365 дней без укрытия', DARK),
        ('Трансляция:', 'Онлайн-камера 24/7 на сайте titanpaints.ru — каждый может смотреть live', DARK),
        ('Контроль:', 'Ежемесячные фото-отчёты + независимая экспертиза ООО «Рутил»', DARK),
        ('Финал:', 'Через 365 дней — торжественное снятие, публикация результатов', DARK),
    ]
    for i, (label, text, color) in enumerate(items_1):
        cy = desc1_y + Inches(0.78) * i
        add_textbox(slide, cx1 + Inches(0.15), cy,
                    Inches(0.8), Inches(0.30),
                    label, font_size=9, color=color, bold=True,
                    align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP)
        add_textbox(slide, cx1 + Inches(0.95), cy,
                    card_w - Inches(1.10), Inches(0.70),
                    text, font_size=9, color=DARK,
                    align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP)

    # Цели
    goals_y = desc1_y + Inches(4.05)
    add_shape(slide, MSO_SHAPE.RECTANGLE,
              cx1 + Inches(0.15), goals_y,
              card_w - Inches(0.30), Inches(0.35),
              fill_color=RED, line_color=None)
    add_textbox(slide, cx1 + Inches(0.15), goals_y,
                card_w - Inches(0.30), Inches(0.35),
                'ЦЕЛЬ', font_size=9, color=YELLOW, bold=True,
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_textbox(slide, cx1 + Inches(0.15), goals_y + Inches(0.40),
                card_w - Inches(0.30), Inches(0.55),
                'Доказать «10 лет без перекраски» не на словах, а публично — в реальном времени',
                font_size=9, color=DARK, bold=True,
                align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP)

    # ===== КАРТОЧКА 2: ЯНДЕКС.ДИРЕКТ =====
    cx2 = cx1 + card_w + gap
    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
              cx2, card_y, card_w, card_h,
              fill_color=WHITE, line_color=DARK_BLUE, line_width=Pt(2.5))
    add_shape(slide, MSO_SHAPE.RECTANGLE,
              cx2, card_y, card_w, Inches(0.80),
              fill_color=DARK_BLUE, line_color=None)
    add_textbox(slide, cx2, card_y + Inches(0.05), card_w, Inches(0.35),
                '🔍  КОНТЕКСТНАЯ РЕКЛАМА', font_size=9, color=YELLOW, bold=True,
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_textbox(slide, cx2, card_y + Inches(0.35), card_w, Inches(0.45),
                'Яндекс.Директ\n(поиск + РСЯ)', font_size=13, color=WHITE, bold=True,
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # Описание
    desc2_y = card_y + Inches(0.95)
    items_2 = [
        ('Как:', 'Человек вводит «краска по ржавчине», «грунт-эмаль 3 в 1», «как покрасить забор» — видит рекламу TITAN', DARK_BLUE),
        ('Формат:', 'Текстово-графические объявления + баннеры РСЯ на сайтах партнёров', DARK),
        ('Таргетинг:', 'Мужчины 25-55 лет, интересы: ремонт, строительство, дача, авто', DARK),
        ('Посыл:', '«TITAN 3в1 — 10 лет без перекраски. Найдите в магазине [сеть]» + ссылка на тест 365', DARK),
        ('Бюджет:', '15-30 ₽ за клик · ~30 000 ₽/мес на регион · охват 50 000+ показов', DARK),
    ]
    for i, (label, text, color) in enumerate(items_2):
        cy = desc2_y + Inches(0.78) * i
        add_textbox(slide, cx2 + Inches(0.15), cy,
                    Inches(0.8), Inches(0.30),
                    label, font_size=9, color=color, bold=True,
                    align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP)
        add_textbox(slide, cx2 + Inches(0.95), cy,
                    card_w - Inches(1.10), Inches(0.70),
                    text, font_size=9, color=DARK,
                    align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP)

    # Цели
    add_shape(slide, MSO_SHAPE.RECTANGLE,
              cx2 + Inches(0.15), goals_y,
              card_w - Inches(0.30), Inches(0.35),
              fill_color=DARK_BLUE, line_color=None)
    add_textbox(slide, cx2 + Inches(0.15), goals_y,
                card_w - Inches(0.30), Inches(0.35),
                'ЦЕЛЬ', font_size=9, color=YELLOW, bold=True,
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_textbox(slide, cx2 + Inches(0.15), goals_y + Inches(0.40),
                card_w - Inches(0.30), Inches(0.55),
                'Ловить горячий спрос — тех, кто уже ищет ЛКМ, и направлять в точки продаж TITAN',
                font_size=9, color=DARK, bold=True,
                align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP)

    # ===== КАРТОЧКА 3: RUTUBE =====
    cx3 = cx2 + card_w + gap
    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
              cx3, card_y, card_w, card_h,
              fill_color=WHITE, line_color=GREEN, line_width=Pt(2.5))
    add_shape(slide, MSO_SHAPE.RECTANGLE,
              cx3, card_y, card_w, Inches(0.80),
              fill_color=GREEN, line_color=None)
    add_textbox(slide, cx3, card_y + Inches(0.05), card_w, Inches(0.35),
                '📺  ВИДЕОРЕКЛАМА', font_size=9, color=YELLOW, bold=True,
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_textbox(slide, cx3, card_y + Inches(0.35), card_w, Inches(0.45),
                'RuTube\n(без VPN в РФ)', font_size=13, color=WHITE, bold=True,
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # Описание
    desc3_y = card_y + Inches(0.95)
    items_3 = [
        ('Почему RuTube:', 'Единственная видеоплатформа, которая работает в РФ без VPN. Аудитория растёт', GREEN),
        ('Формат:', 'Пре-роллы 15-30 сек перед видео о ремонте, строительстве, даче, авто', DARK),
        ('Контент:', 'Демо TITAN vs ПФ-115 + кадры из экстремального теста 365 дней', DARK),
        ('Таргетинг:', 'Мужчины 25-55, каналы: «Ремонт», «Строительство», «Дача», «Гараж», «Сделай сам»', DARK),
        ('Бюджет:', '~10-15 ₽ за просмотр · ~40 000 ₽/мес · охват 100 000+ просмотров', DARK),
    ]
    for i, (label, text, color) in enumerate(items_3):
        cy = desc3_y + Inches(0.78) * i
        add_textbox(slide, cx3 + Inches(0.15), cy,
                    Inches(1.1), Inches(0.30),
                    label, font_size=9, color=color, bold=True,
                    align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP)
        add_textbox(slide, cx3 + Inches(1.25), cy,
                    card_w - Inches(1.40), Inches(0.70),
                    text, font_size=9, color=DARK,
                    align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP)

    # Цели
    add_shape(slide, MSO_SHAPE.RECTANGLE,
              cx3 + Inches(0.15), goals_y,
              card_w - Inches(0.30), Inches(0.35),
              fill_color=GREEN, line_color=None)
    add_textbox(slide, cx3 + Inches(0.15), goals_y,
                card_w - Inches(0.30), Inches(0.35),
                'ЦЕЛЬ', font_size=9, color=YELLOW, bold=True,
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_textbox(slide, cx3 + Inches(0.15), goals_y + Inches(0.40),
                card_w - Inches(0.30), Inches(0.55),
                'Массовый охват целевой аудитории + видеоконтент из теста 365 = синергия',
                font_size=9, color=DARK, bold=True,
                align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP)

    # === НИЖНЯЯ ПОЛОСА — синергия ===
    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
              Inches(0.4), Inches(6.90), SLIDE_W - Inches(0.8), Inches(0.55),
              fill_color=YELLOW, line_color=None)
    add_textbox(slide, Inches(0.4), Inches(6.90), SLIDE_W - Inches(0.8), Inches(0.55),
                '⚡  Синергия: экстремальный тест 365 → контент для RuTube + Я.Директ → трафик в точки продаж → рост третичных продаж',
                font_size=12, color=DARK_RED, bold=True,
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # Футер
    add_textbox(slide, Inches(0.4), Inches(7.30),
                SLIDE_W - Inches(0.8), Inches(0.20),
                'Empils  •  TITAN  •  Outdoor-продвижение  •  Красишь 1 раз — служит 10 лет',
                font_size=8, color=GRAY, bold=False,
                align=PP_ALIGN.CENTER, italic=True)

    prs.save(OUTPUT)
    print(f'OK: {OUTPUT}')
    print(f'Слайдов: {len(prs.slides)}')


if __name__ == '__main__':
    main()
