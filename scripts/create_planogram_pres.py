#!/usr/bin/env python3
"""
Презентация: Планограммы TITAN.
2 варианта:
1. Весь модельный ряд (по категориям/сфере назначения)
2. По классам ТТ (A, B, C, D) — согласно MML
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
import os

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
SHELF_BROWN = RGBColor(0x8B, 0x6F, 0x47)

# RAL colors
RAL_COLORS = {
    'RAL 9005': RGBColor(0x1A, 0x1A, 0x1A),
    'RAL 9010': RGBColor(0xF5, 0xF5, 0xDC),
    'RAL 7004': RGBColor(0xB0, 0xB0, 0xB0),
    'RAL 7024': RGBColor(0x40, 0x40, 0x40),
    'RAL 8017': RGBColor(0x3D, 0x28, 0x17),
    'RAL 6005': RGBColor(0x2D, 0x50, 0x16),
    'RAL 1021': RGBColor(0xE6, 0xB8, 0x00),
    'RAL 5005': RGBColor(0x00, 0x47, 0xAB),
    'RAL 3009': RGBColor(0x7B, 0x3F, 0x00),
    'RAL 1014': RGBColor(0xE8, 0xD4, 0xA0),
    'RAL 9006': RGBColor(0xA0, 0xA0, 0xA0),
    'RAL 1023': RGBColor(0xE6, 0xC8, 0x00),
    'Чёрный': RGBColor(0x1A, 0x1A, 0x1A),
    'Коричневый': RGBColor(0x4A, 0x2C, 0x17),
    'Серебристый': RGBColor(0xC0, 0xC0, 0xC0),
    'Золотистый': RGBColor(0xB8, 0x86, 0x0B),
    'Медный': RGBColor(0xB8, 0x73, 0x33),
    'Тёмно-зелёный': RGBColor(0x35, 0x5E, 0x3B),
    'Серый': RGBColor(0x80, 0x80, 0x80),
}

OUTPUT = '/home/z/my-project/download/Планограммы_TITAN_презентация.pptx'

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


def header_band(slide, title, subtitle=None, color=RED):
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, Inches(0.65),
              fill_color=color, line_color=None)
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0, Inches(0.65), SLIDE_W, Inches(0.04),
              fill_color=YELLOW, line_color=None)
    add_textbox(slide, Inches(0.4), Inches(0.08), SLIDE_W - Inches(0.8), Inches(0.5),
                title, font_size=22, color=WHITE, bold=True,
                align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)
    if subtitle:
        add_textbox(slide, SLIDE_W - Inches(5.0), Inches(0.08),
                    Inches(4.6), Inches(0.5),
                    subtitle, font_size=11, color=YELLOW, bold=False,
                    align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)


def footer(slide):
    add_textbox(slide, Inches(0.4), Inches(7.05),
                SLIDE_W - Inches(0.8), Inches(0.35),
                'Empils  •  TITAN  •  Планограммы выкладки  •  Красишь 1 раз — служит 10 лет',
                font_size=8, color=GRAY, bold=False,
                align=PP_ALIGN.CENTER, italic=True)


def draw_can(slide, x, y, w, h, color_rgb, label_top, label_mid, label_bot,
             effect='gloss'):
    r, g, b = color_rgb[0], color_rgb[1], color_rgb[2]
    brightness = 0.299 * r + 0.587 * g + 0.114 * b
    text_color = DARK if brightness > 140 else WHITE

    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h,
              fill_color=color_rgb, line_color=DARK, line_width=Pt(0.5))
    lid_h = Inches(0.06)
    add_shape(slide, MSO_SHAPE.RECTANGLE,
              x + Inches(0.03), y - lid_h + Inches(0.02),
              w - Inches(0.06), lid_h,
              fill_color=GRAY, line_color=DARK, line_width=Pt(0.3))

    if effect == 'hammer':
        import random
        random.seed(int(x) + int(y))
        w_in = w / 914400 if isinstance(w, (int, float)) else w.inches
        h_in = h / 914400 if isinstance(h, (int, float)) else h.inches
        for _ in range(4):
            dx = random.uniform(0.15, w_in - 0.15)
            dy = random.uniform(0.20, h_in - 0.20)
            dot_r = Inches(0.03)
            add_shape(slide, MSO_SHAPE.OVAL,
                      x + Inches(dx) - dot_r, y + Inches(dy) - dot_r,
                      dot_r * 2, dot_r * 2,
                      fill_color=WHITE, line_color=None)

    tf = slide.shapes.add_textbox(x, y + Inches(0.02), w, h - Inches(0.04)).text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = Pt(0); tf.margin_right = Pt(0)
    tf.margin_top = Pt(0); tf.margin_bottom = Pt(0)
    for i, (text, size, bold) in enumerate([
        (label_top, 6, True),
        (label_mid, 5, True),
        (label_bot, 5, False),
    ]):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = text
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.color.rgb = text_color
        r.font.name = 'Arial'


def draw_shelf(slide, shelf_y, shelf_w, shelf_h, cans, shelf_title,
               accent_color, label=''):
    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
              Inches(0.4), shelf_y, shelf_w, shelf_h,
              fill_color=LIGHT_GRAY, line_color=GRAY, line_width=Pt(1))
    shelf_board_h = Inches(0.20)
    add_shape(slide, MSO_SHAPE.RECTANGLE,
              Inches(0.55), shelf_y + shelf_h - shelf_board_h - Inches(0.03),
              shelf_w - Inches(0.30), shelf_board_h,
              fill_color=SHELF_BROWN, line_color=RGBColor(0x6B, 0x56, 0x35), line_width=Pt(1))

    add_textbox(slide, Inches(0.55), shelf_y + Inches(0.05),
                Inches(6.0), Inches(0.30),
                shelf_title, font_size=10, color=accent_color, bold=True,
                align=PP_ALIGN.LEFT)

    if label:
        add_textbox(slide, shelf_w + Inches(0.2) - Inches(2.0), shelf_y + Inches(0.05),
                    Inches(1.8), Inches(0.30),
                    label, font_size=9, color=accent_color, bold=True,
                    align=PP_ALIGN.RIGHT)

    n = len(cans)
    if n == 0:
        add_textbox(slide, Inches(0.55), shelf_y + Inches(0.40),
                    shelf_w - Inches(0.30), Inches(0.40),
                    'TITAN не представлен', font_size=10, color=RED, bold=True,
                    align=PP_ALIGN.CENTER)
        return

    margin_x = Inches(0.30)
    gap = Inches(0.10)
    can_w = (shelf_w - margin_x * 2 - gap * (n - 1)) / n
    can_h = shelf_h - shelf_board_h - Inches(0.45)

    can_y = shelf_y + shelf_h - shelf_board_h - can_h - Inches(0.10)

    for i, can in enumerate(cans):
        cx = Inches(0.4) + margin_x + (can_w + gap) * i
        color = RAL_COLORS.get(can.get('ral', ''), DARK)
        draw_can(slide, cx, can_y, can_w, can_h, color,
                 can.get('label_top', 'TITAN'),
                 can.get('label_mid', ''),
                 can.get('label_bot', ''),
                 effect=can.get('effect', 'gloss'))


# === ДАННЫЕ SKU ===
# 3в1 по ржавчине — все SKU
sku_3v1 = [
    {'ral': 'RAL 9005', 'label_mid': '3в1 0.75л', 'label_bot': 'Чёрный'},
    {'ral': 'RAL 8017', 'label_mid': '3в1 0.75л', 'label_bot': 'Шоколад'},
    {'ral': 'RAL 7004', 'label_mid': '3в1 0.75л', 'label_bot': 'Серый'},
    {'ral': 'RAL 7024', 'label_mid': '3в1 0.75л', 'label_bot': 'Графит'},
    {'ral': 'RAL 9010', 'label_mid': '3в1 0.75л', 'label_bot': 'Белый'},
    {'ral': 'RAL 6005', 'label_mid': '3в1 0.75л', 'label_bot': 'Зелёный'},
    {'ral': 'RAL 1021', 'label_mid': '3в1 0.75л', 'label_bot': 'Жёлтый'},
    {'ral': 'RAL 5005', 'label_mid': '3в1 0.75л', 'label_bot': 'Синий'},
    {'ral': 'RAL 3009', 'label_mid': '3в1 0.75л', 'label_bot': 'К-кор'},
    {'ral': 'RAL 9005', 'label_mid': '3в1 2л', 'label_bot': 'Чёрный'},
    {'ral': 'RAL 8017', 'label_mid': '3в1 2л', 'label_bot': 'Шоколад'},
    {'ral': 'RAL 7004', 'label_mid': '3в1 2л', 'label_bot': 'Серый'},
    {'ral': 'RAL 7024', 'label_mid': '3в1 2л', 'label_bot': 'Графит'},
    {'ral': 'RAL 9010', 'label_mid': '3в1 2л', 'label_bot': 'Белый'},
]

# Быстросохнущая
sku_fast = [
    {'ral': 'RAL 9005', 'label_mid': 'БЫСТ 0.75л', 'label_bot': 'Чёрный', 'effect': 'matte'},
    {'ral': 'Серый', 'label_mid': 'БЫСТ 0.75л', 'label_bot': 'Серый', 'effect': 'matte'},
    {'ral': 'RAL 8017', 'label_mid': 'БЫСТ 0.75л', 'label_bot': 'Шоколад', 'effect': 'matte'},
    {'ral': 'RAL 9010', 'label_mid': 'БЫСТ 0.75л', 'label_bot': 'Белый', 'effect': 'matte'},
    {'ral': 'RAL 5005', 'label_mid': 'БЫСТ 0.75л', 'label_bot': 'Синий', 'effect': 'matte'},
    {'ral': 'RAL 1023', 'label_mid': 'БЫСТ 0.75л', 'label_bot': 'Жёлтый', 'effect': 'matte'},
    {'ral': 'RAL 9005', 'label_mid': 'БЫСТ 2.2л', 'label_bot': 'Чёрный', 'effect': 'matte'},
]

# Молотковая
sku_hammer = [
    {'ral': 'Чёрный', 'label_mid': 'МОЛ 0.8кг', 'label_bot': 'Чёрный', 'effect': 'hammer'},
    {'ral': 'Коричневый', 'label_mid': 'МОЛ 0.8кг', 'label_bot': 'Корич', 'effect': 'hammer'},
    {'ral': 'Серебристый', 'label_mid': 'МОЛ 0.8кг', 'label_bot': 'Сереб', 'effect': 'hammer'},
    {'ral': 'Золотистый', 'label_mid': 'МОЛ 0.8кг', 'label_bot': 'Золот', 'effect': 'hammer'},
    {'ral': 'Медный', 'label_mid': 'МОЛ 0.8кг', 'label_bot': 'Медн', 'effect': 'hammer'},
    {'ral': 'Тёмно-зелёный', 'label_mid': 'МОЛ 0.8кг', 'label_bot': 'Т-зел', 'effect': 'hammer'},
    {'ral': 'Чёрный', 'label_mid': 'МОЛ 2.5кг', 'label_bot': 'Чёрный', 'effect': 'hammer'},
]

# MML по классам (из файла)
mml_a = [
    {'ral': 'RAL 9005', 'label_mid': '3в1 0.75', 'label_bot': 'Чёрный'},
    {'ral': 'RAL 9010', 'label_mid': '3в1 0.75', 'label_bot': 'Белый'},
    {'ral': 'RAL 7004', 'label_mid': '3в1 0.75', 'label_bot': 'Серый'},
    {'ral': 'RAL 7024', 'label_mid': '3в1 0.75', 'label_bot': 'Графит'},
    {'ral': 'RAL 8017', 'label_mid': '3в1 0.75', 'label_bot': 'Шоколад'},
    {'ral': 'RAL 9005', 'label_mid': '3в1 2л', 'label_bot': 'Чёрный'},
    {'ral': 'RAL 9010', 'label_mid': '3в1 2л', 'label_bot': 'Белый'},
    {'ral': 'RAL 7004', 'label_mid': '3в1 2л', 'label_bot': 'Серый'},
    {'ral': 'RAL 7024', 'label_mid': '3в1 2л', 'label_bot': 'Графит'},
    {'ral': 'RAL 8017', 'label_mid': '3в1 2л', 'label_bot': 'Шоколад'},
    {'ral': 'RAL 9005', 'label_mid': 'БЫСТ 0.75', 'label_bot': 'Чёрный', 'effect': 'matte'},
    {'ral': 'RAL 9005', 'label_mid': 'БЫСТ 2.2л', 'label_bot': 'Чёрный', 'effect': 'matte'},
    {'ral': 'Чёрный', 'label_mid': 'МОЛ 0.8', 'label_bot': 'Чёрный', 'effect': 'hammer'},
    {'ral': 'Чёрный', 'label_mid': 'МОЛ 2.5', 'label_bot': 'Чёрный', 'effect': 'hammer'},
    {'ral': 'Коричневый', 'label_mid': 'МОЛ 0.8', 'label_bot': 'Корич', 'effect': 'hammer'},
    {'ral': 'Серебристый', 'label_mid': 'МОЛ 0.8', 'label_bot': 'Сереб', 'effect': 'hammer'},
]

mml_b = [
    {'ral': 'RAL 9005', 'label_mid': '3в1 0.75', 'label_bot': 'Чёрный'},
    {'ral': 'RAL 9010', 'label_mid': '3в1 0.75', 'label_bot': 'Белый'},
    {'ral': 'RAL 7004', 'label_mid': '3в1 0.75', 'label_bot': 'Серый'},
    {'ral': 'RAL 7024', 'label_mid': '3в1 0.75', 'label_bot': 'Графит'},
    {'ral': 'RAL 8017', 'label_mid': '3в1 0.75', 'label_bot': 'Шоколад'},
    {'ral': 'RAL 9005', 'label_mid': '3в1 2л', 'label_bot': 'Чёрный'},
    {'ral': 'RAL 7004', 'label_mid': '3в1 2л', 'label_bot': 'Серый'},
    {'ral': 'RAL 7024', 'label_mid': '3в1 2л', 'label_bot': 'Графит'},
    {'ral': 'RAL 8017', 'label_mid': '3в1 2л', 'label_bot': 'Шоколад'},
    {'ral': 'RAL 9005', 'label_mid': 'БЫСТ 0.75', 'label_bot': 'Чёрный', 'effect': 'matte'},
    {'ral': 'Чёрный', 'label_mid': 'МОЛ 0.8', 'label_bot': 'Чёрный', 'effect': 'hammer'},
    {'ral': 'Коричневый', 'label_mid': 'МОЛ 0.8', 'label_bot': 'Корич', 'effect': 'hammer'},
]

mml_c = [
    {'ral': 'Чёрный', 'label_mid': 'МОЛ 0.8', 'label_bot': 'Чёрный', 'effect': 'hammer'},
    {'ral': 'Коричневый', 'label_mid': 'МОЛ 0.8', 'label_bot': 'Корич', 'effect': 'hammer'},
]

mml_d = []


# ==============================================================
# СЛАЙДЫ
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
                'TITAN', font_size=72, color=WHITE, bold=True, align=PP_ALIGN.LEFT)
    add_textbox(slide, Inches(0.4), Inches(2.3), Inches(5.0), Inches(0.7),
                'ПЛАНОГРАММЫ', font_size=32, color=WHITE, bold=True, align=PP_ALIGN.LEFT)
    add_textbox(slide, Inches(0.4), Inches(3.1), Inches(5.0), Inches(0.7),
                'ВЫКЛАДКИ', font_size=32, color=WHITE, bold=True, align=PP_ALIGN.LEFT)
    add_shape(slide, MSO_SHAPE.RECTANGLE, Inches(0.4), Inches(4.0),
              Inches(2.5), Inches(0.05), fill_color=YELLOW, line_color=None)
    add_textbox(slide, Inches(0.4), Inches(4.2), Inches(5.0), Inches(0.5),
                '2 варианта выкладки', font_size=18, color=WHITE, bold=False)
    add_textbox(slide, Inches(0.4), Inches(4.75), Inches(5.0), Inches(0.5),
                '3 линейки · 27 SKU · 4 класса ТТ', font_size=16, color=WHITE, italic=True)

    add_textbox(slide, Inches(6.0), Inches(1.0), Inches(7.0), Inches(1.0),
                'Красишь 1 раз', font_size=36, color=GREEN, bold=True, align=PP_ALIGN.LEFT)
    add_textbox(slide, Inches(6.0), Inches(1.9), Inches(7.0), Inches(1.0),
                'служит 10 лет', font_size=36, color=GREEN, bold=True, align=PP_ALIGN.LEFT)
    add_shape(slide, MSO_SHAPE.RECTANGLE, Inches(6.0), Inches(2.95),
              Inches(2.0), Inches(0.05), fill_color=YELLOW, line_color=None)

    blocks = [
        ('1', 'ВЕСЬ МОДЕЛЬНЫЙ РЯД', 'По категориям/сфере назначения', RED),
        ('2', 'ПО КЛАССАМ ТТ (MML)', 'A · B · C · D — по наличию ассортимента', DARK_BLUE),
    ]
    block_y = Inches(3.5)
    for i, (num, name, sub, color) in enumerate(blocks):
        cx = Inches(6.0)
        cy = block_y + Inches(1.8) * i
        add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
                  cx, cy, Inches(7.0), Inches(1.50),
                  fill_color=WHITE, line_color=color, line_width=Pt(2.5))
        add_shape(slide, MSO_SHAPE.RECTANGLE,
                  cx, cy, Inches(0.15), Inches(1.50),
                  fill_color=color, line_color=None)
        add_textbox(slide, cx + Inches(0.4), cy + Inches(0.15),
                    Inches(0.6), Inches(0.5),
                    num, font_size=28, color=color, bold=True)
        add_textbox(slide, cx + Inches(1.1), cy + Inches(0.15),
                    Inches(5.5), Inches(0.5),
                    name, font_size=16, color=color, bold=True)
        add_textbox(slide, cx + Inches(1.1), cy + Inches(0.70),
                    Inches(5.5), Inches(0.5),
                    sub, font_size=12, color=GRAY)

    add_textbox(slide, Inches(6.0), Inches(6.8), Inches(7.0), Inches(0.5),
                'Empils  •  100+ лет на рынке ЛКМ  •  ТОП-3 производителей России',
                font_size=11, color=GRAY, italic=True)


def slide_logic(prs):
    """Слайд 2: Логика выкладки — принципы распределения цветов."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H,
              fill_color=WHITE, line_color=None)

    header_band(slide, 'ЛОГИКА ВЫКЛАДКИ — ПРИНЦИПЫ',
                'Как распределены цвета на полке', color=DARK_BLUE)

    # 4 принципа в 2×2 сетке
    principles = [
        ('1', 'От тёмных к светлым',
         'Цвета слева направо: чёрный → коричневый → серый → графит → белый. '
         'Покупатель сканирует полку слева — тёмные (топ-продажи) первыми.',
         RED),
        ('2', 'Контраст соседей',
         'Схожие оттенки (серый RAL 7004 и графит RAL 7024) не ставим рядом — '
         'они сливаются. Между ними — коричневый для визуального разделения.',
         ORANGE),
        ('3', 'ABC по продажам внутри градации',
         'ТОП-3 цвета (чёрный 7738 шт, графит 5172 шт, серый 5071 шт) — '
         'в начале полки. Дублируем фейсинги на топ-3.',
         GREEN),
        ('4', 'Цветные — отдельно',
         'Зелёный, жёлтый, синий — после нейтральных. Это специфические цвета '
         '(дача, дорожная техника) с меньшим оборотом.',
         DARK_BLUE),
    ]

    card_w = Inches(6.10)
    card_h = Inches(2.65)
    gap = Inches(0.20)
    start_y = Inches(0.95)

    for i, (num, title, desc, color) in enumerate(principles):
        row = i // 2
        col = i % 2
        cx = Inches(0.4) + (card_w + gap) * col
        cy = start_y + (card_h + gap) * row

        add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
                  cx, cy, card_w, card_h,
                  fill_color=WHITE, line_color=color, line_width=Pt(2))
        # Шапка
        add_shape(slide, MSO_SHAPE.RECTANGLE,
                  cx, cy, card_w, Inches(0.55),
                  fill_color=color, line_color=None)
        # Номер
        add_shape(slide, MSO_SHAPE.OVAL,
                  cx + Inches(0.15), cy + Inches(0.10),
                  Inches(0.35), Inches(0.35),
                  fill_color=WHITE, line_color=None)
        add_textbox(slide, cx + Inches(0.15), cy + Inches(0.10),
                    Inches(0.35), Inches(0.35),
                    num, font_size=14, color=color, bold=True,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        # Заголовок
        add_textbox(slide, cx + Inches(0.60), cy + Inches(0.05),
                    card_w - Inches(0.70), Inches(0.45),
                    title, font_size=14, color=WHITE, bold=True,
                    align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)
        # Описание
        add_textbox(slide, cx + Inches(0.20), cy + Inches(0.70),
                    card_w - Inches(0.40), card_h - Inches(0.85),
                    desc, font_size=11, color=DARK,
                    align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP)

    # === Визуальный пример — полоска с цветами ===
    example_y = start_y + (card_h + gap) * 2 + Inches(0.10)
    add_textbox(slide, Inches(0.4), example_y,
                Inches(12.5), Inches(0.35),
                'ПРИМЕР: выкладка 3в1 0.75 л — 5 цветов RAL',
                font_size=12, color=DARK_BLUE, bold=True, align=PP_ALIGN.LEFT)

    # Цветная полоска
    bar_y = example_y + Inches(0.40)
    bar_h = Inches(0.70)
    bar_w = Inches(2.30)
    bar_gap = Inches(0.10)

    colors_order = [
        ('Чёрный', 'RAL 9005', RGBColor(0x1A, 0x1A, 0x1A), '7 738 шт'),
        ('Шоколад', 'RAL 8017', RGBColor(0x3D, 0x28, 0x17), '4 524 шт'),
        ('Серый', 'RAL 7004', RGBColor(0xB0, 0xB0, 0xB0), '5 071 шт'),
        ('Графит', 'RAL 7024', RGBColor(0x40, 0x40, 0x40), '5 172 шт'),
        ('Белый', 'RAL 9010', RGBColor(0xF5, 0xF5, 0xDC), '3 273 шт'),
    ]

    for i, (name, ral, color, qty) in enumerate(colors_order):
        cx = Inches(0.4) + (bar_w + bar_gap) * i
        # Банка-полоска
        add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
                  cx, bar_y, bar_w, bar_h,
                  fill_color=color, line_color=DARK, line_width=Pt(1))
        # Текст на банке
        brightness = 0.299 * color[0] + 0.587 * color[1] + 0.114 * color[2]
        tc = DARK if brightness > 140 else WHITE
        add_multi_text(slide, cx, bar_y, bar_w, bar_h,
                       [
                           (name, 10, tc, True),
                           (ral, 8, tc, False),
                           (qty, 8, tc, False),
                       ])
        # Стрелка между банками (кроме последней)
        if i < len(colors_order) - 1:
            add_shape(slide, MSO_SHAPE.RIGHT_ARROW,
                      cx + bar_w - Inches(0.02), bar_y + bar_h // 2 - Inches(0.10),
                      Inches(0.14), Inches(0.20),
                      fill_color=GRAY, line_color=None)

    # Подпись под полоской
    add_textbox(slide, Inches(0.4), bar_y + bar_h + Inches(0.10),
                Inches(12.5), Inches(0.30),
                'тёмные ←                                                                                    → светлые',
                font_size=9, color=GRAY, italic=True, align=PP_ALIGN.LEFT)

    footer(slide)


def slide_full_range(prs):
    """Слайд 2: Весь модельный ряд — 3 категории на 3 полках."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H,
              fill_color=WHITE, line_color=None)

    header_band(slide, 'ВАРИАНТ 1 — ВЕСЬ МОДЕЛЬНЫЙ РЯД',
                '3 категории · 27 SKU · выкладка по сфере назначения', color=RED)

    # 3 полки
    shelf_w = SLIDE_W - Inches(0.8)
    shelf_y1 = Inches(0.90)
    shelf_y2 = Inches(3.05)
    shelf_y3 = Inches(5.20)
    shelf_h = Inches(1.95)

    # Полка 1: 3в1 по ржавчине
    draw_shelf(slide, shelf_y1, shelf_w, shelf_h, sku_3v1,
               'ПОЛКА 1: Грунт-эмаль 3в1 по ржавчине (9 цветов × 2 объёма = 14 SKU)',
               RED, '14 SKU')

    # Полка 2: Быстросохнущая
    draw_shelf(slide, shelf_y2, shelf_w, shelf_h, sku_fast,
               'ПОЛКА 2: Быстросохнущая матовая (6 цветов 0.75л + 1×2.2л = 7 SKU)',
               DARK, '7 SKU')

    # Полка 3: Молотковая
    draw_shelf(slide, shelf_y3, shelf_w, shelf_h, sku_hammer,
               'ПОЛКА 3: Эмаль с молотковым эффектом (6 цветов 0.8кг + 1×2.5кг = 7 SKU)',
               GOLD, '7 SKU')

    # Подсказка
    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
              Inches(0.4), Inches(7.10), SLIDE_W - Inches(0.8), Inches(0.35),
              fill_color=YELLOW, line_color=None)
    add_textbox(slide, Inches(0.4), Inches(7.10), SLIDE_W - Inches(0.8), Inches(0.35),
                '!  Для точек класса A (федеральная DIY) — выкладка всего модельного ряда. Топ-3 цвета дублировать фейсингами.',
                font_size=10, color=DARK, bold=True,
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)


def slide_class_a(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H,
              fill_color=WHITE, line_color=None)

    header_band(slide, 'ВАРИАНТ 2 — КЛАСС A: ФЕДЕРАЛЬНАЯ СЕТЬ DIY',
                '16 SKU · MML', color=RED)

    # 3 полки
    shelf_w = SLIDE_W - Inches(0.8)
    shelf_h = Inches(1.90)

    # Разделим 16 SKU на 3 полки
    shelf1_cans = mml_a[:6]   # 3в1 0.75л (5 цветов + 1 доп)
    shelf2_cans = mml_a[6:11]  # 3в1 2л (4) + быстросохнущая (2)
    shelf3_cans = mml_a[11:]   # молотковая (4)

    # Полка 1: 3в1 0.75 л
    draw_shelf(slide, Inches(0.90), shelf_w, shelf_h, shelf1_cans,
               'ПОЛКА 1: 3в1 по ржавчине 0.75 л — 5 цветов RAL', RED, '6 SKU')

    # Полка 2: 3в1 2л + быстросохнущая
    draw_shelf(slide, Inches(2.95), shelf_w, shelf_h, shelf2_cans,
               'ПОЛКА 2: 3в1 2 л (4 цвета) + быстросохнущая (2 SKU)', DARK, '5 SKU')

    # Полка 3: молотковая
    draw_shelf(slide, Inches(5.00), shelf_w, shelf_h, shelf3_cans,
               'ПОЛКА 3: молотковый эффект — 4 цвета', GOLD, '5 SKU')

    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
              Inches(0.4), Inches(7.10), SLIDE_W - Inches(0.8), Inches(0.35),
              fill_color=YELLOW, line_color=None)
    add_textbox(slide, Inches(0.4), Inches(7.10), SLIDE_W - Inches(0.8), Inches(0.35),
                '!  Класс A — эталонная выкладка. Топ-3 цвета (чёрный, коричневый, графит) — по 2 фейсинга. Глубина: 6 шт/SKU.',
                font_size=10, color=DARK, bold=True,
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)


def slide_class_b(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H,
              fill_color=WHITE, line_color=None)

    header_band(slide, 'ВАРИАНТ 2 — КЛАСС B: ЛОКАЛЬНАЯ СЕТЬ',
                '12 SKU · MML', color=ORANGE)

    shelf_w = SLIDE_W - Inches(0.8)
    shelf_h = Inches(1.90)

    shelf1_cans = mml_b[:5]   # 3в1 0.75л (5 цветов)
    shelf2_cans = mml_b[5:10]  # 3в1 2л (4) + быстросохнущая (1)
    shelf3_cans = mml_b[10:]   # молотковая (2)

    draw_shelf(slide, Inches(0.90), shelf_w, shelf_h, shelf1_cans,
               'ПОЛКА 1: 3в1 по ржавчине 0.75 л — 5 цветов RAL', ORANGE, '5 SKU')

    draw_shelf(slide, Inches(2.95), shelf_w, shelf_h, shelf2_cans,
               'ПОЛКА 2: 3в1 2 л (4 цвета) + быстросохнущая (1 SKU)', DARK, '5 SKU')

    draw_shelf(slide, Inches(5.00), shelf_w, shelf_h, shelf3_cans,
               'ПОЛКА 3: молотковый эффект — 2 цвета', GOLD, '2 SKU')

    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
              Inches(0.4), Inches(7.10), SLIDE_W - Inches(0.8), Inches(0.35),
              fill_color=YELLOW, line_color=None)
    add_textbox(slide, Inches(0.4), Inches(7.10), SLIDE_W - Inches(0.8), Inches(0.35),
                '!  Класс B — основная выкладка. Нет белого 2л, серебристой молотковой, 2.5 кг. Глубина: 4 шт/SKU.',
                font_size=10, color=DARK, bold=True,
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)


def slide_class_cd(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H,
              fill_color=WHITE, line_color=None)

    header_band(slide, 'ВАРИАНТ 2 — КЛАСС C и D: СУПЕРМАРКЕТ И ХОЗМАГ',
                'C: 2 SKU · D: 0 SKU · MML (текущий)', color=GREEN)

    shelf_w = (SLIDE_W - Inches(1.2)) / 2

    # === ЛЕВАЯ ПОЛОВИНА: КЛАСС C ===
    add_textbox(slide, Inches(0.4), Inches(0.85),
                shelf_w, Inches(0.35),
                'КЛАСС C — СУПЕРМАРКЕТ (2 SKU)', font_size=14, color=GREEN, bold=True,
                align=PP_ALIGN.CENTER)

    # Полка C
    c_shelf_y = Inches(1.30)
    c_shelf_h = Inches(2.50)
    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
              Inches(0.4), c_shelf_y, shelf_w, c_shelf_h,
              fill_color=LIGHT_GRAY, line_color=GREEN, line_width=Pt(2))
    add_shape(slide, MSO_SHAPE.RECTANGLE,
              Inches(0.55), c_shelf_y + c_shelf_h - Inches(0.20) - Inches(0.03),
              shelf_w - Inches(0.30), Inches(0.20),
              fill_color=SHELF_BROWN, line_color=RGBColor(0x6B, 0x56, 0x35), line_width=Pt(1))

    # 2 крупные банки
    for i, can in enumerate(mml_c):
        can_w = Inches(1.80)
        can_h = Inches(1.80)
        gap = Inches(0.50)
        total_w = can_w * 2 + gap
        start_x = Inches(0.4) + (shelf_w - total_w) // 2
        cx = start_x + (can_w + gap) * i
        cy = c_shelf_y + c_shelf_h - Inches(0.20) - can_h - Inches(0.15)
        color = RAL_COLORS.get(can['ral'], DARK)
        draw_can(slide, cx, cy, can_w, can_h, color,
                 'TITAN', can['label_mid'], can['label_bot'],
                 effect=can.get('effect', 'gloss'))

    # Подсказка C
    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
              Inches(0.4), Inches(4.00), shelf_w, Inches(0.55),
              fill_color=LIGHT_RED, line_color=RED, line_width=Pt(1))
    add_textbox(slide, Inches(0.4), Inches(4.00), shelf_w, Inches(0.55),
                '!  Только молотковая. TITAN 3в1 отсутствует.\nРекомендация: добавить 3в1 чёрный + коричневый 0.75л',
                font_size=9, color=DARK, bold=True,
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # === ПРАВАЯ ПОЛОВИНА: КЛАСС D ===
    d_x = Inches(0.4) + shelf_w + Inches(0.4)
    add_textbox(slide, d_x, Inches(0.85),
                shelf_w, Inches(0.35),
                'КЛАСС D — ХОЗМАГ (0 SKU)', font_size=14, color=DARK_BLUE, bold=True,
                align=PP_ALIGN.CENTER)

    # Полка D — пустая
    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
              d_x, c_shelf_y, shelf_w, c_shelf_h,
              fill_color=LIGHT_RED, line_color=RED, line_width=Pt(2))
    add_shape(slide, MSO_SHAPE.RECTANGLE,
              d_x + Inches(0.15), c_shelf_y + c_shelf_h - Inches(0.20) - Inches(0.03),
              shelf_w - Inches(0.30), Inches(0.20),
              fill_color=SHELF_BROWN, line_color=RGBColor(0x6B, 0x56, 0x35), line_width=Pt(1))

    # Сообщение «ПУСТО»
    add_textbox(slide, d_x, c_shelf_y + Inches(0.50),
                shelf_w, Inches(0.8),
                'TITAN', font_size=40, color=RED, bold=True,
                align=PP_ALIGN.CENTER)
    add_textbox(slide, d_x, c_shelf_y + Inches(1.30),
                shelf_w, Inches(0.5),
                'не представлен', font_size=18, color=RED, bold=True,
                align=PP_ALIGN.CENTER)
    add_textbox(slide, d_x, c_shelf_y + Inches(1.85),
                shelf_w, Inches(0.4),
                'Покупатель уходит с ПФ-115', font_size=11, color=GRAY, italic=True,
                align=PP_ALIGN.CENTER)

    # Подсказка D
    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
              d_x, Inches(4.00), shelf_w, Inches(0.55),
              fill_color=LIGHT_RED, line_color=RED, line_width=Pt(1))
    add_textbox(slide, d_x, Inches(4.00), shelf_w, Inches(0.55),
                '!  TITAN полностью отсутствует.\nРекомендация: добавить 3в1 чёрный + коричневый 0.75л',
                font_size=9, color=DARK, bold=True,
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # === НИЖНЯЯ ЧАСТЬ — рекомендация ===
    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
              Inches(0.4), Inches(4.80), SLIDE_W - Inches(0.8), Inches(2.10),
              fill_color=LIGHT_GREEN, line_color=GREEN, line_width=Pt(2))

    add_textbox(slide, Inches(0.6), Inches(4.85),
                SLIDE_W - Inches(1.2), Inches(0.40),
                '🚀  РЕКОМЕНДАЦИЯ: РАСШИРЕНИЕ MML ДЛЯ КЛАССОВ C и D',
                font_size=14, color=GREEN, bold=True,
                align=PP_ALIGN.CENTER)

    rec_items = [
        ('C: +3 SKU', '3в1 чёрный 0.75л + 3в1 коричневый 0.75л + 3в1 серый 0.75л', 'Топ-3 цвета по продажам (7738 + 4524 + 5071 шт)'),
        ('D: +2 SKU', '3в1 чёрный 0.75л + 3в1 коричневый 0.75л', 'Самые ходовые цвета — есть спрос даже в хозмаге'),
    ]
    for i, (title, desc, data) in enumerate(rec_items):
        cy = Inches(5.35) + Inches(0.65) * i
        add_shape(slide, MSO_SHAPE.RECTANGLE,
                  Inches(0.7), cy, Inches(2.0), Inches(0.50),
                  fill_color=GREEN, line_color=None)
        add_textbox(slide, Inches(0.7), cy, Inches(2.0), Inches(0.50),
                    title, font_size=12, color=WHITE, bold=True,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_textbox(slide, Inches(2.9), cy, Inches(4.5), Inches(0.50),
                    desc, font_size=10, color=DARK, bold=True,
                    align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)
        add_textbox(slide, Inches(7.5), cy, Inches(5.0), Inches(0.50),
                    data, font_size=9, color=GRAY,
                    align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)

    add_textbox(slide, Inches(0.6), Inches(6.65),
                SLIDE_W - Inches(1.2), Inches(0.30),
                'Источник: «Титан цвета» — аналитика первичных и вторичных продаж 2025-2026',
                font_size=8, color=GRAY, italic=True,
                align=PP_ALIGN.CENTER)

    footer(slide)


def slide_summary(prs):
    """Слайд 6: Сводная таблица."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H,
              fill_color=WHITE, line_color=None)

    header_band(slide, 'СВОДНАЯ ТАБЛИЦА — MML TITAN ПО КЛАССАМ ТТ',
                'Сравнение по категориям', color=DARK_BLUE)

    # Таблица
    headers = ['Категория / Объём', 'A — Федер. DIY', 'B — Локал. сеть', 'C — Супермаркет', 'D — Хозмаг']
    col_widths = [Inches(3.5), Inches(2.2), Inches(2.2), Inches(2.2), Inches(2.2)]
    table_y = Inches(1.0)
    header_h = Inches(0.45)

    cx = Inches(0.4)
    for i, h in enumerate(headers):
        add_shape(slide, MSO_SHAPE.RECTANGLE,
                  cx, table_y, col_widths[i], header_h,
                  fill_color=DARK_BLUE, line_color=WHITE, line_width=Pt(1))
        add_textbox(slide, cx, table_y, col_widths[i], header_h,
                    h, font_size=10, color=WHITE, bold=True,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        cx += col_widths[i]

    data = [
        ('3в1 по ржавчине, 0.75 л', '5 цветов (9005, 7004, 8017, 7024, 9010)', '5 цветов', '—', '—'),
        ('3в1 по ржавчине, 2 л', '5 цветов', '4 цвета', '—', '—'),
        ('Быстросохнущая, 0.75 л', '1 цвет (9005)', '1 цвет (9005)', '—', '—'),
        ('Быстросохнущая, 2.2 л', '1 цвет (9005)', '—', '—', '—'),
        ('Молотковая, 0.8 кг', '3 цвета (чёрн, кор, сереб)', '2 цвета (чёрн, кор)', '2 цвета (чёрн, кор)', '—'),
        ('Молотковая, 2.5 кг', '1 цвет (чёрный)', '—', '—', '—'),
    ]

    row_h = Inches(0.45)
    for ri, row in enumerate(data):
        ry = table_y + header_h + row_h * ri
        bg = LIGHT_GRAY if ri % 2 == 0 else WHITE
        cx = Inches(0.4)
        for ci, val in enumerate(row):
            add_shape(slide, MSO_SHAPE.RECTANGLE,
                      cx, ry, col_widths[ci], row_h,
                      fill_color=bg, line_color=GRAY, line_width=Pt(0.3))
            val_color = RED if val == '—' else DARK
            val_bold = (ci == 0)
            add_textbox(slide, cx + Inches(0.05), ry,
                        col_widths[ci] - Inches(0.10), row_h,
                        val, font_size=9, color=val_color, bold=val_bold,
                        align=PP_ALIGN.LEFT if ci == 0 else PP_ALIGN.CENTER,
                        anchor=MSO_ANCHOR.MIDDLE)
            cx += col_widths[ci]

    # ИТОГО
    ry = table_y + header_h + row_h * len(data)
    cx = Inches(0.4)
    totals = ['ИТОГО SKU', '16', '12', '2', '0']
    for ci, val in enumerate(totals):
        bg_color = DARK_BLUE if ci == 0 else [RED, ORANGE, GREEN, DARK_BLUE][ci-1]
        add_shape(slide, MSO_SHAPE.RECTANGLE,
                  cx, ry, col_widths[ci], row_h,
                  fill_color=bg_color, line_color=WHITE, line_width=Pt(1))
        add_textbox(slide, cx, ry, col_widths[ci], row_h,
                    val, font_size=14, color=WHITE, bold=True,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        cx += col_widths[ci]

    # Топ-цвета по продажам
    top_y = ry + row_h + Inches(0.30)
    add_textbox(slide, Inches(0.4), top_y,
                Inches(12.5), Inches(0.35),
                'ТОП-5 ЦВЕТОВ ПО ВТОРИЧНЫМ ПРОДАЖАМ (шт):',
                font_size=12, color=DARK_BLUE, bold=True, align=PP_ALIGN.LEFT)

    top_colors = [
        ('1', 'Чёрный RAL 9005', '7 738 шт', '393 АКБ', RGBColor(0x1A, 0x1A, 0x1A)),
        ('2', 'Графитовый RAL 7024', '5 172 шт', '240 АКБ', RGBColor(0x40, 0x40, 0x40)),
        ('3', 'Серый RAL 7004', '5 071 шт', '247 АКБ', RGBColor(0xB0, 0xB0, 0xB0)),
        ('4', 'Шоколад RAL 8017', '4 524 шт', '287 АКБ', RGBColor(0x3D, 0x28, 0x17)),
        ('5', 'Белый RAL 9010', '3 273 шт', '225 АКБ', RGBColor(0xF5, 0xF5, 0xDC)),
    ]
    card_w = Inches(2.45)
    card_h = Inches(1.70)
    card_y = top_y + Inches(0.40)

    for i, (num, name, qty, akb, color) in enumerate(top_colors):
        cx = Inches(0.4) + (card_w + Inches(0.10)) * i
        add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
                  cx, card_y, card_w, card_h,
                  fill_color=WHITE, line_color=color, line_width=Pt(2))
        # Цветной круг
        add_shape(slide, MSO_SHAPE.OVAL,
                  cx + Inches(0.10), card_y + Inches(0.10),
                  Inches(0.50), Inches(0.50),
                  fill_color=color, line_color=DARK, line_width=Pt(1))
        add_textbox(slide, cx + Inches(0.10), card_y + Inches(0.10),
                    Inches(0.50), Inches(0.50),
                    num, font_size=14, color=WHITE if color != RGBColor(0xF5, 0xF5, 0xDC) else DARK, bold=True,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        # Название
        add_textbox(slide, cx + Inches(0.70), card_y + Inches(0.10),
                    card_w - Inches(0.80), Inches(0.50),
                    name, font_size=9, color=DARK, bold=True,
                    align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)
        # Продажи
        add_textbox(slide, cx + Inches(0.10), card_y + Inches(0.70),
                    card_w - Inches(0.20), Inches(0.40),
                    qty, font_size=16, color=RED, bold=True,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        # АКБ
        add_textbox(slide, cx + Inches(0.10), card_y + Inches(1.15),
                    card_w - Inches(0.20), Inches(0.40),
                    akb, font_size=10, color=GRAY,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    footer(slide)


def main():
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H

    slide_title(prs)          # 1 — Титул
    slide_logic(prs)          # 2 — Логика выкладки
    slide_full_range(prs)      # 3 — Весь модельный ряд (3 полки)
    slide_class_a(prs)         # 4 — Класс A
    slide_class_b(prs)         # 5 — Класс B
    slide_class_cd(prs)        # 6 — Класс C и D
    slide_summary(prs)         # 7 — Сводная таблица + топ-5 цветов

    prs.save(OUTPUT)
    print(f'OK: {OUTPUT}')
    print(f'Слайдов: {len(prs.slides)}')


if __name__ == '__main__':
    main()
