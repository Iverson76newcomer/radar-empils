#!/usr/bin/env python3
"""
TITAN — Планограммы по классам ТТ (A, B, C, D).
На основе сводной матрицы TITAN по линейкам.
Для каждого класса — отдельный слайд с полкой и банками.
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
SHELF_DARK_BROWN = RGBColor(0x6B, 0x56, 0x35)
ORANGE = RGBColor(0xF4, 0x79, 0x20)  # оранжевый Эмпилс
YELLOW = RGBColor(0xFB, 0xBF, 0x24)
GOLD = RGBColor(0xB8, 0x86, 0x0B)
BLUE = RGBColor(0x25, 0x63, 0xEB)

# RAL colors (для банок-фигур)
RAL = {
    'RAL 9005': RGBColor(0x1A, 0x1A, 0x1A),  # Чёрный
    'RAL 9010': RGBColor(0xF5, 0xF5, 0xDC),  # Белый (бежевый)
    'RAL 7004': RGBColor(0xB0, 0xB0, 0xB0),  # Серый
    'RAL 7024': RGBColor(0x40, 0x40, 0x40),  # Графитовый серый
    'RAL 8017': RGBColor(0x3D, 0x28, 0x17),  # Шоколадно-коричневый
    'RAL 3009': RGBColor(0x7B, 0x3F, 0x00),  # Красно-коричневый
    'RAL 6005': RGBColor(0x2D, 0x50, 0x16),  # Зелёный
    'RAL 1021': RGBColor(0xE6, 0xB8, 0x00),  # Жёлтый
    'RAL 5005': RGBColor(0x00, 0x47, 0xAB),  # Синий
    'Чёрный (молот)': RGBColor(0x1A, 0x1A, 0x1A),
    'Коричневый (молот)': RGBColor(0x4A, 0x2C, 0x17),
    'Серебристый (молот)': RGBColor(0xC0, 0xC0, 0xC0),
}

OUTPUT = '/home/z/my-project/download/Планограммы_TITAN_по_классам_ТТ.pptx'
CANS_DIR = '/home/z/my-project/download/titan_cans'

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
                'Empils  •  TITAN  •  Планограмма по классам ТТ  •  10 лет без перекраски — 10 лет экономии!',
                font_size=9, color=GRAY, bold=False,
                align=PP_ALIGN.CENTER, font_name='Arial', italic=True)


def draw_can_simple(slide, x, y, w, h, color_rgb, label_top, label_mid, label_bot,
                    effect='gloss'):
    """Нарисовать банку как прямоугольник (если нет фото)."""
    # Цвет текста
    r, g, b = color_rgb[0], color_rgb[1], color_rgb[2]
    brightness = 0.299 * r + 0.587 * g + 0.114 * b
    text_color = DARK if brightness > 140 else WHITE

    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h,
              fill_color=color_rgb, line_color=DARK, line_width=Pt(0.75))

    # Крышка
    lid_h = Inches(0.10)
    add_shape(slide, MSO_SHAPE.RECTANGLE,
              x + Inches(0.05), y - lid_h + Inches(0.03),
              w - Inches(0.10), lid_h,
              fill_color=GRAY, line_color=DARK, line_width=Pt(0.5))

    # Эффект молотковый — точки
    if effect == 'hammer':
        import random
        random.seed(int(x) + int(y))
        w_in = w / 914400
        h_in = h / 914400
        for _ in range(6):
            dx = random.uniform(0.15, w_in - 0.15)
            dy = random.uniform(0.20, h_in - 0.20)
            dot_r = Inches(0.04)
            add_shape(slide, MSO_SHAPE.OVAL,
                      x + Inches(dx) - dot_r, y + Inches(dy) - dot_r,
                      dot_r * 2, dot_r * 2,
                      fill_color=WHITE, line_color=None)

    # Текст
    tf = slide.shapes.add_textbox(x, y + Inches(0.05), w, h - Inches(0.10)).text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = Pt(1); tf.margin_right = Pt(1)
    for i, (text, size, bold) in enumerate([
        (label_top, 9, True),
        (label_mid, 7, True),
        (label_bot, 6, False),
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


def draw_can_photo(slide, x, y, w, h, photo_path, label_text):
    """Вставить фото банки + подпись."""
    if os.path.exists(photo_path):
        slide.shapes.add_picture(photo_path, x, y, w, h)
    # Подпись под банкой
    add_textbox(slide, x - Inches(0.1), y + h + Inches(0.05),
                w + Inches(0.2), Inches(0.3),
                label_text, font_size=8, color=DARK, bold=True,
                align=PP_ALIGN.CENTER)


def draw_shelf_with_cans(slide, shelf_y, shelf_w, shelf_h, cans, shelf_title,
                         accent_color, use_photos=False, photo_map=None):
    """Нарисовать полку с банками.
    cans: список dict с ключами: color, label_top, label_mid, label_bot, effect, photo (опц.)
    use_photos: использовать фото вместо прямоугольников
    photo_map: dict {label_mid: photo_path} для замены
    """
    # Светло-серая подложка
    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
              Inches(0.4), shelf_y, shelf_w, shelf_h,
              fill_color=SHELF_BG, line_color=GRAY, line_width=Pt(1))

    # Коричневая полка снизу
    shelf_board_h = Inches(0.25)
    add_shape(slide, MSO_SHAPE.RECTANGLE,
              Inches(0.55), shelf_y + shelf_h - shelf_board_h - Inches(0.05),
              shelf_w - Inches(0.30), shelf_board_h,
              fill_color=SHELF_BROWN, line_color=SHELF_DARK_BROWN, line_width=Pt(1))

    n = len(cans)
    margin_x = Inches(0.5)
    gap = Inches(0.15)
    can_w = (shelf_w - margin_x * 2 - gap * (n - 1)) / n
    can_h = Inches(3.2)

    can_y = shelf_y + shelf_h - shelf_board_h - can_h - Inches(0.15)

    # Заголовок полки
    add_textbox(slide, Inches(0.55), shelf_y + Inches(0.10),
                Inches(8.0), Inches(0.4),
                shelf_title, font_size=14, color=accent_color, bold=True,
                align=PP_ALIGN.LEFT)

    # Кол-во SKU
    add_textbox(slide, Inches(0.55) + shelf_w - Inches(2.5), shelf_y + Inches(0.15),
                Inches(2.0), Inches(0.5),
                f'{n} SKU', font_size=18, color=accent_color, bold=True,
                align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)

    # Банки
    for i, can in enumerate(cans):
        cx = Inches(0.4) + margin_x + (can_w + gap) * i
        color = RAL.get(can.get('ral', can.get('color', '')), DARK)

        if use_photos and photo_map and can.get('label_mid') in photo_map:
            # Используем фото
            draw_can_photo(slide, cx, can_y, can_w, can_h,
                          photo_map[can['label_mid']],
                          f"{can.get('label_bot', '')}\n{can.get('label_mid', '')}")
        else:
            draw_can_simple(slide, cx, can_y, can_w, can_h, color,
                           can.get('label_top', 'TITAN'),
                           can.get('label_mid', ''),
                           can.get('label_bot', ''),
                           effect=can.get('effect', 'gloss'))


# ==============================================================
# СЛАЙД 1: ТИТУЛ
# ==============================================================

def slide_title(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H,
              fill_color=WHITE, line_color=None)

    # Большая красная плашка слева
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, Inches(5.0), SLIDE_H,
              fill_color=RED, line_color=None)
    add_shape(slide, MSO_SHAPE.RECTANGLE, Inches(5.0), 0, Inches(0.10), SLIDE_H,
              fill_color=YELLOW, line_color=None)

    add_textbox(slide, Inches(0.4), Inches(1.2), Inches(4.5), Inches(1.0),
                'TITAN', font_size=72, color=WHITE, bold=True,
                align=PP_ALIGN.LEFT, font_name='Arial')
    add_textbox(slide, Inches(0.4), Inches(2.4), Inches(4.5), Inches(0.7),
                'ПЛАНОГРАММЫ', font_size=32, color=WHITE, bold=True,
                align=PP_ALIGN.LEFT, font_name='Arial')
    add_textbox(slide, Inches(0.4), Inches(3.2), Inches(4.5), Inches(0.6),
                'ПО КЛАССАМ ТТ', font_size=32, color=WHITE, bold=True,
                align=PP_ALIGN.LEFT, font_name='Arial')

    add_shape(slide, MSO_SHAPE.RECTANGLE, Inches(0.4), Inches(4.1),
              Inches(2.5), Inches(0.05), fill_color=YELLOW, line_color=None)

    add_textbox(slide, Inches(0.4), Inches(4.3), Inches(4.5), Inches(0.6),
                'A — Федеральная сеть DIY',
                font_size=16, color=WHITE, bold=True, align=PP_ALIGN.LEFT)
    add_textbox(slide, Inches(0.4), Inches(4.85), Inches(4.5), Inches(0.6),
                'B — Локальная сеть',
                font_size=16, color=WHITE, bold=True, align=PP_ALIGN.LEFT)
    add_textbox(slide, Inches(0.4), Inches(5.4), Inches(4.5), Inches(0.6),
                'C — Супермаркет',
                font_size=16, color=WHITE, bold=True, align=PP_ALIGN.LEFT)
    add_textbox(slide, Inches(0.4), Inches(5.95), Inches(4.5), Inches(0.6),
                'D — Хозмаг',
                font_size=16, color=WHITE, bold=True, align=PP_ALIGN.LEFT)

    # Правая часть — слоган + сводка
    add_textbox(slide, Inches(5.5), Inches(1.0), Inches(7.5), Inches(1.0),
                'Сводная матрица TITAN',
                font_size=32, color=GREEN, bold=True,
                align=PP_ALIGN.LEFT, font_name='Arial')
    add_textbox(slide, Inches(5.5), Inches(1.9), Inches(7.5), Inches(1.0),
                'по 4 классам ТТ',
                font_size=32, color=GREEN, bold=True,
                align=PP_ALIGN.LEFT, font_name='Arial')

    add_shape(slide, MSO_SHAPE.RECTANGLE, Inches(5.5), Inches(2.95),
              Inches(2.0), Inches(0.05), fill_color=YELLOW, line_color=None)

    add_textbox(slide, Inches(5.5), Inches(3.15), Inches(7.5), Inches(0.6),
                '4 планограммы · 30 SKU TITAN',
                font_size=18, color=DARK, bold=False,
                align=PP_ALIGN.LEFT, italic=True)

    # Сводная таблица на титуле
    table_y = Inches(4.0)
    table_x = Inches(5.5)
    table_w = Inches(7.3)
    row_h = Inches(0.45)

    # Заголовок таблицы
    add_shape(slide, MSO_SHAPE.RECTANGLE, table_x, table_y,
              table_w, row_h, fill_color=ORANGE, line_color=None)
    headers = ['Класс ТТ', 'Канал', 'TITAN SKU', '3в1', 'Быстросох.', 'Молот.']
    col_w = table_w / len(headers)
    for i, h in enumerate(headers):
        add_textbox(slide, table_x + col_w * i, table_y,
                    col_w, row_h,
                    h, font_size=11, color=WHITE, bold=True,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # Данные
    data = [
        ('A', 'Федеральная DIY', '16', '10', '2', '4'),
        ('B', 'Локальная сеть', '12', '9', '1', '2'),
        ('C', 'Супермаркет', '2', '0', '0', '2'),
        ('D', 'Хозмаг', '0', '0', '0', '0'),
    ]
    for ri, row in enumerate(data):
        ry = table_y + row_h * (ri + 1)
        bg = LIGHT_GRAY if ri % 2 == 0 else WHITE
        add_shape(slide, MSO_SHAPE.RECTANGLE, table_x, ry,
                  table_w, row_h, fill_color=bg, line_color=GRAY, line_width=Pt(0.5))
        for ci, val in enumerate(row):
            color = RED if (ci == 2 and int(val) == 0) else DARK
            bold = (ci == 0) or (ci == 2 and int(val) > 0)
            add_textbox(slide, table_x + col_w * ci, ry,
                        col_w, row_h,
                        val, font_size=11, color=color, bold=bold,
                        align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # Подвал
    add_textbox(slide, Inches(5.5), Inches(6.9), Inches(7.5), Inches(0.5),
                'Empils  •  TITAN  •  2026',
                font_size=11, color=GRAY, bold=False,
                align=PP_ALIGN.LEFT, italic=True)


# ==============================================================
# СЛАЙД: КЛАСС A — ФЕДЕРАЛЬНАЯ СЕТЬ DIY (16 SKU)
# ==============================================================

def slide_class_a(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H,
              fill_color=WHITE, line_color=None)

    header_band(slide, 'КЛАСС A — ФЕДЕРАЛЬНАЯ СЕТЬ DIY',
                '16 SKU TITAN  •  3 линейки  •  5 цветов RAL', color=RED)

    # Зона полки
    shelf_y = Inches(0.95)
    shelf_w = SLIDE_W - Inches(0.8)
    shelf_h = Inches(5.6)

    # 3 подзоны: 3в1 / быстросохнущая / молотковая
    # Подзона 1: 3в1 (10 SKU)
    # Подзона 2: Быстросохнущая (2 SKU)
    # Подзона 3: Молотковая (4 SKU)
    # Распределим пропорционально ширине

    # Все 16 SKU в одну линию (узкие банки, но поместятся)
    all_cans = []

    # 3в1 по ржавчине 0.75 л — 5 цветов
    for color_name, ral in [
        ('Чёрный', 'RAL 9005'), ('Серый', 'RAL 7004'),
        ('Коричневый', 'RAL 8017'), ('Графит', 'RAL 7024'),
        ('Белый', 'RAL 9010')
    ]:
        all_cans.append({
            'label_top': 'TITAN',
            'label_mid': '3в1 0.75л',
            'label_bot': f'{color_name} {ral}',
            'ral': ral,
            'effect': 'gloss',
        })

    # 3в1 по ржавчине 2 л — 5 цветов
    for color_name, ral in [
        ('Чёрный', 'RAL 9005'), ('Серый', 'RAL 7004'),
        ('Коричневый', 'RAL 8017'), ('Графит', 'RAL 7024'),
        ('Белый', 'RAL 9010')
    ]:
        all_cans.append({
            'label_top': 'TITAN',
            'label_mid': '3в1 2л',
            'label_bot': f'{color_name} {ral}',
            'ral': ral,
            'effect': 'gloss',
        })

    # Быстросохнущая — 2 SKU (чёрный 0.75 + 2.2 л)
    for vol in ['0.75л', '2.2л']:
        all_cans.append({
            'label_top': 'TITAN',
            'label_mid': f'БЫСТР. {vol}',
            'label_bot': 'Чёрный RAL 9005',
            'ral': 'RAL 9005',
            'effect': 'matte',
        })

    # Молотковая — 4 SKU
    for label, color_name, ral_key in [
        ('Молот. 0.8кг', 'Чёрный', 'Чёрный (молот)'),
        ('Молот. 2.5кг', 'Чёрный', 'Чёрный (молот)'),
        ('Молот. 0.8кг', 'Коричневый', 'Коричневый (молот)'),
        ('Молот. 0.8кг', 'Серебристый', 'Серебристый (молот)'),
    ]:
        all_cans.append({
            'label_top': 'TITAN',
            'label_mid': label,
            'label_bot': color_name,
            'ral': ral_key,
            'effect': 'hammer',
        })

    draw_shelf_with_cans(slide, shelf_y, shelf_w, shelf_h, all_cans,
                         'ПОЛКА: TITAN — полный ассортимент (3в1 + быстросохнущая + молотковая)',
                         RED)

    # Подсказка снизу
    tip_y = Inches(6.65)
    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
              Inches(0.4), tip_y, SLIDE_W - Inches(0.8), Inches(0.35),
              fill_color=YELLOW, line_color=None)
    add_textbox(slide, Inches(0.6), tip_y, SLIDE_W - Inches(1.2), Inches(0.35),
                '!  Класс A — эталонная выкладка. Все 3 линейки TITAN представлены. Дублируем фейсинги на чёрный, коричневый, графит.',
                font_size=11, color=DARK, bold=True,
                anchor=MSO_ANCHOR.MIDDLE)

    footer(slide)


# ==============================================================
# СЛАЙД: КЛАСС B — ЛОКАЛЬНАЯ СЕТЬ (12 SKU)
# ==============================================================

def slide_class_b(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H,
              fill_color=WHITE, line_color=None)

    header_band(slide, 'КЛАСС B — ЛОКАЛЬНАЯ СЕТЬ',
                '12 SKU TITAN  •  3 линейки  •  нет белого 2л, нет серебр.', color=DARK)

    shelf_y = Inches(0.95)
    shelf_w = SLIDE_W - Inches(0.8)
    shelf_h = Inches(5.6)

    all_cans = []

    # 3в1 0.75 л — 5 цветов
    for color_name, ral in [
        ('Чёрный', 'RAL 9005'), ('Серый', 'RAL 7004'),
        ('Коричневый', 'RAL 8017'), ('Графит', 'RAL 7024'),
        ('Белый', 'RAL 9010')
    ]:
        all_cans.append({
            'label_top': 'TITAN',
            'label_mid': '3в1 0.75л',
            'label_bot': f'{color_name} {ral}',
            'ral': ral,
        })

    # 3в1 2 л — 4 цвета (без белого)
    for color_name, ral in [
        ('Чёрный', 'RAL 9005'), ('Серый', 'RAL 7004'),
        ('Коричневый', 'RAL 8017'), ('Графит', 'RAL 7024'),
    ]:
        all_cans.append({
            'label_top': 'TITAN',
            'label_mid': '3в1 2л',
            'label_bot': f'{color_name} {ral}',
            'ral': ral,
        })

    # Быстросохнущая — 1 SKU
    all_cans.append({
        'label_top': 'TITAN',
        'label_mid': 'БЫСТР. 0.75л',
        'label_bot': 'Чёрный RAL 9005',
        'ral': 'RAL 9005',
        'effect': 'matte',
    })

    # Молотковая — 2 SKU
    for color_name, ral_key in [
        ('Чёрный', 'Чёрный (молот)'),
        ('Коричневый', 'Коричневый (молот)'),
    ]:
        all_cans.append({
            'label_top': 'TITAN',
            'label_mid': 'Молот. 0.8кг',
            'label_bot': color_name,
            'ral': ral_key,
            'effect': 'hammer',
        })

    draw_shelf_with_cans(slide, shelf_y, shelf_w, shelf_h, all_cans,
                         'ПОЛКА: TITAN — 3 линейки (без белого 2л и серебристой молот.)',
                         DARK)

    # Подсказка
    tip_y = Inches(6.65)
    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
              Inches(0.4), tip_y, SLIDE_W - Inches(0.8), Inches(0.35),
              fill_color=YELLOW, line_color=None)
    add_textbox(slide, Inches(0.6), tip_y, SLIDE_W - Inches(1.2), Inches(0.35),
                '!  Класс B — основная выкладка. Отсутствует: 3в1 белый 2л, быстросохнущая 2.2л, молотковая серебристая, молотковая 2.5 кг.',
                font_size=11, color=DARK, bold=True,
                anchor=MSO_ANCHOR.MIDDLE)

    footer(slide)


# ==============================================================
# СЛАЙД: КЛАСС C — СУПЕРМАРКЕТ (2 SKU + рекомендуемые)
# ==============================================================

def slide_class_c(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H,
              fill_color=WHITE, line_color=None)

    header_band(slide, 'КЛАСС C — СУПЕРМАРКЕТ',
                'Сейчас: 2 SKU  •  Рекомендуем: 7 SKU  •  добавить 3в1', color=ORANGE)

    # Две полки рядом: «Как сейчас» и «Как предлагаем»
    half_w = (SLIDE_W - Inches(1.2)) / 2

    # ЛЕВАЯ ПОЛКА: «Как сейчас» — 2 SKU
    left_x = Inches(0.4)
    shelf_y = Inches(1.0)
    shelf_h = Inches(5.0)

    # Заголовок
    add_textbox(slide, left_x, shelf_y - Inches(0.05),
                half_w, Inches(0.4),
                'КАК СЕЙЧАС: 2 SKU', font_size=14, color=RED, bold=True,
                align=PP_ALIGN.CENTER)
    add_textbox(slide, left_x, shelf_y + Inches(0.30),
                half_w, Inches(0.3),
                'Только молотковая 0.8 кг', font_size=10, color=GRAY, italic=True,
                align=PP_ALIGN.CENTER)

    shelf_y2 = shelf_y + Inches(0.65)

    # Светлая подложка
    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
              left_x, shelf_y2, half_w, shelf_h - Inches(0.7),
              fill_color=SHELF_BG, line_color=GRAY, line_width=Pt(1))
    # Коричневая полка
    shelf_board_h = Inches(0.25)
    add_shape(slide, MSO_SHAPE.RECTANGLE,
              left_x + Inches(0.15), shelf_y2 + shelf_h - Inches(0.7) - shelf_board_h - Inches(0.05),
              half_w - Inches(0.30), shelf_board_h,
              fill_color=SHELF_BROWN, line_color=SHELF_DARK_BROWN, line_width=Pt(1))

    # 2 банки (крупные)
    cans_now = [
        {'label_top': 'TITAN', 'label_mid': 'Молот. 0.8кг', 'label_bot': 'Чёрный', 'ral': 'Чёрный (молот)', 'effect': 'hammer'},
        {'label_top': 'TITAN', 'label_mid': 'Молот. 0.8кг', 'label_bot': 'Коричневый', 'ral': 'Коричневый (молот)', 'effect': 'hammer'},
    ]
    can_w = Inches(1.6)
    can_h = Inches(3.0)
    gap = Inches(0.5)
    total_w = can_w * 2 + gap
    start_x = left_x + (half_w - total_w) // 2
    can_y = shelf_y2 + shelf_h - Inches(0.7) - shelf_board_h - can_h - Inches(0.15)
    for i, can in enumerate(cans_now):
        cx = start_x + (can_w + gap) * i
        color = RAL.get(can['ral'], DARK)
        draw_can_simple(slide, cx, can_y, can_w, can_h, color,
                       can['label_top'], can['label_mid'], can['label_bot'],
                       effect=can.get('effect', 'gloss'))

    # ПРАВАЯ ПОЛКА: «Как рекомендуем» — 7 SKU (2 + 5)
    right_x = Inches(0.4) + half_w + Inches(0.4)

    # Заголовок
    add_textbox(slide, right_x, shelf_y - Inches(0.05),
                half_w, Inches(0.4),
                'РЕКОМЕНДУЕМ: 7 SKU', font_size=14, color=GREEN, bold=True,
                align=PP_ALIGN.CENTER)
    add_textbox(slide, right_x, shelf_y + Inches(0.30),
                half_w, Inches(0.3),
                'Добавить 3в1 (5 SKU) + быстросохнущая', font_size=10, color=GRAY, italic=True,
                align=PP_ALIGN.CENTER)

    # Подложка
    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
              right_x, shelf_y2, half_w, shelf_h - Inches(0.7),
              fill_color=RGBColor(0xEC, 0xFD, 0xF5), line_color=GREEN, line_width=Pt(2))
    # Полка
    add_shape(slide, MSO_SHAPE.RECTANGLE,
              right_x + Inches(0.15), shelf_y2 + shelf_h - Inches(0.7) - shelf_board_h - Inches(0.05),
              half_w - Inches(0.30), shelf_board_h,
              fill_color=SHELF_BROWN, line_color=SHELF_DARK_BROWN, line_width=Pt(1))

    # 7 банок
    cans_rec = [
        # 3в1 0.75 л — 3 цвета
        {'label_top': 'TITAN', 'label_mid': '3в1 0.75л', 'label_bot': 'Чёрный RAL 9005', 'ral': 'RAL 9005'},
        {'label_top': 'TITAN', 'label_mid': '3в1 0.75л', 'label_bot': 'Корич. RAL 8017', 'ral': 'RAL 8017'},
        {'label_top': 'TITAN', 'label_mid': '3в1 0.75л', 'label_bot': 'Серый RAL 7004', 'ral': 'RAL 7004'},
        # 3в1 2 л — 1 чёрный (для ТСО-ворот)
        {'label_top': 'TITAN', 'label_mid': '3в1 2л', 'label_bot': 'Чёрный RAL 9005', 'ral': 'RAL 9005'},
        # Быстросохнущая 0.75 л
        {'label_top': 'TITAN', 'label_mid': 'БЫСТР. 0.75л', 'label_bot': 'Чёрный RAL 9005', 'ral': 'RAL 9005', 'effect': 'matte'},
        # Молотковая (уже есть)
        {'label_top': 'TITAN', 'label_mid': 'Молот. 0.8кг', 'label_bot': 'Чёрный', 'ral': 'Чёрный (молот)', 'effect': 'hammer'},
        {'label_top': 'TITAN', 'label_mid': 'Молот. 0.8кг', 'label_bot': 'Коричневый', 'ral': 'Коричневый (молот)', 'effect': 'hammer'},
    ]
    n = len(cans_rec)
    margin_x = Inches(0.3)
    gap2 = Inches(0.10)
    can_w2 = (half_w - margin_x * 2 - gap2 * (n - 1)) / n
    can_h2 = Inches(2.8)
    can_y2 = shelf_y2 + shelf_h - Inches(0.7) - shelf_board_h - can_h2 - Inches(0.15)
    for i, can in enumerate(cans_rec):
        cx = right_x + margin_x + (can_w2 + gap2) * i
        color = RAL.get(can['ral'], DARK)
        # Подсветка новых SKU (3в1) зелёной рамкой
        is_new = '3в1' in can['label_mid'] or 'БЫСТР' in can['label_mid']
        if is_new:
            add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
                      cx - Inches(0.04), can_y2 - Inches(0.04),
                      can_w2 + Inches(0.08), can_h2 + Inches(0.08),
                      fill_color=None, line_color=GREEN, line_width=Pt(2))
        draw_can_simple(slide, cx, can_y2, can_w2, can_h2, color,
                       can['label_top'], can['label_mid'], can['label_bot'],
                       effect=can.get('effect', 'gloss'))

    # Подсказка
    tip_y = Inches(6.4)
    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
              Inches(0.4), tip_y, SLIDE_W - Inches(0.8), Inches(0.6),
              fill_color=YELLOW, line_color=None)
    add_textbox(slide, Inches(0.6), tip_y, SLIDE_W - Inches(1.2), Inches(0.6),
                '!  В классе C нет TITAN 3в1 — главная проблема для ТСО-стратегии. Добавление 5 SKU 3в1 даёт покупателю возможность купить продукт, на котором построен калькулятор ТСО.',
                font_size=11, color=DARK, bold=True,
                anchor=MSO_ANCHOR.MIDDLE)

    footer(slide)


# ==============================================================
# СЛАЙД: КЛАСС D — ХОЗМАГ (0 SKU + рекомендуемые 3)
# ==============================================================

def slide_class_d(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H,
              fill_color=WHITE, line_color=None)

    header_band(slide, 'КЛАСС D — ХОЗЯЙСТВЕННЫЙ МАГАЗИН',
                'Сейчас: 0 SKU  •  Рекомендуем: 3 SKU TITAN 3в1', color=GREEN)

    # Две полки
    half_w = (SLIDE_W - Inches(1.2)) / 2
    left_x = Inches(0.4)
    shelf_y = Inches(1.0)
    shelf_h = Inches(5.0)

    # ЛЕВАЯ: «Как сейчас»
    add_textbox(slide, left_x, shelf_y - Inches(0.05),
                half_w, Inches(0.4),
                'КАК СЕЙЧАС: 0 SKU', font_size=14, color=RED, bold=True,
                align=PP_ALIGN.CENTER)
    add_textbox(slide, left_x, shelf_y + Inches(0.30),
                half_w, Inches(0.3),
                'TITAN полностью отсутствует', font_size=10, color=GRAY, italic=True,
                align=PP_ALIGN.CENTER)

    shelf_y2 = shelf_y + Inches(0.65)
    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
              left_x, shelf_y2, half_w, shelf_h - Inches(0.7),
              fill_color=RGBColor(0xFE, 0xE2, 0xE2), line_color=RED, line_width=Pt(2))
    add_shape(slide, MSO_SHAPE.RECTANGLE,
              left_x + Inches(0.15), shelf_y2 + shelf_h - Inches(0.7) - Inches(0.25) - Inches(0.05),
              half_w - Inches(0.30), Inches(0.25),
              fill_color=SHELF_BROWN, line_color=SHELF_DARK_BROWN, line_width=Pt(1))

    # Сообщение «ПУСТО»
    add_textbox(slide, left_x, shelf_y2 + Inches(1.0),
                half_w, Inches(1.0),
                'TITAN', font_size=48, color=RED, bold=True,
                align=PP_ALIGN.CENTER)
    add_textbox(slide, left_x, shelf_y2 + Inches(2.0),
                half_w, Inches(0.6),
                'не представлен', font_size=20, color=RED, bold=True,
                align=PP_ALIGN.CENTER)
    add_textbox(slide, left_x, shelf_y2 + Inches(2.7),
                half_w, Inches(0.5),
                'Покупатель уходит с ПФ-115', font_size=12, color=GRAY, italic=True,
                align=PP_ALIGN.CENTER)

    # ПРАВАЯ: «Как рекомендуем» — 3 SKU
    right_x = Inches(0.4) + half_w + Inches(0.4)

    add_textbox(slide, right_x, shelf_y - Inches(0.05),
                half_w, Inches(0.4),
                'РЕКОМЕНДУЕМ: 3 SKU', font_size=14, color=GREEN, bold=True,
                align=PP_ALIGN.CENTER)
    add_textbox(slide, right_x, shelf_y + Inches(0.30),
                half_w, Inches(0.3),
                'Минимальный набор TITAN 3в1 для хозмага', font_size=10, color=GRAY, italic=True,
                align=PP_ALIGN.CENTER)

    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
              right_x, shelf_y2, half_w, shelf_h - Inches(0.7),
              fill_color=RGBColor(0xEC, 0xFD, 0xF5), line_color=GREEN, line_width=Pt(2))
    add_shape(slide, MSO_SHAPE.RECTANGLE,
              right_x + Inches(0.15), shelf_y2 + shelf_h - Inches(0.7) - Inches(0.25) - Inches(0.05),
              half_w - Inches(0.30), Inches(0.25),
              fill_color=SHELF_BROWN, line_color=SHELF_DARK_BROWN, line_width=Pt(1))

    # 3 крупные банки
    cans_rec = [
        {'label_top': 'TITAN', 'label_mid': '3в1 0.75л', 'label_bot': 'Чёрный\nRAL 9005', 'ral': 'RAL 9005'},
        {'label_top': 'TITAN', 'label_mid': '3в1 0.75л', 'label_bot': 'Коричневый\nRAL 8017', 'ral': 'RAL 8017'},
        {'label_top': 'TITAN', 'label_mid': 'Молот. 0.8кг', 'label_bot': 'Чёрный', 'ral': 'Чёрный (молот)', 'effect': 'hammer'},
    ]
    n = 3
    can_w = Inches(1.5)
    can_h = Inches(3.0)
    gap = Inches(0.5)
    total_w = can_w * n + gap * (n - 1)
    start_x = right_x + (half_w - total_w) // 2
    can_y = shelf_y2 + shelf_h - Inches(0.7) - Inches(0.25) - can_h - Inches(0.15)
    for i, can in enumerate(cans_rec):
        cx = start_x + (can_w + gap) * i
        color = RAL.get(can['ral'], DARK)
        draw_can_simple(slide, cx, can_y, can_w, can_h, color,
                       can['label_top'], can['label_mid'], can['label_bot'],
                       effect=can.get('effect', 'gloss'))

    # Подсказка
    tip_y = Inches(6.4)
    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
              Inches(0.4), tip_y, SLIDE_W - Inches(0.8), Inches(0.6),
              fill_color=YELLOW, line_color=None)
    add_textbox(slide, Inches(0.6), tip_y, SLIDE_W - Inches(1.2), Inches(0.6),
                '!  Класс D — самая массовая точка продаж ЛКМ в России. Отсутствие TITAN = потеря 80% розничного канала. Минимум 3 SKU — обязательно.',
                font_size=11, color=DARK, bold=True,
                anchor=MSO_ANCHOR.MIDDLE)

    footer(slide)


# ==============================================================
# СБОРКА
# ==============================================================

def main():
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H

    # Карта фото для замены банок на реальные фото (где есть)
    photo_map = {
        # label_mid → photo_path
        '3в1 0.75л': os.path.join(CANS_DIR, 'titan_3v1_black_proc.png'),  # заглушка — будем использовать фото только в классе A для демо
    }

    slide_title(prs)
    slide_class_a(prs)
    slide_class_b(prs)
    slide_class_c(prs)
    slide_class_d(prs)

    prs.save(OUTPUT)
    print(f'OK: {OUTPUT}')
    print(f'Слайдов: {len(prs.slides)}')


if __name__ == '__main__':
    main()
