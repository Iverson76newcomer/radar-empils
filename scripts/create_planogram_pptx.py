#!/usr/bin/env python3
"""
TITAN — Планограмма выкладки в формате PowerPoint.
3 линейки продукции · 3 полки · 22 SKU · зона ТСО · POS-материалы.
Слоган: "10 лет без перекраски — 10 лет экономии!"
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
ORANGE = RGBColor(0xEA, 0x58, 0x0C)
YELLOW = RGBColor(0xFB, 0xBF, 0x24)
BLUE = RGBColor(0x00, 0x47, 0xAB)
GOLD = RGBColor(0xB8, 0x86, 0x0B)
SILVER = RGBColor(0xC0, 0xC0, 0xC0)
COPPER = RGBColor(0xB8, 0x73, 0x33)

# RAL colors
RAL = {
    'RAL 9005': RGBColor(0x1A, 0x1A, 0x1A),  # Чёрный
    'RAL 9010': RGBColor(0xF5, 0xF5, 0xDC),  # Белый (бежевый)
    'RAL 3009': RGBColor(0x7B, 0x3F, 0x00),  # Красно-коричневый
    'RAL 6005': RGBColor(0x2D, 0x50, 0x16),  # Зелёный
    'RAL 7004': RGBColor(0xB0, 0xB0, 0xB0),  # Серый
    'RAL 7024': RGBColor(0x40, 0x40, 0x40),  # Графитовый серый
    'RAL 8017': RGBColor(0x3D, 0x28, 0x17),  # Шоколадно-коричневый
    'RAL 1021': RGBColor(0xE6, 0xB8, 0x00),  # Жёлтый
    'RAL 5005': RGBColor(0x00, 0x47, 0xAB),  # Синий
    'Золотистый': GOLD,
    'Серебристый': SILVER,
    'Медный': COPPER,
    'Тёмно-зелёный': RGBColor(0x35, 0x5E, 0x3B),
}

OUTPUT = '/home/z/my-project/download/Планограмма_TITAN.pptx'
PNG_FULL = '/home/z/my-project/download/planogram_titan/planogram_titan_full.png'

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
    """Add a textbox with multiple lines (each line = (text, size, color, bold))."""
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


def draw_can(slide, x, y, w, h, color, label_top, label_mid, label_bot,
             effect='gloss', text_color=None):
    """Нарисовать банку краски.
    color — RGBColor основного цвета
    label_top — TITAN
    label_mid — тип (3в1 / матовая / молот.)
    label_bot — RAL / объём
    effect: 'gloss' / 'matte' / 'hammer'
    """
    # Текст по цвету (для светлых банок — тёмный)
    if text_color is None:
        # Светлые банки
        light_colors = [RAL['RAL 9010'], RAL['RAL 7004'], RAL['RAL 9005'] and False]
        # Проверка яркости
        r, g, b = color[0], color[1], color[2]
        brightness = (0.299 * r + 0.587 * g + 0.114 * b)
        text_color = DARK if brightness > 140 else WHITE

    # Корпус банки (закруглённый прямоугольник)
    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h,
              fill_color=color, line_color=DARK, line_width=Pt(0.75))

    # Крышка сверху (узкий прямоугольник)
    lid_h = Inches(0.10)
    add_shape(slide, MSO_SHAPE.RECTANGLE,
              x + Inches(0.05), y - lid_h + Inches(0.03),
              w - Inches(0.10), lid_h,
              fill_color=GRAY, line_color=DARK, line_width=Pt(0.5))

    # Эффект (матовый — пунктирная рамка; молотковый — точки)
    if effect == 'matte':
        # Тонкая пунктирная рамка — пропустим, оставим индикатор в виде полоски снизу
        add_shape(slide, MSO_SHAPE.RECTANGLE,
                  x + Inches(0.10), y + h - Inches(0.10),
                  w - Inches(0.20), Inches(0.04),
                  fill_color=text_color, line_color=None)
    elif effect == 'hammer':
        # Несколько маленьких точек
        import random
        random.seed(int(x) + int(y))
        # Конвертируем ширину/высоту в дюймы (число)
        w_in = w if isinstance(w, (int, float)) else w.inches if hasattr(w, 'inches') else (w / 914400)
        h_in = h if isinstance(h, (int, float)) else h.inches if hasattr(h, 'inches') else (h / 914400)
        for _ in range(6):
            dx = random.uniform(0.15, w_in - 0.15)
            dy = random.uniform(0.20, h_in - 0.20)
            dot_r = Inches(0.04)
            add_shape(slide, MSO_SHAPE.OVAL,
                      x + Inches(dx) - dot_r, y + Inches(dy) - dot_r,
                      dot_r * 2, dot_r * 2,
                      fill_color=WHITE, line_color=None)

    # Текст: TITAN / тип / RAL / объём
    add_multi_text(slide, x, y + Inches(0.05), w, h - Inches(0.10),
                   [
                       (label_top, 9, text_color, True),
                       (label_mid, 7, text_color, True),
                       (label_bot, 6, text_color, False),
                   ])


def header_band(slide, title, subtitle=None, color=RED):
    """Верхняя красная плашка-заголовок слайда."""
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
    """Футер слайда."""
    add_textbox(slide, Inches(0.4), Inches(7.05),
                SLIDE_W - Inches(0.8), Inches(0.4),
                'Empils  •  TITAN  •  Планограмма выкладки  •  10 лет без перекраски — 10 лет экономии!',
                font_size=9, color=GRAY, bold=False,
                align=PP_ALIGN.CENTER, font_name='Arial', italic=True)


# ==============================================================
# СЛАЙД 1: ТИТУЛ
# ==============================================================

def slide_title(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # Белый фон
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H,
              fill_color=WHITE, line_color=None)

    # Большая красная плашка слева (вся высота)
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, Inches(5.0), SLIDE_H,
              fill_color=RED, line_color=None)
    # Жёлтая разделительная полоса
    add_shape(slide, MSO_SHAPE.RECTANGLE, Inches(5.0), 0, Inches(0.10), SLIDE_H,
              fill_color=YELLOW, line_color=None)

    # TITAN на красном фоне
    add_textbox(slide, Inches(0.4), Inches(1.2), Inches(4.5), Inches(1.0),
                'TITAN', font_size=72, color=WHITE, bold=True,
                align=PP_ALIGN.LEFT, font_name='Arial')
    add_textbox(slide, Inches(0.4), Inches(2.4), Inches(4.5), Inches(0.7),
                'ПЛАНОГРАММА', font_size=32, color=WHITE, bold=True,
                align=PP_ALIGN.LEFT, font_name='Arial')
    add_textbox(slide, Inches(0.4), Inches(3.2), Inches(4.5), Inches(0.6),
                'ВЫКЛАДКИ', font_size=32, color=WHITE, bold=True,
                align=PP_ALIGN.LEFT, font_name='Arial')

    # Линии декора
    add_shape(slide, MSO_SHAPE.RECTANGLE, Inches(0.4), Inches(4.1),
              Inches(2.5), Inches(0.05),
              fill_color=YELLOW, line_color=None)

    add_textbox(slide, Inches(0.4), Inches(4.3), Inches(4.5), Inches(0.6),
                '3 линейки продукции',
                font_size=18, color=WHITE, bold=True,
                align=PP_ALIGN.LEFT, font_name='Arial')
    add_textbox(slide, Inches(0.4), Inches(4.85), Inches(4.5), Inches(0.6),
                '3 полки · 22 SKU',
                font_size=18, color=WHITE, bold=True,
                align=PP_ALIGN.LEFT, font_name='Arial')
    add_textbox(slide, Inches(0.4), Inches(5.4), Inches(4.5), Inches(0.6),
                'Зона ТСО + POS-материалы',
                font_size=18, color=WHITE, bold=True,
                align=PP_ALIGN.LEFT, font_name='Arial')

    # Правая часть: слоган + 3 линейки
    # Большой слоган
    add_textbox(slide, Inches(5.5), Inches(1.0), Inches(7.5), Inches(1.0),
                '10 лет без перекраски',
                font_size=32, color=GREEN, bold=True,
                align=PP_ALIGN.LEFT, font_name='Arial')
    add_textbox(slide, Inches(5.5), Inches(1.9), Inches(7.5), Inches(1.0),
                '10 лет экономии!',
                font_size=32, color=GREEN, bold=True,
                align=PP_ALIGN.LEFT, font_name='Arial')

    add_shape(slide, MSO_SHAPE.RECTANGLE, Inches(5.5), Inches(2.95),
              Inches(2.0), Inches(0.05),
              fill_color=YELLOW, line_color=None)

    # Подзаголовок
    add_textbox(slide, Inches(5.5), Inches(3.15), Inches(7.5), Inches(0.6),
                '3 линейки — 3 повода для выкладки',
                font_size=18, color=DARK, bold=False,
                align=PP_ALIGN.LEFT, font_name='Arial', italic=True)

    # 3 карточки линеек
    card_y = Inches(3.9)
    card_w = Inches(2.45)
    card_h = Inches(2.4)
    cards = [
        ('Полка 1', '3в1 по ржавчине', '9 цветов RAL', RED, 'gloss'),
        ('Полка 2', 'Быстросохнущая', '7 цветов · 30 мин', DARK, 'matte'),
        ('Полка 3', 'Молотковый эффект', '6 цветов · декор', GOLD, 'hammer'),
    ]
    for i, (shelf, name, sub, color, _) in enumerate(cards):
        cx = Inches(5.5) + card_w * i + Inches(0.10) * i
        add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
                  cx, card_y, card_w, card_h,
                  fill_color=WHITE, line_color=color, line_width=Pt(2.5))
        # Цветная полоса сверху
        add_shape(slide, MSO_SHAPE.RECTANGLE,
                  cx, card_y, card_w, Inches(0.5),
                  fill_color=color, line_color=None)
        add_textbox(slide, cx, card_y + Inches(0.05),
                    card_w, Inches(0.4),
                    shelf, font_size=14, color=WHITE, bold=True,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        # Название линейки
        add_textbox(slide, cx + Inches(0.1), card_y + Inches(0.7),
                    card_w - Inches(0.2), Inches(0.8),
                    name, font_size=15, color=DARK, bold=True,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        # Подзаголовок
        add_textbox(slide, cx + Inches(0.1), card_y + Inches(1.5),
                    card_w - Inches(0.2), Inches(0.5),
                    sub, font_size=11, color=GRAY, bold=False,
                    align=PP_ALIGN.CENTER)
        # Большая цифра
        add_textbox(slide, cx + Inches(0.1), card_y + Inches(1.8),
                    card_w - Inches(0.2), Inches(0.5),
                    ['9', '7', '6'][i] + ' SKU',
                    font_size=14, color=color, bold=True,
                    align=PP_ALIGN.CENTER)

    # Футер
    add_textbox(slide, Inches(5.5), Inches(6.7), Inches(7.5), Inches(0.5),
                'Empils  •  TITAN  •  Грунт-эмаль 3-в-1  •  2026',
                font_size=11, color=GRAY, bold=False,
                align=PP_ALIGN.LEFT, font_name='Arial', italic=True)


# ==============================================================
# СЛАЙД 2: ОБЗОРНАЯ ПЛАНОГРАММА (встроенный PNG)
# ==============================================================

def slide_overview(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H,
              fill_color=WHITE, line_color=None)

    header_band(slide, 'ПЛАНОГРАММА ВЫКЛАДКИ TITAN — ОБЩИЙ ВИД',
                '3 полки · 22 SKU · ТСО-зона')

    # Вставляем PNG во всю рабочую область
    if os.path.exists(PNG_FULL):
        # Область под заголовком (с 0.85 до 7.0)
        img_top = Inches(0.85)
        img_h = Inches(6.10)
        # Сохраняем пропорции PNG (20:14 = 1.43)
        # Ширина = height * 1.43 = 8.72 inches
        img_w = Inches(6.10 * 1.43)
        # Если слишком широко — ограничиваем по ширине
        max_w = Inches(12.5)
        if img_w > max_w:
            img_w = max_w
            img_h = Inches(12.5 / 1.43)
        # Центрируем
        img_left = (SLIDE_W - img_w) // 2
        slide.shapes.add_picture(PNG_FULL, img_left, img_top,
                                 width=img_w, height=img_h)

    footer(slide)


# ==============================================================
# СЛАЙДЫ ПОЛОК 3-5
# ==============================================================

def slide_shelf(prs, shelf_num, shelf_name, shelf_color, shelf_subtitle,
                products, accent_color, tip_text):
    """Слайд одной полки с банками.
    products: список (ral_key, label_top, label_mid, label_bot, effect, volume)
    """
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H,
              fill_color=WHITE, line_color=None)

    header_band(slide, f'ПОЛКА {shelf_num}: {shelf_name}',
                shelf_subtitle, color=accent_color)

    # Зона полки (большая серая рамка)
    shelf_x = Inches(0.4)
    shelf_y = Inches(1.0)
    shelf_w = SLIDE_W - Inches(0.8)
    shelf_h = Inches(4.8)

    # Светло-серая подложка
    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
              shelf_x, shelf_y, shelf_w, shelf_h,
              fill_color=SHELF_BG, line_color=GRAY, line_width=Pt(1))

    # Коричневая полка снизу
    shelf_board_h = Inches(0.25)
    add_shape(slide, MSO_SHAPE.RECTANGLE,
              shelf_x + Inches(0.15),
              shelf_y + shelf_h - shelf_board_h - Inches(0.05),
              shelf_w - Inches(0.30), shelf_board_h,
              fill_color=SHELF_BROWN, line_color=SHELF_DARK_BROWN,
              line_width=Pt(1))

    # === БАНКИ ===
    n = len(products)
    # Отступы
    margin_x = Inches(0.5)
    gap = Inches(0.20)
    can_w = (shelf_w - margin_x * 2 - gap * (n - 1)) / n
    can_h = Inches(3.5)

    # Банки стоят на полке
    can_y = shelf_y + shelf_h - shelf_board_h - can_h - Inches(0.15)

    for i, (ral_key, l_top, l_mid, l_bot, effect, _vol) in enumerate(products):
        cx = shelf_x + margin_x + (can_w + gap) * i
        color = RAL[ral_key]
        draw_can(slide, cx, can_y, can_w, can_h, color,
                 l_top, l_mid, l_bot, effect=effect)

    # Подпись полки слева сверху
    add_textbox(slide, shelf_x + Inches(0.20), shelf_y + Inches(0.10),
                Inches(8.0), Inches(0.4),
                f'ПОЛКА {shelf_num}: {shelf_name}',
                font_size=14, color=accent_color, bold=True,
                align=PP_ALIGN.LEFT)
    add_textbox(slide, shelf_x + Inches(0.20), shelf_y + Inches(0.45),
                Inches(8.0), Inches(0.35),
                shelf_subtitle,
                font_size=10, color=GRAY, bold=False,
                align=PP_ALIGN.LEFT, italic=True)

    # Кол-во SKU справа сверху
    add_textbox(slide, shelf_x + shelf_w - Inches(2.0), shelf_y + Inches(0.15),
                Inches(1.8), Inches(0.5),
                f'{n} SKU', font_size=18, color=accent_color, bold=True,
                align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)

    # === ЖЁЛТАЯ ПОДСКАЗКА ВНИЗУ ===
    tip_y = Inches(6.05)
    tip_h = Inches(0.85)
    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
              shelf_x, tip_y, shelf_w, tip_h,
              fill_color=YELLOW, line_color=None)
    # Иконка "!" слева
    add_shape(slide, MSO_SHAPE.OVAL,
              shelf_x + Inches(0.20), tip_y + Inches(0.20),
              Inches(0.45), Inches(0.45),
              fill_color=accent_color, line_color=None)
    add_textbox(slide, shelf_x + Inches(0.20), tip_y + Inches(0.20),
                Inches(0.45), Inches(0.45),
                '!', font_size=22, color=WHITE, bold=True,
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    # Текст подсказки
    add_textbox(slide, shelf_x + Inches(0.85), tip_y + Inches(0.05),
                shelf_w - Inches(1.0), tip_h - Inches(0.10),
                tip_text, font_size=12, color=DARK, bold=True,
                align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)

    footer(slide)


# ==============================================================
# СЛАЙД 6: ЗОНА ТСО
# ==============================================================

def slide_tso(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H,
              fill_color=RGBColor(0xEC, 0xFD, 0xF5), line_color=None)

    header_band(slide, 'ЗОНА ТСО — ВЫКЛАДКА ВНИЗУ СТЕНДА',
                'Калькулятор · Выкрасы · QR', color=GREEN)

    # 3 карточки зон ТСО
    card_y = Inches(1.2)
    card_w = Inches(4.0)
    card_h = Inches(4.5)

    cards = [
        ('💼', '1. ПОСОБИЕ ПРОДАВЦА',
         ['Памятка для продавца:', '• Сравнение TITAN vs ПФ-115',
          '• 3 шага расчёта ТСО', '• Калькулятор экономии 61%',
          '• QR на онлайн-калькулятор'],
         GREEN),
        ('🎨', '2. ВЫКРАСЫ',
         ['Образцы выкрасов на стенде:', '• TITAN 3в1 (через 1 год)',
          '• ПФ-115 (через 1 год)', '• Видимая разница в износе',
          '• Маркировка «до/после»'],
         RED),
        ('📱', '3. QR-КОД',
         ['QR-код в правом нижнем углу:', '• Ссылка на калькулятор ТСО',
          '• Видеоинструкция TITAN', '• Каталог цветов RAL',
          '• Программа лояльности'],
         BLUE),
    ]

    for i, (emoji, title, lines, color) in enumerate(cards):
        cx = Inches(0.5) + (card_w + Inches(0.20)) * i
        # Карточка
        add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
                  cx, card_y, card_w, card_h,
                  fill_color=WHITE, line_color=color, line_width=Pt(2.5))
        # Цветная шапка карточки
        add_shape(slide, MSO_SHAPE.RECTANGLE,
                  cx, card_y, card_w, Inches(0.7),
                  fill_color=color, line_color=None)
        # Эмодзи
        tb = slide.shapes.add_textbox(cx + Inches(0.20), card_y + Inches(0.05),
                                      Inches(0.7), Inches(0.6))
        tf = tb.text_frame
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = emoji
        r.font.size = Pt(28)
        r.font.name = 'Segoe UI Emoji'
        # Заголовок
        add_textbox(slide, cx + Inches(0.95), card_y + Inches(0.10),
                    card_w - Inches(1.0), Inches(0.5),
                    title, font_size=14, color=WHITE, bold=True,
                    anchor=MSO_ANCHOR.MIDDLE)
        # Текст
        text_y = card_y + Inches(0.85)
        for j, line in enumerate(lines):
            add_textbox(slide, cx + Inches(0.20), text_y + Inches(0.35) * j,
                        card_w - Inches(0.4), Inches(0.35),
                        line, font_size=11, color=DARK, bold=(j == 0),
                        align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)

    # Итоговый блок внизу
    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
              Inches(0.5), Inches(6.05),
              SLIDE_W - Inches(1.0), Inches(0.85),
              fill_color=GREEN, line_color=None)
    add_textbox(slide, Inches(0.5), Inches(6.10),
                SLIDE_W - Inches(1.0), Inches(0.75),
                'Зона ТСО — закрывает 70% возражений покупателя по цене',
                font_size=18, color=WHITE, bold=True,
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    footer(slide)


# ==============================================================
# СЛАЙД 7: РЕКОМЕНДАЦИИ ПО ВЫКЛАДКЕ
# ==============================================================

def slide_recommendations(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H,
              fill_color=WHITE, line_color=None)

    header_band(slide, 'РЕКОМЕНДАЦИИ ПО ВЫКЛАДКЕ',
                '7 правил · POS-материалы')

    # 7 рекомендаций в 2 ряда
    recs = [
        ('1', '3в1 — на верхнюю полку',
         'Линейка 3в1 по ржавчине — самая продаваемая. Золотая полка — на уровне глаз.',
         RED),
        ('2', 'Цвета по нарастающей',
         'Тёмные цвета слева → светлые справа. Чёрный, графит, коричневый, серый, белый.',
         DARK),
        ('3', 'Дублируем фейсинги',
         'Топ-3 цвета (чёрный, коричневый, графит) — по 2 фейсинга. Остальные — по 1.',
         ORANGE),
        ('4', 'Молотковая — отдельно',
         'Линейка с молотковым эффектом не смешивается с гладкой. Своя полка, своя подсветка.',
         GOLD),
        ('5', 'Топпер с УТП',
         'Топпер «10 лет без перекраски — 10 лет экономии!» на самом видном месте.',
         GREEN),
        ('6', 'Воблеры на полках',
         'Жёлтые воблеры с УТП «1=3», «Экономия 61%» на каждой полке сверху.',
         YELLOW),
        ('7', 'Зона ТСО внизу',
         'Калькулятор + выкрасы + QR — в нижней части стеллажа, на уровне рук покупателя.',
         BLUE),
    ]

    # Сетка 4 + 3 (нижний ряд центрирован)
    rows = [recs[:4], recs[4:]]
    card_w = Inches(3.0)
    card_h = Inches(2.6)
    gap_x = Inches(0.15)
    gap_y = Inches(0.20)
    start_y = Inches(1.05)

    for row_idx, row in enumerate(rows):
        row_w = card_w * len(row) + gap_x * (len(row) - 1)
        start_x = (SLIDE_W - row_w) // 2
        for col_idx, (num, title, desc, color) in enumerate(row):
            cx = start_x + (card_w + gap_x) * col_idx
            cy = start_y + (card_h + gap_y) * row_idx
            # Карточка
            add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
                      cx, cy, card_w, card_h,
                      fill_color=WHITE, line_color=color, line_width=Pt(2))
            # Цветной круг с цифрой
            circle_d = Inches(0.7)
            add_shape(slide, MSO_SHAPE.OVAL,
                      cx + Inches(0.15), cy + Inches(0.15),
                      circle_d, circle_d,
                      fill_color=color, line_color=None)
            add_textbox(slide, cx + Inches(0.15), cy + Inches(0.15),
                        circle_d, circle_d,
                        num, font_size=24, color=WHITE, bold=True,
                        align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
            # Заголовок
            add_textbox(slide, cx + Inches(0.95), cy + Inches(0.20),
                        card_w - Inches(1.1), Inches(0.6),
                        title, font_size=14, color=color, bold=True,
                        align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)
            # Описание
            add_textbox(slide, cx + Inches(0.15), cy + Inches(1.0),
                        card_w - Inches(0.3), card_h - Inches(1.1),
                        desc, font_size=11, color=DARK, bold=False,
                        align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP)

    # Итоговая полоса внизу
    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
              Inches(0.5), Inches(6.55),
              SLIDE_W - Inches(1.0), Inches(0.45),
              fill_color=RED, line_color=None)
    add_textbox(slide, Inches(0.5), Inches(6.55),
                SLIDE_W - Inches(1.0), Inches(0.45),
                'Цель: повысить долю TITAN в категории ЛКМ с 8% до 15% за 6 месяцев',
                font_size=14, color=WHITE, bold=True,
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    footer(slide)


# ==============================================================
# ДАННЫЕ О ПРОДУКТАХ
# ==============================================================

# Полка 1: 3в1 по ржавчине — 9 цветов RAL
products_shelf1 = [
    ('RAL 9005', 'TITAN', '3в1', 'RAL 9005', 'gloss', '0.75 л'),
    ('RAL 9010', 'TITAN', '3в1', 'RAL 9010', 'gloss', '0.75 л'),
    ('RAL 3009', 'TITAN', '3в1', 'RAL 3009', 'gloss', '0.75 л'),
    ('RAL 6005', 'TITAN', '3в1', 'RAL 6005', 'gloss', '0.75 л'),
    ('RAL 7004', 'TITAN', '3в1', 'RAL 7004', 'gloss', '0.75 л'),
    ('RAL 7024', 'TITAN', '3в1', 'RAL 7024', 'gloss', '0.75 л'),
    ('RAL 8017', 'TITAN', '3в1', 'RAL 8017', 'gloss', '0.75 л'),
    ('RAL 1021', 'TITAN', '3в1', 'RAL 1021', 'gloss', '0.75 л'),
    ('RAL 5005', 'TITAN', '3в1', 'RAL 5005', 'gloss', '0.75 л'),
]

# Полка 2: Быстросохнущая матовая — 7 цветов RAL
products_shelf2 = [
    ('RAL 9010', 'TITAN', 'матовая', 'RAL 9010', 'matte', '0.75 л'),
    ('RAL 9005', 'TITAN', 'матовая', 'RAL 9005', 'matte', '0.75 л'),
    ('RAL 8017', 'TITAN', 'матовая', 'RAL 8017', 'matte', '0.75 л'),
    ('RAL 7024', 'TITAN', 'матовая', 'RAL 7024', 'matte', '0.75 л'),
    ('RAL 7004', 'TITAN', 'матовая', 'RAL 7004', 'matte', '0.75 л'),
    ('RAL 1021', 'TITAN', 'матовая', 'RAL 1021', 'matte', '0.75 л'),
    ('RAL 6005', 'TITAN', 'матовая', 'RAL 6005', 'matte', '0.75 л'),
]

# Полка 3: Молотковый эффект — 6 цветов
products_shelf3 = [
    ('Золотистый', 'TITAN', 'молот.', 'Золотистый', 'hammer', '0.8 л'),
    ('Серебристый', 'TITAN', 'молот.', 'Серебристый', 'hammer', '0.8 л'),
    ('RAL 8017', 'TITAN', 'молот.', 'RAL 8017', 'hammer', '0.8 л'),
    ('RAL 9005', 'TITAN', 'молот.', 'RAL 9005', 'hammer', '0.8 л'),
    ('Медный', 'TITAN', 'молот.', 'Медный', 'hammer', '0.8 л'),
    ('Тёмно-зелёный', 'TITAN', 'молот.', 'Тёмно-зелёный', 'hammer', '0.8 л'),
]


# ==============================================================
# СБОРКА
# ==============================================================

def main():
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H

    # Слайд 1: Титул
    slide_title(prs)

    # Слайд 2: Обзорная планограмма (PNG)
    slide_overview(prs)

    # Слайд 3: Полка 1 — 3в1 по ржавчине
    slide_shelf(prs, 1, 'Грунт-эмаль 3в1 по ржавчине', RED,
                '9 цветов RAL · 10 лет гарантии · по ржавчине без грунта',
                products_shelf1, RED,
                'Флагманская линейка. Самая продаваемая. Топ-3 цвета (чёрный, коричневый, графит) — по 2 фейсинга. Слева направо: от тёмных к светлым.')

    # Слайд 4: Полка 2 — Быстросохнущая матовая
    slide_shelf(prs, 2, 'Грунт-эмаль быстросохнущая (матовая)', DARK,
                '7 цветов RAL · сохнет 30 мин · до -15 °C',
                products_shelf2, DARK,
                'Премиум-сегмент для профессионалов. УТП «сохнет 30 минут» — на воблере. Ставится под флагманской линейкой.')

    # Слайд 5: Полка 3 — Молотковый эффект
    slide_shelf(prs, 3, 'Эмаль с молотковым эффектом', GOLD,
                '6 цветов · декоративный эффект · скрывает дефекты · 10 лет',
                products_shelf3, GOLD,
                'Декоративная линейка. Не смешивать с гладкими эмалями. Золотистый и серебристый — топ-2 SKU, дублируем фейсинги.')

    # Слайд 6: Зона ТСО
    slide_tso(prs)

    # Слайд 7: Рекомендации по выкладке
    slide_recommendations(prs)

    os.makedirs(os.path.dirname(OUTPUT), exist_ok=True)
    prs.save(OUTPUT)
    print(f'OK: {OUTPUT}')
    print(f'Слайдов: {len(prs.slides)}')


if __name__ == '__main__':
    main()
