#!/usr/bin/env python3
"""
TITAN 3в1 — Инструкция за 5 шагов в формате PowerPoint.
Все пиктограммы нарисованы ВЕКТОРНЫМИ ФИГУРАМИ (MSO_SHAPE),
чтобы корректно отображаться в любой версии PowerPoint.
Слоган: "10 лет без перекраски — 10 лет экономии!"
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from lxml import etree
import os

# === COLORS ===
RED = RGBColor(0xC6, 0x28, 0x28)
DARK_RED = RGBColor(0x8E, 0x00, 0x00)
GREEN = RGBColor(0x16, 0xA3, 0x4A)
DARK_GREEN = RGBColor(0x0B, 0x6E, 0x31)
DARK = RGBColor(0x1A, 0x1F, 0x36)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
GRAY = RGBColor(0x9C, 0xA3, 0xAF)
LIGHT_GRAY = RGBColor(0xF3, 0xF4, 0xF6)
MID_GRAY = RGBColor(0xD1, 0xD5, 0xDB)
ORANGE = RGBColor(0xEA, 0x58, 0x0C)
BLUE = RGBColor(0x25, 0x63, 0xEB)
YELLOW = RGBColor(0xFB, 0xBF, 0x24)
BROWN = RGBColor(0x8B, 0x45, 0x13)
METAL = RGBColor(0xC0, 0xC0, 0xC0)
METAL_DARK = RGBColor(0x9A, 0x9A, 0x9A)
RUST = RGBColor(0x8B, 0x45, 0x13)
BLACK = RGBColor(0x00, 0x00, 0x00)

OUTPUT = '/home/z/my-project/download/Инструкция_TITAN_3в1_5_шагов.pptx'

# Slide size 16:9 (13.333 x 7.5 inches)
SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)


def add_textbox(slide, left, top, width, height, text, font_size=18,
                color=DARK, bold=False, align=PP_ALIGN.LEFT, font_name='Arial',
                anchor=MSO_ANCHOR.TOP):
    """Add a textbox with given text."""
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = Pt(2)
    tf.margin_right = Pt(2)
    tf.margin_top = Pt(2)
    tf.margin_bottom = Pt(2)
    tf.vertical_anchor = anchor
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(font_size)
    run.font.color.rgb = color
    run.font.bold = bold
    run.font.name = font_name
    return tb


def add_shape(slide, shape_type, left, top, width, height, fill_color=None,
              line_color=None, line_width=None):
    """Add a shape with fill and line."""
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


def add_shape_with_text(slide, shape_type, left, top, width, height, text,
                        fill_color, text_color=WHITE, font_size=18, bold=True,
                        line_color=None, line_width=None, font_name='Arial'):
    """Add a shape with centered text."""
    shp = add_shape(slide, shape_type, left, top, width, height,
                    fill_color=fill_color, line_color=line_color,
                    line_width=line_width)
    tf = shp.text_frame
    tf.word_wrap = True
    tf.margin_left = Pt(2)
    tf.margin_right = Pt(2)
    tf.margin_top = Pt(2)
    tf.margin_bottom = Pt(2)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = text
    run.font.size = Pt(font_size)
    run.font.color.rgb = text_color
    run.font.bold = bold
    run.font.name = font_name
    return shp


def remove_shadow(shape):
    """Remove shadow from a shape."""
    shape.shadow.inherit = False


# ==============================================================
# VECОRНЫЕ ПИКТОГРАММЫ (только MSO_SHAPE — без эмодзи)
# ==============================================================

def pictogram_clean(slide, cx, cy, scale=1.0):
    """Шаг 1: Очистить — металлическая пластина с ржавчиной + щётка.
    cx, cy — центр в EMU. Все размеры в EMU."""
    s = int(220000 * scale)  # базовый размер

    # Металлическая пластина (серый прямоугольник)
    plate_w = int(s * 1.6)
    plate_h = int(s * 0.55)
    plate = add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
                      cx - plate_w // 2, cy + int(s * 0.05),
                      plate_w, plate_h,
                      fill_color=METAL, line_color=METAL_DARK, line_width=Pt(1))

    # Пятна ржавчины (коричневые круги)
    rust_spots = [
        (cx - int(s * 0.45), cy + int(s * 0.20), int(s * 0.13)),
        (cx - int(s * 0.20), cy + int(s * 0.10), int(s * 0.10)),
        (cx + int(s * 0.05), cy + int(s * 0.25), int(s * 0.12)),
        (cx + int(s * 0.30), cy + int(s * 0.15), int(s * 0.09)),
        (cx - int(s * 0.55), cy + int(s * 0.30), int(s * 0.08)),
    ]
    for sx, sy, sr in rust_spots:
        add_shape(slide, MSO_SHAPE.OVAL,
                  sx - sr, sy - sr, sr * 2, sr * 2,
                  fill_color=RUST, line_color=None)

    # Щётка: ручка (коричневый вертикальный прямоугольник)
    handle_w = int(s * 0.18)
    handle_h = int(s * 0.55)
    add_shape(slide, MSO_SHAPE.RECTANGLE,
              cx + int(s * 0.35), cy - int(s * 0.55),
              handle_w, handle_h,
              fill_color=BROWN, line_color=None)

    # Щетина (жёлтый горизонтальный прямоугольник под ручкой)
    bristle_w = int(s * 0.30)
    bristle_h = int(s * 0.15)
    add_shape(slide, MSO_SHAPE.RECTANGLE,
              cx + int(s * 0.29), cy - int(s * 0.05),
              bristle_w, bristle_h,
              fill_color=YELLOW, line_color=None)

    # Вертикальные полоски на щетине (засечки)
    for i in range(4):
        x = cx + int(s * 0.32) + int(s * 0.06 * i)
        add_shape(slide, MSO_SHAPE.RECTANGLE,
                  x, cy - int(s * 0.05),
                  int(s * 0.025), bristle_h,
                  fill_color=ORANGE, line_color=None)


def pictogram_degrease(slide, cx, cy, scale=1.0):
    """Шаг 2: Обезжирить — пульверизатор с курком + капли."""
    s = int(220000 * scale)

    # Поверхность снизу (серая полоса)
    plate_w = int(s * 1.8)
    plate_h = int(s * 0.20)
    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
              cx - plate_w // 2, cy + int(s * 0.65),
              plate_w, plate_h,
              fill_color=METAL, line_color=METAL_DARK, line_width=Pt(1))

    # Корпус бутылки (синий прямоугольник со скруглением снизу)
    body_w = int(s * 0.70)
    body_h = int(s * 0.95)
    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
              cx - body_w // 2, cy - int(s * 0.05),
              body_w, body_h,
              fill_color=BLUE, line_color=DARK, line_width=Pt(1))

    # Головка пульверизатора (тёмный широкий прямоугольник сверху корпуса)
    head_w = int(s * 0.55)
    head_h = int(s * 0.20)
    add_shape(slide, MSO_SHAPE.RECTANGLE,
              cx - head_w // 2, cy - int(s * 0.30),
              head_w, head_h,
              fill_color=DARK, line_color=None)

    # «Носик» распылителя (узкий прямоугольник, выходящий вправо от головки)
    nose_w = int(s * 0.35)
    nose_h = int(s * 0.08)
    add_shape(slide, MSO_SHAPE.RECTANGLE,
              cx + int(s * 0.05), cy - int(s * 0.27),
              nose_w, nose_h,
              fill_color=DARK, line_color=None)

    # Курок-триггер (закруглённый прямоугольник снизу от головки, под наклоном)
    trig_w = int(s * 0.18)
    trig_h = int(s * 0.30)
    trig = add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
                     cx - int(s * 0.05), cy - int(s * 0.10),
                     trig_w, trig_h,
                     fill_color=DARK, line_color=None)
    trig.rotation = -15.0

    # Капли (3 синих круга справа от носика)
    drop_r = int(s * 0.08)
    for i, (dx, dy) in enumerate([(0.50, -0.30), (0.65, -0.10), (0.55, 0.10)]):
        add_shape(slide, MSO_SHAPE.OVAL,
                  cx + int(s * dx) - drop_r, cy + int(s * dy) - drop_r,
                  drop_r * 2, drop_r * 2,
                  fill_color=BLUE, line_color=None)

    # Звёздочки/блеск (2 жёлтые звёзды слева — чистота)
    star_r = int(s * 0.11)
    add_shape(slide, MSO_SHAPE.STAR_5_POINT,
              cx - int(s * 0.70) - star_r, cy - int(s * 0.20) - star_r,
              star_r * 2, star_r * 2,
              fill_color=YELLOW, line_color=None)
    star_r2 = int(s * 0.08)
    add_shape(slide, MSO_SHAPE.STAR_5_POINT,
              cx - int(s * 0.55) - star_r2, cy + int(s * 0.30) - star_r2,
              star_r2 * 2, star_r2 * 2,
              fill_color=YELLOW, line_color=None)


def pictogram_stir(slide, cx, cy, scale=1.0):
    """Шаг 3: Перемешать — банка с палочкой и пузырьками."""
    s = int(220000 * scale)

    # Корпус банки (красный прямоугольник со скруглением)
    can_w = int(s * 1.0)
    can_h = int(s * 1.15)
    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
              cx - can_w // 2, cy - can_h // 2 + int(s * 0.05),
              can_w, can_h,
              fill_color=RED, line_color=DARK_RED, line_width=Pt(1.5))

    # Крышка (серый прямоугольник сверху)
    lid_w = int(s * 1.1)
    lid_h = int(s * 0.15)
    add_shape(slide, MSO_SHAPE.RECTANGLE,
              cx - lid_w // 2, cy - can_h // 2 - int(s * 0.05),
              lid_w, lid_h,
              fill_color=GRAY, line_color=DARK, line_width=Pt(1))

    # Подпись TITAN на банке (размер шрифта масштабируется, без переноса)
    titan_fs = max(5, int(8 * scale))
    # Расширяем бокс в 2.2 раза шире банки, чтобы TITAN точно поместился
    tb_w = int(can_w * 2.2)
    tb = slide.shapes.add_textbox(cx - tb_w // 2, cy - int(s * 0.10),
                                  tb_w, int(s * 0.22))
    tf = tb.text_frame
    tf.margin_left = Pt(0); tf.margin_right = Pt(0)
    tf.margin_top = Pt(0); tf.margin_bottom = Pt(0)
    tf.word_wrap = False
    tf.auto_size = None
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = 'TITAN'
    r.font.size = Pt(titan_fs)
    r.font.bold = True
    r.font.color.rgb = WHITE
    r.font.name = 'Arial'

    # Палочка для перемешивания (диагональный коричневый прямоугольник)
    stick_w = int(s * 0.08)
    stick_h = int(s * 1.0)
    stick = add_shape(slide, MSO_SHAPE.RECTANGLE,
                      cx + int(s * 0.20), cy - int(s * 0.55),
                      stick_w, stick_h,
                      fill_color=BROWN, line_color=None)
    # Поворот палочки (около 25 градусов)
    stick.rotation = 25.0

    # Завихрение — 5 белых пузырьков разных размеров (крупные, с обводкой)
    bubble_specs = [
        (-0.30, 0.40, 0.13),  # большой, слева внизу
        (-0.15, 0.25, 0.10),  # средний
        (-0.32, 0.20, 0.08),  # маленький
        (0.10, 0.42, 0.11),   # средний справа
        (-0.20, 0.48, 0.09),  # маленький
    ]
    for dx, dy, dr in bubble_specs:
        br = int(s * dr)
        add_shape(slide, MSO_SHAPE.OVAL,
                  cx + int(s * dx) - br, cy + int(s * dy) - br,
                  br * 2, br * 2,
                  fill_color=WHITE, line_color=DARK_RED, line_width=Pt(0.5))


def pictogram_paint(slide, cx, cy, scale=1.0):
    """Шаг 4: Покрасить — кисть на поверхности."""
    s = int(220000 * scale)

    # Поверхность: серая слева, зелёная справа (покрашенная часть)
    plate_w = int(s * 1.7)
    plate_h = int(s * 0.45)
    # Зелёная часть (вся)
    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
              cx - plate_w // 2, cy + int(s * 0.30),
              plate_w, plate_h,
              fill_color=GREEN, line_color=DARK_GREEN, line_width=Pt(1))
    # Серая часть слева (непокрашенная)
    unpaint_w = int(plate_w * 0.45)
    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
              cx - plate_w // 2, cy + int(s * 0.30),
              unpaint_w, plate_h,
              fill_color=METAL, line_color=DARK, line_width=Pt(1))

    # Кисть: ручка (коричневый прямоугольник, наклонён)
    handle_w = int(s * 0.14)
    handle_h = int(s * 0.70)
    handle = add_shape(slide, MSO_SHAPE.RECTANGLE,
                       cx + int(s * 0.05), cy - int(s * 0.50),
                       handle_w, handle_h,
                       fill_color=BROWN, line_color=None)
    handle.rotation = -8.0

    # Металлическое кольцо кисти (серый тонкий прямоугольник)
    ring_w = int(s * 0.22)
    ring_h = int(s * 0.10)
    add_shape(slide, MSO_SHAPE.RECTANGLE,
              cx + int(s * 0.01), cy + int(s * 0.18),
              ring_w, ring_h,
              fill_color=METAL_DARK, line_color=None)

    # Щетина кисти (тёмный прямоугольник)
    bristle_w = int(s * 0.22)
    bristle_h = int(s * 0.18)
    add_shape(slide, MSO_SHAPE.RECTANGLE,
              cx + int(s * 0.01), cy + int(s * 0.28),
              bristle_w, bristle_h,
              fill_color=DARK, line_color=None)

    # Стрелка движения (красная, вправо)
    arrow_w = int(s * 0.55)
    arrow_h = int(s * 0.20)
    add_shape(slide, MSO_SHAPE.RIGHT_ARROW,
              cx - int(s * 0.55), cy + int(s * 0.55),
              arrow_w, arrow_h,
              fill_color=RED, line_color=None)


def pictogram_done(slide, cx, cy, scale=1.0, show_years=True):
    """Шаг 5: Готово — зелёный круг-щит с белой галочкой (нарисована прямоугольниками).
    show_years=False — скрыть надпись '10 лет' (для маленьких версий)."""
    import math
    s = int(220000 * scale)

    # Зелёный круг-щит — увеличенный размер, чтобы галочка была хорошо видна
    shield_d = int(s * 3.0)
    add_shape(slide, MSO_SHAPE.OVAL,
              cx - shield_d // 2, cy - shield_d // 2 + int(s * 0.20),
              shield_d, shield_d,
              fill_color=GREEN, line_color=DARK_GREEN, line_width=Pt(2))

    # Толстая линия как прямоугольник с острыми углами
    def thick_line(x1, y1, x2, y2, thickness, color):
        dx = x2 - x1
        dy = y2 - y1
        length = int(math.sqrt(dx * dx + dy * dy))
        angle = math.degrees(math.atan2(dy, dx))
        lcx = (x1 + x2) // 2
        lcy = (y1 + y2) // 2
        rect = add_shape(slide, MSO_SHAPE.RECTANGLE,
                         lcx - thickness // 2, lcy - length // 2,
                         thickness, length,
                         fill_color=color, line_color=None)
        rect.rotation = angle
        return rect

    # Толщина линий галочки
    line_thickness = max(int(s * 0.14), 30000)

    # Точка V (нижняя вершина галочки)
    vx = cx - int(s * 0.55)
    vy = cy + int(s * 0.30)
    # Средняя точка M (где линии встречаются)
    mx = cx - int(s * 0.15)
    my = cy - int(s * 0.30)
    # Правый верхний конец (далеко и высоко)
    rx = cx + int(s * 0.75)
    ry = cy - int(s * 0.60)

    # Левая часть галочки (короткая, снизу-вверх-вправо)
    thick_line(vx, vy, mx, my, line_thickness, WHITE)
    # Правая часть галочки (длинная, снизу-вверх-вправо, более полого)
    thick_line(mx, my, rx, ry, line_thickness, WHITE)

    # "10 лет" — текст под щитом (только если show_years)
    if show_years:
        years_fs = max(10, int(20 * scale))
        tb = slide.shapes.add_textbox(cx - int(s * 0.7), cy + int(s * 1.50),
                                      int(s * 1.4), int(s * 0.35))
        tf = tb.text_frame
        tf.margin_left = Pt(0); tf.margin_right = Pt(0)
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = '10 лет'
        r.font.size = Pt(years_fs)
        r.font.bold = True
        r.font.color.rgb = GREEN
        r.font.name = 'Arial'


# ==============================================================
# СЛАЙДЫ
# ==============================================================

def slide_title(prs):
    """Слайд 1: Титульный."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank

    # Фоновая плашка (весь слайд — светло-серый)
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H,
              fill_color=WHITE, line_color=None)

    # Красная верхняя плашка-заголовок
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, Inches(2.2),
              fill_color=RED, line_color=None)

    # Тонкая жёлтая полоса под красной плашкой
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0, Inches(2.2), SLIDE_W, Inches(0.08),
              fill_color=YELLOW, line_color=None)

    # TITAN 3в1
    add_textbox(slide, Inches(0.5), Inches(0.4), SLIDE_W - Inches(1.0), Inches(1.0),
                'TITAN 3в1', font_size=54, color=WHITE, bold=True,
                align=PP_ALIGN.CENTER, font_name='Arial')

    # Подзаголовок
    add_textbox(slide, Inches(0.5), Inches(1.45), SLIDE_W - Inches(1.0), Inches(0.6),
                'Инструкция за 5 шагов — 10 лет защиты',
                font_size=24, color=WHITE, bold=False,
                align=PP_ALIGN.CENTER, font_name='Arial')

    # Большой слоган в центре
    add_textbox(slide, Inches(0.8), Inches(2.7), SLIDE_W - Inches(1.6), Inches(1.2),
                '10 лет без перекраски — 10 лет экономии!',
                font_size=40, color=GREEN, bold=True,
                align=PP_ALIGN.CENTER, font_name='Arial')

    # Подпись под слоганом
    add_textbox(slide, Inches(1.0), Inches(3.9), SLIDE_W - Inches(2.0), Inches(0.6),
                'Соблюдайте 5 простых шагов — и покрытие прослужит 10 лет',
                font_size=18, color=DARK, bold=False,
                align=PP_ALIGN.CENTER, font_name='Arial')

    # 5 кругов с цифрами и пиктограммами в ряд
    picto_y = Inches(5.0)
    spacing = Inches(2.3)
    start_x = (SLIDE_W - spacing * 4 - Inches(1.4)) // 2  # центрирование

    picto_fns = [pictogram_clean, pictogram_degrease, pictogram_stir,
                 pictogram_paint, pictogram_done]
    step_names = ['ОЧИСТИТЬ', 'ОБЕЗЖИРИТЬ', 'ПЕРЕМЕШАТЬ',
                  'ПОКРАСИТЬ', 'ГОТОВО!']

    for i in range(5):
        cx = start_x + spacing * i + Inches(0.7)
        # Красный круг
        add_shape(slide, MSO_SHAPE.OVAL,
                  cx - Inches(0.7), picto_y,
                  Inches(1.4), Inches(1.4),
                  fill_color=RED, line_color=DARK_RED, line_width=Pt(2))
        # Цифра вверху круга
        add_textbox(slide, cx - Inches(0.7), picto_y + Inches(0.05),
                    Inches(1.4), Inches(0.35),
                    str(i + 1), font_size=18, color=WHITE, bold=True,
                    align=PP_ALIGN.CENTER)
        # Пиктограмма (в центре круга)
        if i == 4:
            # Пятую пиктограмму (Готово) рисуем без надписи "10 лет"
            # — иначе текст вылезает за пределы маленького круга
            picto_fns[i](slide, cx, picto_y + Inches(0.80), scale=0.55,
                         show_years=False)
        else:
            picto_fns[i](slide, cx, picto_y + Inches(0.80), scale=0.55)
        # Название шага под кругом
        add_textbox(slide, cx - Inches(1.0), picto_y + Inches(1.5),
                    Inches(2.0), Inches(0.4),
                    step_names[i], font_size=11, color=DARK, bold=True,
                    align=PP_ALIGN.CENTER)

    # Футер
    add_textbox(slide, Inches(0.5), Inches(7.0), SLIDE_W - Inches(1.0), Inches(0.4),
                'Empils  •  TITAN  •  Грунт-эмаль по ржавчине 3-в-1',
                font_size=11, color=GRAY, bold=False,
                align=PP_ALIGN.CENTER, font_name='Arial')

    return slide


def slide_step(prs, step_num, title, lines, tip_text,
               picto_fn, table_rows=None):
    """Слайд для одного шага.
    step_num: 1..5
    title: 'ОЧИСТИТЬ'
    lines: список строк описания
    tip_text: текст жёлтой подсказки
    picto_fn: функция пиктограммы
    table_rows: опционально — список (label, value) для таблицы
    """
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # Белый фон
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H,
              fill_color=WHITE, line_color=None)

    # Тонкая красная плашка вверху
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, Inches(0.6),
              fill_color=RED, line_color=None)

    # Название TITAN 3в1 слева вверху
    add_textbox(slide, Inches(0.4), Inches(0.10), Inches(4.0), Inches(0.4),
                'TITAN 3в1  •  Инструкция',
                font_size=14, color=WHITE, bold=True,
                align=PP_ALIGN.LEFT, font_name='Arial')

    # Номер шага справа вверху
    add_textbox(slide, SLIDE_W - Inches(2.0), Inches(0.10),
                Inches(1.6), Inches(0.4),
                f'Шаг {step_num} из 5',
                font_size=14, color=WHITE, bold=True,
                align=PP_ALIGN.RIGHT, font_name='Arial')

    # === ЛЕВАЯ ЧАСТЬ: Большой круг с цифрой и пиктограммой ===
    circle_cx = Inches(3.0)
    circle_cy = Inches(4.1)
    circle_size = Inches(3.6)

    # Красный круг
    add_shape(slide, MSO_SHAPE.OVAL,
              circle_cx - circle_size // 2,
              circle_cy - circle_size // 2,
              circle_size, circle_size,
              fill_color=RED, line_color=DARK_RED, line_width=Pt(3))

    # Цифра шага (большая, вверху круга)
    add_textbox(slide,
                circle_cx - circle_size // 2,
                circle_cy - circle_size // 2 + Inches(0.20),
                circle_size, Inches(0.9),
                str(step_num), font_size=48, color=WHITE, bold=True,
                align=PP_ALIGN.CENTER)

    # Пиктограмма в центре круга
    if step_num == 5:
        picto_fn(slide, circle_cx, circle_cy + Inches(0.55), scale=1.4,
                 show_years=True)
    else:
        picto_fn(slide, circle_cx, circle_cy + Inches(0.65), scale=1.4)

    # === ПРАВАЯ ЧАСТЬ: Заголовок + описание + подсказка ===
    right_x = Inches(5.5)
    right_w = SLIDE_W - right_x - Inches(0.5)

    # Заголовок шага (большой красный текст)
    title_color = GREEN if step_num == 5 else RED
    add_textbox(slide, right_x, Inches(1.0), right_w, Inches(0.9),
                title, font_size=40, color=title_color, bold=True,
                align=PP_ALIGN.LEFT, font_name='Arial')

    # Жёлтая разделительная полоса
    add_shape(slide, MSO_SHAPE.RECTANGLE,
              right_x, Inches(1.85),
              Inches(2.0), Inches(0.06),
              fill_color=YELLOW, line_color=None)

    # Описание (строки)
    desc_y = Inches(2.1)
    line_height = Inches(0.45)
    for i, line in enumerate(lines):
        is_green = line.startswith('GREEN:')
        clean_line = line[6:] if is_green else line
        add_textbox(slide, right_x, desc_y + line_height * i,
                    right_w, line_height,
                    clean_line, font_size=18,
                    color=GREEN if is_green else DARK,
                    bold=is_green,
                    align=PP_ALIGN.LEFT, font_name='Arial')

    # Жёлтая подсказка
    if tip_text:
        tip_y = desc_y + line_height * len(lines) + Inches(0.15)
        tip_h = Inches(0.75)
        # Жёлтый фон
        add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
                  right_x, tip_y, right_w, tip_h,
                  fill_color=YELLOW, line_color=None)
        # Текст подсказки
        tb = slide.shapes.add_textbox(right_x + Inches(0.15),
                                      tip_y + Inches(0.05),
                                      right_w - Inches(0.30), tip_h - Inches(0.10))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf.margin_left = Pt(4); tf.margin_right = Pt(4)
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.LEFT
        r = p.add_run()
        r.text = '!  ' + tip_text
        r.font.size = Pt(13)
        r.font.bold = True
        r.font.color.rgb = DARK
        r.font.name = 'Arial'

    # Таблица (опционально)
    if table_rows:
        tbl_y = desc_y + line_height * len(lines) + Inches(1.10)
        tbl_h = Inches(0.65)
        col_w = right_w // len(table_rows)
        for i, (label, value) in enumerate(table_rows):
            x = right_x + col_w * i
            # Светло-серая плашка
            add_shape(slide, MSO_SHAPE.RECTANGLE,
                      x, tbl_y, col_w, tbl_h,
                      fill_color=LIGHT_GRAY, line_color=GRAY, line_width=Pt(0.5))
            # Текст: метка + значение
            tb = slide.shapes.add_textbox(x + Inches(0.05),
                                          tbl_y + Inches(0.05),
                                          col_w - Inches(0.10), tbl_h - Inches(0.10))
            tf = tb.text_frame
            tf.word_wrap = True
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            r1 = p.add_run()
            r1.text = label + '\n'
            r1.font.size = Pt(11)
            r1.font.bold = True
            r1.font.color.rgb = DARK
            r1.font.name = 'Arial'
            r2 = p.add_run()
            r2.text = value
            r2.font.size = Pt(13)
            r2.font.bold = True
            r2.font.color.rgb = RED
            r2.font.name = 'Arial'

    # Стрелка вниз (между шагами) — внизу справа
    if step_num < 5:
        add_shape(slide, MSO_SHAPE.DOWN_ARROW,
                  SLIDE_W - Inches(1.2), Inches(6.7),
                  Inches(0.5), Inches(0.4),
                  fill_color=GRAY, line_color=None)
        add_textbox(slide, SLIDE_W - Inches(2.0), Inches(6.75),
                    Inches(0.8), Inches(0.35),
                    'далее', font_size=10, color=GRAY, bold=False,
                    align=PP_ALIGN.RIGHT)

    # Футер
    add_textbox(slide, Inches(0.5), Inches(7.1), SLIDE_W - Inches(1.0), Inches(0.35),
                'Empils  •  TITAN  •  Грунт-эмаль 3-в-1  •  10 лет без перекраски — 10 лет экономии!',
                font_size=10, color=GRAY, bold=False,
                align=PP_ALIGN.CENTER, font_name='Arial')

    return slide


def slide_tso(prs):
    """Слайд 7: ТСО — итоговый слоган."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # Зелёный фон (весь слайд)
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H,
              fill_color=RGBColor(0xEC, 0xFD, 0xF5), line_color=None)

    # Зелёная плашка-заголовок
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, Inches(2.0),
              fill_color=GREEN, line_color=None)

    # Жёлтая полоса
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0, Inches(2.0), SLIDE_W, Inches(0.08),
              fill_color=YELLOW, line_color=None)

    # TITAN 3в1 вверху
    add_textbox(slide, Inches(0.5), Inches(0.3), SLIDE_W - Inches(1.0), Inches(0.7),
                'TITAN 3в1  •  Итог', font_size=22, color=WHITE, bold=True,
                align=PP_ALIGN.CENTER, font_name='Arial')

    # Большой слоган
    add_textbox(slide, Inches(0.5), Inches(0.9), SLIDE_W - Inches(1.0), Inches(1.0),
                '10 лет без перекраски — 10 лет экономии!',
                font_size=40, color=WHITE, bold=True,
                align=PP_ALIGN.CENTER, font_name='Arial')

    # === 3 КРУПНЫЕ ЦИФРЫ ===
    # 1) 1 банка вместо 3
    card1_x = Inches(0.8)
    card_y = Inches(2.6)
    card_w = Inches(3.7)
    card_h = Inches(3.5)

    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
              card1_x, card_y, card_w, card_h,
              fill_color=WHITE, line_color=GREEN, line_width=Pt(2))
    add_textbox(slide, card1_x, card_y + Inches(0.3),
                card_w, Inches(1.4),
                '1', font_size=120, color=GREEN, bold=True,
                align=PP_ALIGN.CENTER)
    add_textbox(slide, card1_x, card_y + Inches(1.7),
                card_w, Inches(0.5),
                'банка вместо 3', font_size=22, color=DARK, bold=True,
                align=PP_ALIGN.CENTER)
    add_textbox(slide, card1_x, card_y + Inches(2.3),
                card_w, Inches(1.0),
                'грунт + эмаль +\nпреобразователь ржавчины\nв одной банке',
                font_size=14, color=GRAY, bold=False,
                align=PP_ALIGN.CENTER)

    # 2) 61% экономии
    card2_x = Inches(4.8)
    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
              card2_x, card_y, card_w, card_h,
              fill_color=WHITE, line_color=GREEN, line_width=Pt(2))
    add_textbox(slide, card2_x, card_y + Inches(0.3),
                card_w, Inches(1.4),
                '61%', font_size=100, color=GREEN, bold=True,
                align=PP_ALIGN.CENTER)
    add_textbox(slide, card2_x, card_y + Inches(1.7),
                card_w, Inches(0.5),
                'экономия за 10 лет', font_size=22, color=DARK, bold=True,
                align=PP_ALIGN.CENTER)
    add_textbox(slide, card2_x, card_y + Inches(2.3),
                card_w, Inches(1.0),
                'покрасил 1 раз —\nзабыл на 10 лет',
                font_size=14, color=GRAY, bold=False,
                align=PP_ALIGN.CENTER)

    # 3) 10 лет защиты
    card3_x = Inches(8.8)
    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
              card3_x, card_y, card_w, card_h,
              fill_color=GREEN, line_color=DARK_GREEN, line_width=Pt(2))
    add_textbox(slide, card3_x, card_y + Inches(0.3),
                card_w, Inches(1.4),
                '10', font_size=120, color=WHITE, bold=True,
                align=PP_ALIGN.CENTER)
    add_textbox(slide, card3_x, card_y + Inches(1.7),
                card_w, Inches(0.5),
                'лет защиты', font_size=22, color=WHITE, bold=True,
                align=PP_ALIGN.CENTER)
    add_textbox(slide, card3_x, card_y + Inches(2.3),
                card_w, Inches(1.0),
                'срок службы\nпокрытия без\nперекраски',
                font_size=14, color=WHITE, bold=False,
                align=PP_ALIGN.CENTER)

    # Итоговый слоган внизу
    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
              Inches(0.8), Inches(6.4),
              SLIDE_W - Inches(1.6), Inches(0.8),
              fill_color=GREEN, line_color=None)
    add_textbox(slide, Inches(0.8), Inches(6.45),
                SLIDE_W - Inches(1.6), Inches(0.7),
                'TITAN 3в1 — 10 лет без перекраски, 10 лет экономии!',
                font_size=22, color=WHITE, bold=True,
                align=PP_ALIGN.CENTER, font_name='Arial',
                anchor=MSO_ANCHOR.MIDDLE)

    return slide


# ==============================================================
# СБОРКА ПРЕЗЕНТАЦИИ
# ==============================================================

def main():
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H

    # Слайд 1: Титульный
    slide_title(prs)

    # Слайд 2: Шаг 1 — Очистить
    slide_step(prs, 1, 'ОЧИСТИТЬ',
               ['Удалите рыхлую ржавчину и старую',
                'краску металлической щёткой или',
                'наждачной бумагой (Р40-Р80).'],
               'Плотная ржавчина — можно оставить. TITAN красит по ржавчине.',
               pictogram_clean)

    # Слайд 3: Шаг 2 — Обезжирить
    slide_step(prs, 2, 'ОБЕЗЖИРИТЬ',
               ['Протрите поверхность тряпкой с',
                'растворителем (ацетон, ксилол).',
                'Поверхность должна быть сухой.'],
               'Не используйте уайт-спирит. После обезжиривания не трогайте руками.',
               pictogram_degrease)

    # Слайд 4: Шаг 3 — Перемешать
    slide_step(prs, 3, 'ПЕРЕМЕШАТЬ',
               ['Открыть банку и тщательно перемешать',
                'до однородного цвета. Со дна вверх.',
                'GREEN:Готово к нанесению — грунтовать не нужно!'],
               'TITAN 3в1 = грунт + эмаль + преобразователь ржавчины в одной банке.',
               pictogram_stir)

    # Слайд 5: Шаг 4 — Покрасить
    slide_step(prs, 4, 'ПОКРАСИТЬ',
               ['Нанесите 2-3 слоя кистью, валиком',
                'или распылителем. Расход: 80-100 г/м².',
                'Между слоями: 5 часов при +20 °C.'],
               'Не красьте под прямым солнцем и при влажности выше 80%.',
               pictogram_paint,
               table_rows=[
                   ('Кисть / валик', '5 ч'),
                   ('Распылитель', '8 ч'),
                   ('Температура', '+5...+30 °C'),
               ])

    # Слайд 6: Шаг 5 — Готово
    slide_step(prs, 5, 'ГОТОВО! 10 ЛЕТ',
               ['Покрытие сохнет 5 часов. Не кантовать',
                'до полного высыхания. Эксплуатация —',
                'GREEN:через 7 дней. Срок службы: 10 лет.'],
               'Первые 7 дней покрытие набирает прочность — избегайте царапин.',
               pictogram_done)

    # Слайд 7: ТСО
    slide_tso(prs)

    os.makedirs(os.path.dirname(OUTPUT), exist_ok=True)
    prs.save(OUTPUT)
    print(f'OK: {OUTPUT}')
    print(f'Слайдов: {len(prs.slides)}')


if __name__ == '__main__':
    main()
