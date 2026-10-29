"""
Создаём новый слайд «СОБСТВЕННАЯ ТОРГОВАЯ МАРКА (СТМ)» в презентации для новых партнёров.
Стиль: цепочка из 5 этапов с кругами, разделёнными пунктирными стрелками.
Адаптация примера под бренд-цвета Эмпилс: оранжевый #F47629 + тёмно-серый фон.
5 этапов: 1) Бриф и дизайн · 2) Разработка рецептуры · 3) Лабораторные испытания ·
4) Производство · 5) Упаковка и отгрузка.
"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

SRC = '/home/z/my-project/download/Презентация_для_новых_партнёров_с_историей.pptx'
OUT = '/home/z/my-project/download/Презентация_для_новых_партнёров_с_историей_и_СТМ.pptx'

# ---------- Цвета ----------
ORANGE      = RGBColor(0xF4, 0x76, 0x29)   # акцент Эмпилс
ORANGE_LT   = RGBColor(0xFF, 0xC1, 0x80)   # светлый оранжевый для пассивной части кольца
DARK_BG     = RGBColor(0x4A, 0x4A, 0x4A)   # тёмно-серый фон рабочей зоны
GREY_TEXT   = RGBColor(0x6B, 0x72, 0x80)   # серый текст
GREY_DARK   = RGBColor(0x31, 0x31, 0x31)   # тёмно-серый текст
GREY_LIGHT  = RGBColor(0xC0, 0xC0, 0xC0)   # светло-серый
WHITE       = RGBColor(0xFF, 0xFF, 0xFF)
GREY_RING   = RGBColor(0xD8, 0xD8, 0xD8)  # серое пассивное кольцо

FONT_BOLD = 'Inter Extra Bold'
FONT_REG  = 'Inter Regular'

p = Presentation(SRC)
slide_w, slide_h = p.slide_width, p.slide_height  # 20" × 11.25"

blank_layout = p.slide_layouts[4]
slide = p.slides.add_slide(blank_layout)

# ---------- утилиты ----------
def add_rect(slide, left, top, width, height, fill_color, line_color=None):
    sh = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(left), Inches(top),
                                 Inches(width), Inches(height))
    sh.fill.solid()
    sh.fill.fore_color.rgb = fill_color
    if line_color is None:
        sh.line.fill.background()
    else:
        sh.line.color.rgb = line_color
        sh.line.width = Pt(0.5)
    sh.shadow.inherit = False
    return sh

def add_rounded_rect(slide, left, top, width, height, fill_color, line_color=None, radius=0.1):
    sh = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                 Inches(left), Inches(top),
                                 Inches(width), Inches(height))
    sh.fill.solid()
    sh.fill.fore_color.rgb = fill_color
    if line_color is None:
        sh.line.fill.background()
    else:
        sh.line.color.rgb = line_color
        sh.line.width = Pt(0.5)
    # Установка радиуса скругления
    try:
        sh.adjustments[0] = radius
    except Exception:
        pass
    sh.shadow.inherit = False
    return sh

def add_text(slide, left, top, width, height, text, font_size=11, bold=False,
             color=DARK_BG, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, font_name=FONT_REG,
             italic=False):
    tb = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.02)
    tf.margin_right = Inches(0.02)
    tf.margin_top = Inches(0.02)
    tf.margin_bottom = Inches(0.02)
    tf.vertical_anchor = anchor
    para = tf.paragraphs[0]
    para.alignment = align
    r = para.add_run()
    r.text = text
    r.font.size = Pt(font_size)
    r.font.bold = bold
    r.font.italic = italic
    r.font.name = font_name
    r.font.color.rgb = color
    return tb

def add_oval(slide, left, top, width, height, fill_color, line_color=None, line_width_pt=0.5):
    sh = slide.shapes.add_shape(MSO_SHAPE.OVAL,
                                Inches(left), Inches(top),
                                Inches(width), Inches(height))
    sh.fill.solid()
    sh.fill.fore_color.rgb = fill_color
    if line_color is None:
        sh.line.fill.background()
    else:
        sh.line.color.rgb = line_color
        sh.line.width = Pt(line_width_pt)
    sh.shadow.inherit = False
    return sh

def add_arc_segment(slide, cx, cy, radius, start_angle, end_angle, color, thickness=0.25):
    """
    Добавляет дуговое кольцо (сегмент окружности) с заданными углами.
    Используется для создания эффекта «заполненной части кольца».
    """
    # В python-pptx нет прямой поддержки дуг, поэтому рисуем через BLOCK_ARC
    # или через OVAL с обрезкой. Простейший путь — нарисовать овал и поверх него
    # частичный блок другого цвета. Для простоты используем OVAL частично перекрытый.
    pass

def add_dashed_arrow(slide, x1, y1, x2, y2, color, width_pt=1.5):
    """Пунктирная стрелка между двумя точками."""
    from pptx.enum.shapes import MSO_CONNECTOR
    line = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT,
                                       Inches(x1), Inches(y1),
                                       Inches(x2), Inches(y2))
    line.line.color.rgb = color
    line.line.width = Pt(width_pt)
    # Делаем пунктир
    ln = line.line._get_or_add_ln()
    prstDash = ln.makeelement(qn('a:prstDash'), {'val': 'dash'})
    ln.append(prstDash)
    # Добавляем стрелку на конце
    tail = ln.makeelement(qn('a:tailEnd'), {'type': 'triangle', 'w': 'med', 'len': 'med'})
    ln.append(tail)
    line.shadow.inherit = False
    return line

# ============================================================
# ФОН И ВЕРХНЯЯ ПОЛОСА
# ============================================================
# Фон — белый (по умолчанию)

# Заголовок слайда (как на других слайдах — серый)
add_text(slide, 0.55, 0.45, 14.0, 0.7,
         'СОБСТВЕННАЯ ТОРГОВАЯ МАРКА (СТМ)',
         font_size=36, bold=True, color=RGBColor(0x65, 0x65, 0x65),
         font_name=FONT_BOLD, anchor=MSO_ANCHOR.MIDDLE)

# Подзаголовок — оранжевый акцент
add_text(slide, 0.55, 1.15, 18.5, 0.5,
         'Изготовим под вашей торговой маркой специализированный продукт '
         'в необходимой вам таре и объёме',
         font_size=16, bold=True, color=ORANGE,
         font_name=FONT_BOLD, anchor=MSO_ANCHOR.TOP)

# Тонкая оранжевая линия-разделитель
add_rect(slide, 0.55, 1.75, 19.0, 0.04, ORANGE)

# ============================================================
# ОПИСАНИЕ (краткое, под линией)
# ============================================================
add_text(slide, 0.55, 1.95, 18.5, 0.5,
         'У нас есть полный спектр компетенций — от разработки рецептуры до '
         'контроля качества готовой партии.',
         font_size=13, bold=False, color=GREY_TEXT,
         font_name=FONT_REG, anchor=MSO_ANCHOR.TOP, italic=True)

# ============================================================
# ТЁМНО-СЕРАЯ РАБОЧАЯ ЗОНА С 5 КРУГАМИ
# ============================================================
# Большой тёмно-серый прямоугольник-фон для цепочки этапов
BG_TOP = 2.7
BG_HEIGHT = 5.2
add_rect(slide, 0.55, BG_TOP, 18.9, BG_HEIGHT, DARK_BG)

# Параметры 5 кругов
NUM_CIRCLES = 5
CIRCLE_DIAM = 1.4   # диаметр каждого круга
CIRCLE_Y = BG_TOP + (BG_HEIGHT - CIRCLE_DIAM) / 2 - 0.3
LABEL_Y = CIRCLE_Y - 0.7

# Горизонтальные позиции ЦЕНТРОВ кругов — равномерно распределены
# Учитываем, что круги имеют диаметр CIRCLE_DIAM и должны помещаться в зоне [0.8, 19.2]
LEFT_CENTER = 0.8 + CIRCLE_DIAM/2    # центр первого круга
RIGHT_CENTER = 19.2 - CIRCLE_DIAM/2  # центр последнего круга
TOTAL_WIDTH = RIGHT_CENTER - LEFT_CENTER
STEP = TOTAL_WIDTH / (NUM_CIRCLES - 1)
circle_centers_x = [LEFT_CENTER + i * STEP for i in range(NUM_CIRCLES)]
# Левый край каждого круга (для add_oval)
circle_x_positions = [cx - CIRCLE_DIAM/2 for cx in circle_centers_x]

# Заголовок цепочки внутри серой зоны (белый, над кругами)
add_text(slide, 0.55, BG_TOP + 0.2, 18.9, 0.4,
         'ПОЛНЫЙ ЦИКЛ РАЗРАБОТКИ — 5 ЭТАПОВ',
         font_size=14, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, font_name=FONT_BOLD)

# Подписи этапов над кругами
stage_labels = [
    'Бриф и дизайн',
    'Разработка\nрецептуры',
    'Лабораторные\nиспытания',
    'Производство\nпродукта',
    'Упаковка и\nотгрузка',
]

# Этапы — номера над кругами (оранжевые)
LABEL_WIDTH = 2.4   # ширина блока подписи (больше диаметра круга, чтобы текст помещался)
for i, x in enumerate(circle_x_positions):
    # x — это левый край круга. Подпись центрируем по центру круга.
    label_left = circle_centers_x[i] - LABEL_WIDTH/2
    # Номер этапа (оранжевый, над подписью)
    add_text(slide, label_left, BG_TOP + 0.7, LABEL_WIDTH, 0.4,
             f'ЭТАП {i+1}',
             font_size=10, bold=True, color=ORANGE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, font_name=FONT_BOLD)
    # Название этапа (белый, над кругом)
    label_lines = stage_labels[i].split('\n')
    for j, line in enumerate(label_lines):
        add_text(slide, label_left, BG_TOP + 1.15 + j * 0.35, LABEL_WIDTH, 0.35,
                 line,
                 font_size=11, bold=False, color=WHITE,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, font_name=FONT_REG)

# Сами круги — рисуем на Y = CIRCLE_Y
CIRCLE_TOP = BG_TOP + 2.05
CIRCLE_LEFT_BASE = circle_x_positions[0]

for i, x in enumerate(circle_x_positions):
    # 1) Внешний круг-обводка (серый, как пассивная часть кольца)
    add_oval(slide, x, CIRCLE_TOP, CIRCLE_DIAM, CIRCLE_DIAM,
             fill_color=GREY_RING, line_color=None)
    # 2) Внутренний тёмный круг (для эффекта «кольца с сегментом»)
    inner_padding = 0.18
    inner_size = CIRCLE_DIAM - 2 * inner_padding
    add_oval(slide, x + inner_padding, CIRCLE_TOP + inner_padding,
             inner_size, inner_size,
             fill_color=DARK_BG, line_color=None)
    # 3) Внутренний белый круг с цифрой (по центру)
    inner_white_padding = 0.32
    inner_white_size = CIRCLE_DIAM - 2 * inner_white_padding
    add_oval(slide, x + inner_white_padding, CIRCLE_TOP + inner_white_padding,
             inner_white_size, inner_white_size,
             fill_color=WHITE, line_color=None)
    # 4) Номер этапа в центре круга (оранжевый, жирный)
    add_text(slide, x, CIRCLE_TOP, CIRCLE_DIAM, CIRCLE_DIAM,
             str(i+1),
             font_size=54, bold=True, color=ORANGE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, font_name=FONT_BOLD)

# Пунктирные стрелки между кругами
for i in range(NUM_CIRCLES - 1):
    x1 = circle_x_positions[i] + CIRCLE_DIAM + 0.05
    x2 = circle_x_positions[i+1] - 0.05
    y = CIRCLE_TOP + CIRCLE_DIAM / 2
    add_dashed_arrow(slide, x1, y, x2, y, ORANGE_LT, width_pt=2.0)

# ============================================================
# НИЖНИЙ БЛОК — ИТОГОВЫЙ С ВЫВОДОМ
# ============================================================
# Белый прямоугольник со скруглёнными углами, частично наложенный на серую зону
BOTTOM_TOP = BG_TOP + BG_HEIGHT - 0.5   # частично перекрывает серую зону
BOTTOM_HEIGHT = 0.95
BOTTOM_LEFT = 3.0
BOTTOM_WIDTH = 14.0
add_rounded_rect(slide, BOTTOM_LEFT, BOTTOM_TOP, BOTTOM_WIDTH, BOTTOM_HEIGHT,
                 fill_color=WHITE, line_color=ORANGE_LT, radius=0.15)

# Зелёная галочка-кружок слева от текста
CHECK_SIZE = 0.55
CHECK_LEFT = BOTTOM_LEFT + 0.3
CHECK_TOP = BOTTOM_TOP + (BOTTOM_HEIGHT - CHECK_SIZE) / 2
add_oval(slide, CHECK_LEFT, CHECK_TOP, CHECK_SIZE, CHECK_SIZE,
         fill_color=ORANGE, line_color=None)
# Белая галочка внутри
add_text(slide, CHECK_LEFT, CHECK_TOP, CHECK_SIZE, CHECK_SIZE,
         '✓',
         font_size=22, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, font_name=FONT_BOLD)

# Текст справа от галочки
add_text(slide, CHECK_LEFT + CHECK_SIZE + 0.3, BOTTOM_TOP,
         BOTTOM_WIDTH - CHECK_SIZE - 0.6, BOTTOM_HEIGHT,
         'Сопровождается контролем качества на каждом этапе',
         font_size=18, bold=True, color=GREY_DARK,
         align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE, font_name=FONT_BOLD)

# ============================================================
# БЛОКИ С КОМПЕТЕНЦИЯМИ (под серой зоной)
# ============================================================
COMP_TOP = BG_TOP + BG_HEIGHT + 0.7   # 8.6
COMP_HEIGHT = 1.8
COMP_LEFT = 0.55
COMP_WIDTH = 18.9
# Раздел: «Наши компетенции»
add_text(slide, COMP_LEFT, COMP_TOP, COMP_WIDTH, 0.5,
         'НАШИ КОМПЕТЕНЦИИ',
         font_size=14, bold=True, color=ORANGE,
         font_name=FONT_BOLD, anchor=MSO_ANCHOR.TOP)

# Сетка компетенций 3 × 2 (6 ячеек)
competences = [
    ('Разработка рецептуры',     'Собственная лаборатория, 2 НИЛ\nпод любые типы ЛКМ'),
    ('Дизайн упаковки',            'Брендирование этикетки\nи тары под СТМ'),
    ('Все виды тары',              '0,4 / 0,75 / 2 / 5 / 8,5 л\nбанки, аэрозоли, бочки'),
    ('Цвет по каталогу RAL',       'Колеровка в любой цвет\nRAL, NCS, собственный'),
    ('Сертификация и ТР ТС',       'Все необходимые документы\nвключая СТБ, ГОСТ, ISO'),
    ('Контроль качества партии',   'Каждая партия проходит\nвыходной контроль в лаборатории'),
]

cell_w = COMP_WIDTH / 3
cell_h = 0.85
for i, (title, desc) in enumerate(competences):
    row = i // 3
    col = i % 3
    x = COMP_LEFT + col * cell_w
    y = COMP_TOP + 0.5 + row * cell_h
    # Заголовок компетенции
    add_text(slide, x + 0.1, y, cell_w - 0.2, 0.3,
             title,
             font_size=11, bold=True, color=GREY_DARK,
             align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, font_name=FONT_BOLD)
    # Описание
    desc_lines = desc.split('\n')
    for j, line in enumerate(desc_lines):
        add_text(slide, x + 0.1, y + 0.32 + j * 0.22, cell_w - 0.2, 0.25,
                 line,
                 font_size=10, bold=False, color=GREY_TEXT,
                 align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, font_name=FONT_REG)

# ============================================================
# ФИНАЛЬНЫЙ БЛОК — ПРИЗЫВ К ДЕЙСТВИЮ ВНИЗУ
# ============================================================
add_text(slide, 0.55, 10.7, 18.9, 0.4,
         'Готовы обсудить ваш проект СТМ? Свяжитесь с региональным менеджером Эмпилс.',
         font_size=12, bold=True, color=ORANGE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, font_name=FONT_BOLD, italic=True)

# ---------- Сохранение ----------
p.save(OUT)
print(f'OK saved: {OUT}')
import os
print(f'Size: {os.path.getsize(OUT)} bytes')
print(f'Total slides: {len(p.slides)}')
