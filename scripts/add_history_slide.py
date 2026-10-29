"""
Создаём новый слайд «История компании Эмпилс» в стиле примера Kudo.
- Горизонтальная временная линия с крупными годами
- Чередование блоков сверху/снизу
- Оранжевый акцент #F47629 (бренд-цвет Эмпилс)
- Шрифт Inter (Extra Bold для годов/заголовков, Regular для описаний)
- Лёгкие серые годы как «декоративные якоря», жирные оранжевые события
- Размер слайда: 20" × 11.25" (как в исходной презентации)
"""
import copy
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

SRC = '/home/z/my-project/upload/Презентация_для_новых_партнёров.pptx'
OUT = '/home/z/my-project/download/Презентация_для_новых_партнёров_с_историей.pptx'

# Цвета
ORANGE      = RGBColor(0xF4, 0x76, 0x29)   # акцент Эмпилс
DARK        = RGBColor(0x31, 0x31, 0x31)   # тёмный текст
GREY_LIGHT  = RGBColor(0xC0, 0xC0, 0xC0)   # серый для годов-якорей
GREY_TEXT   = RGBColor(0x4A, 0x4A, 0x4A)   # серый для описаний
LINE_GREY   = RGBColor(0xD0, 0xD0, 0xD0)   # серый для соединительных линий
WHITE       = RGBColor(0xFF, 0xFF, 0xFF)

FONT_BOLD = 'Inter Extra Bold'
FONT_REG  = 'Inter Regular'

p = Presentation(SRC)
slide_w, slide_h = p.slide_width, p.slide_height  # 20" × 11.25"

blank_layout = p.slide_layouts[4]  # Blank
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

def add_text(slide, left, top, width, height, text, font_size=11, bold=False,
             color=DARK, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, font_name=FONT_REG,
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

def add_multiline(slide, left, top, width, height, lines, font_size=10, color=GREY_TEXT,
                  font_name=FONT_REG, line_spacing=1.2):
    """lines: list of (text, bold, color) — каждая строка отдельным параграфом."""
    tb = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.02)
    tf.margin_right = Inches(0.02)
    tf.margin_top = Inches(0.02)
    tf.margin_bottom = Inches(0.02)
    for i, (text, bold, c) in enumerate(lines):
        if i == 0:
            para = tf.paragraphs[0]
        else:
            para = tf.add_paragraph()
        para.alignment = PP_ALIGN.LEFT
        para.line_spacing = line_spacing
        r = para.add_run()
        r.text = text
        r.font.size = Pt(font_size)
        r.font.bold = bold
        r.font.name = font_name
        r.font.color.rgb = c
    return tb

def add_line(slide, x1, y1, x2, y2, color, width_pt=1.0, dash=False):
    """Добавляет линию (соединительную)."""
    from pptx.enum.shapes import MSO_CONNECTOR
    line = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT,
                                       Inches(x1), Inches(y1),
                                       Inches(x2), Inches(y2))
    line.line.color.rgb = color
    line.line.width = Pt(width_pt)
    if dash:
        ln = line.line._get_or_add_ln()
        prstDash = ln.makeelement(qn('a:prstDash'), {'val': 'dash'})
        ln.append(prstDash)
    line.shadow.inherit = False
    return line

# ============================================================
# ФОН СЛАЙДА
# ============================================================
# Белый фон (по умолчанию) — оставляем как есть

# ============================================================
# ЗАГОЛОВОК СЛАЙДА
# ============================================================
# Серый заголовок «ИСТОРИЯ КОМПАНИИ «ЭМПИЛС»» (как на других слайдах)
add_text(slide, 0.55, 0.45, 12.0, 0.7,
         'ИСТОРИЯ КОМПАНИИ «ЭМПИЛС»',
         font_size=36, bold=True, color=RGBColor(0x65, 0x65, 0x65),
         font_name=FONT_BOLD, anchor=MSO_ANCHOR.MIDDLE)

# Подзаголовок — оранжевая «нитка» заголовка
add_text(slide, 0.55, 1.10, 18.0, 0.4,
         '124 года развития — от основания до современного производства полного цикла',
         font_size=14, bold=False, color=GREY_TEXT, italic=True,
         font_name=FONT_REG, anchor=MSO_ANCHOR.TOP)

# Тонкая оранжевая линия под заголовком
add_rect(slide, 0.55, 1.55, 19.0, 0.04, ORANGE)

# ============================================================
# ГОРИЗОНТАЛЬНАЯ ВРЕМЕННАЯ ЛИНИЯ
# ============================================================
# Линия по центру слайда, на Y = 5.6"
LINE_Y = 5.6
LINE_LEFT = 0.7
LINE_RIGHT = 19.3
LINE_WIDTH = LINE_RIGHT - LINE_LEFT

# Главная горизонтальная линия (оранжевая)
add_rect(slide, LINE_LEFT, LINE_Y - 0.025, LINE_WIDTH, 0.05, ORANGE)

# ============================================================
# КРУПНЫЕ СЕРЫЕ ГОДЫ КАК «ДЕКОРАТИВНЫЕ ЯКОРЯ»
# (в стиле Kudo — большие бледные цифры за линией)
# ============================================================
# Годы распределены равномерно по линии
# Позиции X (центр каждой «точки»): 7 точек
year_positions = [1.5, 4.5, 7.5, 10.5, 13.5, 16.5, 18.8]
year_labels = ['1901', '1930', '1960', '1990', '2007', '2015', '2026']

for x, year in zip(year_positions, year_labels):
    # Крупный серый год (фон)
    add_text(slide, x - 1.0, LINE_Y - 1.0, 2.0, 0.9,
             year, font_size=54, bold=True, color=GREY_LIGHT,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, font_name=FONT_BOLD)

# ============================================================
# ТОЧКИ-МАРКЕРЫ НА ЛИНИИ (оранжевые круги)
# ============================================================
for x in year_positions:
    # Круглый маркер на линии
    marker_size = 0.16
    sh = slide.shapes.add_shape(MSO_SHAPE.OVAL,
                                 Inches(x - marker_size/2),
                                 Inches(LINE_Y - marker_size/2),
                                 Inches(marker_size),
                                 Inches(marker_size))
    sh.fill.solid()
    sh.fill.fore_color.rgb = ORANGE
    sh.line.color.rgb = WHITE
    sh.line.width = Pt(2)
    sh.shadow.inherit = False

# ============================================================
# СОБЫТИЯ — ЧЕРЕДОВАНИЕ СВЕРХУ/СНИЗУ
# ============================================================
# Формат каждой записи: (x_position, is_top, [(title, events_list, year_color)])
# is_top=True — блок сверху от линии (заголовок близко к линии, описание выше)
# is_top=False — блок снизу (заголовок близко к линии, описание ниже)

events = [
    # 1901 — Основание (сверху)
    (1.5, True, 'Основание завода', [
        'Завод «Русская краска»',
        'в Ростове-на-Дону',
        'Начало производства',
        'органических ЛКМ',
    ]),
    # 1930 — Становление (снизу)
    (4.5, False, 'Индустриализация', [
        'Расширение производства',
        'Выпуск ЛКМ для',
        'промышленных объектов',
        'страны',
    ]),
    # 1960 — Модернизация (сверху)
    (7.5, True, 'Модернизация', [
        'Внедрение новых',
        'технологий варки',
        'лаков и эмалей',
        'Расширение ассортимента',
    ]),
    # 1990 — Приватизация (снизу)
    (10.5, False, 'Приватизация', [
        'Акционирование завода',
        'Сохранение команды',
        'и технологий',
        'Новый этап развития',
    ]),
    # 2007 — Новый этап (сверху)
    (13.5, True, 'Новый этап', [
        'Запуск брендов',
        'ОРЕОЛ, NEWTONE,',
        'AVANTGARDE',
        'Модернизация линий',
    ]),
    # 2015 — TITAN (снизу)
    (16.5, False, 'Запуск TITAN', [
        'Премиальная линейка',
        'по металлу —',
        'грунт-эмаль 3в1',
        'срок службы 10 лет',
    ]),
    # 2026 — Сегодня (сверху, справа)
    (18.8, True, 'Сегодня', [
        '39,3 тыс. тонн/год',
        'ISO 9001:2015',
        '8% рынка ЛКМ РФ',
        '124 года истории',
    ]),
]

# Координаты размещения блоков
# Блок сверху: заголовок на Y=4.4 (близко к линии), описание на Y=3.0
# Блок снизу: заголовок на Y=6.0 (близко к линии), описание на Y=6.7
BLOCK_WIDTH = 2.4
TOP_DESC_Y = 2.4
TOP_TITLE_Y = 4.6
BOT_TITLE_Y = 6.0
BOT_DESC_Y = 6.7

# Тонкая серая вертикальная линия-связка от точки к блоку
for x, is_top, title, _ in events:
    if is_top:
        # Вертикальная линия от точки (Y=LINE_Y) вверх к блоку заголовка
        add_line(slide, x, LINE_Y - 0.08, x, TOP_TITLE_Y + 0.2, LINE_GREY, width_pt=0.75, dash=True)
    else:
        add_line(slide, x, LINE_Y + 0.08, x, BOT_TITLE_Y, LINE_GREY, width_pt=0.75, dash=True)

for x, is_top, title, desc_lines in events:
    # Заголовок события (жирный, оранжевый)
    if is_top:
        add_text(slide, x - BLOCK_WIDTH/2, TOP_TITLE_Y, BLOCK_WIDTH, 0.5,
                 title, font_size=14, bold=True, color=ORANGE,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.TOP, font_name=FONT_BOLD)
        # Описание (над заголовком, если блок сверху)
        lines_data = [(line, False, GREY_TEXT) for line in desc_lines]
        add_multiline(slide, x - BLOCK_WIDTH/2, TOP_DESC_Y, BLOCK_WIDTH, 1.8,
                      lines_data, font_size=10, color=GREY_TEXT, line_spacing=1.15)
    else:
        # Заголовок сразу под линией
        add_text(slide, x - BLOCK_WIDTH/2, BOT_TITLE_Y, BLOCK_WIDTH, 0.5,
                 title, font_size=14, bold=True, color=ORANGE,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.TOP, font_name=FONT_BOLD)
        # Описание под заголовком
        lines_data = [(line, False, GREY_TEXT) for line in desc_lines]
        add_multiline(slide, x - BLOCK_WIDTH/2, BOT_DESC_Y, BLOCK_WIDTH, 1.8,
                      lines_data, font_size=10, color=GREY_TEXT, line_spacing=1.15)

# ============================================================
# ДЕКОР — ТОЧКИ-РАЗДЕЛИТЕЛИ ВДОЛЬ ЛИНИИ (опционально)
# ============================================================
# Маленькие точки-засечки на линии между маркерами годов
import numpy as np
minor_ticks = 30
for i in range(1, minor_ticks):
    t = i / minor_ticks
    x = LINE_LEFT + t * LINE_WIDTH
    # Пропускаем близко к маркерам
    too_close = any(abs(x - xp) < 0.4 for xp in year_positions)
    if too_close:
        continue
    # Маленькая серая засечка
    add_rect(slide, x - 0.015, LINE_Y - 0.015, 0.03, 0.03, LINE_GREY)

# ============================================================
# ФИНАЛЬНЫЙ БЛОК — ЦИФРА/ВЫВОД ВНИЗУ СПРАВА
# ============================================================
# «1901 → 2026» как декоративный элемент

# Подпись внизу слева
add_text(slide, 0.55, 10.2, 8.0, 0.5,
         '124 года непрерывного развития — от основания до современного производства полного цикла',
         font_size=11, bold=False, color=GREY_TEXT, italic=True,
         anchor=MSO_ANCHOR.MIDDLE)

# Большой оранжевый знак «124» в правом нижнем углу
add_text(slide, 17.0, 9.8, 2.5, 1.2,
         '124',
         font_size=64, bold=True, color=ORANGE,
         align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE, font_name=FONT_BOLD)
add_text(slide, 17.0, 10.7, 2.5, 0.4,
         'года истории',
         font_size=11, bold=False, color=GREY_TEXT,
         align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.TOP, font_name=FONT_REG)

# ---------- Сохранение ----------
p.save(OUT)
print(f'OK saved: {OUT}')
import os
print(f'Size: {os.path.getsize(OUT)} bytes')
print(f'Total slides: {len(p.slides)}')
