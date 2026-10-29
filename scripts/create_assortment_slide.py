#!/usr/bin/env python3
"""
Новый слайд: Ассортиментная матрица Эмпилс.
Стиль: как в презентации — оранжевый акцент F47629, шрифт Inter/Proxima Nova,
тёмно-серый текст 313131, белый фон, оранжевые подчёркивания.
Вставляется в презентацию после слайда 10 (ценовое позиционирование).
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Cm, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
import os

SRC = '/home/z/my-project/upload/Презентация для потенциальных партнеров 04.08.2026.pptx'
OUT = '/home/z/my-project/download/Презентация_партнёрам_с_ассортиментом.pptx'

# Цвета из презентации
COL_ORANGE = RGBColor(0xF4, 0x76, 0x29)     # F47629 — акцент
COL_ORANGE_LIGHT = RGBColor(0xF1, 0x83, 0x24) # F18324 — подчёркивания
COL_DARK = RGBColor(0x31, 0x31, 0x31)        # 313131 — заголовки
COL_BLACK = RGBColor(0x00, 0x00, 0x00)       # 000000 — текст
COL_WHITE = RGBColor(0xFF, 0xFF, 0xFF)
COL_GRAY = RGBColor(0x6B, 0x72, 0x80)
COL_LIGHT_BG = RGBColor(0xF8, 0xF8, 0xF8)
COL_BLUE = RGBColor(0x1E, 0x40, 0xAF)
COL_GREEN = RGBColor(0x16, 0xA3, 0x4A)
COL_RED = RGBColor(0xC6, 0x28, 0x28)

# Цвета для брендов (пастельные)
BRAND_COLORS = {
    'AVANTGARDE': RGBColor(0xE0, 0xE7, 0xFF),   # светло-голубой
    'NEWTONE': RGBColor(0xDB, 0xEA, 0xFE),        # голубой
    'ОРЕОЛ': RGBColor(0xFE, 0xF9, 0xC3),          # светло-жёлтый
    'TITAN': RGBColor(0xFE, 0xE2, 0xE2),          # светло-красный
    'PROFIWOOD': RGBColor(0xDC, 0xFC, 0xE7),      # светло-зелёный
    'РАСЦВЕТ': RGBColor(0xFF, 0xED, 0xD5),        # светло-оранжевый
    'ПРОСТОКРАШЕНО!': RGBColor(0xF3, 0xF4, 0xF6), # светло-серый
}

# Открываем исходную презентацию
prs = Presentation(SRC)
slide_width = prs.slide_width
slide_height = prs.slide_height

# Создаём новый слайд
slide = prs.slides.add_slide(prs.slide_layouts[0])  # DEFAULT layout

# === БЕЛЫЙ ФОН ===
bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, slide_width, slide_height)
bg.fill.solid(); bg.fill.fore_color.rgb = COL_WHITE
bg.line.fill.background()
bg.shadow.inherit = False

# === ЗАГОЛОВОК ===
title_box = slide.shapes.add_textbox(Cm(1.7), Cm(1.2), Cm(25), Cm(1.2))
tf = title_box.text_frame; tf.word_wrap = True
p = tf.paragraphs[0]; p.alignment = PP_ALIGN.LEFT
r = p.add_run()
r.text = "АССОРТИМЕНТНАЯ МАТРИЦА"
r.font.name = 'Proxima Nova Regular'; r.font.size = Pt(36); r.font.color.rgb = COL_DARK

# Оранжевое подчёркивание
underline = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Cm(1.7), Cm(2.4), Cm(5.5), Cm(0.12))
underline.fill.solid(); underline.fill.fore_color.rgb = COL_ORANGE
underline.line.fill.background()
underline.shadow.inherit = False

# Подзаголовок
sub = slide.shapes.add_textbox(Cm(1.7), Cm(2.7), Cm(25), Cm(0.8))
tf = sub.text_frame; tf.word_wrap = True
p = tf.paragraphs[0]; p.alignment = PP_ALIGN.LEFT
r = p.add_run()
r.text = "7 брендов · 5 ценовых сегментов · полный спектр ЛКМ для любых задач"
r.font.name = 'Inter Regular'; r.font.size = Pt(14); r.font.color.rgb = COL_GRAY

# === 7 БРЕНДОВ — карточки ===
# Размещение: 4 в верхнем ряду + 3 в нижнем
# Размеры: ширина ~11.5 см, высота ~8.5 см

brands = [
    {
        'name': 'AVANTGARDE',
        'segment': 'ПРЕМИУМ',
        'desc': 'Гибридные технологии: экологичность водных + износостойкость органических',
        'categories': [
            'Краска интерьерная моющаяся (база A/C)',
            'Краска для фасадов (база A/C)',
            'Лак яхтный высокой твёрдости',
            'Лак паркетный высокой твёрдости',
        ],
        'surfaces': 'Интерьер · Фасад · Дерево',
        'color': BRAND_COLORS['AVANTGARDE'],
    },
    {
        'name': 'NEWTONE',
        'segment': 'СРЕДНИЙ+',
        'desc': 'Передовые технологии и современные материалы',
        'categories': [
            'Краска интерьерная (база A/C)',
            'Краска фасадная (база A/C)',
            'Грунтовка глубокого проникновения',
            'Шпатлёвка суперфинишная',
        ],
        'surfaces': 'Интерьер · Фасад',
        'color': BRAND_COLORS['NEWTONE'],
    },
    {
        'name': 'ОРЕОЛ',
        'segment': 'СРЕДНИЙ+',
        'desc': 'Надёжный помощник в создании уюта — краски, лаки, грунтовки',
        'categories': [
            'Краска для стен и потолков моющаяся',
            'Краска интерьерная / фасадная',
            'Краска для плит OSB',
            'Грунтовка бетоноконтакт',
            'Шпатлёвка, клей, олифа, растворитель',
        ],
        'surfaces': 'Интерьер · Фасад · Дерево',
        'color': BRAND_COLORS['ОРЕОЛ'],
    },
    {
        'name': 'TITAN',
        'segment': 'СРЕДНИЙ+',
        'desc': 'Грунт-эмаль 3в1 по ржавчине — грунт + эмаль + преобразователь в одной банке',
        'categories': [
            'Грунт-эмаль 3в1 по ржавчине (10 цветов)',
            'Быстросохнущая матовая (7 цветов)',
            'Молотковая алкидно-стирольная (6 цветов)',
            'Слоган: «Красишь 1 раз — служит 10 лет»',
        ],
        'surfaces': 'Металл · Оцинковка · Дерево',
        'color': BRAND_COLORS['TITAN'],
    },
    {
        'name': 'PROFIWOOD',
        'segment': 'СРЕДНИЙ',
        'desc': 'Профессиональная защита и декор древесины',
        'categories': [
            'Антисептик кроющий / лаковый',
            'Защитно-декоративные покрытия',
            'Лак паркетный / по дереву и камню',
            'Огнебиозащитные составы',
        ],
        'surfaces': 'Дерево · Камень',
        'color': BRAND_COLORS['PROFIWOOD'],
    },
    {
        'name': 'РАСЦВЕТ',
        'segment': 'СРЕДНИЙ / ЭКОНОМ',
        'desc': 'Классические ЛКМ по ГОСТ — эмали, грунт-эмали, грунтовки',
        'categories': [
            'Эмаль ПФ-115 (ГОСТ)',
            'Грунт-эмаль по ржавчине',
            'Эмаль с молотковым эффектом',
            'Грунтовка ГФ-021 / быстросохнущая',
            'Растворители',
        ],
        'surfaces': 'Металл · Дерево',
        'color': BRAND_COLORS['РАСЦВЕТ'],
    },
    {
        'name': 'ПРОСТОКРАШЕНО!',
        'segment': 'ЭКОНОМ',
        'desc': 'Базовые ЛКМ для массового потребителя — просто и недорого',
        'categories': [
            'Эмаль ПФ-115 (БАУЦЕНТР)',
            'Краска водоэмульсионная',
            'Грунтовка / олифа',
            'Лак / растворитель',
        ],
        'surfaces': 'Металл · Дерево · Интерьер',
        'color': BRAND_COLORS['ПРОСТОКРАШЕНО!'],
    },
]

# Размещение карточек
# 4 в верхнем ряду, 3 в нижнем
card_w = Cm(11.2)
card_h = Cm(9.0)
gap_x = Cm(0.5)
gap_y = Cm(0.6)
start_x = Cm(1.5)
start_y = Cm(4.0)

for i, brand in enumerate(brands):
    if i < 4:
        # Верхний ряд — 4 карточки
        x = start_x + i * (card_w + gap_x)
        y = start_y
    else:
        # Нижний ряд — 3 карточки, центрируем
        x = start_x + Cm(0.5) + (i - 4) * (card_w + gap_x)
        y = start_y + card_h + gap_y
    
    # Карточка
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, card_w, card_h)
    card.fill.solid(); card.fill.fore_color.rgb = brand['color']
    card.line.color.rgb = COL_ORANGE; card.line.width = Pt(1)
    card.shadow.inherit = False
    
    # Название бренда
    name_box = slide.shapes.add_textbox(x + Cm(0.4), y + Cm(0.2), card_w - Cm(0.8), Cm(1.0))
    tf = name_box.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.LEFT
    r = p.add_run()
    r.text = brand['name']
    r.font.name = 'Proxima Nova Regular'; r.font.size = Pt(18); r.font.bold = True; r.font.color.rgb = COL_DARK
    
    # Ценовой сегмент
    p2 = tf.add_paragraph(); p2.alignment = PP_ALIGN.LEFT
    p2.space_before = Pt(0)
    r = p2.add_run()
    r.text = brand['segment']
    r.font.name = 'Inter Regular'; r.font.size = Pt(9); r.font.color.rgb = COL_ORANGE
    r.font.bold = True
    
    # Оранжевое подчёркивание
    ul = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x + Cm(0.4), y + Cm(1.5), Cm(3.5), Cm(0.06))
    ul.fill.solid(); ul.fill.fore_color.rgb = COL_ORANGE
    ul.line.fill.background()
    ul.shadow.inherit = False
    
    # Описание
    desc_box = slide.shapes.add_textbox(x + Cm(0.4), y + Cm(1.7), card_w - Cm(0.8), Cm(1.4))
    tf = desc_box.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.LEFT
    r = p.add_run()
    r.text = brand['desc']
    r.font.name = 'Inter Regular'; r.font.size = Pt(9); r.font.color.rgb = COL_BLACK
    
    # Категории
    cat_box = slide.shapes.add_textbox(x + Cm(0.4), y + Cm(3.3), card_w - Cm(0.8), Cm(4.5))
    tf = cat_box.text_frame; tf.word_wrap = True
    for ci, cat in enumerate(brand['categories']):
        if ci == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        p.space_before = Pt(2)
        p.space_after = Pt(2)
        r = p.add_run()
        r.text = f"•  {cat}"
        r.font.name = 'Inter Regular'; r.font.size = Pt(9); r.font.color.rgb = COL_BLACK
    
    # Поверхности
    surf_box = slide.shapes.add_textbox(x + Cm(0.4), y + Cm(8.0), card_w - Cm(0.8), Cm(0.8))
    tf = surf_box.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.LEFT
    r = p.add_run()
    r.text = brand['surfaces']
    r.font.name = 'Inter Regular'; r.font.size = Pt(8); r.font.italic = True; r.font.color.rgb = COL_GRAY

# === НИЖНЯЯ ИНФО-ПОЛОСА ===
footer_y = Cm(26.5)
footer = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, footer_y, slide_width, Cm(1.5))
footer.fill.solid(); footer.fill.fore_color.rgb = COL_DARK
footer.line.fill.background()
footer.shadow.inherit = False

footer_text = slide.shapes.add_textbox(Cm(1.5), footer_y + Cm(0.2), Cm(35), Cm(1.1))
tf = footer_text.text_frame; tf.word_wrap = True
tf.vertical_anchor = MSO_ANCHOR.MIDDLE
p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
r = p.add_run()
r.text = "7 брендов покрывают весь спектр потребностей: от эконом до премиум · "
r.font.name = 'Inter Regular'; r.font.size = Pt(11); r.font.color.rgb = COL_WHITE
r2 = p.add_run()
r2.text = "металл · дерево · интерьер · фасад · спецпродукты"
r2.font.name = 'Inter Regular'; r2.font.size = Pt(11); r2.font.bold = True; r2.font.color.rgb = COL_ORANGE

# === Перемещаем слайд на позицию 11 (после ценового позиционирования) ===
sldIdLst = prs.slides._sldIdLst
slides_list = list(sldIdLst)
# Новый слайд добавлен в конец (индекс -1)
new_slide_elem = slides_list[-1]
sldIdLst.remove(new_slide_elem)
# Вставляем после слайда 10 (индекс 9, т.е. 10-й элемент)
tenth_elem = list(sldIdLst)[9]
tenth_elem.addnext(new_slide_elem)

prs.save(OUT)
print(f'OK: {OUT}')
print(f'Размер: {os.path.getsize(OUT) / 1024:.1f} KB')
print(f'Слайдов: {len(prs.slides)}')

# Проверка порядка
prs2 = Presentation(OUT)
for i, s in enumerate(prs2.slides, 1):
    texts = []
    for shape in s.shapes:
        if shape.has_text_frame:
            t = shape.text_frame.text.strip()
            if t and len(t) > 3:
                texts.append(t[:60])
    label = ' | '.join(texts[:1])[:80] if texts else '(пустой)'
    print(f"  Слайд {i}: {label}")
