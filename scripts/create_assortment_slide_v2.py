#!/usr/bin/env python3
"""
Новый слайд: Ассортиментная матрица Эмпилс — все 11 брендов.
4+4+3 = 3 ряда карточек.
Стиль: как в презентации — оранжевый акцент F47629, Inter/Proxima Nova.
"""

from pptx import Presentation
from pptx.util import Pt, Cm
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
import os

SRC = '/home/z/my-project/upload/Презентация для потенциальных партнеров 04.08.2026.pptx'
OUT = '/home/z/my-project/download/Презентация_партнёрам_с_ассортиментом_v2.pptx'

# Цвета из презентации
COL_ORANGE = RGBColor(0xF4, 0x76, 0x29)
COL_DARK = RGBColor(0x31, 0x31, 0x31)
COL_BLACK = RGBColor(0x00, 0x00, 0x00)
COL_WHITE = RGBColor(0xFF, 0xFF, 0xFF)
COL_GRAY = RGBColor(0x6B, 0x72, 0x80)

# Цвета для брендов
BRAND_COLORS = {
    'AVANTGARDE':       RGBColor(0xE0, 0xE7, 0xFF),  # светло-фиолетовый
    'NEWTONE':          RGBColor(0xDB, 0xEA, 0xFE),  # голубой
    'ОРЕОЛ':            RGBColor(0xFE, 0xF9, 0xC3),  # жёлтый
    'TITAN':            RGBColor(0xFE, 0xE2, 0xE2),  # красный
    'PROFIWOOD':        RGBColor(0xDC, 0xFC, 0xE7),  # зелёный
    'РАСЦВЕТ':          RGBColor(0xFF, 0xED, 0xD5),  # оранжевый
    'ОПТИМА':           RGBColor(0xCC, 0xFB, 0xF1),  # бирюзовый
    'ОРЕОЛ Дисконт':    RGBColor(0xFE, 0xF3, 0xC7),  # светло-бежевый
    'ПРОСТОКРАШЕНО!':   RGBColor(0xF3, 0xF4, 0xF6),  # серый
    'EMPILS':           RGBColor(0xD1, 0xFA, 0xE5),  # тёмно-бирюзовый (промышл.)
    'СТРОЙКОМПЛЕКТ':    RGBColor(0xFE, 0xE5, 0xD3),  # коричневый
}

prs = Presentation(SRC)
slide_width = prs.slide_width
slide_height = prs.slide_height

slide = prs.slides.add_slide(prs.slide_layouts[0])

# === БЕЛЫЙ ФОН ===
bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, slide_width, slide_height)
bg.fill.solid(); bg.fill.fore_color.rgb = COL_WHITE
bg.line.fill.background()
bg.shadow.inherit = False

# === ЗАГОЛОВОК ===
title_box = slide.shapes.add_textbox(Cm(1.7), Cm(0.8), Cm(25), Cm(1.0))
tf = title_box.text_frame; tf.word_wrap = True
p = tf.paragraphs[0]; p.alignment = PP_ALIGN.LEFT
r = p.add_run()
r.text = "АССОРТИМЕНТНАЯ МАТРИЦА"
r.font.name = 'Proxima Nova Regular'; r.font.size = Pt(32); r.font.color.rgb = COL_DARK

# Оранжевое подчёркивание
ul = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Cm(1.7), Cm(1.8), Cm(5.5), Cm(0.1))
ul.fill.solid(); ul.fill.fore_color.rgb = COL_ORANGE
ul.line.fill.background()
ul.shadow.inherit = False

# Подзаголовок
sub = slide.shapes.add_textbox(Cm(1.7), Cm(2.0), Cm(30), Cm(0.6))
tf = sub.text_frame; tf.word_wrap = True
p = tf.paragraphs[0]; p.alignment = PP_ALIGN.LEFT
r = p.add_run()
r.text = "11 брендов · от эконом до премиум · полный спектр ЛКМ для любых поверхностей и задач"
r.font.name = 'Inter Regular'; r.font.size = Pt(12); r.font.color.rgb = COL_GRAY

# === 11 БРЕНДОВ ===
brands = [
    {
        'name': 'AVANTGARDE',
        'segment': 'ПРЕМИУМ',
        'desc': 'Гибридные технологии: экологичность водных + износостойкость органических',
        'categories': [
            'Краска интерьерная моющаяся (база A/C)',
            'Краска для фасадов (база A/C)',
            'Лак яхтный / паркетный',
        ],
        'surfaces': 'Интерьер · Фасад · Дерево',
        'color': BRAND_COLORS['AVANTGARDE'],
    },
    {
        'name': 'NEWTONE',
        'segment': 'СРЕДНИЙ+',
        'desc': 'Передовые технологии и современные материалы',
        'categories': [
            'Краска интерьерная / фасадная (база A/C)',
            'Грунтовка глубокого проникновения',
            'Шпатлёвка суперфинишная',
        ],
        'surfaces': 'Интерьер · Фасад',
        'color': BRAND_COLORS['NEWTONE'],
    },
    {
        'name': 'ОРЕОЛ',
        'segment': 'СРЕДНИЙ+',
        'desc': 'Краски, лаки, грунтовки — надёжный помощник в ремонте',
        'categories': [
            'Краска для стен и потолков моющаяся',
            'Краска для плит OSB / фасадная',
            'Грунтовка бетоноконтакт / шпатлёвка',
            'Клей, олифа, растворитель',
        ],
        'surfaces': 'Интерьер · Фасад · Дерево',
        'color': BRAND_COLORS['ОРЕОЛ'],
    },
    {
        'name': 'TITAN',
        'segment': 'СРЕДНИЙ+',
        'desc': 'Грунт-эмаль 3в1 — грунт + эмаль + преобразователь ржавчины',
        'categories': [
            '3в1 по ржавчине (10 цветов, RAL)',
            'Быстросохнущая матовая (7 цветов)',
            'Молотковая алкидно-стирольная (6 цветов)',
            '«Красишь 1 раз — служит 10 лет»',
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
            'Грунтовка ГФ-021 / растворители',
        ],
        'surfaces': 'Металл · Дерево',
        'color': BRAND_COLORS['РАСЦВЕТ'],
    },
    {
        'name': 'ОПТИМА',
        'segment': 'ЭКСПОРТ',
        'desc': 'Покрытия на органической основе для стран ближнего зарубежья',
        'categories': [
            'Эмали цветные (яркие оттенки)',
            'Грунтовка ГФ-021',
            'Грунт-эмали',
            'Оптимальное соотношение цены и качества',
        ],
        'surfaces': 'Металл · Дерево',
        'color': BRAND_COLORS['ОПТИМА'],
    },
    {
        'name': 'ОРЕОЛ Дисконт',
        'segment': 'ЭКОНОМ',
        'desc': 'Водно-дисперсионные покрытия по доступной цене',
        'categories': [
            'Краска для стен и потолков',
            'Шпатлёвка масляно-клеевая',
            'Грунтовка',
            '«Покрасил и забыл!»',
        ],
        'surfaces': 'Интерьер',
        'color': BRAND_COLORS['ОРЕОЛ Дисконт'],
    },
    {
        'name': 'ПРОСТОКРАШЕНО!',
        'segment': 'ЭКОНОМ',
        'desc': 'Базовые ЛКМ для массового потребителя',
        'categories': [
            'Эмаль ПФ-115 (БАУЦЕНТР)',
            'Краска водоэмульсионная',
            'Грунтовка / олифа / лак',
            'Растворитель',
        ],
        'surfaces': 'Металл · Дерево · Интерьер',
        'color': BRAND_COLORS['ПРОСТОКРАШЕНО!'],
    },
    {
        'name': 'EMPILS',
        'segment': 'ПРОМЫШЛЕННЫЕ',
        'desc': 'Покрытия промышленного назначения — для крупных предприятий',
        'categories': [
            'Огнебиозащитные составы',
            'Защитно-декоративные покрытия',
            'Краски-базы под колеровку',
            'Партнёр: «Ростсельмаш»',
        ],
        'surfaces': 'Металл · Дерево · Промышл.',
        'color': BRAND_COLORS['EMPILS'],
    },
    {
        'name': 'СТРОЙКОМПЛЕКТ',
        'segment': 'СТРОИТЕЛЬНАЯ ФАСОВКА',
        'desc': 'Крупная фасовка для строительных объектов большой площади',
        'categories': [
            'Эмаль ПФ-115 (крупная фасовка)',
            'Краска интерьерная / фасадная',
            'Универсальность · оптимальная цена',
            'Яркая цветовая палитра',
        ],
        'surfaces': 'Металл · Дерево · Интерьер',
        'color': BRAND_COLORS['СТРОЙКОМПЛЕКТ'],
    },
]

# === РАЗМЕЩЕНИЕ: 4+4+3 = 3 ряда ===
card_w = Cm(11.3)
card_h = Cm(7.3)
gap_x = Cm(0.5)
gap_y = Cm(0.4)
start_x = Cm(1.5)
start_y = Cm(2.9)

for i, brand in enumerate(brands):
    # Определяем ряд и позицию в ряду
    if i < 4:
        row, col = 0, i
    elif i < 8:
        row, col = 1, i - 4
    else:
        row, col = 2, i - 8
    
    x = start_x + col * (card_w + gap_x)
    y = start_y + row * (card_h + gap_y)
    
    # Для последнего ряда (3 карточки) — центрируем
    if row == 2:
        x = start_x + Cm(1.75) + col * (card_w + gap_x)
    
    # Карточка
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, card_w, card_h)
    card.fill.solid(); card.fill.fore_color.rgb = brand['color']
    card.line.color.rgb = COL_ORANGE; card.line.width = Pt(1)
    card.shadow.inherit = False
    
    # Название бренда
    name_box = slide.shapes.add_textbox(x + Cm(0.3), y + Cm(0.15), card_w - Cm(0.6), Cm(0.8))
    tf = name_box.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.LEFT
    r = p.add_run()
    r.text = brand['name']
    r.font.name = 'Proxima Nova Regular'; r.font.size = Pt(15); r.font.bold = True; r.font.color.rgb = COL_DARK
    
    # Ценовой сегмент
    p2 = tf.add_paragraph(); p2.alignment = PP_ALIGN.LEFT
    r = p2.add_run()
    r.text = brand['segment']
    r.font.name = 'Inter Regular'; r.font.size = Pt(8); r.font.color.rgb = COL_ORANGE
    r.font.bold = True
    
    # Оранжевое подчёркивание
    ul = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x + Cm(0.3), y + Cm(1.2), Cm(3.0), Cm(0.05))
    ul.fill.solid(); ul.fill.fore_color.rgb = COL_ORANGE
    ul.line.fill.background()
    ul.shadow.inherit = False
    
    # Описание
    desc_box = slide.shapes.add_textbox(x + Cm(0.3), y + Cm(1.4), card_w - Cm(0.6), Cm(1.0))
    tf = desc_box.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.LEFT
    r = p.add_run()
    r.text = brand['desc']
    r.font.name = 'Inter Regular'; r.font.size = Pt(8); r.font.color.rgb = COL_BLACK
    
    # Категории
    cat_box = slide.shapes.add_textbox(x + Cm(0.3), y + Cm(2.6), card_w - Cm(0.6), Cm(3.5))
    tf = cat_box.text_frame; tf.word_wrap = True
    for ci, cat in enumerate(brand['categories']):
        if ci == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        p.space_before = Pt(1)
        p.space_after = Pt(1)
        r = p.add_run()
        r.text = f"•  {cat}"
        r.font.name = 'Inter Regular'; r.font.size = Pt(8); r.font.color.rgb = COL_BLACK
    
    # Поверхности
    surf_box = slide.shapes.add_textbox(x + Cm(0.3), y + Cm(6.5), card_w - Cm(0.6), Cm(0.6))
    tf = surf_box.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.LEFT
    r = p.add_run()
    r.text = brand['surfaces']
    r.font.name = 'Inter Regular'; r.font.size = Pt(7); r.font.italic = True; r.font.color.rgb = COL_GRAY

# === НИЖНЯЯ ИНФО-ПОЛОСА ===
footer_y = Cm(26.8)
footer = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, footer_y, slide_width, Cm(1.5))
footer.fill.solid(); footer.fill.fore_color.rgb = COL_DARK
footer.line.fill.background()
footer.shadow.inherit = False

footer_text = slide.shapes.add_textbox(Cm(1.5), footer_y + Cm(0.2), Cm(40), Cm(1.1))
tf = footer_text.text_frame; tf.word_wrap = True
tf.vertical_anchor = MSO_ANCHOR.MIDDLE
p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
r = p.add_run()
r.text = "11 брендов покрывают весь спектр: от эконом до премиум · "
r.font.name = 'Inter Regular'; r.font.size = Pt(10); r.font.color.rgb = COL_WHITE
r2 = p.add_run()
r2.text = "металл · дерево · интерьер · фасад · промышленные · экспорт"
r2.font.name = 'Inter Regular'; r2.font.size = Pt(10); r2.font.bold = True; r2.font.color.rgb = COL_ORANGE

# === Перемещаем слайд на позицию 11 ===
sldIdLst = prs.slides._sldIdLst
slides_list = list(sldIdLst)
new_slide_elem = slides_list[-1]
sldIdLst.remove(new_slide_elem)
tenth_elem = list(sldIdLst)[9]
tenth_elem.addnext(new_slide_elem)

prs.save(OUT)
print(f'OK: {OUT}')
print(f'Размер: {os.path.getsize(OUT) / 1024:.1f} KB')
print(f'Слайдов: {len(prs.slides)}')

# Проверка
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
