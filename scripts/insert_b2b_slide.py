#!/usr/bin/env python3
"""
PPT: вставка нового слайда между 9 и 10 в дорожную карту TITAN.
Слайд: ЭТАП 5+ — B2B-ИНСТРУМЕНТЫ ПРОДВИЖЕНИЯ
2 инструмента: листовка по TITAN + буклет B2B с пробником 0.4 кг
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Cm, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from copy import deepcopy
from lxml import etree
import os

SRC = '/home/z/my-project/upload/Titan -Дорожная_карта_fin_16.07.2026.pptx'
OUT = '/home/z/my-project/download/Titan_Дорожная_карта_fin_16.07.2026_с_B2B.pptx'

# Цвета (как на слайдах 9, 10)
COL_GREEN = RGBColor(0x16, 0xA3, 0x4A)
COL_ORANGE = RGBColor(0xF4, 0x79, 0x20)
COL_RED = RGBColor(0xC6, 0x28, 0x28)
COL_BLUE = RGBColor(0x1E, 0x3A, 0x5F)
COL_YELLOW = RGBColor(0xFB, 0xBF, 0x24)
COL_BODY = RGBColor(0x1A, 0x1F, 0x36)
COL_GRAY = RGBColor(0x6B, 0x72, 0x80)
COL_WHITE = RGBColor(0xFF, 0xFF, 0xFF)
COL_LIGHT_GREEN = RGBColor(0xDC, 0xFC, 0xE7)


# Открываем исходную презентацию
prs = Presentation(SRC)

# Создаём новый слайд в конце (потом переместим)
blank_layout = prs.slide_layouts[6]
new_slide = prs.slides.add_slide(blank_layout)

# === Фон слайда (белый) ===
bg = new_slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
bg.fill.solid()
bg.fill.fore_color.rgb = COL_WHITE
bg.line.fill.background()
bg.shadow.inherit = False

# === Верхняя зелёная полоса (как на слайде 10) ===
top_bar = new_slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, Inches(0.7))
top_bar.fill.solid()
top_bar.fill.fore_color.rgb = COL_GREEN
top_bar.line.fill.background()
top_bar.shadow.inherit = False

# === Тонкая жёлтая полоса под зелёной ===
yellow_bar = new_slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, Inches(0.7), prs.slide_width, Inches(0.05))
yellow_bar.fill.solid()
yellow_bar.fill.fore_color.rgb = COL_YELLOW
yellow_bar.line.fill.background()
yellow_bar.shadow.inherit = False

# === Заголовок ===
title = new_slide.shapes.add_textbox(Inches(0.4), Inches(0.1), Inches(8.5), Inches(0.5))
tf = title.text_frame
p = tf.paragraphs[0]; p.alignment = PP_ALIGN.LEFT
r = p.add_run()
r.text = 'ЭТАП 5+ — B2B-ИНСТРУМЕНТЫ ПРОДВИЖЕНИЯ'
r.font.name = 'Arial'; r.font.size = Pt(22); r.font.bold = True; r.font.color.rgb = COL_WHITE

# === Подзаголовок справа ===
subtitle = new_slide.shapes.add_textbox(Inches(8.5), Inches(0.1), Inches(4.5), Inches(0.5))
tf = subtitle.text_frame
p = tf.paragraphs[0]; p.alignment = PP_ALIGN.RIGHT
r = p.add_run()
r.text = '2 инструмента · для кабинетов закупок ФС'
r.font.name = 'Arial'; r.font.size = Pt(13); r.font.color.rgb = COL_WHITE; r.font.italic = True


# === 2 КАРТОЧКИ ===
# Размеры как на слайде 10, но шире (по 6 inch каждая)
card_y = Inches(0.95)
card_w = Inches(6.2)
card_h = Inches(5.5)
gap = Inches(0.3)

cards = [
    {
        'x': Inches(0.4),
        'header_color': COL_ORANGE,
        'icon': '📄',
        'title': 'ЛИСТОВКА ПО TITAN',
        'subtitle': 'A4 · для ЛПР федеральных сетей',
        'rows': [
            ('Где:', 'Кабинеты закупок федеральных сетей (Леруа Мерлен, Максидом, ВсеИнструменты.ру, Лемана Про)'),
            ('Что:', 'Листовка А4 с УТП TITAN, расчётом маржинальности, кейсом «выкрас через год», топ-5 SKU'),
            ('Зачем:', 'ЛПР фед. сетей не ходят в торговый зал — листовка работает в кабинете закупок, где принимаются решения'),
            ('Тираж:', '500 экз. · 1 листовка = 1 ЛПР · точечная раздача на встречах'),
            ('Призыв:', '«Включите TITAN в MML — +15-20% маржи vs ПФ-115, −80% возвратов»'),
        ],
        'goal_title': 'ЦЕЛЬ',
        'goal_text': 'Убедить ЛПР федеральной сети включить TITAN в MML → запуск в 50-100 точках за квартал',
    },
    {
        'x': Inches(0.4) + card_w + gap,
        'header_color': COL_BLUE,
        'icon': '📘',
        'title': 'БУКЛЕТ B2B + ПРОБНИК 0.4 кг',
        'subtitle': 'A5 + баночка-пробник · для ЛПР и ТП',
        'rows': [
            ('Где:', 'Встречи с ЛПР дистрибьюторов + те же кабинеты ФС · в комплекте с КП'),
            ('Что:', 'Буклет А5 с полным портфелем TITAN (3в1, быстросохнущая, молотковая) + прикреплённая баночка 0.4 кг TITAN 3в1 чёрный RAL 9005'),
            ('Зачем:', 'ЛПР может ПОТРОГАТЬ продукт, открыть банку, оценить консистенцию — это сильнее любого КП на бумаге'),
            ('Тираж:', '300 комплектов · 1 буклет+пробник = 1 ЛПР · на ключевую встречу'),
            ('Призыв:', '«Откройте, потрогайте, попробуйте на пластине — и подпишем контракт на MML»'),
        ],
        'goal_title': 'ЦЕЛЬ',
        'goal_text': 'Конвертировать ЛПР из «рассматриваем» в «заключаем контракт» → +5-10 новых дистрибьюторов в Q3-Q4',
    },
]

for card in cards:
    # Фон карточки (белый)
    card_bg = new_slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, card['x'], card_y, card_w, card_h)
    card_bg.fill.solid(); card_bg.fill.fore_color.rgb = COL_WHITE
    card_bg.line.color.rgb = COL_GRAY; card_bg.line.width = Pt(0.5)
    card_bg.shadow.inherit = False
    
    # Шапка карточки (цветная)
    header_h = Inches(0.9)
    header = new_slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, card['x'], card_y, card_w, header_h)
    header.fill.solid(); header.fill.fore_color.rgb = card['header_color']
    header.line.fill.background()
    header.shadow.inherit = False
    
    # Иконка + заголовок
    title_box = new_slide.shapes.add_textbox(card['x'] + Inches(0.2), card_y + Inches(0.1), card_w - Inches(0.4), Inches(0.5))
    tf = title_box.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.LEFT
    r = p.add_run()
    r.text = f"{card['icon']}  {card['title']}"
    r.font.name = 'Arial'; r.font.size = Pt(15); r.font.bold = True; r.font.color.rgb = COL_WHITE
    
    # Подзаголовок карточки
    sub_box = new_slide.shapes.add_textbox(card['x'] + Inches(0.2), card_y + Inches(0.55), card_w - Inches(0.4), Inches(0.35))
    tf = sub_box.text_frame
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.LEFT
    r = p.add_run()
    r.text = card['subtitle']
    r.font.name = 'Arial'; r.font.size = Pt(11); r.font.italic = True; r.font.color.rgb = COL_WHITE
    
    # Строки с описанием
    row_y = card_y + header_h + Inches(0.1)
    for label, text in card['rows']:
        # Label
        lbl = new_slide.shapes.add_textbox(card['x'] + Inches(0.2), row_y, Inches(1.0), Inches(0.3))
        tf = lbl.text_frame
        p = tf.paragraphs[0]; p.alignment = PP_ALIGN.LEFT
        r = p.add_run()
        r.text = label
        r.font.name = 'Arial'; r.font.size = Pt(10); r.font.bold = True; r.font.color.rgb = card['header_color']
        
        # Text
        txt = new_slide.shapes.add_textbox(card['x'] + Inches(1.2), row_y, card_w - Inches(1.4), Inches(0.6))
        tf = txt.text_frame; tf.word_wrap = True
        p = tf.paragraphs[0]; p.alignment = PP_ALIGN.LEFT
        r = p.add_run()
        r.text = text
        r.font.name = 'Arial'; r.font.size = Pt(9.5); r.font.color.rgb = COL_BODY
        
        # Увеличиваем Y для следующей строки (примерно 0.6-0.8 inch в зависимости от длины текста)
        row_y += Inches(0.7) if len(text) < 80 else Inches(0.85)
    
    # Блок ЦЕЛЬ внизу карточки
    goal_y = card_y + card_h - Inches(1.1)
    goal_bar = new_slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, card['x'] + Inches(0.2), goal_y, card_w - Inches(0.4), Inches(0.3))
    goal_bar.fill.solid(); goal_bar.fill.fore_color.rgb = card['header_color']
    goal_bar.line.fill.background()
    goal_bar.shadow.inherit = False
    
    goal_title = new_slide.shapes.add_textbox(card['x'] + Inches(0.2), goal_y, card_w - Inches(0.4), Inches(0.3))
    tf = goal_title.text_frame
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = card['goal_title']
    r.font.name = 'Arial'; r.font.size = Pt(10); r.font.bold = True; r.font.color.rgb = COL_WHITE
    
    goal_text = new_slide.shapes.add_textbox(card['x'] + Inches(0.2), goal_y + Inches(0.35), card_w - Inches(0.4), Inches(0.7))
    tf = goal_text.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.LEFT
    r = p.add_run()
    r.text = card['goal_text']
    r.font.name = 'Arial'; r.font.size = Pt(10); r.font.color.rgb = COL_BODY; r.font.bold = True

# === ЖЁЛТАЯ ПЛАШКА С СИНЕРГИЕЙ ВНИЗУ ===
synergy_y = card_y + card_h + Inches(0.15)
synergy_bar = new_slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.4), synergy_y, Inches(12.5), Inches(0.55))
synergy_bar.fill.solid(); synergy_bar.fill.fore_color.rgb = COL_YELLOW
synergy_bar.line.fill.background()
synergy_bar.shadow.inherit = False

syn_text = new_slide.shapes.add_textbox(Inches(0.4), synergy_y, Inches(12.5), Inches(0.55))
tf = syn_text.text_frame; tf.word_wrap = True
tf.vertical_anchor = MSO_ANCHOR.MIDDLE
p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
r = p.add_run()
r.text = '⚡  Синергия: листовка (внимание ЛПР) → буклет с пробником (тест продукта) → контракт на MML → запуск TITAN в сети'
r.font.name = 'Arial'; r.font.size = Pt(11); r.font.bold = True; r.font.color.rgb = COL_BODY

# === Подпись внизу (как на др. слайдах) ===
footer = new_slide.shapes.add_textbox(Inches(0.4), Inches(7.25), Inches(12.5), Inches(0.2))
tf = footer.text_frame
p = tf.paragraphs[0]; p.alignment = PP_ALIGN.LEFT
r = p.add_run()
r.text = 'Empils  •  TITAN  •  Дорожная карта развития продаж'
r.font.name = 'Arial'; r.font.size = Pt(8); r.font.color.rgb = COL_GRAY; r.font.italic = True


# === ПЕРЕМЕЩАЕМ НОВЫЙ СЛАЙД НА ПОЗИЦИЮ 10 (между 9 и 10) ===
# В python-pptx слайды хранятся в XML, нужно переставить
# Получаем XML-дерево
sldIdLst = prs.slides._sldIdLst  # <p:sldIdLst>

# Новый слайд добавлен в конец — берём последний элемент
slides_list = list(sldIdLst)
new_slide_elem = slides_list[-1]
sldIdLst.remove(new_slide_elem)

# Вставляем после 9-го элемента (index 8, т.е. 9-й слайд)
# После удаления нового слайда, в списке осталось 19 элементов
# 9-й слайд = index 8
ninth_elem = list(sldIdLst)[8]
ninth_elem.addnext(new_slide_elem)

# Сохраняем
prs.save(OUT)
print(f'OK: {OUT}')
print(f'Размер: {os.path.getsize(OUT) / 1024:.1f} KB')
print(f'Слайдов: {len(prs.slides)}')

# Проверка порядка слайдов
prs2 = Presentation(OUT)
print('\n=== Новый порядок слайдов ===')
for i, slide in enumerate(prs2.slides, 1):
    texts = []
    for shape in slide.shapes:
        if shape.has_text_frame:
            t = shape.text_frame.text.strip()
            if t and len(t) > 3:
                texts.append(t)
                if len(texts) >= 1:
                    break
    label = texts[0][:80] if texts else '(пустой)'
    print(f"  Слайд {i}: {label}")
