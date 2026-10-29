#!/usr/bin/env python3
"""
Создать слайд про работу с B2B-партнёрами по TITAN
и вставить его в презентацию Titan_Дорожная_карта_Лемана_v1.pptx.

Цели для B2B-партнёров (от пользователя):
1. Привлечение — доходность продукта дистрибьютора позволяет и мотивирует ли работать с b2b партерами
2. Пробный тест-образец (пробник) TITAN 0,4 глянец
3. Листовка, направленная на B2B
4. Высчитывание маржинальности в сторону смещения к крупной фасовке
5. Выявление регионального инфлюенсера, имеющего влияние в сфере B2B

Дополнительная информация (от меня):
6. Калькулятор маржинальности для дистрибьютора (Excel)
7. Отраслевые кейсы: B2B-объекты с расчётом выгоды
8. Прямой контакт Эмпилс с B2B-клиентом дистрибьютора
9. Тестовые образцы для демонстрации свойств на объекте

Главный вопрос, на который отвечает слайд:
"Что нам надо сделать, чтобы B2B-партнёры захотели работать с TITAN?"
"""

import os, shutil
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn

SRC = "/home/z/my-project/download/Titan_Дорожная_карта_Лемана_v1.pptx"
DST = "/home/z/my-project/download/Titan_Дорожная_карта_B2B_v1.pptx"
shutil.copy(SRC, DST)

prs = Presentation(DST)
print(f"Слайдов в исходной презентации: {len(prs.slides)}")

# Вставляем новый слайд после слайда 6 (Лемана) — позиция 7 (индекс 6)
slide_layout = prs.slide_layouts[6]  # Blank
new_slide = prs.slides.add_slide(slide_layout)

sldIdLst = prs.slides._sldIdLst
sldIds = list(sldIdLst)
new_sldId = sldIds[-1]
sldIdLst.remove(new_sldId)
sldIdLst.insert(6, new_sldId)  # позиция 7 (после Лемана)

slide = new_slide

# === Цвета фирменные ===
ORANGE = RGBColor(0xF4, 0x76, 0x29)
ORANGE_DARK = RGBColor(0xD5, 0x5E, 0x0A)
ORANGE_LIGHT = RGBColor(0xFB, 0xE6, 0xD5)
DARK = RGBColor(0x31, 0x31, 0x31)
GRAY = RGBColor(0x6B, 0x6B, 0x6B)
LIGHT_GRAY = RGBColor(0xF5, 0xF5, 0xF5)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
BG = RGBColor(0xFA, 0xFA, 0xFA)
BLUE = RGBColor(0x15, 0x65, 0xC0)
BLUE_DARK = RGBColor(0x0D, 0x47, 0xA1)
BLUE_LIGHT = RGBColor(0xE3, 0xF2, 0xFD)
GREEN = RGBColor(0x2E, 0x7D, 0x32)
GREEN_LIGHT = RGBColor(0xE8, 0xF5, 0xE9)

# === Данные по 7 инструментам работы с B2B-партнёрами ===
TOOLS = [
    {
        "num": "01",
        "title": "Доходность для дистрибьютора",
        "desc": "Расчёт маржинальности B2B-канала. Демонстрируем дистрибьютору, что работа с B2B-партнёрами выгодна — доходность выше, чем в рознице, при меньших операционных затратах",
        "tag": "МОТИВАЦИЯ",
        "color": ORANGE,
    },
    {
        "num": "02",
        "title": "Пробник TITAN 0,4 л (глянец)",
        "desc": "Тестовый образец для B2B-партнёра. Малая фасовка — низкий порог входа. Партнёр может попробовать продукт без больших затрат, оценить качество и принять решение о закупке",
        "tag": "ТЕСТ-ОБРАЗЕЦ",
        "color": ORANGE_DARK,
    },
    {
        "num": "03",
        "title": "Листовка для B2B-сегмента",
        "desc": "Отдельная листовка с расчётом выгоды TITAN для B2B-клиента: срок службы 10 лет, совокупная стоимость владения, экономия vs ПФ-115. Фокус на окупаемость, а не на цену банки",
        "tag": "ИНСТРУМЕНТ",
        "color": ORANGE,
    },
    {
        "num": "04",
        "title": "Смещение к крупной фасовке",
        "desc": "Расчёт маржинальности по фасовкам. B2B-клиенту выгоднее 2 л / 2,5 кг, чем 0,75 л. Дистрибьютор зарабатывает больше на крупной фасовке — мотивируем его продвигать именно её",
        "tag": "МАРЖА",
        "color": ORANGE_DARK,
    },
    {
        "num": "05",
        "title": "Региональный B2B-инфлюенсер",
        "desc": "Выявление регионального лидера мнений в сфере B2B (производитель металлоконструкций, владелец крупного цеха, руководитель строительной компании). Его рекомендация работает лучше любой рекламы",
        "tag": "ВЛИЯНИЕ",
        "color": ORANGE,
    },
    {
        "num": "06",
        "title": "Калькулятор маржинальности",
        "desc": "Excel-калькулятор для дистрибьютора: вводит объём закупки и фасовку → получает расчёт прибыли. Дистрибьютор сам видит, сколько он заработает на B2B-канале с TITAN",
        "tag": "ИНСТРУМЕНТ",
        "color": ORANGE_DARK,
    },
    {
        "num": "07",
        "title": "Прямой контакт с B2B-клиентом",
        "desc": "Эмпилс подключается к переговорам дистрибьютора с B2B-клиентом. Техническая экспертиза, демонстрация на объекте, расчёт ТСО. Дистрибьютор не остаётся один на один со сложным клиентом",
        "tag": "ПОДДЕРЖКА",
        "color": ORANGE,
    },
]

# === Вспомогательные функции ===

def add_rect(slide, left, top, width, height, fill_color, line_color=None, line_width=None):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    if line_color is None:
        shape.line.fill.background()
    else:
        shape.line.color.rgb = line_color
        if line_width:
            shape.line.width = line_width
    shape.shadow.inherit = False
    return shape

def add_round_rect(slide, left, top, width, height, fill_color, line_color=None, line_width=None):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    if line_color is None:
        shape.line.fill.background()
    else:
        shape.line.color.rgb = line_color
        if line_width:
            shape.line.width = line_width
    shape.shadow.inherit = False
    try:
        shape.adjustments[0] = 0.05
    except:
        pass
    return shape

def add_text(slide, left, top, width, height, text, font_size=10, bold=False,
             color=DARK, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, font_name="Calibri",
             italic=False, line_spacing=1.0):
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.margin_left = Pt(0); tf.margin_right = Pt(0)
    tf.margin_top = Pt(0); tf.margin_bottom = Pt(0)
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    
    lines = text.split('\n')
    for i, line in enumerate(lines):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.alignment = align
        if line_spacing:
            p.line_spacing = line_spacing
        r = p.add_run()
        r.text = line
        r.font.name = font_name
        r.font.size = Pt(font_size)
        r.font.bold = bold
        r.font.italic = italic
        r.font.color.rgb = color
    return tb

# === Размеры слайда ===
SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)

# === 1. Фон слайда ===
add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, BG)

# === 2. Верхняя оранжевая полоса ===
add_rect(slide, 0, 0, SLIDE_W, Inches(0.12), ORANGE)

# === 3. Заголовок ===
add_text(slide, Inches(0.5), Inches(0.22), Inches(11), Inches(0.45),
         "TITAN B2B — ЧТО НУЖНО СДЕЛАТЬ, ЧТОБЫ ПАРТНЁРЫ ЗАХОТЕЛИ",
         font_size=19, bold=True, color=DARK, anchor=MSO_ANCHOR.MIDDLE)

add_text(slide, Inches(0.5), Inches(0.70), Inches(12), Inches(0.3),
         "7 инструментов привлечения B2B-партнёров · Фокус: мотивация дистрибьютора + инструменты продаж",
         font_size=10, italic=True, color=GRAY)

# Бренд-марка справа сверху
add_text(slide, Inches(11.0), Inches(0.22), Inches(2.0), Inches(0.45),
         "TITAN · B2B",
         font_size=12, bold=True, color=ORANGE, align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)

# === 4. Главная плашка-вопрос ===
QUESTION_TOP = Inches(1.15)
add_round_rect(slide, Inches(0.5), QUESTION_TOP, Inches(12.33), Inches(0.45), DARK)
add_text(slide, Inches(0.5), QUESTION_TOP, Inches(12.33), Inches(0.45),
         "🎯 ГЛАВНЫЙ ВОПРОС:  Что нам надо сделать, чтобы B2B-партнёры захотели работать с TITAN?",
         font_size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# === 5. 7 карточек инструментов (4 + 3) ===
GRID_TOP = Inches(1.85)
CARD_WIDTH = Inches(3.05)
CARD_HEIGHT = Inches(2.20)
CARD_GAP_X = Inches(0.12)
CARD_GAP_Y = Inches(0.15)
CARD_LEFT_START = Inches(0.5)

for i, tool in enumerate(TOOLS):
    row = i // 4
    col = i % 4
    left = CARD_LEFT_START + col * (CARD_WIDTH + CARD_GAP_X)
    top = GRID_TOP + row * (CARD_HEIGHT + CARD_GAP_Y)
    
    accent = tool["color"]
    
    # Фон карточки
    add_round_rect(slide, left, top, CARD_WIDTH, CARD_HEIGHT, WHITE,
                   line_color=LIGHT_GRAY, line_width=Pt(0.75))
    
    # Левая акцентная полоска
    add_rect(slide, left, top, Inches(0.08), CARD_HEIGHT, accent)
    
    # Номер + тег
    add_text(slide, left + Inches(0.20), top + Inches(0.10), Inches(0.7), Inches(0.3),
             tool["num"], font_size=18, bold=True, color=accent, anchor=MSO_ANCHOR.TOP)
    add_text(slide, left + Inches(0.95), top + Inches(0.13), CARD_WIDTH - Inches(1.1), Inches(0.25),
             tool["tag"], font_size=8, bold=True, color=GRAY, align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.TOP)
    
    # Заголовок
    add_text(slide, left + Inches(0.20), top + Inches(0.50), CARD_WIDTH - Inches(0.35), Inches(0.45),
             tool["title"], font_size=11, bold=True, color=DARK, anchor=MSO_ANCHOR.TOP, line_spacing=1.1)
    
    # Описание
    add_text(slide, left + Inches(0.20), top + Inches(0.95), CARD_WIDTH - Inches(0.35), Inches(1.15),
             tool["desc"], font_size=8.5, color=GRAY, anchor=MSO_ANCHOR.TOP, line_spacing=1.25)

# === 6. Нижний блок — 3 принципа работы с B2B ===
PRIN_TOP = Inches(6.55)
PRIN_HEIGHT = Inches(0.65)
PRIN_WIDTH = Inches(4.05)
PRIN_GAP = Inches(0.10)

PRINCIPLES = [
    ("💰", "ДОХОДНОСТЬ ВЫШЕ", "Маржа на крупной фасовке выше, чем в рознице. B2B — самый маржинальный канал"),
    ("🤝", "ПОДДЕРЖКА ЭМПИЛС", "Эмпилс подключается к переговорам с B2B-клиентом. Техническая экспертиза + расчёт ТСО"),
    ("📈", "РАСТУЩИЙ РЫНОК", "Промышленность восстанавливается — спрос на защитные покрытия растёт"),
]

for i, (icon, title, desc) in enumerate(PRINCIPLES):
    left = Inches(0.5) + i * (PRIN_WIDTH + PRIN_GAP)
    add_round_rect(slide, left, PRIN_TOP, PRIN_WIDTH, PRIN_HEIGHT, ORANGE_LIGHT,
                   line_color=ORANGE, line_width=Pt(0.75))
    
    # Иконка
    add_text(slide, left + Inches(0.15), PRIN_TOP, Inches(0.45), PRIN_HEIGHT,
             icon, font_size=18, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    
    # Заголовок
    add_text(slide, left + Inches(0.65), PRIN_TOP + Inches(0.05), PRIN_WIDTH - Inches(0.8), Inches(0.25),
             title, font_size=10, bold=True, color=ORANGE_DARK, anchor=MSO_ANCHOR.TOP)
    
    # Описание
    add_text(slide, left + Inches(0.65), PRIN_TOP + Inches(0.28), PRIN_WIDTH - Inches(0.8), Inches(0.35),
             desc, font_size=8, color=DARK, anchor=MSO_ANCHOR.TOP, line_spacing=1.2)

# === 7. Подвал ===
FOOTER_TOP = Inches(7.30)
add_rect(slide, 0, FOOTER_TOP, SLIDE_W, Inches(0.20), ORANGE)

# Сохраняем
prs.save(DST)
print(f"\n✅ Файл сохранён: {DST}")
print(f"   Размер: {os.path.getsize(DST):,} байт")
print(f"   Слайд вставлен на позицию 7 (после Лемана)")
print(f"   Всего слайдов: {len(prs.slides)}")
