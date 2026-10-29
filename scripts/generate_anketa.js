#!/usr/bin/env python3
"""
Generate modernized questionnaire for Lemana PRO delegation visit to Empils.
"""
from docx import Document
from docx.shared import Pt, Cm, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

doc = Document()

# Page margins
for section in doc.sections:
    section.top_margin = Cm(2)
    section.bottom_margin = Cm(2)
    section.left_margin = Cm(2.5)
    section.right_margin = Cm(2)

# Styles
style = doc.styles['Normal']
font = style.font
font.name = 'Arial'
font.size = Pt(10.5)
style.element.rPr.rFonts.set(qn('w:eastAsia'), 'Arial')

# ===== TITLE =====
title = doc.add_paragraph()
title.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = title.add_run('АНКЕТА\nдля делегации «Лемана ПРО»')
run.bold = True
run.font.size = Pt(16)
run.font.color.rgb = RGBColor(0, 51, 102)

subtitle = doc.add_paragraph()
subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = subtitle.add_run('Подготовка к визиту на завод «Эмпилс»')
run.font.size = Pt(12)
run.font.color.rgb = RGBColor(0, 51, 102)

# Intro
intro = doc.add_paragraph()
intro.alignment = WD_ALIGN_PARAGRAPH.LEFT
run = intro.add_run(
    'Анонимно. Ответы помогут нам адаптировать программу визита под ваши потребности '
    'и сделать презентацию максимально полезной.\n'
    'Пожалуйста, заполните до 27 июня.'
)
run.font.size = Pt(10)
run.italic = True

doc.add_paragraph()  # spacer

def add_block_header(text):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.bold = True
    run.font.size = Pt(13)
    run.font.color.rgb = RGBColor(0, 51, 102)
    # Add bottom border
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '8')
    bottom.set(qn('w:space'), '4')
    bottom.set(qn('w:color'), '003366')
    pBdr.append(bottom)
    pPr.append(pBdr)

def add_question(num, text, note=''):
    p = doc.add_paragraph()
    run = p.add_run(f'{num}. {text}')
    run.bold = True
    run.font.size = Pt(10.5)
    if note:
        run2 = p.add_run(f'\n({note})')
        run2.font.size = Pt(9)
        run2.italic = True
        run2.font.color.rgb = RGBColor(100, 100, 100)

def add_option(text, indent=True):
    p = doc.add_paragraph()
    if indent:
        p.paragraph_format.left_indent = Cm(1)
    run = p.add_run(f'○  {text}')
    run.font.size = Pt(10)

def add_option_multi(text, indent=True):
    p = doc.add_paragraph()
    if indent:
        p.paragraph_format.left_indent = Cm(1)
    run = p.add_run(f'☐  {text}')
    run.font.size = Pt(10)

def add_free_field(lines=1):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(1)
    run = p.add_run('________________________________')
    run.font.size = Pt(10)
    run.font.color.rgb = RGBColor(150, 150, 150)

# ============================================================
# БЛОК 1. РОЛЬ И ОПЫТ
# ============================================================
add_block_header('БЛОК 1. Ваша роль и опыт')

add_question(1, 'Ваша должность в «Лемана ПРО»')
add_option('Руководитель отдела (закупки / категорийный менеджмент)')
add_option('Менеджер по работе с поставщиками')
add_option('Консультант / продавец торгового зала')
add_option('Мерчандайзер / супервайзер')
add_option('Другое: ________________')

add_question(2, 'Ваш опыт работы в категории ЛКМ')
add_option('Менее 1 года')
add_option('1–3 года')
add_option('3–5 лет')
add_option('Более 5 лет')

add_question(3, 'В каком формате вы чаще всего общаетесь с покупателями ЛКМ?')
add_option('Ежедневно помогаю выбрать краску в зале')
add_option('Периодически консультирую по запросу')
add_option('Редко сталкиваюсь с ЛКМ — другая категория')
add_option('Не работаю с покупателями напрямую')

# ============================================================
# БЛОК 2. УЗНАВАЕМОСТЬ БРЕНДОВ ЭМПИЛС
# ============================================================
add_block_header('БЛОК 2. Узнаваемость брендов «Эмпилс»')

add_question(4, 'Какие бренды ЛКМ вы можете вспомнить без подсказок?\n(перечислите в свободной форме — любые, не только Эмпилс)')
add_free_field(3)

add_question(5, 'Какие торговые марки компании «Эмпилс» вам известны?\n(отметьте все, которые знаете)')

# Use a table for checkbox list - 2 columns
table = doc.add_table(rows=5, cols=2)
table.alignment = WD_TABLE_ALIGNMENT.CENTER
brands = [
    ('☐  ПростоКРАСКА', '☐  Ореол Premium'),
    ('☐  Ореол', '☐  TITAN'),
    ('☐  AvantGarde', '☐  Luxens'),
    ('☐  Расцвет', '☐  ЭКОНОМ Л'),
    ('☐  ГФ-021', '☐  Другие: __________'),
]
for i, (left, right) in enumerate(brands):
    for j, text in enumerate([left, right]):
        cell = table.cell(i, j)
        p = cell.paragraphs[0]
        run = p.add_run(text)
        run.font.size = Pt(10)

doc.add_paragraph()

add_question(6, 'Какая ассоциация возникает у вас в первую очередь при слове «Эмпилс»?\n(выберите одну)')
add_option('Надёжность, проверенное качество')
add_option('Советское наследие / старый завод')
add_option('Устаревшие продукты')
add_option('Современные технологии и инновации')
add_option('Доступная цена / эконом-сегмент')
add_option('Широкий ассортимент')
add_option('Не могу сказать — мало знаком(а)')

# ============================================================
# БЛОК 3. ВОСПРИЯТИЕ ПРОДУКЦИИ НА ПОЛКЕ ЛЕМАНЫ
# ============================================================
add_block_header('БЛОК 3. Восприятие продукции «Эмпилс» на полке «Лемана ПРО»')

add_question(7, 'Как вы оцениваете продукцию «Эмпилс» по сравнению с другими брендами\nв вашей категории ЛКМ?', 'по 5-балльной шкале, где 1 — значительно хуже, 5 — значительно лучше')

# Rating table
table = doc.add_table(rows=7, cols=3)
table.alignment = WD_TABLE_ALIGNMENT.CENTER
table.cell(0, 0).text = 'Критерий'
table.cell(0, 1).text = 'Оценка (1–5)'
table.cell(0, 2).text = 'Затрудняюсь'
criteria = ['Качество покрытия', 'Цена / ценность', 'Ассортимент цветов', 
            'Узнаваемость для покупателя', 'Удобство фасовки', 'Рекламная поддержка']
for i, c in enumerate(criteria):
    table.cell(i+1, 0).text = c
    table.cell(i+1, 1).text = '1    2    3    4    5'
    table.cell(i+1, 2).text = '☐'

# Style table header
for cell in table.rows[0].cells:
    for p in cell.paragraphs:
        for run in p.runs:
            run.bold = True
            run.font.size = Pt(9.5)

doc.add_paragraph()

add_question(8, 'Какие бренды-конкуренты наиболее часто покупают вместо продукции\n«Эмпилс» в ваших магазинах?\n(выберите до 3-х)')
add_option_multi('Лакра')
add_option_multi('Текс')
add_option_multi('Новоколор')
add_option_multi('Профилюкс')
add_option_multi('СТМ «Лемана ПРО» (собственная марка)')
add_option_multi('Другое: ________________')

add_question(9, 'Как вы считаете, какого сегмента в основном продукция «Эмпилс»?')
add_option('Только эконом')
add_option('В основном эконом, есть средний')
add_option('Эконом + средний + премиум — широкий охват')
add_option('В основном средний сегмент')
add_option('Затрудняюсь оценить')

# ============================================================
# БЛОК 4. ЗНАНИЕ И ВОСПРИЯТИЕ TITAN 3в1
# ============================================================
add_block_header('БЛОК 4. Грунт-эмаль TITAN 3в1 — знание и восприятие')

add_question(10, 'Знаете ли вы о грунт-эмали по ржавчине TITAN 3в1 в ассортименте «Лемана ПРО»?')
add_option('Да, хорошо знаю и могу объяснить покупателю')
add_option('Слышал(а), но подробностей не знаю')
add_option('Вижу на полке, но никогда не рекомендовал(а)')
add_option('Нет, впервые слышу')

add_question(11, 'Что, на ваш взгляд, означает формат «3в1» для грунт-эмали TITAN?')
add_option('Грунтовка + краска + антикоррозийная защита — три функции в одном')
add_option('Грунтовка + краска + лак')
add_option('Грунтовка + антисептик + краска')
add_option('Три слоя покрытия')
add_option('Затрудняюсь ответить')

add_question(12, 'Если бы покупатель спросил: «Какую краску взять для ржавых ворот на даче?» —\nчто бы вы порекомендовали?')
add_option('Грунт-эмаль 3в1 (TITAN или аналог)')
add_option('Сначала грунт ГФ-021, потом эмаль ПФ-115')
add_option('СТМ сети (лучшая цена)')
add_option('Уточнил(а) бы бюджет и тогда порекомендовал(а)')
add_option('Другое: ________________')

add_question(13, 'Что, на ваш взгляд, может стать главным препятствием для продаж\nTITAN 3в1 в «Лемана ПРО»?\n(выберите 1–2 варианта)')
add_option_multi('Цена выше, чем у обычной эмали ПФ-115')
add_option_multi('Покупатель не знает бренд «Титан»')
add_option_multi('Недоверие к формату «3в1» — сомнения, что грунт + краска в одной банке работают')
add_option_multi('Конкуренция с СТМ «Лемана ПРО»')
add_option_multi('Сложность объяснить выгоду покупателю')
add_option_multi('Недостаточная рекламная поддержка на полке')
add_option_multi('Другое: ________________')

# ============================================================
# БЛОК 5. ГОТОВНОСТЬ РЕКОМЕНДОВАТЬ И ПОТРЕБНОСТИ
# ============================================================
add_block_header('БЛОК 5. Готовность рекомендовать и потребности в поддержке')

add_question(14, 'Насколько вы готовы рекомендовать TITAN 3в1 покупателю,\nкоторый ищет краску для металла?', 
             'по шкале от 0 до 10, где 0 — точно не рекомендую, 10 — точно рекомендую')
add_option('0  1  2  3  4  5  6  7  8  9  10')

add_question(15, 'Что помогло бы вам увереннее продавать TITAN 3в1?\n(выберите до 3-х)')
add_option_multi('Краткая шпаргалка / подсказка на полке (какие вопросы задать покупателю)')
add_option_multi('Примеры образцов покрытия (до/после)')
add_option_multi('Обучение: как за 2 минуты объяснить выгоду формата 3в1')
add_option_multi('Калькулятор ТСО (полная стоимость покраски с учётом грунта, времени, слоёв)')
add_option_multi('Промо-материалы для покупателя (листовки, видео на экране)')
add_option_multi('Сравнительная таблица TITAN 3в1 vs конкуренты')
add_option_multi('Мотивационная программа для продавцов')
add_option_multi('Другое: ________________')

add_question(16, 'Использовали ли вы лично продукцию «Эмпилс» (любую)?')
add_option('Да, регулярно')
add_option('Да, иногда')
add_option('Пробовал(а) один раз')
add_option('Нет, не доводилось')

add_question(17, 'Какой один вопрос о «Эмпилс» или TITAN 3в1 вы бы хотели\nзадать на заводе?', 
             'открытый вопрос')
add_free_field(2)

# ===== FOOTER =====
doc.add_paragraph()
footer = doc.add_paragraph()
footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = footer.add_run('Спасибо за ваше время! Результаты будут использованы для адаптации программы визита.')
run.font.size = Pt(9)
run.italic = True
run.font.color.rgb = RGBColor(100, 100, 100)

# Save
output_path = '/home/z/my-project/download/Анкета_Лемана_ПРО_модернизированная.docx'
doc.save(output_path)
print(f"Saved: {output_path}")
