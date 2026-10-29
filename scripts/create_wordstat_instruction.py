#!/usr/bin/env python3
"""
Word-документ: инструкция для ручного сбора данных из Wordstat.
Пользователь идёт по 8 запросам, вписывает цифры в таблицу.
"""

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

OUT = '/home/z/my-project/download/Инструкция_Wordstat_сбор_данных.docx'

COL_PRIMARY = RGBColor(0x1E, 0x3A, 0x5F)
COL_ACCENT = RGBColor(0xF4, 0x79, 0x20)
COL_BODY = RGBColor(0x1A, 0x1F, 0x36)
COL_GRAY = RGBColor(0x6B, 0x72, 0x80)
COL_RED = RGBColor(0xC6, 0x28, 0x28)
COL_GREEN = RGBColor(0x16, 0xA3, 0x4A)


def add_para(doc, text='', size=11, bold=False, color=COL_BODY, align=None,
             font='Arial', space_before=0, space_after=4, line_spacing=1.3):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    if align is not None:
        p.alignment = align
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(space_after)
    pf.line_spacing = line_spacing
    if text:
        run = p.add_run(text)
        run.font.name = font
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.color.rgb = color
        rPr = run._element.get_or_add_rPr()
        rFonts = rPr.find(qn('w:rFonts'))
        if rFonts is None:
            rFonts = OxmlElement('w:rFonts')
            rPr.append(rFonts)
        rFonts.set(qn('w:ascii'), font)
        rFonts.set(qn('w:hAnsi'), font)
        rFonts.set(qn('w:cs'), font)
    return p


def add_bullet(doc, text, level=0, size=11, bold=False, color=COL_BODY):
    p = doc.add_paragraph(style='List Bullet')
    pf = p.paragraph_format
    pf.left_indent = Cm(1.0 + level * 0.7)
    pf.space_before = Pt(0)
    pf.space_after = Pt(2)
    pf.line_spacing = 1.3
    run = p.add_run(text)
    run.font.name = 'Arial'
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.append(rFonts)
    rFonts.set(qn('w:ascii'), 'Arial')
    rFonts.set(qn('w:hAnsi'), 'Arial')


def add_heading(doc, text, level=1):
    if level == 1:
        size, color, sb, sa = 16, COL_PRIMARY, 18, 8
    elif level == 2:
        size, color, sb, sa = 13, COL_PRIMARY, 12, 6
    else:
        size, color, sb, sa = 11, COL_ACCENT, 8, 4
    return add_para(doc, text, size=size, bold=True, color=color,
                    align=WD_ALIGN_PARAGRAPH.LEFT, space_before=sb, space_after=sa)


def set_cell_shading(cell, hex_color):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    tc_pr.append(shd)


def set_cell_borders(cell, color='D1D5DB', sz='4'):
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_borders = OxmlElement('w:tcBorders')
    for edge in ('top', 'left', 'bottom', 'right'):
        b = OxmlElement(f'w:{edge}')
        b.set(qn('w:val'), 'single')
        b.set(qn('w:sz'), sz)
        b.set(qn('w:color'), color)
        tc_borders.append(b)
    tc_pr.append(tc_borders)


def style_cell(cell, value, font_size=10, bold=False, color=COL_BODY, fill=None, align='left'):
    if fill:
        set_cell_shading(cell, fill)
    set_cell_borders(cell)
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.line_spacing = 1.2
    if align == 'center':
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    elif align == 'right':
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    else:
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run(value)
    run.font.name = 'Arial'
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.color.rgb = color


# =========================================================
# ДОКУМЕНТ
# =========================================================
doc = Document()

for section in doc.sections:
    section.top_margin = Cm(2.0)
    section.bottom_margin = Cm(2.0)
    section.left_margin = Cm(2.0)
    section.right_margin = Cm(2.0)

style = doc.styles['Normal']
style.font.name = 'Arial'
style.font.size = Pt(11)
style.font.color.rgb = COL_BODY

# Заголовок
add_para(doc, 'ИНСТРУКЦИЯ ПО СБОРУ ДАННЫХ ИЗ WORDSTAT', size=10, bold=True, color=COL_ACCENT,
         align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
add_para(doc, 'TITAN vs конкуренты — узнаваемость брендов в Яндексе', size=18, bold=True, color=COL_PRIMARY,
         align=WD_ALIGN_PARAGRAPH.CENTER, space_after=4)
add_para(doc, '8 запросов × 2 разреза (Россия + Сибирский ФО) = 16 цифр для ввода. Время: ~15-20 минут.',
         size=11, color=COL_GRAY, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=20)

# === ШАГ 1 ===
add_heading(doc, 'Шаг 1. Откройте Wordstat и авторизуйтесь', level=2)
add_para(doc, 'Перейдите по ссылке: https://wordstat.yandex.ru/', size=11, color=COL_BODY, space_after=4)
add_para(doc, 'Если не авторизованы — нажмите «Войти» и введите логин/пароль Яндекс ID. Без авторизации данные показываются, но с ограничениями и капчей.',
         size=11, color=COL_BODY, space_after=8, line_spacing=1.4)

# === ШАГ 2 ===
add_heading(doc, 'Шаг 2. Для каждого из 8 запросов соберите 2 цифры', level=2)
add_para(doc, 'Для каждого запроса вам нужно получить 2 показателя:', size=11, color=COL_BODY, space_after=4)
add_bullet(doc, 'Россия в целом — цифра над графиком «История запросов» (показывает за месяц)')
add_bullet(doc, 'Сибирский ФО — цифра из таблицы «По регионам» для строки «Сибирский федеральный округ»')

add_para(doc, '', size=4, space_after=4)
add_para(doc, 'ВАЖНО: На вкладке «По словам» в поле ввода копируйте запрос БЕЗ кавычек. После ввода нажмите кнопку «Подобрать».',
         size=10, bold=True, color=COL_RED, space_after=8)

# === ТАБЛИЦА ДЛЯ ЗАПОЛНЕНИЯ ===
add_heading(doc, 'Шаг 3. Таблица для заполнения', level=2)
add_para(doc, 'Впишите цифры в жёлтые ячейки. Для каждого запроса — 2 числа (Россия + Сибирский ФО).',
         size=11, color=COL_BODY, space_after=6)

# Таблица: 8 запросов
queries = [
    ('1', 'TITAN краска', 'общий бренд'),
    ('2', 'TITAN грунт-эмаль', 'основная категория TITAN'),
    ('3', 'Dufa краска', 'конкурент премиум'),
    ('4', 'Tikkurila краска', 'конкурент премиум'),
    ('5', 'Лакра краска', 'конкурент масс-маркет'),
    ('6', 'Elcon краска', 'конкурент средний'),
    ('7', 'Dali краска по ржавчине', 'конкурент 3в1'),
    ('8', 'грунт-эмаль 3 в 1', 'категория без бренда'),
]

table = doc.add_table(rows=1 + len(queries), cols=5)
table.alignment = WD_TABLE_ALIGNMENT.CENTER
table.autofit = False
col_widths = [Cm(0.8), Cm(6.5), Cm(4.5), Cm(3.0), Cm(3.0)]
for i, w in enumerate(col_widths):
    table.columns[i].width = w

# Шапка
headers = ['№', 'Что вводить в Wordstat', 'Зачем', 'Россия', 'Сибирский ФО']
for i, h in enumerate(headers):
    cell = table.cell(0, i)
    cell.width = col_widths[i]
    style_cell(cell, h, font_size=10, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF), fill='1E3A5F', align='center')

# Данные
yellow_fill = 'FEF3C7'
for ri, (n, q, purpose) in enumerate(queries, 1):
    fill = 'FFFFFF' if ri % 2 == 1 else 'F9FAFB'
    style_cell(table.cell(ri, 0), n, font_size=10, bold=True, color=COL_PRIMARY, fill=fill, align='center')
    style_cell(table.cell(ri, 1), q, font_size=10, bold=True, color=COL_BODY, fill=fill)
    style_cell(table.cell(ri, 2), purpose, font_size=9, color=COL_GRAY, fill=fill)
    # Пустые жёлтые ячейки для заполнения
    style_cell(table.cell(ri, 3), '', font_size=11, bold=True, color=COL_PRIMARY, fill=yellow_fill, align='center')
    style_cell(table.cell(ri, 4), '', font_size=11, bold=True, color=COL_PRIMARY, fill=yellow_fill, align='center')

# === ПОДСКАЗКА: ГДЕ ИСКАТЬ ЦИФРЫ ===
add_para(doc, '', size=4, space_after=8)
add_heading(doc, 'Где какие цифры смотреть в Wordstat', level=2)

# Таблица-подсказка
hint_table = doc.add_table(rows=4, cols=2)
hint_table.alignment = WD_TABLE_ALIGNMENT.CENTER
hint_table.autofit = False
hint_widths = [Cm(5.5), Cm(12.5)]
for i, w in enumerate(hint_widths):
    hint_table.columns[i].width = w

hints = [
    ('Что ищем', 'Где в Wordstat смотреть'),
    ('Цифра по России', 'Большое число над графиком «История запросов» в самом верху страницы. Это суммарное число показов за последний месяц по всем регионам.'),
    ('Цифра по Сибирскому ФО', 'Прокрутите вниз до блока «По регионам». Найдите строку «Сибирский федеральный округ» — цифра во втором столбце.'),
    ('Динамика 12 месяцев (опц.)', 'График «История запросов» — просто посмотрите, растёт или падает кривая за год. Это пригодится для выводов, но в таблицу цифры не вносим.'),
]

for ri, (label, text) in enumerate(hints):
    if ri == 0:
        style_cell(hint_table.cell(ri, 0), label, font_size=10, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF), fill='1E3A5F')
        style_cell(hint_table.cell(ri, 1), text, font_size=10, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF), fill='1E3A5F')
    else:
        fill = 'FFFFFF' if ri % 2 == 1 else 'F9FAFB'
        style_cell(hint_table.cell(ri, 0), label, font_size=10, bold=True, color=COL_PRIMARY, fill=fill)
        style_cell(hint_table.cell(ri, 1), text, font_size=10, color=COL_BODY, fill=fill)

# === ПРИМЕР ===
add_para(doc, '', size=4, space_after=8)
add_heading(doc, 'Пример заполнения (для запроса «TITAN краска»)', level=2)
add_para(doc, 'После ввода «TITAN краска» в Wordstat вы видите:', size=11, color=COL_BODY, space_after=4)
add_bullet(doc, 'Над графиком: «7 253 показа в месяц» → вносим 7253 в колонку «Россия»')
add_bullet(doc, 'В таблице «По регионам»: «Сибирский федеральный округ — 412» → вносим 412 в колонку «Сибирский ФО»')
add_para(doc, 'И так для всех 8 запросов.', size=11, color=COL_BODY, space_after=8)

# === ЧТО ДАЛЬШЕ ===
add_para(doc, '', size=4, space_after=8)
add_heading(doc, 'Шаг 4. Что делать с заполненной таблицей', level=2)
add_para(doc, 'После заполнения всех 16 жёлтых ячеек:', size=11, color=COL_BODY, space_after=4)
add_bullet(doc, 'Вариант 1: пришлите мне заполненный Word-файл (или просто текстом 16 цифр в сообщении).')
add_bullet(doc, 'Вариант 2: сделайте 8 скриншотов — по одному на каждый запрос (видны и Россия, и регионы).')

add_para(doc, '', size=4, space_after=8)
add_para(doc, 'Я на основе этих данных соберу:', size=11, color=COL_BODY, space_after=4, bold=True)
add_bullet(doc, 'Excel-таблицу «Узнаваемость TITAN vs конкурентов в Яндексе» с расчётом доли рынка в запросах и индекса по регионам.')
add_bullet(doc, 'PPT-слайд с горизонтальной диаграммой сравнения брендов и выводами для программы продвижения.')
add_bullet(doc, 'Анализ «где TITAN отстаёт / опережает конкурентов по узнаваемости» с разбивкой по регионам.')

# Подпись
add_para(doc, '', size=8, space_after=10)
add_para(doc, 'Источник: https://wordstat.yandex.ru/ (требуется авторизация Яндекс ID).',
         size=9, color=COL_GRAY, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=10)

doc.save(OUT)
print(f'OK: {OUT}')
import os
print(f'Размер: {os.path.getsize(OUT) / 1024:.1f} KB')
