#!/usr/bin/env python3
"""Generate DOCX with all TCO mockups for designer."""

import os
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

OUT_PATH = '/home/z/my-project/download/Макеты_ТСО_на_банке_для_дизайнера.docx'
IMG_DIR = '/home/z/my-project/download/tco_mockups'

P = {
    'primary': RGBColor(0xC6, 0x28, 0x28),
    'dark': RGBColor(0x1A, 0x1F, 0x36),
    'secondary': RGBColor(0x5A, 0x60, 0x80),
    'green': RGBColor(0x16, 0xA3, 0x4A),
}


def set_cell_bg(cell, color_hex):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), color_hex)
    tc_pr.append(shd)


def add_h(doc, text, level=1, color=None):
    if color is None:
        color = P['primary'] if level == 1 else P['dark']
    h = doc.add_heading('', level=level)
    run = h.add_run(text)
    run.font.color.rgb = color
    run.font.name = 'Calibri'


def add_p(doc, text, size=11, bold=False, italic=False, color=None, indent=False):
    p = doc.add_paragraph()
    if indent:
        p.paragraph_format.first_line_indent = Cm(0.5)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.line_spacing = 1.25
    run = p.add_run(text)
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.name = 'Calibri'
    if color:
        run.font.color.rgb = color
    return p


def add_image(doc, path, width_cm=24):
    if os.path.exists(path):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run()
        run.add_picture(path, width=Cm(width_cm))


doc = Document()
style = doc.styles['Normal']
style.font.name = 'Calibri'
style.font.size = Pt(11)
style.paragraph_format.line_spacing = 1.25

section = doc.sections[0]
section.orientation = WD_ORIENT.LANDSCAPE
section.page_width = Cm(29.7)
section.page_height = Cm(21.0)
section.left_margin = Cm(2.0)
section.right_margin = Cm(2.0)
section.top_margin = Cm(1.5)
section.bottom_margin = Cm(1.5)

# Cover
cover = doc.add_table(rows=1, cols=1)
cell = cover.rows[0].cells[0]
set_cell_bg(cell, 'C62828')
cell.text = ''
p = cell.paragraphs[0]
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(30)
p.paragraph_format.space_after = Pt(10)
run = p.add_run('TITAN | Эмпилс')
run.font.size = Pt(32)
run.font.bold = True
run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
p2 = cell.add_paragraph()
p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
p2.paragraph_format.space_after = Pt(25)
run = p2.add_run('Визуальные макеты ТСО-элементов на банке')
run.font.size = Pt(20)
run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
p3 = cell.add_paragraph()
p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
p3.paragraph_format.space_after = Pt(30)
run = p3.add_run('7 макетов для передачи дизайнеру\nИнтеграция ТСО и цепочек нанесения на этикетку')
run.font.size = Pt(13)
run.font.italic = True
run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

doc.add_page_break()

# Introduction
add_h(doc, 'Задача', level=1)
add_p(doc,
    'Интегрировать элементы ТСО (совокупной стоимости владения) и информацию о '
    'технологических цепочках нанесения прямо на банку TITAN 3в1. Ниже — 7 визуальных '
    'макетов с пояснениями. Макеты демонстрируют концепцию и расположение, '
    'но не являются финальным дизайном.', indent=True)

add_p(doc, 'Принцип: «Банка — крючок, экран — деталь».', bold=True, color=P['primary'])
add_p(doc,
    'На банке размещаем только самое важное — то, что цепляет за 3 секунды. '
    'Полный расчёт ТСО — по QR-коду на телефоне покупателя.', indent=True)

doc.add_page_break()

# Mockup 1
add_h(doc, 'Макет 1. QR-код → калькулятор ТСО', level=1)
add_image(doc, os.path.join(IMG_DIR, '1_qr_block.png'), width_cm=22)
add_p(doc, 'Размещение: обратная сторона этикетки, внизу. Размер QR: 2×2 см.', bold=True)
add_p(doc,
    'Покупатель сканирует QR → попадает на калькулятор ТСО + видео-инструкцию 2 мин. '
    'Работает 24/7 без продавца. Самый компактный инструмент — занимает минимум места, '
    'даёт максимум информации.', indent=True)

doc.add_page_break()

# Mockup 2
add_h(doc, 'Макет 2. «1 банка = 3 продукта»', level=1)
add_image(doc, os.path.join(IMG_DIR, '2_one_equals_three.png'), width_cm=24)
add_p(doc, 'Размещение: лицевая сторона, под названием продукта. Размер: ~3×2 см.', bold=True)
add_p(doc,
    'Визуальная формула: 1 банка TITAN = 3 отдельных продукта (грунт + эмаль + преобразователь). '
    'Понятно без слов, даже ребёнку. Кружок «61%» — акцент на экономии.', indent=True)

doc.add_page_break()

# Mockup 3
add_h(doc, 'Макет 3. Цепочка нанесения (сравнение)', level=1)
add_image(doc, os.path.join(IMG_DIR, '3_chain.png'), width_cm=25)
add_p(doc, 'Размещение: оборот этикетки. Размер: ~5×4 см.', bold=True)
add_p(doc,
    'Сравнение двух цепочек: классическая (3 продукта, 2 дня) vs TITAN 3в1 (1 продукт, 1 день). '
    'Покупатель видит: меньше шагов, меньше времени, меньше денег. Внизу — итоговая экономия.', indent=True)

doc.add_page_break()

# Mockup 4
add_h(doc, 'Макет 4. «Цена за год службы»', level=1)
add_image(doc, os.path.join(IMG_DIR, '4_price_per_year.png'), width_cm=20)
add_p(doc, 'Размещение: лицевая сторона, рядом с ценником. Размер: ~3×2 см.', bold=True)
add_p(doc,
    'Переводит цену банки в «цену за год». 5 720 руб. → 532 руб./год (при 10 годах службы). '
    'Сравнение с ПФ-115 (1 386 руб./год) — показывает, что TITAN дешевле на 62% в пересчёте на год.', indent=True)

doc.add_page_break()

# Mockup 5
add_h(doc, 'Макет 5. Иконки ТСО (компактный блок)', level=1)
add_image(doc, os.path.join(IMG_DIR, '5_icons.png'), width_cm=22)
add_p(doc, 'Размещение: лицевая сторона, сбоку. Размер: 4×1.5 см.', bold=True)
add_p(doc,
    'Три крупные цифры в кружках: 1 банка, 2 слоя, 10 лет. Под каждой — «вместо». '
    'Минимум текста, максимум смысла. Понятно за 2 секунды.', indent=True)

doc.add_page_break()

# Mockup 6
add_h(doc, 'Макет 6. Шкала ТСО (сравнение за 10 лет)', level=1)
add_image(doc, os.path.join(IMG_DIR, '6_scale.png'), width_cm=24)
add_p(doc, 'Размещение: оборот этикетки. Размер: 6×3 см.', bold=True)
add_p(doc,
    'Визуальная шкала: два столбика разной длины. ПФ-115+грунт = 13 860 руб. (длинный красный), '
    'TITAN 3в1 = 5 320 руб. (короткий зелёный). Экономия 8 540 руб. (−61%). Внизу — призыв сканировать QR.', indent=True)

doc.add_page_break()

# Mockup 7 - Full can
add_h(doc, 'Макет 7. Общий вид банки (все элементы вместе)', level=1)
add_image(doc, os.path.join(IMG_DIR, '7_full_can.png'), width_cm=26)
add_p(doc, 'Лицевая сторона + оборот — все элементы ТСО на одной банке.', bold=True)

add_p(doc, 'Лицевая сторона содержит:', bold=True, color=P['primary'])
add_p(doc, '• Бренд + название продукта\n'
           '• «1=3» — визуальная формула (2×1.5 см)\n'
           '• Иконки: 1 банка / 2 слоя / 10 лет (4×1.5 см)\n'
           '• «532 руб./год службы» — зелёный блок (2×1.5 см)\n'
           '• Цвет RAL + образец цвета\n'
           '• Фасовка')

add_p(doc, 'Оборот содержит:', bold=True, color=P['primary'])
add_p(doc, '• Шкала ТСО: ПФ-115 vs TITAN за 10 лет (6×2 см)\n'
           '• Цепочка нанесения: классическая vs TITAN 3в1 (5×3 см)\n'
           '• Инструкция по нанесению (стандартная)\n'
           '• QR-код → калькулятор ТСО (2×2 см)')

doc.add_page_break()

# Summary
add_h(doc, 'Сводная таблица элементов', level=1)

from docx.enum.table import WD_TABLE_ALIGNMENT

def make_table(doc, headers, rows, col_widths=None):
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    if col_widths:
        for i, w in enumerate(col_widths):
            for cell in table.columns[i].cells:
                cell.width = Cm(w)
    for i, h in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = ''
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(h)
        run.font.bold = True
        run.font.size = Pt(10)
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        set_cell_bg(cell, 'C62828')
    for r_idx, row_data in enumerate(rows):
        for c_idx, cell_text in enumerate(row_data):
            cell = table.rows[r_idx + 1].cells[c_idx]
            cell.text = ''
            p = cell.paragraphs[0]
            run = p.add_run(str(cell_text))
            run.font.size = Pt(9)
            if r_idx % 2 == 0:
                set_cell_bg(cell, 'FAF5F5')

summary = [
    ['1', 'QR-код → калькулятор', 'Оборот, внизу', '2×2 см', 'Калькулятор ТСО + видео'],
    ['2', '«1=3» формула', 'Лицевая, под названием', '3×2 см', '1 банка вместо 3 продуктов'],
    ['3', 'Цепочка нанесения', 'Оборот', '5×4 см', 'Сравнение схем нанесения'],
    ['4', 'Цена за год службы', 'Лицевая, рядом с ценой', '3×2 см', '532 руб./год вместо 5 720'],
    ['5', 'Иконки 1/2/10', 'Лицевая, сбоку', '4×1.5 см', '1 банка, 2 слоя, 10 лет'],
    ['6', 'Шкала ТСО', 'Оборот, сверху', '6×3 см', 'Сравнение за 10 лет'],
    ['7', 'Все вместе', 'Лицевая + оборот', '—', 'Общий макет'],
]
make_table(doc, ['№', 'Элемент', 'Где', 'Размер', 'Что передаёт'],
           summary, col_widths=[1, 5, 5, 3, 7])

add_p(doc, '')
add_p(doc, 'Минимальный набор (если место ограничено):', bold=True, color=P['primary'])
add_p(doc, '• QR-код (2×2 см) — на оборот\n'
           '• «1=3» (3×2 см) — на лицевую\n'
           '• «Цена/год» (2×1.5 см) — на лицевую\n'
           '• Шкала ТСО (6×3 см) — на оборот\n'
           'Итого: 4 элемента, ~15% площади этикетки.')

# Final note
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(20)
run = p.add_run('Empils • TITAN • Макеты ТСО для дизайнера • 2026')
run.font.size = Pt(10)
run.font.italic = True
run.font.color.rgb = P['secondary']

doc.save(OUT_PATH)
print(f'Document saved: {OUT_PATH}')
print(f'Size: {os.path.getsize(OUT_PATH)} bytes')
