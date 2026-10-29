#!/usr/bin/env python3
"""
Word-документ: 4 варианта смыслового посыла для брендированного ценника TITAN TCO.
Для распечатки и обсуждения с дизайнером.
"""

from docx import Document
from docx.shared import Pt, Cm, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

OUT = '/home/z/my-project/download/TITAN_ценник_4_варианта_TCO.docx'

COL_PRIMARY = RGBColor(0x1E, 0x3A, 0x5F)
COL_ACCENT = RGBColor(0xF4, 0x79, 0x20)
COL_BODY = RGBColor(0x1A, 0x1F, 0x36)
COL_GRAY = RGBColor(0x6B, 0x72, 0x80)
COL_RED = RGBColor(0xC6, 0x28, 0x28)
COL_GREEN = RGBColor(0x16, 0xA3, 0x4A)
COL_WHITE = RGBColor(0xFF, 0xFF, 0xFF)


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


def add_heading(doc, text, level=1):
    if level == 1:
        size, color, sb, sa = 16, COL_PRIMARY, 18, 8
    elif level == 2:
        size, color, sb, sa = 13, COL_PRIMARY, 12, 6
    else:
        size, color, sb, sa = 11, COL_ACCENT, 8, 4
    return add_para(doc, text, size=size, bold=True, color=color,
                    align=WD_ALIGN_PARAGRAPH.LEFT, space_before=sb, space_after=sa)


def add_bullet(doc, text, size=10.5, bold=False, color=COL_BODY):
    p = doc.add_paragraph(style='List Bullet')
    pf = p.paragraph_format
    pf.left_indent = Cm(0.8)
    pf.space_before = Pt(0)
    pf.space_after = Pt(2)
    pf.line_spacing = 1.25
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


def style_cell(cell, value, font_size=9, bold=False, color=COL_BODY, fill=None, align='left', valign='center'):
    if fill:
        set_cell_shading(cell, fill)
    set_cell_borders(cell)
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.line_spacing = 1.15
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


def add_mockup(doc, title, lines, accent_color_hex='F47920', mockup_fill='F3F4F6'):
    """Добавляет визуальный макет ценника как таблицу."""
    add_heading(doc, title, level=2)
    
    # Таблица-макет (1 колонка, имитирующая ценник)
    table = doc.add_table(rows=len(lines), cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    table.columns[0].width = Cm(12)
    
    for ri, (text, font_size, bold, color_hex, bg_hex, align) in enumerate(lines):
        cell = table.cell(ri, 0)
        cell.width = Cm(12)
        bg = bg_hex if bg_hex else mockup_fill
        color = RGBColor.from_string(color_hex) if color_hex else COL_BODY
        style_cell(cell, text, font_size=font_size, bold=bold, color=color, fill=bg, align=align)
    
    doc.add_paragraph()


# =========================================================
# ДОКУМЕНТ
# =========================================================
doc = Document()

for section in doc.sections:
    section.top_margin = Cm(1.8)
    section.bottom_margin = Cm(1.8)
    section.left_margin = Cm(2.0)
    section.right_margin = Cm(2.0)

style = doc.styles['Normal']
style.font.name = 'Arial'
style.font.size = Pt(11)
style.font.color.rgb = COL_BODY

# === ШАПКА ===
add_para(doc, 'БРЕНДИРОВАННЫЙ ЦЕННИК TITAN — TCO', size=10, bold=True, color=COL_ACCENT,
         align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
add_para(doc, '4 варианта смыслового посыла', size=18, bold=True, color=COL_PRIMARY,
         align=WD_ALIGN_PARAGRAPH.CENTER, space_after=4)
add_para(doc, 'Для распечатки и обсуждения с дизайнером · Август 2026',
         size=10, color=COL_GRAY, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=16)

# === ПРОБЛЕМА ===
add_heading(doc, 'Проблема текущих вариантов', level=1)
add_para(doc,
    'Все 3 текущих варианта ценника показывают «Цена за год: ___/ГОД» — но без точки отсчёта. '
    'Покупатель видит цифру 49 ₽/год и не понимает: это много или мало? Не с чем сравнить. '
    'Нужен ЯКОРЬ — цена, которую покупатель знает (ПФ-115), рядом с ценой TITAN.',
    size=11, space_after=8, line_spacing=1.4)

add_para(doc, 'Принцип: «Цифра + Якорь + Вывод»', size=11, bold=True, color=COL_ACCENT, space_after=4)
add_bullet(doc, 'Цифра — цена TITAN за год (49 ₽) — маленькая, не пугает')
add_bullet(doc, 'Якорь — цена ПФ-115 за год (300 ₽) — даёт точку отсчёта, мозг сам сравнивает')
add_bullet(doc, 'Вывод — одна фраза-заключение: «в 6 раз дешевле» / «экономия 2 000 ₽» / «5 банок vs 1 банка»')

doc.add_paragraph()

# =========================================================
# ВАРИАНТ 1
# =========================================================
add_heading(doc, 'Вариант 1 — «В 6 раз дешевле»', level=1)

add_para(doc, 'Концепция: две цены рядом, мозг сам делает вывод.', size=10, color=COL_GRAY, space_after=6)

add_mockup(doc, 'Макет ценника', [
    ('TITAN 3в1 · 0.75 л · RAL 9005', 10, False, '1E3A5F', 'F3F4F6', 'left'),
    ('', 6, False, '000000', 'F3F4F6', 'center'),
    ('Цена: 494 ₽  →  49 ₽ / год', 14, True, '1E3A5F', 'FFFFFF', 'center'),
    ('', 4, False, '000000', 'F3F4F6', 'center'),
    ('ПФ-115: ~300 ₽/год    TITAN: 49 ₽/год', 11, True, 'C62828', 'FEE2E2', 'center'),
    ('← в 6 раз дешевле', 12, True, '16A34A', 'DCFCE7', 'center'),
    ('', 4, False, '000000', 'F3F4F6', 'center'),
    ('КРАСИШЬ 1 РАЗ — СЛУЖИТ 10 ЛЕТ', 11, True, '1E3A5F', 'F3F4F6', 'center'),
])

add_para(doc, 'Почему работает:', size=10, bold=True, color=COL_GREEN, space_after=2)
add_bullet(doc, 'Покупатель видит 2 цифры рядом — 300 и 49. Мозг сам делает вывод: «в 6 раз дешевле»')
add_bullet(doc, 'Не нужно считать — сравнение моментальное')
add_bullet(doc, '«В 6 раз дешевле» — конкретная, измеримая выгода')

doc.add_page_break()

# =========================================================
# ВАРИАНТ 2
# =========================================================
add_heading(doc, 'Вариант 2 — «5 банок ПФ vs 1 банка TITAN»', level=1)

add_para(doc, 'Концепция: сравнение за 10 лет, а не за год. Не цена за год, а количество перекрасок.', size=10, color=COL_GRAY, space_after=6)

add_mockup(doc, 'Макет ценника', [
    ('TITAN 3в1 · 0.75 л · RAL 9005', 10, False, '1E3A5F', 'F3F4F6', 'left'),
    ('', 6, False, '000000', 'F3F4F6', 'center'),
    ('ПФ-115:  красить 5 раз  =  2 500 ₽', 11, True, 'C62828', 'FEE2E2', 'left'),
    ('TITAN:    красить 1 раз  =    494 ₽', 11, True, '16A34A', 'DCFCE7', 'left'),
    ('', 4, False, '000000', 'F3F4F6', 'center'),
    ('ЭКОНОМИЯ: 2 000 ₽', 14, True, 'F47920', 'FFFFFF', 'center'),
    ('', 4, False, '000000', 'F3F4F6', 'center'),
    ('КРАСИШЬ 1 РАЗ — СЛУЖИТ 10 ЛЕТ', 11, True, '1E3A5F', 'F3F4F6', 'center'),
])

add_para(doc, 'Почему работает:', size=10, bold=True, color=COL_GREEN, space_after=2)
add_bullet(doc, 'Не требует вычислений — 5 vs 1, мозг считает мгновенно')
add_bullet(doc, 'Показывает конкретную экономию — 2 000 ₽, а не абстрактные «49 ₽/год»')
add_bullet(doc, 'Объясняет WHY — «почему TITAN дешевле? Потому что не нужно перекрашивать»')
add_bullet(doc, 'Работает для любого объёма и цвета — можно подставить любые цифры')

doc.add_page_break()

# =========================================================
# ВАРИАНТ 3
# =========================================================
add_heading(doc, 'Вариант 3 — «Цена + Цена за год + Якорь» (компактный)', level=1)

add_para(doc, 'Концепция: минимум текста. Три цены по убыванию. Скобки говорят: «а у других — вот столько».', size=10, color=COL_GRAY, space_after=6)

add_mockup(doc, 'Макет ценника', [
    ('TITAN 3в1 · 0.75 л', 10, False, '1E3A5F', 'F3F4F6', 'left'),
    ('', 6, False, '000000', 'F3F4F6', 'center'),
    ('Цена: 494 ₽', 12, True, '1E3A5F', 'FFFFFF', 'center'),
    ('Цена за год: 49 ₽', 14, True, 'F47920', 'FFFFFF', 'center'),
    ('(ПФ-115: 300 ₽/год)', 9, False, '6B7280', 'F3F4F6', 'center'),
    ('', 4, False, '000000', 'F3F4F6', 'center'),
    ('КРАСИШЬ 1 РАЗ — СЛУЖИТ 10 ЛЕТ', 11, True, '1E3A5F', 'F3F4F6', 'center'),
])

add_para(doc, 'Почему работает:', size=10, bold=True, color=COL_GREEN, space_after=2)
add_bullet(doc, 'Минимум текста — помещается на самый маленький ценник')
add_bullet(doc, 'Три цены по убыванию: 494 → 49 → (300 в скобках)')
add_bullet(doc, 'Скобки = «а у конкурентов — вот столько» — мягкое сравнение без агрессии')
add_bullet(doc, 'Подходит для полки, где мало места')

doc.add_page_break()

# =========================================================
# ВАРИАНТ 4
# =========================================================
add_heading(doc, 'Вариант 4 — «Дешевле чашки кофе» (эмоциональный)', level=1)

add_para(doc, 'Концепция: сравнение с понятным бытовым якорем. 49 ₽/год — это смешно дёшево, и это правда.', size=10, color=COL_GRAY, space_after=6)

add_mockup(doc, 'Макет ценника', [
    ('TITAN 3в1 · 0.75 л', 10, False, '1E3A5F', 'F3F4F6', 'left'),
    ('', 6, False, '000000', 'F3F4F6', 'center'),
    ('49 ₽ / ГОД', 18, True, 'F47920', 'FFFFFF', 'center'),
    ('дешевле чашки кофе', 10, False, '6B7280', 'FFFFFF', 'center'),
    ('', 4, False, '000000', 'F3F4F6', 'center'),
    ('ПФ-115: 300 ₽/год    TITAN: 49 ₽/год', 10, True, '1E3A5F', 'F3F4F6', 'center'),
    ('', 4, False, '000000', 'F3F4F6', 'center'),
    ('КРАСИШЬ 1 РАЗ — СЛУЖИТ 10 ЛЕТ', 11, True, '1E3A5F', 'F3F4F6', 'center'),
])

add_para(doc, 'Почему работает:', size=10, bold=True, color=COL_GREEN, space_after=2)
add_bullet(doc, 'Сравнение с понятным бытовым якорем (чашка кофе ≈ 150-200 ₽)')
add_bullet(doc, '«Дешевле кофе» — эмоционально. 49 ₽/год звучит смешно дёшево — и это правда')
add_bullet(doc, 'Запоминается — ассоциация «TITAN = дешевле кофе» остаётся в голове')
add_bullet(doc, 'Хорошо работает для DIY-покупателей, не профессионалов')

doc.add_page_break()

# =========================================================
# СВОДНАЯ ТАБЛИЦА
# =========================================================
add_heading(doc, 'Сводная таблица вариантов', level=1)

table = doc.add_table(rows=5, cols=4)
table.alignment = WD_TABLE_ALIGNMENT.CENTER
table.autofit = False
col_widths = [Cm(3.5), Cm(5.0), Cm(4.0), Cm(4.5)]
for i, w in enumerate(col_widths):
    table.columns[i].width = w

headers = ['Вариант', 'Главный посыл', 'Цифры на ценнике', 'Когда использовать']
for i, h in enumerate(headers):
    cell = table.cell(0, i)
    cell.width = col_widths[i]
    style_cell(cell, h, font_size=10, bold=True, color=COL_WHITE, fill='1E3A5F', align='center')

data = [
    ('1. «В 6 раз дешевле»', 'Сравнение цен за год', '494 ₽ → 49 ₽/год\nПФ: 300 ₽/год\nВ 6 раз дешевле', 'Основной ценник, средний размер'),
    ('2. «5 банок vs 1»', 'Сравнение за 10 лет', 'ПФ: 5 раз = 2 500 ₽\nTITAN: 1 раз = 494 ₽\nЭкономия: 2 000 ₽', 'Презентация, буклет, крупный ценник'),
    ('3. «Цена + Якорь»', 'Три цены по убыванию', 'Цена: 494 ₽\nЗа год: 49 ₽\n(ПФ: 300 ₽/год)', 'Маленький ценник, мало места'),
    ('4. «Дешевле кофе»', 'Эмоциональное сравнение', '49 ₽/ГОД\nдешевле чашки кофе\nПФ: 300 ₽/год', 'DIY-покупатели, соцсети, маркетинг'),
]
for ri, (v, msg, nums, when) in enumerate(data, 1):
    fill = 'F9FAFB' if ri % 2 == 1 else 'FFFFFF'
    style_cell(table.cell(ri, 0), v, font_size=9, bold=True, color=COL_PRIMARY, fill=fill, align='left')
    style_cell(table.cell(ri, 1), msg, font_size=9, fill=fill, align='left')
    style_cell(table.cell(ri, 2), nums, font_size=9, fill=fill, align='left')
    style_cell(table.cell(ri, 3), when, font_size=9, fill=fill, align='left')

doc.add_paragraph()

# === РЕКОМЕНДАЦИЯ ===
add_heading(doc, 'Рекомендация', level=1)

add_para(doc,
    'Вариант 2 («5 банок vs 1 банка TITAN») — самый понятный для покупателя, потому что:',
    size=11, space_after=4, line_spacing=1.4)
add_bullet(doc, 'Не требует вычислений — 5 vs 1, мозг считает мгновенно')
add_bullet(doc, 'Показывает конкретную экономию — 2 000 ₽, а не абстрактные «49 ₽/год»')
add_bullet(doc, 'Объясняет WHY — «почему TITAN дешевле? Потому что не нужно перекрашивать»')
add_bullet(doc, 'Работает для любого объёма и цвета — можно подставить любые цифры')

add_para(doc,
    'Фраза «Красишь 1 раз — служит 10 лет» — отличная, оставляем. '
    'Но она должна быть ВЫВОДОМ, а не единственным аргументом. '
    'Перед ней нужен расчёт, который делает вывод очевидным.',
    size=11, bold=True, color=COL_PRIMARY, space_after=10, line_spacing=1.4)

# === ЗАДАЧА ДИЗАЙНЕРУ ===
add_heading(doc, 'Задача дизайнеру', level=1)

add_para(doc, 'Что нужно изменить в текущих вариантах:', size=11, bold=True, color=COL_ACCENT, space_after=4)
add_bullet(doc, 'Убрать пустое поле «___/ГОД» — цена должна быть заполнена конкретной цифрой', bold=True)
add_bullet(doc, 'Добавить ЯКОРЬ — цену ПФ-115 за тот же период (или «5 банок vs 1 банка»)', bold=True)
add_bullet(doc, 'Убрать «Цена за год» как заголовок — заменить на «10 лет без перекраски» или «Экономия за 10 лет»', bold=True)
add_bullet(doc, 'Показать конкретные цифры — 494 ₽, 49 ₽/год, 2 000 ₽ экономии', bold=True)

doc.save(OUT)
print(f'OK: {OUT}')
import os
print(f'Размер: {os.path.getsize(OUT) / 1024:.1f} KB')
