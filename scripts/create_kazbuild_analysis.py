#!/usr/bin/env python3
"""
Word-документ: сводка по участникам павильона 10 KazBuild 2026.
Анализ: кто из участников связан с ЛКМ, кто — нет.
"""

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

OUT = '/home/z/my-project/download/KazBuild_2026_павильон_10_анализ.docx'

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
    if valign == 'center':
        from docx.enum.table import WD_ALIGN_VERTICAL
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

# === ЗАГОЛОВОК ===
add_para(doc, 'KazBuild 2026 · Алматы, Казахстан', size=10, bold=True, color=COL_ACCENT,
         align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
add_para(doc, 'Павильон 10 — анализ участников', size=18, bold=True, color=COL_PRIMARY,
         align=WD_ALIGN_PARAGRAPH.CENTER, space_after=4)
add_para(doc, 'Стенд Эмпилс: 10-679 (сектор E, центральная нижняя часть павильона)',
         size=11, color=COL_GRAY, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=20)

# === РАЗДЕЛ 1: ПРОИЗВОДИТЕЛИ ЛКМ ===
add_heading(doc, '1. Производители и дистрибьюторы ЛКМ на павильоне', level=1)

add_para(doc,
    'Из ~120 компаний-участников павильона 10 выявлено 2 производителя лакокрасочных '
    'материалов и 1 компания, требующая уточнения. Эмпилс — один из немногих производителей '
    'красок на выставке, что создаёт конкурентное преимущество.',
    size=11, space_after=10, line_spacing=1.4)

# Таблица ЛКМ
table = doc.add_table(rows=4, cols=5)
table.alignment = WD_TABLE_ALIGNMENT.CENTER
table.autofit = False
col_widths = [Cm(3.5), Cm(1.8), Cm(2.2), Cm(2.0), Cm(7.5)]
for i, w in enumerate(col_widths):
    table.columns[i].width = w

headers = ['Компания', 'Стенд', 'Страна', 'ЛКМ?', 'Описание']
for i, h in enumerate(headers):
    cell = table.cell(0, i)
    cell.width = col_widths[i]
    style_cell(cell, h, font_size=10, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF), fill='1E3A5F', align='center')

painters = [
    ('FAB PAINT', '10-696', 'Азербайджан', 'ДА', 'Производитель красок, лаков, растворителей, эпоксидных смол. Бренд FAB Boya. Основан в 2001 г., г. Баку. Воднодисперсионные, синтетические, целлюлозные краски. Прямой конкурент по категории.', 'FEE2E2'),
    ('GYNEX CHEMALEX', '10-610', 'Словакия', 'ДА', 'Производитель строительной и промышленной химии. Краски и лаки под брендом Cyklon (базовые/металлические цвета, защитные, маркировочные). Также герметики, клеи, ПУ-пены. С 1997 г.', 'FEE2E2'),
    ('DECOPMIX', '10-635', 'Не определена', 'Возможно', 'По названию — декоративные штукатурки/покрытия. Точных данных нет. Похожие бренды: DECOMIX (РФ, СПб — мозаичная/фактурная штукатурка). Уточнить на стенде.', 'FEF9C3'),
]

for ri, (name, stand, country, paint, desc, fill) in enumerate(painters, 1):
    style_cell(table.cell(ri, 0), name, font_size=9, bold=True, color=COL_PRIMARY, fill=fill)
    style_cell(table.cell(ri, 1), stand, font_size=9, align='center', fill=fill)
    style_cell(table.cell(ri, 2), country, font_size=9, align='center', fill=fill)
    style_cell(table.cell(ri, 3), paint, font_size=9, bold=True, color=COL_RED if paint == 'ДА' else COL_ACCENT, align='center', fill=fill)
    style_cell(table.cell(ri, 4), desc, font_size=8.5, fill=fill)

doc.add_paragraph()

# === РАЗДЕЛ 2: СОСЕДИ ЭМПИЛС ===
add_heading(doc, '2. Соседи стенда Эмпилс (10-679)', level=1)

add_para(doc, 'Эмпилс находится в секторе E, центральная нижняя часть павильона, у входной стены. Ближайшие соседи:', size=11, space_after=6)

table2 = doc.add_table(rows=6, cols=3)
table2.alignment = WD_TABLE_ALIGNMENT.CENTER
table2.autofit = False
col_widths2 = [Cm(2.5), Cm(4.5), Cm(7.5)]
for i, w in enumerate(col_widths2):
    table2.columns[i].width = w

headers2 = ['Стенд', 'Компания', 'Чем занимается']
for i, h in enumerate(headers2):
    cell = table2.cell(0, i)
    cell.width = col_widths2[i]
    style_cell(cell, h, font_size=10, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF), fill='1E3A5F', align='center')

neighbors = [
    ('10-680 (справа)', 'MARRIONI', 'Межкомнатные двери (Италия)'),
    ('10-678 (слева)', 'RESERVED', 'Свободно — 9 кв.м. (можно арендовать)'),
    ('10-677', 'PHILIPS EASYKEY', 'Системы контроля доступа, умные замки'),
    ('10-682', 'SANATT TEJARAT PARDIS', 'Крупная иранская торговая компания (36 кв.м.)'),
    ('10-617', 'HISSARIA', '— (уточнить на стенде)'),
]
for ri, (stand, name, desc) in enumerate(neighbors, 1):
    fill = 'F9FAFB' if ri % 2 == 1 else 'FFFFFF'
    style_cell(table2.cell(ri, 0), stand, font_size=9, bold=True, color=COL_PRIMARY, fill=fill)
    style_cell(table2.cell(ri, 1), name, font_size=9, bold=True, fill=fill)
    style_cell(table2.cell(ri, 2), desc, font_size=9, fill=fill)

doc.add_paragraph()

# === РАЗДЕЛ 3: КАТЕГОРИИ УЧАСТНИКОВ ===
add_heading(doc, '3. Основные категории участников павильона 10', level=1)

add_para(doc, 'Большинство компаний павильона НЕ связаны с ЛКМ. Основные категории:', size=11, space_after=6)

categories = [
    ('Окна, двери, фасады', 'KAPELLI DVERI, ALUBOND, BOWU WINDOW, ALUMIN STROY, JUNG, PHILIPS EASYKEY'),
    ('Напольные покрытия', 'ALISBA, KINGWINS, MALAYSIAN TIMBER COUNCIL, OAKONCEPT WOOD, KDE UNIDECO (LVT), TIMBERSTONE'),
    ('Стеновые панели', 'GIPSOVINIL (Казахстан — первый завод гипсовиниловых панелей), KAZROOF, WILLVIC, WOLTA, WISKIND'),
    ('Строительная химия (НЕ краски)', 'SOMAFIX (Турция — пена, герметики, клеи), ERA PLAST, ALTER PLAST'),
    ('Регионы России', 'MOSCOW REGION, CHELYABINSK REGION, KRASNOYARSK REGION, NIZHNY NOVGOROD REGION, KRASNODAR REGION'),
    ('Отделочные материалы', 'DE BAGET, VEKS (отделочный камень, Армения), RENDI (ПВХ-плиты, Китай)'),
    ('Прочее', 'Искусственный газон (KAZ SAISI), композитные панели (YARET), профили, крепёж, искусственный газон'),
]

for cat, companies in categories:
    add_bullet(doc, f'{cat}: {companies}', size=10.5)

doc.add_paragraph()

# === РАЗДЕЛ 4: ВЫВОДЫ ===
add_heading(doc, '4. Выводы для Эмпилс', level=1)

add_para(doc, '1. Прямых конкурентов-производителей ЛКМ на павильоне всего 2', size=11, bold=True, color=COL_PRIMARY, space_after=4)
add_para(doc, 'FAB PAINT (Азербайджан) и GYNEX CHEMALEX (Словакия, бренд Cyklon). Эмпилс — один из немногих производителей красок на выставке, что даёт преимущество.',
         size=11, space_after=8, line_spacing=1.4)

add_para(doc, '2. FAB PAINT (10-696) — ближайший по категории', size=11, bold=True, color=COL_PRIMARY, space_after=4)
add_para(doc, 'Стоит посетить стенд, оценить ассортимент и цены. Азербайджанский производитель — вероятно, конкурирует в нижнем/среднем ценовом сегменте. Бренд FAB Boya: воднодисперсионные, синтетические, целлюлозные краски.',
         size=11, space_after=8, line_spacing=1.4)

add_para(doc, '3. GYNEX CHEMALEX (10-610) — европейский производитель', size=11, bold=True, color=COL_PRIMARY, space_after=4)
add_para(doc, 'Бренд Cyklon: краски и лаки (базовые/металлические цвета, защитные, маркировочные). Может быть интересен как партнёр для дистрибуции в РФ, или как конкурент в сегменте спецэмалей.',
         size=11, space_after=8, line_spacing=1.4)

add_para(doc, '4. DECOPMIX (10-635) — уточнить на стенде', size=11, bold=True, color=COL_PRIMARY, space_after=4)
add_para(doc, 'Если это декоративные штукатурки — не прямой конкурент, но смежная категория. Возможен обмен клиентами.',
         size=11, space_after=8, line_spacing=1.4)

add_para(doc, '5. Большинство соседей — не ЛКМ', size=11, bold=True, color=COL_PRIMARY, space_after=4)
add_para(doc, 'Это двери, окна, полы, панели. Эмпилс выделяется как единственный крупный производитель красок в этой зоне павильона. Использовать это в переговорах с посетителями.',
         size=11, space_after=10, line_spacing=1.4)

# Подвал
add_para(doc, '', size=6, space_after=4)
add_para(doc, 'Источник: план павильона 10 KazBuild 2026 (Алматы). Анализ выполнен по визуальному распознаванию плана и веб-поиску по компаниям. Информация по 5 из 10 проверенных компаний подтверждена в официальном списке участников KazBuild 2026 (PDF от 20.07.2026).',
         size=8, color=COL_GRAY, align=WD_ALIGN_PARAGRAPH.LEFT, space_before=10)

doc.save(OUT)
print(f'OK: {OUT}')
import os
print(f'Размер: {os.path.getsize(OUT) / 1024:.1f} KB')
