"""
Справка по Central Asia Coatings Show — Астана, Казахстан.
1 страница A4.
"""
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import os

doc = Document()

# === ЕДИНЫЕ СТИЛИ ===
style = doc.styles['Normal']
font = style.font
font.name = 'Calibri'
font.size = Pt(10)
rPr = style.element.get_or_add_rPr()
rFonts = rPr.find(qn('w:rFonts'))
if rFonts is None:
    rFonts = OxmlElement('w:rFonts')
    rPr.insert(0, rFonts)
rFonts.set(qn('w:ascii'), 'Calibri')
rFonts.set(qn('w:hAnsi'), 'Calibri')

# Поля — компактные для 1 страницы
for section in doc.sections:
    section.top_margin = Cm(1.2)
    section.bottom_margin = Cm(1.2)
    section.left_margin = Cm(1.8)
    section.right_margin = Cm(1.5)

BLUE = RGBColor(0x1F, 0x4E, 0x79)
RED = RGBColor(0xC0, 0x00, 0x00)
GREEN = RGBColor(0x37, 0x56, 0x23)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
BLACK = RGBColor(0x1A, 0x1A, 0x1A)
GRAY = RGBColor(0x59, 0x59, 0x59)
ROW_ALT = "F2F7FC"

# === Утилиты ===
def set_cell_shading(cell, color_hex):
    shading = OxmlElement('w:shd')
    shading.set(qn('w:fill'), color_hex)
    shading.set(qn('w:val'), 'clear')
    cell._element.get_or_add_tcPr().append(shading)

def set_cell_borders(cell, color='BFBFBF', size='4'):
    tcPr = cell._element.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for bn in ['top', 'left', 'bottom', 'right']:
        b = OxmlElement(f'w:{bn}')
        b.set(qn('w:val'), 'single')
        b.set(qn('w:sz'), size)
        b.set(qn('w:color'), color)
        tcBorders.append(b)
    tcPr.append(tcBorders)

def add_cell(cell, text, bold=False, size=9, color=BLACK, align='left', fill=None):
    cell.text = ''
    p = cell.paragraphs[0]
    if align == 'center': p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    elif align == 'right': p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after = Pt(1)
    lines = text.split('\n')
    for i, line in enumerate(lines):
        if i > 0:
            run = p.add_run()
            run.add_break()
        run = p.add_run(line)
        run.font.name = 'Calibri'
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.color.rgb = color
    if fill: set_cell_shading(cell, fill)
    set_cell_borders(cell)

def add_heading(text, size=11, color=BLUE, space_before=8, space_after=2):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    run = p.add_run(text)
    run.font.name = 'Calibri'
    run.font.size = Pt(size)
    run.font.bold = True
    run.font.color.rgb = color
    return p

def add_text(text, size=9, color=BLACK, bold=False, align='left', space_after=2):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(space_after)
    if align == 'center': p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(text)
    run.font.name = 'Calibri'
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    return p

def add_hr():
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    pPr = p._element.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '6')
    bottom.set(qn('w:space'), '1')
    bottom.set(qn('w:color'), '1F4E79')
    pBdr.append(bottom)
    pPr.append(pBdr)

# ====================== ДОКУМЕНТ ======================

# === ЗАГОЛОВОК ===
add_text('Central Asia Coatings Show', size=15, bold=True, color=BLUE, align='center', space_after=1)
add_text('4-я Международная выставка лакокрасочных материалов, покрытий, сырья и оборудования', 
         size=9, color=GRAY, align='center', space_after=4)

add_hr()

# === КЛЮЧЕВЫЕ ДАННЫЕ (таблица 2×4) ===
table = doc.add_table(rows=2, cols=4)
table.alignment = WD_TABLE_ALIGNMENT.CENTER
for i in range(4):
    table.rows[0].cells[i].width = Cm(4.3)
    table.rows[1].cells[i].width = Cm(4.3)

key_data = [
    [('ДАТЫ', '27-29 января 2027'), ('МЕСТО', 'IEC EXPO, Астана'), ('ПЛОЩАДЬ', '20 000 м²'), ('ПОСЕТИТЕЛИ', '3 000+ чел.')],
]

# Row 0: headers
for ci, (label, val) in enumerate(key_data[0]):
    add_cell(table.rows[0].cells[ci], f'{label}\n{val}', bold=True, size=9, color=WHITE, align='center', fill='1F4E79')

# Row 1: secondary stats
secondary = [
    ('СТРАНЫ', '10+'), ('ЭКСПОНЕНТЫ', '100+ (из 15 стран)'), ('СПИКЕРЫ', '25 в бизнес-программе'), ('ИСТОРИЯ', 'более 5 лет, 4-я выставка'),
]
for ci, (label, val) in enumerate(secondary):
    add_cell(table.rows[1].cells[ci], f'{label}\n{val}', bold=False, size=8, color=GRAY, align='center', fill=ROW_ALT)

doc.add_paragraph().paragraph_format.space_after = Pt(2)

# === ЧАСЫ РАБОТЫ ===
add_heading('Часы работы', size=10, space_before=4, space_after=1)
add_text('27 января (ср): 10:00–18:00  ·  28 января (чт): 10:00–18:00  ·  29 января (пт): 10:00–16:00', size=9, space_after=1)
add_text('Вход для специалистов отрасли. Онлайн-регистрация на сайте centralasiacoatingshow.com', size=8, color=GRAY, space_after=2)

# === АДРЕС ===
add_heading('Адрес', size=10, space_before=2, space_after=1)
add_text('МВЦ «EXPO» (IEC Astana Expo), 53/1 пр. Мәңгілік Ел (Mangilik El), Астана, 020000, Казахстан', size=9, space_after=2)

# === О ВЫСТАВКЕ ===
add_heading('О выставке', size=10, space_before=2, space_after=1)
add_text(
    'Central Asia Coatings Show — крупнейшая в Центральной Азии международная B2B-выставка лакокрасочных материалов, '
    'сырья, оборудования и технологий для их производства. Выставка проводится ежегодно с 2021 года, поддерживается '
    'Министерством промышленности и строительства РК, Национальной палатой «Атамекен» и Казахстанским союзом химической '
    'промышленности. В 2026 году выставка прошла совместно с United Chemical Show и объединила 4 000+ посетителей, '
    '100+ экспонентов из 15 стран.', size=9, space_after=2)

# === СЕКЦИИ ВЫСТАВКИ ===
add_heading('Секции выставки', size=10, space_before=2, space_after=1)

sections_data = [
    ['Сырьё и материалы', 'Смолы, пигменты, наполнители, растворители, добавки, отвердители, полимеры'],
    ['Готовая продукция', 'Порошковые, антикоррозийные, строительные, промышленные, автомобильные ЛКМ, клеи, герметики'],
    ['Оборудование', 'Производство ЛКМ, подготовка поверхностей, нанесение, сушка, дозирование, колеровочное оборудование'],
    ['Упаковка и новые технологии', 'Фасовка, этикетирование, экологичные покрытия, научные разработки, переработка отходов'],
]

table = doc.add_table(rows=len(sections_data), cols=2)
table.alignment = WD_TABLE_ALIGNMENT.CENTER
for ri, (cat, desc) in enumerate(sections_data):
    table.rows[ri].cells[0].width = Cm(3.5)
    table.rows[ri].cells[1].width = Cm(13.7)
    add_cell(table.rows[ri].cells[0], cat, bold=True, size=8, color=BLUE, fill=ROW_ALT if ri % 2 == 0 else None)
    add_cell(table.rows[ri].cells[1], desc, size=8, fill=ROW_ALT if ri % 2 == 0 else None)

doc.add_paragraph().paragraph_format.space_after = Pt(1)

# === ЦЕЛЕВАЯ АУДИТОРИЯ ===
add_heading('Целевая аудитория', size=10, space_before=2, space_after=1)
add_text(
    'Строительная отрасль, химическая и нефтегазовая промышленность, деревообработка и мебельное производство, '
    'автомобильная промышленность, авиа- и судостроение, производители ЛКМ, дистрибьюторы, инженеры-технологи.', 
    size=9, space_after=2)

# === КЛЮЧЕВЫЕ ЭКСПОНЕНТЫ (2026) ===
add_heading('Ключевые участники выставки 2026', size=10, space_before=2, space_after=1)

exhibitors = [
    ['Страна', 'Компания', 'Профиль'],
    ['Казахстан', 'SP Bely Dom LLP', 'Производитель ЛКМ (Иран-Казахстан), 40 000 т/год'],
    ['Казахстан', 'ASIAN PAINTS LLP', 'Эмали, грунтовки, растворители, дорожная разметка'],
    ['Казахстан', 'ALMAMIX', 'Декоративные штукатурки, краски, покрытия'],
    ['Казахстан', 'KazSmartCompany LLP', 'Промышленное оборудование для нефтегаза'],
    ['Китай', 'Hebei Haoshuo Chemical', 'Эфиры целлюлозы (HPMC, MHEC, HEC)'],
    ['Китай', 'Farfly', 'Оборудование: смесители, диспергаторы, мельницы'],
    ['Китай', 'Shandong Boctok Chemical', 'Сырьё для ЛКМ (USI Chemical)'],
    ['Глобально', 'Tintaly', 'Колеровочные системы и ПО'],
    ['Глобально', 'Rheonics', 'Системы контроля вязности (Швейцария)'],
    ['Россия', 'НПО Лакокраспокрытие', 'Промышленные покрытия с 1960 г.'],
]

table = doc.add_table(rows=len(exhibitors), cols=3)
table.alignment = WD_TABLE_ALIGNMENT.CENTER
for ri, row_data in enumerate(exhibitors):
    table.rows[ri].cells[0].width = Cm(2.5)
    table.rows[ri].cells[1].width = Cm(4.5)
    table.rows[ri].cells[2].width = Cm(10.2)
    if ri == 0:
        for ci, val in enumerate(row_data):
            add_cell(table.rows[ri].cells[ci], val, bold=True, size=8, color=WHITE, align='center', fill='1F4E79')
    else:
        fill = ROW_ALT if ri % 2 == 0 else None
        add_cell(table.rows[ri].cells[0], row_data[0], bold=True, size=8, color=BLUE, fill=fill, align='center')
        add_cell(table.rows[ri].cells[1], row_data[1], bold=True, size=8, fill=fill)
        add_cell(table.rows[ri].cells[2], row_data[2], size=8, fill=fill)

doc.add_paragraph().paragraph_format.space_after = Pt(1)

# === ПАРТНЁРЫ И ПОДДЕРЖКА ===
add_heading('Поддержка и партнёры', size=10, space_before=2, space_after=1)
add_text(
    'Министерство промышленности и строительства РК · Национальная палата предпринимателей «Атамекен» · '
    'Акимат г. Астаны · Казахстанский союз химической промышленности · Ассоциация производителей ЛКМ «Союзкраска» · '
    'Ассоциация химической промышленности KAZХИМИЯ · Союз строителей Казахстана', 
    size=8, color=GRAY, space_after=2)

# === КОНТАКТЫ ===
add_heading('Контакты', size=10, space_before=2, space_after=1)

contacts_data = [
    ['Сайт', 'centralasiacoatingshow.com', 'Оргкомитет (KZ)', '+90 (536) 810 68 98 (WhatsApp)'],
    ['Офис в Астане', '54A Syganak St, Business Centre A, Office 203', 'Совместно с', 'United Chemical Show'],
]
table = doc.add_table(rows=2, cols=4)
table.alignment = WD_TABLE_ALIGNMENT.CENTER
for ri, row_data in enumerate(contacts_data):
    for ci, val in enumerate(row_data):
        table.rows[ri].cells[ci].width = Cm(4.3)
        is_label = ci % 2 == 0
        add_cell(table.rows[ri].cells[ci], val, bold=is_label, size=8, 
                 color=BLUE if is_label else BLACK, fill=ROW_ALT if ri == 0 else None)

# === Финальная строка ===
add_hr()
add_text('Справка подготовлена по данным официального сайта centralasiacoatingshow.com (октябрь 2026)', 
         size=7, color=GRAY, align='center', space_after=0)

# === СОХРАНИТЬ ===
out_path = "/home/z/my-project/download/Справка_Central_Asia_Coatings_Show.docx"
doc.save(out_path)
print(f"✓ Файл сохранён: {out_path}")
print(f"  Размер: {os.path.getsize(out_path):,} байт")
