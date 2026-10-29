"""
TITAN — Email-письмо для региональной торговой команды
Запрос обратной связи по 5 POSm-материалам: листовка, шелфтокер, воблер, ценник, планограмма.
Лаконичный формат с таблицей параметров.
"""
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

ORANGE = RGBColor(0xF4, 0x76, 0x29)
DARK   = RGBColor(0x31, 0x31, 0x31)
GREY   = RGBColor(0x6B, 0x72, 0x80)
WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
BLACK_TEXT = RGBColor(0x1A, 0x1F, 0x36)

doc = Document()

style = doc.styles['Normal']
style.font.name = 'Arial'
style.font.size = Pt(11)
style.element.rPr.rFonts.set(qn('w:eastAsia'), 'Arial')

for section in doc.sections:
    section.top_margin = Cm(2)
    section.bottom_margin = Cm(2)
    section.left_margin = Cm(2.2)
    section.right_margin = Cm(2.2)
    section.page_height = Cm(29.7)
    section.page_width = Cm(21.0)

# ---------- утилиты ----------
def set_spacing(p, before=0, after=6, line=1.3):
    pf = p.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.line_spacing = line

def add_para(text, size=11, bold=False, color=None, italic=False, align='left', after=6, before=0):
    p = doc.add_paragraph()
    set_spacing(p, before=before, after=after, line=1.3)
    p.alignment = {'center': WD_ALIGN_PARAGRAPH.CENTER, 'right': WD_ALIGN_PARAGRAPH.RIGHT,
                   'left': WD_ALIGN_PARAGRAPH.LEFT, 'justify': WD_ALIGN_PARAGRAPH.JUSTIFY}.get(align, WD_ALIGN_PARAGRAPH.LEFT)
    r = p.add_run(text)
    r.font.size = Pt(size)
    r.bold = bold
    r.italic = italic
    if color:
        r.font.color.rgb = color
    r.font.name = 'Arial'
    return p

def add_bullet(text, size=11, bold=False):
    p = doc.add_paragraph(style='List Bullet')
    set_spacing(p, before=0, after=3, line=1.25)
    p.paragraph_format.left_indent = Cm(0.6)
    r = p.runs[0] if p.runs else p.add_run()
    r.text = text
    r.font.size = Pt(size)
    r.font.name = 'Arial'
    if bold:
        r.bold = True
    return p

def set_cell_bg(cell, color_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), color_hex)
    tcPr.append(shd)

def set_cell_borders(cell, color='D0D0D0', sz='4'):
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for edge in ('top', 'left', 'bottom', 'right'):
        el = OxmlElement(f'w:{edge}')
        el.set(qn('w:val'), 'single')
        el.set(qn('w:sz'), sz)
        el.set(qn('w:color'), color)
        tcBorders.append(el)
    tcPr.append(tcBorders)

def add_table(headers, rows, col_widths, font_size=10, header_bg='313131'):
    tbl = doc.add_table(rows=1 + len(rows), cols=len(headers))
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    for i, w in enumerate(col_widths):
        for cell in tbl.columns[i].cells:
            cell.width = Cm(w)
    for i, h in enumerate(headers):
        c = tbl.rows[0].cells[i]
        set_cell_bg(c, header_bg)
        set_cell_borders(c, color='313131', sz='4')
        c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_spacing(p, before=2, after=2, line=1.2)
        r = p.add_run(h)
        r.bold = True
        r.font.size = Pt(font_size)
        r.font.color.rgb = WHITE
        r.font.name = 'Arial'
    for ri, row in enumerate(rows):
        for ci, val in enumerate(row):
            c = tbl.rows[ri + 1].cells[ci]
            if ri % 2 == 1:
                set_cell_bg(c, 'F8F4F0')
            set_cell_borders(c, color='D0D0D0', sz='4')
            c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = c.paragraphs[0]
            set_spacing(p, before=2, after=2, line=1.25)
            r = p.add_run(str(val))
            r.font.size = Pt(font_size)
            r.font.name = 'Arial'
            if ci == 0:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                r.bold = True
                r.font.color.rgb = ORANGE
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p = doc.add_paragraph()
    set_spacing(p, before=0, after=4, line=1.0)
    return tbl

def add_callout(title, body, bg='FFF4ED', bar='F47629'):
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    cell = tbl.cell(0, 0)
    cell.width = Cm(16.5)
    set_cell_bg(cell, bg)
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for edge, c in (('top', bg), ('right', bg), ('bottom', bg), ('left', bar)):
        el = OxmlElement(f'w:{edge}')
        el.set(qn('w:val'), 'single')
        el.set(qn('w:sz'), '36' if edge == 'left' else '4')
        el.set(qn('w:color'), c)
        tcBorders.append(el)
    tcPr.append(tcBorders)
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    if title:
        p = cell.paragraphs[0]
        set_spacing(p, before=2, after=2, line=1.2)
        r = p.add_run(title)
        r.bold = True
        r.font.size = Pt(11)
        r.font.color.rgb = ORANGE
        r.font.name = 'Arial'
    if body:
        p2 = cell.add_paragraph()
        set_spacing(p2, before=0, after=2, line=1.3)
        r2 = p2.add_run(body)
        r2.font.size = Pt(10)
        r2.font.color.rgb = DARK
        r2.font.name = 'Arial'
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for edge, val in (('top', '120'), ('left', '200'), ('bottom', '120'), ('right', '200')):
        el = OxmlElement(f'w:{edge}')
        el.set(qn('w:w'), val)
        el.set(qn('w:type'), 'dxa')
        tcMar.append(el)
    tcPr.append(tcMar)
    p = doc.add_paragraph()
    set_spacing(p, before=0, after=4, line=1.0)
    return tbl

def hr():
    p = doc.add_paragraph()
    set_spacing(p, before=2, after=6, line=1.0)
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '6')
    bottom.set(qn('w:color'), 'F47629')
    pBdr.append(bottom)
    pPr.append(pBdr)

print("Utils ready")

# ============================================================
# ТЕМА ПИСЬМА
# ============================================================
p = doc.add_paragraph()
set_spacing(p, before=0, after=2, line=1.2)
r = p.add_run('Тема: ')
r.bold = True
r.font.size = Pt(11)
r.font.color.rgb = DARK
r.font.name = 'Arial'
r2 = p.add_run('TITAN POSm — просьба дать обратную связь по макетам (до 28.08)')
r2.font.size = Pt(11)
r2.font.color.rgb = ORANGE
r2.bold = True
r2.font.name = 'Arial'

hr()

# ============================================================
# ПРИВЕТСТВИЕ
# ============================================================
add_para('Коллеги, добрый день!', size=12, bold=True, after=8)

# ============================================================
# ОСНОВНОЙ ТЕКСТ
# ============================================================
add_para(
    'Дизайнеры подготовили финальные макеты 5 POSm-материалов TITAN: '
    'листовка, шелфтокер, воблер, брендированный ценник и планограмма. '
    'Вам с этими материалами работать в точках — прошу дать короткую обратную связь.',
    after=8
)

add_para(
    'Макеты во вложении. Срок обратной связи — до 28 августа 2026.',
    bold=True, after=10
)

# ============================================================
# ТАБЛИЦА POSm С РАЗМЕРАМИ И МАТЕРИАЛАМИ
# ============================================================
add_para('Что проверяем', size=13, bold=True, color=ORANGE, before=4, after=6)

add_table(
    headers=['№', 'Материал', 'Размер (Ш×В)', 'Материал', 'Покрытие', 'Где ставится'],
    rows=[
        ['1', 'Листовка (двусторонняя, А4)',
         '210 × 297 мм',
         'Бумага мелованная 300 г/м²',
         'Ламинация матовая, 2 стороны',
         'Зона B2B / прилавок'],
        ['2', 'Шелфтокер (на торец полки)',
         '300 × 100 мм',
         'Картон мелованный 300 г/м²',
         'Ламинация глянцевая',
         'Край полки с TITAN'],
        ['3', 'Воблер (на гибкой ножке)',
         '120 × 80 мм',
         'Картон мелованный 350 г/м²',
         'Ламинация матовая, 2 стороны',
         'Над верхней полкой TITAN'],
        ['4', 'Брендированный ценник',
         '70 × 40 мм',
         'Картон мелованный 250 г/м²',
         'Ламинация глянцевая (лицо)',
         'Рядом с банками TITAN'],
        ['5', 'Планограмма выкладки (А4 горизонт.)',
         '297 × 210 мм',
         'Бумага мелованная 200 г/м²',
         'Без ламинации',
         'Документ для ТП / не POSm'],
    ],
    col_widths=[0.7, 4.0, 2.5, 3.5, 3.0, 2.8],
    font_size=9
)

# ============================================================
# ВОПРОСЫ ДЛЯ ОБРАТНОЙ СВЯЗИ
# ============================================================
add_para('Вопросы к вам', size=13, bold=True, color=ORANGE, before=8, after=4)

add_bullet('Понятна ли вам планограмма выкладки TITAN?')
add_bullet('Какой из трёх предложенных вариантов макета (см. вложения 1, 2, 3) вам больше нравится?')
add_bullet('Понятны ли тексты и слоганы на POSm — особенно «Красишь 1 раз — служит 10 лет»?')
add_bullet('Подходят ли размеры под реальные полки ваших точек — особенно шелфтокер 300 мм?')
add_bullet('Есть ли что-то, что мы упустили, но эта информация точно должна быть на макетах?')
add_bullet('Считывается ли QR-код на шелфтокере камерой вашего телефона с расстояния 30–50 см?')

# ============================================================
# ДЕДЛАЙН
# ============================================================
add_callout('Дедлайн: 28 августа 2026 г.',
            'После 28.08 макеты уходят в печать. Правки после этой даты уже не принимаются — '
            'они войдут в следующую итерацию (через 2 месяца после пилота).')

# ============================================================
# ПОДПИСЬ
# ============================================================
add_para('', after=6, before=8)
add_para('Спасибо за помощь!', size=11, bold=True, after=4)
add_para('С уважением,', size=11, after=2)
add_para('Команда бренда TITAN / ЗАО «Эмпилс»', size=11, bold=True, color=DARK, after=2)
add_para('телефон: +7 (863) 200-00-00  ·  e-mail: titan@empils.ru', size=10, color=GREY, after=2)

# ---------- Финал ----------
hr()
add_para('TITAN POSm  ·  Запрос обратной связи  ·  Август 2026',
         size=9, italic=True, color=GREY, align='center', after=2)
add_para('Во вложении: 5 изображений макетов + ТЗ_POSM_TITAN_SHORT.docx с техническими параметрами.',
         size=9, italic=True, color=GREY, align='center')

# ---------- Сохранение ----------
for table in doc.tables:
    if len(table.rows) > 0:
        tr = table.rows[0]._tr
        trPr = tr.get_or_add_trPr()
        tblHeader = OxmlElement('w:tblHeader')
        tblHeader.set(qn('w:val'), 'true')
        trPr.append(tblHeader)
        for row in table.rows:
            trPr = row._tr.get_or_add_trPr()
            cantSplit = OxmlElement('w:cantSplit')
            cantSplit.set(qn('w:val'), 'true')
            trPr.append(cantSplit)

out = '/home/z/my-project/download/TITAN_POSm_Email_команде.docx'
doc.save(out)
print(f'OK saved: {out}')
import os
print(f'Size: {os.path.getsize(out)} bytes')
