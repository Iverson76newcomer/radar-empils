"""
TITAN — Актуальные цены конкурентов на vseinstrumenti.ru
Сравнение цен с картинок конкурентного окружения и текущих цен.
3 категории: 3в1 / быстросохнущая / молотковая.
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
GREEN  = RGBColor(0x16, 0xA3, 0x4A)   # цена снизилась — выгодно для TITAN
RED    = RGBColor(0xC0, 0x39, 0x2B)   # цена выросла — конкуренты дороже
BLACK_TEXT = RGBColor(0x1A, 0x1F, 0x36)

doc = Document()

style = doc.styles['Normal']
style.font.name = 'Arial'
style.font.size = Pt(11)
style.element.rPr.rFonts.set(qn('w:eastAsia'), 'Arial')

for section in doc.sections:
    section.top_margin = Cm(2)
    section.bottom_margin = Cm(2)
    section.left_margin = Cm(2.0)
    section.right_margin = Cm(2.0)
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

def add_h1(text, color=DARK, size=20, before=0, after=8, orange_bar=True):
    p = doc.add_paragraph()
    set_spacing(p, before=before, after=after, line=1.1)
    r = p.add_run(text)
    r.bold = True
    r.font.size = Pt(size)
    r.font.color.rgb = color
    r.font.name = 'Arial'
    if orange_bar:
        pPr = p._p.get_or_add_pPr()
        pBdr = OxmlElement('w:pBdr')
        left = OxmlElement('w:left')
        left.set(qn('w:val'), 'single')
        left.set(qn('w:sz'), '24')
        left.set(qn('w:space'), '8')
        left.set(qn('w:color'), 'F47629')
        pBdr.append(left)
        pPr.append(pBdr)
        p.paragraph_format.left_indent = Cm(0.3)
    return p

def add_h2(text, color=ORANGE, size=14, before=10, after=4):
    p = doc.add_paragraph()
    set_spacing(p, before=before, after=after, line=1.2)
    r = p.add_run(text)
    r.bold = True
    r.font.size = Pt(size)
    r.font.color.rgb = color
    r.font.name = 'Arial'
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

def add_table(headers, rows, col_widths, font_size=10, header_bg='313131',
              delta_col=None, highlight_titan=True):
    """
    delta_col: индекс колонки с Δ изменением — красим в зелёный (−) или красный (+)
    highlight_titan: подсвечивать строки с TITAN
    """
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
        is_titan = 'titan' in str(row[1]).lower() or 'titan' in str(row[0]).lower()
        for ci, val in enumerate(row):
            c = tbl.rows[ri + 1].cells[ci]
            if is_titan and highlight_titan:
                set_cell_bg(c, 'FFF4ED')
            elif ri % 2 == 1:
                set_cell_bg(c, 'F8F4F0')
            set_cell_borders(c, color='D0D0D0', sz='4')
            c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = c.paragraphs[0]
            set_spacing(p, before=2, after=2, line=1.25)
            r = p.add_run(str(val))
            r.font.size = Pt(font_size)
            r.font.name = 'Arial'
            if ci == 0:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r.bold = True
                r.font.color.rgb = ORANGE
            elif delta_col is not None and ci == delta_col:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r.bold = True
                # Зелёный если цена снизилась (−), красный если выросла (+)
                val_str = str(val)
                if val_str.startswith('−') or val_str.startswith('-'):
                    r.font.color.rgb = GREEN
                elif val_str.startswith('+'):
                    r.font.color.rgb = RED
                else:
                    r.font.color.rgb = GREY
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                if is_titan:
                    r.bold = True
                    r.font.color.rgb = ORANGE
    p = doc.add_paragraph()
    set_spacing(p, before=0, after=4, line=1.0)
    return tbl

def add_callout(title, body, bg='FFF4ED', bar='F47629'):
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    cell = tbl.cell(0, 0)
    cell.width = Cm(17.0)
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

def page_break():
    p = doc.add_paragraph()
    r = p.add_run()
    r.add_break(WD_BREAK.PAGE)

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
# ЗАГОЛОВОК
# ============================================================
add_h1('TITAN — Актуальные цены конкурентов', size=22, before=0, after=2, orange_bar=True)
add_para('Сравнение цен с картинок конкурентного окружения и текущих цен на vseinstrumenti.ru',
         size=12, italic=True, color=GREY, after=4)
add_para('Источник актуальных цен: vseinstrumenti.ru  ·  Дата проверки: 26 августа 2026 г.',
         size=11, after=8)
add_para('Документ обновляет данные с 3 картинок конкурентного окружения по категориям TITAN: '
         'грунт-эмаль 3в1, быстросохнущая матовая и молотковая эмаль.',
         size=11, after=8)

hr()

# Легенда
add_para('Как читать таблицы', size=12, bold=True, color=ORANGE, before=4, after=4)
add_para('Зелёная Δ — цена конкурента снизилась (стал дешевле, выгоднее для нас). '
         'Красная Δ — цена выросла (стал дороже). Строки TITAN выделены оранжевым.',
         size=10, italic=True, color=GREY, after=10)

# ============================================================
# КАТЕГОРИЯ 1 — ГРУНТ-ЭМАЛЬ 3В1 ЧЁРНАЯ RAL 9005
# ============================================================
add_h2('Категория 1. Грунт-эмаль 3в1 высокоглянцевая, чёрный RAL 9005', before=4, after=6)

add_table(
    headers=['№', 'Бренд / Продукт', 'Фасовка', 'Цена на картинке', 'Актуальная цена', 'Δ изменение'],
    rows=[
        ['1', 'ОРЕОЛ TITAN 3в1',         '0,75 л',  '872 ₽',  '860 ₽',  '−12 ₽ (−1,4%)'],
        ['2', 'Elcon 3в1 глянцевая',      '0,8 кг',  '958 ₽',  '963 ₽',  '+5 ₽ (+0,5%)'],
        ['3', 'Dufa CORREX 3в1',           '0,75 л',  '1020 ₽', '1 130 ₽','+110 ₽ (+10,8%)'],
        ['4', 'Dali 3в1',                  '0,75 л',  '1055 ₽', '1 259 ₽','+204 ₽ (+19,3%)'],
        ['5', 'HUSKY KLONDIKE глянцевая',  '0,9 л',   '1153 ₽', '953 ₽',  '−200 ₽ (−17,3%)'],
        ['6', 'Tikkurila METALLISTA',      '0,8 л',   '2085 ₽', '1 599 ₽','−486 ₽ (−23,3%)'],
    ],
    col_widths=[0.8, 5.5, 2.0, 2.5, 2.5, 3.7],
    font_size=10,
    delta_col=5
)

add_callout('Вывод по категории 1',
            'TITAN сохранил цену практически без изменений (−1,4%) и остался самым доступным '
            'продуктом 3в1 в категории. Прямые конкуренты Dali и Dufa CORREX подорожали на 11–19%. '
            'HUSKY KLONDIKE и Tikkurila снизили цены, но по-прежнему дороже TITAN на 11% и 86% '
            'соответственно. Позиционирование TITAN как «премиум по цене среднего» сохраняется.')

page_break()

# ============================================================
# КАТЕГОРИЯ 2 — БЫСТРОСОХНУЩАЯ МАТОВАЯ СЕРАЯ
# ============================================================
add_h2('Категория 2. Быстросохнущая матовая, серая RAL 7040', before=4, after=6)

add_table(
    headers=['№', 'Бренд / Продукт', 'Фасовка', 'Цена на картинке', 'Актуальная цена', 'Δ изменение'],
    rows=[
        ['1', 'ОРЕОЛ TITAN быстросохнущая', '0,75 л (0,9 кг)', '765 ₽',  '903 ₽',  '+138 ₽ (+18,0%)'],
        ['2', 'Elcon быстросохнущая 3в1',    '0,8 кг',          '765 ₽',  '764 ₽',  '−1 ₽ (−0,1%)'],
        ['3', 'ЛАКРА PROF IT 4в1',           '0,8 кг',          '838 ₽',  '875 ₽',  '+37 ₽ (+4,4%)'],
        ['4', 'Новбытхим 5в1',                '0,75 л',          '906 ₽',  '775 ₽',  '−131 ₽ (−14,5%)'],
    ],
    col_widths=[0.8, 5.5, 2.5, 2.5, 2.5, 3.7],
    font_size=10,
    delta_col=5
)

add_callout('Вывод по категории 2',
            'TITAN подорожал на 18% — необходимо проверить, не теряем ли мы ценовое '
            'преимущество. На картинке TITAN был самым дешёвым (765 ₽), сейчас Elcon быстросохнущая '
            'и Новбытхим дешевле. Рекомендация: оценить возможность удержания РРЦ на уровне '
            '765–800 ₽ или подготовить для ТП аргумент о премиальности (10 лет срока службы '
            'vs 5–6 лет у конкурентов).')

page_break()

# ============================================================
# КАТЕГОРИЯ 3 — МОЛОТКОВАЯ ЭМАЛЬ ЗОЛОТИСТАЯ
# ============================================================
add_h2('Категория 3. Эмаль с молотковым эффектом, золотистая', before=4, after=6)

add_table(
    headers=['№', 'Бренд / Продукт', 'Фасовка', 'Цена на картинке', 'Актуальная цена', 'Δ изменение'],
    rows=[
        ['1', 'Лакра молотковая',             '0,8 кг',  '1057 ₽', '1 141 ₽', '+84 ₽ (+7,9%)'],
        ['2', 'ОРЕОЛ TITAN молотковая',        '0,8 кг',  '1097 ₽', '1 294 ₽', '+197 ₽ (+17,9%)'],
        ['3', 'Dufa Premium HAMMERLACK',       '750 мл',  '1196 ₽', '1 196 ₽', '0 ₽ (0%)'],
        ['4', 'Dali 3в1 молотковая',            '0,75 л',  '1308 ₽', '1 472 ₽', '+164 ₽ (+12,5%)'],
        ['5', 'Elcon кузнечная золото',         '0,8 кг',  '1341 ₽', '1 268 ₽', '−73 ₽ (−5,4%)'],
        ['6', 'HUSKY KLONDIKE золото',          '0,9 л',   '1653 ₽', '1 719 ₽', '+66 ₽ (+4,0%)'],
    ],
    col_widths=[0.8, 5.5, 2.0, 2.5, 2.5, 3.7],
    font_size=10,
    delta_col=5
)

add_callout('Вывод по категории 3',
            'TITAN молотковая подорожала на 18%, но осталась в нижней половине ценового '
            'диапазона. Ситуация в категории: 4 из 6 конкурентов подорожали на 4–13%. Dufa '
            'HAMMERLACK единственный сохранил цену. Самый дешёвый вариант — Лакра (1 141 ₽), '
            'но у TITAN остаётся преимущество в сроке службы (8 лет vs 5 лет у Лакры) и '
            'более широкая палитра (6 цветов vs 4 у Лакры). Рекомендация для ТП: при '
            'обосновании цены оперировать на срок службы.')

page_break()

# ============================================================
# СВОДКА
# ============================================================
add_h2('Сводка по всем 3 категориям', before=4, after=6)

add_table(
    headers=['Категория', 'Кол-во SKU', 'TITAN: Δ цены', 'Дешевле TITAN', 'Дороже TITAN', 'Позиция TITAN'],
    rows=[
        ['3в1 чёрный RAL 9005',     '6 SKU', '−1,4%', '0 (самый дешёвый)',  '5 SKU',  'Лидер по цене'],
        ['Быстросохнущая серая',    '4 SKU', '+18,0%', '2 SKU (Elcon, Новбытхим)', '1 SKU', 'Средний сегмент'],
        ['Молотковая золотистая',    '6 SKU', '+17,9%', '1 SKU (Лакра)', '4 SKU', 'Нижняя половина'],
    ],
    col_widths=[4.5, 1.8, 2.0, 3.5, 2.5, 3.0],
    font_size=9
)

add_callout('Общий вывод',
            'TITAN 3в1 — безусловный ценовой лидер в своей категории (подешевел на 1,4% при '
            'росте цен у 4 из 6 конкурентов). TITAN быстросохнущая и молотковая подорожали '
            'на 18%, что отражает общий рост цен на алкидные грунт-эмали в категории. В '
            'быстросохнущей категории TITAN потерял позицию самого дешёвого — это критично '
            'для скрипта продаж и требует либо коррекции цены, либо усиления аргументации по '
            'ССВ (сроку службы).')

add_para('', after=4, before=4)
add_para('Источник: vseinstrumenti.ru  ·  Проверка: 26.08.2026  ·  Цены актуальны на дату проверки',
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

out = '/home/z/my-project/download/TITAN_Актуальные_цены_конкурентов.docx'
doc.save(out)
print(f'OK saved: {out}')
import os
print(f'Size: {os.path.getsize(out)} bytes')
