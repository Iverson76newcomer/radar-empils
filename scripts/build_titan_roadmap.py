"""
TITAN — RoadMap для дистрибьютора (объединённый документ)
Соединяет основной Sales Guide и Дополнение (15 разделов) в пошаговую
последовательность из 23 шагов. Текст исходных документов НЕ меняется —
копируется как XML-deepcopy с сохранением форматирования.
"""
import copy
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ---------- Цвета ----------
ORANGE = RGBColor(0xF4, 0x76, 0x29)
DARK   = RGBColor(0x31, 0x31, 0x31)
WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
GREY   = RGBColor(0x6B, 0x6B, 0x6B)
LIGHT  = RGBColor(0xFA, 0xF6, 0xF2)

# ---------- Источники ----------
MAIN_PATH = '/home/z/my-project/upload/TITAN_Sales_Guide_дистрибьюторам_21.08.2026.docx'
SUPP_PATH = '/home/z/my-project/upload/TITAN_Sales_Guide_Дополнение_15_разделов.docx'
OUT_PATH  = '/home/z/my-project/download/TITAN_RoadMap_дистрибьютора.docx'

main_doc = Document(MAIN_PATH)
supp_doc = Document(SUPP_PATH)
new_doc = Document()

# ---------- Базовый стиль нового документа ----------
style = new_doc.styles['Normal']
style.font.name = 'Arial'
style.font.size = Pt(11)
style.element.rPr.rFonts.set(qn('w:eastAsia'), 'Arial')

# A4 + поля
for section in new_doc.sections:
    section.top_margin = Cm(1.8)
    section.bottom_margin = Cm(1.8)
    section.left_margin = Cm(2.0)
    section.right_margin = Cm(2.0)
    section.page_height = Cm(29.7)
    section.page_width = Cm(21.0)

# ---------- Утилиты ----------
def set_paragraph_spacing(p, before=0, after=6, line=1.3):
    pf = p.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.line_spacing = line

def add_run(p, text, size=11, bold=False, color=None, italic=False):
    r = p.add_run(text)
    r.font.size = Pt(size)
    r.bold = bold
    r.italic = italic
    if color:
        r.font.color.rgb = color
    r.font.name = 'Arial'
    return r

def add_para(text, size=11, bold=False, color=None, italic=False, align='left', after=6, before=0):
    p = new_doc.add_paragraph()
    set_paragraph_spacing(p, before=before, after=after, line=1.3)
    p.alignment = {
        'center': WD_ALIGN_PARAGRAPH.CENTER,
        'right':  WD_ALIGN_PARAGRAPH.RIGHT,
        'left':   WD_ALIGN_PARAGRAPH.LEFT,
    }.get(align, WD_ALIGN_PARAGRAPH.LEFT)
    add_run(p, text, size=size, bold=bold, color=color, italic=italic)
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

def page_break():
    p = new_doc.add_paragraph()
    r = p.add_run()
    r.add_break(WD_BREAK.PAGE)

def hr(color='F47629'):
    p = new_doc.add_paragraph()
    set_paragraph_spacing(p, before=2, after=6, line=1.0)
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '6')
    bottom.set(qn('w:color'), color)
    pBdr.append(bottom)
    pPr.append(pBdr)

def add_step_banner(step_num, step_title, stage, source, source_doc):
    """
    Плашка-разделитель перед каждым шагом:
    тёмный хедер с номером шага + оранжевая полоса + метаданные.
    """
    # Таблица 1×1 как «карточка шага»
    tbl = new_doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    cell = tbl.cell(0, 0)
    cell.width = Cm(17.0)
    set_cell_bg(cell, '313131')

    # Границы: левая оранжевая толстая, остальные тёмные
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for edge in ('top', 'right', 'bottom'):
        el = OxmlElement(f'w:{edge}')
        el.set(qn('w:val'), 'single')
        el.set(qn('w:sz'), '4')
        el.set(qn('w:color'), '313131')
        tcBorders.append(el)
    left_el = OxmlElement('w:left')
    left_el.set(qn('w:val'), 'single')
    left_el.set(qn('w:sz'), '36')
    left_el.set(qn('w:color'), 'F47629')
    tcBorders.append(left_el)
    tcPr.append(tcBorders)

    # Внутренние поля
    tcMar = OxmlElement('w:tcMar')
    for edge, val in (('top', '160'), ('left', '220'), ('bottom', '160'), ('right', '220')):
        el = OxmlElement(f'w:{edge}')
        el.set(qn('w:w'), val)
        el.set(qn('w:type'), 'dxa')
        tcMar.append(el)
    tcPr.append(tcMar)

    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    # Строка 1: ШАГ N / ЭТАП
    p1 = cell.paragraphs[0]
    set_paragraph_spacing(p1, before=0, after=2, line=1.2)
    r1 = p1.add_run(f'ШАГ {step_num}  ·  {stage}')
    r1.bold = True
    r1.font.size = Pt(10)
    r1.font.color.rgb = ORANGE
    r1.font.name = 'Arial'

    # Строка 2: название шага
    p2 = cell.add_paragraph()
    set_paragraph_spacing(p2, before=0, after=4, line=1.2)
    r2 = p2.add_run(step_title)
    r2.bold = True
    r2.font.size = Pt(16)
    r2.font.color.rgb = WHITE
    r2.font.name = 'Arial'

    # Строка 3: источник
    p3 = cell.add_paragraph()
    set_paragraph_spacing(p3, before=0, after=0, line=1.2)
    r3a = p3.add_run('Источник: ')
    r3a.font.size = Pt(9)
    r3a.italic = True
    r3a.font.color.rgb = RGBColor(0xC0, 0xC0, 0xC0)
    r3a.font.name = 'Arial'
    r3b = p3.add_run(source)
    r3b.font.size = Pt(9)
    r3b.font.color.rgb = RGBColor(0xFF, 0xC0, 0x80)
    r3b.bold = True
    r3b.font.name = 'Arial'

    # Небольшой отступ после плашки
    add_para('', after=2, before=0)

# ---------- Копирование элементов из исходных документов ----------
def _insert_before_sectpr(new_elem):
    """Вставляет элемент в body ПЕРЕД sectPr (чтобы сохранить структуру)."""
    body = new_doc.element.body
    sectPr = body.find(qn('w:sectPr'))
    if sectPr is not None:
        sectPr.addprevious(new_elem)
    else:
        body.append(new_elem)

def copy_main_section(start_marker_text, end_marker_text, exclude_markers=True):
    """
    Копирует все элементы основного документа между параграфами,
    текст которых начинается с start_marker_text и заканчивается
    end_marker_text (или до конца документа, если end_marker_text=None).
    """
    body = main_doc.element.body
    elems = list(body.iterchildren())
    in_section = False
    copied = 0
    for elem in elems:
        tag = elem.tag.split('}')[-1]
        if tag == 'p':
            text = ''.join(node.text or '' for node in elem.iter(qn('w:t')))
            text = text.strip()
            if not in_section:
                if text.startswith(start_marker_text):
                    in_section = True
                    if exclude_markers:
                        continue
            else:
                if end_marker_text and text.startswith(end_marker_text):
                    break
        if in_section:
            # Пропускаем sectPr — он не должен копироваться
            if tag == 'sectPr':
                continue
            new_elem = copy.deepcopy(elem)
            _insert_before_sectpr(new_elem)
            copied += 1
    return copied

def copy_supp_section(section_header_text):
    """
    Копирует весь раздел дополнения начиная с параграфа
    'Раздел N. ...' до следующего раздела 'Раздел N+1. ...' или до конца.
    Пропускает служебные плашки «Категория: ... Позиция в Sales Guide: ...».
    """
    body = supp_doc.element.body
    elems = list(body.iterchildren())
    in_section = False
    copied = 0
    for elem in elems:
        tag = elem.tag.split('}')[-1]
        if tag == 'p':
            text = ''.join(node.text or '' for node in elem.iter(qn('w:t')))
            text = text.strip()
            if not in_section:
                if text.startswith(section_header_text):
                    in_section = True
                    continue
            else:
                if text.startswith('Раздел ') and not text.startswith(section_header_text):
                    break
                if text.startswith('Категория:'):
                    continue
        if in_section:
            # Пропускаем sectPr
            if tag == 'sectPr':
                continue
            new_elem = copy.deepcopy(elem)
            _insert_before_sectpr(new_elem)
            copied += 1
    return copied

# ---------- Импорт стилей из исходных документов ----------
def import_styles_from(source_doc):
    """Импортирует все стили из source_doc в new_doc (если их ещё нет)."""
    src_styles_elem = source_doc.part.styles.element
    new_styles_elem = new_doc.part.styles.element
    existing_ids = set()
    for st in new_styles_elem.findall(qn('w:style')):
        style_id = st.get(qn('w:styleId'))
        if style_id:
            existing_ids.add(style_id)
    # Копируем недостающие
    imported = 0
    for st in src_styles_elem.findall(qn('w:style')):
        style_id = st.get(qn('w:styleId'))
        if style_id and style_id not in existing_ids:
            new_st = copy.deepcopy(st)
            new_styles_elem.append(new_st)
            existing_ids.add(style_id)
            imported += 1
    return imported

# Импортируем стили из обоих источников
imp1 = import_styles_from(main_doc)
imp2 = import_styles_from(supp_doc)
print(f'Импортировано стилей: из основного={imp1}, из дополнения={imp2}')

# ---------- Импорт нумерации ----------
def import_numbering_from(source_doc):
    """Копирует numbering.xml из source_doc в new_doc (объединяя)."""
    try:
        src_elem = source_doc.part.numbering_part.element
    except Exception:
        return 0
    try:
        new_elem = new_doc.part.numbering_part.element
    except Exception:
        # Если в новом документе нет numbering — нужно создать часть
        # python-docx не exposes public API, поэтому используем пакетный подход
        from docx.parts.numbering import NumberingPart
        from docx.opc.packuri import PackURI
        from io import BytesIO
        # Минимальный numbering.xml
        xml_bytes = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'
        partname = PackURI('/word/numbering.xml')
        new_num_part = NumberingPart.new()
        new_doc.part.relate_to(new_num_part, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering')
        new_elem = new_num_part.element
    # Копируем abstractNum и num элементы
    existing_abstract = set()
    for a in new_elem.findall(qn('w:abstractNum')):
        aid = a.get(qn('w:abstractNumId'))
        if aid is not None:
            existing_abstract.add(aid)
    existing_num = set()
    for n in new_elem.findall(qn('w:num')):
        nid = n.get(qn('w:numId'))
        if nid is not None:
            existing_num.add(nid)
    imported = 0
    for a in src_elem.findall(qn('w:abstractNum')):
        aid = a.get(qn('w:abstractNumId'))
        if aid and aid not in existing_abstract:
            new_elem.append(copy.deepcopy(a))
            existing_abstract.add(aid)
            imported += 1
    for n in src_elem.findall(qn('w:num')):
        nid = n.get(qn('w:numId'))
        if nid and nid not in existing_num:
            new_elem.append(copy.deepcopy(n))
            existing_num.add(nid)
            imported += 1
    return imported

imp_num1 = import_numbering_from(main_doc)
imp_num2 = import_numbering_from(supp_doc)
print(f'Импортировано элементов нумерации: основной={imp_num1}, дополнение={imp_num2}')

print("Utils ready")

# ============================================================
# ОБЛОЖКА
# ============================================================
# Большой оранжевый заголовок
p = new_doc.add_paragraph()
set_paragraph_spacing(p, before=60, after=4, line=1.0)
p.alignment = WD_ALIGN_PARAGRAPH.LEFT
r = p.add_run('TITAN')
r.bold = True
r.font.size = Pt(72)
r.font.color.rgb = ORANGE
r.font.name = 'Arial'

p = new_doc.add_paragraph()
set_paragraph_spacing(p, before=0, after=4, line=1.1)
r = p.add_run('RoadMap для дистрибьютора')
r.font.size = Pt(28)
r.bold = True
r.font.color.rgb = DARK
r.font.name = 'Arial'

p = new_doc.add_paragraph()
set_paragraph_spacing(p, before=0, after=20, line=1.3)
r = p.add_run('Пошаговая программа запуска продаж TITAN с полки — '
              'от знакомства с продуктом до операционного управления и Плана «Б»')
r.font.size = Pt(13)
r.italic = True
r.font.color.rgb = GREY
r.font.name = 'Arial'

hr()

# Описание
add_para('Что это за документ', bold=True, size=14, color=DARK, before=8, after=4)
add_para(
    'RoadMap объединяет два документа в единую пошаговую последовательность: '
    '«TITAN — Руководство для дистрибьюторов по продаже TITAN с полки» (основной Sales Guide) и '
    '«Дополнение к Sales Guide — 15 недостающих разделов». Текст исходных документов сохранён '
    'без изменений — изменена только последовательность и нумерация шагов.',
    after=8
)
add_para(
    'Дорожная карта выстроена так, чтобы дистрибьютор, идя от Шага 1 к Шагу 23, прошёл весь '
    'путь запуска продаж: знакомство с брендом и продуктом → расчёт экономики → подготовка '
    'торговой команды → выкладка и POSm → техника продаж → операционное управление → План «Б» '
    'и контакты для эскалации. Каждый шаг самостоятелен и снабжён плашкой с указанием источника '
    '(Основной Sales Guide или Дополнение) и этапа программы.',
    after=12
)

# Таблица 23 шагов
add_para('Карта дороги — 23 шага', bold=True, size=14, color=DARK, before=8, after=6)

steps_map = [
    ('1',  'Этап I. Знакомство с продуктом',   'Что такое TITAN',                                    'Дополнение, разд. 1'),
    ('2',  'Этап I. Знакомство с продуктом',   'Сертификаты и регалии',                              'Дополнение, разд. 3'),
    ('3',  'Этап I. Знакомство с продуктом',   'УТП TITAN — 5 причин выбрать',                       'Дополнение, разд. 2'),
    ('4',  'Этап II. Экономика продукта',     'Расчёт ССВ — экономика для клиента',                 'Дополнение, разд. 4'),
    ('5',  'Этап II. Экономика продукта',     'Цены — рекомендуемая розничная цена на полке',       'Дополнение, разд. 5'),
    ('6',  'Этап II. Экономика продукта',     'Маржинальность дистрибьютора',                       'Дополнение, разд. 6'),
    ('7',  'Этап III. Рынок и техника',       'Конкурентное окружение',                             'Дополнение, разд. 7'),
    ('8',  'Этап III. Рынок и техника',       'Технические характеристики (TDS)',                  'Дополнение, разд. 8'),
    ('9',  'Этап III. Рынок и техника',       'Инструкция по нанесению — 5 шагов',                  'Дополнение, разд. 9'),
    ('10', 'Этап IV. Канал и точки',          'В какие точки мы ставим TITAN',                       'Основной, разд. 1'),
    ('11', 'Этап IV. Канал и точки',          'Ассортиментная матрица, цены, доля цвета',           'Основной, разд. 2.1'),
    ('12', 'Этап IV. Канал и точки',          'Планограмма выкладки TITAN',                          'Основной, разд. 2.2'),
    ('13', 'Этап IV. Канал и точки',          'Где размещаем TITAN в категории «Грунт-эмали»',      'Основной, разд. 3'),
    ('14', 'Этап V. Выкладка и POSm',         'POS-материалы и рекомендации по размещению',         'Основной, разд. 4'),
    ('15', 'Этап VI. Техника продаж',         'Скрипт продаж для продавца',                         'Дополнение, разд. 11'),
    ('16', 'Этап VI. Техника продаж',         'Обработка возражений',                              'Дополнение, разд. 12'),
    ('17', 'Этап VII. Контент и обучение',    'Контент для продажи TITAN',                          'Основной, разд. 6'),
    ('18', 'Этап VII. Контент и обучение',    'Ссылка на облако с материалами',                     'Дополнение, разд. 13'),
    ('19', 'Этап VIII. Дополнительный канал', 'B2B-канал — продажи юридическим лицам',               'Дополнение, разд. 10'),
    ('20', 'Этап IX. Операционное управление','Обучение торговых представителей',                    'Основной, разд. 7'),
    ('21', 'Этап IX. Операционное управление','Предоставление статистики',                          'Основной, разд. 8'),
    ('22', 'Этап X. Управление отклонениями', 'План «Б» — корректирующие мероприятия',             'Дополнение, разд. 14'),
    ('23', 'Этап XI. Поддержка',              'Контакты и эскалация',                               'Дополнение, разд. 15'),
]

# Создаём таблицу карты
tbl = new_doc.add_table(rows=1 + len(steps_map), cols=3)
tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
tbl.autofit = False
# ширины
for i, w in enumerate([1.2, 9.0, 6.8]):
    for cell in tbl.columns[i].cells:
        cell.width = Cm(w)

# заголовок
for i, h in enumerate(['Шаг', 'Этап · Название', 'Источник']):
    c = tbl.rows[0].cells[i]
    set_cell_bg(c, '313131')
    set_cell_borders(c, color='313131', sz='4')
    c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    p = c.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_paragraph_spacing(p, before=2, after=2, line=1.2)
    r = p.add_run(h)
    r.bold = True
    r.font.size = Pt(10)
    r.font.color.rgb = WHITE
    r.font.name = 'Arial'

# тело
for ri, (num, stage, title, source) in enumerate(steps_map):
    # полное название этапа + название шага
    stage_title = f'{stage}. {title}'
    for ci, val in enumerate([num, f'{stage_title}', source]):
        c = tbl.rows[ri + 1].cells[ci]
        if ri % 2 == 1:
            set_cell_bg(c, 'F8F4F0')
        set_cell_borders(c, color='D0D0D0', sz='4')
        c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = c.paragraphs[0]
        set_paragraph_spacing(p, before=3, after=3, line=1.3)
        r = p.add_run(val)
        r.font.size = Pt(10)
        r.font.name = 'Arial'
        if ci == 0:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r.bold = True
            r.font.color.rgb = ORANGE
        elif ci == 1:
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            r.font.color.rgb = DARK
        else:
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            r.italic = True
            r.font.color.rgb = GREY

# Небольшой отступ + итог
add_para('', after=8, before=0)
add_para(
    'Дистрибьютор, который последовательно проходит все 23 шага, проходит полный цикл запуска '
    'продаж TITAN: от знакомства с продуктом и расчёта экономики — к выкладке и технике продаж — '
    'к операционному управлению и Плану «Б» при отклонениях. Каждый шаг — это обязательный '
    'контрольный пункт, после которого дистрибьютор переходит к следующему этапу программы.',
    italic=True, color=GREY, size=10, after=10
)

page_break()
print("Обложка готова")

# ============================================================
# 23 ШАГА — КОПИРОВАНИЕ СОДЕРЖИМОГО
# ============================================================

# Шаг 1. Что такое TITAN (Дополнение, разд. 1)
add_step_banner('1', 'Что такое TITAN', 'Этап I. Знакомство с продуктом',
                'Дополнение, разд. 1', supp_doc)
copy_supp_section('Раздел 1. Что такое TITAN')
page_break()

# Шаг 2. Сертификаты и регалии (Дополнение, разд. 3)
add_step_banner('2', 'Сертификаты и регалии', 'Этап I. Знакомство с продуктом',
                'Дополнение, разд. 3', supp_doc)
copy_supp_section('Раздел 3. Сертификаты и регалии')
page_break()

# Шаг 3. УТП TITAN — 5 причин выбрать (Дополнение, разд. 2)
add_step_banner('3', 'УТП TITAN — 5 причин выбрать', 'Этап I. Знакомство с продуктом',
                'Дополнение, разд. 2', supp_doc)
copy_supp_section('Раздел 2. УТП TITAN — 5 причин выбрать')
page_break()

# Шаг 4. Расчёт ССВ — экономика для клиента (Дополнение, разд. 4)
add_step_banner('4', 'Расчёт ССВ — экономика для клиента', 'Этап II. Экономика продукта',
                'Дополнение, разд. 4', supp_doc)
copy_supp_section('Раздел 4. Расчёт ССВ — экономика для клиента')
page_break()

# Шаг 5. Цены — рекомендуемая розничная цена на полке (Дополнение, разд. 5)
add_step_banner('5', 'Цены — рекомендуемая розничная цена на полке', 'Этап II. Экономика продукта',
                'Дополнение, разд. 5', supp_doc)
copy_supp_section('Раздел 5. Цены — рекомендуемая розничная цена на полке')
page_break()

# Шаг 6. Маржинальность дистрибьютора (Дополнение, разд. 6)
add_step_banner('6', 'Маржинальность дистрибьютора', 'Этап II. Экономика продукта',
                'Дополнение, разд. 6', supp_doc)
copy_supp_section('Раздел 6. Маржинальность дистрибьютора')
page_break()

# Шаг 7. Конкурентное окружение (Дополнение, разд. 7)
add_step_banner('7', 'Конкурентное окружение', 'Этап III. Рынок и техника',
                'Дополнение, разд. 7', supp_doc)
copy_supp_section('Раздел 7. Конкурентное окружение')
page_break()

# Шаг 8. Технические характеристики (TDS) (Дополнение, разд. 8)
add_step_banner('8', 'Технические характеристики (TDS)', 'Этап III. Рынок и техника',
                'Дополнение, разд. 8', supp_doc)
copy_supp_section('Раздел 8. Технические характеристики (TDS)')
page_break()

# Шаг 9. Инструкция по нанесению — 5 шагов (Дополнение, разд. 9)
add_step_banner('9', 'Инструкция по нанесению — 5 шагов', 'Этап III. Рынок и техника',
                'Дополнение, разд. 9', supp_doc)
copy_supp_section('Раздел 9. Инструкция по нанесению — 5 шагов')
page_break()

# Шаг 10. В какие точки мы ставим TITAN (Основной, разд. 1)
add_step_banner('10', 'В какие точки мы ставим TITAN', 'Этап IV. Канал и точки',
                'Основной Sales Guide, разд. 1', main_doc)
# Копируем: заголовок «1. В какие точки...» + текст + таблица классов ТТ
# Это элементы 2..4 (текст до заголовка «2. Как мы делаем выкладку»)
copy_main_section('1. В какие точки мы ставим TITAN', '2. Как мы делаем выкладку')
page_break()

# Шаг 11. Ассортиментная матрица и цены (Основной, разд. 2.1)
add_step_banner('11', 'Ассортиментная матрица, цены, доля цвета', 'Этап IV. Канал и точки',
                'Основной Sales Guide, разд. 2.1', main_doc)
copy_main_section('2.1. Ассортиментная матрица', '2.2. Планограмма')
page_break()

# Шаг 12. Планограмма выкладки (Основной, разд. 2.2)
add_step_banner('12', 'Планограмма выкладки TITAN', 'Этап IV. Канал и точки',
                'Основной Sales Guide, разд. 2.2', main_doc)
copy_main_section('2.2. Планограмма', '3. Где размещаем TITAN')
page_break()

# Шаг 13. Где размещаем TITAN в категории (Основной, разд. 3)
add_step_banner('13', 'Где размещаем TITAN в категории «Грунт-эмали»', 'Этап IV. Канал и точки',
                'Основной Sales Guide, разд. 3', main_doc)
copy_main_section('3. Где размещаем TITAN', '4. POS-материалы')
page_break()

# Шаг 14. POS-материалы (Основной, разд. 4)
add_step_banner('14', 'POS-материалы и рекомендации по размещению', 'Этап V. Выкладка и POSm',
                'Основной Sales Guide, разд. 4', main_doc)
copy_main_section('4. POS-материалы', '6. Контент для продажи')
page_break()

# Шаг 15. Скрипт продаж (Дополнение, разд. 11)
add_step_banner('15', 'Скрипт продаж для продавца', 'Этап VI. Техника продаж',
                'Дополнение, разд. 11', supp_doc)
copy_supp_section('Раздел 11. Скрипт продаж')
page_break()

# Шаг 16. Обработка возражений (Дополнение, разд. 12)
add_step_banner('16', 'Обработка возражений', 'Этап VI. Техника продаж',
                'Дополнение, разд. 12', supp_doc)
copy_supp_section('Раздел 12. Обработка возражений')
page_break()

# Шаг 17. Контент для продажи TITAN (Основной, разд. 6)
add_step_banner('17', 'Контент для продажи TITAN', 'Этап VII. Контент и обучение',
                'Основной Sales Guide, разд. 6', main_doc)
copy_main_section('6. Контент для продажи', '7. Обучение торговых представителей')
page_break()

# Шаг 18. Ссылка на облако (Дополнение, разд. 13)
add_step_banner('18', 'Ссылка на облако с материалами', 'Этап VII. Контент и обучение',
                'Дополнение, разд. 13', supp_doc)
copy_supp_section('Раздел 13. Ссылка на облако')
page_break()

# Шаг 19. B2B-канал (Дополнение, разд. 10)
add_step_banner('19', 'B2B-канал — продажи юридическим лицам', 'Этап VIII. Дополнительный канал',
                'Дополнение, разд. 10', supp_doc)
copy_supp_section('Раздел 10. B2B-канал')
page_break()

# Шаг 20. Обучение торговых представителей (Основной, разд. 7)
add_step_banner('20', 'Обучение торговых представителей', 'Этап IX. Операционное управление',
                'Основной Sales Guide, разд. 7', main_doc)
copy_main_section('7. Обучение торговых представителей', '8. Предоставление статистики')
page_break()

# Шаг 21. Предоставление статистики (Основной, разд. 8)
add_step_banner('21', 'Предоставление статистики', 'Этап IX. Операционное управление',
                'Основной Sales Guide, разд. 8', main_doc)
copy_main_section('8. Предоставление статистики', None)  # до конца
page_break()

# Шаг 22. План «Б» (Дополнение, разд. 14)
add_step_banner('22', 'План «Б» — корректирующие мероприятия', 'Этап X. Управление отклонениями',
                'Дополнение, разд. 14', supp_doc)
copy_supp_section('Раздел 14. План «Б»')
page_break()

# Шаг 23. Контакты и эскалация (Дополнение, разд. 15)
add_step_banner('23', 'Контакты и эскалация', 'Этап XI. Поддержка',
                'Дополнение, разд. 15', supp_doc)
copy_supp_section('Раздел 15. Контакты и эскалация')

# ---------- Финал документа ----------
hr()
add_para('TITAN RoadMap для дистрибьютора  ·  23 шага  ·  Версия 1.0  ·  Август 2026',
         size=10, italic=True, color=GREY, align='center', after=4)
add_para('Документ объединяет «TITAN — Руководство для дистрибьюторов по продаже TITAN с полки» '
         '(от 21.08.2026) и «Дополнение к Sales Guide — 15 разделов» в единую пошаговую '
         'последовательность. Текст исходных документов сохранён без изменений.',
         size=10, italic=True, color=GREY, align='center')

# ---------- Повтор шапок таблиц + cantSplit ----------
for table in new_doc.tables:
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

# ---------- Сохранение ----------
new_doc.save(OUT_PATH)
print(f'OK saved: {OUT_PATH}')
import os
print(f'Size: {os.path.getsize(OUT_PATH)} bytes')
