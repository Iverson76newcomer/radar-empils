"""
Добавляем раздел «Контент онлайн» после таблицы POSm (со строкой Планограмма)
в файл TITAN_Road_map_Short_27 08 2026_v2.docx.

Размещение: сразу после закрывающей строки таблицы POSm,
перед плашкой «Этап IX. Контент и обучение».
"""
import shutil
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

SRC = '/home/z/my-project/upload/TITAN_Road_map_Short_27 08 2026_v2.docx'
OUT = '/home/z/my-project/download/TITAN_Road_map_Short_27_08_2026_v2.docx'
shutil.copy(SRC, OUT)

doc = Document(OUT)

ORANGE = RGBColor(0xF4, 0x76, 0x29)
DARK   = RGBColor(0x31, 0x31, 0x31)
GREY   = RGBColor(0x6B, 0x72, 0x80)
WHITE  = RGBColor(0xFF, 0xFF, 0xFF)

# ---------- утилиты ----------
def set_spacing(p, before=0, after=6, line=1.3):
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

def add_callout(title, body, bg='FFF4ED', bar='F47629'):
    """Оранжевая плашка-выноска с одним абзацем текста."""
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    cell = tbl.cell(0, 0)
    cell.width = Cm(17.0)
    # Заливка фона
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), bg)
    tcPr.append(shd)
    # Границы: левая оранжевая толстая, остальные под цвет фона
    tcBorders = OxmlElement('w:tcBorders')
    for edge, c in (('top', bg), ('right', bg), ('bottom', bg), ('left', bar)):
        el = OxmlElement(f'w:{edge}')
        el.set(qn('w:val'), 'single')
        el.set(qn('w:sz'), '36' if edge == 'left' else '4')
        el.set(qn('w:color'), c)
        tcBorders.append(el)
    tcPr.append(tcBorders)
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    # Заголовок
    if title:
        p = cell.paragraphs[0]
        set_spacing(p, before=2, after=2, line=1.2)
        add_run(p, title, size=11, bold=True, color=ORANGE)
    # Тело
    if body:
        p2 = cell.add_paragraph()
        set_spacing(p2, before=0, after=2, line=1.3)
        add_run(p2, body, size=10, color=DARK)
    # Внутренние поля
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for edge, val in (('top', '120'), ('left', '200'), ('bottom', '120'), ('right', '200')):
        el = OxmlElement(f'w:{edge}')
        el.set(qn('w:w'), val)
        el.set(qn('w:type'), 'dxa')
        tcMar.append(el)
    tcPr.append(tcMar)
    return tbl

# Текст раздела (6-8 строк)
SECTION_TITLE = 'Контент онлайн: сайты и интернет-магазины дистрибьютора'
SECTION_BODY = (
    'Информация о TITAN на сайте и в интернет-магазине дистрибьютора должна полностью '
    'соответствовать актуальной банке и характеристикам на titanpaints.ru. '
    'Проверяется: правильное наименование продукта (Ореол TITAN 3в1 / быстросохнущая / '
    'молотковая), верный артикул и фасовка, корректные технические параметры (расход, '
    'время высыхания, срок службы), соответствие изображения банки реальной этикетке, '
    'наличие сертификатов и актуальной розничной цены. Региональный менеджер Эмпилс '
    'проводит аудит онлайн-витрины раз в квартал; при обнаружении расхождений '
    'дистрибьютор обязан внести правки в течение 5 рабочих дней.'
)

# ---------- ищем место для вставки ----------
# Нужно найти элемент после последней строки таблицы POSm (со строкой "Планограмма")
# и перед плашкой "Этап IX. Контент и обучение"
body = doc.element.body
elems = list(body.iterchildren())

# Найдём индекс параграфа, который содержит «Этап IX. Контент и обучение»
target_idx = None
for i, e in enumerate(elems):
    tag = e.tag.split('}')[-1]
    if tag == 'p':
        text = ''.join(node.text or '' for node in e.iter(qn('w:t')))
        if 'Этап IX' in text and 'Контент' in text:
            target_idx = i
            break

if target_idx is None:
    # Возможно, плашка в другой структуре — поищем шире
    for i, e in enumerate(elems):
        tag = e.tag.split('}')[-1]
        if tag == 'tbl':
            # проверяем, есть ли в таблице плашка с «Этап IX»
            text = ''.join(node.text or '' for node in e.iter(qn('w:t')))
            if 'Этап IX' in text and 'Контент' in text:
                target_idx = i
                break

print(f'target_idx = {target_idx}')

if target_idx is None:
    print('ERROR: не найден раздел «Этап IX. Контент и обучение»')
else:
    # Создаём новые элементы: параграф-заголовок + параграф-тело
    # Используем стандартное создание через doc.add_paragraph, затем перенесём XML

    # Заголовок раздела (оранжевый, жирный, с полосой слева)
    p_title = doc.add_paragraph()
    set_spacing(p_title, before=12, after=4, line=1.2)
    pPr = p_title._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    left_border = OxmlElement('w:left')
    left_border.set(qn('w:val'), 'single')
    left_border.set(qn('w:sz'), '24')
    left_border.set(qn('w:space'), '8')
    left_border.set(qn('w:color'), 'F47629')
    pBdr.append(left_border)
    pPr.append(pBdr)
    p_title.paragraph_format.left_indent = Cm(0.3)
    add_run(p_title, SECTION_TITLE, size=13, bold=True, color=DARK)

    # Тело раздела
    p_body = doc.add_paragraph()
    set_spacing(p_body, before=0, after=8, line=1.3)
    p_body.alignment = WD_ALIGN_PARAGRAPH.LEFT
    add_run(p_body, SECTION_BODY, size=11, color=DARK)

    # Переносим оба новых элемента из конца body в нужную позицию
    # Сейчас они добавлены в конец body — нужно переместить их перед target_idx
    target_elem = elems[target_idx]
    # Новые элементы — это последние два в body
    new_p_title = body[-2]  # предпоследний
    new_p_body  = body[-1]  # последний
    # Удаляем из конца и вставляем перед target_elem
    body.remove(new_p_title)
    body.remove(new_p_body)
    target_elem.addprevious(new_p_title)
    target_elem.addprevious(new_p_body)

    doc.save(OUT)
    print(f'OK saved: {OUT}')
    import os
    print(f'Size: {os.path.getsize(OUT)} bytes')
