"""
КП Вегос-М (Бурятия) — на основе структуры КП Очаг v2.
Единый шрифт Calibri, блок прибыли наверху, симметричные отступы.
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
rFonts.set(qn('w:cs'), 'Calibri')

for section in doc.sections:
    section.top_margin = Cm(1.5)
    section.bottom_margin = Cm(1.5)
    section.left_margin = Cm(2.0)
    section.right_margin = Cm(1.5)

# === КОНСТАНТЫ ===
BLUE = RGBColor(0x1F, 0x4E, 0x79)
RED = RGBColor(0xC0, 0x00, 0x00)
GREEN = RGBColor(0x37, 0x56, 0x23)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
BLACK = RGBColor(0x1A, 0x1A, 0x1A)
GRAY = RGBColor(0x59, 0x59, 0x59)
LIGHT_BLUE = "D6E4F0"
LIGHT_GREEN = "E2EFDA"
LIGHT_RED = "FFC7CE"
ROW_ALT = "F2F7FC"

# === УТИЛИТЫ ===
def set_cell_shading(cell, color_hex):
    shading = OxmlElement('w:shd')
    shading.set(qn('w:fill'), color_hex)
    shading.set(qn('w:val'), 'clear')
    cell._element.get_or_add_tcPr().append(shading)

def set_cell_borders(cell, color='BFBFBF', size='4'):
    tcPr = cell._element.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for border_name in ['top', 'left', 'bottom', 'right']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'single')
        border.set(qn('w:sz'), size)
        border.set(qn('w:color'), color)
        tcBorders.append(border)
    tcPr.append(tcBorders)

def add_cell_text(cell, text, bold=False, size=10, color=BLACK, align='left'):
    cell.text = ''
    p = cell.paragraphs[0]
    if align == 'center':
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    elif align == 'right':
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    else:
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
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

def add_table(doc, rows, cols, data, col_widths=None):
    table = doc.add_table(rows=rows, cols=cols)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = True
    if col_widths:
        for i, width in enumerate(col_widths):
            for row in table.rows:
                row.cells[i].width = Cm(width)
    for ri, row_data in enumerate(data):
        for ci, cell_data in enumerate(row_data):
            cell = table.rows[ri].cells[ci]
            text = cell_data if isinstance(cell_data, str) else str(cell_data)
            if ri == 0:
                add_cell_text(cell, text, bold=True, size=10, color=WHITE, align='center')
                set_cell_shading(cell, '1F4E79')
            else:
                is_bold = text.startswith('ИТОГО') or text.startswith('Итого')
                fill = ROW_ALT if ri % 2 == 0 else None
                add_cell_text(cell, text, bold=is_bold, size=10, color=BLACK, align='left')
                if fill:
                    set_cell_shading(cell, fill)
                if is_bold:
                    set_cell_shading(cell, LIGHT_BLUE)
            set_cell_borders(cell)
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(4)
    return table

def add_heading(doc, text, size=12, color=BLUE, bold=True, space_before=12, space_after=4):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    run = p.add_run(text)
    run.font.name = 'Calibri'
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    return p

def add_text(doc, text, size=10, color=BLACK, bold=False, italic=False, align='left', space_after=4):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(space_after)
    if align == 'center':
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    elif align == 'right':
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = p.add_run(text)
    run.font.name = 'Calibri'
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    return p

def add_hr(doc):
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
add_text(doc, 'Коммерческое предложение ЗАО «Эмпилс»', size=16, bold=True, color=BLUE, align='center', space_after=2)
add_text(doc, 'Партнёр: Вегос-М (Урбаева В.С.) · Республика Бурятия', size=11, color=GRAY, align='center', space_after=8)

add_hr(doc)

# === БЛОК ПРИБЫЛИ (ВВЕРХУ) ===
add_heading(doc, 'ЗАРАБОТОК ПАРТНЁРА', size=13, space_before=8, space_after=6)

profit_data = [
    ['Период', 'Товарооборот', 'Доход от наценки (35%)', 'Бонусы Эмпилс', 'ИТОГО ДОХОД', 'Рост'],
    ['Год 1', '10 млн ₽', '4 млн ₽', '0,9 млн ₽', '4,9 млн ₽', '—'],
    ['Год 2', '21 млн ₽', '7 млн ₽', '2,1 млн ₽', '9,1 млн ₽', '×1,9'],
    ['Год 3', '50 млн ₽', '18 млн ₽', '5,5 млн ₽', '23,5 млн ₽', '×2,6'],
    ['ИТОГО за 3 года', '81 млн ₽', '29 млн ₽', '8,5 млн ₽', '37,5 млн ₽', ''],
]

table = doc.add_table(rows=5, cols=6)
table.alignment = WD_TABLE_ALIGNMENT.CENTER
col_widths_profit = [2.5, 2.5, 3.0, 2.5, 2.8, 1.5]
for i, w in enumerate(col_widths_profit):
    for row in table.rows:
        row.cells[i].width = Cm(w)

for ri, row_data in enumerate(profit_data):
    for ci, cell_data in enumerate(row_data):
        cell = table.rows[ri].cells[ci]
        text = str(cell_data)
        if ri == 0:
            add_cell_text(cell, text, bold=True, size=9, color=WHITE, align='center')
            set_cell_shading(cell, '1F4E79')
        elif ri == 4:
            add_cell_text(cell, text, bold=True, size=10, color=BLUE, align='center')
            set_cell_shading(cell, LIGHT_BLUE)
        else:
            is_highlight = (ci == 4)
            add_cell_text(cell, text, bold=is_highlight, size=10, 
                          color=GREEN if is_highlight else BLACK, align='center')
            if is_highlight:
                set_cell_shading(cell, LIGHT_GREEN)
            elif ri % 2 == 1:
                set_cell_shading(cell, ROW_ALT)
        set_cell_borders(cell)

doc.add_paragraph().paragraph_format.space_after = Pt(2)

add_text(doc, '50 млн ₽ — целевой товарооборот Эмпилс в Год 3 = 8% потенциала региона. Заработок партнёра: 18 млн ₽/год при наценке 35%.', 
         size=10, bold=True, color=BLUE, space_after=4)
add_text(doc, 'Покрытие АКБ: 200 РТТ. Дистрибьютор: 1 ЭТП + собственная команда торговых представителей.', 
         size=10, color=GRAY, space_after=4)

add_hr(doc)

# === ТОВАРООБОРОТ И ЗАРАБОТОК ПО ГОДАМ ===
add_heading(doc, 'ТОВАРООБОРОТ И ЗАРАБОТОК ПАРТНЁРА ПО ГОДАМ', size=12)

to_data = [
    ['Год', 'Товарооборот', 'Наценка 35% (доход партнёра)', 'Комментарий'],
    ['Год 1', '10 млн ₽', '4 млн ₽', '1 ЭТП + ТП дистрибьютора, 300 SKU'],
    ['Год 2', '21 млн ₽', '7 млн ₽', 'Расширение АКБ до РТТ'],
    ['Год 3', '50 млн ₽', '18 млн ₽', 'Полное покрытие 8% потенциала'],
    ['ИТОГО', '81 млн ₽', '29 млн ₽', 'За 3 года'],
]
add_table(doc, 5, 4, to_data, col_widths=[2.0, 3.0, 4.5, 6.0])

# === СИСТЕМА БОНУСИРОВАНИЯ ===
add_heading(doc, 'СИСТЕМА БОНУСИРОВАНИЯ', size=12)

add_text(doc, 'Условия: выполнение плана по ассортименту (наличие обязательных SKU) + выполнение ежемесячного плана товарооборота + выполнение годового плана. Бонус выплачивается ретро-бонусом (кредит-нотой) по итогам года.', 
         size=10, space_after=4)

bonus_data = [
    ['Год', 'Товарооборот', 'Бонус за ассортимент (4%)', 'Бонус за ежемес. ТО (4%)', 'Бонус за год. план', 'Итого бонус', 'Итого доход (наценка + бонус)'],
    ['Год 1', '10 млн ₽', '0,4 млн ₽', '0,4 млн ₽', '0,1 млн ₽ (1%)', '0,9 млн ₽ (9%)', '4,9 млн ₽'],
    ['Год 2', '21 млн ₽', '0,84 млн ₽', '0,84 млн ₽', '0,42 млн ₽ (2%)', '2,1 млн ₽ (10%)', '9,1 млн ₽'],
    ['Год 3', '50 млн ₽', '2,0 млн ₽', '2,0 млн ₽', '1,5 млн ₽ (3%)', '5,5 млн ₽ (11%)', '23,5 млн ₽'],
    ['ИТОГО', '81 млн ₽', '3,24 млн ₽', '3,24 млн ₽', '2,02 млн ₽', '8,5 млн ₽', '37,5 млн ₽'],
]
add_table(doc, 5, 7, bonus_data, col_widths=[1.5, 2.0, 2.8, 2.8, 2.2, 2.5, 3.0])

add_hr(doc)

# === ДЕТАЛИЗАЦИЯ РАСЧЁТА ===
add_text(doc, '↓ ДЕТАЛИЗАЦИЯ РАСЧЁТА ↓', size=11, bold=True, color=BLUE, align='center', space_after=8)

# === 1. ПОТЕНЦИАЛ РЕГИОНА — НАСЕЛЕНИЕ ===
add_heading(doc, '1. ПОТЕНЦИАЛ РЕГИОНА — НАСЕЛЕНИЕ', size=12)
add_text(doc, 'База расчёта: городское население — 560 000 чел.', size=10, bold=True, space_after=4)

pop_data = [
    ['Регион', 'Население (чел.)', 'Ключевые города', 'Города (чел.)'],
    ['Республика Бурятия', '982 000', 'Улан-Удэ', '435 000'],
    ['+ Гусиноозёрск, Северобайкальск', '—', '2 города', '50 000'],
    ['+ Северобайкальск, Кяхта, Бабушкин', '—', '3 города', '75 000'],
    ['ИТОГО', '982 000', '6 городов', '560 000'],
]
add_table(doc, 5, 4, pop_data, col_widths=[5.5, 3.0, 3.5, 3.5])

# === 2. ПОТЕНЦИАЛ РЫНКА ЛКМ ===
add_heading(doc, '2. ПОТЕНЦИАЛ РЫНКА ЛКМ (по городскому населению)', size=12)

pot_data = [
    ['Параметр', 'В рублях', 'В тоннах'],
    ['Потребление ЛКМ на душу', '~1123 ₽/чел.', '~14.59 кг/чел.'],
    ['Расчёт', '560 000 × 1123 ₽', '560 000 × 14.59 кг'],
    ['ПОТЕНЦИАЛ РЫНКА ЛКМ', '629 млн ₽', '8 170 т'],
]
add_table(doc, 4, 3, pot_data, col_widths=[5.0, 5.5, 5.0])

# Сводная таблица
summary_data = [
    ['ПОТЕНЦИАЛ РЫНКА', 'ФАКТИЧЕСКИЙ ТО', 'НЕДОЗАГРУЗКА', 'ДОЛЯ ЭМПИЛС 8%'],
    ['629 млн ₽\n8 170 т', '158 млн ₽', '471 млн ₽\n75% рынка', '50 млн ₽\nцель на 3 года'],
]
table = doc.add_table(rows=2, cols=4)
table.alignment = WD_TABLE_ALIGNMENT.CENTER
for ri, row_data in enumerate(summary_data):
    for ci, cell_data in enumerate(row_data):
        cell = table.rows[ri].cells[ci]
        if ri == 0:
            add_cell_text(cell, cell_data, bold=True, size=10, color=WHITE, align='center')
            set_cell_shading(cell, '1F4E79')
        else:
            colors = [BLUE, BLACK, RED, GREEN]
            fills = [LIGHT_BLUE, ROW_ALT, LIGHT_RED, LIGHT_GREEN]
            add_cell_text(cell, cell_data, bold=True, size=11, color=colors[ci], align='center')
            set_cell_shading(cell, fills[ci])
        set_cell_borders(cell)
        cell.width = Cm(3.7)

doc.add_paragraph().paragraph_format.space_after = Pt(2)

# === 3. ТЕКУЩИЙ РЫНОК — ДИСТРИБЬЮТОРЫ ===
add_heading(doc, '3. ТЕКУЩИЙ РЫНОК — ФАКТИЧЕСКИЕ ДИСТРИБЬЮТОРЫ', size=12)

dist_data = [
    ['Дистрибьютор', 'Портфель конкурентов', 'Территория'],
    ['Вегос-М (Урбаева В.С.)', 'АВС Фарбен (82,6%), Тиккурилла (12,8%), Мефферт, Новый Дом, Лакра', 'Республика Бурятия'],
    ['ИТОГО (1 дистрибьютор)', 'Совокупный ТО: 158 млн ₽', '1 регион'],
]
add_table(doc, 3, 3, dist_data, col_widths=[4.5, 7.0, 4.0])

# === НЕДОЗАГРУЗКА ===
add_heading(doc, 'НЕДОЗАГРУЗКА РЕГИОНА', size=12, color=RED)

add_text(doc, '⚠ Регион недозагружен на 75% — потенциал 471 млн ₽ не освоен', size=10, bold=True, color=RED, space_after=4)

underload_data = [
    ['Показатель', 'Млрд ₽', 'Тонн', 'Доля'],
    ['Теоретический потенциал', '0.63', '8 170', '100%'],
    ['Фактический ТО дистрибьюторов', '0.16', '158 млн ₽', '25%'],
    ['НЕДОЗАГРУЗКА', '0.47', '6 130 т', '75%'],
]
add_table(doc, 4, 4, underload_data, col_widths=[5.0, 2.5, 3.5, 2.0])

# === 4. ДОЛЯ ЭМПИЛС ===
add_heading(doc, '4. ДОЛЯ ЭМПИЛС — 8% РЫНКА', size=12)

add_text(doc, '50 млн ₽ — целевой товарооборот Эмпилс, достигаемый за 3 года. Покрытие АКБ: 200 РТТ.', size=10, bold=True, space_after=4)

share_data = [
    ['Показатель', 'В рублях', 'В тоннах'],
    ['Потенциал рынка (100%)', '629 млн ₽', '8 170 т'],
    ['Доля Эмпилс (8%)', '50 млн ₽', '654 т'],
]
add_table(doc, 3, 3, share_data, col_widths=[5.0, 5.5, 5.0])

# === 5. ЦЕНООБРАЗОВАНИЕ ===
add_heading(doc, '5. ЦЕНООБРАЗОВАНИЕ — НАЦЕНКИ', size=12)

add_text(doc, 'Пример: при заводской цене 100 ₽ → дистрибьютор продаёт в РТТ по 135 ₽ → РТТ продаёт потребителю по ~202 ₽.', size=10, space_after=4)

markup_data = [
    ['Уровень', 'Цепочка', 'Наценка', 'Описание'],
    ['Первичная', 'Дистрибьютор → РТТ', '35%', 'Базовая наценка дистрибьютора при продаже в РТТ'],
    ['Вторичная', 'РТТ → Потребитель', '50%', 'Наценка РТТ при продаже конечному клиенту'],
    ['Итого', 'Дистрибьютор → Потребитель', '~102%', 'Совокупная наценка от заводской цены до цены на полке'],
]
add_table(doc, 4, 4, markup_data, col_widths=[2.5, 4.5, 2.0, 7.0])

# === 6. ОТСРОЧКА ===
add_heading(doc, '6. ОТСРОЧКА', size=12)
add_text(doc, 'Диапазон 30-45 дней.', size=10, space_after=4)

# === 7. ОРГАНИЗАЦИЯ ПРОДАЖ ===
add_heading(doc, '7. ОРГАНИЗАЦИЯ ПРОДАЖ — ЭКСКЛЮЗИВНЫЕ ТП', size=12)
add_text(doc, '75% ЭТП нанимается в период первичной загрузки дистрибьютора.', size=10, space_after=4)

etp_data = [
    ['Параметр', 'Значение'],
    ['1 ЭТП покрывает', '120 РТТ ОКБ (80 АКБ)'],
    ['Товарооборот через ЭТП', '6 млн ₽/год'],
    ['Товарооборот через ТП дистрибьютора', '4 млн ₽/год'],
    ['ИТОГО год 1', '10 млн ₽'],
]
add_table(doc, 5, 2, etp_data, col_widths=[6.0, 10.0])

# === 8. ПРОГНОЗ ТОВАРООБОРОТА ===
add_heading(doc, '8. ПРОГНОЗ ТОВАРООБОРОТА ПО ГОДАМ', size=12)

forecast_data = [
    ['Год', 'Товарооборот', 'Рост', 'Комментарий'],
    ['Год 1', '10 млн ₽', '—', '1 ЭТП + ТП дистрибьютора, 300 SKU'],
    ['Год 2', '21 млн ₽', '×2,1', 'Расширение АКБ до 300+ РТТ'],
    ['Год 3', '50 млн ₽', '×2,4', 'Полное покрытие 8% потенциала + ЭТП'],
]
add_table(doc, 4, 4, forecast_data, col_widths=[2.0, 3.0, 2.0, 9.0])

# === 9. АССОРТИМЕНТ ===
add_heading(doc, '9. АССОРТИМЕНТ НАЧАЛА ВЗАИМОДЕЙСТВИЯ', size=12)

add_text(doc, 'Стартовый ассортимент: 300 SKU из портфеля Эмпилс', size=10, bold=True, space_after=2)
add_text(doc, 'Бренды: · TITAN · PROFIWOOD · ОРЕОЛ · ОРЕОЛ ДИСКОНТ · РАСЦВЕТ · ПРОСТОКРАШЕНО', size=10, space_after=2)
add_text(doc, 'Категории: краски · фасадные · лаки · эмали · грунт-эмали · грунтовки · шпатлёвки · ЗДП', size=10, space_after=4)

# === 10. ПОДДЕРЖКА ===
add_heading(doc, '10. ПОДДЕРЖКА', size=12)

support_data = [
    ['Программа', 'Описание'],
    ['Программа мотивации общего прайса', 'Бонусы для ТП, продающих весь ассортимент Эмпилс'],
    ['Мотивация сотрудников РТТ', 'Программа поощрения продавцов торговых точек за рекомендации продукции Эмпилс'],
    ['Продвижение фокусных марок', 'Целевое продвижение топ-SKU (TITAN, PROFIWOOD) с повышенной маржой'],
    ['Гибкая отсрочка', 'Индивидуальные условия отсрочки платежа в зависимости от объёма закупки и истории'],
    ['Команда ЭТП', 'Старт с 1 эксклюзивным торговым представителем, обученным и мотивируемым Эмпилс'],
    ['CISLINK DTS', 'Подключение к системе аналитики дистрибьюторов'],
]
add_table(doc, 7, 2, support_data, col_widths=[5.0, 11.0])

# === ФУТЕР ===
add_hr(doc)
add_text(doc, 'Предложение действительно до 15.12.2026', size=10, bold=True, align='right')
add_text(doc, 'Региональный директор__________________  ', size=10, bold=True, align='right')

# === СОХРАНИТЬ ===
out_path = "/home/z/my-project/download/КП_Вегос-М_коммерческое_предложение.docx"
doc.save(out_path)
print(f"✓ Файл сохранён: {out_path}")
print(f"  Размер: {os.path.getsize(out_path):,} байт")
