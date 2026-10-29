"""
Памятка v2 — с правильным порядком разделов.
"""
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import os

doc = Document()

style = doc.styles['Normal']
font = style.font
font.name = 'Calibri'
font.size = Pt(8)
rPr = style.element.get_or_add_rPr()
rFonts = rPr.find(qn('w:rFonts'))
if rFonts is None:
    rFonts = OxmlElement('w:rFonts')
    rPr.insert(0, rFonts)
rFonts.set(qn('w:ascii'), 'Calibri')
rFonts.set(qn('w:hAnsi'), 'Calibri')

for section in doc.sections:
    section.top_margin = Cm(0.5)
    section.bottom_margin = Cm(0.4)
    section.left_margin = Cm(1.0)
    section.right_margin = Cm(0.8)

BLUE = RGBColor(0x1F, 0x4E, 0x79)
RED = RGBColor(0xC0, 0x00, 0x00)
GREEN = RGBColor(0x37, 0x56, 0x23)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
BLACK = RGBColor(0x1A, 0x1A, 0x1A)
GRAY = RGBColor(0x59, 0x59, 0x59)
ROW_ALT = "F2F7FC"
WARN_FILL = "FFF2CC"
GREEN_FILL = "E2EFDA"

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

def add_cell(cell, text, bold=False, size=7, color=BLACK, align='left', fill=None):
    cell.text = ''
    p = cell.paragraphs[0]
    if align == 'center': p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    elif align == 'right': p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    lines = text.split('\n')
    for i, line in enumerate(lines):
        if i > 0:
            r = p.add_run()
            r.add_break()
        r = p.add_run(line)
        r.font.name = 'Calibri'
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.color.rgb = color
    if fill: set_cell_shading(cell, fill)
    set_cell_borders(cell)

def add_heading(text, size=9, color=BLUE, space_before=4, space_after=1):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    r = p.add_run(text)
    r.font.name = 'Calibri'
    r.font.size = Pt(size)
    r.font.bold = True
    r.font.color.rgb = color
    return p

def add_text(text, size=8, color=BLACK, bold=False, align='left', space_after=1):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(space_after)
    if align == 'center': p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(text)
    r.font.name = 'Calibri'
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.color.rgb = color
    return p

def add_table_simple(data, col_widths, header_fill='1F4E79', header_color=WHITE, body_size=7):
    rows = len(data)
    cols = len(data[0])
    table = doc.add_table(rows=rows, cols=cols)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, w in enumerate(col_widths):
        for row in table.rows:
            row.cells[i].width = Cm(w)
    for ri, row_data in enumerate(data):
        for ci, val in enumerate(row_data):
            cell = table.rows[ri].cells[ci]
            if ri == 0:
                add_cell(cell, val, bold=True, size=body_size, color=header_color, align='center', fill=header_fill)
            else:
                is_bold = str(val).startswith('ИТОГО') or str(val).startswith('Итого')
                fill = ROW_ALT if ri % 2 == 0 else None
                add_cell(cell, val, bold=is_bold, size=body_size, color=BLACK, fill=fill)
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(1)
    p.paragraph_format.space_before = Pt(0)
    for run in p.runs:
        run.font.size = Pt(1)
    return table

# ====================== ПАМЯТКА ======================

add_text('ПАМЯТКА РЕГИОНАЛЬНОГО МЕНЕДЖЕРА', size=13, bold=True, color=BLUE, align='center', space_after=0)
add_text('Запуск колеровочного оборудования Эмпилс в торговой точке', size=8, color=GRAY, align='center', space_after=2)

# === 1. ГДЕ СТАВИТЬ ===
add_heading('1. ГДЕ СТАВИТЬ ОБОРУДОВАНИЕ В МАГАЗИНЕ')
place_data = [
    ['Тип магазина', 'Площадь', 'Где ставить', 'Зона'],
    ['Маленький DIY', 'до 100 м²', 'У дальней стены или в углу. Видна с входа.', '1–1,5 м²'],
    ['Средний', '100–500 м²', '«Колеровочный остров» в центре или вдоль стены', '2–4 м²'],
    ['Крупный DIY', '500+ м²', 'Отдельная зона со стендом, компьютером, принтером', '4–6 м²'],
]
add_table_simple(place_data, [3.0, 2.5, 7.0, 2.5])

# === 2. ПРИНЦИПЫ РАЗМЕЩЕНИЯ ===
add_heading('2. ПРИНЦИПЫ РАЗМЕЩЕНИЯ')
principles_data = [
    ['Принцип', 'Что делать'],
    ['Видимость с входа', 'Зона колеровки видна сразу при входе — сигнал «у нас любой цвет»'],
    ['Рядом с красками', 'Рядом с полками AvantGarde, NEWTONE, ОРЕОЛ — взял банку, дошёл до машины'],
    ['Ровный пол', 'Бетонная основа 1×1 м (для устойчивости оборудования)'],
    ['Электропитание', '2 евро-розетки, провод 2,5 мм², заземление'],
    ['Температура', 'Не ниже +15°C (колоранты густеют при холоде)'],
    ['Освещение', 'Желательно 3 типа: дневной, люминесцентный, накаливания'],
    ['Техдоступ', '0,5 м со стороны обслуживания для замены колорантов и ремонта'],
]
add_table_simple(principles_data, [3.0, 12.0])

# === 3. КОЛЕРОВОЧНОЕ ОБОРУДОВАНИЕ ===
add_heading('3. КОЛЕРОВОЧНОЕ ОБОРУДОВАНИЕ (предоставляет Эмпилс)')
equip_data = [
    ['Оборудование', 'Производитель', 'Ключевые характеристики'],
    ['Дозатор COROB D180', 'Corob, Италия', '16 канистр, bellows pump, мин. доза 0,051 мл, 1250×710×880 мм, 90 кг'],
    ['Миксер MultiSpeed CABO', 'CABO, Италия', 'Гироскопический, 3 скорости, 800×860×1040 мм, 170 кг, рольставни'],
    ['Рекламная стойка NEWTONE', 'Эмпилс', 'Фирменный POS-стенд для привлечения внимания'],
    ['Веер RAL K7 + NCS index 1950', 'Эмпилс', '2 базовых каталога цветов (213 + 1950 цветов)'],
    ['ПО Innovatint 3/4', 'Vibrantz', '8 каталогов цветов, управление дозатором, печать этикеток'],
    ['16 колорантов MONICOLOR PRO', 'Закупка дистра', 'Доля пигмента 7–65%, фасовка 950 мл'],
]
add_table_simple(equip_data, [4.0, 3.0, 8.0])

# === 4. КАКОЙ АССОРТИМЕНТ ЗАВЕСТИ ===
add_heading('4. КАКОЙ АССОРТИМЕНТ ЗАВЕСТИ (колеровочные SKU)')

add_text('AvantGarde (премиум):', size=8, bold=True, color=BLUE, space_after=0)
avant_data = [
    ['Продукт', 'База', 'Фасовки'],
    ['Краска интерьерная моющаяся AvantGarde', 'А и С', '0,7 / 4,5 / 8,5 л'],
    ['Краска фасадная AvantGarde', 'А и С', '4,5 / 8,5 л'],
]
add_table_simple(avant_data, [7.5, 2.0, 5.5])

add_text('NEWTONE (средний сегмент):', size=8, bold=True, color=BLUE, space_after=0)
newtone_data = [
    ['Продукт', 'База', 'Фасовки'],
    ['NEWTONE INTERIO суперстойкая интерьерная', 'А и С', '0,8 / 4,5 / 8,5 л'],
    ['NEWTONE MODERN PRO фасадная силиконовая', 'А и С', '0,8 / 4,5 / 8,5 л'],
    ['NEWTONE LOTUS 3/7/20 моющаяся интерьерная', 'А и С', '0,8 / 4,5 / 8,5 л'],
]
add_table_simple(newtone_data, [7.5, 2.0, 5.5])

add_text('ОРЕОЛ + РАСЦВЕТ + ПРОСТОКРАШЕНО (массовый сегмент):', size=8, bold=True, color=BLUE, space_after=0)
oreol_data = [
    ['Продукт', 'База', 'Фасовки'],
    ['Краска ОРЕОЛ для интерьеров моющаяся', 'А и С', '2,2 / 9 л'],
    ['Краска ОРЕОЛ для фасадов', 'А и С', '4,5 / 9 л'],
    ['Эмаль акриловая ОРЕОЛ', 'А и С', '0,9 / 1,9 кг'],
    ['Эмаль РАСЦВЕТ ПФ-115 (белая)', '—', '0,9 / 1,9 / 2,7 кг'],
    ['Эмаль ПРОСТОКРАШЕНО ПФ-115 (белая)', '—', '0,9 / 1,9 / 2,7 кг'],
]
add_table_simple(oreol_data, [7.5, 2.0, 5.5])

# === 5. ЦВЕТОВЫЕ ВЕЕРА И КАТАЛОГИ ===
add_heading('5. ЦВЕТОВЫЕ ВЕЕРА И КАТАЛОГИ (поддерживаются ПО Innovatint)')
catalogs_data = [
    ['Каталог', 'Кол-во цветов', 'Тип'],
    ['RAL 840 HR', '213', 'Международный'],
    ['RAL Design D2', '1 625', 'Международный'],
    ['NCS index 1950', '1 950', 'Международный'],
    ['Symphony I & II (Tikkurila)', '2 436', 'Фирменный'],
    ['Monicolor Nova 2024', '1 600', 'Фирменный'],
    ['Caparol 3D PLUS', '1 600', 'Фирменный'],
    ['Dulux Trade Color Palette 5', '1 200', 'Фирменный'],
]
add_table_simple(catalogs_data, [6.5, 3.5, 5.0])
add_text('Итого: 12 000+ цветов из 8 каталогов — покупатель может прийти с любым веером.', size=7, bold=True, color=GREEN, space_after=1)

# === 6. ОФОРМЛЕНИЕ ЗОНЫ ===
add_heading('6. ОФОРМЛЕНИЕ ЗОНЫ КОЛЕРОВКИ В РТТ — КЛЮЧ К ОБЪЁМАМ', color=RED)

table = doc.add_table(rows=1, cols=1)
table.alignment = WD_TABLE_ALIGNMENT.CENTER
table.rows[0].cells[0].width = Cm(15.0)
add_cell(table.rows[0].cells[0],
    'БЕЗ ОФОРМЛЕНИЯ ЗОНЫ КОЛЕРОВКИ ОБЪЁМЫ НЕ ПОЙДУТ.\n'
    'Зона должна быть яркой, заметной и привлекать внимание покупателя с момента входа в магазин.',
    bold=True, size=8, color=RED, align='center', fill=WARN_FILL)

p = doc.add_paragraph()
p.paragraph_format.space_after = Pt(1)
for run in p.runs:
    run.font.size = Pt(1)

zone_data = [
    ['Элемент оформления', 'Зачем'],
    ['Баннер «Колеровка» над зоной', 'Виден с входа, привлекает внимание'],
    ['Входная группа (указатель)', 'Направляет покупателя к зоне колеровки'],
    ['Рекламная стойка NEWTONE', 'Фирменный POS-стенд, узнаваемость бренда'],
    ['Стенд с выкрасками на стене', 'Покупатель видит реальные цвета, а не бумагу'],
    ['Белый парус с логотипом Эмпилс', 'Контраст, подсветка, фирменный стиль'],
    ['Подсветка (3 типа света)', 'Критично! Покупатель видит цвет как дома'],
    ['Серый фон (алюминий)', 'Подчёркивает яркость выкрасок, скрывает пыль'],
    ['Веера RAL + NCS на стенде', 'Покупатель выбирает цвет сам — ускоряет продажу'],
    ['Знак «Любой цвет за 5 минут»', 'Снимает барьер «долго ждать»'],
]
add_table_simple(zone_data, [6.0, 9.0])

# === 7. ПРОМОПЕРСОНАЛ ===
add_heading('7. ПРОМОПЕРСОНАЛ — ДВИЖУЩАЯ СИЛА ОБЪЁМОВ', color=RED)

table = doc.add_table(rows=1, cols=1)
table.alignment = WD_TABLE_ALIGNMENT.CENTER
table.rows[0].cells[0].width = Cm(15.0)
add_cell(table.rows[0].cells[0],
    'Зона колеровки без мотивированного промоперсонала = мёртвый капитал.\n'
    'Нужен оператор, обученный работе с дозатором и замотивированный на объёмы продаж.\n\n'
    '✓ Обученный оператор (обучение у специалистов Эмпилс при пуско-наладке)\n'
    '✓ KPI: не менее 2 000 кг колерованной краски в месяц\n'
    '✓ Активная работа с покупателем: предложение цвета, помощь в выборе, upsell\n'
    '✓ Ежедневно: промывка дозатора утром/вечером, контроль колорантов (≥900 мл)\n'
    '✓ 1 ЭТП от Эмпилс покрывает 120 РТТ — работает в связке с оператором',
    size=7, color=BLACK, align='left', fill=GREEN_FILL)

# === 8. РЕКОМЕНДУЕМЫЕ ОБЪЁМЫ ===
add_heading('8. РЕКОМЕНДУЕМЫЕ ОБЪЁМЫ КОЛЕРОВКИ (минимум для окупаемости)')
vol_data = [
    ['Бренд', 'Объём в месяц', 'Годовой объём'],
    ['NEWTONE', '1 500 кг', '18 000 кг/год'],
    ['AvantGarde', '2 000 кг', '24 000 кг/год'],
    ['ОРЕОЛ', '500 кг (минимум)', '6 000 кг/год'],
    ['ИТОГО с 1 машины', '4 000 кг/мес', '48 000 кг/год'],
]
add_table_simple(vol_data, [4.0, 5.5, 5.5])

table = doc.add_table(rows=1, cols=1)
table.alignment = WD_TABLE_ALIGNMENT.CENTER
table.rows[0].cells[0].width = Cm(15.0)
add_cell(table.rows[0].cells[0],
    '⚠ МИНИМУМ: 24 тонны/год (2 тонны/мес) с одной машины во втором сезоне.\n'
    'При невыполнении — Эмпилс пересматривает условия аренды (700 руб/мес).',
    bold=True, size=7, color=RED, align='center', fill=WARN_FILL)

p = doc.add_paragraph()
p.paragraph_format.space_after = Pt(1)
for run in p.runs:
    run.font.size = Pt(1)

# === 9. ФИНАНСОВЫЕ УСЛОВИЯ ===
add_heading('9. ФИНАНСОВЫЕ УСЛОВИЯ (коротко)')
fin_data = [
    ['Что', 'Условие'],
    ['Аренда оборудования', '700 руб/мес (дозатор + миксер + стойка + веера + ПО)'],
    ['Колоранты MONICOLOR PRO (16 шт.)', 'Закупка дистрибьютора по прайсу Эмпилс'],
    ['Пуско-наладка и обучение', 'Бесплатно (специалисты Эмпилс)'],
    ['Сервисное обслуживание', 'За счёт Эмпилс (при соблюдении регламента)'],
    ['Компьютер + принтер этикеток', 'За счёт дистрибьютора'],
]
add_table_simple(fin_data, [6.0, 9.0])

# === 10. КОНТАКТЫ ===
add_heading('10. КОНТАКТЫ КОЛЕРОВОЧНОЙ СЛУЖБЫ ЭМПИЛС')
contacts_data = [
    ['Специалист', 'Телефон', 'E-mail'],
    ['Козловский Станислав (гл. специалист)', '+7 (988) 556-07-86, доб. 4222', 'kozlovskiy@empils.ru'],
    ['Каретин Алексей (специалист)', '+7 (988) 589-82-35, доб. 4142', 'karetin@empils.ru'],
    ['Горячая линия', '8 (800) 200-49-11 (бесплатно по РФ)', '—'],
]
add_table_simple(contacts_data, [5.0, 6.0, 4.0])

add_text('8 (800) 200-49-11 · www.empils.ru', size=7, bold=True, color=BLUE, align='center', space_after=0)

# === СОХРАНИТЬ ===
out_path = "/home/z/my-project/download/Памятка_РМ_колеровочное_оборудование_v2.docx"
doc.save(out_path)
print(f"✓ Файл сохранён: {out_path}")
print(f"  Размер: {os.path.getsize(out_path):,} байт")
