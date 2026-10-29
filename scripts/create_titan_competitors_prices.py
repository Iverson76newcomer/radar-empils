#!/usr/bin/env python3
"""
Excel-шаблон: Конкурентное окружение TITAN — цены из презентации + колонки для актуальных цен.
Источник: презентация TITAN 10.06.2026_грунт-эмаль.pptx, слайды 11-13.
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

OUT = '/home/z/my-project/download/TITAN_конкуренты_цены_ВсеИнструменты.xlsx'

# Стили
HEADER_FONT = Font(name='Arial', size=11, bold=True, color='FFFFFF')
HEADER_FILL = PatternFill('solid', fgColor='1E3A5F')
TITLE_FONT = Font(name='Arial', size=14, bold=True, color='1E3A5F')
SUBTITLE_FONT = Font(name='Arial', size=10, italic=True, color='6B7280')
NORMAL_FONT = Font(name='Arial', size=10, color='1A1F36')
BOLD_FONT = Font(name='Arial', size=10, bold=True, color='1A1F36')
RED_FONT = Font(name='Arial', size=10, bold=True, color='C62828')
GREEN_FONT = Font(name='Arial', size=10, bold=True, color='16A34A')
ORANGE_FONT = Font(name='Arial', size=10, bold=True, color='F47920')
GRAY_FONT = Font(name='Arial', size=9, color='6B7280')
WHITE_BOLD = Font(name='Arial', size=11, bold=True, color='FFFFFF')

LIGHT_GRAY = PatternFill('solid', fgColor='F3F4F6')
LIGHT_RED = PatternFill('solid', fgColor='FEE2E2')
LIGHT_GREEN = PatternFill('solid', fgColor='DCFCE7')
LIGHT_YELLOW = PatternFill('solid', fgColor='FEF9C3')
LIGHT_ORANGE = PatternFill('solid', fgColor='FFEDD5')
LIGHT_BLUE = PatternFill('solid', fgColor='DBEAFE')

TITAN_FILL = PatternFill('solid', fgColor='FEF3C7')  # выделяем TITAN жёлтым

THIN = Side(style='thin', color='D1D5DB')
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CENTER = Alignment(horizontal='center', vertical='center', wrap_text=True)
LEFT = Alignment(horizontal='left', vertical='center', wrap_text=True)
RIGHT = Alignment(horizontal='right', vertical='center', wrap_text=True)


def style_cell(ws, row, col, value, font=NORMAL_FONT, fill=None, align=CENTER, border=BORDER, num_format=None):
    cell = ws.cell(row, col, value)
    cell.font = font
    if fill:
        cell.fill = fill
    cell.alignment = align
    cell.border = border
    if num_format:
        cell.number_format = num_format
    return cell


# ========================================================
# ДАННЫЕ ИЗ ПРЕЗЕНТАЦИИ TITAN 10.06.2026 (слайды 11-13)
# Источник цен: сайт ВсеИнструменты.ру на момент 10.06.2026
# ========================================================

products = [
    # === Слайд 11. TITAN 3в1 ПО РЖАВЧИНЕ ГЛЯНЦЕВАЯ (чёрный RAL 9005) ===
    {
        'cat': '3в1 по ржавчине (глянцевая)',
        'brand': 'TITAN',
        'product': 'Грунт-эмаль по ржавчине Empils ОРЕОЛ TITAN чёрный RAL 9005',
        'vol': '0.75 л (0.75 кг)',
        'ral': 'RAL 9005 (чёрный)',
        'price_old': 872,
        'search_query': 'ОРЕОЛ TITAN чёрный RAL 9005 0.75',
    },
    {
        'cat': '3в1 по ржавчине (глянцевая)',
        'brand': 'Elcon',
        'product': 'Грунт-эмаль 3 в 1 Elcon глянцевая чёрная RAL 9005',
        'vol': '0.8 кг',
        'ral': 'RAL 9005 (чёрный)',
        'price_old': 958,
        'search_query': 'Elcon грунт-эмаль 3 в 1 глянцевая RAL 9005 0.8 кг',
    },
    {
        'cat': '3в1 по ржавчине (глянцевая)',
        'brand': 'Dufa',
        'product': 'Грунт-эмаль Dufa CORREX 3 в 1 для металла по ржавчине RAL 9005',
        'vol': '0.75 л',
        'ral': 'RAL 9005 (чёрный)',
        'price_old': 1020,
        'search_query': 'Dufa CORREX 3 в 1 RAL 9005 0.75',
    },
    {
        'cat': '3в1 по ржавчине (глянцевая)',
        'brand': 'Dali',
        'product': 'Грунт-эмаль по ржавчине Dali 3 в 1, чёрная RAL 9005',
        'vol': '0.75 л',
        'ral': 'RAL 9005 (чёрный)',
        'price_old': 1055,
        'search_query': 'Dali 3 в 1 грунт-эмаль RAL 9005 0.75',
    },
    {
        'cat': '3в1 по ржавчине (глянцевая)',
        'brand': 'HUSKY KLONDIKE',
        'product': 'Краска по металлу HUSKY KLONDIKE глянцевая, чёрная, RAL 9005',
        'vol': '0.9 л',
        'ral': 'RAL 9005 (чёрный)',
        'price_old': 1153,
        'search_query': 'HUSKY KLONDIKE краска по металлу RAL 9005 0.9',
    },
    {
        'cat': '3в1 по ржавчине (глянцевая)',
        'brand': 'Tikkurila',
        'product': 'Краска для металла Tikkurila METALLISTA по ржавчине 3 в 1, чёрная',
        'vol': '0.8 л',
        'ral': 'чёрный',
        'price_old': 2085,
        'search_query': 'Tikkurila METALLISTA 3 в 1 0.8 чёрная',
    },
    # === Слайд 12. TITAN БЫСТРОСОХНУЩАЯ МАТОВАЯ (серый RAL 7040 / графит RAL 7024) ===
    {
        'cat': 'Быстросохнущая (матовая)',
        'brand': 'Новбытхим',
        'product': 'Грунт-эмаль 5 в 1 Новбытхим, серый, RAL 7040',
        'vol': '0.75 л',
        'ral': 'RAL 7040 (серый)',
        'price_old': 765,
        'search_query': 'Новбытхим грунт-эмаль 5 в 1 RAL 7040 0.75',
    },
    {
        'cat': 'Быстросохнущая (матовая)',
        'brand': 'TITAN',
        'product': 'Грунт-эмаль быстросохнущая ОРЕОЛ TITAN матовая серая',
        'vol': '0.75 л (0.9 кг)',
        'ral': 'серый',
        'price_old': 765,
        'search_query': 'ОРЕОЛ TITAN быстросохнущая матовая серая 0.75',
    },
    {
        'cat': 'Быстросохнущая (матовая)',
        'brand': 'Elcon',
        'product': 'Быстросохнущая грунт-эмаль по металлу Elcon 3 в 1 по ржавчине, матовая, серая',
        'vol': '0.8 кг',
        'ral': 'серый',
        'price_old': 838,
        'search_query': 'Elcon быстросохнущая 3 в 1 матовая серая 0.8',
    },
    {
        'cat': 'Быстросохнущая (матовая)',
        'brand': 'Лакра',
        'product': 'Грунт-эмаль по ржавчине ЛАКРА PROF IT 4 в 1, быстросохнущая, графит, RAL 7024',
        'vol': '0.8 кг',
        'ral': 'RAL 7024 (графит)',
        'price_old': 906,
        'search_query': 'Лакра PROF IT 4 в 1 быстросохнущая RAL 7024 0.8',
    },
    # === Слайд 13. TITAN МОЛОТКОВАЯ (золотистая) ===
    {
        'cat': 'Молотковая',
        'brand': 'Dali',
        'product': 'Грунт-эмаль по ржавчине Dali 3 в 1 молотковая золотистая',
        'vol': '0.75 л',
        'ral': 'золотистый',
        'price_old': 1057,
        'search_query': 'Dali молотковая золотистая 3 в 1 0.75',
    },
    {
        'cat': 'Молотковая',
        'brand': 'Elcon',
        'product': 'Кузнечная краска с молотковым эффектом Elcon золото',
        'vol': '0.8 кг',
        'ral': 'золото',
        'price_old': 1097,
        'search_query': 'Elcon кузнечная краска молотковая золото 0.8',
    },
    {
        'cat': 'Молотковая',
        'brand': 'HUSKY KLONDIKE',
        'product': 'Краска по металлу с молотковым эффектом HUSKY KLONDIKE золото',
        'vol': '0.9 л',
        'ral': 'золото',
        'price_old': 1196,
        'search_query': 'HUSKY KLONDIKE молотковая золото 0.9',
    },
    {
        'cat': 'Молотковая',
        'brand': 'TITAN',
        'product': 'Эмаль с молотковым эффектом Titan алкидно-стирольная, золотистая',
        'vol': '0.8 кг',
        'ral': 'золотистый',
        'price_old': 1308,
        'search_query': 'TITAN молотковая алкидно-стирольная золотистая 0.8',
    },
    {
        'cat': 'Молотковая',
        'brand': 'Лакра',
        'product': 'Эмаль с молотковым эффектом Лакра золото',
        'vol': '0.8 кг',
        'ral': 'золото',
        'price_old': 1341,
        'search_query': 'Лакра молотковая золото 0.8',
    },
    {
        'cat': 'Молотковая',
        'brand': 'Dufa',
        'product': 'Эмаль Dufa Premium HAMMERLACK молотковая, золотая',
        'vol': '0.75 л (750 мл)',
        'ral': 'золотая',
        'price_old': 1653,
        'search_query': 'Dufa Premium HAMMERLACK молотковая золотая 750 мл',
    },
]


# ========================================================
# СОЗДАНИЕ EXCEL
# ========================================================
wb = openpyxl.Workbook()

# ========================================================
# ЛИСТ 1: ЦЕНЫ — БЫЛО / СТАЛО
# ========================================================
ws = wb.active
ws.title = 'Цены было-стало'

# Заголовок
ws.merge_cells('A1:K1')
ws['A1'] = 'TITAN — конкурентное окружение: цены ВсеПлитрументы.ру (было / стало)'
ws['A1'].font = TITLE_FONT
ws['A1'].alignment = CENTER
ws.row_dimensions[1].height = 25

ws.merge_cells('A2:K2')
ws['A2'] = ('Источник «было»: презентация «TITAN 10.06.2026_грунт-эмаль.pptx», слайды 11-13. '
            'Колонка «стало» — заполнить актуальными ценами с vseinstrumenti.ru. '
            'Цены — за 1 шт. в рознице, с НДС.')
ws['A2'].font = SUBTITLE_FONT
ws['A2'].alignment = LEFT
ws.row_dimensions[2].height = 30

# Заголовки таблицы
headers = [
    '№',
    'Категория TITAN',
    'Бренд',
    'Продукт',
    'Объём',
    'RAL / цвет',
    'Цена 10.06.2026 (₽)',
    'Цена актуальная (₽)',
    'Δ цены (₽)',
    'Δ цены (%)',
    'Поисковый запрос для vseinstrumenti.ru',
]
for ci, h in enumerate(headers, 1):
    style_cell(ws, 4, ci, h, font=HEADER_FONT, fill=HEADER_FILL)
ws.row_dimensions[4].height = 50

# Данные
for ri, p in enumerate(products, 1):
    r = 4 + ri
    is_titan = p['brand'] == 'TITAN'
    fill = TITAN_FILL if is_titan else (LIGHT_BLUE if ri % 2 == 0 else LIGHT_GRAY)
    
    style_cell(ws, r, 1, ri, font=BOLD_FONT, fill=fill, align=CENTER)
    style_cell(ws, r, 2, p['cat'], font=NORMAL_FONT, fill=fill, align=LEFT)
    style_cell(ws, r, 3, p['brand'], font=BOLD_FONT if is_titan else NORMAL_FONT, fill=fill, align=LEFT)
    style_cell(ws, r, 4, p['product'], font=NORMAL_FONT, fill=fill, align=LEFT)
    style_cell(ws, r, 5, p['vol'], font=NORMAL_FONT, fill=fill, align=CENTER)
    style_cell(ws, r, 6, p['ral'], font=NORMAL_FONT, fill=fill, align=CENTER)
    style_cell(ws, r, 7, p['price_old'], font=BOLD_FONT, fill=fill, align=RIGHT, num_format='#,##0 ₽')
    # Колонка "Цена актуальная" — пустая, для заполнения
    style_cell(ws, r, 8, '', font=BOLD_FONT, fill=LIGHT_YELLOW, align=RIGHT, num_format='#,##0 ₽')
    # Δ цены — формула
    ws.cell(r, 9).value = f'=IF(H{r}="","",H{r}-G{r})'
    ws.cell(r, 9).font = BOLD_FONT
    ws.cell(r, 9).fill = fill
    ws.cell(r, 9).alignment = RIGHT
    ws.cell(r, 9).border = BORDER
    ws.cell(r, 9).number_format = '+#,##0 ₽;-#,##0 ₽;0 ₽'
    # Δ % — формула
    ws.cell(r, 10).value = f'=IF(H{r}="","",IF(G{r}=0,"",(H{r}-G{r})/G{r}))'
    ws.cell(r, 10).font = BOLD_FONT
    ws.cell(r, 10).fill = fill
    ws.cell(r, 10).alignment = RIGHT
    ws.cell(r, 10).border = BORDER
    ws.cell(r, 10).number_format = '+0%;-0%;0%'
    # Поисковый запрос
    style_cell(ws, r, 11, p['search_query'], font=GRAY_FONT, fill=fill, align=LEFT)

# ИТОГО — средние цены
r_total = 4 + len(products) + 1
ws.merge_cells(start_row=r_total, start_column=1, end_row=r_total, end_column=6)
style_cell(ws, r_total, 1, 'СРЕДНЯЯ ЦЕНА ПО ВСЕМ SKU', font=WHITE_BOLD, fill=HEADER_FILL, align=LEFT)
ws.cell(r_total, 7).value = f'=AVERAGE(G5:G{4+len(products)})'
ws.cell(r_total, 8).value = f'=IFERROR(AVERAGE(H5:H{4+len(products)}),"—")'
for ci in [7, 8]:
    ws.cell(r_total, ci).font = WHITE_BOLD
    ws.cell(r_total, ci).fill = HEADER_FILL
    ws.cell(r_total, ci).alignment = RIGHT
    ws.cell(r_total, ci).border = BORDER
    ws.cell(r_total, ci).number_format = '#,##0 ₽'
ws.cell(r_total, 9).value = f'=IF(H{r_total}="—","",H{r_total}-G{r_total})'
ws.cell(r_total, 10).value = f'=IF(H{r_total}="—","",IF(G{r_total}=0,"",(H{r_total}-G{r_total})/G{r_total}))'
for ci in [9, 10]:
    ws.cell(r_total, ci).font = WHITE_BOLD
    ws.cell(r_total, ci).fill = HEADER_FILL
    ws.cell(r_total, ci).alignment = RIGHT
    ws.cell(r_total, ci).border = BORDER
ws.cell(r_total, 9).number_format = '+#,##0 ₽;-#,##0 ₽;0 ₽'
ws.cell(r_total, 10).number_format = '+0%;-0%;0%'
style_cell(ws, r_total, 11, '', fill=HEADER_FILL)

# ИТОГО по TITAN
r_titan = r_total + 1
ws.merge_cells(start_row=r_titan, start_column=1, end_row=r_titan, end_column=6)
style_cell(ws, r_titan, 1, 'СРЕДНЯЯ ЦЕНА TITAN (3 SKU)', font=Font(name='Arial', size=11, bold=True, color='1E3A5F'), fill=TITAN_FILL, align=LEFT)
# TITAN — это строки 5, 12, 17 (1-я в каждой категории). Найдём их
titan_rows = [4 + i + 1 for i, p in enumerate(products) if p['brand'] == 'TITAN']
titan_g = ','.join(f'G{r}' for r in titan_rows)
titan_h = ','.join(f'H{r}' for r in titan_rows)
ws.cell(r_titan, 7).value = f'=AVERAGE({titan_g})'
ws.cell(r_titan, 8).value = f'=IFERROR(AVERAGE({titan_h}),"—")'
for ci in [7, 8]:
    ws.cell(r_titan, ci).font = Font(name='Arial', size=11, bold=True, color='1E3A5F')
    ws.cell(r_titan, ci).fill = TITAN_FILL
    ws.cell(r_titan, ci).alignment = RIGHT
    ws.cell(r_titan, ci).border = BORDER
    ws.cell(r_titan, ci).number_format = '#,##0 ₽'
ws.cell(r_titan, 9).value = f'=IF(H{r_titan}="—","",H{r_titan}-G{r_titan})'
ws.cell(r_titan, 10).value = f'=IF(H{r_titan}="—","",IF(G{r_titan}=0,"",(H{r_titan}-G{r_titan})/G{r_titan}))'
for ci in [9, 10]:
    ws.cell(r_titan, ci).font = Font(name='Arial', size=11, bold=True, color='1E3A5F')
    ws.cell(r_titan, ci).fill = TITAN_FILL
    ws.cell(r_titan, ci).alignment = RIGHT
    ws.cell(r_titan, ci).border = BORDER
ws.cell(r_titan, 9).number_format = '+#,##0 ₽;-#,##0 ₽;0 ₽'
ws.cell(r_titan, 10).number_format = '+0%;-0%;0%'
style_cell(ws, r_titan, 11, '', fill=TITAN_FILL)

# Ширина колонок
col_widths = [4, 24, 16, 50, 14, 18, 14, 14, 12, 10, 35]
for i, w in enumerate(col_widths, 1):
    ws.column_dimensions[get_column_letter(i)].width = w


# ========================================================
# ЛИСТ 2: ИНСТРУКЦИЯ
# ========================================================
ws2 = wb.create_sheet('Инструкция')

ws2.merge_cells('A1:F1')
ws2['A1'] = 'Как заполнить колонку «Цена актуальная»'
ws2['A1'].font = TITLE_FONT
ws2['A1'].alignment = LEFT
ws2.row_dimensions[1].height = 25

instructions = [
    ('Шаг 1.', 'Откройте сайт https://vseinstrumenti.ru/ в браузере'),
    ('Шаг 2.', 'В поиске введите текст из колонки «Поисковый запрос» (последняя колонка таблицы)'),
    ('Шаг 3.', 'Найдите в выдаче товар, совпадающий по бренду, объёму и цвету'),
    ('Шаг 4.', 'Скопируйте цену (за 1 шт., с НДС, без учёта акций) в колонку «Цена актуальная (₽)»'),
    ('Шаг 5.', 'Δ цены и Δ % посчитаются автоматически по формуле'),
    ('', ''),
    ('⚠️ Важно:', 'Если товара нет в наличии — отметьте в колонке «Заметки» (на этом листе нет, можно добавить)'),
    ('⚠️ Важно:', 'Если цена акционная (со скидкой) — берите обычную цену, а не акционную'),
    ('⚠️ Важно:', 'Если у TITAN цена изменилась — это тоже вносим, не только у конкурентов'),
    ('⚠️ Важно:', 'Все цены — за 1 шт., не за набор и не за упаковку'),
    ('', ''),
    ('Срок актуальности:', 'Цены собраны 10.06.2026. Новые цены — текущая дата заполнения'),
    ('Источник:', 'Презентация TITAN 10.06.2026_грунт-эмаль.pptx, слайды 11-13'),
]

for ri, (label, text) in enumerate(instructions, 3):
    style_cell(ws2, ri, 1, label, font=BOLD_FONT, fill=LIGHT_GRAY, align=LEFT)
    ws2.merge_cells(start_row=ri, start_column=2, end_row=ri, end_column=6)
    style_cell(ws2, ri, 2, text, font=NORMAL_FONT, fill=LIGHT_GRAY, align=LEFT)

col_widths2 = [16, 25, 25, 25, 25, 25]
for i, w in enumerate(col_widths2, 1):
    ws2.column_dimensions[get_column_letter(i)].width = w


# Сохранение
wb.save(OUT)
print(f'OK: {OUT}')
import os
print(f'Размер: {os.path.getsize(OUT) / 1024:.1f} KB')
print(f'\nЛисты:')
for sn in wb.sheetnames:
    print(f'  "{sn}"')
print(f'\nВсего SKU: {len(products)}')
print(f'Из них TITAN: {sum(1 for p in products if p["brand"] == "TITAN")}')
print(f'Конкурентов: {sum(1 for p in products if p["brand"] != "TITAN")}')
