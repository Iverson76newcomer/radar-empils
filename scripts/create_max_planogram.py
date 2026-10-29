#!/usr/bin/env python3
"""
Максимальная планограмма TITAN — все SKU, все тары, все цвета.
Источник: прайс-лист 10.07.2026 (листы ОРЕОЛ TITAN и ОРЕОЛ TITAN(л))
+ добавляем грунт-эмаль быстросохнущую графит RAL 7024 (по продажам — есть спрос).
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

OUT = '/home/z/my-project/download/Планограмма_TITAN_максимум.xlsx'

# Стили
HEADER_FONT = Font(name='Arial', size=11, bold=True, color='FFFFFF')
HEADER_FILL = PatternFill('solid', fgColor='1E3A5F')
TITLE_FONT = Font(name='Arial', size=14, bold=True, color='1E3A5F')
NORMAL_FONT = Font(name='Arial', size=10, color='1A1F36')
BOLD_FONT = Font(name='Arial', size=10, bold=True, color='1A1F36')
GRAY_FONT = Font(name='Arial', size=9, color='6B7280')
WHITE_BOLD = Font(name='Arial', size=11, bold=True, color='FFFFFF')

# Цвета категорий
CAT_3V1 = PatternFill('solid', fgColor='FEE2E2')  # розовый
CAT_FAST = PatternFill('solid', fgColor='FFEDD5')  # оранжевый
CAT_HAMMER = PatternFill('solid', fgColor='FEF9C3')  # жёлтый
NEW_FILL = PatternFill('solid', fgColor='DCFCE7')  # зелёный (новая позиция)

THIN = Side(style='thin', color='D1D5DB')
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CENTER = Alignment(horizontal='center', vertical='center', wrap_text=True)
LEFT = Alignment(horizontal='left', vertical='center', wrap_text=True)
RIGHT = Alignment(horizontal='right', vertical='center', wrap_text=True)


def style_cell(ws, row, col, value, font=NORMAL_FONT, fill=None, align=CENTER, num_format=None):
    cell = ws.cell(row, col, value)
    cell.font = font
    if fill:
        cell.fill = fill
    cell.alignment = align
    cell.border = BORDER
    if num_format:
        cell.number_format = num_format
    return cell


# === ДАННЫЕ ИЗ ПРАЙС-ЛИСТА ===

# Категория 1: Грунт-эмаль по ржавчине 3в1 (10 цветов × 5 фасовок = 50 SKU)
products_3v1 = [
    # (цвет, RAL, [(код, объём, розница, опт), ...])
    ('Белый', 'RAL 9010', [
        ('89260', '0.4 л', 403.52, 322.81, True),  # True = новинка
        ('87269', '0.75 л', 675.92, 540.73, False),
        ('87270', '2 л', 1705.01, 1364.01, False),
        ('89261', '5 л', 3524.58, 2819.66, True),
    ]),
    ('Графитовый серый', 'RAL 7024', [
        ('89262', '0.4 л', 403.52, 322.81, True),
        ('87286', '0.75 л', 616.94, 493.55, False),
        ('87287', '2 л', 1591.69, 1273.35, False),
        ('89263', '5 л', 3524.58, 2819.66, True),
    ]),
    ('Жёлтый', 'RAL 1021', [
        ('89264', '0.4 л', 403.52, 322.81, True),
        ('87272', '0.75 л', 669.57, 535.65, False),
        ('87273', '2 л', 1618.67, 1294.93, False),
        ('89265', '5 л', 3524.58, 2819.66, True),
    ]),
    ('Зелёный', 'RAL 6005', [
        ('89266', '0.4 л', 403.52, 322.81, True),
        ('87275', '0.75 л', 616.94, 493.55, False),
        ('87276', '2 л', 1591.69, 1273.35, False),
        ('89267', '5 л', 3524.58, 2819.66, True),
    ]),
    ('Красно-коричневый', 'RAL 3009', [
        ('89270', '0.4 л', 403.52, 322.81, True),
        ('87278', '0.75 л', 616.94, 493.55, False),
        ('87279', '2 л', 1591.69, 1273.35, False),
        ('89271', '5 л', 3524.58, 2819.66, True),
    ]),
    ('Серебро', 'RAL 9006', [
        ('89277', '0.4 л', 373.63, 298.90, True),
        ('89278', '0.75 л', 571.24, 456.99, True),
        ('89279', '2 л', 1578.72, 1262.97, True),
        ('89280', '5 л', 3263.50, 2610.80, True),
    ]),
    ('Серый', 'RAL 7004', [
        ('89272', '0.4 л', 403.52, 322.81, True),
        ('87281', '0.75 л', 616.94, 493.55, False),
        ('87282', '2 л', 1591.69, 1273.35, False),
        ('89273', '5 л', 3524.58, 2819.66, True),
    ]),
    ('Синий', 'RAL 5005', [
        ('89282', '0.4 л', 373.63, 298.90, True),
        ('89283', '0.75 л', 571.24, 456.99, True),
        ('89284', '2 л', 1473.86, 1179.08, True),
        ('89285', '5 л', 3263.50, 2610.80, True),
    ]),
    ('Слоновая кость', 'RAL 1014', [
        ('89287', '0.4 л', 373.63, 298.90, True),
        ('89288', '0.75 л', 571.24, 456.99, True),
        ('89289', '2 л', 1473.86, 1179.08, True),
        ('89290', '5 л', 3263.50, 2610.80, True),
    ]),
    ('Чёрный', 'RAL 9005', [
        ('89274', '0.4 л', 403.52, 322.81, True),
        ('87289', '0.75 л', 616.94, 493.55, False),
        ('87290', '2 л', 1591.69, 1273.35, False),
        ('89275', '5 л', 3524.58, 2819.66, True),
    ]),
    ('Шоколадно-коричневый', 'RAL 8017', [
        ('89268', '0.4 л', 403.52, 322.81, True),
        ('87292', '0.75 л', 616.94, 493.55, False),
        ('87293', '2 л', 1591.69, 1273.35, False),
        ('89269', '5 л', 3524.58, 2819.66, True),
    ]),
]

# Категория 2: Быстросохнущая матовая (6 цветов + графит 7024 = 7 × 2 фасовки = 14 SKU)
products_fast = [
    ('Белая', 'RAL 9010', [
        ('83760', '0.9 л', 709.71, 567.76, False),
        ('83762', '2.7 л', 2035.34, 1628.27, False),
    ]),
    ('Жёлтая', 'RAL 1023', [
        ('83764', '0.9 л', 703.06, 562.44, False),
        ('83766', '2.7 л', 1940.90, 1552.72, False),
    ]),
    ('Зелёная', 'RAL 6005', [
        ('83768', '0.9 л', 647.77, 518.22, False),
        ('83770', '2.7 л', 1857.67, 1486.13, False),
    ]),
    ('Серая', '—', [
        ('83773', '0.9 л', 647.77, 518.22, False),
        ('83775', '2.7 л', 1857.67, 1486.13, False),
    ]),
    ('Графитовая', 'RAL 7024', [  # ⚠ ДОБАВЛЕН по продажам (5 172 шт за период) — пока нет в прайсе
        ('—', '0.9 л', 647.77, 518.22, True),  # цена по аналогии с серой
        ('—', '2.7 л', 1857.67, 1486.13, True),
    ]),
    ('Чёрная', 'RAL 9005', [
        ('83777', '0.9 л', 647.77, 518.22, False),
        ('83779', '2.7 л', 1857.67, 1486.13, False),
    ]),
    ('Шоколадно-коричневая', 'RAL 8017', [
        ('83781', '0.9 л', 647.77, 518.22, False),
        ('83783', '2.7 л', 1857.67, 1486.13, False),
    ]),
]

# Категория 3: Молотковая алкидно-стирольная (6 цветов × 2 фасовки = 12 SKU)
products_hammer = [
    ('Золотистая', '—', [
        ('77639', '0.8 кг', 876.70, 701.37, False),
        ('77640', '2.5 кг', 2594.51, 2075.61, False),
    ]),
    ('Коричневая', '—', [
        ('77660', '0.8 кг', 876.70, 701.37, False),
        ('77661', '2.5 кг', 2594.51, 2075.61, False),
    ]),
    ('Медная', '—', [
        ('77654', '0.8 кг', 876.70, 701.37, False),
        ('77655', '2.5 кг', 2594.51, 2075.61, False),
    ]),
    ('Серебристая', '—', [
        ('77642', '0.8 кг', 876.70, 701.37, False),
        ('77643', '2.5 кг', 2594.51, 2075.61, False),
    ]),
    ('Тёмно-зелёная', '—', [
        ('77657', '0.8 кг', 899.58, 719.67, False),
        ('77658', '2.5 кг', 2662.20, 2129.75, False),
    ]),
    ('Чёрная', '—', [
        ('77636', '0.8 кг', 876.70, 701.37, False),
        ('77637', '2.5 кг', 2594.51, 2075.61, False),
    ]),
]


# ========================================================
# СОЗДАНИЕ EXCEL
# ========================================================
wb = openpyxl.Workbook()

# === ЛИСТ 1: ПЛАНОГРАММА ===
ws = wb.active
ws.title = 'Планограмма максимум'

# Заголовок
ws.merge_cells('A1:H1')
ws['A1'] = 'TITAN — МАКСИМАЛЬНАЯ ПЛАНОГРАММА (все SKU, все тары, все цвета)'
ws['A1'].font = TITLE_FONT
ws['A1'].alignment = CENTER
ws.row_dimensions[1].height = 25

ws.merge_cells('A2:H2')
ws['A2'] = 'Источник: прайс-лист 10.07.2026. Цены — розница с НДС. Добавлена быстросохнущая графит RAL 7024 (по продажам — есть спрос, в прайсе пока нет).'
ws['A2'].font = GRAY_FONT
ws['A2'].alignment = LEFT
ws.row_dimensions[2].height = 18

# Шапка
headers = ['№', 'Категория', 'Цвет', 'RAL', 'Код SKU', 'Объём', 'Цена розница (₽)', 'Цена опт (₽)']
for ci, h in enumerate(headers, 1):
    style_cell(ws, 4, ci, h, font=HEADER_FONT, fill=HEADER_FILL)
ws.row_dimensions[4].height = 30

# Данные
current_row = 5
sku_count = 0

def add_category(ws, start_row, cat_name, cat_fill, products):
    """Добавляет категорию с продуктами."""
    r = start_row
    sku_num = 0
    for color_name, ral, skus in products:
        for i, (code, vol, price_ret, price_opt, is_new) in enumerate(skus):
            sku_num += 1
            # Заливка: новинка или категорийная
            fill = NEW_FILL if is_new else cat_fill
            
            style_cell(ws, r, 1, sku_num, font=BOLD_FONT, fill=fill, align=CENTER)
            style_cell(ws, r, 2, cat_name if i == 0 else '', font=NORMAL_FONT, fill=fill, align=LEFT)
            style_cell(ws, r, 3, color_name if i == 0 else '', font=BOLD_FONT if i == 0 else NORMAL_FONT, fill=fill, align=LEFT)
            style_cell(ws, r, 4, ral, font=NORMAL_FONT, fill=fill, align=CENTER)
            style_cell(ws, r, 5, code, font=BOLD_FONT, fill=fill, align=CENTER)
            style_cell(ws, r, 6, vol, font=NORMAL_FONT, fill=fill, align=CENTER)
            style_cell(ws, r, 7, price_ret, font=NORMAL_FONT, fill=fill, align=RIGHT, num_format='#,##0 ₽')
            style_cell(ws, r, 8, price_opt, font=NORMAL_FONT, fill=fill, align=RIGHT, num_format='#,##0 ₽')
            r += 1
        # Пустая строка между цветами
        r += 1
    return r, sku_num

# 3в1
current_row, n1 = add_category(ws, current_row, 'Грунт-эмаль по ржавчине 3в1', CAT_3V1, products_3v1)
sku_count += n1

# Быстросохнущая
current_row, n2 = add_category(ws, current_row, 'Быстросохнущая матовая', CAT_FAST, products_fast)
sku_count += n2

# Молотковая
current_row, n3 = add_category(ws, current_row, 'Молотковая алкидно-стирольная', CAT_HAMMER, products_hammer)
sku_count += n3

# ИТОГО
current_row += 1
ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=5)
style_cell(ws, current_row, 1, f'ИТОГО: {sku_count} SKU', font=WHITE_BOLD, fill=HEADER_FILL, align=LEFT)
style_cell(ws, current_row, 6, '', fill=HEADER_FILL)
style_cell(ws, current_row, 7, '', fill=HEADER_FILL)
style_cell(ws, current_row, 8, '', fill=HEADER_FILL)

# Ширина
col_widths = [4, 28, 22, 12, 10, 10, 14, 14]
for i, w in enumerate(col_widths, 1):
    ws.column_dimensions[get_column_letter(i)].width = w


# === ЛИСТ 2: СВОДКА ПО КАТЕГОРИЯМ ===
ws2 = wb.create_sheet('Сводка по категориям')

ws2.merge_cells('A1:F1')
ws2['A1'] = 'TITAN — сводка по категориям'
ws2['A1'].font = TITLE_FONT
ws2['A1'].alignment = CENTER

# Шапка
headers2 = ['Категория', 'Цветов', 'Фасовок', 'SKU', 'Мин. цена (₽)', 'Макс. цена (₽)']
for ci, h in enumerate(headers2, 1):
    style_cell(ws2, 3, ci, h, font=HEADER_FONT, fill=HEADER_FILL)
ws2.row_dimensions[3].height = 30

# Данные
cats = [
    ('3в1 по ржавчине', 11, '0.4 / 0.75 / 2 / 5 л', n1, 322.81, 3524.58, CAT_3V1),
    ('Быстросохнущая матовая', 7, '0.9 / 2.7 л', n2, 518.22, 2035.34, CAT_FAST),
    ('Молотковая', 6, '0.8 / 2.5 кг', n3, 701.37, 2662.20, CAT_HAMMER),
]
for ri, (cat, colors, fasovki, sku, min_p, max_p, fill) in enumerate(cats, 4):
    style_cell(ws2, ri, 1, cat, font=BOLD_FONT, fill=fill, align=LEFT)
    style_cell(ws2, ri, 2, colors, font=NORMAL_FONT, fill=fill)
    style_cell(ws2, ri, 3, fasovki, font=NORMAL_FONT, fill=fill, align=LEFT)
    style_cell(ws2, ri, 4, sku, font=BOLD_FONT, fill=fill)
    style_cell(ws2, ri, 5, min_p, font=NORMAL_FONT, fill=fill, align=RIGHT, num_format='#,##0 ₽')
    style_cell(ws2, ri, 6, max_p, font=NORMAL_FONT, fill=fill, align=RIGHT, num_format='#,##0 ₽')

# ИТОГО
r = 7
style_cell(ws2, r, 1, 'ИТОГО', font=WHITE_BOLD, fill=HEADER_FILL, align=LEFT)
style_cell(ws2, r, 2, '24', font=WHITE_BOLD, fill=HEADER_FILL)
style_cell(ws2, r, 3, '—', font=WHITE_BOLD, fill=HEADER_FILL)
style_cell(ws2, r, 4, sku_count, font=WHITE_BOLD, fill=HEADER_FILL)
style_cell(ws2, r, 5, 322.81, font=WHITE_BOLD, fill=HEADER_FILL, align=RIGHT, num_format='#,##0 ₽')
style_cell(ws2, r, 6, 3524.58, font=WHITE_BOLD, fill=HEADER_FILL, align=RIGHT, num_format='#,##0 ₽')

col_widths2 = [28, 10, 28, 10, 16, 16]
for i, w in enumerate(col_widths2, 1):
    ws2.column_dimensions[get_column_letter(i)].width = w


# === ЛИСТ 3: ПО ЦВЕТАМ ===
ws3 = wb.create_sheet('По цветам')

ws3.merge_cells('A1:D1')
ws3['A1'] = 'TITAN — все цвета по категориям'
ws3['A1'].font = TITLE_FONT
ws3['A1'].alignment = CENTER

headers3 = ['Категория', 'Цвет', 'RAL', 'Кол-во SKU']
for ci, h in enumerate(headers3, 1):
    style_cell(ws3, 3, ci, h, font=HEADER_FONT, fill=HEADER_FILL)

r = 4
for color_name, ral, skus in products_3v1:
    style_cell(ws3, r, 1, '3в1 по ржавчине', font=NORMAL_FONT, fill=CAT_3V1, align=LEFT)
    style_cell(ws3, r, 2, color_name, font=BOLD_FONT, fill=CAT_3V1, align=LEFT)
    style_cell(ws3, r, 3, ral, font=NORMAL_FONT, fill=CAT_3V1, align=CENTER)
    style_cell(ws3, r, 4, len(skus), font=NORMAL_FONT, fill=CAT_3V1)
    r += 1

for color_name, ral, skus in products_fast:
    style_cell(ws3, r, 1, 'Быстросохнущая', font=NORMAL_FONT, fill=CAT_FAST, align=LEFT)
    style_cell(ws3, r, 2, color_name, font=BOLD_FONT, fill=CAT_FAST, align=LEFT)
    style_cell(ws3, r, 3, ral, font=NORMAL_FONT, fill=CAT_FAST, align=CENTER)
    style_cell(ws3, r, 4, len(skus), font=NORMAL_FONT, fill=CAT_FAST)
    r += 1

for color_name, ral, skus in products_hammer:
    style_cell(ws3, r, 1, 'Молотковая', font=NORMAL_FONT, fill=CAT_HAMMER, align=LEFT)
    style_cell(ws3, r, 2, color_name, font=BOLD_FONT, fill=CAT_HAMMER, align=LEFT)
    style_cell(ws3, r, 3, ral, font=NORMAL_FONT, fill=CAT_HAMMER, align=CENTER)
    style_cell(ws3, r, 4, len(skus), font=NORMAL_FONT, fill=CAT_HAMMER)
    r += 1

col_widths3 = [25, 25, 15, 12]
for i, w in enumerate(col_widths3, 1):
    ws3.column_dimensions[get_column_letter(i)].width = w


# Сохранение
wb.save(OUT)
print(f'OK: {OUT}')
import os
print(f'Размер: {os.path.getsize(OUT) / 1024:.1f} KB')
print(f'\nЛисты:')
for sn in wb.sheetnames:
    print(f'  "{sn}"')
print(f'\nВсего SKU: {sku_count}')
print(f'  3в1 по ржавчине: {n1}')
print(f'  Быстросохнущая: {n2} (включая графит 7024 — 3 новых SKU)')
print(f'  Молотковая: {n3}')
