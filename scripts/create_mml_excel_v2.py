#!/usr/bin/env python3
"""
Excel: MML TITAN — Было → Стало + продажи.
3 листа:
1. Сравнение MML (было/стало по классам + продажи)
2. Продажи по цветам (вторичка vs MML)
3. Сводка по классам (с колонкой "Текущие позиции MML" + обоснование с разбивкой по цветам)

v2: переработана "Сводка по классам" — добавлена колонка текущих позиций MML
    и в "Обосновании" продажи привязаны к конкретным цветам.
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

OUT = '/home/z/my-project/download/MML_TITAN_с_продажами_v2.xlsx'  # v2 без добавлений в D

# Стили
header_font = Font(name='Arial', size=10, bold=True, color='FFFFFF')
header_fill = PatternFill('solid', fgColor='1E3A5F')
title_font = Font(name='Arial', size=14, bold=True, color='1E3A5F')
subtitle_font = Font(name='Arial', size=11, bold=True, color='1E3A5F')
normal_font = Font(name='Arial', size=10, color='1A1F36')
bold_font = Font(name='Arial', size=10, bold=True, color='1A1F36')
red_font = Font(name='Arial', size=10, bold=True, color='C62828')
green_font = Font(name='Arial', size=10, bold=True, color='16A34A')
orange_font = Font(name='Arial', size=10, bold=True, color='F47920')
gray_font = Font(name='Arial', size=9, color='6B7280')
white_bold = Font(name='Arial', size=11, bold=True, color='FFFFFF')
small_font = Font(name='Arial', size=9, color='1A1F36')
small_bold = Font(name='Arial', size=9, bold=True, color='1A1F36')

light_gray = PatternFill('solid', fgColor='F3F4F6')
light_red = PatternFill('solid', fgColor='FEE2E2')
light_green = PatternFill('solid', fgColor='DCFCE7')
light_orange = PatternFill('solid', fgColor='FFEDD5')
light_blue = PatternFill('solid', fgColor='DBEAFE')
yellow_fill = PatternFill('solid', fgColor='FEF3C7')
light_yellow = PatternFill('solid', fgColor='FEF9C3')

thin = Side(style='thin', color='D1D5DB')
border = Border(left=thin, right=thin, top=thin, bottom=thin)

center = Alignment(horizontal='center', vertical='center', wrap_text=True)
left = Alignment(horizontal='left', vertical='center', wrap_text=True)
left_top = Alignment(horizontal='left', vertical='top', wrap_text=True)
right = Alignment(horizontal='right', vertical='center', wrap_text=True)


def style_cell(ws, row, col, value, font=normal_font, fill=None, align=center, border_style=border, num_format=None):
    cell = ws.cell(row, col, value)
    cell.font = font
    if fill:
        cell.fill = fill
    cell.alignment = align
    cell.border = border_style
    if num_format:
        cell.number_format = num_format
    return cell


wb = openpyxl.Workbook()

# ==============================================================
# ЛИСТ 1: СРАВНЕНИЕ MML — БЫЛО / СТАЛО + ПРОДАЖИ
# ==============================================================
ws1 = wb.active
ws1.title = 'MML сравнение'

ws1.merge_cells('A1:L1')
ws1['A1'] = 'MML TITAN — СРАВНЕНИЕ: БЫЛО → СТАЛО (с привязкой к продажам)'
ws1['A1'].font = title_font
ws1['A1'].alignment = center
ws1.row_dimensions[1].height = 30

ws1.merge_cells('A2:L2')
ws1['A2'] = 'Источник: MML скорректированный 28.06.2026 + «Титан цвета» — вторичные продажи'
ws1['A2'].font = gray_font
ws1['A2'].alignment = center

headers = [
    'Код SKU', 'Категория', 'Цвет', 'RAL', 'Объём',
    'Цена Эмпилс (₽)', 'Вторичка (шт)', 'АКБ',
    'MML: A', 'MML: B', 'MML: C', 'MML: D',
]
for col, h in enumerate(headers, 1):
    style_cell(ws1, 4, col, h, font=header_font, fill=header_fill)
ws1.row_dimensions[4].height = 35

all_sku = [
    # 3в1 ПО РЖАВЧИНЕ — ТЕКУЩИЕ
    ('87269', '3в1 по ржавчине', 'Белый', 'RAL 9010', '0.75 л', 541, 3273, 225, '✅', '✅', '✅', '✅', '—', '—', '—', '—'),
    ('87270', '3в1 по ржавчине', 'Белый', 'RAL 9010', '2 л', 1364, 0, 0, '✅', '✅', '—', '—', '—', '—', '—', '—'),
    ('87281', '3в1 по ржавчине', 'Серый', 'RAL 7004', '0.75 л', 494, 5071, 247, '✅', '✅', '✅', '✅', '—', '🆕', '—', '—'),
    ('87282', '3в1 по ржавчине', 'Серый', 'RAL 7004', '2 л', 1273, 0, 0, '✅', '✅', '✅', '✅', '—', '—', '—', '—'),
    ('87286', '3в1 по ржавчине', 'Графитовый', 'RAL 7024', '0.75 л', 494, 5172, 240, '✅', '✅', '✅', '✅', '—', '—', '—', '—'),
    ('87287', '3в1 по ржавчине', 'Графитовый', 'RAL 7024', '2 л', 1273, 0, 0, '✅', '✅', '✅', '✅', '—', '—', '—', '—'),
    ('87289', '3в1 по ржавчине', 'Чёрный', 'RAL 9005', '0.75 л', 494, 7738, 393, '✅', '✅', '✅', '✅', '—', '🆕', '—', '—'),  # без D
    ('87290', '3в1 по ржавчине', 'Чёрный', 'RAL 9005', '2 л', 1273, 0, 0, '✅', '✅', '✅', '✅', '—', '—', '—', '—'),
    ('87292', '3в1 по ржавчине', 'Шоколад', 'RAL 8017', '0.75 л', 494, 4524, 287, '✅', '✅', '✅', '✅', '—', '🆕', '—', '—'),  # без D
    ('87293', '3в1 по ржавчине', 'Шоколад', 'RAL 8017', '2 л', 1273, 0, 0, '✅', '✅', '✅', '✅', '—', '—', '—', '—'),
    # 3в1 — РЕКОМЕНДОВАННЫЕ (НЕТ В MML)
    ('87275', '3в1 по ржавчине', 'Зелёный', 'RAL 6005', '0.75 л', 457, 1893, 111, '—', '🆕', '—', '🆕', '—', '—', '—', '—'),
    # БЫСТРОСОХНУЩАЯ — ТЕКУЩИЕ
    ('83777', 'Быстросохнущая', 'Чёрный', 'RAL 9005', '0.75 л', 518, 5381, 303, '✅', '✅', '✅', '✅', '—', '🆕', '—', '—'),
    ('83779', 'Быстросохнущая', 'Чёрный', 'RAL 9005', '2.2 л', 1486, 0, 0, '✅', '✅', '—', '—', '—', '—', '—', '—'),
    # БЫСТРОСОХНУЩАЯ — РЕКОМЕНДОВАННЫЕ
    ('—', 'Быстросохнущая', 'Серый', 'RAL 7004', '0.75 л', 0, 2695, 199, '—', '🆕', '—', '🆕', '—', '🆕', '—', '—'),
    ('—', 'Быстросохнущая', 'Шоколад', 'RAL 8017', '0.75 л', 0, 2232, 193, '—', '🆕', '—', '🆕', '—', '—', '—', '—'),
    ('—', 'Быстросохнущая', 'Белый', 'RAL 9010', '0.75 л', 0, 1507, 119, '—', '🆕', '—', '—', '—', '—', '—', '—'),
    # МОЛОТКОВАЯ — ТЕКУЩИЕ
    ('77636', 'Молотковая', 'Чёрный', '—', '0.8 кг', 701, 3536, 243, '✅', '✅', '✅', '✅', '✅', '✅', '—', '—'),
    ('77637', 'Молотковая', 'Чёрный', '—', '2.5 кг', 2076, 0, 0, '✅', '✅', '—', '—', '—', '—', '—', '—'),
    ('77642', 'Молотковая', 'Серебристый', '—', '0.8 кг', 701, 2394, 198, '✅', '✅', '—', '🆕', '—', '—', '—', '—'),
    ('77660', 'Молотковая', 'Коричневый', '—', '0.8 кг', 701, 2913, 207, '✅', '✅', '✅', '✅', '✅', '✅', '—', '—'),
    # МОЛОТКОВАЯ — РЕКОМЕНДОВАННЫЕ
    ('—', 'Молотковая', 'Медный', '—', '0.8 кг', 0, 1162, 102, '—', '🆕', '—', '—', '—', '—', '—', '—'),
]

row = 5
current_cat = ''
for sku in all_sku:
    code, cat, color, ral, vol, price, qty, akb, a_was, a_now, b_was, b_now, c_was, c_now, d_was, d_now = sku
    if cat != current_cat:
        current_cat = cat
        if row > 5:
            row += 1
    cat_bg = {
        '3в1 по ржавчине': light_red,
        'Быстросохнущая': light_orange,
        'Молотковая': light_yellow,
    }.get(cat, light_gray)
    style_cell(ws1, row, 1, code, font=bold_font, fill=cat_bg, align=center)
    style_cell(ws1, row, 2, cat, font=normal_font, fill=cat_bg, align=left)
    style_cell(ws1, row, 3, color, font=normal_font, fill=cat_bg, align=left)
    style_cell(ws1, row, 4, ral, font=normal_font, fill=cat_bg, align=center)
    style_cell(ws1, row, 5, vol, font=normal_font, fill=cat_bg, align=center)
    style_cell(ws1, row, 6, price if price else '—', font=normal_font, fill=cat_bg, align=right, num_format='#,##0')
    style_cell(ws1, row, 7, qty if qty else '—', font=bold_font if qty > 1000 else normal_font, fill=cat_bg, align=right, num_format='#,##0')
    style_cell(ws1, row, 8, akb if akb else '—', font=normal_font, fill=cat_bg, align=center)
    for ci, (was, now) in enumerate([(a_was, a_now), (b_was, b_now), (c_was, c_now), (d_was, d_now)]):
        col = 9 + ci
        if was == '✅' and now == '✅':
            style_cell(ws1, row, col, '✅', font=green_font, fill=light_green, align=center)
        elif was == '—' and now == '🆕':
            style_cell(ws1, row, col, '🆕', font=red_font, fill=light_red, align=center)
        elif was == '✅' and now == '🆕':
            style_cell(ws1, row, col, '✅→🆕', font=orange_font, fill=light_orange, align=center)
        else:
            style_cell(ws1, row, col, '—', font=gray_font, fill=cat_bg, align=center)
    row += 1

# ИТОГО
row += 1
ws1.merge_cells(start_row=row, start_column=1, end_row=row, end_column=5)
style_cell(ws1, row, 1, 'ИТОГО', font=white_bold, fill=header_fill, align=center)
style_cell(ws1, row, 6, '', fill=header_fill)
style_cell(ws1, row, 7, '', fill=header_fill)
style_cell(ws1, row, 8, '', fill=header_fill)
for ci, (was, now) in enumerate([('16', '21'), ('12', '16'), ('2', '7'), ('0', '0')]):  # D без изменений
    col = 9 + ci
    style_cell(ws1, row, col, f'{was} → {now}', font=white_bold, fill=header_fill, align=center)

col_widths = [10, 18, 15, 10, 8, 12, 12, 8, 10, 10, 10, 10]
for i, w in enumerate(col_widths, 1):
    ws1.column_dimensions[get_column_letter(i)].width = w


# ==============================================================
# ЛИСТ 2: ПРОДАЖИ ПО ЦВЕТАМ — ВТОРИЧКА vs MML
# ==============================================================
ws2 = wb.create_sheet('Продажи vs MML')

ws2.merge_cells('A1:H1')
ws2['A1'] = 'ВТОРИЧНЫЕ ПРОДАЖИ TITAN — РАНГ ПО ШТУКАМ vs MML'
ws2['A1'].font = title_font
ws2['A1'].alignment = center
ws2.row_dimensions[1].height = 30

ws2.merge_cells('A2:H2')
ws2['A2'] = 'Сортировка: от самых продаваемых к наименее. Видно, какие цвета формируют MML.'
ws2['A2'].font = gray_font
ws2['A2'].alignment = center

headers2 = ['Ранг', 'Категория', 'Цвет', 'RAL', 'Продажи (шт)', 'АКБ', 'В MML?', 'Классы MML']
for col, h in enumerate(headers2, 1):
    style_cell(ws2, 4, col, h, font=header_font, fill=header_fill)
ws2.row_dimensions[4].height = 30

sales_data = [
    (1, '3в1 по ржавчине', 'Чёрный', 'RAL 9005', 7738, 393, '✅ ДА', 'A, B'),
    (2, '3в1 по ржавчине', 'Графитовый', 'RAL 7024', 5172, 240, '✅ ДА', 'A, B'),
    (3, '3в1 по ржавчине', 'Серый', 'RAL 7004', 5071, 247, '✅ ДА', 'A, B'),
    (4, '3в1 по ржавчине', 'Шоколад', 'RAL 8017', 4524, 287, '✅ ДА', 'A, B'),
    (5, '3в1 по ржавчине', 'Белый', 'RAL 9010', 3273, 225, '✅ ДА', 'A, B'),
    (6, '3в1 по ржавчине', 'Зелёный', 'RAL 6005', 1893, 111, '⚠️ НЕТ', '—'),
    (7, '3в1 по ржавчине', 'Серый (без RAL)', '—', 1655, 154, '⚠️ НЕТ', '—'),
    (8, '3в1 по ржавчине', 'Жёлтый', 'RAL 1021', 1276, 80, '⚠️ НЕТ', '—'),
    (9, '3в1 по ржавчине', 'Красно-корич.', '—', 633, 67, '⚠️ НЕТ', '—'),
    (10, '3в1 по ржавчине', 'Красно-корич.', 'RAL 3009', 570, 49, '⚠️ НЕТ', '—'),
    (11, 'Быстросохнущая', 'Чёрный', 'RAL 9005', 5381, 303, '✅ ДА', 'A, B'),
    (12, 'Быстросохнущая', 'Серый', 'RAL 7004', 2695, 199, '⚠️ НЕТ', '—'),
    (13, 'Быстросохнущая', 'Шоколад', 'RAL 8017', 2232, 193, '⚠️ НЕТ', '—'),
    (14, 'Быстросохнущая', 'Белый', 'RAL 9010', 1507, 119, '⚠️ НЕТ', '—'),
    (15, 'Быстросохнущая', 'Зелёный', 'RAL 6005', 748, 55, '⚠️ НЕТ', '—'),
    (16, 'Быстросохнущая', 'Жёлтый', 'RAL 1023', 365, 29, '⚠️ НЕТ', '—'),
    (17, 'Молотковая', 'Чёрный', '—', 3536, 243, '✅ ДА', 'A, B, C'),
    (18, 'Молотковая', 'Коричневый', '—', 2913, 207, '✅ ДА', 'A, B, C'),
    (19, 'Молотковая', 'Серебристый', '—', 2394, 198, '✅ ДА', 'A'),
    (20, 'Молотковая', 'Медный', '—', 1162, 102, '⚠️ НЕТ', '—'),
    (21, 'Молотковая', 'Золотистый', '—', 1153, 98, '⚠️ НЕТ', '—'),
    (22, 'Молотковая', 'Тёмно-зелёный', '—', 425, 70, '⚠️ НЕТ', '—'),
]

row = 5
current_cat = ''
for data in sales_data:
    rank, cat, color, ral, qty, akb, in_mml, classes = data
    if cat != current_cat:
        current_cat = cat
        if row > 5:
            row += 1
    cat_bg = {
        '3в1 по ржавчине': light_red,
        'Быстросохнущая': light_orange,
        'Молотковая': light_yellow,
    }.get(cat, light_gray)
    mml_bg = light_green if '✅' in in_mml else yellow_fill
    mml_font = green_font if '✅' in in_mml else orange_font
    style_cell(ws2, row, 1, rank, font=bold_font, fill=cat_bg, align=center)
    style_cell(ws2, row, 2, cat, font=normal_font, fill=cat_bg, align=left)
    style_cell(ws2, row, 3, color, font=bold_font if rank <= 5 else normal_font, fill=cat_bg, align=left)
    style_cell(ws2, row, 4, ral, font=normal_font, fill=cat_bg, align=center)
    style_cell(ws2, row, 5, qty, font=bold_font, fill=cat_bg, align=right, num_format='#,##0')
    style_cell(ws2, row, 6, akb, font=normal_font, fill=cat_bg, align=center)
    style_cell(ws2, row, 7, in_mml, font=mml_font, fill=mml_bg, align=center)
    style_cell(ws2, row, 8, classes, font=normal_font, fill=cat_bg, align=center)
    row += 1

# ИТОГО
row += 1
ws2.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
style_cell(ws2, row, 1, 'ИТОГО', font=white_bold, fill=header_fill, align=center)
total_qty = sum(d[4] for d in sales_data)
total_akb = sum(d[5] for d in sales_data)
style_cell(ws2, row, 5, total_qty, font=white_bold, fill=header_fill, align=right, num_format='#,##0')
style_cell(ws2, row, 6, total_akb, font=white_bold, fill=header_fill, align=center)
style_cell(ws2, row, 7, '', fill=header_fill)
style_cell(ws2, row, 8, '', fill=header_fill)

# Вывод
row += 2
ws2.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8)
style_cell(ws2, row, 1,
    'Вывод: ТОП-5 цветов 3в1 (чёрный, графит, серый, шоколад, белый) — в MML (классы A, B). '
    'Но 8 цветов с продажами 1000+ шт — НЕ в MML. Это упущенная выручка.',
    font=orange_font, fill=yellow_fill, align=left)

col_widths2 = [6, 18, 18, 10, 12, 8, 10, 12]
for i, w in enumerate(col_widths2, 1):
    ws2.column_dimensions[get_column_letter(i)].width = w


# ==============================================================
# ЛИСТ 3: СВОДКА ПО КЛАССАМ (v2 — с колонкой "Текущие позиции MML")
# ==============================================================
ws3 = wb.create_sheet('Сводка по классам')

ws3.merge_cells('A1:H1')
ws3['A1'] = 'СВОДКА: MML TITAN ПО КЛАССАМ ТТ'
ws3['A1'].font = title_font
ws3['A1'].alignment = center
ws3.row_dimensions[1].height = 30

ws3.merge_cells('A2:H2')
ws3['A2'] = 'Было → Стало. Текущие позиции + что добавляется + обоснование продажами по каждому цвету'
ws3['A2'].font = gray_font
ws3['A2'].alignment = center

# Заголовки (8 колонок)
headers3 = [
    'Класс ТТ',           # 1
    'Канал',              # 2
    'Было (SKU)',         # 3
    'Стало (SKU)',        # 4
    'Изменение',          # 5
    'Текущие позиции MML',# 6  ← НОВОЕ
    'Что добавлено',      # 7
    'Обоснование (продажи по цветам)',  # 8 ← детализация по цветам
]
for col, h in enumerate(headers3, 1):
    style_cell(ws3, 3, col, h, font=header_font, fill=header_fill)
ws3.row_dimensions[3].height = 45

# Данные по классам
# Структура:
#   cls, channel, was, now, diff,
#   current_positions (строка с разбивкой по категориям),
#   added_positions (строка с добавлениями),
#   justification (продажи по каждому цвету: цвет → шт/АКБ + Σ)

class_data = [
    # =========================================================
    # КЛАСС A — Федеральная DIY
    # =========================================================
    (
        'A', 'Федеральная DIY', 16, 21, '+5',
        # Текущие позиции MML (16 SKU)
        '3в1 по ржавчине (10 SKU):\n'
        '• чёрный RAL 9005 (0.75л + 2л)\n'
        '• серый RAL 7004 (0.75л + 2л)\n'
        '• шоколад RAL 8017 (0.75л + 2л)\n'
        '• графитовый RAL 7024 (0.75л + 2л)\n'
        '• белый RAL 9010 (0.75л + 2л)\n'
        '\n'
        'Быстросохнущая (2 SKU):\n'
        '• чёрный RAL 9005 (0.75л + 2.2л)\n'
        '\n'
        'Молотковая (4 SKU):\n'
        '• чёрный (0.8кг + 2.5кг)\n'
        '• коричневый (0.8кг)\n'
        '• серебристый (0.8кг)',
        # Что добавлено (5 SKU)
        '3в1 по ржавчине (1):\n'
        '• зелёный RAL 6005 (0.75л)\n'
        '\n'
        'Быстросохнущая (3):\n'
        '• серый RAL 7004 (0.75л)\n'
        '• шоколад RAL 8017 (0.75л)\n'
        '• белый RAL 9010 (0.75л)\n'
        '\n'
        'Молотковая (1):\n'
        '• медный (0.8кг)',
        # Обоснование — продажи по каждому цвету
        '• 3в1 зелёный RAL 6005 — 1 893 шт / 111 АКБ\n'
        '• Быстросохн. серый RAL 7004 — 2 695 шт / 199 АКБ\n'
        '• Быстросохн. шоколад RAL 8017 — 2 232 шт / 193 АКБ\n'
        '• Быстросохн. белый RAL 9010 — 1 507 шт / 119 АКБ\n'
        '• Молотковая медный — 1 162 шт / 102 АКБ\n'
        '\n'
        'Σ = 9 489 шт (ΣАКБ = 724)',
    ),
    # =========================================================
    # КЛАСС B — Локальная сеть
    # =========================================================
    (
        'B', 'Локальная сеть', 12, 16, '+4',
        # Текущие позиции MML (12 SKU)
        '3в1 по ржавчине (9 SKU):\n'
        '• чёрный RAL 9005 (0.75л + 2л)\n'
        '• серый RAL 7004 (0.75л + 2л)\n'
        '• шоколад RAL 8017 (0.75л + 2л)\n'
        '• графитовый RAL 7024 (0.75л + 2л)\n'
        '• белый RAL 9010 (0.75л)\n'
        '\n'
        'Быстросохнущая (1 SKU):\n'
        '• чёрный RAL 9005 (0.75л)\n'
        '\n'
        'Молотковая (2 SKU):\n'
        '• чёрный (0.8кг)\n'
        '• коричневый (0.8кг)',
        # Что добавлено (4 SKU)
        '3в1 по ржавчине (1):\n'
        '• зелёный RAL 6005 (0.75л)\n'
        '\n'
        'Быстросохнущая (2):\n'
        '• серый RAL 7004 (0.75л)\n'
        '• шоколад RAL 8017 (0.75л)\n'
        '\n'
        'Молотковая (1):\n'
        '• серебристый (0.8кг)',
        # Обоснование — продажи по каждому цвету
        '• 3в1 зелёный RAL 6005 — 1 893 шт / 111 АКБ\n'
        '• Быстросохн. серый RAL 7004 — 2 695 шт / 199 АКБ\n'
        '• Быстросохн. шоколад RAL 8017 — 2 232 шт / 193 АКБ\n'
        '• Молотковая серебристый — 2 394 шт / 198 АКБ\n'
        '\n'
        'Σ = 9 214 шт (ΣАКБ = 701)',
    ),
    # =========================================================
    # КЛАСС C — Супермаркет
    # =========================================================
    (
        'C', 'Супермаркет', 2, 7, '+5',
        # Текущие позиции MML (2 SKU)
        'Молотковая (2 SKU):\n'
        '• чёрный (0.8кг)\n'
        '• коричневый (0.8кг)\n'
        '\n'
        '3в1 по ржавчине — НЕТ\n'
        'Быстросохнущая — НЕТ',
        # Что добавлено (5 SKU)
        '3в1 по ржавчине (3):\n'
        '• чёрный RAL 9005 (0.75л)\n'
        '• шоколад RAL 8017 (0.75л)\n'
        '• серый RAL 7004 (0.75л)\n'
        '\n'
        'Быстросохнущая (2):\n'
        '• чёрный RAL 9005 (0.75л)\n'
        '• серый RAL 7004 (0.75л)',
        # Обоснование — продажи по каждому цвету
        '• 3в1 чёрный RAL 9005 — 7 738 шт / 393 АКБ\n'
        '• 3в1 шоколад RAL 8017 — 4 524 шт / 287 АКБ\n'
        '• 3в1 серый RAL 7004 — 5 071 шт / 247 АКБ\n'
        '• Быстросохн. чёрный RAL 9005 — 5 381 шт / 303 АКБ\n'
        '• Быстросохн. серый RAL 7004 — 2 695 шт / 199 АКБ\n'
        '\n'
        'Σ = 25 409 шт (ΣАКБ = 1 429)',
    ),
    # =========================================================
    # КЛАСС D — Хозмаг (БЕЗ ДОБАВЛЕНИЙ)
    # =========================================================
    (
        'D', 'Хозмаг', 0, 0, '0',
        # Текущие позиции MML (0 SKU)
        'TITAN в MML — НЕТ\n'
        '(в MML класса D только Эмаль ПФ-115 "Простокрашено!" и грунт-эмаль "Расцвет")',
        # Что добавлено — НЕТ
        'Добавлений нет\n'
        '(категория D исключена из расширения MML TITAN)',
        # Обоснование — НЕТ
        '—',
    ),
]

cls_colors = {'A': 'C62828', 'B': 'F47920', 'C': '16A34A', 'D': '1E3A5F'}
cls_light = {
    'A': PatternFill('solid', fgColor='FEE2E2'),
    'B': PatternFill('solid', fgColor='FFEDD5'),
    'C': PatternFill('solid', fgColor='DCFCE7'),
    'D': PatternFill('solid', fgColor='DBEAFE'),
}

for ri, (cls, channel, was, now, diff, current_pos, added, reason) in enumerate(class_data):
    r = 4 + ri
    cls_color = cls_colors[cls]
    light_fill = cls_light[cls]
    
    # Колонка 1: Класс ТТ
    style_cell(ws3, r, 1, cls,
               font=Font(name='Arial', size=14, bold=True, color='FFFFFF'),
               fill=PatternFill('solid', fgColor=cls_color), align=center)
    # Колонка 2: Канал
    style_cell(ws3, r, 2, channel, font=bold_font, fill=light_fill, align=left)
    # Колонка 3: Было
    style_cell(ws3, r, 3, was, font=red_font, fill=light_red, align=center)
    # Колонка 4: Стало
    style_cell(ws3, r, 4, now, font=green_font, fill=light_green, align=center)
    # Колонка 5: Изменение
    style_cell(ws3, r, 5, diff, font=green_font, fill=light_green, align=center)
    # Колонка 6: Текущие позиции MML (новое)
    style_cell(ws3, r, 6, current_pos, font=small_font, fill=light_fill, align=left_top)
    # Колонка 7: Что добавлено
    style_cell(ws3, r, 7, added, font=small_font, fill=PatternFill('solid', fgColor='FEF9C3'), align=left_top)
    # Колонка 8: Обоснование
    style_cell(ws3, r, 8, reason, font=small_font, fill=light_gray, align=left_top)
    
    # Высота строки — увеличенная для содержимого
    ws3.row_dimensions[r].height = 200 if cls in ['A', 'B'] else (80 if cls == 'D' else 150)

# ИТОГО
r = 8
ws3.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
style_cell(ws3, r, 1, 'ИТОГО', font=white_bold, fill=header_fill, align=center)
style_cell(ws3, r, 3, 30, font=white_bold, fill=header_fill, align=center)
style_cell(ws3, r, 4, 44, font=white_bold, fill=PatternFill('solid', fgColor='16A34A'), align=center)
style_cell(ws3, r, 5, '+14', font=white_bold, fill=PatternFill('solid', fgColor='16A34A'), align=center)
ws3.merge_cells(start_row=r, start_column=6, end_row=r, end_column=8)
style_cell(ws3, r, 6, 'Все добавления основаны на реальных вторичных продажах (источник: «Титан цвета»)',
           font=white_bold, fill=header_fill, align=center)
ws3.row_dimensions[r].height = 30

# Примечание под таблицей
r = 10
ws3.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
style_cell(ws3, r, 1,
    'Легенда: Σ — сумма по добавленным SKU. АКБ — активная клиентская база (сколько клиентов купило цвет). '
    'ΣАКБ может быть меньше суммы АКБ по отдельным цветам, т.к. один клиент мог купить несколько цветов.',
    font=gray_font, fill=light_gray, align=left)
ws3.row_dimensions[r].height = 30

# Ширина колонок
col_widths3 = [6, 16, 8, 8, 9, 38, 28, 38]
for i, w in enumerate(col_widths3, 1):
    ws3.column_dimensions[get_column_letter(i)].width = w


# ==============================================================
# Сохранение
# ==============================================================
wb.save(OUT)
print(f'OK: {OUT}')
print(f'Листов: {len(wb.sheetnames)}')
for name in wb.sheetnames:
    print(f'  "{name}"')
