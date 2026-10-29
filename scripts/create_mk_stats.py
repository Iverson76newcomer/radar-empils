#!/usr/bin/env python3
"""
Таблица для сбора статистики на мастер-классе TITAN в Лемана ПРО.
2 листа: 1) Чек-лист подготовки  2) Статистика по часам
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, numbers
from openpyxl.utils import get_column_letter

OUT = '/home/z/my-project/download/Статистика_мастер-класс_TITAN.xlsx'

# Стили
header_font = Font(name='Arial', size=10, bold=True, color='FFFFFF')
header_fill = PatternFill('solid', fgColor='C62828')
title_font = Font(name='Arial', size=14, bold=True, color='1E3A5F')
subtitle_font = Font(name='Arial', size=11, bold=True, color='1E3A5F')
normal_font = Font(name='Arial', size=10, color='1A1F36')
bold_font = Font(name='Arial', size=10, bold=True, color='1A1F36')
red_font = Font(name='Arial', size=10, bold=True, color='C62828')
green_font = Font(name='Arial', size=10, bold=True, color='16A34A')
gray_font = Font(name='Arial', size=9, color='6B7280')

light_gray = PatternFill('solid', fgColor='F3F4F6')
light_red = PatternFill('solid', fgColor='FEE2E2')
light_green = PatternFill('solid', fgColor='DCFCE7')
light_orange = PatternFill('solid', fgColor='FFEDD5')
light_blue = PatternFill('solid', fgColor='DBEAFE')
yellow_fill = PatternFill('solid', fgColor='FEF3C7')

thin = Side(style='thin', color='D1D5DB')
border = Border(left=thin, right=thin, top=thin, bottom=thin)

center = Alignment(horizontal='center', vertical='center', wrap_text=True)
left = Alignment(horizontal='left', vertical='center', wrap_text=True)


wb = openpyxl.Workbook()

# ==============================================================
# ЛИСТ 1: ЧЕК-ЛИСТ ПОДГОТОВКИ
# ==============================================================
ws1 = wb.active
ws1.title = 'Чек-лист подготовки'

ws1.merge_cells('A1:D1')
ws1['A1'] = 'ЧЕК-ЛИСТ ПОДГОТОВКИ К МАСТЕР-КЛАССУ TITAN'
ws1['A1'].font = title_font
ws1['A1'].alignment = center
ws1.row_dimensions[1].height = 30

ws1.merge_cells('A2:D2')
ws1['A2'] = 'Заполнить до начала мастер-класса'
ws1['A2'].font = gray_font
ws1['A2'].alignment = center

# Параметры МК
params = [
    ('Дата', '', ''),
    ('Магазин (адрес)', '', ''),
    ('Время начала', '', ''),
    ('Время окончания', '', ''),
    ('Сотрудник 1 (ФИО)', '', ''),
    ('Сотрудник 2 (ФИО)', '', ''),
    ('Количество банок на столе', '', ''),
    ('Количество пластинок (неокраш.)', '', ''),
    ('Количество подарков (кисть+перчатки)', '', ''),
    ('Количество сертификатов (500₽)', '', ''),
    ('Ролл-ап установлен?', '', 'Да / Нет'),
    ('Стол для выкрашивания готов?', '', 'Да / Нет'),
    ('Скатерть, салфетки, мешки для мусора?', '', 'Да / Нет'),
    ('Фартуки, перчатки для покупателей?', '', 'Да / Нет'),
    ('Готовые пластины RAL 7024 и 8017?', '', 'Да / Нет'),
]

ws1['A4'] = 'Параметр'
ws1['B4'] = 'Значение'
ws1['C4'] = 'Примечание'
for col in ['A', 'B', 'C']:
    ws1[f'{col}4'].font = header_font
    ws1[f'{col}4'].fill = header_fill
    ws1[f'{col}4'].alignment = center
    ws1[f'{col}4'].border = border

for i, (param, val, note) in enumerate(params):
    r = 5 + i
    ws1.cell(r, 1, param).font = bold_font
    ws1.cell(r, 1).fill = light_gray
    ws1.cell(r, 1).border = border
    ws1.cell(r, 1).alignment = left
    ws1.cell(r, 2, val).font = normal_font
    ws1.cell(r, 2).border = border
    ws1.cell(r, 2).alignment = left
    ws1.cell(r, 3, note).font = gray_font
    ws1.cell(r, 3).border = border
    ws1.cell(r, 3).alignment = left

ws1.column_dimensions['A'].width = 35
ws1.column_dimensions['B'].width = 30
ws1.column_dimensions['C'].width = 20


# ==============================================================
# ЛИСТ 2: СТАТИСТИКА ПО ЧАСАМ
# ==============================================================
ws2 = wb.create_sheet('Статистика по часам')

ws2.merge_cells('A1:K1')
ws2['A1'] = 'СТАТИСТИКА МАСТЕР-КЛАССА TITAN — ПО ЧАСАМ'
ws2['A1'].font = title_font
ws2['A1'].alignment = center
ws2.row_dimensions[1].height = 30

ws2.merge_cells('A2:K2')
ws2['A2'] = 'Отмечайте каждый час. Цифры — количество человек за этот час.'
ws2['A2'].font = gray_font
ws2['A2'].alignment = center

# Заголовки
headers = [
    'Час',
    'Подошли к столу',
    'Получили консультацию',
    'Участвовали в выкрасе',
    'Заинтересовались (взяли банку в руки)',
    'Совершили покупку TITAN',
    'Сумма чеков (₽)',
    'Подарок: кисть+перчатки (шт)',
    'Сертификат 500₽ (шт)',
    'Возражений по цене',
    'Примечания / вопросы покупателей',
]

for col, h in enumerate(headers, 1):
    cell = ws2.cell(4, col, h)
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = center
    cell.border = border
ws2.row_dimensions[4].height = 45

# Часовые интервалы (14:30 - 19:30)
hours = [
    '14:30–15:00',
    '15:00–15:30',
    '15:30–16:00',
    '16:00–16:30',
    '16:30–17:00',
    '17:00–17:30',
    '17:30–18:00',
    '18:00–18:30',
    '18:30–19:00',
    '19:00–19:30',
]

for ri, hour in enumerate(hours):
    r = 5 + ri
    ws2.cell(r, 1, hour).font = bold_font
    ws2.cell(r, 1).fill = light_blue
    ws2.cell(r, 1).border = border
    ws2.cell(r, 1).alignment = center
    for col in range(2, 12):
        cell = ws2.cell(r, col, '')
        cell.font = normal_font
        cell.border = border
        cell.alignment = center
        if ri % 2 == 0:
            cell.fill = light_gray

# Строка ИТОГО
total_row = 5 + len(hours)
ws2.cell(total_row, 1, 'ИТОГО').font = Font(name='Arial', size=11, bold=True, color='FFFFFF')
ws2.cell(total_row, 1).fill = PatternFill('solid', fgColor='1E3A5F')
ws2.cell(total_row, 1).border = border
ws2.cell(total_row, 1).alignment = center

for col in range(2, 11):
    cell = ws2.cell(total_row, col)
    col_letter = get_column_letter(col)
    cell.value = f'=SUM({col_letter}5:{col_letter}{total_row-1})'
    cell.font = Font(name='Arial', size=10, bold=True, color='FFFFFF')
    cell.fill = PatternFill('solid', fgColor='1E3A5F')
    cell.border = border
    cell.alignment = center

# Примечания — без формулы
ws2.cell(total_row, 11, '').fill = PatternFill('solid', fgColor='1E3A5F')
ws2.cell(total_row, 11).border = border

# Ширина колонок
col_widths = [14, 14, 16, 14, 16, 14, 14, 14, 14, 14, 35]
for i, w in enumerate(col_widths, 1):
    ws2.column_dimensions[get_column_letter(i)].width = w

# ==============================================================
# ЛИСТ 3: ОПРОС ПОКУПАТЕЛЕЙ
# ==============================================================
ws3 = wb.create_sheet('Опрос покупателей')

ws3.merge_cells('A1:F1')
ws3['A1'] = 'ОПРОС ПОКУПАТЕЛЕЙ — КЛЮЧЕВЫЕ КРИТЕРИИ И БАРЬЕРЫ'
ws3['A1'].font = title_font
ws3['A1'].alignment = center
ws3.row_dimensions[1].height = 30

ws3.merge_cells('A2:F2')
ws3['A2'] = 'Задавайте эти вопросы во время МК. Записывайте ответы кратко.'
ws3['A2'].font = gray_font
ws3['A2'].alignment = center

headers3 = ['№', 'Время', 'Что красит?', 'Ключевой критерий выбора ЛКМ', 'Главный барьер к покупке TITAN', 'Купил?']

for col, h in enumerate(headers3, 1):
    cell = ws3.cell(4, col, h)
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = center
    cell.border = border
ws3.row_dimensions[4].height = 35

# 15 пустых строк
for ri in range(15):
    r = 5 + ri
    ws3.cell(r, 1, ri + 1).font = bold_font
    ws3.cell(r, 1).fill = light_blue
    ws3.cell(r, 1).border = border
    ws3.cell(r, 1).alignment = center
    for col in range(2, 7):
        cell = ws3.cell(r, col, '')
        cell.font = normal_font
        cell.border = border
        cell.alignment = left
        if ri % 2 == 0:
            cell.fill = light_gray

col_widths3 = [5, 10, 20, 30, 30, 12]
for i, w in enumerate(col_widths3, 1):
    ws3.column_dimensions[get_column_letter(i)].width = w

# Подсказки
ws3.merge_cells('A21:F21')
ws3['A21'] = 'Подсказки: критерии выбора — цена / долговечность / цвет / скорость сушки / бренд. Барьеры — дорого / не знаю бренд / беру ПФ-115 / нет нужного цвета.'
ws3['A21'].font = gray_font
ws3['A21'].alignment = left
ws3.row_dimensions[21].height = 30


# ==============================================================
# ЛИСТ 4: СВОДКА (заполняется в конце)
# ==============================================================
ws4 = wb.create_sheet('Сводка')

ws4.merge_cells('A1:C1')
ws4['A1'] = 'СВОДКА МАСТЕР-КЛАССА — ИТОГИ'
ws4['A1'].font = title_font
ws4['A1'].alignment = center
ws4.row_dimensions[1].height = 30

ws4.merge_cells('A2:C2')
ws4['A2'] = 'Заполнить в конце дня'
ws4['A2'].font = gray_font
ws4['A2'].alignment = center

# Итоговые показатели (ссылки на лист Статистика)
metrics = [
    ('Всего подошло к столу', "='Статистика по часам'!B15", 'чел'),
    ('Получили консультацию', "='Статистика по часам'!C15", 'чел'),
    ('Участвовали в выкрасе', "='Статистика по часам'!D15", 'чел'),
    ('Заинтересовались', "='Статистика по часам'!E15", 'чел'),
    ('Совершили покупку', "='Статистика по часам'!F15", 'чел'),
    ('Сумма чеков', "='Статистика по часам'!G15", '₽'),
    ('Подарено: кисть+перчатки', "='Статистика по часам'!H15", 'шт'),
    ('Подарено: сертификаты', "='Статистика по часам'!I15", 'шт'),
    ('Возражений по цене', "='Статистика по часам'!J15", 'шт'),
]

headers4 = ['Показатель', 'Значение', 'Ед. изм.']
for col, h in enumerate(headers4, 1):
    cell = ws4.cell(4, col, h)
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = center
    cell.border = border

for i, (name, formula, unit) in enumerate(metrics):
    r = 5 + i
    ws4.cell(r, 1, name).font = bold_font
    ws4.cell(r, 1).fill = light_gray
    ws4.cell(r, 1).border = border
    ws4.cell(r, 1).alignment = left
    ws4.cell(r, 2, formula).font = Font(name='Arial', size=12, bold=True, color='C62828')
    ws4.cell(r, 2).border = border
    ws4.cell(r, 2).alignment = center
    ws4.cell(r, 3, unit).font = gray_font
    ws4.cell(r, 3).border = border
    ws4.cell(r, 3).alignment = center

# Конверсия
conv_row = 5 + len(metrics) + 1
ws4.cell(conv_row, 1, 'КОНВЕРСИЯ').font = subtitle_font
ws4.cell(conv_row, 1).fill = yellow_fill
ws4.cell(conv_row, 1).border = border

conversions = [
    ('Конверсия в покупку (от подошедших)', "=IFERROR(B6/B5,0)", '%'),
    ('Конверсия в выкрас (от подошедших)', "=IFERROR(B7/B5,0)", '%'),
    ('Средний чек', "=IFERROR(B10/B9,0)", '₽'),
]

for i, (name, formula, unit) in enumerate(conversions):
    r = conv_row + 1 + i
    ws4.cell(r, 1, name).font = bold_font
    ws4.cell(r, 1).fill = light_gray
    ws4.cell(r, 1).border = border
    ws4.cell(r, 1).alignment = left
    ws4.cell(r, 2, formula).font = green_font
    ws4.cell(r, 2).border = border
    ws4.cell(r, 2).alignment = center
    ws4.cell(r, 2).number_format = '0.0%' if unit == '%' else '#,##0'
    ws4.cell(r, 3, unit).font = gray_font
    ws4.cell(r, 3).border = border
    ws4.cell(r, 3).alignment = center

# Примечание
note_row = conv_row + 1 + len(conversions) + 1
ws4.merge_cells(f'A{note_row}:C{note_row}')
ws4.cell(note_row, 1, 'Передать результаты: Зингеровской Е.Г. — опрос покупателей (критерии и барьеры)')
ws4.cell(note_row, 1).font = red_font
ws4.cell(note_row, 1).alignment = left

ws4.column_dimensions['A'].width = 40
ws4.column_dimensions['B'].width = 15
ws4.column_dimensions['C'].width = 12


# Сохранение
wb.save(OUT)
print(f'OK: {OUT}')
print(f'Листов: {len(wb.sheetnames)}')
for name in wb.sheetnames:
    print(f'  "{name}"')
