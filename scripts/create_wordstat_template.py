#!/usr/bin/env python3
"""
Excel-шаблон + Word-инструкция для сбора данных Wordstat.
14 запросов × 9 регионов (Россия + 8 ФО) = 126 ячеек для заполнения.

3 уровня запросов:
- Производитель (6 шт) — Empils, Dufa, Tikkurila, Лакра, Рогнеда, Elcon
- ТМ продукта (8 шт) — TITAN, CORREX, HAMMERLACK, METALLISTA, PROF IT, Dali, KLONDIKE + TITAN грунт-эмаль
- Категория (2 шт) — грунт-эмаль 3 в 1, краска по ржавчине
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

OUT = '/home/z/my-project/download/Wordstat_шаблон_сбора.xlsx'

# Стили
HEADER_FONT = Font(name='Arial', size=11, bold=True, color='FFFFFF')
HEADER_FILL = PatternFill('solid', fgColor='1E3A5F')
TITLE_FONT = Font(name='Arial', size=14, bold=True, color='1E3A5F')
SUBTITLE_FONT = Font(name='Arial', size=11, italic=True, color='6B7280')
NORMAL_FONT = Font(name='Arial', size=10, color='1A1F36')
BOLD_FONT = Font(name='Arial', size=10, bold=True, color='1A1F36')
GRAY_FONT = Font(name='Arial', size=9, color='6B7280')
ACCENT_FONT = Font(name='Arial', size=10, bold=True, color='F47920')
WHITE_BOLD = Font(name='Arial', size=11, bold=True, color='FFFFFF')

LIGHT_GRAY = PatternFill('solid', fgColor='F3F4F6')
YELLOW = PatternFill('solid', fgColor='FEF3C7')  # для заполнения
LIGHT_BLUE = PatternFill('solid', fgColor='DBEAFE')  # для производителя
LIGHT_GREEN = PatternFill('solid', fgColor='DCFCE7')  # для ТМ
LIGHT_ORANGE = PatternFill('solid', fgColor='FFEDD5')  # для категории

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


# 14 запросов — 3 уровня
queries = [
    # === УРОВЕНЬ 1: ПРОИЗВОДИТЕЛЬ (6 шт) — голубой ===
    {'n': 1, 'level': 'Производитель', 'query': 'Empils краска', 'tm_note': 'TITAN, ОРЕОЛ — производитель Empils'},
    {'n': 2, 'level': 'Производитель', 'query': 'Dufa краска', 'tm_note': 'CORREX, HAMMERLACK — производитель Dufa'},
    {'n': 3, 'level': 'Производитель', 'query': 'Tikkurila краска', 'tm_note': 'METALLISTA — производитель Tikkurila'},
    {'n': 4, 'level': 'Производитель', 'query': 'Лакра краска', 'tm_note': 'PROF IT — производитель Лакра'},
    {'n': 5, 'level': 'Производитель', 'query': 'Рогнеда краска', 'tm_note': 'Dali — производитель Рогнеда'},
    {'n': 6, 'level': 'Производитель', 'query': 'Elcon краска', 'tm_note': 'без отдельной ТМ'},
    # === УРОВЕНЬ 2: ТМ ПРОДУКТА (8 шт) — зелёный ===
    {'n': 7, 'level': 'ТМ продукта', 'query': 'TITAN краска', 'tm_note': 'основной бренд Empils'},
    {'n': 8, 'level': 'ТМ продукта', 'query': 'TITAN грунт-эмаль', 'tm_note': 'категорийный запрос TITAN'},
    {'n': 9, 'level': 'ТМ продукта', 'query': 'CORREX краска', 'tm_note': 'ТМ Dufa — 3в1'},
    {'n': 10, 'level': 'ТМ продукта', 'query': 'HAMMERLACK', 'tm_note': 'ТМ Dufa — молотковая'},
    {'n': 11, 'level': 'ТМ продукта', 'query': 'METALLISTA Tikkurila', 'tm_note': 'ТМ Tikkurila — по металлу'},
    {'n': 12, 'level': 'ТМ продукта', 'query': 'Dali краска по ржавчине', 'tm_note': 'ТМ Рогнеда — 3в1'},
    {'n': 13, 'level': 'ТМ продукта', 'query': 'HUSKY KLONDIKE', 'tm_note': 'ТМ HUSKY — молотковая/3в1'},
    {'n': 14, 'level': 'ТМ продукта', 'query': 'Лакра PROF IT', 'tm_note': 'ТМ Лакра — профессиональная'},
    # === УРОВЕНЬ 3: КАТЕГОРИЯ (2 шт) — оранжевый ===
    {'n': 15, 'level': 'Категория', 'query': 'грунт-эмаль 3 в 1', 'tm_note': 'объём рынка в запросах'},
    {'n': 16, 'level': 'Категория', 'query': 'краска по ржавчине', 'tm_note': 'объём рынка в запросах'},
]

# 9 регионов: Россия + 8 ФО
regions = ['Россия (всего)', 'Центральный ФО', 'Северо-Западный ФО', 'Южный ФО',
           'Северо-Кавказский ФО', 'Приволжский ФО', 'Уральский ФО',
           'Сибирский ФО', 'Дальневосточный ФО']


# ========================================================
# СОЗДАНИЕ EXCEL
# ========================================================
wb = openpyxl.Workbook()

# === ЛИСТ 1: ШАБЛОН ДЛЯ ЗАПОЛНЕНИЯ ===
ws = wb.active
ws.title = 'Шаблон сбора'

# Заголовок
ws.merge_cells('A1:M1')
ws['A1'] = 'Wordstat — сбор данных по узнаваемости брендов TITAN и конкурентов'
ws['A1'].font = TITLE_FONT
ws['A1'].alignment = CENTER
ws.row_dimensions[1].height = 25

ws.merge_cells('A2:M2')
ws['A2'] = '16 запросов × 9 регионов = 144 ячейки для заполнения. Источник: https://wordstat.yandex.ru/'
ws['A2'].font = SUBTITLE_FONT
ws['A2'].alignment = CENTER
ws.row_dimensions[2].height = 18

# Шапка таблицы
headers = ['№', 'Уровень', 'Что вводить в Wordstat', 'Зачем / привязка к ТМ'] + regions
for ci, h in enumerate(headers, 1):
    style_cell(ws, 4, ci, h, font=HEADER_FONT, fill=HEADER_FILL)
ws.row_dimensions[4].height = 45

# Данные — 16 строк
for ri, q in enumerate(queries, 1):
    r = 4 + ri
    # Цвет по уровню
    if q['level'] == 'Производитель':
        level_fill = LIGHT_BLUE
    elif q['level'] == 'ТМ продукта':
        level_fill = LIGHT_GREEN
    else:
        level_fill = LIGHT_ORANGE
    
    style_cell(ws, r, 1, q['n'], font=BOLD_FONT, fill=level_fill, align=CENTER)
    style_cell(ws, r, 2, q['level'], font=ACCENT_FONT, fill=level_fill, align=LEFT)
    style_cell(ws, r, 3, q['query'], font=BOLD_FONT, fill=level_fill, align=LEFT)
    style_cell(ws, r, 4, q['tm_note'], font=GRAY_FONT, fill=level_fill, align=LEFT)
    
    # 9 колонок регионов — пустые, для заполнения (жёлтые)
    for ci in range(5, 14):
        style_cell(ws, r, ci, '', font=BOLD_FONT, fill=YELLOW, align=RIGHT, num_format='#,##0')

# ИТОГО — формулы по строкам
r_total = 4 + len(queries) + 1
ws.merge_cells(start_row=r_total, start_column=1, end_row=r_total, end_column=4)
style_cell(ws, r_total, 1, 'СУММА ПО ВСЕМ ЗАПРОСАМ', font=WHITE_BOLD, fill=HEADER_FILL, align=LEFT)
for ci in range(5, 14):
    col_letter = get_column_letter(ci)
    formula = f'=SUM({col_letter}5:{col_letter}{4+len(queries)})'
    cell = ws.cell(r_total, ci, formula)
    cell.font = WHITE_BOLD
    cell.fill = HEADER_FILL
    cell.alignment = RIGHT
    cell.border = BORDER
    cell.number_format = '#,##0'

# Ширина колонок
col_widths = [4, 16, 26, 30, 12, 14, 14, 14, 16, 14, 12, 12, 14]
for i, w in enumerate(col_widths, 1):
    ws.column_dimensions[get_column_letter(i)].width = w


# === ЛИСТ 2: ИНСТРУКЦИЯ ===
ws2 = wb.create_sheet('Инструкция')

ws2.merge_cells('A1:F1')
ws2['A1'] = 'Как заполнить шаблон'
ws2['A1'].font = TITLE_FONT
ws2['A1'].alignment = LEFT
ws2.row_dimensions[1].height = 25

instructions = [
    ('Шаг 1.', 'Откройте https://wordstat.yandex.ru/ и авторизуйтесь через Яндекс ID'),
    ('Шаг 2.', 'В поле ввода скопируйте запрос из колонки «Что вводить в Wordstat» (без кавычек)'),
    ('Шаг 3.', 'Нажмите «Подобрать». Откроется страница с данными.'),
    ('Шаг 4.', 'Цифра «Россия (всего)» — большое число над графиком «История запросов» вверху страницы'),
    ('Шаг 5.', 'Цифры по 8 ФО — прокрутите вниз до блока «По регионам». Найдите строки для каждого федерального округа.'),
    ('Шаг 6.', 'Внесите все 9 цифр в соответствующую строку шаблона (жёлтые ячейки)'),
    ('Шаг 7.', 'Повторите для всех 16 запросов. Итого: 16 × 9 = 144 цифры'),
    ('', ''),
    ('⚠️ Важно:', 'Запросы вводите БЕЗ кавычек, если в шаблоне они не указаны'),
    ('⚠️ Важно:', 'Если цифра по региону < 5 — Wordstat показывает «менее 5». Вносите 0'),
    ('⚠️ Важно:', 'Если запроса нет вообще (0 показов) — оставьте ячейку пустой'),
    ('⚠️ Важно:', 'Цифры могут отличаться на ±5% из-за сезонности — это нормально'),
    ('', ''),
    ('Легенда цветов:', ''),
    ('🔵 Голубой', 'Производитель — запросы по названию компании (Empils, Dufa, Tikkurila и т.д.)'),
    ('🟢 Зелёный', 'ТМ продукта — запросы по конкретному продукту (TITAN, CORREX, METALLISTA и т.д.)'),
    ('🟠 Оранжевый', 'Категория — запросы без бренда, по типу продукта (грунт-эмаль 3 в 1)'),
    ('🟡 Жёлтый', 'Ячейки для заполнения — впишите туда цифры из Wordstat'),
    ('', ''),
    ('Сколько времени:', '~30-40 минут на все 16 запросов (по 2 минуты на запрос)'),
    ('Источник:', 'https://wordstat.yandex.ru/ (требуется авторизация Яндекс ID)'),
]

for ri, (label, text) in enumerate(instructions, 3):
    # Выбор цвета заливки
    if 'Голубой' in label:
        color_hex = 'DBEAFE'
    elif 'Зелёный' in label:
        color_hex = 'DCFCE7'
    elif 'Оранжевый' in label:
        color_hex = 'FFEDD5'
    elif label.startswith('🟡') or 'Жёлтый' in label:
        color_hex = 'FEF3C7'
    else:
        color_hex = 'F3F4F6'
    
    # Простая запись — без style_cell, чтобы избежать ошибок
    cell_a = ws2.cell(ri, 1, label)
    cell_a.font = Font(name='Arial', size=10, bold=True, color='1A1F36')
    cell_a.fill = PatternFill(start_color=color_hex, end_color=color_hex, fill_type='solid')
    cell_a.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
    cell_a.border = BORDER
    
    if text:
        cell_b = ws2.cell(ri, 2, text)
        cell_b.font = Font(name='Arial', size=10, color='1A1F36')
        cell_b.fill = PatternFill(start_color=color_hex, end_color=color_hex, fill_type='solid')
        cell_b.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
        cell_b.border = BORDER
        ws2.merge_cells(start_row=ri, start_column=2, end_row=ri, end_column=6)
        # Заполним merged ячейки тоже
        for c in range(3, 7):
            mc = ws2.cell(ri, c)
            mc.fill = PatternFill(start_color=color_hex, end_color=color_hex, fill_type='solid')
            mc.border = BORDER
    elif label:
        # Пустая строка-заголовок (объединяем A:F)
        ws2.merge_cells(start_row=ri, start_column=1, end_row=ri, end_column=6)
        for c in range(2, 7):
            mc = ws2.cell(ri, c)
            mc.fill = PatternFill(start_color=color_hex, end_color=color_hex, fill_type='solid')
            mc.border = BORDER

col_widths2 = [16, 30, 20, 20, 20, 20]
for i, w in enumerate(col_widths2, 1):
    ws2.column_dimensions[get_column_letter(i)].width = w


# === ЛИСТ 3: РЕГИОНЫ РФ — ПОДСКАЗКА ===
ws3 = wb.create_sheet('Регионы ФО')

ws3.merge_cells('A1:C1')
ws3['A1'] = 'Регионы, входящие в каждый федеральный округ'
ws3['A1'].font = TITLE_FONT
ws3['A1'].alignment = LEFT
ws3.row_dimensions[1].height = 25

ws3.merge_cells('A2:C2')
ws3['A2'] = 'Подсказка: в Wordstat цифры по ФО — это сумма всех регионов внутри округа. Используйте для проверки.'
ws3['A2'].font = SUBTITLE_FONT
ws3['A2'].alignment = LEFT

# Заголовки
for ci, h in enumerate(['Федеральный округ', 'Главные регионы', 'Код в Wordstat'], 1):
    style_cell(ws3, 4, ci, h, font=HEADER_FONT, fill=HEADER_FILL)
ws3.row_dimensions[4].height = 25

fo_data = [
    ('Центральный ФО', 'Москва, Московская обл., Воронежская, Тульская, Ярославская', 'ЦФО'),
    ('Северо-Западный ФО', 'Санкт-Петербург, Ленинградская обл., Калининградская, Архангельская', 'СЗФО'),
    ('Южный ФО', 'Краснодарский край, Ростовская обл., Волгоградская, Астраханская', 'ЮФО'),
    ('Северо-Кавказский ФО', 'Ставропольский край, Дагестан, Чечня, Кабардино-Балкария', 'СКФО'),
    ('Приволжский ФО', 'Нижегородская, Самарская, Татарстан, Башкортостан, Пермский край', 'ПФО'),
    ('Уральский ФО', 'Свердловская, Челябинская, Тюменская обл., ХМАО, ЯНАО', 'УрФО'),
    ('Сибирский ФО', 'Новосибирская, Красноярский край, Иркутская, Кемеровская, Омская, Алтайский край', 'СибФО'),
    ('Дальневосточный ФО', 'Приморский край, Хабаровский край, Якутия, Амурская обл.', 'ДВФО'),
]

for ri, (fo, regions_list, code) in enumerate(fo_data, 5):
    fill_hex = 'FFFFFF' if ri % 2 == 1 else 'F9FAFB'
    fill_obj = PatternFill('solid', fgColor=fill_hex)
    style_cell(ws3, ri, 1, fo, font=BOLD_FONT, fill=fill_obj, align=LEFT)
    fill_obj2 = PatternFill('solid', fgColor=fill_hex)
    style_cell(ws3, ri, 2, regions_list, font=NORMAL_FONT, fill=fill_obj2, align=LEFT)
    fill_obj3 = PatternFill('solid', fgColor=fill_hex)
    style_cell(ws3, ri, 3, code, font=ACCENT_FONT, fill=fill_obj3, align=CENTER)

col_widths3 = [25, 60, 15]
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
print(f'\nЗапросов: {len(queries)}')
print(f'Регионов: {len(regions)} (Россия + 8 ФО)')
print(f'Ячеек для заполнения: {len(queries) * len(regions)}')
