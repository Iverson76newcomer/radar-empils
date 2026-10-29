#!/usr/bin/env python3
"""
Excel-таблица с обоснованием рекомендаций по корректировке MML.
Основание: данные первичных и вторичных продаж + текущий MML.
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

OUTPUT = '/home/z/my-project/download/Рекомендации_MML_с_обоснованием.xlsx'

wb = openpyxl.Workbook()

# === СТИЛИ ===
thin_border = Border(
    left=Side(style='thin', color='D1D5DB'),
    right=Side(style='thin', color='D1D5DB'),
    top=Side(style='thin', color='D1D5DB'),
    bottom=Side(style='thin', color='D1D5DB'),
)

header_font = Font(name='Arial', size=10, bold=True, color='FFFFFF')
header_fill = PatternFill('solid', fgColor='1E3A5F')
title_font = Font(name='Arial', size=14, bold=True, color='1E3A5F')
subtitle_font = Font(name='Arial', size=11, bold=True, color='1E3A5F')
normal_font = Font(name='Arial', size=10, color='1A1F36')
bold_font = Font(name='Arial', size=10, bold=True, color='1A1F36')
red_font = Font(name='Arial', size=10, bold=True, color='C62828')
green_font = Font(name='Arial', size=10, bold=True, color='16A34A')
orange_font = Font(name='Arial', size=10, bold=True, color='F47920')

light_gray = PatternFill('solid', fgColor='F3F4F6')
light_red = PatternFill('solid', fgColor='FEE2E2')
light_green = PatternFill('solid', fgColor='DCFCE7')
light_orange = PatternFill('solid', fgColor='FFEDD5')
light_yellow = PatternFill('solid', fgColor='FEF9C3')

center = Alignment(horizontal='center', vertical='center', wrap_text=True)
left = Alignment(horizontal='left', vertical='center', wrap_text=True)
right = Alignment(horizontal='right', vertical='center', wrap_text=True)


# ==============================================================
# ЛИСТ 1: РЕКОМЕНДАЦИИ ПО ДОБАВЛЕНИЮ SKU В MML
# ==============================================================
ws1 = wb.active
ws1.title = 'Рекомендации MML'

# Заголовок
ws1.merge_cells('A1:I1')
ws1['A1'] = 'РЕКОМЕНДАЦИИ ПО КОРРЕКТИРОВКЕ MML — TITAN'
ws1['A1'].font = title_font
ws1['A1'].alignment = center
ws1.row_dimensions[1].height = 30

ws1.merge_cells('A2:I2')
ws1['A2'] = 'Основание: данные первичных (завод→дистр.) и вторичных (дистр.→розница) продаж TITAN за период'
ws1['A2'].font = Font(name='Arial', size=10, italic=True, color='6B7280')
ws1['A2'].alignment = center
ws1.row_dimensions[2].height = 20

# Заголовки таблицы
headers = [
    '№',
    'Категория',
    'Цвет / RAL',
    'Объём',
    'Вторичка (шт)',
    'Вторичка (АКБ)',
    'Первичка (кг)',
    'Текущий MML\n(классы)',
    'Рекомендация\n(куда добавить)',
]
for col, h in enumerate(headers, 1):
    cell = ws1.cell(row=4, column=col, value=h)
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = center
    cell.border = thin_border
ws1.row_dimensions[4].height = 40

# Данные рекомендаций
recommendations = [
    # (№, категория, цвет, объём, вторичка_шт, вторичка_АКБ, первичка_кг, текущий_MML, рекомендация)
    # === 3в1 ПО РЖАВЧИНЕ — добавить в C и D ===
    (1, '3в1 по ржавчине', 'Чёрный RAL 9005', '0.75 л',
     7738, 393, 13909,
     'A, B',
     'ДОБАВИТЬ в C и D\nТоп-1 по продажам\n(7 738 шт, 393 АКБ)\nБазовый цвет для любого магазина'),
    
    (2, '3в1 по ржавчине', 'Шоколад-коричневый RAL 8017', '0.75 л',
     4524, 287, 9827,
     'A, B',
     'ДОБАВИТЬ в C и D\nТоп-4 по продажам\n(4 524 шт, 287 АКБ)\nВторой по спросу цвет'),
    
    (3, '3в1 по ржавчине', 'Серый RAL 7004', '0.75 л',
     5071, 247, 13362,
     'A, B',
     'ДОБАВИТЬ в C\nТоп-3 по продажам\n(5 071 шт, 247 АКБ)\nХодовой цвет для супермаркета'),
    
    # === 3в1 — НОВЫЕ ЦВЕТА (хорошо продаются, но нет в MML вообще) ===
    (4, '3в1 по ржавчине', 'Зелёный RAL 6005', '0.75 л',
     1893, 111, 3994,
     'НЕТ в MML',
     'ДОБАВИТЬ в A и B\n6-е место по продажам\n(1 893 шт, 111 АКБ)\nЗелёный — спрос на дачные заборы'),
    
    (5, '3в1 по ржавчине', 'Жёлтый RAL 1021', '0.75 л',
     1276, 80, 2524,
     'НЕТ в MML',
     'ДОБАВИТЬ в A\n8-е место по продажам\n(1 276 шт, 80 АКБ)\nСпрос от дорожников/строительной техники'),
    
    # === БЫСТРОСОХНУЩАЯ — добавить цвета (сейчас только чёрный) ===
    (6, 'Быстросохнущая', 'Серый', '0.75 л',
     2695, 199, 4193,
     'НЕТ в MML\n(только чёрный)',
     'ДОБАВИТЬ в A и B\n2-е место по продажам\n(2 695 шт, 199 АКБ)\nСерый — второй цвет после чёрного'),
    
    (7, 'Быстросохнущая', 'Шоколадно-кор. RAL 8017', '0.75 л',
     2232, 193, 2794,
     'НЕТ в MML\n(только чёрный)',
     'ДОБАВИТЬ в A и B\n3-е место по продажам\n(2 232 шт, 193 АКБ)\nКоричневый — ходовой цвет'),
    
    (8, 'Быстросохнущая', 'Белый RAL 9010', '0.75 л',
     1507, 119, 2117,
     'НЕТ в MML\n(только чёрный)',
     'ДОБАВИТЬ в A\n4-е место по продажам\n(1 507 шт, 119 АКБ)\nБелый — для интерьерных работ'),
    
    # === МОЛОТКОВАЯ — добавить серебристый в B, золотистый в A ===
    (9, 'Молотковая', 'Серебристый', '0.8 кг',
     2394, 198, 2190,
     'Только A',
     'ДОБАВИТЬ в B\n3-е место по продажам\n(2 394 шт, 198 АКБ)\nСеребристый востребован в локальных сетях'),
    
    (10, 'Молотковая', 'Золотистый', '0.8 кг',
     1153, 98, 873,
     'НЕТ в MML',
     'ДОБАВИТЬ в A\n5-е место по продажам\n(1 153 шт, 98 АКБ)\nЗолотистый — декоративный эффект'),
    
    (11, 'Молотковая', 'Медный', '0.8 кг',
     1162, 102, 634,
     'НЕТ в MML',
     'ДОБАВИТЬ в A\n4-е место по продажам\n(1 162 шт, 102 АКБ)\nМедный — растущий спрос'),
]

row_num = 5
for rec in recommendations:
    num, cat, color, vol, sec_qty, sec_akb, prim_kg, current_mml, recommendation = rec
    
    # Цвет фона по категории
    if '3в1' in cat:
        bg = light_red
    elif 'Быстро' in cat:
        bg = light_orange
    else:
        bg = light_yellow
    
    # Цвет фона для рекомендации
    if 'ДОБАВИТЬ в C и D' in recommendation:
        rec_bg = PatternFill('solid', fgColor='16A34A')  # зелёный
        rec_font = Font(name='Arial', size=9, bold=True, color='FFFFFF')
    elif 'ДОБАВИТЬ в A и B' in recommendation:
        rec_bg = PatternFill('solid', fgColor='F47920')  # оранжевый
        rec_font = Font(name='Arial', size=9, bold=True, color='FFFFFF')
    elif 'ДОБАВИТЬ в A\n' in recommendation or recommendation == 'ДОБАВИТЬ в A':
        rec_bg = PatternFill('solid', fgColor='FBBF24')  # жёлтый
        rec_font = Font(name='Arial', size=9, bold=True, color='1A1F36')
    else:
        rec_bg = bg
        rec_font = normal_font
    
    values = [num, cat, color, vol, sec_qty, sec_akb, prim_kg, current_mml, recommendation]
    for col, val in enumerate(values, 1):
        cell = ws1.cell(row=row_num, column=col, value=val)
        cell.border = thin_border
        cell.alignment = center if col != 9 else left
        
        if col == 1:
            cell.font = bold_font
            cell.fill = bg
        elif col == 2:
            cell.font = bold_font
            cell.fill = bg
        elif col == 3:
            cell.font = normal_font
            cell.fill = bg
        elif col == 4:
            cell.font = normal_font
            cell.fill = bg
        elif col == 5:
            cell.font = bold_font if sec_qty > 1000 else normal_font
            cell.fill = bg
            cell.number_format = '#,##0'
        elif col == 6:
            cell.font = normal_font
            cell.fill = bg
            cell.number_format = '#,##0'
        elif col == 7:
            cell.font = normal_font
            cell.fill = bg
            cell.number_format = '#,##0'
        elif col == 8:
            cell.font = red_font if 'НЕТ' in str(val) else green_font
            cell.fill = bg
        elif col == 9:
            cell.font = rec_font
            cell.fill = rec_bg
    
    ws1.row_dimensions[row_num].height = 60
    row_num += 1

# Ширина колонок
col_widths = [5, 20, 25, 10, 12, 12, 12, 18, 35]
for i, w in enumerate(col_widths, 1):
    ws1.column_dimensions[get_column_letter(i)].width = w

# Итоговая строка
row_num += 1
ws1.merge_cells(f'A{row_num}:I{row_num}')
cell = ws1.cell(row=row_num, column=1, value='ИТОГО: 11 рекомендаций — добавить 11 SKU в MML на основе реальных продаж')
cell.font = Font(name='Arial', size=11, bold=True, color='FFFFFF')
cell.fill = PatternFill('solid', fgColor='1E3A5F')
cell.alignment = center
ws1.row_dimensions[row_num].height = 25


# ==============================================================
# ЛИСТ 2: ПОЛНАЯ ТАБЛИЦА — все цвета по продажам vs MML
# ==============================================================
ws2 = wb.create_sheet('Полный анализ')

ws2.merge_cells('A1:H1')
ws2['A1'] = 'ПОЛНЫЙ АНАЛИЗ: ПРОДАЖИ vs MML — по всем цветам TITAN'
ws2['A1'].font = title_font
ws2['A1'].alignment = center
ws2.row_dimensions[1].height = 30

headers2 = [
    'Категория',
    'Цвет / RAL',
    'Вторичка (шт)',
    'Вторичка (АКБ)',
    'Первичка (кг)',
    'Ранг по\nпродажам',
    'Текущий MML\n(классы)',
    'Статус',
]
for col, h in enumerate(headers2, 1):
    cell = ws2.cell(row=3, column=col, value=h)
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = center
    cell.border = thin_border
ws2.row_dimensions[3].height = 35

# Все данные
all_data = [
    # 3в1 по ржавчине
    ('3в1 по ржавчине', 'Чёрный RAL 9005', 7738, 393, 13909, 'A, B', '✅ В MML'),
    ('3в1 по ржавчине', 'Графитовый серый RAL 7024', 5172, 240, 11625, 'A, B', '✅ В MML'),
    ('3в1 по ржавчине', 'Серый RAL 7004', 5071, 247, 13362, 'A, B', '✅ В MML'),
    ('3в1 по ржавчине', 'Шоколад-коричневый RAL 8017', 4524, 287, 9827, 'A, B', '✅ В MML'),
    ('3в1 по ржавчине', 'Белый RAL 9010', 3273, 225, 6793, 'A, B', '✅ В MML'),
    ('3в1 по ржавчине', 'Зелёный RAL 6005', 1893, 111, 3994, '—', '⚠️ Нет в MML'),
    ('3в1 по ржавчине', 'Серый (без RAL)', 1655, 154, 0, '—', '⚠️ Нет в MML'),
    ('3в1 по ржавчине', 'Жёлтый RAL 1021', 1276, 80, 2524, '—', '⚠️ Нет в MML'),
    ('3в1 по ржавчине', 'Красно-коричневый', 633, 67, 0, '—', 'Низкие продажи'),
    ('3в1 по ржавчине', 'Красно-коричневый RAL 3009', 570, 49, 1339, '—', 'Низкие продажи'),
    # Быстросохнущая
    ('Быстросохнущая', 'Чёрный RAL 9005', 5381, 303, 6846, 'A, B', '✅ В MML'),
    ('Быстросохнущая', 'Серый', 2695, 199, 4193, '—', '⚠️ Нет в MML'),
    ('Быстросохнущая', 'Шоколадно-кор. RAL 8017', 2232, 193, 2794, '—', '⚠️ Нет в MML'),
    ('Быстросохнущая', 'Белый RAL 9010', 1507, 119, 2117, '—', '⚠️ Нет в MML'),
    ('Быстросохнущая', 'Зелёный RAL 6005', 748, 55, 714, '—', 'Низкие продажи'),
    ('Быстросохнущая', 'Шоколадно-коричневый RAL 8017', 704, 78, 0, '—', 'Низкие продажи'),
    ('Быстросохнущая', 'Жёлтый RAL 1023', 365, 29, 473, '—', 'Низкие продажи'),
    # Молотковая
    ('Молотковая', 'Чёрный', 3536, 243, 3713, 'A, B, C', '✅ В MML'),
    ('Молотковая', 'Коричневый', 2913, 207, 3497, 'A, B, C', '✅ В MML'),
    ('Молотковая', 'Серебристый', 2394, 198, 2190, 'Только A', '⚠️ Нет в B'),
    ('Молотковая', 'Медный', 1162, 102, 634, '—', '⚠️ Нет в MML'),
    ('Молотковая', 'Золотистый', 1153, 98, 873, '—', '⚠️ Нет в MML'),
    ('Молотковая', 'Тёмно-зелёный', 425, 70, 370, '—', 'Низкие продажи'),
]

row_num = 4
current_cat = ''
rank = 0
for data in all_data:
    cat, color, sec_qty, sec_akb, prim_kg, mml_classes, status = data
    
    if cat != current_cat:
        current_cat = cat
        rank = 0
    
    rank += 1
    
    # Цвет фона
    if '3в1' in cat:
        bg = light_red
    elif 'Быстро' in cat:
        bg = light_orange
    else:
        bg = light_yellow
    
    if '⚠️' in status:
        status_bg = PatternFill('solid', fgColor='FEF3C7')
        status_font = orange_font
    elif '✅' in status:
        status_bg = light_green
        status_font = green_font
    else:
        status_bg = light_gray
        status_font = Font(name='Arial', size=9, color='6B7280')
    
    values = [cat, color, sec_qty, sec_akb, prim_kg, rank, mml_classes, status]
    for col, val in enumerate(values, 1):
        cell = ws2.cell(row=row_num, column=col, value=val)
        cell.border = thin_border
        
        if col == 1:
            cell.font = bold_font
            cell.fill = bg
            cell.alignment = left
        elif col == 2:
            cell.font = normal_font
            cell.fill = bg
            cell.alignment = left
        elif col in [3, 4, 5]:
            cell.font = bold_font if (col == 3 and sec_qty > 1000) else normal_font
            cell.fill = bg
            cell.alignment = right
            cell.number_format = '#,##0'
        elif col == 6:
            cell.font = bold_font
            cell.fill = bg
            cell.alignment = center
        elif col == 7:
            cell.font = red_font if '—' in str(val) else green_font
            cell.fill = bg
            cell.alignment = center
        elif col == 8:
            cell.font = status_font
            cell.fill = status_bg
            cell.alignment = center
    
    ws2.row_dimensions[row_num].height = 22
    row_num += 1

col_widths2 = [20, 30, 12, 12, 12, 8, 15, 18]
for i, w in enumerate(col_widths2, 1):
    ws2.column_dimensions[get_column_letter(i)].width = w


# ==============================================================
# ЛИСТ 3: СВОДКА ПО КЛАССАМ ТТ
# ==============================================================
ws3 = wb.create_sheet('Сводка по классам')

ws3.merge_cells('A1:F1')
ws3['A1'] = 'СВОДКА: ТЕКУЩИЙ vs РЕКОМЕНДУЕМЫЙ MML ПО КЛАССАМ ТТ'
ws3['A1'].font = title_font
ws3['A1'].alignment = center
ws3.row_dimensions[1].height = 30

headers3 = [
    'Класс ТТ',
    'Канал',
    'Текущий MML\n(SKU TITAN)',
    'Рекомендуемый MML\n(SKU TITAN)',
    'Изменение',
    'Что добавить',
]
for col, h in enumerate(headers3, 1):
    cell = ws3.cell(row=3, column=col, value=h)
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = center
    cell.border = thin_border
ws3.row_dimensions[3].height = 40

class_data = [
    ('A', 'Федеральная сеть DIY', '16', '20 (+4)',
     '+4 SKU',
     '3в1 зелёный RAL 6005 (0.75л)\n3в1 жёлтый RAL 1021 (0.75л)\nБыстросохнущая серый (0.75л)\nМолотковая золотистый (0.8кг)\nМолотковая медный (0.8кг)'),
    
    ('B', 'Локальная сеть', '12', '16 (+4)',
     '+4 SKU',
     '3в1 зелёный RAL 6005 (0.75л)\nБыстросохнущая серый (0.75л)\nБыстросохнущая коричневый (0.75л)\nМолотковая серебристый (0.8кг)'),
    
    ('C', 'Супермаркет', '2', '5 (+3)',
     '+3 SKU',
     '3в1 чёрный RAL 9005 (0.75л)\n3в1 коричневый RAL 8017 (0.75л)\n3в1 серый RAL 7004 (0.75л)'),
    
    ('D', 'Хозмаг', '0', '2 (+2)',
     '+2 SKU',
     '3в1 чёрный RAL 9005 (0.75л)\n3в1 коричневый RAL 8017 (0.75л)'),
]

row_num = 4
for cls, channel, current, recommended, change, what_add in class_data:
    values = [cls, channel, current, recommended, change, what_add]
    
    # Цвет по классу
    cls_colors = {'A': 'C62828', 'B': 'F47920', 'C': '16A34A', 'D': '1E3A5F'}
    cls_color = cls_colors[cls]
    
    for col, val in enumerate(values, 1):
        cell = ws3.cell(row=row_num, column=col, value=val)
        cell.border = thin_border
        
        if col == 1:
            cell.font = Font(name='Arial', size=14, bold=True, color='FFFFFF')
            cell.fill = PatternFill('solid', fgColor=cls_color)
            cell.alignment = center
        elif col == 2:
            cell.font = bold_font
            cell.fill = light_gray
            cell.alignment = left
        elif col == 3:
            cell.font = red_font
            cell.fill = light_red
            cell.alignment = center
        elif col == 4:
            cell.font = green_font
            cell.fill = light_green
            cell.alignment = center
        elif col == 5:
            cell.font = green_font
            cell.fill = light_green
            cell.alignment = center
        elif col == 6:
            cell.font = normal_font
            cell.fill = light_gray
            cell.alignment = left
    
    ws3.row_dimensions[row_num].height = 80
    row_num += 1

# Итог
row_num += 1
ws3.merge_cells(f'A{row_num}:F{row_num}')
cell = ws3.cell(row=row_num, column=1, value='ИТОГО: текущий MML = 30 SKU → рекомендуемый = 43 SKU (+13 SKU основано на реальных продажах)')
cell.font = Font(name='Arial', size=11, bold=True, color='FFFFFF')
cell.fill = PatternFill('solid', fgColor='1E3A5F')
cell.alignment = center
ws3.row_dimensions[row_num].height = 25

col_widths3 = [8, 25, 12, 16, 10, 45]
for i, w in enumerate(col_widths3, 1):
    ws3.column_dimensions[get_column_letter(i)].width = w


# ==============================================================
# СОХРАНЕНИЕ
# ==============================================================
wb.save(OUTPUT)
print(f'OK: {OUTPUT}')
print(f'Листов: {len(wb.sheetnames)}')
for name in wb.sheetnames:
    print(f'  "{name}"')
