"""
Анализ статистики И.П. Рыгалов (Пенза) — TITAN
Сравнение 2025 vs 2026 по классам точек A/B/C/D/E/F/G/J
С полями для замеров после программы TITAN.
"""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import json

# Цвета
ORANGE_FILL = PatternFill('solid', fgColor='F47629')
DARK_FILL = PatternFill('solid', fgColor='313131')
WHITE_FONT = Font(name='Arial', size=11, bold=True, color='FFFFFF')
ORANGE_FONT = Font(name='Arial', size=11, bold=True, color='F47629')
RED_FONT = Font(name='Arial', size=11, bold=True, color='C0392B')
GREEN_FONT = Font(name='Arial', size=11, bold=True, color='16A34A')
NORMAL_FONT = Font(name='Arial', size=10)
ITALIC_FONT = Font(name='Arial', size=10, italic=True, color='6B7280')
THIN_BORDER = Border(
    left=Side(style='thin', color='D0D0D0'), right=Side(style='thin', color='D0D0D0'),
    top=Side(style='thin', color='D0D0D0'), bottom=Side(style='thin', color='D0D0D0'),
)
EMPTY_FILL = PatternFill('solid', fgColor='FFFDE7')  # жёлтый для пустых полей замеров

# Данные из Excel
# Структура: класс -> {итоговые АКБ и продажи} + {SKU -> {АКБ, продажи}}
data = {
    'B': {
        'summary_2025': {'akb': 2, 'sales': 2020678.20},
        'summary_2026': {'akb': 2, 'sales': 1712189.40},
        'sku_count_2025': 28,  # количество SKU с данными
        'sku_count_2026': 25,
    },
    'C': {
        'summary_2025': {'akb': 2, 'sales': 12199.80},
        'summary_2026': {'akb': 2, 'sales': 9890.60},
        'sku_count_2025': 11,
        'sku_count_2026': 10,
    },
    'D': {
        'summary_2025': {'akb': 25, 'sales': 76398.40},
        'summary_2026': {'akb': 29, 'sales': 226658.80},
        'sku_count_2025': 21,
        'sku_count_2026': 24,
    },
    'E': {
        'summary_2025': {'akb': 36, 'sales': 349602.00},
        'summary_2026': {'akb': 42, 'sales': 263856.80},
        'sku_count_2025': 23,
        'sku_count_2026': 23,
    },
    'F': {
        'summary_2025': {'akb': 2, 'sales': 11994.00},
        'summary_2026': {'akb': 1, 'sales': 4064.00},
        'sku_count_2025': 6,
        'sku_count_2026': 3,
    },
    'G': {
        'summary_2025': {'akb': 3, 'sales': 17711.20},
        'summary_2026': {'akb': 1, 'sales': 5279.00},
        'sku_count_2025': 11,
        'sku_count_2026': 4,
    },
    'J': {
        'summary_2025': {'akb': 12, 'sales': 67199.50},
        'summary_2026': {'akb': 15, 'sales': 80181.90},
        'sku_count_2025': 15,
        'sku_count_2026': 14,
    },
}

# Итого по всем
total_2025_akb = 82
total_2025_sales = 2555783.10
total_2026_akb = 97
total_2026_sales = 2314763.10

# Рекомендации из Sales Guide
recommended = {
    'A': {'sku': 23, 'desc': 'Федеральная сеть, полная матрица'},
    'B': {'sku': 23, 'desc': 'Региональная сеть, расширенная матрица'},
    'C': {'sku': 16, 'desc': 'Супермаркет, топ-цвета + молотковая'},
    'D': {'sku': 6, 'desc': 'Хозмаг, топ-6 цветов'},
}

wb = Workbook()

# ============================================================
# ЛИСТ 1: СВОДКА ПО КЛАССАМ
# ============================================================
ws1 = wb.active
ws1.title = 'Сводка по классам'

ws1.merge_cells('A1:K1')
ws1['A1'] = 'TITAN — Статистика И.П. Рыгалов (Пенза) | Анализ 2025 vs 2026 + поля для замеров после программы'
ws1['A1'].font = WHITE_FONT
ws1['A1'].fill = ORANGE_FILL
ws1['A1'].alignment = Alignment(horizontal='center', vertical='center')
ws1.row_dimensions[1].height = 30

ws1.merge_cells('A2:K2')
ws1['A2'] = 'Источник: выгрузка продаж И.П. Рыгалов 09.09.2026 | Критерий 1: соблюдение рекомендованного количества SKU TITAN в точках класса A, B, C, D'
ws1['A2'].font = ITALIC_FONT
ws1['A2'].alignment = Alignment(horizontal='center')

# Заголовки
headers = ['Класс ТТ', 'Описание', 'Реком. SKU', 'SKU 2025', 'SKU 2026', 'Δ SKU 25→26',
           'АКБ 2025', 'АКБ 2026', 'Продажи 2025, ₽', 'Продажи 2026, ₽', 'Δ продаж, %']
for i, h in enumerate(headers, 1):
    c = ws1.cell(row=4, column=i, value=h)
    c.font = WHITE_FONT
    c.fill = DARK_FILL
    c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    c.border = THIN_BORDER
ws1.row_dimensions[4].height = 40

# Данные
classes_order = ['B', 'C', 'D', 'E', 'F', 'G', 'J']
class_desc = {
    'B': 'Региональная сеть (расширенная матрица)',
    'C': 'Супермаркет (топ-цвета + молотковая)',
    'D': 'Хозмаг (топ-6 цветов)',
    'E': 'Малая ТТ (ниже D)',
    'F': 'Очень малая ТТ',
    'G': 'Микро-ТТ',
    'J': 'Спец. канал',
}

for ri, cls in enumerate(classes_order, 1):
    d = data[cls]
    rec_sku = recommended.get(cls, {}).get('sku', '—')
    sku25 = d['sku_count_2025']
    sku26 = d['sku_count_2026']
    delta_sku = sku26 - sku25
    akb25 = d['summary_2025']['akb']
    akb26 = d['summary_2026']['akb']
    sales25 = d['summary_2025']['sales']
    sales26 = d['summary_2026']['sales']
    delta_sales = ((sales26 - sales25) / sales25 * 100) if sales25 else 0

    row_data = [cls, class_desc.get(cls, ''), rec_sku, sku25, sku26, delta_sku,
                akb25, akb26, f'{sales25:,.0f}', f'{sales26:,.0f}', f'{delta_sales:+.1f}%']
    for ci, v in enumerate(row_data, 1):
        c = ws1.cell(row=4 + ri, column=ci, value=v)
        c.font = NORMAL_FONT
        c.border = THIN_BORDER
        if ci == 1:
            c.font = ORANGE_FONT
            c.alignment = Alignment(horizontal='center')
        elif ci == 3:
            c.alignment = Alignment(horizontal='center')
            c.font = Font(name='Arial', size=10, bold=True, color='1E3A5F')
        elif ci == 6:
            c.alignment = Alignment(horizontal='center')
            if delta_sku > 0:
                c.font = GREEN_FONT
            elif delta_sku < 0:
                c.font = RED_FONT
        elif ci in (7, 8):
            c.alignment = Alignment(horizontal='center')
        elif ci in (9, 10):
            c.alignment = Alignment(horizontal='center')
        elif ci == 11:
            c.alignment = Alignment(horizontal='center')
            if delta_sales > 0:
                c.font = GREEN_FONT
            elif delta_sales < 0:
                c.font = RED_FONT

# Итого
row_idx = 4 + len(classes_order) + 1
total_delta_sales = ((total_2026_sales - total_2025_sales) / total_2025_sales * 100)
row_data = ['ИТОГО', 'Все классы', '—', '—', '—', '—',
            total_2025_akb, total_2026_akb, f'{total_2025_sales:,.0f}', f'{total_2026_sales:,.0f}', f'{total_delta_sales:+.1f}%']
for ci, v in enumerate(row_data, 1):
    c = ws1.cell(row=row_idx, column=ci, value=v)
    c.font = Font(name='Arial', size=10, bold=True, color='FFFFFF')
    c.fill = ORANGE_FILL
    c.border = THIN_BORDER
    c.alignment = Alignment(horizontal='center')

# Колонки
for col, w in {'A': 10, 'B': 30, 'C': 10, 'D': 10, 'E': 10, 'F': 12,
               'G': 10, 'H': 10, 'I': 16, 'J': 16, 'K': 12}.items():
    ws1.column_dimensions[col].width = w

# ============================================================
# ЛИСТ 2: ПОЛЯ ДЛЯ ЗАМЕРОВ ПОСЛЕ ПРОГРАММЫ
# ============================================================
ws2 = wb.create_sheet('Замеры после программы')

ws2.merge_cells('A1:J1')
ws2['A1'] = 'ЗАМЕРЫ ПОСЛЕ ПРОГРАММЫ TITAN — пустые поля для заполнения'
ws2['A1'].font = WHITE_FONT
ws2['A1'].fill = ORANGE_FILL
ws2['A1'].alignment = Alignment(horizontal='center', vertical='center')
ws2.row_dimensions[1].height = 30

ws2.merge_cells('A2:J2')
ws2['A2'] = 'Жёлтые поля — для заполнения после завершения программы TITAN (план: 6 месяцев)'
ws2['A2'].font = ITALIC_FONT
ws2['A2'].alignment = Alignment(horizontal='center')

headers2 = ['Класс ТТ', 'Реком. SKU', 'SKU до программы (2026)', 'SKU после программы',
            'Δ SKU', 'АКБ до (2026)', 'АКБ после', 'Продажи до, ₽ (2026)',
            'Продажи после, ₽', 'Δ продаж, %']
for i, h in enumerate(headers2, 1):
    c = ws2.cell(row=4, column=i, value=h)
    c.font = WHITE_FONT
    c.fill = DARK_FILL
    c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    c.border = THIN_BORDER
ws2.row_dimensions[4].height = 40

for ri, cls in enumerate(classes_order, 1):
    d = data[cls]
    rec_sku = recommended.get(cls, {}).get('sku', '—')
    sku26 = d['sku_count_2026']
    akb26 = d['summary_2026']['akb']
    sales26 = d['summary_2026']['sales']

    # Колонки 1-3, 7-8 — заполнены (до программы)
    # Колонки 4, 5, 8, 9, 10 — пустые (после программы) — жёлтый фон
    row_data = [cls, rec_sku, sku26, '', '', akb26, '', f'{sales26:,.0f}', '', '']
    for ci, v in enumerate(row_data, 1):
        c = ws2.cell(row=4 + ri, column=ci, value=v)
        c.font = NORMAL_FONT
        c.border = THIN_BORDER
        c.alignment = Alignment(horizontal='center')
        if ci in (4, 5, 7, 9, 10):
            c.fill = EMPTY_FILL  # жёлтый — для заполнения
        if ci == 1:
            c.font = ORANGE_FONT

# Итого
row_idx2 = 4 + len(classes_order) + 1
row_data2 = ['ИТОГО', '—', '—', '', '', total_2026_akb, '', f'{total_2026_sales:,.0f}', '', '']
for ci, v in enumerate(row_data2, 1):
    c = ws2.cell(row=row_idx2, column=ci, value=v)
    c.font = Font(name='Arial', size=10, bold=True, color='FFFFFF')
    c.fill = ORANGE_FILL
    c.border = THIN_BORDER
    c.alignment = Alignment(horizontal='center')

for col, w in {'A': 10, 'B': 10, 'C': 15, 'D': 15, 'E': 10, 'F': 12,
               'G': 12, 'H': 16, 'I': 16, 'J': 12}.items():
    ws2.column_dimensions[col].width = w

# ============================================================
# ЛИСТ 3: ТРЕНДЫ
# ============================================================
ws3 = wb.create_sheet('Тренды 2025→2026')

ws3.merge_cells('A1:H1')
ws3['A1'] = 'ТРЕНДЫ 2025 → 2026 — анализ изменений'
ws3['A1'].font = WHITE_FONT
ws3['A1'].fill = ORANGE_FILL
ws3['A1'].alignment = Alignment(horizontal='center', vertical='center')
ws3.row_dimensions[1].height = 30

headers3 = ['Класс', 'SKU 2025', 'SKU 2026', 'Δ SKU', 'Тренд SKU',
            'Продажи 2025, ₽', 'Продажи 2026, ₽', 'Тренд продаж']
for i, h in enumerate(headers3, 1):
    c = ws3.cell(row=3, column=i, value=h)
    c.font = WHITE_FONT
    c.fill = DARK_FILL
    c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    c.border = THIN_BORDER
ws3.row_dimensions[3].height = 35

for ri, cls in enumerate(classes_order, 1):
    d = data[cls]
    sku25 = d['sku_count_2025']
    sku26 = d['sku_count_2026']
    delta_sku = sku26 - sku25
    sales25 = d['summary_2025']['sales']
    sales26 = d['summary_2026']['sales']
    delta_sales = ((sales26 - sales25) / sales25 * 100) if sales25 else 0

    if delta_sku > 0:
        trend_sku = '🟢 Рост'
    elif delta_sku < 0:
        trend_sku = '🔴 Снижение'
    else:
        trend_sku = '⚪ Стабильно'

    if delta_sales > 10:
        trend_sales = '🟢 Рост'
    elif delta_sales < -10:
        trend_sales = '🔴 Снижение'
    else:
        trend_sales = '🟡 Стабильно'

    row_data = [cls, sku25, sku26, delta_sku, trend_sku,
                f'{sales25:,.0f}', f'{sales26:,.0f}', f'{delta_sales:+.1f}% {trend_sales}']
    for ci, v in enumerate(row_data, 1):
        c = ws3.cell(row=3 + ri, column=ci, value=v)
        c.font = NORMAL_FONT
        c.border = THIN_BORDER
        c.alignment = Alignment(horizontal='center')
        if ci == 1:
            c.font = ORANGE_FONT
        elif ci == 4:
            if delta_sku > 0: c.font = GREEN_FONT
            elif delta_sku < 0: c.font = RED_FONT

# Итого
row_idx3 = 3 + len(classes_order) + 1
row_data3 = ['ИТОГО', '—', '—', '—', '—',
             f'{total_2025_sales:,.0f}', f'{total_2026_sales:,.0f}',
             f'{total_delta_sales:+.1f}%']
for ci, v in enumerate(row_data3, 1):
    c = ws3.cell(row=row_idx3, column=ci, value=v)
    c.font = Font(name='Arial', size=10, bold=True, color='FFFFFF')
    c.fill = ORANGE_FILL
    c.border = THIN_BORDER
    c.alignment = Alignment(horizontal='center')

for col, w in {'A': 10, 'B': 10, 'C': 10, 'D': 10, 'E': 15, 'F': 16, 'G': 16, 'H': 20}.items():
    ws3.column_dimensions[col].width = w

# ============================================================
# ЛИСТ 4: ДЕТАЛИЗАЦИЯ ПО SKU (класс B)
# ============================================================
ws4 = wb.create_sheet('Детализация класс B')
ws4.merge_cells('A1:E1')
ws4['A1'] = 'ДЕТАЛИЗАЦИЯ ПО SKU — КЛАСС B (2 точки, 28→25 SKU)'
ws4['A1'].font = WHITE_FONT
ws4['A1'].fill = ORANGE_FILL
ws4['A1'].alignment = Alignment(horizontal='center')

headers4 = ['SKU', 'АКБ 2025', 'Продажи 2025, ₽', 'АКБ 2026', 'Продажи 2026, ₽']
for i, h in enumerate(headers4, 1):
    c = ws4.cell(row=3, column=i, value=h)
    c.font = WHITE_FONT
    c.fill = DARK_FILL
    c.alignment = Alignment(horizontal='center')
    c.border = THIN_BORDER

# Читаем данные из исходного файла
from openpyxl import load_workbook as lw
src = lw('/home/z/my-project/upload/Титан Пенза И.П, Рыгалов 09.09.2026.xlsx', data_only=True)
src_ws = src['Лист1']

row = 4
in_class_b = False
for r in src_ws.iter_rows(min_row=7, values_only=True):
    name = r[2] if r[2] else ''
    akb25 = r[3] if r[3] else ''
    sales25 = r[4] if r[4] else ''
    akb26 = r[5] if r[5] else ''
    sales26 = r[6] if r[6] else ''

    # Определяем класс
    if name in ('B', 'C', 'D', 'E', 'F', 'G', 'J'):
        in_class_b = (name == 'B')
        continue

    if in_class_b and name:
        row_data = [str(name)[:70], akb25 if akb25 else '—',
                    f'{sales25:,.0f}' if sales25 and isinstance(sales25, (int, float)) else '—',
                    akb26 if akb26 else '—',
                    f'{sales26:,.0f}' if sales26 and isinstance(sales26, (int, float)) else '—']
        for ci, v in enumerate(row_data, 1):
            c = ws4.cell(row=row, column=ci, value=v)
            c.font = NORMAL_FONT
            c.border = THIN_BORDER
            if ci == 1:
                c.alignment = Alignment(horizontal='left')
            else:
                c.alignment = Alignment(horizontal='center')
        row += 1

for col, w in {'A': 50, 'B': 10, 'C': 16, 'D': 10, 'E': 16}.items():
    ws4.column_dimensions[col].width = w

# Сохранение
out_path = '/home/z/my-project/download/TITAN_Рыгалов_анализ_статистики.xlsx'
wb.save(out_path)
print(f'OK saved: {out_path}')
import os
print(f'Size: {os.path.getsize(out_path)} bytes')
