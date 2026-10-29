"""Создание файла анализа статистики TITAN — Ваш Дом.
   ТОЛЬКО ОРЕОЛ TITAN (грунт-эмали 3 в 1).
   Структура: 24 магазина, данные 2026 (Jan-Sep), MML=23.
"""
import json
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

with open('/home/z/my-project/scripts/vashdom_stats.json', 'r', encoding='utf-8') as f:
    stats = json.load(f)

# Стилизация
ORANGE_FILL = PatternFill('solid', fgColor='F47629')
DARK_FILL = PatternFill('solid', fgColor='313131')
WHITE_FONT = Font(name='Arial', size=11, bold=True, color='FFFFFF')
ORANGE_FONT = Font(name='Arial', size=11, bold=True, color='F47629')
RED_FONT = Font(name='Arial', size=11, bold=True, color='C0392B')
GREEN_FONT = Font(name='Arial', size=11, bold=True, color='16A34A')
NORMAL_FONT = Font(name='Arial', size=10)
TITLE_FONT = Font(name='Arial', size=14, bold=True, color='313131')
SUBTITLE_FONT = Font(name='Arial', size=10, italic=True, color='666666')
MML_FONT = Font(name='Arial', size=10, bold=True, color='1E3A5F')
THIN_BORDER = Border(
    left=Side(style='thin', color='D0D0D0'),
    right=Side(style='thin', color='D0D0D0'),
    top=Side(style='thin', color='D0D0D0'),
    bottom=Side(style='thin', color='D0D0D0')
)
EMPTY_FILL = PatternFill('solid', fgColor='FFFDE7')

CENTER = Alignment(horizontal='center', vertical='center', wrap_text=True)
LEFT = Alignment(horizontal='left', vertical='center', wrap_text=True)

# MML
MML = 23

# Порядок цветов RAL (по убыванию продаж)
colors_order = [
    'RAL 7004 — Серый',
    'RAL 7024 — Графитовый серый',
    'RAL 8017 — Шоколадно-коричневый',
    'RAL 9005 — Чёрный',
    'RAL 9010 — Белый',
    'RAL 1021 — Жёлтый',
    'RAL 6005 — Зелёный',
    'RAL 3009 — Красно-коричневый',
    'Прочее',
]

# Создаём книгу
wb = Workbook()

# ============================================================
# ЛИСТ 1: Сводка по цветам RAL
# ============================================================
ws1 = wb.active
ws1.title = 'Сводка по цветам RAL'

ws1.cell(row=1, column=1, value='TITAN — Статистика Ваш Дом (Ростов-на-Дону)').font = TITLE_FONT
ws1.merge_cells(start_row=1, start_column=1, end_row=1, end_column=10)

ws1.cell(row=2, column=1, value='ОРЕОЛ TITAN — грунт-эмали 3 в 1 | Анализ продаж 2026 (янв-сен)').font = SUBTITLE_FONT
ws1.merge_cells(start_row=2, start_column=1, end_row=2, end_column=10)

ws1.cell(row=3, column=1, value='Источник: выгрузка Ваш Дом 09.09.2026 | Период: 01.01.2026 - 06.09.2026 | MML (реком.): 23 SKU').font = SUBTITLE_FONT
ws1.merge_cells(start_row=3, start_column=1, end_row=3, end_column=10)

ws1.row_dimensions[1].height = 22
ws1.row_dimensions[2].height = 18
ws1.row_dimensions[3].height = 18

# Заголовок таблицы
headers = ['Цвет RAL', 'SKU всего', 'Активных ТТ (макс.)',
           'Поступило, шт.', 'Остаток, шт.', 'Продано, шт.', 'Доля продаж, %',
           'Продано после', 'Δ продаж, шт.', 'Δ продаж, %']
for ci, h in enumerate(headers, 1):
    c = ws1.cell(row=4, column=ci, value=h)
    c.font = WHITE_FONT; c.fill = DARK_FILL
    c.alignment = CENTER; c.border = THIN_BORDER
ws1.row_dimensions[4].height = 42

# Данные
tot_sku = tot_qty = tot_stock = tot_sold = 0
existing_colors = []
for ri, color in enumerate(colors_order, 1):
    d = stats['categories'].get(color, None)
    if not d or d['sku_count'] == 0:
        continue
    existing_colors.append(color)
    sku_count = d['sku_count']
    qty = d['total_qty']
    stock = sum(sku['stock'] for sku in d.get('sku_list', []))
    sold = d['total_sales']
    active_shops_max = max((sku['active_shops'] for sku in d.get('sku_list', [])), default=0)

    row_data = [color, sku_count, active_shops_max,
                int(qty), int(stock), int(sold), '',
                '', '', '']
    for ci, v in enumerate(row_data, 1):
        c = ws1.cell(row=4 + ri, column=ci, value=v)
        c.font = NORMAL_FONT; c.border = THIN_BORDER
        c.alignment = CENTER if ci != 1 else LEFT
        if ci == 1:
            c.font = ORANGE_FONT
        elif ci in (8, 9, 10):
            c.fill = EMPTY_FILL

    tot_sku += sku_count; tot_qty += qty; tot_stock += stock; tot_sold += sold

# Доля продаж
for ri, color in enumerate(existing_colors, 1):
    sold = stats['categories'][color]['total_sales']
    share = (sold / tot_sold * 100) if tot_sold > 0 else 0
    c = ws1.cell(row=4 + ri, column=7, value=f'{share:.1f}%')
    c.font = NORMAL_FONT; c.alignment = CENTER; c.border = THIN_BORDER

# Итог
row_idx = 4 + len(existing_colors) + 1
row_data = ['ИТОГО', tot_sku, '', int(tot_qty), int(tot_stock), int(tot_sold),
            '100.0%', '', '', '']
for ci, v in enumerate(row_data, 1):
    c = ws1.cell(row=row_idx, column=ci, value=v)
    c.font = Font(name='Arial', size=10, bold=True, color='FFFFFF')
    c.fill = ORANGE_FILL; c.border = THIN_BORDER
    c.alignment = CENTER if ci != 1 else LEFT

# MML compliance note
note_row = row_idx + 2
c = ws1.cell(row=note_row, column=1,
             value=f'Соблюдение MML: {tot_sku} SKU из {MML} рекомендованных → {tot_sku/MML*100:.0f}% ({"✅ соответствует" if tot_sku >= MML else "⚠️ ниже нормы"})')
c.font = SUBTITLE_FONT
ws1.merge_cells(start_row=note_row, start_column=1, end_row=note_row, end_column=10)

# Ширины
widths = {'A': 36, 'B': 12, 'C': 16, 'D': 14, 'E': 14, 'F': 14, 'G': 14, 'H': 14, 'I': 14, 'J': 14}
for col, w in widths.items():
    ws1.column_dimensions[col].width = w

ws1.freeze_panes = 'A5'

# ============================================================
# ЛИСТ 2: Сводка по магазинам
# ============================================================
ws2 = wb.create_sheet('Сводка по магазинам')

ws2.cell(row=1, column=1, value='TITAN — Сводка по магазинам Ваш Дом').font = TITLE_FONT
ws2.merge_cells(start_row=1, start_column=1, end_row=1, end_column=8)
ws2.cell(row=2, column=1, value='ОРЕОЛ TITAN | 24 магазина в Ростове-на-Дону и области | Период: 01.01.2026 - 06.09.2026').font = SUBTITLE_FONT
ws2.merge_cells(start_row=2, start_column=1, end_row=2, end_column=8)
ws2.row_dimensions[1].height = 22
ws2.row_dimensions[2].height = 18

headers2 = ['№', 'Магазин (адрес)', 'SKU TITAN', 'Поступило, шт.', 'Остаток, шт.',
            'Продано, шт.', 'Доля, %', 'Статус']
for ci, h in enumerate(headers2, 1):
    c = ws2.cell(row=4, column=ci, value=h)
    c.font = WHITE_FONT; c.fill = DARK_FILL
    c.alignment = CENTER; c.border = THIN_BORDER
ws2.row_dimensions[4].height = 38

shop_totals = stats['shop_totals']
sorted_shops = sorted(shop_totals.items(), key=lambda x: -x[1]['sales'])
tot_shop_sold = sum(d['sales'] for d in shop_totals.values())

for ri, (shop_name, d) in enumerate(sorted_shops, 1):
    share = (d['sales'] / tot_shop_sold * 100) if tot_shop_sold > 0 else 0
    if d['sales'] >= 100:
        status = '🟢 Топ-точка'
    elif d['sales'] >= 30:
        status = '🟡 Средняя'
    elif d['sales'] >= 10:
        status = '🟠 Низкая'
    else:
        status = '🔴 Слабая'
    row_data = [ri, shop_name, d['sku_count'], int(d['qty']), int(d['stock']), int(d['sales']),
                f'{share:.1f}%', status]
    for ci, v in enumerate(row_data, 1):
        c = ws2.cell(row=4 + ri, column=ci, value=v)
        c.font = NORMAL_FONT; c.border = THIN_BORDER
        c.alignment = CENTER if ci != 2 else LEFT
        if ci == 2:
            c.font = Font(name='Arial', size=10, bold=True, color='1E3A5F')
        elif ci == 6:
            if d['sales'] >= 100:
                c.font = GREEN_FONT
            elif d['sales'] >= 30:
                c.font = Font(name='Arial', size=10, bold=True, color='D97706')
            elif d['sales'] < 10:
                c.font = RED_FONT
        elif ci == 8:
            if '🟢' in str(v):
                c.font = GREEN_FONT
            elif '🔴' in str(v):
                c.font = RED_FONT

# Итог
row_idx2 = 4 + len(sorted_shops) + 1
tot_qty_shop = sum(d['qty'] for d in shop_totals.values())
tot_stock_shop = sum(d['stock'] for d in shop_totals.values())
row_data2 = ['', 'ИТОГО', '', int(tot_qty_shop), int(tot_stock_shop), int(tot_shop_sold), '100.0%', '—']
for ci, v in enumerate(row_data2, 1):
    c = ws2.cell(row=row_idx2, column=ci, value=v)
    c.font = Font(name='Arial', size=10, bold=True, color='FFFFFF')
    c.fill = ORANGE_FILL; c.border = THIN_BORDER
    c.alignment = CENTER if ci != 2 else LEFT

# Неактивные магазины
inactive = [s for s in stats['shops'] if s not in shop_totals]
if inactive:
    row_idx2 += 2
    c = ws2.cell(row=row_idx2, column=1,
                 value=f'Неактивные магазины ({len(inactive)} из 24 — нет продаж TITAN за период):')
    c.font = Font(name='Arial', size=10, italic=True, bold=True, color='C0392B')
    ws2.merge_cells(start_row=row_idx2, start_column=1, end_row=row_idx2, end_column=8)
    for ri, shop in enumerate(inactive, 1):
        c = ws2.cell(row=row_idx2 + ri, column=2, value=f'  • {shop}')
        c.font = NORMAL_FONT; c.alignment = LEFT

widths2 = {'A': 6, 'B': 38, 'C': 10, 'D': 14, 'E': 14, 'F': 14, 'G': 10, 'H': 16}
for col, w in widths2.items():
    ws2.column_dimensions[col].width = w

ws2.freeze_panes = 'A5'

# ============================================================
# ЛИСТ 3: Замеры после программы
# ============================================================
ws3 = wb.create_sheet('Замеры после программы')

ws3.cell(row=1, column=1, value='ЗАМЕРЫ ПОСЛЕ ПРОГРАММЫ TITAN — жёлтые поля для заполнения').font = TITLE_FONT
ws3.merge_cells(start_row=1, start_column=1, end_row=1, end_column=10)
ws3.cell(row=2, column=1, value='Жёлтые поля — для заполнения после завершения программы (план: 6 месяцев)').font = SUBTITLE_FONT
ws3.merge_cells(start_row=2, start_column=1, end_row=2, end_column=10)
ws3.row_dimensions[1].height = 22
ws3.row_dimensions[2].height = 18

headers3 = ['Цвет RAL / Магазин', 'MML', 'SKU до (2026)', 'SKU после', 'Δ SKU',
            'Продано до, шт.', 'Продано после, шт.', 'Δ продаж, шт.', 'Δ продаж, %', 'Остаток, шт.']
for ci, h in enumerate(headers3, 1):
    c = ws3.cell(row=4, column=ci, value=h)
    c.font = WHITE_FONT; c.fill = DARK_FILL
    c.alignment = CENTER; c.border = THIN_BORDER
ws3.row_dimensions[4].height = 42

# По цветам RAL
for ri, color in enumerate(existing_colors, 1):
    d = stats['categories'][color]
    row_data = [color, MML if ri == 1 else '—', d['sku_count'], '', '', int(d['total_sales']), '', '', '', '']
    for ci, v in enumerate(row_data, 1):
        c = ws3.cell(row=4 + ri, column=ci, value=v)
        c.font = NORMAL_FONT; c.border = THIN_BORDER
        c.alignment = CENTER if ci != 1 else LEFT
        if ci in (4, 5, 7, 8, 9):
            c.fill = EMPTY_FILL
        if ci == 1:
            c.font = ORANGE_FONT
        elif ci == 2 and ri == 1:
            c.font = MML_FONT

# Итог по цветам
row_idx3 = 4 + len(existing_colors) + 1
row_data3 = ['ИТОГО по цветам', MML, tot_sku, '', '', int(tot_sold), '', '', '', '']
for ci, v in enumerate(row_data3, 1):
    c = ws3.cell(row=row_idx3, column=ci, value=v)
    c.font = Font(name='Arial', size=10, bold=True, color='FFFFFF')
    c.fill = ORANGE_FILL; c.border = THIN_BORDER
    c.alignment = CENTER if ci != 1 else LEFT

# По топ-10 магазинам
row_idx3 += 2
c = ws3.cell(row=row_idx3, column=1, value='Топ-10 магазинов для замеров (по продажам ОРЕОЛ TITAN):')
c.font = SUBTITLE_FONT
row_idx3 += 1
# Заголовок для топ-10
for ci, h in enumerate(headers3, 1):
    c = ws3.cell(row=row_idx3, column=ci, value=h)
    c.font = WHITE_FONT; c.fill = DARK_FILL
    c.alignment = CENTER; c.border = THIN_BORDER
row_idx3 += 1
for ri, (shop_name, d) in enumerate(sorted_shops[:10], 1):
    row_data = [shop_name, '—', d['sku_count'], '', '', int(d['sales']), '', '', '', '']
    for ci, v in enumerate(row_data, 1):
        c = ws3.cell(row=row_idx3 + ri, column=ci, value=v)
        c.font = NORMAL_FONT; c.border = THIN_BORDER
        c.alignment = CENTER if ci != 1 else LEFT
        if ci in (4, 5, 7, 8, 9):
            c.fill = EMPTY_FILL
        if ci == 1:
            c.font = Font(name='Arial', size=10, bold=True, color='1E3A5F')

widths3 = {'A': 36, 'B': 10, 'C': 14, 'D': 14, 'E': 10, 'F': 16, 'G': 16, 'H': 16, 'I': 12, 'J': 12}
for col, w in widths3.items():
    ws3.column_dimensions[col].width = w

ws3.freeze_panes = 'A5'

# ============================================================
# ЛИСТ 4: Детализация по SKU
# ============================================================
ws4 = wb.create_sheet('Детализация по SKU')

ws4.cell(row=1, column=1, value='ДЕТАЛИЗАЦИЯ ПО SKU — ОРЕОЛ TITAN (грунт-эмали 3 в 1)').font = TITLE_FONT
ws4.merge_cells(start_row=1, start_column=1, end_row=1, end_column=8)
ws4.cell(row=2, column=1, value=f'MML (рекомендованный): {MML} SKU | Период: 01.01.2026 - 06.09.2026 | Сортировка: по продажам (убывание)').font = SUBTITLE_FONT
ws4.merge_cells(start_row=2, start_column=1, end_row=2, end_column=8)
ws4.row_dimensions[1].height = 22
ws4.row_dimensions[2].height = 18

headers4 = ['Код', 'Наименование SKU', 'ABC', 'Статус', 'Активных ТТ', 'Поступило, шт.', 'Остаток, шт.', 'Продано, шт.']
for ci, h in enumerate(headers4, 1):
    c = ws4.cell(row=4, column=ci, value=h)
    c.font = WHITE_FONT; c.fill = DARK_FILL
    c.alignment = CENTER; c.border = THIN_BORDER
ws4.row_dimensions[4].height = 38

# Сбор всех SKU
all_skus = []
for color, d in stats['categories'].items():
    for sku in d['sku_list']:
        all_skus.append({**sku, 'color': color})

all_skus.sort(key=lambda x: -x['sales'])

for ri, sku in enumerate(all_skus, 1):
    row_data = [sku['code'], sku['name'], sku['abc'], sku['status'],
                sku['active_shops'], int(sku['qty']), int(sku['stock']), int(sku['sales'])]
    for ci, v in enumerate(row_data, 1):
        c = ws4.cell(row=4 + ri, column=ci, value=v)
        c.font = NORMAL_FONT; c.border = THIN_BORDER
        c.alignment = CENTER if ci != 2 else LEFT
        if ci == 5:  # Активных ТТ
            if sku['active_shops'] >= 10:
                c.font = GREEN_FONT
            elif sku['active_shops'] < 5:
                c.font = RED_FONT
        elif ci == 8:  # Продано
            if sku['sales'] >= 50:
                c.font = GREEN_FONT
            elif sku['sales'] < 10:
                c.font = RED_FONT
        elif ci == 4:  # Статус
            if 'Вывод' in str(v):
                c.font = RED_FONT
            elif 'Действующий' in str(v):
                c.font = GREEN_FONT

# Итог
row_idx4 = 4 + len(all_skus) + 1
row_data4 = ['', f'ИТОГО ОРЕОЛ TITAN ({len(all_skus)} SKU из {MML} реком.)', '', '', '', int(tot_qty), int(tot_stock), int(tot_sold)]
for ci, v in enumerate(row_data4, 1):
    c = ws4.cell(row=row_idx4, column=ci, value=v)
    c.font = Font(name='Arial', size=10, bold=True, color='FFFFFF')
    c.fill = ORANGE_FILL; c.border = THIN_BORDER
    c.alignment = CENTER if ci != 2 else LEFT

widths4 = {'A': 10, 'B': 60, 'C': 8, 'D': 14, 'E': 12, 'F': 14, 'G': 14, 'H': 14}
for col, w in widths4.items():
    ws4.column_dimensions[col].width = w

ws4.freeze_panes = 'A5'

# ============================================================
# ЛИСТ 5: Матрица SKU × Магазин
# ============================================================
ws5 = wb.create_sheet('Матрица SKU × Магазин')

ws5.cell(row=1, column=1, value='МАТРИЦА ОРЕОЛ TITAN: SKU × Магазин (продано, шт.)').font = TITLE_FONT
ws5.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(sorted_shops) + 2)
ws5.cell(row=2, column=1, value='Пустая ячейка = нет продаж | 0 = есть остаток, но нет продаж').font = SUBTITLE_FONT
ws5.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(sorted_shops) + 2)
ws5.row_dimensions[1].height = 22
ws5.row_dimensions[2].height = 18

# Заголовок: SKU | ABC | shop1 | shop2 | ... | Итого
headers5 = ['Код SKU', 'Наименование'] + [s[0][:20] for s in sorted_shops] + ['Итого']
for ci, h in enumerate(headers5, 1):
    c = ws5.cell(row=4, column=ci, value=h)
    c.font = WHITE_FONT; c.fill = DARK_FILL
    c.alignment = CENTER; c.border = THIN_BORDER
ws5.row_dimensions[4].height = 60

for ri, sku in enumerate(all_skus, 1):
    row_data = [sku['code'], sku['name'][:55]]
    row_total = 0
    for shop_name, _ in sorted_shops:
        sales = sku.get('shops', {}).get(shop_name, {}).get('sales', 0)
        row_data.append(int(sales) if sales > 0 else '')
        row_total += sales
    row_data.append(int(row_total))
    for ci, v in enumerate(row_data, 1):
        c = ws5.cell(row=4 + ri, column=ci, value=v)
        c.font = NORMAL_FONT; c.border = THIN_BORDER
        c.alignment = CENTER if ci != 2 else LEFT
        if ci == 2:
            c.alignment = LEFT
        elif ci == len(headers5):  # Итого
            if row_total >= 50:
                c.font = GREEN_FONT
            elif row_total < 10:
                c.font = RED_FONT
        elif isinstance(v, int) and v > 0:
            if v >= 10:
                c.font = GREEN_FONT; c.fill = PatternFill('solid', fgColor='E8F5E9')
            elif v < 3:
                c.font = RED_FONT

# Итог по магазинам
row_idx5 = 4 + len(all_skus) + 1
row_data5 = ['', 'ИТОГО']
for shop_name, d in sorted_shops:
    row_data5.append(int(d['sales']) if d['sales'] > 0 else '')
row_data5.append(int(tot_sold))
for ci, v in enumerate(row_data5, 1):
    c = ws5.cell(row=row_idx5, column=ci, value=v)
    c.font = Font(name='Arial', size=10, bold=True, color='FFFFFF')
    c.fill = ORANGE_FILL; c.border = THIN_BORDER
    c.alignment = CENTER if ci != 2 else LEFT

# Ширины
ws5.column_dimensions['A'].width = 10
ws5.column_dimensions['B'].width = 50
for ci in range(3, len(headers5) + 1):
    ws5.column_dimensions[chr(64 + ci) if ci <= 26 else 'A' + chr(64 + ci - 26)].width = 10

ws5.freeze_panes = 'C5'

# Сохраняем
out_path = '/home/z/my-project/download/TITAN_ВашДом_анализ_статистики.xlsx'
wb.save(out_path)
print(f'OK saved: {out_path}')
print(f'\n=== ИТОГИ (только ОРЕОЛ TITAN) ===')
print(f'Всего SKU: {tot_sku} (MML = {MML}, соблюдение {tot_sku/MML*100:.0f}%)')
print(f'Всего поступило: {int(tot_qty)} шт.')
print(f'Всего остаток: {int(tot_stock)} шт.')
print(f'Всего продано: {int(tot_sold)} шт.')
print(f'Активных магазинов: {len(shop_totals)} из 24')
print(f'\nТоп-3 цвета RAL:')
for color in existing_colors[:3]:
    d = stats['categories'][color]
    print(f'  - {color}: {d["sku_count"]} SKU, продано {int(d["total_sales"])} шт.')
print(f'\nТоп-3 магазина:')
for shop, d in sorted_shops[:3]:
    print(f'  - {shop}: продано {int(d["sales"])} шт.')
