"""
КОРРЕКТНАЯ финальная сборка v3:
1. Из исходного файла взять оригинальные 300 SKU с их idx и именами
2. Применить изменения: удалить 5, заменить 3, добавить 5 новых
3. Перенумеровать ОСТАВШИЕСЯ с 1 по 295
4. Добавить 5 новых с idx 296-300
5. Создать маппинг новый_idx → цена (через имя)
6. Обновить матрицы и журнал изменений
"""
import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from collections import defaultdict, Counter
import json
import os
import shutil

SRC_PATH = "/home/z/my-project/upload/Ассортимент_Олимп_300_SKU.xlsx"
OUT_PATH = "/home/z/my-project/download/Ассортимент_Олимп_300_SKU_финал_v3.xlsx"

# Начать с оригинала
shutil.copy(SRC_PATH, OUT_PATH)
wb = load_workbook(OUT_PATH)

# === Стили ===
HEAD_FILL = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
HEAD_FONT = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
HEAD_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)
TITLE_FONT = Font(name="Calibri", size=14, bold=True, color="1F4E79")
SUBTITLE_FONT = Font(name="Calibri", size=10, italic=True, color="595959")
BODY_FONT = Font(name="Calibri", size=10)
BODY_ALIGN = Alignment(horizontal="left", vertical="center", wrap_text=True)
NUM_ALIGN = Alignment(horizontal="right", vertical="center")
CENTER_ALIGN = Alignment(horizontal="center", vertical="center")
PRICE_FONT = Font(name="Calibri", size=10, bold=True, color="1F4E79")
NEW_ROW_FILL = PatternFill(start_color="C6E0B4", end_color="C6E0B4", fill_type="solid")
REPLACED_FILL = PatternFill(start_color="FFE699", end_color="FFE699", fill_type="solid")
DELETED_FILL = PatternFill(start_color="F4B084", end_color="F4B084", fill_type="solid")
THIN = Side(border_style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

# === Изменения ===
DELETE_IDX = [38, 56, 61, 62, 63]

REPLACEMENTS = {
    37: {
        'new_name': 'Грунт-эмаль "РАСЦВЕТ" 3 в 1 красная RAL 3020 1.9 кг',
        'new_sphere': 'Металл', 'new_purpose': 'Грунт-эмаль по металлу',
        'new_base': 'Органическая', 'new_brand': 'Расцвет', 'new_type': 'Грунт-Эмаль',
        'new_grp_code': 65199, 'new_abc': 'А',
        'отгрузочная': 610.78, 'базовая': 763.48, 'фасовка': 1.9, 'лист': 'Расцвет',
        'match_type': 'replacement',
        'ценовое_имя': 'Грунт-эмаль "РАСЦВЕТ" 3 в 1 красная RAL 3020',
        'status': 'ЗАМЕНА: вишневая → красная RAL 3020 (та же фасовка 1.9 кг, ближайший цвет)',
    },
    150: {
        'new_name': 'Грунт-концентрат влагоизолирующий водно-дисперсионный для наружных и внутренних работ "ОРЕОЛ" 1 кг',
        'new_sphere': 'Помещения (интерьер)', 'new_purpose': 'Прочее',
        'new_base': 'Водно-дисперсионная', 'new_brand': 'Ореол', 'new_type': 'Грунтовки',
        'new_grp_code': 66403, 'new_abc': 'С',
        'отгрузочная': 315.58, 'базовая': 394.47, 'фасовка': 1, 'лист': 'Ореол',
        'match_type': 'replacement',
        'ценовое_имя': 'Грунт-концентрат влагоизолирующий водно-дисперсионный для наружных и внутренних работ "ОРЕОЛ"',
        'status': 'ЗАМЕНА: АКВАСТОП → влагоизолирующий грунт-концентрат ОРЕОЛ (та же фасовка 1 кг)',
    },
    169: {
        'new_name': 'Краска для интерьеров моющаяся глубокоматовая «NEWTONE» водно-дисперсионная полиакриловая для внутренних работ черная RAL 9005 10 л 13.9 кг',
        'new_sphere': 'Помещения (интерьер)', 'new_purpose': 'Стены и потолки (интерьер)',
        'new_base': 'Водно-дисперсионная', 'new_brand': 'NEWTONE', 'new_type': 'Краски',
        'new_grp_code': 89508, 'new_abc': 'А',
        'отгрузочная': 2436.34, 'базовая': 3045.43, 'фасовка': 13.9, 'лист': 'NEWTONE',
        'match_type': 'replacement',
        'ценовое_имя': 'Краска для интерьеров моющаяся глубокоматовая «NEWTONE» водно-дисперсионная полиакриловая для внутренних работ черная RAL 9005',
        'status': 'ЗАМЕНА: ABSOLUTE BLACK → NEWTONE черная RAL 9005 (та же фасовка 13.9 кг)',
    },
}

NEW_ITEMS = [
    {'idx': 296, 'sphere': 'Дерево', 'purpose': 'Защита и декор дерева', 'base': 'Водно-дисперсионная',
     'brand': 'NoName', 'type': 'Краски', 'grp_code': 63180, 'abc': 'А',
     'name': 'Краска для садовых деревьев водно-дисперсионная полиакриловая 3 кг',
     'отгрузочная': 221.27, 'базовая': 276.59, 'фасовка': 3, 'лист': 'NoName',
     'match_type': 'new_from_abc',
     'ценовое_имя': 'Краска для садовых деревьев водно-дисперсионная полиакриловая',
     'status': 'НОВАЯ: топ-1 ABC-анализа (cumulative=0.641)'},
    {'idx': 297, 'sphere': 'Металл', 'purpose': 'Эмаль по металлу', 'base': 'Органическая',
     'brand': 'Простокрашено! БАУ', 'type': 'Эмали Алкидные', 'grp_code': 79754, 'abc': 'А',
     'name': 'Эмаль ПФ-115 "ПРОСТОКРАШЕНО!" синяя БАУ 0.8 кг',
     'отгрузочная': 179.02, 'базовая': 223.78, 'фасовка': 0.8, 'лист': 'Простокрашено! БАУ',
     'match_type': 'new_from_abc',
     'ценовое_имя': 'Эмаль ПФ-115 "ПРОСТОКРАШЕНО!" синяя БАУ',
     'status': 'НОВАЯ: топ-2 ABC-анализа. Альтернатива Простокрашено! БАУ 0.8 кг'},
    {'idx': 298, 'sphere': 'Фасад', 'purpose': 'Фасад', 'base': 'Водно-дисперсионная',
     'brand': 'СТРОЙКОМПЛЕКТ', 'type': 'Краски', 'grp_code': 73122, 'abc': 'А',
     'name': 'Краска фасадная атмосферостойкая водно-дисперсионная полиакриловая СТРОЙКОМПЛЕКТ 13 кг',
     'отгрузочная': 1503.92, 'базовая': 1879.9, 'фасовка': 13, 'лист': 'СТРОЙКОМПЛЕКТ',
     'match_type': 'new_from_abc',
     'ценовое_имя': 'Краска фасадная атмосферостойкая водно-дисперсионная полиакриловая СТРОЙКОМПЛЕКТ',
     'status': 'НОВАЯ: топ-3 ABC-анализа. Новый бренд: СТРОЙКОМПЛЕКТ'},
    {'idx': 299, 'sphere': 'Металл', 'purpose': 'Эмаль по металлу', 'base': 'Органическая',
     'brand': 'EMPILS ЭКОНОМ', 'type': 'Эмали Алкидные', 'grp_code': 7803, 'abc': 'А',
     'name': 'Эмаль EMPILS ПФ-115 «ЭКОНОМ» чёрная 20 кг',
     'отгрузочная': 2578.87, 'базовая': 3223.59, 'фасовка': 20, 'лист': 'EMPILS ЭКОНОМ',
     'match_type': 'new_from_abc',
     'ценовое_имя': 'Эмаль EMPILS ПФ-115 «ЭКОНОМ» чёрная',
     'status': 'НОВАЯ: топ-4 ABC-анализа. Новый бренд: EMPILS ЭКОНОМ'},
    {'idx': 300, 'sphere': 'Металл', 'purpose': 'Эмаль по металлу', 'base': 'Органическая',
     'brand': 'СТРОЙКОМПЛЕКТ', 'type': 'Эмали Алкидные', 'grp_code': 72676, 'abc': 'А',
     'name': 'Эмаль ПФ-115 СТРОЙКОМПЛЕКТ чёрная 10 кг',
     'отгрузочная': 1640.57, 'базовая': 2050.71, 'фасовка': 10, 'лист': 'СТРОЙКОМПЛЕКТ',
     'match_type': 'new_from_abc',
     'ценовое_имя': 'Эмаль ПФ-115 СТРОЙКОМПЛЕКТ чёрная',
     'status': 'НОВАЯ: топ-5 ABC-анализа. Бренд СТРОЙКОМПЛЕКТ (2-я позиция)'},
]

# === 1. Загрузить оригинальные SKU и их цены (по имени) ===
print("=== Загрузка оригинальных SKU и цен ===")
wb_orig = load_workbook(SRC_PATH, data_only=True)
ws_orig = wb_orig["300 SKU — детально"]

# Словарь: orig_idx → item_data
orig_items = {}
for r in range(4, ws_orig.max_row + 1):
    idx = ws_orig.cell(row=r, column=1).value
    if idx is None:
        continue
    orig_items[idx] = {
        'sphere': ws_orig.cell(row=r, column=2).value,
        'purpose': ws_orig.cell(row=r, column=3).value,
        'base': ws_orig.cell(row=r, column=4).value,
        'brand': ws_orig.cell(row=r, column=5).value,
        'type': ws_orig.cell(row=r, column=6).value,
        'name': ws_orig.cell(row=r, column=7).value,
        'grp_code': ws_orig.cell(row=r, column=8).value,
        'abc': ws_orig.cell(row=r, column=9).value,
    }
print(f"  Оригинальных SKU: {len(orig_items)}")

# Загрузить предыдущие результаты (цены по idx)
with open('/home/z/my-project/scripts/results.json', 'r', encoding='utf-8') as f:
    prev_results = json.load(f)
prev_prices_by_idx = {r['idx']: r for r in prev_results}

# === 2. Создать список финальных SKU ===
final_items = []

# 2a. Оригинальные SKU (кроме удалённых), с заменами для #37, #150, #169
for idx in sorted(orig_items.keys()):
    if idx in DELETE_IDX:
        continue  # пропустить удалённые
    
    orig = orig_items[idx]
    
    if idx in REPLACEMENTS:
        # Замена
        repl = REPLACEMENTS[idx]
        item = {
            'orig_idx': idx,
            'sphere': repl['new_sphere'],
            'purpose': repl['new_purpose'],
            'base': repl['new_base'],
            'brand': repl['new_brand'],
            'type': repl['new_type'],
            'name': repl['new_name'],
            'grp_code': repl['new_grp_code'],
            'abc': repl['new_abc'],
            'price_source': 'replacement',
            'orig_name_for_price': None,
        }
    else:
        # Оригинал
        item = {
            'orig_idx': idx,
            'sphere': orig['sphere'],
            'purpose': orig['purpose'],
            'base': orig['base'],
            'brand': orig['brand'],
            'type': orig['type'],
            'name': orig['name'],
            'grp_code': orig['grp_code'],
            'abc': orig['abc'],
            'price_source': 'original',
            'orig_name_for_price': orig['name'],
        }
    final_items.append(item)

print(f"  После удаления 5 и без добавления новых: {len(final_items)} SKU")

# 2b. Добавить 5 новых SKU
for ni in NEW_ITEMS:
    final_items.append({
        'orig_idx': None,  # новый
        'sphere': ni['sphere'],
        'purpose': ni['purpose'],
        'base': ni['base'],
        'brand': ni['brand'],
        'type': ni['type'],
        'name': ni['name'],
        'grp_code': ni['grp_code'],
        'abc': ni['abc'],
        'price_source': 'new',
        'orig_name_for_price': None,
    })

print(f"  После добавления 5 новых: {len(final_items)} SKU")

# 2c. Перенумеровать с 1 по 300
for i, item in enumerate(final_items, 1):
    item['new_idx'] = i
print(f"  Перенумеровано: 1-{len(final_items)}")

# === 3. Обновить лист "300 SKU — детально" ===
print("\n=== Обновление '300 SKU — детально' ===")
ws_det = wb["300 SKU — детально"]

# Очистить старые данные
max_row = ws_det.max_row
for r in range(4, max_row + 1):
    for c in range(1, 10):
        ws_det.cell(row=r, column=c, value=None)

# Записать новые данные
for i, item in enumerate(final_items):
    r = 4 + i
    ws_det.cell(row=r, column=1, value=item['new_idx'])
    ws_det.cell(row=r, column=2, value=item['sphere'])
    ws_det.cell(row=r, column=3, value=item['purpose'])
    ws_det.cell(row=r, column=4, value=item['base'])
    ws_det.cell(row=r, column=5, value=item['brand'])
    ws_det.cell(row=r, column=6, value=item['type'])
    ws_det.cell(row=r, column=7, value=item['name'])
    ws_det.cell(row=r, column=8, value=item['grp_code'])
    ws_det.cell(row=r, column=9, value=item['abc'])

print(f"  ✓ Записано {len(final_items)} строк")

# === 4. Пересоздать лист "300 SKU с ценами" ===
print("\n=== Пересоздание '300 SKU с ценами' ===")
if "300 SKU с ценами" in wb.sheetnames:
    del wb["300 SKU с ценами"]
ws_prices = wb.create_sheet("300 SKU с ценами", 2)

# Заголовок
ws_prices.merge_cells("A1:N1")
ws_prices["A1"] = "Стартовый ассортимент Олимп — 300 SKU с отгрузочными ценами (ФИНАЛ v3)"
ws_prices["A1"].font = TITLE_FONT
ws_prices["A1"].alignment = Alignment(horizontal="left", vertical="center")
ws_prices.row_dimensions[1].height = 28

ws_prices.merge_cells("A2:N2")
ws_prices["A2"] = ("Источник цен: прайс-лист «16.10.2026 РОССИЯ с НДС, руб.» | "
                   "Удалено: 5 SKU РАСЦВЕТ без альтернатив | Заменено на альтернативы: 3 SKU | "
                   "Добавлено из ABC-анализа: 5 SKU | ИТОГО: 300 SKU, все с ценами")
ws_prices["A2"].font = SUBTITLE_FONT
ws_prices["A2"].alignment = Alignment(horizontal="left", vertical="center")
ws_prices.row_dimensions[2].height = 18

HEADERS = [
    ("#", 5), ("Сфера", 14), ("Назначение", 28), ("Основа", 16),
    ("Бренд", 16), ("Тип продукции", 22), ("Наименование", 60),
    ("Гр. код", 9), ("ABC", 6),
    ("Отгрузочная цена, руб", 14), ("Базовая цена, руб", 13),
    ("Фасовка в прайсе, кг", 11), ("Лист прайса", 16), ("Статус сопоставления", 32),
]
for col_idx, (title, width) in enumerate(HEADERS, 1):
    cell = ws_prices.cell(row=4, column=col_idx, value=title)
    cell.fill = HEAD_FILL
    cell.font = HEAD_FONT
    cell.alignment = HEAD_ALIGN
    cell.border = BORDER
    ws_prices.column_dimensions[get_column_letter(col_idx)].width = width
ws_prices.row_dimensions[4].height = 38

# Заполнить строки с ценами
row_i = 5
stats = {'matched_original': 0, 'replaced': 0, 'new': 0, 'no_price': 0}

for item in final_items:
    # Ищем цену
    price_info = None
    is_replaced = False
    is_new = False
    
    if item['price_source'] == 'replacement':
        # Для заменённых — берём цену из REPLACEMENTS (по orig_idx)
        repl = REPLACEMENTS[item['orig_idx']]
        price_info = {
            'отгрузочная': repl['отгрузочная'],
            'базовая': repl['базовая'],
            'фасовка': repl['фасовка'],
            'лист': repl['лист'],
            'status': repl['status'],
        }
        stats['replaced'] += 1
        is_replaced = True
    elif item['price_source'] == 'new':
        # Для новых — берём цену из NEW_ITEMS (по grp_code)
        for ni in NEW_ITEMS:
            if ni['grp_code'] == item['grp_code']:
                price_info = {
                    'отгрузочная': ni['отгрузочная'],
                    'базовая': ni['базовая'],
                    'фасовка': ni['фасовка'],
                    'лист': ni['лист'],
                    'status': ni['status'],
                }
                break
        stats['new'] += 1
        is_new = True
    else:
        # Для оригинальных — берём цену из prev_prices_by_idx (по orig_idx)
        prev = prev_prices_by_idx.get(item['orig_idx'])
        if prev and prev.get('отгрузочная') is not None:
            price_info = {
                'отгрузочная': prev['отгрузочная'],
                'базовая': prev['базовая'],
                'фасовка': prev.get('фасовка_прайс'),
                'лист': prev.get('лист', ''),
                'status': f"Цена: {prev.get('match_type', 'matched')}",
            }
            stats['matched_original'] += 1
        else:
            stats['no_price'] += 1
    
    vals = [
        item['new_idx'], item['sphere'], item['purpose'], item['base'],
        item['brand'], item['type'], item['name'],
        item['grp_code'], item['abc'],
        price_info['отгрузочная'] if price_info else None,
        price_info['базовая'] if price_info else None,
        price_info['фасовка'] if price_info else None,
        price_info['лист'] if price_info else '',
        price_info['status'] if price_info else 'не найдено в прайс-листе',
    ]
    
    fill = NEW_ROW_FILL if is_new else (REPLACED_FILL if is_replaced else None)
    
    for col_idx, val in enumerate(vals, 1):
        cell = ws_prices.cell(row=row_i, column=col_idx, value=val)
        cell.border = BORDER
        cell.font = BODY_FONT
        if col_idx == 1:
            cell.alignment = CENTER_ALIGN
        elif col_idx in (10, 11, 12):
            cell.alignment = NUM_ALIGN
            if col_idx == 10 and val is not None:
                cell.font = PRICE_FONT
                cell.number_format = '# ##0.00'
            elif col_idx == 11 and val is not None:
                cell.number_format = '# ##0.00'
            elif col_idx == 12 and val is not None:
                cell.number_format = '0.##'
        elif col_idx == 7:
            cell.alignment = BODY_ALIGN
        elif col_idx == 14:
            cell.alignment = BODY_ALIGN
            if is_new:
                cell.font = Font(name="Calibri", size=9, italic=True, color="375623")
            elif is_replaced:
                cell.font = Font(name="Calibri", size=9, italic=True, color="7F6000")
        else:
            cell.alignment = BODY_ALIGN
        
        if fill:
            cell.fill = fill
    
    if is_new:
        ws_prices.row_dimensions[row_i].height = 40
    elif is_replaced:
        ws_prices.row_dimensions[row_i].height = 36
    
    row_i += 1

ws_prices.freeze_panes = "B5"
last_col = get_column_letter(14)
ws_prices.auto_filter.ref = f"A4:{last_col}{row_i - 1}"

print(f"  ✓ Заполнено {row_i - 5} строк с ценами")
print(f"  Статистика: {stats}")

# === 5. Пересчёт матриц на листе "Сводка" ===
print("\n=== Пересчёт матриц 'Сводка' ===")
ws_sv = wb["Сводка"]

spheres_order = ['Металл', 'Помещения (интерьер)', 'Дерево', 'Пол', 'Фасад']
bases_order = ['Органическая', 'Водно-дисперсионная']

# Матрица 1: Основа × Сфера
matrix_base_sphere = {(b, s): 0 for b in bases_order for s in spheres_order}
for it in final_items:
    key = (it['base'], it['sphere'])
    if key in matrix_base_sphere:
        matrix_base_sphere[key] += 1

for row_offset, base in enumerate(bases_order):
    r = 5 + row_offset
    ws_sv.cell(row=r, column=1, value=base)
    for col_offset, sphere in enumerate(spheres_order):
        ws_sv.cell(row=r, column=2 + col_offset, value=matrix_base_sphere[(base, sphere)])
    total_row = sum(matrix_base_sphere[(base, s)] for s in spheres_order)
    ws_sv.cell(row=r, column=7, value=total_row)
    pct = total_row / len(final_items)
    ws_sv.cell(row=r, column=8, value=f"{pct*100:.1f}%")

r = 7
ws_sv.cell(row=r, column=1, value='ИТОГО')
for col_offset, sphere in enumerate(spheres_order):
    col_total = sum(matrix_base_sphere[(b, sphere)] for b in bases_order)
    ws_sv.cell(row=r, column=2 + col_offset, value=col_total)
ws_sv.cell(row=r, column=7, value=len(final_items))
ws_sv.cell(row=r, column=8, value='100%')

# Матрица 2: СФЕРА × БРЕНД
brand_counts = Counter(it['brand'] for it in final_items)
original_brands = ['ОРЕОЛ TITAN', 'Ореол', 'Ореол Дисконт', 'NEWTONE', 'AvantGarde',
                    'Расцвет', 'Простокрашено! БАУ', 'PROFIWOOD']
new_brands = [b for b in brand_counts if b not in original_brands]
all_brands = original_brands + new_brands

for r in range(10, 17):
    for c in range(1, 15):
        ws_sv.cell(row=r, column=c, value=None)

ws_sv.cell(row=10, column=1, value='СФЕРА / БРЕНД')
for col_offset, brand in enumerate(all_brands):
    ws_sv.cell(row=10, column=2 + col_offset, value=brand)
ws_sv.cell(row=10, column=2 + len(all_brands), value='ИТОГО')

for c in range(1, 3 + len(all_brands)):
    cell = ws_sv.cell(row=10, column=c)
    cell.fill = HEAD_FILL
    cell.font = HEAD_FONT
    cell.alignment = HEAD_ALIGN
    cell.border = BORDER

matrix_sphere_brand = {(s, b): 0 for s in spheres_order for b in all_brands}
for it in final_items:
    if it['brand'] in all_brands:
        matrix_sphere_brand[(it['sphere'], it['brand'])] += 1

for row_offset, sphere in enumerate(spheres_order):
    r = 11 + row_offset
    ws_sv.cell(row=r, column=1, value=sphere)
    for col_offset, brand in enumerate(all_brands):
        val = matrix_sphere_brand[(sphere, brand)]
        if val > 0:
            ws_sv.cell(row=r, column=2 + col_offset, value=val)
    total_row = sum(matrix_sphere_brand[(sphere, b)] for b in all_brands)
    ws_sv.cell(row=r, column=2 + len(all_brands), value=total_row)
    
    for c in range(1, 3 + len(all_brands)):
        cell = ws_sv.cell(row=r, column=c)
        cell.font = BODY_FONT
        cell.border = BORDER
        if c > 1:
            cell.alignment = CENTER_ALIGN
        else:
            cell.alignment = Alignment(horizontal="left", vertical="center")

r = 16
ws_sv.cell(row=r, column=1, value='ИТОГО')
for col_offset, brand in enumerate(all_brands):
    ws_sv.cell(row=r, column=2 + col_offset, value=brand_counts[brand])
ws_sv.cell(row=r, column=2 + len(all_brands), value=len(final_items))

for c in range(1, 3 + len(all_brands)):
    cell = ws_sv.cell(row=r, column=c)
    cell.font = Font(name="Calibri", size=10, bold=True)
    cell.fill = PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid")
    cell.border = BORDER
    if c > 1:
        cell.alignment = CENTER_ALIGN
    else:
        cell.alignment = Alignment(horizontal="left", vertical="center")

for col_offset in range(len(all_brands) + 1):
    ws_sv.column_dimensions[get_column_letter(2 + col_offset)].width = 13
ws_sv.column_dimensions[get_column_letter(2 + len(all_brands))].width = 8
ws_sv.column_dimensions['A'].width = 24

# Матрица 3: ДЕТАЛЬНАЯ КЛАССИФИКАЦИЯ ПО НАЗНАЧЕНИЮ
purposes_order = [
    'Грунт-эмаль по металлу', 'Эмаль по металлу', 'Пол',
    'Стены и потолки (интерьер)', 'Защита и декор дерева',
    'Эмали ВД прочие', 'Лаки прочие', 'Грунтовки', 'Фасад',
    'Прочее', 'Яхтный лак', 'Шпатлёвки', 'Паркетный лак',
    'Молотковая эмаль', 'Потолки', 'Резиновая краска',
]

matrix_purpose_sphere = {(p, s): 0 for p in purposes_order for s in spheres_order}
purpose_counts = Counter()
for it in final_items:
    key = (it['purpose'], it['sphere'])
    if key in matrix_purpose_sphere:
        matrix_purpose_sphere[key] += 1
        purpose_counts[it['purpose']] += 1

purposes_sorted = sorted(purposes_order, key=lambda p: -purpose_counts.get(p, 0))

for r in range(20, 36):
    for c in range(1, 8):
        ws_sv.cell(row=r, column=c, value=None)

ws_sv.cell(row=19, column=1, value='Сфера / Назначение')
sphere_cols = ['Металл', 'Помещения (интерьер)', 'Дерево', 'Пол', 'Фасад']
for col_offset, sp in enumerate(sphere_cols):
    ws_sv.cell(row=19, column=2 + col_offset, value=sp)
ws_sv.cell(row=19, column=7, value='ИТОГО')

for c in range(1, 8):
    cell = ws_sv.cell(row=19, column=c)
    cell.fill = HEAD_FILL
    cell.font = HEAD_FONT
    cell.alignment = HEAD_ALIGN
    cell.border = BORDER

for row_offset, purpose in enumerate(purposes_sorted):
    r = 20 + row_offset
    ws_sv.cell(row=r, column=1, value=purpose)
    for col_offset, sp in enumerate(sphere_cols):
        val = matrix_purpose_sphere[(purpose, sp)]
        if val > 0:
            ws_sv.cell(row=r, column=2 + col_offset, value=val)
    total_row = purpose_counts.get(purpose, 0)
    if total_row > 0:
        ws_sv.cell(row=r, column=7, value=total_row)
    
    for c in range(1, 8):
        cell = ws_sv.cell(row=r, column=c)
        cell.font = BODY_FONT
        cell.border = BORDER
        if c > 1:
            cell.alignment = CENTER_ALIGN
        else:
            cell.alignment = Alignment(horizontal="left", vertical="center")

print(f"  ✓ Матрицы обновлены")

# === 6. Лист "Изменения v3 (10 SKU)" ===
if "Изменения v3 (10 SKU)" in wb.sheetnames:
    del wb["Изменения v3 (10 SKU)"]
if "8 SKU — альтернативы" in wb.sheetnames:
    del wb["8 SKU — альтернативы"]
if "8 несопоставленных — альтернативы" in wb.sheetnames:
    del wb["8 несопоставленных — альтернативы"]

ws_log = wb.create_sheet("Изменения v3 (10 SKU)")

ws_log.merge_cells("A1:G1")
ws_log["A1"] = "Журнал изменений v3: удалено 5 SKU, заменено 3 SKU, добавлено 5 SKU"
ws_log["A1"].font = TITLE_FONT
ws_log.row_dimensions[1].height = 28

ws_log.merge_cells("A2:G2")
ws_log["A2"] = ("В соответствии с заданием: 5 SKU РАСЦВЕТ 3 в 1 (без альтернатив в прайс-листе) удалены и заменены "
                "5 новыми популярными SKU из ABC-анализа (топ-5 группы А). 3 SKU (с найденными альтернативами в прайс-листе) "
                "заменены на эти альтернативы с проставлением цен. После изменений выполнена перенумерация всех SKU 1-300.")
ws_log["A2"].font = SUBTITLE_FONT
ws_log["A2"].alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
ws_log.row_dimensions[2].height = 50

log_headers = [
    ("Действие", 14), ("# SKU (исх.)", 12), ("Бренд", 18), ("Сфера", 16),
    ("Наименование (новое)", 55), ("Отгруз. цена, руб", 14), ("Комментарий", 75),
]
for col_idx, (title, width) in enumerate(log_headers, 1):
    cell = ws_log.cell(row=4, column=col_idx, value=title)
    cell.fill = HEAD_FILL
    cell.font = HEAD_FONT
    cell.alignment = HEAD_ALIGN
    cell.border = BORDER
    ws_log.column_dimensions[get_column_letter(col_idx)].width = width
ws_log.row_dimensions[4].height = 38

row_i = 5
ws_log.cell(row=row_i, column=1, value="УДАЛЕНО (5 SKU РАСЦВЕТ 3 в 1 без альтернатив в прайс-листе)")
ws_log.cell(row=row_i, column=1).font = Font(name="Calibri", size=11, bold=True, color="C00000")
for c in range(1, 8):
    ws_log.cell(row=row_i, column=c).fill = DELETED_FILL
ws_log.merge_cells(f"A{row_i}:G{row_i}")
row_i += 1

deleted_items_info = [
    (38, 'Расцвет', 'Металл', 'Грунт-эмаль "РАСЦВЕТ" 3 в 1 голубая 0.9 кг',
     'Цвет ГОЛУБАЯ для 3 в 1 отсутствует в прайс-листе. Доступные цвета 3 в 1: красная RAL 3020, серый графит RAL 7024.'),
    (56, 'Расцвет', 'Металл', 'Грунт-эмаль "РАСЦВЕТ" 3 в 1 тёмно-зелёная 0.9 кг',
     'Цвет ТЁМНО-ЗЕЛЁНАЯ для 3 в 1 отсутствует в прайс-листе.'),
    (61, 'Расцвет', 'Металл', 'Грунт-эмаль "РАСЦВЕТ" 3 в 1 шоколадная 0.9 кг',
     'Цвет ШОКОЛАДНАЯ для 3 в 1 отсутствует в прайс-листе.'),
    (62, 'Расцвет', 'Металл', 'Грунт-эмаль "РАСЦВЕТ" 3 в 1 шоколадная 1.9 кг',
     'Цвет ШОКОЛАДНАЯ для 3 в 1 отсутствует в прайс-листе.'),
    (63, 'Расцвет', 'Металл', 'Грунт-эмаль "РАСЦВЕТ" 3 в 1 шоколадная 2.7 кг',
     'Цвет ШОКОЛАДНАЯ для 3 в 1 отсутствует в прайс-листе.'),
]
for idx, brand, sphere, name, comment in deleted_items_info:
    vals = ["Удалено", idx, brand, sphere, name, None, comment]
    for col_idx, val in enumerate(vals, 1):
        cell = ws_log.cell(row=row_i, column=col_idx, value=val)
        cell.border = BORDER
        cell.font = BODY_FONT
        cell.fill = DELETED_FILL
        if col_idx == 1:
            cell.alignment = CENTER_ALIGN
            cell.font = Font(name="Calibri", size=10, bold=True, color="C00000")
        elif col_idx == 2:
            cell.alignment = CENTER_ALIGN
        elif col_idx == 6:
            cell.alignment = NUM_ALIGN
        else:
            cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    ws_log.row_dimensions[row_i].height = 36
    row_i += 1

row_i += 1
ws_log.cell(row=row_i, column=1, value="ЗАМЕНЕНО (3 SKU — на найденные альтернативы в прайс-листе)")
ws_log.cell(row=row_i, column=1).font = Font(name="Calibri", size=11, bold=True, color="7F6000")
for c in range(1, 8):
    ws_log.cell(row=row_i, column=c).fill = REPLACED_FILL
ws_log.merge_cells(f"A{row_i}:G{row_i}")
row_i += 1

for idx, repl in REPLACEMENTS.items():
    vals = [
        "Замена", idx, repl['new_brand'], repl['new_sphere'],
        repl['new_name'], repl['отгрузочная'], repl['status'],
    ]
    for col_idx, val in enumerate(vals, 1):
        cell = ws_log.cell(row=row_i, column=col_idx, value=val)
        cell.border = BORDER
        cell.font = BODY_FONT
        cell.fill = REPLACED_FILL
        if col_idx == 1:
            cell.alignment = CENTER_ALIGN
            cell.font = Font(name="Calibri", size=10, bold=True, color="7F6000")
        elif col_idx == 2:
            cell.alignment = CENTER_ALIGN
        elif col_idx == 6:
            cell.alignment = NUM_ALIGN
            if val is not None:
                cell.number_format = '# ##0.00'
                cell.font = Font(name="Calibri", size=10, bold=True, color="7F6000")
        else:
            cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    ws_log.row_dimensions[row_i].height = 48
    row_i += 1

row_i += 1
ws_log.cell(row=row_i, column=1, value="ДОБАВЛЕНО (5 SKU из топ-5 ABC-анализа группы А, нумерация 296-300)")
ws_log.cell(row=row_i, column=1).font = Font(name="Calibri", size=11, bold=True, color="375623")
for c in range(1, 8):
    ws_log.cell(row=row_i, column=c).fill = NEW_ROW_FILL
ws_log.merge_cells(f"A{row_i}:G{row_i}")
row_i += 1

for item in NEW_ITEMS:
    vals = [
        "Новая", item['idx'], item['brand'], item['sphere'],
        item['name'], item['отгрузочная'], item['status'],
    ]
    for col_idx, val in enumerate(vals, 1):
        cell = ws_log.cell(row=row_i, column=col_idx, value=val)
        cell.border = BORDER
        cell.font = BODY_FONT
        cell.fill = NEW_ROW_FILL
        if col_idx == 1:
            cell.alignment = CENTER_ALIGN
            cell.font = Font(name="Calibri", size=10, bold=True, color="375623")
        elif col_idx == 2:
            cell.alignment = CENTER_ALIGN
        elif col_idx == 6:
            cell.alignment = NUM_ALIGN
            if val is not None:
                cell.number_format = '# ##0.00'
                cell.font = Font(name="Calibri", size=10, bold=True, color="375623")
        else:
            cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    ws_log.row_dimensions[row_i].height = 50
    row_i += 1

ws_log.freeze_panes = "B5"

# === Сохранить ===
wb.properties.creator = "Z.ai"
wb.properties.title = "Ассортимент Олимп 300 SKU — финал v3"
wb.properties.subject = "Удалено 5, заменено 3, добавлено 5 SKU из ABC-анализа"

print(f"\nЛисты в финальном файле: {wb.sheetnames}")
wb.save(OUT_PATH)
print(f"\n✓ Файл сохранён: {OUT_PATH}")
print(f"  Размер: {os.path.getsize(OUT_PATH):,} байт")
