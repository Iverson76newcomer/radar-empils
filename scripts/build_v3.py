"""
Финальная сборка файла v3:
1. Удалить 5 SKU РАСЦВЕТ 3 в 1 (без альтернатив)
2. Заменить 3 SKU (#37, #150, #169) на найденные альтернативы с ценами
3. Добавить 5 новых SKU из ABC-анализа с ценами
4. Пересчитать матрицы: Основа × Сфера, Сфера × Бренд, Назначение
"""
import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import shutil
import os

SRC_PATH = "/home/z/my-project/upload/Ассортимент_Олимп_300_SKU.xlsx"
OUT_PATH = "/home/z/my-project/download/Ассортимент_Олимп_300_SKU_финал_v3.xlsx"

# Копируем оригинал
shutil.copy(SRC_PATH, OUT_PATH)
wb = load_workbook(OUT_PATH)
print(f"Листы в файле: {wb.sheetnames}")

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
NEW_ROW_FILL = PatternFill(start_color="C6E0B4", end_color="C6E0B4", fill_type="solid")  # зелёный для новых
REPLACED_FILL = PatternFill(start_color="FFE699", end_color="FFE699", fill_type="solid")  # жёлтый для заменённых
DELETED_FILL = PatternFill(start_color="F4B084", end_color="F4B084", fill_type="solid")  # оранжевый для удалённых

THIN = Side(border_style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

# === Данные для изменений ===
# 5 SKU для удаления (idx → row in '300 SKU — детально' sheet)
# Из предыдущей проверки: строки 41, 42, 60, 65, 66, 67 → но индексы: 38, 56, 61, 62, 63
DELETE_IDX = [38, 56, 61, 62, 63]

# 3 SKU для замены: (idx, новое_имя, новая_основа, отгруз, базовая, фасовка, лист, тип_совпадения, комментарий)
REPLACEMENTS = {
    37: {
        'new_name': 'Грунт-эмаль "РАСЦВЕТ" 3 в 1 красная RAL 3020 1.9 кг',
        'new_sphere': 'Металл',
        'new_purpose': 'Грунт-эмаль по металлу',
        'new_base': 'Органическая',
        'new_brand': 'Расцвет',
        'new_type': 'Грунт-Эмаль',
        'new_grp_code': 65199,
        'new_abc': 'А',
        'отгрузочная': 610.78,
        'базовая': 763.48,
        'фасовка': 1.9,
        'лист': 'Расцвет',
        'match_type': 'replacement',
        'ценовое_имя': 'Грунт-эмаль "РАСЦВЕТ" 3 в 1 красная RAL 3020',
        'comment': 'ЗАМЕНА: вишневая → красная RAL 3020 (та же фасовка 1.9 кг, ближайший цвет)',
    },
    150: {
        'new_name': 'Грунт-концентрат влагоизолирующий водно-дисперсионный для наружных и внутренних работ "ОРЕОЛ" 1 кг',
        'new_sphere': 'Помещения (интерьер)',
        'new_purpose': 'Прочее',
        'new_base': 'Водно-дисперсионная',  # было Органическая — ИЗМЕНИЛИ
        'new_brand': 'Ореол',
        'new_type': 'Грунтовки',
        'new_grp_code': 66403,
        'new_abc': 'С',
        'отгрузочная': 315.58,
        'базовая': 394.47,
        'фасовка': 1,
        'лист': 'Ореол',
        'match_type': 'replacement',
        'ценовое_имя': 'Грунт-концентрат влагоизолирующий водно-дисперсионный для наружных и внутренних работ "ОРЕОЛ"',
        'comment': 'ЗАМЕНА: АКВАСТОП (серия отсутствует в прайсе) → влагоизолирующий грунт-концентрат ОРЕОЛ (та же фасовка 1 кг, похожий функционал)',
    },
    169: {
        'new_name': 'Краска для интерьеров моющаяся глубокоматовая «NEWTONE» водно-дисперсионная полиакриловая для внутренних работ черная RAL 9005 10 л 13.9 кг',
        'new_sphere': 'Помещения (интерьер)',
        'new_purpose': 'Стены и потолки (интерьер)',
        'new_base': 'Водно-дисперсионная',
        'new_brand': 'NEWTONE',
        'new_type': 'Краски',
        'new_grp_code': 89508,
        'new_abc': 'А',
        'отгрузочная': 2436.34,
        'базовая': 3045.43,
        'фасовка': 13.9,
        'лист': 'NEWTONE',
        'match_type': 'replacement',
        'ценовое_имя': 'Краска для интерьеров моющаяся глубокоматовая «NEWTONE» водно-дисперсионная полиакриловая для внутренних работ черная RAL 9005',
        'comment': 'ЗАМЕНА: серия ABSOLUTE BLACK (нет в прайсе) → NEWTONE черная RAL 9005 (тот же цвет, та же фасовка 13.9 кг, та же характеристика моющаяся глубокоматовая)',
    },
}

# 5 новых SKU (idx 296-300)
NEW_ITEMS = [
    {
        'idx': 296, 'sphere': 'Дерево', 'purpose': 'Защита и декор дерева', 'base': 'Водно-дисперсионная',
        'brand': 'NoName', 'type': 'Краски', 'grp_code': 63180, 'abc': 'А',
        'name': 'Краска для садовых деревьев водно-дисперсионная полиакриловая 3 кг',
        'отгрузочная': 221.27, 'базовая': 276.59, 'фасовка': 3, 'лист': 'NoName',
        'match_type': 'new_from_abc',
        'ценовое_имя': 'Краска для садовых деревьев водно-дисперсионная полиакриловая',
        'abc_cumulative': 0.641,
        'comment': 'НОВАЯ ПОЗИЦИЯ из ABC-анализа (топ-1 по популярности в группе А, не входившие в ассортимент)',
    },
    {
        'idx': 297, 'sphere': 'Металл', 'purpose': 'Эмаль по металлу', 'base': 'Органическая',
        'brand': 'Простокрашено! БАУ', 'type': 'Эмали Алкидные', 'grp_code': 79754, 'abc': 'А',
        'name': 'Эмаль ПФ-115 "ПРОСТОКРАШЕНО!" синяя БАУ 0.8 кг',
        'отгрузочная': 179.02, 'базовая': 223.78, 'фасовка': 0.8, 'лист': 'Простокрашено! БАУ',
        'match_type': 'new_from_abc',
        'ценовое_имя': 'Эмаль ПФ-115 "ПРОСТОКРАШЕНО!" синяя БАУ',
        'abc_cumulative': 0.647,
        'comment': 'НОВАЯ ПОЗИЦИЯ из ABC-анализа (топ-2). В прайсе NoName нет — взята ближайшая по фасовке (0.8 кг) от Простокрашено! БАУ',
    },
    {
        'idx': 298, 'sphere': 'Фасад', 'purpose': 'Фасад', 'base': 'Водно-дисперсионная',
        'brand': 'СТРОЙКОМПЛЕКТ', 'type': 'Краски', 'grp_code': 73122, 'abc': 'А',
        'name': 'Краска фасадная атмосферостойкая водно-дисперсионная полиакриловая СТРОЙКОМПЛЕКТ 13 кг',
        'отгрузочная': 1503.92, 'базовая': 1879.9, 'фасовка': 13, 'лист': 'СТРОЙКОМПЛЕКТ',
        'match_type': 'new_from_abc',
        'ценовое_имя': 'Краска фасадная атмосферостойкая водно-дисперсионная полиакриловая СТРОЙКОМПЛЕКТ',
        'abc_cumulative': 0.653,
        'comment': 'НОВАЯ ПОЗИЦИЯ из ABC-анализа (топ-3). Новый бренд в матрице: СТРОЙКОМПЛЕКТ',
    },
    {
        'idx': 299, 'sphere': 'Металл', 'purpose': 'Эмаль по металлу', 'base': 'Органическая',
        'brand': 'EMPILS ЭКОНОМ', 'type': 'Эмали Алкидные', 'grp_code': 7803, 'abc': 'А',
        'name': 'Эмаль EMPILS ПФ-115 «ЭКОНОМ» чёрная 20 кг',
        'отгрузочная': 2578.87, 'базовая': 3223.59, 'фасовка': 20, 'лист': 'EMPILS ЭКОНОМ',
        'match_type': 'new_from_abc',
        'ценовое_имя': 'Эмаль EMPILS ПФ-115 «ЭКОНОМ» чёрная',
        'abc_cumulative': 0.655,
        'comment': 'НОВАЯ ПОЗИЦИЯ из ABC-анализа (топ-4). Новый бренд в матрице: EMPILS ЭКОНОМ',
    },
    {
        'idx': 300, 'sphere': 'Металл', 'purpose': 'Эмаль по металлу', 'base': 'Органическая',
        'brand': 'СТРОЙКОМПЛЕКТ', 'type': 'Эмали Алкидные', 'grp_code': 72676, 'abc': 'А',
        'name': 'Эмаль ПФ-115 СТРОЙКОМПЛЕКТ чёрная 10 кг',
        'отгрузочная': 1640.57, 'базовая': 2050.71, 'фасовка': 10, 'лист': 'СТРОЙКОМПЛЕКТ',
        'match_type': 'new_from_abc',
        'ценовое_имя': 'Эмаль ПФ-115 СТРОЙКОМПЛЕКТ чёрная',
        'abc_cumulative': 0.673,
        'comment': 'НОВАЯ ПОЗИЦИЯ из ABC-анализа (топ-5). Бренд СТРОЙКОМПЛЕКТ (2-я позиция этого бренда)',
    },
]

# === 1. Обновить лист "300 SKU — детально" ===
print("\n=== Обновление листа '300 SKU — детально' ===")
ws_det = wb["300 SKU — детально"]

# Найти строки по idx и обновить
row_by_idx = {}
for r in range(4, ws_det.max_row + 1):
    idx = ws_det.cell(row=r, column=1).value
    if idx is not None:
        row_by_idx[idx] = r

# Заменить 3 SKU
for idx, repl in REPLACEMENTS.items():
    r = row_by_idx[idx]
    ws_det.cell(row=r, column=2, value=repl['new_sphere'])
    ws_det.cell(row=r, column=3, value=repl['new_purpose'])
    ws_det.cell(row=r, column=4, value=repl['new_base'])
    ws_det.cell(row=r, column=5, value=repl['new_brand'])
    ws_det.cell(row=r, column=6, value=repl['new_type'])
    ws_det.cell(row=r, column=7, value=repl['new_name'])
    ws_det.cell(row=r, column=8, value=repl['new_grp_code'])
    ws_det.cell(row=r, column=9, value=repl['new_abc'])
    print(f"  ✓ Заменён #{idx} (строка {r}): {repl['new_name'][:60]}")

# Удалить 5 SKU (с конца, чтобы не сбить нумерацию)
for idx in sorted(DELETE_IDX, reverse=True):
    r = row_by_idx[idx]
    ws_det.delete_rows(r, 1)
    print(f"  ✗ Удалён #{idx} (строка {r})")

# Перенумеровать # оставшихся SKU (чтобы 1-295 шли подряд)
# После удаления 5 и замены 3 (без удаления) = 295 SKU
# Найдём все оставшиеся строки и перенумеруем
all_rows = []
for r in range(4, ws_det.max_row + 1):
    idx_val = ws_det.cell(row=r, column=1).value
    if idx_val is not None:
        all_rows.append(r)

print(f"\n  После удаления: {len(all_rows)} строк с SKU")

# Перенумеруем с 1 по 295
for new_idx, r in enumerate(all_rows, 1):
    old_idx = ws_det.cell(row=r, column=1).value
    if old_idx != new_idx:
        ws_det.cell(row=r, column=1, value=new_idx)
print(f"  Перенумеровано: 1-{len(all_rows)}")

# Добавить 5 новых SKU в конец (новые idx 296-300)
last_row = ws_det.max_row
for i, item in enumerate(NEW_ITEMS):
    r = last_row + 1 + i
    ws_det.cell(row=r, column=1, value=item['idx'])
    ws_det.cell(row=r, column=2, value=item['sphere'])
    ws_det.cell(row=r, column=3, value=item['purpose'])
    ws_det.cell(row=r, column=4, value=item['base'])
    ws_det.cell(row=r, column=5, value=item['brand'])
    ws_det.cell(row=r, column=6, value=item['type'])
    ws_det.cell(row=r, column=7, value=item['name'])
    ws_det.cell(row=r, column=8, value=item['grp_code'])
    ws_det.cell(row=r, column=9, value=item['abc'])
    print(f"  + Добавлен #{item['idx']} (строка {r}): {item['name'][:60]}")

# === 2. Обновить лист "300 SKU с ценами" — пересоздать на основе обновлённого детального листа ===
print("\n=== Пересоздание листа '300 SKU с ценами' ===")
if "300 SKU с ценами" in wb.sheetnames:
    del wb["300 SKU с ценами"]

ws_prices = wb.create_sheet("300 SKU с ценами")

# Заголовок
ws_prices.merge_cells("A1:N1")
ws_prices["A1"] = "Стартовый ассортимент Олимп — 300 SKU с отгрузочными ценами (ФИНАЛ v3)"
ws_prices["A1"].font = TITLE_FONT
ws_prices["A1"].alignment = Alignment(horizontal="left", vertical="center")
ws_prices.row_dimensions[1].height = 28

ws_prices.merge_cells("A2:N2")
ws_prices["A2"] = ("Источник цен: прайс-лист «16.10.2026 РОССИЯ с НДС, руб.» | "
                   "Удалено: 5 SKU РАСЦВЕТ без альтернатив | Заменено на альтернативы: 3 SKU | "
                   "Добавлено из ABC-анализа: 5 SKU | ИТОГО: 300 SKU")
ws_prices["A2"].font = SUBTITLE_FONT
ws_prices["A2"].alignment = Alignment(horizontal="left", vertical="center")
ws_prices.row_dimensions[2].height = 18

# Шапка (строка 4)
HEADERS = [
    ("#", 5), ("Сфера", 14), ("Назначение", 28), ("Основа", 16),
    ("Бренд", 16), ("Тип продукции", 22), ("Наименование", 60),
    ("Гр. код", 9), ("ABC", 6),
    ("Отгрузочная цена, руб", 14), ("Базовая цена, руб", 13),
    ("Фасовка в прайсе, кг", 11), ("Лист прайса", 16), ("Статус сопоставления", 22),
]
for col_idx, (title, width) in enumerate(HEADERS, 1):
    cell = ws_prices.cell(row=4, column=col_idx, value=title)
    cell.fill = HEAD_FILL
    cell.font = HEAD_FONT
    cell.alignment = HEAD_ALIGN
    cell.border = BORDER
    ws_prices.column_dimensions[get_column_letter(col_idx)].width = width
ws_prices.row_dimensions[4].height = 38

# Прочитать обновлённый детальный лист и заполнить цены
# Сначала загрузим цены из прайс-листа 16.10.2026
from collections import defaultdict
import re

PRICE_PATH = "/home/z/my-project/upload/16.10.2026 РОССИЯ с НДС, руб..xlsx"
wb_price = load_workbook(PRICE_PATH, data_only=True)
price_dict = defaultdict(list)
for sn in wb_price.sheetnames:
    ws = wb_price[sn]
    current_name = None
    for row in ws.iter_rows(min_row=14, max_row=ws.max_row, values_only=True):
        f_val = row[5] if len(row) > 5 else None
        m_val = row[12] if len(row) > 12 else None
        n_val = row[13] if len(row) > 13 else None
        h_val = row[7] if len(row) > 7 else None
        if f_val is not None and isinstance(f_val, str) \
            and not f_val.startswith('СТО') and not f_val.startswith('ТУ') and not f_val.startswith('ГОСТ'):
            current_name = f_val.strip()
        if (m_val is not None and isinstance(m_val, (int, float))) or (n_val is not None and isinstance(n_val, (int, float))):
            price_dict[sn].append({
                'name': current_name, 'fасовка': h_val, 'базовая': m_val, 'отгрузочная': n_val,
            })

# Используем ранее созданный алгоритм сопоставления
# (импортируем функции из build_final.py через exec)
import sys
sys.path.insert(0, '/home/z/my-project/scripts')

# Загрузим ранее рассчитанные результаты из results.json
import json
with open('/home/z/my-project/scripts/results.json', 'r', encoding='utf-8') as f:
    prev_results = json.load(f)
# Преобразуем в словарь по idx
prev_results_by_idx = {r['idx']: r for r in prev_results}

# Также добавим замены и новые
def get_price_for_item(item):
    """Returns price info for an item (idx, name, brand)."""
    idx = item.get('idx')
    
    # Если это заменённый SKU — вернуть данные из REPLACEMENTS
    if idx in REPLACEMENTS:
        repl = REPLACEMENTS[idx]
        return {
            'отгрузочная': repl['отгрузочная'],
            'базовая': repl['базовая'],
            'фасовка': repl['фасовка'],
            'лист': repl['лист'],
            'match_type': 'replacement',
            'ценовое_имя': repl['ценовое_имя'],
            'status': f"ЗАМЕНА: {repl['comment']}",
        }
    
    # Если это новый SKU из ABC — вернуть данные из NEW_ITEMS
    for ni in NEW_ITEMS:
        if ni['idx'] == idx:
            return {
                'отгрузочная': ni['отгрузочная'],
                'базовая': ni['базовая'],
                'фасовка': ni['фасовка'],
                'лист': ni['лист'],
                'match_type': 'new_from_abc',
                'ценовое_имя': ni['ценовое_имя'],
                'status': f"НОВАЯ: {ni['comment']}",
            }
    
    # Иначе — использовать ранее рассчитанный результат
    prev = prev_results_by_idx.get(idx)
    if prev and prev.get('отгрузочная') is not None:
        return {
            'отгрузочная': prev['отгрузочная'],
            'базовая': prev['базовая'],
            'фасовка': prev.get('фасовка_прайс'),
            'лист': prev.get('лист', ''),
            'match_type': prev.get('match_type', ''),
            'ценовое_имя': prev.get('ценовое_имя', ''),
            'status': f"Цена: {prev.get('match_type', 'matched')}",
        }
    return None

# Заполнить строки с ценами
row_i = 5
for r in range(4, ws_det.max_row + 1):
    idx = ws_det.cell(row=r, column=1).value
    if idx is None:
        continue
    
    item = {
        'idx': idx,
        'sphere': ws_det.cell(row=r, column=2).value,
        'purpose': ws_det.cell(row=r, column=3).value,
        'base': ws_det.cell(row=r, column=4).value,
        'brand': ws_det.cell(row=r, column=5).value,
        'type': ws_det.cell(row=r, column=6).value,
        'name': ws_det.cell(row=r, column=7).value,
        'grp_code': ws_det.cell(row=r, column=8).value,
        'abc': ws_det.cell(row=r, column=9).value,
    }
    
    price_info = get_price_for_item(item)
    
    vals = [
        item['idx'],
        item['sphere'],
        item['purpose'],
        item['base'],
        item['brand'],
        item['type'],
        item['name'],
        item['grp_code'],
        item['abc'],
        price_info['отгрузочная'] if price_info else None,
        price_info['базовая'] if price_info else None,
        price_info['фасовка'] if price_info else None,
        price_info['лист'] if price_info else '',
        price_info['status'] if price_info else 'не найдено в прайс-листе',
    ]
    
    is_new = idx in [ni['idx'] for ni in NEW_ITEMS]
    is_replaced = idx in REPLACEMENTS
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

# Закрепить заголовок и автофильтр
ws_prices.freeze_panes = "B5"
last_col = get_column_letter(14)
ws_prices.auto_filter.ref = f"A4:{last_col}{row_i - 1}"

print(f"  ✓ Заполнено {row_i - 5} строк с ценами")

# === 3. Обновить лист "Сводка" — пересчитать матрицы ===
print("\n=== Пересчёт матриц на листе 'Сводка' ===")
ws_sv = wb["Сводка"]

# Сначала соберём статистику из обновлённого листа "300 SKU — детально"
from collections import Counter
items_all = []
for r in range(4, ws_det.max_row + 1):
    idx = ws_det.cell(row=r, column=1).value
    if idx is None:
        continue
    items_all.append({
        'idx': idx,
        'sphere': ws_det.cell(row=r, column=2).value,
        'purpose': ws_det.cell(row=r, column=3).value,
        'base': ws_det.cell(row=r, column=4).value,
        'brand': ws_det.cell(row=r, column=5).value,
        'type': ws_det.cell(row=r, column=6).value,
    })

print(f"  Всего SKU в обновлённом детальном листе: {len(items_all)}")

# === Матрица 1: Основа × Сфера (rows 5-7) ===
# Headers: Основа / Сфера, Металл, Помещения (интерьер), Дерево, Пол, Фасад, ИТОГО, Доля
spheres_order = ['Металл', 'Помещения (интерьер)', 'Дерево', 'Пол', 'Фасад']
bases_order = ['Органическая', 'Водно-дисперсионная']

# Считаем
matrix_base_sphere = {(b, s): 0 for b in bases_order for s in spheres_order}
for it in items_all:
    key = (it['base'], it['sphere'])
    if key in matrix_base_sphere:
        matrix_base_sphere[key] += 1

# Записываем
for row_offset, base in enumerate(bases_order):
    r = 5 + row_offset
    ws_sv.cell(row=r, column=1, value=base)
    for col_offset, sphere in enumerate(spheres_order):
        ws_sv.cell(row=r, column=2 + col_offset, value=matrix_base_sphere[(base, sphere)])
    # ИТОГО
    total_row = sum(matrix_base_sphere[(base, s)] for s in spheres_order)
    ws_sv.cell(row=r, column=7, value=total_row)
    # Доля
    pct = total_row / len(items_all)
    ws_sv.cell(row=r, column=8, value=f"{pct*100:.1f}%")

# Строка ИТОГО
r = 7
ws_sv.cell(row=r, column=1, value='ИТОГО')
for col_offset, sphere in enumerate(spheres_order):
    col_total = sum(matrix_base_sphere[(b, sphere)] for b in bases_order)
    ws_sv.cell(row=r, column=2 + col_offset, value=col_total)
ws_sv.cell(row=r, column=7, value=len(items_all))
ws_sv.cell(row=r, column=8, value='100%')

print(f"  ✓ Матрица Основа × Сфера обновлена")

# === Матрица 2: СФЕРА × БРЕНД (rows 10-16) ===
# Найдём все бренды в обновлённом ассортименте
brand_counts = Counter(it['brand'] for it in items_all)
print(f"  Бренды в обновлённом ассортименте: {dict(brand_counts)}")

# Порядок брендов: оригинальные + новые в конце
original_brands = ['ОРЕОЛ TITAN', 'Ореол', 'Ореол Дисконт', 'NEWTONE', 'AvantGarde',
                    'Расцвет', 'Простокрашено! БАУ', 'PROFIWOOD']
new_brands = [b for b in brand_counts if b not in original_brands]
all_brands = original_brands + new_brands
print(f"  Все бренды (с новыми): {all_brands}")

# Перепишем заголовок строки 10 (чтобы добавить новые бренды)
# Очистим старую строку 10 и 11-16
for r in range(10, 17):
    for c in range(1, 12):
        ws_sv.cell(row=r, column=c, value=None)

# Заголовок строки 10
ws_sv.cell(row=10, column=1, value='СФЕРА / БРЕНД')
for col_offset, brand in enumerate(all_brands):
    ws_sv.cell(row=10, column=2 + col_offset, value=brand)
ws_sv.cell(row=10, column=2 + len(all_brands), value='ИТОГО')

# Применим форматирование к заголовку
for c in range(1, 3 + len(all_brands)):
    cell = ws_sv.cell(row=10, column=c)
    cell.fill = HEAD_FILL
    cell.font = HEAD_FONT
    cell.alignment = HEAD_ALIGN
    cell.border = BORDER

# Считаем матрицу
matrix_sphere_brand = {(s, b): 0 for s in spheres_order for b in all_brands}
for it in items_all:
    if it['brand'] in all_brands:
        matrix_sphere_brand[(it['sphere'], it['brand'])] += 1

# Записываем строки 11-15 (по сферам)
for row_offset, sphere in enumerate(spheres_order):
    r = 11 + row_offset
    ws_sv.cell(row=r, column=1, value=sphere)
    for col_offset, brand in enumerate(all_brands):
        val = matrix_sphere_brand[(sphere, brand)]
        if val > 0:
            ws_sv.cell(row=r, column=2 + col_offset, value=val)
    # ИТОГО по сфере
    total_row = sum(matrix_sphere_brand[(sphere, b)] for b in all_brands)
    ws_sv.cell(row=r, column=2 + len(all_brands), value=total_row)
    
    # Применить форматирование
    for c in range(1, 3 + len(all_brands)):
        cell = ws_sv.cell(row=r, column=c)
        cell.font = BODY_FONT
        cell.border = BORDER
        if c > 1:
            cell.alignment = CENTER_ALIGN
        else:
            cell.alignment = Alignment(horizontal="left", vertical="center")

# Строка ИТОГО (16)
r = 16
ws_sv.cell(row=r, column=1, value='ИТОГО')
for col_offset, brand in enumerate(all_brands):
    col_total = brand_counts[brand]
    ws_sv.cell(row=r, column=2 + col_offset, value=col_total)
ws_sv.cell(row=r, column=2 + len(all_brands), value=len(items_all))

# Форматирование ИТОГО
for c in range(1, 3 + len(all_brands)):
    cell = ws_sv.cell(row=r, column=c)
    cell.font = Font(name="Calibri", size=10, bold=True)
    cell.fill = PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid")
    cell.border = BORDER
    if c > 1:
        cell.alignment = CENTER_ALIGN
    else:
        cell.alignment = Alignment(horizontal="left", vertical="center")

# Ширина колонок для брендов
for col_offset in range(len(all_brands) + 1):
    ws_sv.column_dimensions[get_column_letter(2 + col_offset)].width = 14
ws_sv.column_dimensions[get_column_letter(2 + len(all_brands))].width = 8
ws_sv.column_dimensions['A'].width = 24

print(f"  ✓ Матрица СФЕРА × БРЕНД обновлена (с новыми брендами: {new_brands})")

# === Матрица 3: ДЕТАЛЬНАЯ КЛАССИФИКАЦИЯ ПО НАЗНАЧЕНИЮ (rows 19-35) ===
# Назначения в исходном файле:
purposes_order = [
    'Грунт-эмаль по металлу', 'Эмаль по металлу', 'Пол',
    'Стены и потолки (интерьер)', 'Защита и декор дерева',
    'Эмали ВД прочие', 'Лаки прочие', 'Грунтовки', 'Фасад',
    'Прочее', 'Яхтный лак', 'Шпатлёвки', 'Паркетный лак',
    'Молотковая эмаль', 'Потолки', 'Резиновая краска',
]

# Считаем матрицу назначений
matrix_purpose_sphere = {(p, s): 0 for p in purposes_order for s in spheres_order}
purpose_counts = Counter()
for it in items_all:
    key = (it['purpose'], it['sphere'])
    if key in matrix_purpose_sphere:
        matrix_purpose_sphere[key] += 1
        purpose_counts[it['purpose']] += 1

# Сортируем назначения по убыванию количества
purposes_sorted = sorted(purposes_order, key=lambda p: -purpose_counts.get(p, 0))

# Перепишем строки 20-35
# Сначала очистим
for r in range(20, 36):
    for c in range(1, 8):
        ws_sv.cell(row=r, column=c, value=None)

# Заголовок строки 19
ws_sv.cell(row=19, column=1, value='Сфера / Назначение')
sphere_cols = ['Металл', 'Помещения (интерьер)', 'Дерево', 'Пол', 'Фасад']
for col_offset, sp in enumerate(sphere_cols):
    ws_sv.cell(row=19, column=2 + col_offset, value=sp)
ws_sv.cell(row=19, column=7, value='ИТОГО')

# Форматирование заголовка
for c in range(1, 8):
    cell = ws_sv.cell(row=19, column=c)
    cell.fill = HEAD_FILL
    cell.font = HEAD_FONT
    cell.alignment = HEAD_ALIGN
    cell.border = BORDER

# Данные
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
    
    # Форматирование
    for c in range(1, 8):
        cell = ws_sv.cell(row=r, column=c)
        cell.font = BODY_FONT
        cell.border = BORDER
        if c > 1:
            cell.alignment = CENTER_ALIGN
        else:
            cell.alignment = Alignment(horizontal="left", vertical="center")

print(f"  ✓ Матрица ДЕТАЛЬНАЯ КЛАССИФИКАЦИЯ ПО НАЗНАЧЕНИЮ обновлена ({len(purposes_sorted)} назначений)")

# === 4. Удалить старый лист "8 SKU — альтернативы" если есть, создать новый "Изменения v3" ===
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
ws_log["A2"] = ("В соответствии с заданием: 5 SKU РАСЦВЕТ (без альтернатив в прайс-листе) удалены и заменены "
                "5 новыми популярными SKU из ABC-анализа. 3 SKU (с найденными альтернативами) заменены на эти альтернативы.")
ws_log["A2"].font = SUBTITLE_FONT
ws_log["A2"].alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
ws_log.row_dimensions[2].height = 36

# Шапка
log_headers = [
    ("Действие", 14), ("# SKU", 7), ("Бренд", 18), ("Сфера", 16),
    ("Наименование (новое)", 55), ("Отгруз. цена, руб", 14), ("Комментарий", 70),
]
for col_idx, (title, width) in enumerate(log_headers, 1):
    cell = ws_log.cell(row=4, column=col_idx, value=title)
    cell.fill = HEAD_FILL
    cell.font = HEAD_FONT
    cell.alignment = HEAD_ALIGN
    cell.border = BORDER
    ws_log.column_dimensions[get_column_letter(col_idx)].width = width
ws_log.row_dimensions[4].height = 38

# === Удалённые ===
row_i = 5
ws_log.cell(row=row_i, column=1, value="УДАЛЕНО (5)").font = Font(name="Calibri", size=11, bold=True, color="C00000")
ws_log.cell(row=row_i, column=1).fill = DELETED_FILL
for c in range(2, 8):
    ws_log.cell(row=row_i, column=c).fill = DELETED_FILL
ws_log.merge_cells(f"A{row_i}:G{row_i}")
row_i += 1

deleted_items_info = [
    (38, 'Расцвет', 'Металл', 'Грунт-эмаль "РАСЦВЕТ" 3 в 1 голубая 0.9 кг',
     'Цвет ГОЛУБАЯ для 3 в 1 отсутствует в прайс-листе. Доступные цвета: красная RAL 3020, серый графит RAL 7024.'),
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

# === Заменённые ===
row_i += 1
ws_log.cell(row=row_i, column=1, value="ЗАМЕНЕНО (3)").font = Font(name="Calibri", size=11, bold=True, color="7F6000")
ws_log.cell(row=row_i, column=1).fill = REPLACED_FILL
for c in range(2, 8):
    ws_log.cell(row=row_i, column=c).fill = REPLACED_FILL
ws_log.merge_cells(f"A{row_i}:G{row_i}")
row_i += 1

for idx, repl in REPLACEMENTS.items():
    vals = [
        "Замена",
        idx,
        repl['new_brand'],
        repl['new_sphere'],
        repl['new_name'],
        repl['отгрузочная'],
        repl['comment'],
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

# === Новые ===
row_i += 1
ws_log.cell(row=row_i, column=1, value="ДОБАВЛЕНО из ABC-анализа (5)").font = Font(name="Calibri", size=11, bold=True, color="375623")
ws_log.cell(row=row_i, column=1).fill = NEW_ROW_FILL
for c in range(2, 8):
    ws_log.cell(row=row_i, column=c).fill = NEW_ROW_FILL
ws_log.merge_cells(f"A{row_i}:G{row_i}")
row_i += 1

for item in NEW_ITEMS:
    vals = [
        "Новая",
        item['idx'],
        item['brand'],
        item['sphere'],
        item['name'],
        item['отгрузочная'],
        f"{item['comment']} | ABC cumulative={item['abc_cumulative']:.3f} (топ в группе А) | Цена: {item['ценовое_имя']}",
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

# Убедиться, что порядок листов правильный
print(f"\nЛисты в финальном файле: {wb.sheetnames}")

wb.save(OUT_PATH)
print(f"\n✓ Файл сохранён: {OUT_PATH}")
print(f"  Размер: {os.path.getsize(OUT_PATH):,} байт")
