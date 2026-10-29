"""Сбор статистики TITAN по дистрибьютору Ваш Дом.
   Структура файла: .xls, 1 лист, 26 магазинов в колонках, SKU в строках.
   Период: 01.01.2026 - 06.09.2026 (только 2026, без 2025).
"""
import xlrd
import json
from collections import defaultdict

wb = xlrd.open_workbook('/home/z/my-project/upload/Титан Ваш Дом 09.09.2026.xls')
ws = wb.sheet_by_name('TDSheet')

# Заголовок с магазинами — строка 4 (0-индекс: 3)
# Колонки 0-6: Код, "", "", Номенклатура, Поставщик, ABC, Статус
# С колонки 7 начинаются магазины (по 3 подколонки: Количество, Остаток, Продажи)
# Последние 3 колонки — Итого

# Соберём список магазинов
# Заголовок с магазинами — строка 5 (0-индекс: 4) — после "Параметры:", "Тип параметра", пустая строка
# Найдём строку с "Код" в начале — это заголовок
header_row = None
for ri in range(0, 8):
    if str(ws.cell_value(ri, 0)).strip() == 'Код':
        header_row = ri
        break
print(f"Header row index: {header_row}")
# Колонки 0-6: Код, "", "", Номенклатура, Поставщик, ABC, Статус
# С колонки 7 начинаются магазины (по 3 подколонки: Количество, Остаток, Продажи)
# Последние 3 колонки — Итого

shops = []
ci = 7
while ci < ws.ncols - 3:
    shop_name = ws.cell_value(header_row, ci)
    if shop_name and shop_name != 'Итого' and str(shop_name).strip():
        shops.append({'name': str(shop_name).strip(), 'col_qty': ci, 'col_stock': ci+1, 'col_sales': ci+2})
        ci += 3
    else:
        break

print(f"Найдено магазинов: {len(shops)}")
for s in shops:
    print(f"  - {s['name']}")

# Категории продуктов TITAN:
# ОРЕОЛ TITAN — грунт-эмали 3 в 1 (основной продукт)
# ОРЕОЛ ДИСКОНТ — водно-дисперсионные краски (побочный)
# ОРЕОЛ — эмали радиаторов/акриловые/по пластику

# Сбор данных по SKU
# Структура: строка начинается с кода (если это SKU) или пустая (если это категория)
sku_data = {}  # code -> {name, abc, status, shops: {shop_name: {qty, stock, sales}}}

# Данные начинаются со строки после header_row + 1 (строка с подзаголовками Количество/Остаток/Продажи)
data_start = header_row + 2

for ri in range(data_start, ws.nrows):
    code = ws.cell_value(ri, 0)
    name = ws.cell_value(ri, 3)
    supplier = ws.cell_value(ri, 4)
    abc = ws.cell_value(ri, 5)
    status = ws.cell_value(ri, 6)
    if not code or not name:
        continue  # пропускаем строки-категории
    code = str(int(code)) if isinstance(code, float) else str(code)
    name = str(name).strip()
    if 'TITAN' not in name.upper() and 'ОРЕОЛ' not in name.upper():
        continue  # только продукты TITAN/ОРЕОЛ
    if 'Уценка' in name or 'уценка' in name:
        continue
    # Оставляем ТОЛЬКО ОРЕОЛ TITAN (грунт-эмали 3 в 1)
    if 'ОРЕОЛ TITAN' not in name.upper():
        continue
    sku_data[code] = {
        'name': name, 'abc': str(abc), 'status': str(status),
        'shops': {},
        'total_qty': 0, 'total_stock': 0, 'total_sales': 0.0,
        'active_shops': 0,
    }
    for shop in shops:
        try:
            qty = ws.cell_value(ri, shop['col_qty'])
            stock = ws.cell_value(ri, shop['col_stock'])
            sales = ws.cell_value(ri, shop['col_sales'])
            qty = float(qty) if qty != '' else 0
            stock = float(stock) if stock != '' else 0
            sales = float(sales) if sales != '' else 0
        except (ValueError, TypeError):
            qty = stock = sales = 0
        if qty > 0 or sales > 0 or stock > 0:
            sku_data[code]['shops'][shop['name']] = {'qty': qty, 'stock': stock, 'sales': sales}
            sku_data[code]['total_qty'] += qty
            sku_data[code]['total_stock'] += stock
            sku_data[code]['total_sales'] += sales
            if sales > 0:
                sku_data[code]['active_shops'] += 1

# Категоризация SKU — только ОРЕОЛ TITAN, по цветам RAL
def categorize(name):
    n = name.upper()
    # Извлекаем RAL
    import re
    m = re.search(r'RAL\s*(\d{4})', name)
    if m:
        ral = m.group(1)
        ral_names = {
            '9010': 'RAL 9010 — Белый',
            '9005': 'RAL 9005 — Чёрный',
            '7004': 'RAL 7004 — Серый',
            '7024': 'RAL 7024 — Графитовый серый',
            '6005': 'RAL 6005 — Зелёный',
            '8017': 'RAL 8017 — Шоколадно-коричневый',
            '1021': 'RAL 1021 — Жёлтый',
            '3009': 'RAL 3009 — Красно-коричневый',
        }
        return ral_names.get(ral, f'RAL {ral}')
    return 'Прочее'

# Группировка по категориям
categories = defaultdict(lambda: {
    'sku_count': 0, 'total_qty': 0, 'total_sales': 0.0,
    'active_shops_sum': 0, 'sku_list': [],
})

for code, d in sku_data.items():
    cat = categorize(d['name'])
    categories[cat]['sku_count'] += 1
    categories[cat]['total_qty'] += d['total_qty']
    categories[cat]['total_sales'] += d['total_sales']
    categories[cat]['active_shops_sum'] += d['active_shops']
    categories[cat]['sku_list'].append({
        'code': code, 'name': d['name'], 'abc': d['abc'], 'status': d['status'],
        'qty': d['total_qty'], 'stock': d['total_stock'], 'sales': d['total_sales'],
        'active_shops': d['active_shops'],
    })

# Печать сводки по категориям
print("\n" + "=" * 100)
print(f"{'Категория':<45} | {'SKU':<5} | {'Кол-во':<8} | {'Продажи, ₽':<14} | {'Активн. ТТ':<10}")
print("=" * 100)
for cat, d in sorted(categories.items(), key=lambda x: -x[1]['total_sales']):
    print(f"{cat:<45} | {d['sku_count']:<5} | {int(d['total_qty']):<8} | {d['total_sales']:<14,.2f} | {d['active_shops_sum']:<10}")

# Сводка по магазинам
shop_totals = defaultdict(lambda: {'qty': 0, 'stock': 0, 'sales': 0.0, 'sku_count': 0})
for code, d in sku_data.items():
    for shop_name, sd in d['shops'].items():
        if sd['sales'] > 0:
            shop_totals[shop_name]['qty'] += sd['qty']
            shop_totals[shop_name]['stock'] += sd['stock']
            shop_totals[shop_name]['sales'] += sd['sales']
            shop_totals[shop_name]['sku_count'] += 1

print("\n" + "=" * 100)
print(f"{'Магазин':<45} | {'SKU':<5} | {'Продажи, ₽':<14}")
print("=" * 100)
for shop, d in sorted(shop_totals.items(), key=lambda x: -x[1]['sales']):
    print(f"{shop[:45]:<45} | {d['sku_count']:<5} | {d['sales']:<14,.2f}")

# Итоги
total_sku = len(sku_data)
total_qty = sum(d['total_qty'] for d in sku_data.values())
total_sales = sum(d['total_sales'] for d in sku_data.values())
active_shops_count = len(shop_totals)
print(f"\n=== ИТОГО ===")
print(f"Всего SKU TITAN: {total_sku}")
print(f"Всего продано штук: {int(total_qty)}")
print(f"Всего продажи: {total_sales:,.2f} ₽")
print(f"Активных магазинов: {active_shops_count} из {len(shops)}")

# Сохраняем в JSON
out = {
    'shops': [s['name'] for s in shops],
    'categories': {k: dict(v) for k, v in categories.items()},
    'sku_data': {k: {kk: vv for kk, vv in v.items() if kk != 'shops'} | {'shops': v['shops']}
                  for k, v in sku_data.items()},
    'shop_totals': {k: dict(v) for k, v in shop_totals.items()},
    'totals': {
        'sku_count': total_sku, 'total_qty': int(total_qty),
        'total_sales': total_sales, 'active_shops': active_shops_count,
    }
}
with open('/home/z/my-project/scripts/vashdom_stats.json', 'w', encoding='utf-8') as f:
    json.dump(out, f, ensure_ascii=False, indent=2, default=str)
print("\nДанные сохранены в vashdom_stats.json")
