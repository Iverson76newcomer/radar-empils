"""Сбор статистики по БРОЗЭКС — анализ классов, подсчёт SKU/АКБ/продаж."""
from openpyxl import load_workbook
from collections import defaultdict

wb = load_workbook('/home/z/my-project/upload/Титан БРОЗЭКС.xlsx', data_only=True)
ws = wb['Лист1']

# Пропустим первые 2 строки (заголовок)
classes_data = defaultdict(lambda: {
    'sku_2025': set(),   # уникальные SKU с продажами 2025
    'sku_2026': set(),   # уникальные SKU с продажами 2026
    'akb_2025': [],      # список значений АКБ 2025
    'akb_2026': [],      # список значений АКБ 2026
    'sales_2025': 0.0,
    'sales_2026': 0.0,
    'sku_data': defaultdict(lambda: {'akb25': 0, 'sales25': 0.0, 'akb26': 0, 'sales26': 0.0}),
})

for row in ws.iter_rows(min_row=3, max_row=ws.max_row, values_only=True):
    if not row or len(row) < 8:
        continue
    dist, series, cls, name, akb25, sales25, akb26, sales26 = row[0], row[1], row[2], row[3], row[4], row[5], row[6], row[7]
    if dist is None or cls is None:
        continue
    cls = str(cls).strip()
    # пропускаем строки "X Итог"
    if 'Итог' in cls or 'итог' in cls:
        continue
    if cls.endswith(' Итог') or 'Итог' in str(cls):
        continue
    # также класс может быть "A Итог" и т.п.
    parts = cls.split()
    if len(parts) > 1 and parts[-1] == 'Итог':
        continue
    if name is None or str(name).strip() == '':
        continue  # это строка-итог без наименования
    # Парсим числа
    try:
        akb25 = int(akb25) if akb25 not in (None, '') else 0
    except (ValueError, TypeError):
        akb25 = 0
    try:
        akb26 = int(akb26) if akb26 not in (None, '') else 0
    except (ValueError, TypeError):
        akb26 = 0
    try:
        sales25 = float(sales25) if sales25 not in (None, '') else 0.0
    except (ValueError, TypeError):
        sales25 = 0.0
    try:
        sales26 = float(sales26) if sales26 not in (None, '') else 0.0
    except (ValueError, TypeError):
        sales26 = 0.0

    if cls not in classes_data:
        classes_data[cls]  # init
    d = classes_data[cls]
    name = str(name).strip()
    # SKU с продажами в 2025
    if akb25 > 0 or sales25 > 0:
        d['sku_2025'].add(name)
        d['akb_2025'].append(akb25)
    if akb26 > 0 or sales26 > 0:
        d['sku_2026'].add(name)
        d['akb_2026'].append(akb26)
    d['sales_2025'] += sales25
    d['sales_2026'] += sales26
    # детализация
    sku_d = d['sku_data'][name]
    sku_d['akb25'] += akb25
    sku_d['sales25'] += sales25
    sku_d['akb26'] += akb26
    sku_d['sales26'] += sales26

# Печать результатов
print("=" * 100)
print(f"{'Класс':<8} | {'SKU 2025':<10} | {'SKU 2026':<10} | {'АКБ 2025 (avg)':<16} | {'АКБ 2026 (avg)':<16} | {'Прод 2025':<15} | {'Прод 2026':<15}")
print("=" * 100)
for cls in sorted(classes_data.keys()):
    d = classes_data[cls]
    avg_akb25 = sum(d['akb_2025']) / len(d['akb_2025']) if d['akb_2025'] else 0
    avg_akb26 = sum(d['akb_2026']) / len(d['akb_2026']) if d['akb_2026'] else 0
    print(f"{cls:<8} | {len(d['sku_2025']):<10} | {len(d['sku_2026']):<10} | {avg_akb25:<16.2f} | {avg_akb26:<16.2f} | {d['sales_2025']:<15,.2f} | {d['sales_2026']:<15,.2f}")

print("\n\n=== ИТОГО ===")
tot_sku25 = sum(len(d['sku_2025']) for d in classes_data.values())
tot_sku26 = sum(len(d['sku_2026']) for d in classes_data.values())
tot_akb25 = sum(sum(d['akb_2025']) for d in classes_data.values())
tot_akb26 = sum(sum(d['akb_2026']) for d in classes_data.values())
tot_sales25 = sum(d['sales_2025'] for d in classes_data.values())
tot_sales26 = sum(d['sales_2026'] for d in classes_data.values())
print(f"SKU 2025: {tot_sku25}, SKU 2026: {tot_sku26}")
print(f"АКБ 2025 (sum): {tot_akb25}, АКБ 2026 (sum): {tot_akb26}")
print(f"Продажи 2025: {tot_sales25:,.2f}, Продажи 2026: {tot_sales26:,.2f}")
print(f"Δ продаж: {(tot_sales26-tot_sales25)/tot_sales25*100:+.1f}%")

# Сохраняем в pickle для дальнейшего использования
import json
out = {}
for cls, d in classes_data.items():
    avg_akb25 = round(sum(d['akb_2025']) / len(d['akb_2025'])) if d['akb_2025'] else 0
    avg_akb26 = round(sum(d['akb_2026']) / len(d['akb_2026'])) if d['akb_2026'] else 0
    out[cls] = {
        'sku25': len(d['sku_2025']),
        'sku26': len(d['sku_2026']),
        'akb25_avg': avg_akb25,
        'akb26_avg': avg_akb26,
        'akb25_sum': sum(d['akb_2025']),
        'akb26_sum': sum(d['akb_2026']),
        'sales25': d['sales_2025'],
        'sales26': d['sales_2026'],
        'sku_list': sorted(d['sku_data'].keys()),
        'sku_data': {k: dict(v) for k, v in d['sku_data'].items()},
    }
with open('/home/z/my-project/scripts/brozex_stats.json', 'w', encoding='utf-8') as f:
    json.dump(out, f, ensure_ascii=False, indent=2, default=str)
print("\nДанные сохранены в brozex_stats.json")
