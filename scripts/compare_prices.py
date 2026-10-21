import openpyxl
from openpyxl import load_workbook
from collections import defaultdict
import re

# === Загрузить ОБА прайс-листа ===
def load_prices(path):
    wb = load_workbook(path, data_only=True)
    price_dict = defaultdict(list)
    for sn in wb.sheetnames:
        ws = wb[sn]
        current_name = None
        for row_idx, row in enumerate(ws.iter_rows(min_row=14, max_row=ws.max_row, values_only=True), 14):
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
    return price_dict

old_prices = load_prices("/home/z/my-project/upload/21.08.2026 РОССИЯ с НДС, руб..xlsx")
new_prices = load_prices("/home/z/my-project/upload/16.10.2026 РОССИЯ с НДС, руб..xlsx")

print(f"Старый прайс (21.08): {sum(len(v) for v in old_prices.values())} позиций")
print(f"Новый прайс (16.10): {sum(len(v) for v in new_prices.values())} позиций")

# Сравнить цены на нескольких позициях
print("\n=== Сравнение цен на 5 случайных позиций ===")
checks = [
    ('AvantGarde', 'Защитно-декоративное покрытие для дерева AvantGarde белое', 0.8),
    ('ОРЕОЛ TITAN', 'Грунт-эмаль быстросохнущая ОРЕОЛ TITAN матовая белая RAL 9010', 0.9),
    ('Простокрашено! БАУ', 'Грунт-эмаль по ржавчине 3в1 "ПРОСТОКРАШЕНО!" белая', 0.8),
    ('Расцвет', 'Эмаль ПФ-115 белая. ТЗ "РАСЦВЕТ"', 0.9),
    ('Ореол', 'Грунт-концентрат влагоизолирующий водно-дисперсионный для наружных и внутренних работ "ОРЕОЛ"', 1),
]
for sheet, name, fas in checks:
    print(f"\n  Лист: {sheet} | Фасовка: {fas}")
    print(f"  Имя: {name}")
    old_match = next((p for p in old_prices[sheet] if p['name'] == name and p['fасовка'] == fas), None)
    new_match = next((p for p in new_prices[sheet] if p['name'] == name and p['fасовка'] == fas), None)
    print(f"  Старый (21.08): отгруз={old_match['отгрузочная'] if old_match else 'N/A'} | базовая={old_match['базовая'] if old_match else 'N/A'}")
    print(f"  Новый  (16.10): отгруз={new_match['отгрузочная'] if new_match else 'N/A'} | базовая={new_match['базовая'] if new_match else 'N/A'}")
    if old_match and new_match:
        diff = new_match['отгрузочная'] - old_match['отгрузочная']
        pct = (diff / old_match['отгрузочная']) * 100 if old_match['отгрузочная'] else 0
        print(f"  Изменение: {diff:+.2f} руб ({pct:+.2f}%)")

# Подсчёт изменённых цен
print("\n\n=== Анализ изменений цен ===")
changed = 0
unchanged = 0
not_found_new = 0
for sn in old_prices:
    for op in old_prices[sn]:
        # Найти соответствующую позицию в новом прайс-листе
        np_match = next((p for p in new_prices[sn] if p['name'] == op['name'] and p['fасовка'] == op['fасовка']), None)
        if np_match is None:
            not_found_new += 1
        elif abs(np_match['отгрузочная'] - op['отгрузочная']) > 0.01:
            changed += 1
        else:
            unchanged += 1
print(f"Без изменений: {unchanged}")
print(f"Изменились:    {changed}")
print(f"Не найдено в новом прайс-листе: {not_found_new}")
