"""
Найти замену для NoName SKU (гр.код 63180) — топ-новинку с чётким брендом, не входящую в ассортимент.
"""
import openpyxl
from openpyxl import load_workbook
import json

# Загрузить ассортимент v4
wb1 = load_workbook("/home/z/my-project/upload/Ассортимент_Олимп_300_SKU_финал_v4.xlsx", data_only=True)
ws1 = wb1["300 SKU — детально"]

# Собрать существующие гр.коды
existing_codes = set()
existing_skus = []
for row in ws1.iter_rows(min_row=4, max_row=ws1.max_row, values_only=True):
    if row[0] is None:
        continue
    code = row[7]
    if code:
        existing_codes.add(int(code))
    existing_skus.append({
        'idx': row[0], 'sphere': row[1], 'purpose': row[2], 'base': row[3],
        'brand': row[4], 'type': row[5], 'name': row[6], 'grp_code': code, 'abc': row[8]
    })

print(f"Ассортимент v4: {len(existing_skus)} SKU, {len(existing_codes)} уникальных гр.кодов")

# Загрузить новинки
wb2 = load_workbook("/home/z/my-project/upload/Новинки за 2026 год первичка и вторичка.xlsx", data_only=True)
ws2 = wb2["Первичные и вторичные продажи"]

novinki = []
for row in ws2.iter_rows(min_row=3, max_row=ws2.max_row, values_only=True):
    if row[3] is None or not isinstance(row[3], int):
        continue
    novinki.append({
        'code': row[3], 'brand': row[2], 'name': row[4], 'type': row[7],
        'vtorichka_rub': row[11] or 0,
    })

print(f"\nНовинок: {len(novinki)}")

# Топ новинок, которых НЕТ в ассортименте v4 (по гр.коду)
not_in_assortment = [n for n in novinki if n['code'] not in existing_codes]
not_in_assortment.sort(key=lambda x: -x['vtorichka_rub'])

print(f"\nНовинок нет в ассортименте v4: {len(not_in_assortment)}")
print("\nТоп-10 кандидатов на замену NoName (по убыванию вторички):")
print(f"{'Код':<7} {'Бренд':<14} {'Вторичка':<10} {'Номенклатура'}")
for n in not_in_assortment[:10]:
    print(f"{n['code']:<7} {n['brand']:<14} {n['vtorichka_rub']:<10} {n['name'][:55]}")

# NoName SKU, которую нужно заменить
print("\n\n=== NoName SKU, которую нужно заменить ===")
noname_sku = next(s for s in existing_skus if s['brand'] == 'NoName')
print(f"  idx={noname_sku['idx']} | гр.код={noname_sku['grp_code']} | бренд={noname_sku['brand']}")
print(f"  Имя: {noname_sku['name']}")
print(f"  Сфера: {noname_sku['sphere']} | Назначение: {noname_sku['purpose']} | Тип: {noname_sku['type']}")

# Подходящие замены: PROFIWOOD (для согласованности с категорией "Защита и декор дерева")
print("\n\n=== Лучшие кандидаты для замены ===")
print("(критерий: чёткий бренд, не в ассортименте, желательно та же категория — Дерево)")

# Покажем топ-3 PROFIWOOD новинок, которых нет в ассортименте
profiwood_new = [n for n in not_in_assortment if n['brand'] == 'PROFIWOOD']
print(f"\nPROFIWOOD новинки, не в ассортименте ({len(profiwood_new)}):")
for n in profiwood_new[:5]:
    print(f"  Код {n['code']} | {n['name'][:60]} | втор={n['vtorichka_rub']} руб")

# Найти цену для топ-кандидата
print("\n\n=== Поиск цены для лучшего кандидата в прайс-листе ===")
import re

wb3 = load_workbook("/home/z/my-project/upload/16.10.2026 РОССИЯ с НДС, руб..xlsx", data_only=True)
price_dict = {}
for sn in wb3.sheetnames:
    ws = wb3[sn]
    current_name = None
    for r in range(14, ws.max_row + 1):
        f_val = ws.cell(row=r, column=6).value
        h_val = ws.cell(row=r, column=8).value
        m_val = ws.cell(row=r, column=13).value
        n_val = ws.cell(row=r, column=14).value
        if f_val and isinstance(f_val, str) \
            and not f_val.startswith('СТО') and not f_val.startswith('ТУ') and not f_val.startswith('ГОСТ'):
            current_name = f_val.strip()
        if (m_val and isinstance(m_val, (int, float))) or (n_val and isinstance(n_val, (int, float))):
            price_dict.setdefault(sn, []).append({
                'name': current_name, 'fasovka': h_val,
                'базовая': m_val, 'отгрузочная': n_val,
            })

# Найти цену для топ-3 кандидатов
top_candidates = profiwood_new[:3]
for cand in top_candidates:
    print(f"\n--- Кандидат: {cand['name']} (гр.код {cand['code']}, вторичка {cand['vtorichka_rub']} руб) ---")
    # Найти в листе PROFIWOOD
    matches = []
    for p in price_dict.get('PROFIWOOD', []):
        if p['name'] and cand['name'][:30].lower() in p['name'].lower():
            matches.append(p)
    
    if matches:
        for p in matches[:5]:
            print(f"  ✓ '{p['name']}' | фас={p['fasovka']} | отгруз={p['отгрузочная']} | базовая={p['базовая']}")
    else:
        # Попробуем по ключевым словам
        keywords = [w for w in cand['name'].split() if len(w) > 4][:3]
        for p in price_dict.get('PROFIWOOD', []):
            if p['name'] and all(kw.lower() in p['name'].lower() for kw in keywords[:2]):
                matches.append(p)
        if matches:
            print(f"  По ключевым словам {keywords}:")
            for p in matches[:5]:
                print(f"    '{p['name']}' | фас={p['fasovka']} | отгруз={p['отгрузочная']}")
        else:
            print(f"  Не найдено")

# Сохранить данные для замены
import json
with open('/home/z/my-project/scripts/replacement_data.json', 'w', encoding='utf-8') as f:
    json.dump({
        'noname_sku': noname_sku,
        'candidates': profiwood_new[:5],
    }, f, ensure_ascii=False, indent=2, default=str)
