"""Дополнительный поиск — по частичным совпадениям и по ИНН."""
import openpyxl
from openpyxl import load_workbook
import re

# === 1. Загрузить общую таблицу ===
wb_general = load_workbook("/home/z/my-project/upload/Belye_pyatna_obschaya_tablitsa_21_09_2026.xlsx", data_only=True)
ws_general = wb_general["Белые пятна"]

general_partners = []
for r in range(2, ws_general.max_row + 1):
    client_name = ws_general.cell(row=r, column=4).value
    inn = ws_general.cell(row=r, column=5).value
    city = ws_general.cell(row=r, column=3).value
    region = ws_general.cell(row=r, column=2).value
    if client_name and str(client_name).strip():
        general_partners.append({
            'row': r,
            'name': str(client_name).strip(),
            'inn': inn,
            'city': city,
            'region': region,
        })

# === 2. Искомые компании из файла Москвы (без строки-заголовка) ===
missing_companies = [
    {'name': 'ООО "СТРОЙТЕРМИНАЛ ЦЕНТР КРАСОК"', 'short': 'СТРОЙТЕРМИНАЛ', 'territory': 'Московская обл.', 'to': 920},
    {'name': 'ООО "СТРОЙБАЗА - СПЕКТР"', 'short': 'СТРОЙБАЗА', 'territory': 'Липецкая обл.', 'to': 264},
    {'name': 'ООО Стройоптторг', 'short': 'Стройоптторг', 'territory': 'Вологодская обл.', 'to': 110},
    {'name': 'ООО "КОНТИНЕНТАЛЬ"', 'short': 'КОНТИНЕНТАЛЬ', 'territory': 'Воронежская обл.', 'to': 107},
    {'name': 'ООО "Аплайн"', 'short': 'Аплайн', 'territory': 'Вологодская обл.', 'to': 55},
]

# Также проверим "ВЕРТИКАЛЬ" из Московской обл. (ТО 1664 млн) — есть ли она отдельно?
extra_check = {'name': 'ООО "ВЕРТИКАЛЬ" (Московская обл.)', 'short': 'ВЕРТИКАЛЬ', 'territory': 'Московская обл.', 'to': 1664}

# === 3. Поиск по частичным совпадениям ===
print("=" * 100)
print("ПОИСК НЕНАЙДЕННЫХ КОМПАНИЙ В ОБЩЕЙ ТАБЛИЦЕ (по частичным совпадениям)")
print("=" * 100)

for missing in missing_companies + [extra_check]:
    short = missing['short'].upper()
    print(f"\n--- Ищем: {missing['name']} (короткое имя: {short}) ---")
    print(f"    Территория: {missing['territory']}, ТО: {missing['to']} млн руб")
    
    # Поиск по частичному совпадению
    matches = []
    for g in general_partners:
        g_name_upper = g['name'].upper()
        if short in g_name_upper:
            matches.append(g)
    
    if matches:
        print(f"  ✓ НАЙДЕНЫ совпадения:")
        for m in matches:
            print(f"    Строка {m['row']}: '{m['name']}' | город={m['city']} | ИНН={m['inn']}")
    else:
        # Поиск по части имени (без кавычек, ООО и т.д.)
        partial_matches = []
        for g in general_partners:
            g_name_clean = re.sub(r'[«»""„"\']', '', g['name'].upper())
            g_name_clean = re.sub(r'\bООО\b|\bОАО\b|\bЗАО\b', '', g_name_clean).strip()
            short_clean = re.sub(r'[«»""„"\']', '', short).strip()
            
            if short_clean in g_name_clean or g_name_clean in short_clean:
                partial_matches.append(g)
        
        if partial_matches:
            print(f"  ⚠ ЧАСТИЧНЫЕ совпадения:")
            for m in partial_matches:
                print(f"    Строка {m['row']}: '{m['name']}' | город={m['city']} | ИНН={m['inn']}")
        else:
            print(f"  ✗ НЕ НАЙДЕН в общей таблице")

# === 4. Все записи с "ВЕРТИКАЛЬ" в общей таблице ===
print("\n\n=== Все записи с 'ВЕРТИКАЛЬ' в общей таблице ===")
for g in general_partners:
    if 'ВЕРТИКАЛЬ' in g['name'].upper():
        print(f"  Строка {g['row']}: '{g['name']}' | город={g['city']} | ИНН={g['inn']}")

# === 5. Все записи с 'СТРОЙ' в общей таблице (для поиска Стройтерминал, Стройбаза, Стройоптторг) ===
print("\n=== Все записи с 'СТРОЙ' в общей таблице ===")
for g in general_partners:
    if 'СТРОЙ' in g['name'].upper():
        print(f"  Строка {g['row']}: '{g['name']}' | город={g['city']} | ИНН={g['inn']}")

# === 6. Все записи с 'АПЛАЙН', 'КОНТИНЕНТАЛЬ', 'АПЛАЙН' ===
print("\n=== Поиск 'АПЛАЙН', 'КОНТИНЕНТАЛЬ', 'APLAIN', 'CONTINENTAL' ===")
for g in general_partners:
    name_upper = g['name'].upper()
    if any(kw in name_upper for kw in ['АПЛАЙН', 'APLAIN', 'APLAYN', 'КОНТИНЕНТАЛЬ', 'CONTINENTAL', 'КОНТИНЕНТ']):
        print(f"  Строка {g['row']}: '{g['name']}' | город={g['city']} | ИНН={g['inn']}")

# === 7. Все партнёры из Москвы/Московской обл. в общей таблице ===
print("\n=== Все партнёры из Москвы/Московской обл. в общей таблице ===")
for g in general_partners:
    city_str = str(g['city']).upper() if g['city'] else ''
    region_str = str(g['region']).upper() if g['region'] else ''
    if 'МОСКВ' in city_str or 'МОСКВ' in region_str:
        print(f"  Строка {g['row']}: '{g['name']}' | город={g['city']} | регион={g['region']}")

