"""Сравнить партнёров из файла Москвы с общей таблицей."""
import openpyxl
from openpyxl import load_workbook

# === 1. Загрузить общую таблицу ===
wb_general = load_workbook("/home/z/my-project/upload/Belye_pyatna_obschaya_tablitsa_21_09_2026.xlsx", data_only=True)
ws_general = wb_general["Белые пятна"]

headers_general = [ws_general.cell(row=1, column=c).value for c in range(1, ws_general.max_column + 1)]

# Прочитать всех партнёров из общей таблицы
general_partners = []
for r in range(2, ws_general.max_row + 1):
    client_name = ws_general.cell(row=r, column=4).value  # Колонка "Клиент"
    inn = ws_general.cell(row=r, column=5).value  # ИНН
    city = ws_general.cell(row=r, column=3).value  # Город
    region = ws_general.cell(row=r, column=2).value  # Федеральный округ
    if client_name and str(client_name).strip():
        general_partners.append({
            'row': r,
            'name': str(client_name).strip(),
            'inn': inn,
            'city': city,
            'region': region,
        })

print(f"Всего партнёров в общей таблице: {len(general_partners)}")

# === 2. Загрузить файл Москвы — извлечь список дистрибьюторов ===
wb_msk = load_workbook("/home/z/my-project/upload/Белые пятна - Москва.xlsx", data_only=True)
ws_msk = wb_msk["Лист1"]

# Структура файла: строки с названием дистрибьютора в колонке C (3), 
# портфель конкурентов в колонке D (4), территория в E (5), ТО в F (6)
# Под каждым дистрибьютором — строки с дополнительными брендами (колонка D)

msk_distributors = []
current_distributor = None
current_brands = []

for r in range(2, ws_msk.max_row + 1):
    col_c = ws_msk.cell(row=r, column=3).value  # Дистрибьютор
    col_d = ws_msk.cell(row=r, column=4).value   # Портфель конкурентов
    col_e = ws_msk.cell(row=r, column=5).value   # Территория
    col_f = ws_msk.cell(row=r, column=6).value   # ТО, млн. руб.
    col_g = ws_msk.cell(row=r, column=7).value   # Контактное лицо
    
    if col_c and str(col_c).strip():
        # New distributor
        if current_distributor:
            msk_distributors.append({
                'name': current_distributor,
                'brands': current_brands,
                'territory': current_territory,
                'to': current_to,
                'contact': current_contact,
            })
        current_distributor = str(col_c).strip()
        current_brands = []
        if col_d and str(col_d).strip():
            current_brands.append(str(col_d).strip())
        current_territory = str(col_e).strip() if col_e else ''
        current_to = col_f
        current_contact = str(col_g).strip() if col_g else ''
    elif col_d and str(col_d).strip():
        # Additional brand for current distributor
        current_brands.append(str(col_d).strip())

# Don't forget the last one
if current_distributor:
    msk_distributors.append({
        'name': current_distributor,
        'brands': current_brands,
        'territory': current_territory,
        'to': current_to,
        'contact': current_contact,
    })

print(f"\nДистрибьюторов в файле Москвы: {len(msk_distributors)}")
print("\n=== Дистрибьюторы из файла Москвы ===")
for i, d in enumerate(msk_distributors, 1):
    print(f"\n{i}. {d['name']}")
    print(f"   Территория: {d['territory']}")
    print(f"   ТО: {d['to']} млн руб")
    print(f"   Портфель: {', '.join(d['brands'])}")

# === 3. Сравнение: есть ли каждый дистрибьютор из Москвы в общей таблице ===
print("\n\n" + "=" * 100)
print("СРАВНЕНИЕ: ЕСТЬ ЛИ ДИСТРИБЬЮТОРЫ ИЗ ФАЙЛА МОСКВЫ В ОБЩЕЙ ТАБЛИЦЕ")
print("=" * 100)

def normalize(s):
    """Нормализовать строку для сравнения."""
    if s is None:
        return ''
    s = str(s).upper().strip()
    # Убрать кавычки, ООО, лишние пробелы
    import re
    s = re.sub(r'[«»""„"\']', '', s)
    s = re.sub(r'\bООО\b', '', s)
    s = re.sub(r'\s+', ' ', s).strip()
    return s

found_count = 0
not_found_count = 0

for d in msk_distributors:
    msk_name_norm = normalize(d['name'])
    
    # Ищем в общей таблице
    found = None
    for g in general_partners:
        g_name_norm = normalize(g['name'])
        if msk_name_norm == g_name_norm or msk_name_norm in g_name_norm or g_name_norm in msk_name_norm:
            found = g
            break
    
    if found:
        found_count += 1
        print(f"\n✓ НАЙДЕН: {d['name']}")
        print(f"   В файле Москвы: ТО={d['to']} млн руб, территория={d['territory']}")
        print(f"   В общей таблице: строка {found['row']}, город={found['city']}, ИНН={found['inn']}")
    else:
        not_found_count += 1
        # Попробуем частичное совпадение
        partial = None
        for g in general_partners:
            g_name_norm = normalize(g['name'])
            # Проверить, совпадает ли хотя бы 3+ слова
            msk_words = set(msk_name_norm.split())
            g_words = set(g_name_norm.split())
            common = msk_words & g_words
            if len(common) >= 2 and any(len(w) > 3 for w in common):
                partial = g
                break
        
        if partial:
            print(f"\n⚠ ЧАСТИЧНОЕ СОВПАДЕНИЕ: {d['name']}")
            print(f"   В файле Москвы: ТО={d['to']} млн руб, территория={d['territory']}")
            print(f"   Похоже на: {partial['name']} (строка {partial['row']}, город={partial['city']})")
        else:
            print(f"\n✗ НЕ НАЙДЕН: {d['name']}")
            print(f"   ТО={d['to']} млн руб, территория={d['territory']}")
            print(f"   Портфель: {', '.join(d['brands'])}")

# === 4. Итог ===
print("\n\n" + "=" * 100)
print(f"ИТОГ: {found_count} из {len(msk_distributors)} найдены в общей таблице, {not_found_count} не найдены")
print("=" * 100)
