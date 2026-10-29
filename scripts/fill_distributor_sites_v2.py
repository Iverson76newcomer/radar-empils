#!/usr/bin/env python3
"""
Заполнить колонку "Сайт" (E) в Distributors_2026.xlsx на основе 4 файлов коллег.
Учитывает merged cells: для каждого объединённого диапазона E записывает URL в верхнюю ячейку.
Для нераспознанных — веб-поиск.
"""

import os
import re
import shutil
import json
import time
from urllib.parse import urlparse
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

SRC = "/home/z/my-project/upload/Distributors_2026.xlsx"
DST = "/home/z/my-project/download/Distributors_2026_с_сайтами.xlsx"
shutil.copy(SRC, DST)

# === Утилиты ===

def clean_url(url):
    """Очистить URL: убрать лишние пробелы, особенно внутри 'http:// www.'"""
    if not url:
        return None
    s = str(url).strip()
    if not s:
        return None
    # Удалить пробелы после http:// и https://
    s = re.sub(r'(https?://)\s+', r'\1', s, flags=re.IGNORECASE)
    # Удалить другие внутренние пробелы
    s = re.sub(r'\s+', '', s)
    return s

def extract_base_url(url):
    """Извлечь базовый URL: scheme + host."""
    if not url:
        return None
    s = clean_url(url)
    if not s:
        return None
    if not s.startswith(('http://', 'https://')):
        s = 'http://' + s
    try:
        parsed = urlparse(s)
        if not parsed.netloc:
            return None
        return f"{parsed.scheme}://{parsed.netloc}"
    except Exception:
        return None

def normalize(name):
    """Нормализовать имя: lower, без лишних пробелов и кавычек."""
    if not name:
        return ""
    s = str(name).lower().strip()
    s = re.sub(r'\s+', ' ', s.replace('\xa0', ' '))
    s = re.sub(r'["«»\'\"„""]+', '', s)
    s = s.rstrip(' ,.')
    return s

def normalize_city(city):
    """Нормализовать город: lower, без 'г.'."""
    if not city:
        return ""
    s = str(city).lower().strip()
    s = re.sub(r'^г\.\s*', '', s)
    s = re.sub(r'\s+', ' ', s)
    return s.strip()

# === 1. Строим маппинг (имя, город) → base_url из 4 файлов ===

# mapping[(norm_name, norm_city)] = base_url
# также mapping_by_name[norm_name] = list of (city, url) для дубликатов
mapping = {}
mapping_by_name = {}

def add_mapping(name, url, city=None):
    if not name or not url:
        return
    base = extract_base_url(url)
    if not base:
        return
    norm = normalize(name)
    if not norm:
        return
    norm_city = normalize_city(city) if city else None
    key = (norm, norm_city)
    if key not in mapping:
        mapping[key] = base
    # Также в by_name
    if norm not in mapping_by_name:
        mapping_by_name[norm] = []
    entry = (norm_city, base)
    if entry not in mapping_by_name[norm]:
        mapping_by_name[norm].append(entry)

# Файл 1: ДВ и ВС
wb1 = load_workbook("/home/z/my-project/upload/Сайты дистрибьюторов ДВ и ВС_2024.xlsx", data_only=True)
ws1 = wb1["Лист1"]
for row in ws1.iter_rows(min_row=3, values_only=True):
    region, num, legal_name, comm_name, site, ftm, city = row[:7]
    # Колонка G (index 6) — город, но он может быть как "Улан-Удэ, Чита"
    if comm_name and site:
        if city:
            for c in str(city).split(','):
                add_mapping(comm_name, site, c.strip())
        else:
            add_mapping(comm_name, site, None)
    if legal_name and site:
        if city:
            for c in str(city).split(','):
                add_mapping(legal_name, site, c.strip())
        else:
            add_mapping(legal_name, site, None)

# Файлы 2-4: с "fill-down" имени
files_234 = [
    "/home/z/my-project/upload/Ссылки на сайты дистрибьюторов Москва, Центр, Северо-Запад.xlsx",
    "/home/z/my-project/upload/Ссылки на сайты дистрибьюторов Поволжье Юг.xlsx",
    "/home/z/my-project/upload/Ссылки на сайты дистрибьюторов Урал-Западная Сибирь.xlsx",
]

for fpath in files_234:
    wb = load_workbook(fpath, data_only=True)
    ws = wb["Лист1"]
    current_name = None
    current_city = None

    for row in ws.iter_rows(min_row=3, values_only=True):
        sbu = row[1] if len(row) > 1 else None
        name = row[2] if len(row) > 2 else None
        city = row[3] if len(row) > 3 else None
        tm = row[4] if len(row) > 4 else None
        link = row[5] if len(row) > 5 else None

        if name and str(name).strip():
            current_name = str(name).strip()
            current_city = str(city).strip() if city else None
        if link and str(link).strip() and current_name:
            add_mapping(current_name, link, current_city)

print(f"=== Построен маппинг: {len(mapping)} записей (имя+город) ===")
print(f"=== Уникальных имён: {len(mapping_by_name)} ===\n")

# === 2. Открываем основной файл ===

wb_main = load_workbook(DST)
ws_main = wb_main["Лист1"]

# Соберём все merged ranges для E
merged_E_ranges = []
for mr in ws_main.merged_cells.ranges:
    if mr.min_col == 5 and mr.max_col == 5:
        merged_E_ranges.append((mr.min_row, mr.max_row))
merged_E_ranges.sort()

# Также найдём merged D ranges для синхронизации
merged_D_ranges = []
for mr in ws_main.merged_cells.ranges:
    if mr.min_col == 4 and mr.max_col == 4:
        merged_D_ranges.append((mr.min_row, mr.max_row))
merged_D_ranges.sort()

# Соберём список всех клиентов (для merged D - одна запись, для non-merged - построчно)
clients_to_process = []

# Сначала merged ranges
for d_min, d_max in merged_D_ranges:
    client = ws_main.cell(row=d_min, column=4).value
    if client:
        clients_to_process.append({
            'row': d_min,
            'row_end': d_max,
            'client': str(client).strip(),
            'merged': True,
        })

# Теперь для строк без merged D — добавим каждую отдельно
# Но также проверим, что строка не входит в merged E range (там сайт уже общий для группы)
merged_E_rows = set()
for e_min, e_max in merged_E_ranges:
    for r in range(e_min, e_max + 1):
        merged_E_rows.add(r)

for row_idx in range(2, ws_main.max_row + 1):
    if row_idx in merged_E_rows:
        continue  # эта строка — часть merged E, пропускаем
    client = ws_main.cell(row=row_idx, column=4).value
    if client and str(client).strip():
        # Проверим, не входит ли эта строка в merged D
        in_merged_D = False
        for d_min, d_max in merged_D_ranges:
            if d_min <= row_idx <= d_max and d_min != row_idx:
                in_merged_D = True
                break
        if not in_merged_D:
            clients_to_process.append({
                'row': row_idx,
                'row_end': row_idx,
                'client': str(client).strip(),
                'merged': False,
            })

print(f"=== Клиентов к обработке: {len(clients_to_process)} ===")
print(f"   Из них в merged ranges: {sum(1 for c in clients_to_process if c['merged'])}")
print(f"   Отдельных строк:       {sum(1 for c in clients_to_process if not c['merged'])}\n")

# === 3. Извлекаем город из имени клиента (если есть в скобках) ===

def extract_city_from_name(name):
    """Извлечь город из имени вида 'Альянс ООО (Иркутск)'."""
    m = re.search(r'\(([^)]+)\)', name)
    if m:
        content = m.group(1).strip()
        # Если содержит ООО/ИП и т.п., не берём как город
        if re.search(r'\b(ооо|ип|ао|зао|ю-трейд|цитадель|бывший|бывшая)\b', content.lower()):
            return None
        return content
    return None

# === 4. Функция поиска сайта в маппинге ===

def find_site(client_name):
    """Возвращает (url, source) — url и источник ('exact'/'fuzzy'/'none')."""
    norm = normalize(client_name)
    # Извлечь город из скобок в имени
    city_in_name = extract_city_from_name(client_name)
    norm_city_in_name = normalize_city(city_in_name) if city_in_name else None

    # Точное совпадение по имени+городу
    if norm_city_in_name:
        key = (norm, norm_city_in_name)
        if key in mapping:
            return mapping[key], 'exact'

    # Точное совпадение только по имени
    if norm in mapping_by_name:
        entries = mapping_by_name[norm]
        # Если есть запись без города — возьмём её
        for c, u in entries:
            if c is None:
                return u, 'exact'
        # Иначе берём первую
        return entries[0][1], 'exact'

    # Fuzzy: ищем ключи, в которых есть ключевые токены клиента
    # Удаляем OPF и общие слова
    cleaned = re.sub(r'\b(ооо|оао|зао|ип|ао|пао|сеть|компания|торговый дом|тд|бывший|бывшая)\b\s*', '', norm).strip()
    cleaned = re.sub(r'\([^)]*\)', '', cleaned).strip()  # удаляем всё в скобках
    cleaned = re.sub(r'[^a-zа-яё0-9\s-]', '', cleaned)
    tokens = set(t for t in re.split(r'[\s\-]+', cleaned) if len(t) >= 4)

    if tokens:
        best_score = 0
        best_url = None
        for key_name, entries in mapping_by_name.items():
            key_cleaned = re.sub(r'\b(ооо|оао|зао|ип|ао|пао|сеть|компания|торговый дом|тд)\b\s*', '', key_name).strip()
            key_cleaned = re.sub(r'\([^)]*\)', '', key_cleaned).strip()
            key_cleaned = re.sub(r'[^a-zа-яё0-9\s-]', '', key_cleaned)
            key_tokens = set(t for t in re.split(r'[\s\-]+', key_cleaned) if len(t) >= 4)
            common = tokens & key_tokens
            if len(common) > best_score:
                best_score = len(common)
                best_url = entries[0][1]
        if best_score >= 2 and best_url:  # минимум 2 общих токена
            return best_url, 'fuzzy'

    return None, 'none'

# === 5. Обрабатываем клиентов ===

# Стили
fill_exact = PatternFill("solid", fgColor="C8E6C9")   # зелёный — точное совпадение
fill_fuzzy = PatternFill("solid", fgColor="FFE0B2")   # оранжевый — fuzzy
fill_search = PatternFill("solid", fgColor="BBDEFB")   # голубой — web search
fill_unmatched = PatternFill("solid", fgColor="FFCDD2") # красный — не найден

font_url = Font(name="Calibri", size=10, color="0563C1", underline="single")

matched_exact = []
matched_fuzzy = []
unmatched = []

for c in clients_to_process:
    url, source = find_site(c['client'])
    if url:
        cell = ws_main.cell(row=c['row'], column=5)
        cell.value = url
        cell.font = font_url
        cell.fill = fill_exact if source == 'exact' else fill_fuzzy
        cell.alignment = Alignment(horizontal="left", vertical="center")
        if source == 'exact':
            matched_exact.append((c['row'], c['client'], url))
        else:
            matched_fuzzy.append((c['row'], c['client'], url))
    else:
        unmatched.append((c['row'], c['client']))

print(f"=== ЭТАП 1: Заполнение из файлов коллег ===")
print(f"  ✅ Точное совпадение:    {len(matched_exact)}")
print(f"  ⚠️  Fuzzy совпадение:    {len(matched_fuzzy)}")
print(f"  ❌ Не найдено в файлах:  {len(unmatched)}")
print(f"     (потребуется web-поиск)\n")

# Сохраняем промежуточно
wb_main.save(DST)

# Выводим fuzzy совпадения для ручной проверки
if matched_fuzzy:
    print("=== Fuzzy совпадения (требуют проверки) ===")
    for row, name, url in matched_fuzzy:
        print(f"  R{row}: {name} → {url}")
    print()

print(f"=== НЕ НАЙДЕННЫЕ (для web-поиска) ===")
for row, name in unmatched:
    print(f"  R{row}: {name}")

# Сохраним список unmatched для использования в следующем скрипте
with open("/tmp/unmatched_distributors.json", "w", encoding="utf-8") as f:
    json.dump(unmatched, f, ensure_ascii=False, indent=2)

print(f"\n✅ Этап 1 завершен. Файл сохранён: {DST}")
print(f"   Размер: {os.path.getsize(DST):,} байт")
