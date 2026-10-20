"""
Финальный скрипт: сопоставление товаров из ассортимента с прайс-листом
и генерация Excel-файла с колонкой "Отгрузочная цена, руб".
"""
import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, NamedStyle
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule
from collections import defaultdict
import re
import shutil
import os

# === ПУТИ ===
PRICE_PATH = "/home/z/my-project/upload/21.08.2026 РОССИЯ с НДС, руб..xlsx"
ASSORT_PATH = "/home/z/my-project/upload/Ассортимент_Олимп_300_SKU.xlsx"
OUT_PATH = "/home/z/my-project/download/Ассортимент_Олимп_300_SKU_с_отгрузочными_ценами.xlsx"

# === 1. Чтение прайс-листа ===
print("Чтение прайс-листа...")
wb2 = load_workbook(PRICE_PATH, data_only=True)

price_dict = defaultdict(list)
for sn in wb2.sheetnames:
    ws = wb2[sn]
    current_name = None
    for row_idx, row in enumerate(ws.iter_rows(min_row=14, max_row=ws.max_row, values_only=True), 14):
        f_val = row[5] if len(row) > 5 else None
        m_val = row[12] if len(row) > 12 else None
        n_val = row[13] if len(row) > 13 else None
        h_val = row[7] if len(row) > 7 else None
        if f_val is not None and isinstance(f_val, str) \
            and not f_val.startswith('СТО') \
            and not f_val.startswith('ТУ') \
            and not f_val.startswith('ГОСТ'):
            current_name = f_val.strip()
        if (m_val is not None and isinstance(m_val, (int, float))) or (n_val is not None and isinstance(n_val, (int, float))):
            price_dict[sn].append({
                'name': current_name,
                'fасовка': h_val,
                'базовая': m_val,
                'отгрузочная': n_val,
                'sheet': sn,
            })
print(f"  Загружено {sum(len(v) for v in price_dict.values())} ценовых позиций")

# === 2. Утилиты нормализации ===
QUOTE_RE = re.compile(r'[«»""„"\']')
# ИСПРАВЛЕНО: \s* вместо \s+ (чтобы «0.75 л 0.9 кг» очищалось полностью)
FASOVKA_RE = re.compile(r'\s*\d+(?:[.,]\d+)?\s*(?:кг|л|l)\s*', re.IGNORECASE)
# Штуки (для PROFIWOOD "6 шт.")
SHT_RE = re.compile(r'\s*\d+\s*шт\.?\s*', re.IGNORECASE)

def strip_quotes(s):
    return QUOTE_RE.sub(' ', s)

def norm_for_match(name):
    if name is None:
        return ''
    s = strip_quotes(str(name))
    s = FASOVKA_RE.sub(' ', s)
    s = SHT_RE.sub(' ', s)
    # Заменить пунктуацию на пробелы
    s = re.sub(r'[.,;:!?()\[\]{}]', ' ', s)
    # Дефисы — на пробелы (для токенизации лак-антисептик → лак антисептик)
    s = re.sub(r'-', ' ', s)
    s = s.lower()
    s = re.sub(r'\s+', ' ', s)
    return s.strip()

def tokenize(s):
    tokens = set(re.findall(r'[a-zа-яё]+|\d+(?:\.\d+)?', s))
    return tokens

# «Шумовые» токены, не влияющие на идентичность товара
FILLER_TOKENS = {
    # Предлоги/союзы
    'с', 'в', 'и', 'по', 'для', 'на', 'из', 'от', 'до', 'или', 'со', 'об', 'о',
    # Стандарты
    'гост', 'тз', 'сто', 'ту',
    # Акронимы и сокращения
    'бау', 'шт', 'м',  # М = матовая (иногда)
    # Аббревиатуры типов ЛКМ
    'вд', 'ва', 'ак', 'пф', 'гф', 'нц', 'вда', 'вдаа',
    # Тех. характеристики ЛКМ
    'водно', 'дисперсионная', 'дисперсионный', 'полиакриловая', 'полиакриловый',
    'латексная', 'латексный', 'силиконовая', 'силиконовый', 'акриловая', 'акриловый',
    'алкидная', 'алкидный', 'алкидно', 'стирольная', 'стирольный',
    'уретановый', 'уретановая',
    # Дескрипторы свойств
    'универсальная', 'универсальный', 'атмосферостойкий', 'атмосферостойкая',
    'влагостойкая', 'влагостойкий', 'износостойкий', 'износостойкая',
    'быстросохнущая', 'быстросохнущий', 'моющаяся',
    'матовая', 'глянцевая', 'полуматовая', 'глубокоматовая', 'суперстойкая',
    'интерьерная', 'фасадная', 'термостойкая',
    # Маркетинговые слова
    'супербелая',
    # Единицы измерения
    'кг', 'л', 'мл', 'мм',
    # Базовые коды (однобуквенные) — могут пересекаться с "База А", "База С"
    'а', 'с',
}

# Распознавание цвета (для требования совпадения цвета)
COLOR_RE = re.compile(
    r'бел\w*|черн\w*|чёрн\w*|красн\w*|сер\w*|зелён\w*|зелен\w*|син\w*|голуб\w*|'
    r'жёлт\w*|желт\w*|коричнев\w*|вишнёв\w*|вишнев\w*|шоколад\w*|бежев\w*|'
    r'серебрист\w*|золотист\w*|медн\w*|графитов\w*|оранжев\w*|салатов\w*|'
    r'бирюзов\w*|васильков\w*|ультрабел\w*|скандинавск\w*|горьк\w*|'
    r'тёмно\w*|темно\w*|светло\w*|ярко\w*|красно\w*|'
    r'слонов\w*|кость\w*|бесцветн\w*|'
    r'ультра\w*|супер\w*',
    re.IGNORECASE
)

def extract_colors(s):
    if not s:
        return set()
    matches = COLOR_RE.findall(s)
    return set(m.lower() for m in matches if m)

# Извлечение фасовки из имени в ассортименте
KG_RE = re.compile(r'(\d+(?:[.,]\d+)?)\s*кг', re.IGNORECASE)
L_RE = re.compile(r'(\d+(?:[.,]\d+)?)\s*[лl]', re.IGNORECASE)

def extract_fasovka(name):
    if name is None:
        return (None, None)
    s = str(name).replace(',', '.')
    kg_m = KG_RE.search(s)
    l_m = L_RE.search(s)
    kg = float(kg_m.group(1)) if kg_m else None
    l = float(l_m.group(1)) if l_m else None
    return (kg, l)

# === 3. Соответствие брендов листам прайс-листа ===
brand_to_sheets = {
    'AvantGarde':         ['AvantGarde'],
    'NEWTONE':            ['NEWTONE'],
    'PROFIWOOD':          ['PROFIWOOD'],
    'ОРЕОЛ TITAN':        ['ОРЕОЛ TITAN', 'ОРЕОЛ TITAN(л)'],
    'Ореол':              ['Ореол'],
    'Ореол Дисконт':      ['Ореол Дисконт', 'Ореол'],   # Ищем и в "Ореол"
    'Простокрашено! БАУ': ['Простокрашено! БАУ'],
    'Расцвет':            ['Расцвет'],
}

# === 4. Функция сопоставления ===
def match_price(item_brand, item_name):
    sheets = brand_to_sheets.get(item_brand, [])
    if not sheets:
        return None
    
    name_norm = norm_for_match(item_name)
    if not name_norm:
        return None
    a_tokens = tokenize(name_norm)
    a_tokens_clean = a_tokens - FILLER_TOKENS
    a_colors = extract_colors(name_norm)
    
    if not a_tokens_clean:
        return None
    
    kg, l = extract_fasovka(item_name)
    
    candidates = []
    for sn in sheets:
        for p in price_dict[sn]:
            if p['name'] is None:
                continue
            pn = norm_for_match(p['name'])
            if not pn:
                continue
            b_tokens = tokenize(pn)
            b_tokens_clean = b_tokens - FILLER_TOKENS
            b_colors = extract_colors(pn)
            
            if not b_tokens_clean:
                continue
            
            # Проверка цвета: если оба имени содержат цвета, они должны пересекаться
            color_ok = True
            if a_colors and b_colors:
                if not a_colors.issubset(b_colors) and not b_colors.issubset(a_colors):
                    overlap = len(a_colors & b_colors)
                    if overlap == 0:
                        color_ok = False
            
            if not color_ok:
                continue
            
            # Подмножество (subset)
            a_in_b = a_tokens_clean.issubset(b_tokens_clean)
            b_in_a = b_tokens_clean.issubset(a_tokens_clean)
            intersection = len(a_tokens_clean & b_tokens_clean)
            union = len(a_tokens_clean | b_tokens_clean)
            jaccard = intersection / union if union else 0
            
            # Точное совпадение
            if pn == name_norm:
                candidates.append((sn, p, 'exact', 1.0))
                continue
            
            # Подмножество (subset)
            if (a_in_b or b_in_a) and jaccard >= 0.5:
                candidates.append((sn, p, 'subset', jaccard))
                continue
            
            # Подстрока
            if name_norm in pn or pn in name_norm:
                candidates.append((sn, p, 'substr', 0.7))
                continue
            
            # Высокий Jaccard
            if jaccard >= 0.7:
                candidates.append((sn, p, 'jaccard', jaccard))
                continue
    
    if not candidates:
        return None
    
    # Сортируем по качеству
    candidates.sort(key=lambda c: ({'exact': 0, 'subset': 1, 'substr': 2, 'jaccard': 3}[c[2]], -c[3]))
    
    top_mt = candidates[0][2]
    top_ratio = candidates[0][3]
    same_quality = [c for c in candidates if c[2] == top_mt and abs(c[3] - top_ratio) < 0.05]
    
    # Если в ассортименте есть фасовка — ищем ближайшую
    if kg is not None or l is not None:
        target_values = []
        if kg is not None:
            target_values.append(kg)
        if l is not None:
            target_values.append(l)
        
        best = None
        best_diff = float('inf')
        for sn, p, mt, ratio in same_quality:
            f = p['fасовка']
            if f is None or not isinstance(f, (int, float)):
                continue
            for tv in target_values:
                diff = abs(f - tv)
                if diff < best_diff:
                    best_diff = diff
                    best = (sn, p, mt, ratio)
        
        if best is not None and best_diff <= 0.5:
            return {
                'базовая': best[1]['базовая'],
                'отгрузочная': best[1]['отгрузочная'],
                'фасовка_прайс': best[1]['fасовка'],
                'лист': best[0],
                'ценовое_имя': best[1]['name'],
                'match_type': best[2],
                'tolerance': best_diff,
            }
    
    # Без фасовки или без совпадения фасовки — берём наименьшую
    with_fas = [c for c in same_quality if isinstance(c[1]['fасовка'], (int, float))]
    if with_fas:
        with_fas.sort(key=lambda c: c[1]['fасовка'])
        sn, p, mt, ratio = with_fas[0]
        return {
            'базовая': p['базовая'],
            'отгрузочная': p['отгрузочная'],
            'фасовка_прайс': p['fасовка'],
            'лист': sn,
            'ценовое_имя': p['name'],
            'match_type': mt + '_мин',
            'tolerance': None,
        }
    
    sn, p, mt, ratio = candidates[0]
    return {
        'базовая': p['базовая'],
        'отгрузочная': p['отгрузочная'],
        'фасовка_прайс': p['fасовка'],
        'лист': sn,
        'ценовое_имя': p['name'],
        'match_type': mt + '_первое',
        'tolerance': None,
    }

# === 5. Чтение ассортимента и сопоставление ===
print("\nЧтение ассортимента и сопоставление...")
wb1 = load_workbook(ASSORT_PATH, data_only=True)
ws1 = wb1['300 SKU — детально']

results = []
matched = 0
unmatched = []
match_types = defaultdict(int)

for row in ws1.iter_rows(min_row=4, max_row=ws1.max_row, values_only=True):
    if row[0] is None:
        continue
    item = {
        'idx': row[0], 'sphere': row[1], 'purpose': row[2], 'base': row[3],
        'brand': row[4], 'type': row[5], 'name': row[6], 'grp_code': row[7], 'abc': row[8]
    }
    match = match_price(item['brand'], item['name'])
    if match:
        item.update(match)
        matched += 1
        match_types[match['match_type']] += 1
    else:
        unmatched.append(item)
    results.append(item)

print(f"\nСопоставлено: {matched} / {len(results)} ({matched*100//len(results)}%)")
print(f"Не сопоставлено: {len(unmatched)}")
print(f"\nТипы сопоставления:")
for mt, cnt in sorted(match_types.items(), key=lambda x: -x[1]):
    print(f"  {mt}: {cnt}")

print("\nНе сопоставлено:")
for item in unmatched:
    print(f"  [{item['idx']}] {item['brand']} | {item['name']}")

# Сохраним результаты в JSON для последующего использования
import json
def _ser(v):
    return round(v, 4) if isinstance(v, float) else v
with open('/home/z/my-project/scripts/results.json', 'w', encoding='utf-8') as f:
    out = [{k: _ser(v) for k, v in item.items()} for item in results]
    json.dump(out, f, ensure_ascii=False, indent=2, default=str)
print(f"\nСохранено в results.json")
