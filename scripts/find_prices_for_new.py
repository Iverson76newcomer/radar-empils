"""
Найти цены в прайс-листе (16.10.2026) для топ-5 кандидатов из ABC-анализа.
"""
import openpyxl
from openpyxl import load_workbook
from collections import defaultdict
import re

PRICE_PATH = "/home/z/my-project/upload/16.10.2026 РОССИЯ с НДС, руб..xlsx"
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
            and not f_val.startswith('СТО') and not f_val.startswith('ТУ') and not f_val.startswith('ГОСТ'):
            current_name = f_val.strip()
        if (m_val is not None and isinstance(m_val, (int, float))) or (n_val is not None and isinstance(n_val, (int, float))):
            price_dict[sn].append({
                'name': current_name, 'fасовка': h_val, 'базовая': m_val, 'отгрузочная': n_val,
            })

# === Топ-5 кандидатов из ABC-анализа ===
top5 = [
    {'sbu': 'Урал', 'brand': 'NoName', 'type': 'Краски', 'grp_code': 63180, 'name': 'Краска ВД для садовых деревьев 3 кг', 'cumulative': 0.640668},
    {'sbu': 'Урал', 'brand': 'NoName', 'type': 'Эмали Алкидные', 'grp_code': 79754, 'name': 'Эмаль ПФ-115 синяя 0.7 кг', 'cumulative': 0.646645},
    {'sbu': 'Урал', 'brand': 'СТРОЙКОМПЛЕКТ', 'type': 'Краски', 'grp_code': 73122, 'name': 'Краска фасадная СТРОЙКОМПЛЕКТ 13 кг', 'cumulative': 0.652555},
    {'sbu': 'Урал', 'brand': 'EMPILS ЭКОНОМ', 'type': 'Эмали Алкидные', 'grp_code': 7803, 'name': 'Эмаль EMPILS ПФ-115 ЭКОНОМ чёрная 20 кг', 'cumulative': 0.655463},
    {'sbu': 'Урал', 'brand': 'СТРОЙКОМПЛЕКТ', 'type': 'Эмали Алкидные', 'grp_code': 72676, 'name': 'Эмаль ПФ-115 СТРОЙКОМПЛЕКТ чёрная 10 кг', 'cumulative': 0.672533},
]

# Определить сферу для каждого товара
def determine_sphere(item):
    name_lower = item['name'].lower()
    if 'садовых деревьев' in name_lower or 'дерев' in name_lower or 'антисептик' in name_lower:
        return 'Дерево'
    if 'фасад' in name_lower:
        return 'Фасад'
    if 'пф-115' in name_lower or 'эмаль' in name_lower or 'по ржавчине' in name_lower:
        return 'Металл'
    if 'интерьер' in name_lower or 'потол' in name_lower:
        return 'Помещения (интерьер)'
    if 'пол' in name_lower and 'для пола' in name_lower:
        return 'Пол'
    return 'Металл'  # default

def determine_purpose(item):
    name_lower = item['name'].lower()
    if 'садовых деревьев' in name_lower:
        return 'Антисептик для древесины'
    if 'фасад' in name_lower:
        return 'Краска фасадная'
    if 'пф-115' in name_lower:
        return 'Эмаль ПФ-115'
    if 'для пола' in name_lower:
        return 'Эмаль для пола'
    return item.get('type', '')

def determine_base(item):
    if item['type'] in ('Краски', 'Грунтовки') or 'ВД' in item['name'] or 'водно-дисперсионная' in item['name'].lower():
        return 'Водно-дисперсионная'
    if 'Эмали' in item['type'] or 'пф' in item['name'].lower() or 'эмаль' in item['name'].lower():
        return 'Органическая'
    return 'Органическая'

# Нормализация для поиска
def norm(s):
    if s is None:
        return ''
    s = re.sub(r'[«»""„"\']', ' ', str(s))
    s = re.sub(r'\s*\d+(?:[.,]\d+)?\s*(?:кг|л|l)\s*', ' ', s, flags=re.IGNORECASE)
    s = re.sub(r'[.,;:!?()\[\]{}-]', ' ', s)
    return s.lower().strip()

# Извлечь фасовку из имени
def extract_fas(name):
    m = re.search(r'(\d+(?:\.\d+)?)\s*кг', str(name), re.IGNORECASE)
    return float(m.group(1)) if m else None

# Brand → sheets
brand_to_sheets = {
    'NoName': ['NoName'],
    'СТРОЙКОМПЛЕКТ': ['СТРОЙКОМПЛЕКТ'],
    'EMPILS ЭКОНОМ': ['EMPILS ЭКОНОМ'],
    'AvantGarde': ['AvantGarde'],
    'NEWTONE': ['NEWTONE'],
    'PROFIWOOD': ['PROFIWOOD'],
    'ОРЕОЛ TITAN': ['ОРЕОЛ TITAN', 'ОРЕОЛ TITAN(л)'],
    'Ореол': ['Ореол'],
    'Ореол Дисконт': ['Ореол Дисконт'],
    'Простокрашено! БАУ': ['Простокрашено! БАУ'],
    'Расцвет': ['Расцвет'],
}

# === Поиск цен для каждого топ-5 ===
print("=" * 100)
print("ПОИСК ЦЕН ДЛЯ 5 НОВЫХ SKU ИЗ ABC-АНАЛИЗА (топ-5 группы А)")
print("=" * 100)

results = []
for cand in top5:
    print(f"\n--- {cand['name']} (cumulative={cand['cumulative']:.3f}) ---")
    brand = cand['brand']
    sheets = brand_to_sheets.get(brand, [])
    print(f"  Бренд: {brand} | Листы: {sheets}")
    
    cand_name_norm = norm(cand['name'])
    cand_fas = extract_fas(cand['name'])
    print(f"  Нормализованное имя: '{cand_name_norm}' | Фасовка: {cand_fas}")
    
    # Поиск по точному и подстрочному совпадению
    matches = []
    for sn in sheets:
        for p in price_dict[sn]:
            if p['name'] is None:
                continue
            pn = norm(p['name'])
            if not pn:
                continue
            
            # Точное совпадение
            if pn == cand_name_norm:
                matches.append((sn, p, 'exact'))
            elif cand_name_norm in pn or pn in cand_name_norm:
                # Проверка на значимое пересечение
                a_tokens = set(re.findall(r'[a-zа-яё]+', cand_name_norm))
                b_tokens = set(re.findall(r'[a-zа-яё]+', pn))
                overlap = a_tokens & b_tokens
                if len(overlap) >= max(2, len(a_tokens) // 2):
                    matches.append((sn, p, 'substring'))
    
    if matches:
        # Если есть фасовка, выбрать совпадающую
        chosen = None
        if cand_fas:
            for sn, p, mt in matches:
                if isinstance(p['fасовка'], (int, float)) and abs(p['fасовка'] - cand_fas) <= 0.5:
                    chosen = (sn, p, mt)
                    break
        if not chosen:
            # Берём первое
            chosen = matches[0]
        
        sn, p, mt = chosen
        print(f"  ✓ НАЙДЕНО [{sn}]: '{p['name']}' | фас={p['fасовка']} | отгруз={p['отгрузочная']} | базовая={p['базовая']} | тип={mt}")
        results.append({
            'cand': cand,
            'match': {
                'name': p['name'], 'fасовка': p['fасовка'],
                'базовая': p['базовая'], 'отгрузочная': p['отгрузочная'],
                'лист': sn, 'match_type': mt,
            }
        })
    else:
        print(f"  ✗ НЕ НАЙДЕНО в прайс-листе")
        # Покажем все позиции этого бренда для отладки
        if sheets:
            print(f"  Все позиции бренда в листе '{sheets[0]}':")
            for p in price_dict[sheets[0]][:10]:
                print(f"    '{p['name']}' | фас={p['fасовка']} | отгруз={p['отгрузочная']}")

# === Сводка ===
print("\n" + "=" * 100)
print("ИТОГ:")
print("=" * 100)
for r in results:
    c = r['cand']
    m = r['match']
    print(f"  Код {c['grp_code']} | {c['brand']} | {c['name']}")
    print(f"    → [{m['лист']}] {m['name']}")
    print(f"    фасовка={m['fасовка']} | отгрузочная={m['отгрузочная']} | базовая={m['базовая']}")
    print(f"    Сфера: {determine_sphere(c)} | Назначение: {determine_purpose(c)} | Основа: {determine_base(c)}")
