import openpyxl
from openpyxl import load_workbook
from collections import defaultdict
import re

# === 1. Read price list ===
wb2 = load_workbook("/home/z/my-project/upload/21.08.2026 РОССИЯ с НДС, руб..xlsx", data_only=True)

price_dict = defaultdict(list)
for sn in wb2.sheetnames:
    ws = wb2[sn]
    current_name = None
    for row_idx, row in enumerate(ws.iter_rows(min_row=14, max_row=ws.max_row, values_only=True), 14):
        f_val = row[5] if len(row) > 5 else None
        m_val = row[12] if len(row) > 12 else None
        n_val = row[13] if len(row) > 13 else None
        h_val = row[7] if len(row) > 7 else None
        if f_val is not None and isinstance(f_val, str) and not f_val.startswith('СТО') and not f_val.startswith('ТУ') and not f_val.startswith('ГОСТ'):
            current_name = f_val.strip()
        if (m_val is not None and isinstance(m_val, (int, float))) or (n_val is not None and isinstance(n_val, (int, float))):
            price_dict[sn].append({
                'name': current_name,
                'fасовка': h_val,
                'базовая': m_val,
                'отгрузочная': n_val,
                'sheet': sn,
            })

# === 2. Normalization helpers ===
QUOTE_RE = re.compile(r'[«»""„"\']')
FASOVKA_RE = re.compile(r'\s+\d+(?:\.\d+)?\s*(?:кг|л|l)\s*', re.IGNORECASE)

def strip_quotes(s):
    return QUOTE_RE.sub(' ', s)

def norm_for_match(name):
    if name is None:
        return ''
    s = strip_quotes(str(name))
    s = FASOVKA_RE.sub(' ', s)
    # Replace punctuation with spaces
    s = re.sub(r'[.,;:!?]', ' ', s)
    s = s.lower()
    s = re.sub(r'\s+', ' ', s)
    return s.strip()

def tokenize(s):
    s = re.sub(r'[\(\)]', ' ', s)
    tokens = set(re.findall(r'[a-zа-яё]+(?:-[a-zа-яё]+)*|\d+(?:\.\d+)?', s))
    return tokens

# Filler tokens to ignore
FILLER_TOKENS = {
    'с', 'в', 'и', 'по', 'для', 'на', 'из', 'от', 'до',  # prepositions
    'гост', 'тз', 'сто', 'ту',  # standards
    'бау',  # acronym
    'а', 'с',  # base codes (single letters)
    'алкидно', 'стирольная',  # combined but adjective
    'водно', 'дисперсионная', 'полиакриловая', 'латексная', 'силиконовая',
    'универсальная', 'атмосферостойкий', 'ак', 'пф', 'вд', 'вда',
    'кг', 'л', 'м',  # units
    'мл', 'мм',
    # Descriptors that don't affect identity
    'матричная', 'матовая', 'глянцевая', 'полуматовая', 'глубокоматовая', 'суперстойкая',
    'интерьерная', 'фасадная', 'термостойкая', 'быстросохнущая', 'моющаяся',
}

# Color detection: extract all color-like tokens
COLOR_PATTERNS = [
    r'бел\w+', r'черн\w+', r'чёрн\w+', r'красн\w+', r'сер\w+', r'зелён\w+', r'зелен\w+',
    r'син\w+', r'голуб\w+', r'жёлт\w+', r'желт\w+', r'коричнев\w+', r'вишнёв\w+', r'вишнев\w+',
    r'шоколад\w+', r'бежев\w+', r'серебрист\w+', r'золотист\w+', r'медн\w+', r'графитов\w+',
    r'оранжев\w+', r'салатов\w+', r'бирюзов\w+', r'васильков\w+', r'ультрабел\w+',
    r'скандинавск\w+', r'горьк\w+',
    r'тёмно-\w+', r'темно-\w+', r'светло-\w+', r'ярко-\w+', r'красно-\w+',
    r'тёмнозелён\w+', r'темнозелен\w+', r'тёмносер\w+', r'темносер\w+',
    r'светлосер\w+', r'светлозелён\w+', r'ярко-жёлт\w+', r'ярко-желт\w+', r'ярко-зелен\w+',
    r'слонов\w+', r'кость\w*',
    r'красно-коричнев\w+',
    # AvantGarde / PROFIWOOD special
    r'скандинавск\w+', r'туман\w*',
    r'бесцветн\w+',
    r'золотисто-коричнев\w+', r'золотисто-жёлт\w+',
    # Color word "RAL XXXX" handled by digit pattern
]
COLOR_RE = re.compile('|'.join(COLOR_PATTERNS), re.IGNORECASE)

def extract_colors(s):
    """Extract set of color tokens from normalized name."""
    matches = COLOR_RE.findall(s)
    return set(m.lower() for m in matches if m)

KG_RE = re.compile(r'(\d+(?:\.\d+)?)\s*кг')
L_RE = re.compile(r'(\d+(?:\.\d+)?)\s*[лl]')

def extract_fasovka(name):
    if name is None:
        return (None, None)
    s = str(name)
    kg_m = KG_RE.search(s)
    l_m = L_RE.search(s)
    kg = float(kg_m.group(1)) if kg_m else None
    l = float(l_m.group(1)) if l_m else None
    return (kg, l)

# === 3. Brand → sheet mapping ===
brand_to_sheets = {
    'AvantGarde': ['AvantGarde'],
    'NEWTONE': ['NEWTONE'],
    'PROFIWOOD': ['PROFIWOOD'],
    'ОРЕОЛ TITAN': ['ОРЕОЛ TITAN', 'ОРЕОЛ TITAN(л)'],
    'Ореол': ['Ореол'],
    'Ореол Дисконт': ['Ореол Дисконт'],
    'Простокрашено! БАУ': ['Простокрашено! БАУ'],
    'Расцвет': ['Расцвет'],
}

# === 4. Match function v3 ===
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
            
            # Color check: if both have colors, they must overlap
            color_ok = True
            if a_colors and b_colors:
                if not a_colors.issubset(b_colors) and not b_colors.issubset(a_colors):
                    # Check overlap — at least 50%
                    overlap = len(a_colors & b_colors)
                    if overlap == 0:
                        color_ok = False
            
            if not color_ok:
                continue
            
            # Subset check (with filler removed)
            a_in_b = a_tokens_clean.issubset(b_tokens_clean)
            b_in_a = b_tokens_clean.issubset(a_tokens_clean)
            
            # Jaccard similarity
            intersection = len(a_tokens_clean & b_tokens_clean)
            union = len(a_tokens_clean | b_tokens_clean)
            jaccard = intersection / union if union else 0
            
            # Exact match
            if pn == name_norm:
                candidates.append((sn, p, 'exact', 1.0))
                continue
            
            # Subset match
            if (a_in_b or b_in_a) and jaccard >= 0.5:
                candidates.append((sn, p, 'subset', jaccard))
                continue
            
            # Substring match
            if name_norm in pn or pn in name_norm:
                candidates.append((sn, p, 'substr', 0.7))
                continue
            
            # High Jaccard even without subset
            if jaccard >= 0.65:
                candidates.append((sn, p, 'jaccard', jaccard))
                continue
    
    if not candidates:
        return None
    
    # Sort by match quality
    candidates.sort(key=lambda c: ({'exact': 0, 'subset': 1, 'substr': 2, 'jaccard': 3}[c[2]], -c[3]))
    
    # Filter to top quality
    top_mt = candidates[0][2]
    top_ratio = candidates[0][3]
    same_quality = [c for c in candidates if c[2] == top_mt and abs(c[3] - top_ratio) < 0.01]
    
    # If фасовka extracted, find best фасовка match among top candidates
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
                'фасовка': best[1]['fасовка'],
                'лист': best[0],
                'ценовое_имя': best[1]['name'],
                'match_type': best[2],
                'tolerance': best_diff,
            }
    
    # No фасовка match — take smallest from top candidates
    with_fas = [c for c in same_quality if isinstance(c[1]['fасовка'], (int, float))]
    if with_fas:
        with_fas.sort(key=lambda c: c[1]['fасовка'])
        sn, p, mt, ratio = with_fas[0]
        return {
            'базовая': p['базовая'],
            'отгрузочная': p['отгрузочная'],
            'фасовка': p['fасовка'],
            'лист': sn,
            'ценовое_имя': p['name'],
            'match_type': mt + '_smallest',
            'tolerance': None,
        }
    
    # Last resort: first candidate
    sn, p, mt, ratio = candidates[0]
    return {
        'базовая': p['базовая'],
        'отгрузочная': p['отгрузочная'],
        'фасовка': p['fасовка'],
        'лист': sn,
        'ценовое_имя': p['name'],
        'match_type': mt + '_first',
        'tolerance': None,
    }

# === 5. Process all assortment items ===
wb1 = load_workbook("/home/z/my-project/upload/Ассортимент_Олимп_300_SKU.xlsx", data_only=True)
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

print(f"Сопоставлено: {matched} / {len(results)}")
print(f"Не сопоставлено: {len(unmatched)}")
print(f"\nТипы сопоставления:")
for mt, cnt in sorted(match_types.items(), key=lambda x: -x[1]):
    print(f"  {mt}: {cnt}")

# Print all unmatched
print("\n=== ВСЕ не сопоставленные ===")
for item in unmatched:
    print(f"  [{item['idx']}] {item['brand']} | {item['name']}")
