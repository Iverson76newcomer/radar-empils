import openpyxl
from openpyxl import load_workbook
from collections import defaultdict
import re
import shutil

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

def strip_quotes(s):
    return QUOTE_RE.sub(' ', s)

def norm(s):
    if s is None:
        return ''
    s = str(s).strip()
    s = strip_quotes(s)
    s = s.lower()
    s = re.sub(r'\s+', ' ', s)
    return s.strip()

# Extract фасовка (kg and L) from assortment name
KG_RE = re.compile(r'(\d+(?:\.\d+)?)\s*кг')
L_RE = re.compile(r'(\d+(?:\.\d+)?)\s*[лl]')

def extract_fasovka(name):
    """Return (kg_value, l_value) — both can be None."""
    if name is None:
        return (None, None)
    s = str(name)
    kg_m = KG_RE.search(s)
    l_m = L_RE.search(s)
    kg = float(kg_m.group(1)) if kg_m else None
    l = float(l_m.group(1)) if l_m else None
    return (kg, l)

# Remove фасовка suffix from name for matching
FASOVKA_RE = re.compile(r'\s+(\d+(?:\.\d+)?)\s*(кг|л|l)\s*(\d+(?:\.\d+)?)\s*(кг|л|l)?', re.IGNORECASE)

def norm_for_match(name):
    """Normalize and strip фасовка info from name for matching."""
    if name is None:
        return ''
    s = strip_quotes(str(name))
    # Remove фасовка patterns: "0.9 кг", "0.75 л 0.9 кг", "9 л 14.5 кг", "10 кг"
    s = re.sub(r'\s+\d+(?:\.\d+)?\s*(?:кг|л|l)\s*', ' ', s, flags=re.IGNORECASE)
    s = s.lower()
    s = re.sub(r'\s+', ' ', s)
    return s.strip()

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

# === 4. Match function ===
def match_price(item_brand, item_name):
    """Returns dict with matched price info or None."""
    sheets = brand_to_sheets.get(item_brand, [])
    if not sheets:
        return None
    
    name_norm = norm_for_match(item_name)
    if not name_norm:
        return None
    
    kg, l = extract_fasovka(item_name)
    
    # Collect candidates: price entries whose normalized name == our normalized name
    # or where our name is substring of price name or vice versa
    candidates = []
    for sn in sheets:
        for p in price_dict[sn]:
            if p['name'] is None:
                continue
            pn = norm_for_match(p['name'])
            if not pn:
                continue
            # Exact match preferred
            if pn == name_norm:
                candidates.append((sn, p, 'exact'))
            elif name_norm in pn or pn in name_norm:
                # Substring match — check key tokens match (color, RAL, base)
                # Tokenize
                a_tokens = set(re.findall(r'[a-zа-яё]+(?:-[a-zа-яё]+)?|\d+', name_norm))
                b_tokens = set(re.findall(r'[a-zа-яё]+(?:-[a-zа-яё]+)?|\d+', pn))
                # require >=80% of smaller set to match
                if a_tokens and b_tokens:
                    overlap = a_tokens & b_tokens
                    smaller = min(len(a_tokens), len(b_tokens))
                    if len(overlap) / smaller >= 0.8:
                        candidates.append((sn, p, 'token'))
    
    if not candidates:
        return None
    
    # Prefer exact matches
    exact = [c for c in candidates if c[2] == 'exact']
    if exact:
        candidates = exact
    
    # If фасовка extracted, try to find closest фасовка
    if kg is not None or l is not None:
        # Try kg match first
        target_values = []
        if kg is not None:
            target_values.append(kg)
        if l is not None:
            target_values.append(l)
        
        best = None
        best_diff = float('inf')
        for sn, p, mt in candidates:
            f = p['fасовка']
            if f is None or not isinstance(f, (int, float)):
                continue
            for tv in target_values:
                diff = abs(f - tv)
                if diff < best_diff:
                    best_diff = diff
                    best = (sn, p, mt)
        
        # Accept if within 0.3 kg/L tolerance
        if best is not None and best_diff <= 0.3:
            return {
                'базовая': best[1]['базовая'],
                'отгрузочная': best[1]['отгрузочная'],
                'фасовка': best[1]['fасовка'],
                'лист': best[0],
                'ценовое_имя': best[1]['name'],
                'match_type': best[2],
                'tolerance': best_diff,
            }
    
    # Fallback: take smallest фасовка (first entry)
    # Or if only one candidate, take it
    if len(candidates) == 1:
        sn, p, mt = candidates[0]
        return {
            'базовая': p['базовая'],
            'отгрузочная': p['отгрузочная'],
            'фасовка': p['fасовка'],
            'лист': sn,
            'ценовое_имя': p['name'],
            'match_type': mt,
            'tolerance': None,
        }
    
    # Take smallest фасовка
    valid = [c for c in candidates if c[1]['fасовка'] is not None]
    if valid:
        # Sort by фасовка
        valid_sorted = sorted(valid, key=lambda c: c[1]['fасовка'] if isinstance(c[1]['fасовка'], (int, float)) else 999)
        sn, p, mt = valid_sorted[0]
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
    sn, p, mt = candidates[0]
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
    else:
        unmatched.append(item)
    results.append(item)

print(f"Сопоставлено: {matched} / {len(results)}")
print(f"Не сопоставлено: {len(unmatched)}")

# Show unmatched sample
print("\n=== Не сопоставлено (первые 30) ===")
for item in unmatched[:30]:
    print(f"  [{item['idx']}] {item['brand']} | {item['name']}")

# Save intermediate results for review
import json
with open('/home/z/my-project/scripts/match_results.json', 'w', encoding='utf-8') as f:
    json.dump({
        'matched': matched,
        'unmatched_count': len(unmatched),
        'unmatched_sample': [
            {'idx': it['idx'], 'brand': it['brand'], 'name': it['name']} for it in unmatched[:50]
        ]
    }, f, ensure_ascii=False, indent=2)
print("\nSaved match_results.json")
