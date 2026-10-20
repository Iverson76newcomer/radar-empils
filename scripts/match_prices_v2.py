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
    """Normalize name: lowercase, strip quotes, strip фасовка, collapse whitespace."""
    if name is None:
        return ''
    s = strip_quotes(str(name))
    s = FASOVKA_RE.sub(' ', s)
    s = s.lower()
    s = re.sub(r'\s+', ' ', s)
    return s.strip()

def tokenize(s):
    """Extract tokens (words and numbers), treating parenthesized as separate."""
    # Replace parens with spaces
    s = re.sub(r'[\(\)]', ' ', s)
    # Find tokens: words (incl. hyphenated) or numbers
    tokens = set(re.findall(r'[a-zа-яё]+(?:-[a-zа-яё]+)*|\d+(?:\.\d+)?', s))
    # Remove common stopwords: с, в, и, по, для — single letters/short words
    # Keep them actually because they help distinguish
    return tokens

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

# === 4. Match function (improved) ===
def match_price(item_brand, item_name):
    sheets = brand_to_sheets.get(item_brand, [])
    if not sheets:
        return None
    
    name_norm = norm_for_match(item_name)
    if not name_norm:
        return None
    a_tokens = tokenize(name_norm)
    if not a_tokens:
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
            if not b_tokens:
                continue
            
            # Exact match
            if pn == name_norm:
                candidates.append((sn, p, 'exact', 1.0))
                continue
            
            # Bidirectional subset: all tokens of one side in the other
            a_in_b = a_tokens.issubset(b_tokens)
            b_in_a = b_tokens.issubset(a_tokens)
            
            if a_in_b or b_in_a:
                # Color/base must align (subset already ensures this)
                overlap = len(a_tokens & b_tokens)
                smaller = min(len(a_tokens), len(b_tokens))
                ratio = overlap / smaller if smaller else 0
                if ratio >= 0.85:
                    candidates.append((sn, p, 'subset', ratio))
                continue
            
            # Substring match (fallback)
            if name_norm in pn or pn in name_norm:
                candidates.append((sn, p, 'substr', 0.7))
                continue
    
    if not candidates:
        return None
    
    # Sort by match quality: exact > subset (by ratio) > substr
    candidates.sort(key=lambda c: ({'exact': 0, 'subset': 1, 'substr': 2}[c[2]], -c[3]))
    
    # If фасовka extracted, find best фасовка match among candidates
    if kg is not None or l is not None:
        target_values = []
        if kg is not None:
            target_values.append(kg)
        if l is not None:
            target_values.append(l)
        
        best = None
        best_diff = float('inf')
        for sn, p, mt, ratio in candidates:
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
        # If no фасовка match within tolerance, but there's a unique candidate, take it
        if len(candidates) == 1:
            sn, p, mt, ratio = candidates[0]
            return {
                'базовая': p['базовая'],
                'отгрузочная': p['отгрузочная'],
                'фасовка': p['fасовка'],
                'лист': sn,
                'ценовое_имя': p['name'],
                'match_type': mt + '_no_fasovka_match',
                'tolerance': best_diff if best else None,
            }
    
    # No фасовка in name — take smallest фасовка from top candidates
    # Prefer exact > subset
    if candidates:
        # Take the best candidate (top of sorted list)
        # If best has multiple фасовка options, take smallest
        best_mt = candidates[0][2]
        same_quality = [c for c in candidates if c[2] == best_mt and c[3] == candidates[0][3]]
        if same_quality:
            with_fas = [c for c in same_quality if isinstance(c[1]['fасовка'], (int, float))]
            if with_fas:
                # Take smallest фасовка
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
        # Fall back to first
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
    
    return None

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

print("\n=== Не сопоставлено (первые 40) ===")
for item in unmatched[:40]:
    print(f"  [{item['idx']}] {item['brand']} | {item['name']}")
