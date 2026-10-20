import openpyxl
from openpyxl import load_workbook
from collections import defaultdict
import re

# Reload everything
exec(open('/home/z/my-project/scripts/match_prices_v3.py').read().split('# === 5. Process all assortment items ===')[0])

# Test NEWTONE SKY
name = 'Краска ВД «NEWTONE SKY» для потолков ультрабелая 9 л 14.5 кг'
print(f"Testing: '{name}'")
result = match_price('NEWTONE', name)
print(f"Result: {result}")

# Debug
name_norm = norm_for_match(name)
print(f"Normalized: '{name_norm}'")
a_tokens = tokenize(name_norm)
a_tokens_clean = a_tokens - FILLER_TOKENS
print(f"Tokens clean: {a_tokens_clean}")
a_colors = extract_colors(name_norm)
print(f"Colors: {a_colors}")
kg, l = extract_fasovka(name)
print(f"Фасовка: kg={kg}, l={l}")

print("\n=== Candidates from NEWTONE sheet ===")
for p in price_dict['NEWTONE']:
    if p['name'] and 'sky' in p['name'].lower():
        pn = norm_for_match(p['name'])
        print(f"\nPrice: '{p['name']}'")
        print(f"  Norm: '{pn}'")
        b_tokens = tokenize(pn)
        b_tokens_clean = b_tokens - FILLER_TOKENS
        print(f"  Tokens clean: {b_tokens_clean}")
        b_colors = extract_colors(pn)
        print(f"  Colors: {b_colors}")
        a_in_b = a_tokens_clean.issubset(b_tokens_clean)
        b_in_a = b_tokens_clean.issubset(a_tokens_clean)
        intersection = len(a_tokens_clean & b_tokens_clean)
        union = len(a_tokens_clean | b_tokens_clean)
        jaccard = intersection / union if union else 0
        print(f"  a_in_b={a_in_b}, b_in_a={b_in_a}, jaccard={jaccard:.3f}")
        print(f"  Фасовка: {p['fасовка']}")
