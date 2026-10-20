import sys
sys.path.insert(0, '/home/z/my-project/scripts')

# Re-exec build_final.py to get all the functions in scope
exec(open('/home/z/my-project/scripts/build_final.py').read().split('# Сохраним результаты')[0])

# Test
test_cases = [
    ('Ореол Дисконт', 'Краска ВД Ореол Дисконт интерьерная 1.5 кг'),
    ('Ореол Дисконт', 'Краска ВД Ореол Дисконт для стен и потолков моющаяся 1.5 кг'),
    ('Ореол Дисконт', 'Грунтовка ВД универсальная ОРЕОЛ ДИСКОНТ 1 кг'),
    ('Ореол', 'Краска ВД Ореол белоснежная для стен и потолков 3 кг'),
    ('Ореол', 'Грунт-концентрат АКВАСТОП ОРЕОЛ 1 кг'),
]

for brand, name in test_cases:
    print(f"\n=== {brand} | {name} ===")
    name_norm = norm_for_match(name)
    print(f"Normalized: '{name_norm}'")
    a_tokens = tokenize(name_norm)
    a_tokens_clean = a_tokens - FILLER_TOKENS
    a_colors = extract_colors(name_norm)
    kg, l = extract_fasovka(name)
    print(f"Tokens clean: {a_tokens_clean}")
    print(f"Colors: {a_colors}")
    print(f"Фасовка: kg={kg}, l={l}")
    
    sheets = brand_to_sheets[brand]
    print(f"Sheets to search: {sheets}")
    for sn in sheets:
        for p in price_dict[sn]:
            if p['name'] is None:
                continue
            pn = norm_for_match(p['name'])
            b_tokens = tokenize(pn)
            b_tokens_clean = b_tokens - FILLER_TOKENS
            b_colors = extract_colors(pn)
            
            # Quick filter
            if not (a_tokens_clean & b_tokens_clean):
                continue
            # Show only ones with reasonable overlap
            overlap = len(a_tokens_clean & b_tokens_clean)
            if overlap < 2:
                continue
            
            a_in_b = a_tokens_clean.issubset(b_tokens_clean)
            b_in_a = b_tokens_clean.issubset(a_tokens_clean)
            intersection = overlap
            union = len(a_tokens_clean | b_tokens_clean)
            jaccard = intersection / union if union else 0
            
            print(f"  [{sn}] '{p['name']}' фас={p['fасовка']}")
            print(f"    Norm: '{pn}'")
            print(f"    Tokens clean: {b_tokens_clean}")
            print(f"    Colors: {b_colors}")
            print(f"    a_in_b={a_in_b}, b_in_a={b_in_a}, jaccard={jaccard:.3f}")
            
            # Color check
            if a_colors and b_colors:
                if not a_colors.issubset(b_colors) and not b_colors.issubset(a_colors):
                    overlap_c = len(a_colors & b_colors)
                    if overlap_c == 0:
                        print(f"    COLOR MISMATCH (a={a_colors}, b={b_colors})")
