#!/usr/bin/env python3
"""
Дополнительный сбор цен для товаров, где первый поиск не дал результата.
Стратегия: 
1. Поиск с другим запросом (с ценой или купить)
2. Фильтрация цен по разумному диапазону для 0.75-0.9 л
"""

import subprocess
import json
import re
import time

# Товары, для которых нужно найти цену
# Из первого прохода — где price_new = None или > 3000 (некорректная цена)
products_to_retry = [
    {'n': 2, 'brand': 'Elcon', 'product': 'Грунт-эмаль 3 в 1 Elcon глянцевая чёрная RAL 9005 0.8 кг', 'price_old': 958,
     'queries': [
        'купить Elcon грунт-эмаль 3 в 1 глянцевая RAL 9005 0.8 кг цена',
        'Elcon 3 в 1 глянцевая черная RAL 9005 0.8 кг всеИнструменты',
        'грунт-эмаль Elcon 3 в 1 RAL 9005 0.8 цена',
     ]},
    {'n': 4, 'brand': 'Dali', 'product': 'Грунт-эмаль по ржавчине Dali 3 в 1, чёрная RAL 9005 0.75 л', 'price_old': 1055,
     'queries': [
        'купить Dali 3 в 1 грунт-эмаль RAL 9005 0.75 л цена',
        'Dali 3 в 1 молотковая 0.75 л всеИнструменты цена',
        'грунт-эмаль Dali 3 в 1 черная RAL 9005 0.75 цена',
     ]},
    {'n': 5, 'brand': 'HUSKY KLONDIKE', 'product': 'Краска по металлу HUSKY KLONDIKE глянцевая RAL 9005 0.9 л', 'price_old': 1153,
     'queries': [
        'купить HUSKY KLONDIKE краска по металлу RAL 9005 0.9 л цена',
        'HUSKY KLONDIKE глянцевая черная RAL 9005 0.9 всеИнструменты',
     ]},
    {'n': 6, 'brand': 'Tikkurila', 'product': 'Tikkurila METALLISTA 3 в 1, 0.8 л, чёрная', 'price_old': 2085,
     'queries': [
        'купить Tikkurila METALLISTA 3 в 1 0.8 чёрная цена',
        'Tikkurila METALLISTA 0.8 л всеИнструменты',
     ]},
    {'n': 7, 'brand': 'Новбытхим', 'product': 'Грунт-эмаль 5 в 1 Новбытхим серый RAL 7040 0.75 л', 'price_old': 765,
     'queries': [
        'купить Новбытхим 5 в 1 грунт-эмаль RAL 7040 0.75 л цена',
        'Новбытхим 5в1 серый RAL 7040 0.75 всеИнструменты',
     ]},
    {'n': 8, 'brand': 'TITAN', 'product': 'Грунт-эмаль быстросохнущая ОРЕОЛ TITAN матовая серая 0.75 л', 'price_old': 765,
     'queries': [
        'купить ОРЕОЛ TITAN быстросохнущая матовая серая 0.75 л цена',
        'TITAN быстросохнущая матовая серая 0.75 всеИнструменты цена',
        'ОРЕОЛ TITAN матовая серая 0.75 0.9 кг грунт-эмаль цена',
     ]},
    {'n': 9, 'brand': 'Elcon', 'product': 'Быстросохнущая грунт-эмаль Elcon 3 в 1 матовая серая 0.8 кг', 'price_old': 838,
     'queries': [
        'купить Elcon быстросохнущая 3 в 1 матовая серая 0.8 цена',
        'Elcon 3 в 1 быстросохнущая матовая серая 0.8 всеИнструменты',
     ]},
    {'n': 10, 'brand': 'Лакра', 'product': 'Грунт-эмаль ЛАКРА PROF IT 4 в 1 быстросохнущая RAL 7024 0.8 кг', 'price_old': 906,
     'queries': [
        'купить Лакра PROF IT 4 в 1 быстросохнущая RAL 7024 0.8 цена',
        'Лакра PROF IT 4 в 1 графит 0.8 всеИнструменты',
     ]},
    {'n': 11, 'brand': 'Dali', 'product': 'Грунт-эмаль Dali 3 в 1 молотковая золотистая 0.75 л', 'price_old': 1057,
     'queries': [
        'купить Dali молотковая золотистая 3 в 1 0.75 цена',
        'Dali 3 в 1 молотковая золотистая 0.75 всеИнструменты',
     ]},
    {'n': 14, 'brand': 'TITAN', 'product': 'Эмаль TITAN молотковая алкидно-стирольная золотистая 0.8 кг', 'price_old': 1308,
     'queries': [
        'купить TITAN молотковая алкидно-стирольная золотистая 0.8 цена',
        'ОРЕОЛ TITAN молотковая золотистая 0.8 всеИнструменты',
        'Эмаль TITAN молотковая 0.8 кг золотистая всеИнструменты цена',
     ]},
    {'n': 15, 'brand': 'Лакра', 'product': 'Эмаль Лакра молотковая золото 0.8 кг', 'price_old': 1341,
     'queries': [
        'купить Лакра молотковая золото 0.8 кг цена',
        'Лакра молотковая золото 0.8 всеИнструменты',
     ]},
    {'n': 16, 'brand': 'Dufa', 'product': 'Эмаль Dufa Premium HAMMERLACK молотковая золотая 750 мл', 'price_old': 1653,
     'queries': [
        'купить Dufa Premium HAMMERLACK молотковая 750 мл цена',
        'Dufa HAMMERLACK золотая 750 всеИнструменты',
     ]},
]


def search_web(query, num=10):
    try:
        cmd = ['z-ai', 'function', '-n', 'web_search', '-a', json.dumps({'query': query, 'num': num})]
        r = subprocess.run(cmd, capture_output=True, text=True, timeout=60)
        if r.returncode != 0:
            return []
        output = r.stdout
        idx = output.find('[')
        end_idx = output.rfind(']')
        if idx < 0 or end_idx < idx:
            return []
        json_str = output[idx:end_idx+1]
        try:
            return json.loads(json_str)
        except:
            return []
    except Exception as e:
        return []


def extract_prices_filtered(text, min_price=400, max_price=3500):
    """Извлекает цены и фильтрует по диапазону для грунт-эмали 0.75-0.9л."""
    if not text:
        return []
    prices = []
    patterns = [
        r'(\d{1,3}[\s\u00A0]?\d{3})\s*[₽]',
        r'(\d{3,5})\s*[₽]',
        r'(\d{1,3}[\s\u00A0]?\d{3})\s*руб',
    ]
    for pat in patterns:
        for m in re.finditer(pat, text, re.IGNORECASE):
            price_str = m.group(1).replace(' ', '').replace('\u00A0', '')
            try:
                p = int(price_str)
                if min_price <= p <= max_price:
                    prices.append(p)
            except:
                continue
    return sorted(set(prices), reverse=True)


# Обработка
print("="*80)
print("ПОВТОРНЫЙ СБОР ЦЕН")
print("="*80)

results = {}
for p in products_to_retry:
    print(f"\n[{p['n']}] {p['brand']} — {p['product'][:50]}...")
    found_price = None
    found_url = ''
    
    for qi, query in enumerate(p['queries']):
        print(f"  Query {qi+1}: {query[:70]}")
        search_results = search_web(query, num=10)
        if not search_results:
            print(f"    нет результатов")
            continue
        
        # Ищем vseinstrumenti
        vse_results = [r for r in search_results if 'vseinstrumenti' in r.get('host_name', '').lower() or 'vseinstrumenti' in r.get('url', '').lower()]
        if not vse_results:
            vse_results = search_results  # берём все
        
        for r in vse_results[:5]:
            snippet = r.get('snippet', '')
            name = r.get('name', '')
            combined = f"{name} {snippet}"
            prices = extract_prices_filtered(combined)
            if prices:
                # Берём максимальную (без акции)
                found_price = prices[0]
                found_url = r.get('url', '')
                print(f"    ✅ Цена: {found_price} ₽ (все: {prices})")
                print(f"       URL: {found_url[:80]}")
                break
        
        if found_price:
            break
        time.sleep(1)
    
    if not found_price:
        print(f"  ❌ Цена не найдена после {len(p['queries'])} запросов")
    
    results[p['n']] = {
        'brand': p['brand'],
        'product': p['product'],
        'price_old': p['price_old'],
        'price_new': found_price,
        'url': found_url,
    }

# Сохранение
out_json = '/home/z/my-project/scripts/prices_retry_results.json'
with open(out_json, 'w', encoding='utf-8') as f:
    json.dump(results, f, ensure_ascii=False, indent=2)

print("\n" + "="*80)
print("ИТОГИ ПОВТОРНОГО СБОРА")
print("="*80)
found = sum(1 for r in results.values() if r['price_new'])
print(f"Найдено: {found} из {len(results)}")
print(f"Сохранено: {out_json}")
print()
print(f"{'N':<3} | {'Бренд':<15} | {'Цена 10.06':<10} | {'Цена акт.':<10}")
print('-' * 50)
for n, r in results.items():
    pn = r['price_new'] if r['price_new'] else '—'
    print(f"{n:<3} | {r['brand']:<15} | {r['price_old']:<10} | {str(pn):<10}")
