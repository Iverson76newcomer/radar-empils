"""Дополнительные поиски для продуктов, у которых нет цены."""
import json
import subprocess
import re
import time

with open('/home/z/my-project/scripts/vsi_prices_final.json', 'r', encoding='utf-8') as f:
    results = json.load(f)

# Какие продукты без цены
missing = [r for r in results if not r.get('final_price')]
print(f"Продуктов без цены: {len(missing)}")

# Уточнённые запросы для каждого
extra_queries = {
    # Категория 1
    "1_Elcon_3в1 глянцевая": "Elcon грунт-эмаль 3 в 1 глянцевая RAL 9005 0.8 кг 00-00750063 купить цена",
    "1_Dufa CORREX_3в1": "Dufa CORREX 3 в 1 RAL 9005 0.75 л МП00-013659 купить цена",
    # Категория 2
    "2_ОРЕОЛ TITAN_быстросохнущая серая": "ОРЕОЛ TITAN грунт-эмаль быстросохнущая матовая серая 0.75 л 83764 купить цена",
    "2_ЛАКРА PROF IT_4в1": "ЛАКРА PROF IT грунт-эмаль 4 в 1 серая 0.8 кг купить цена",
    # Категория 3
    "3_ОРЕОЛ TITAN_молотковая": "ОРЕОЛ TITAN эмаль молотковая золотистая 0.8 кг 77639 купить цена",
    "3_Dufa Premium HAMMERLACK_молотковая": "Dufa Premium HAMMERLACK молотковая 750 мл купить цена",
    "3_HUSKY KLONDIKE_золото": "HUSKY Klondike краска по металлу золото молотковая 0.9 купить цена",
    # Категория 4
    "4_Лакра_5в1 грунт-эмаль": "Лакра грунт-эмаль 5 в 1 серая 0.8 купить цена",
    # Категория 5
    "5_TOPGRADE_грунт-эмаль": "Topgrade by Hammerite краска гладкая 0.75 купить цена",
}

def search_with_snippets(query, num=10):
    args = json.dumps({"query": query, "num": num})
    out = subprocess.run(['z-ai', 'function', '-n', 'web_search', '-a', args],
                         capture_output=True, text=True, timeout=90)
    try:
        output = out.stdout
        json_start = output.find('[')
        json_end = output.rfind(']')
        if json_start >= 0 and json_end > json_start:
            return json.loads(output[json_start:json_end+1])
    except:
        pass
    return []

def extract_prices_from_snippet(snippet):
    if not snippet:
        return None, None
    # "от X рублей"
    m = re.search(r'(?:цен[ае]\s+)?от\s+(\d{1,3}(?:\s?\d{3})*)\s*(?:руб|₽)?', snippet, re.IGNORECASE)
    if m:
        try:
            val = int(m.group(1).replace(' ', ''))
            if 100 < val < 100000:
                return val, None
        except:
            pass
    prices = []
    for m in re.finditer(r'(\d{1,3}(?:\s\d{3})+|\d{3,4})\s*₽', snippet):
        try:
            val = int(m.group(1).replace(' ', ''))
            if 100 < val < 15000:
                prices.append(val)
        except:
            pass
    if not prices:
        return None, None
    if len(prices) == 1:
        return prices[0], None
    p1, p2 = prices[0], prices[1]
    if 'вместо' in snippet.lower():
        return p1, (p2 if p2 < 15000 else None)
    return min(p1, p2), max(p1, p2)

# Обработка пропущенных
for r in missing:
    key = f"{r['cat']}_{r['brand']}_{r['product']}"
    query = extra_queries.get(key, f"{r['brand']} {r['product']} 0.8 купить цена")
    print(f"\n[{r['cat']}.{r['brand']}] query: {query[:60]}")
    data = search_with_snippets(query)
    found_price = None
    found_url = None
    for d in data:
        url = d.get('url', '')
        snippet = d.get('snippet', '') or ''
        cur, old = extract_prices_from_snippet(snippet)
        if cur or old:
            final = old if old else cur
            print(f"  URL: {url[:80]}")
            print(f"  snippet: {snippet[:200]}")
            print(f"  → cur={cur}, old={old}, final={final}")
            found_price = final
            if 'vseinstrumenti.ru/product/' in url:
                found_url = url
            break
    if found_price:
        r['final_price'] = found_price
        r['old_price'] = old
        r['current_price'] = cur
        if found_url:
            r['url'] = found_url
        r['source'] = 'extra_search'
    else:
        print(f"  ✗ Не найдено")
    time.sleep(1.5)

# Сохраняем обновлённые результаты
with open('/home/z/my-project/scripts/vsi_prices_final.json', 'w', encoding='utf-8') as f:
    json.dump(results, f, ensure_ascii=False, indent=2)

print(f"\n=== ИТОГ после доп. поиска ===")
found = sum(1 for r in results if r.get('final_price'))
print(f"Цен найдено: {found} из {len(results)}")
print(f"\n{'#':<3} {'Cat':<4} {'Brand':<25} {'Цена':<10} {'URL':<70}")
print("-" * 110)
for i, r in enumerate(results, 1):
    price = r.get('final_price')
    url = r.get('url', '')[:70]
    print(f"{i:<3} {r['cat']:<4} {r['brand']:<25} {price!s:<10} {url}")
