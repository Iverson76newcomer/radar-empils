"""Сбор цен через сниппеты web_search.
   VSI в сниппетах отдаёт цены в формате "860 ₽. 987 ₽" (текущая | старая).
   Также ищем "цена от X рублей" как ориентир.
   Правило: если есть перечёркнутая (старая) цена — берём её (постоянная, без скидки).
"""
import json
import subprocess
import re
import time

# Все продукты с уточнёнными запросами
products = [
    # Категория 1: Грунт-эмаль 3в1 высокоглянцевая, чёрный RAL 9005
    {'cat': 1, 'brand': 'ОРЕОЛ TITAN', 'product': '3в1 глянцевая RAL 9005', 'facovka': '0,75 л',
     'query': 'ОРЕОЛ TITAN грунт-эмаль по ржавчине 3 в 1 RAL 9005 0.75 vseinstrumenti.ru'},
    {'cat': 1, 'brand': 'Elcon', 'product': '3в1 глянцевая', 'facovka': '0,8 кг',
     'query': 'Elcon грунт-эмаль 3 в 1 по ржавчине глянцевая RAL 9005 0.8 vseinstrumenti.ru'},
    {'cat': 1, 'brand': 'Dufa CORREX', 'product': '3в1', 'facovka': '0,75 л',
     'query': 'Dufa CORREX грунт-эмаль 3 в 1 RAL 9005 0.75 vseinstrumenti.ru'},
    {'cat': 1, 'brand': 'Dali', 'product': '3в1', 'facovka': '0,75 л',
     'query': 'Dali грунт-эмаль 3 в 1 черная RAL 9005 0.75 vseinstrumenti.ru'},
    {'cat': 1, 'brand': 'HUSKY KLONDIKE', 'product': '3в1 глянцевая', 'facovka': '0,9 л',
     'query': 'HUSKY Klondike грунт-эмаль 3 в 1 черная RAL 9005 0.9 vseinstrumenti.ru'},
    {'cat': 1, 'brand': 'Tikkurila METALLISTA', 'product': '3в1', 'facovka': '0,8 л',
     'query': 'Tikkurila METALLISTA краска по ржавчине 3 в 1 RAL 9005 0.8 vseinstrumenti.ru'},

    # Категория 2: Быстросохнущая матовая, серая RAL 7040
    {'cat': 2, 'brand': 'ОРЕОЛ TITAN', 'product': 'быстросохнущая серая', 'facovka': '0,75 л (0,9 кг)',
     'query': 'ОРЕОЛ TITAN грунт-эмаль быстросохнущая матовая серая RAL 7040 0.75 vseinstrumenti.ru'},
    {'cat': 2, 'brand': 'Elcon', 'product': 'быстросохнущая 3в1', 'facovka': '0,8 кг',
     'query': 'Elcon быстросохнущая грунт-эмаль 3 в 1 RAL 7040 0.8 vseinstrumenti.ru'},
    {'cat': 2, 'brand': 'ЛАКРА PROF IT', 'product': '4в1', 'facovka': '0,8 кг',
     'query': 'Лакра PROF IT грунт-эмаль 4 в 1 серая 0.8 vseinstrumenti.ru'},
    {'cat': 2, 'brand': 'Новбытхим', 'product': '5в1', 'facovka': '0,75 л',
     'query': 'Новбытхим грунт-эмаль 5в1 RAL 7024 графитовый 0.75 vseinstrumenti.ru'},

    # Категория 3: Эмаль с молотковым эффектом, золотистая
    {'cat': 3, 'brand': 'Лакра', 'product': 'молотковая', 'facovka': '0,8 кг',
     'query': 'Лакра эмаль с молотковым эффектом золото 0.8 vseinstrumenti.ru'},
    {'cat': 3, 'brand': 'ОРЕОЛ TITAN', 'product': 'молотковая', 'facovka': '0,8 кг',
     'query': 'ОРЕОЛ TITAN эмаль с молотковым эффектом золотистая 0.8 vseinstrumenti.ru'},
    {'cat': 3, 'brand': 'Dufa Premium HAMMERLACK', 'product': 'молотковая', 'facovka': '750 мл',
     'query': 'Dufa Premium HAMMERLACK молотковая 750 мл vseinstrumenti.ru'},
    {'cat': 3, 'brand': 'Dali', 'product': '3в1 молотковая', 'facovka': '0,75 л',
     'query': 'Dali 3 в 1 молотковая черная 0.75 vseinstrumenti.ru'},
    {'cat': 3, 'brand': 'Elcon', 'product': 'кузнечная золото', 'facovka': '0,8 кг',
     'query': 'Elcon кузнечная краска молотковая золото 0.8 vseinstrumenti.ru'},
    {'cat': 3, 'brand': 'HUSKY KLONDIKE', 'product': 'золото', 'facovka': '0,9 л',
     'query': 'HUSKY Klondike краска по металлу молотковая золото 0.9 vseinstrumenti.ru'},

    # Категория 4: Лакра 5 в 1 Грунт-эмаль (НОВАЯ)
    {'cat': 4, 'brand': 'Лакра', 'product': '5в1 грунт-эмаль', 'facovka': '0,8 кг',
     'query': 'Лакра грунт-эмаль 5 в 1 0.8 vseinstrumenti.ru'},

    # Категория 5: Грунт-эмаль TOPGRADE (НОВАЯ)
    {'cat': 5, 'brand': 'TOPGRADE', 'product': 'грунт-эмаль', 'facovka': '0,75 л',
     'query': 'Topgrade by Hammerite грунт краска 3 в 1 0.75 vseinstrumenti.ru'},
]

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
    """Извлечь цены из сниппета.
       Корректно обрабатываем формат "X XXX ₽" с пробелом-разделителем.
       "1 210" → 1210 (не 121!).
    """
    if not snippet:
        return None, None

    # Ищем "от X рублей" или "от X ₽"
    m = re.search(r'(?:цен[ае]\s+)?от\s+(\d{1,3}(?:\s?\d{3})*)\s*(?:руб|₽)?', snippet, re.IGNORECASE)
    if m:
        try:
            val = int(m.group(1).replace(' ', ''))
            if 100 < val < 100000:
                return val, None
        except:
            pass

    # Ищем ВСЕ цены в формате "X XXX ₽" или "XXX ₽"
    # ВАЖНО: пробел между цифрами — это разделитель тысяч, а не разделитель цен!
    # Поэтому сначала ищем "X XXX ₽" (где X - цифры), потом уже "XXX" без ₽
    # Регекс: \d{1,3}(\s\d{3})+ — означает "1-3 цифры + пробел + 3 цифры (повторяется)"
    # Также захватываем одиночные 3-значные числа с ₽
    prices = []
    # Сначала ищем 4-значные и более цены с ₽ (с пробелом или без)
    for m in re.finditer(r'(\d{1,3}(?:\s\d{3})+|\d{3,4})\s*₽', snippet):
        try:
            val = int(m.group(1).replace(' ', ''))
            if 100 < val < 15000:
                prices.append(val)
        except:
            pass

    # Если ничего с ₽ — ищем "от X рублей"
    if not prices:
        m = re.search(r'цен[ае]\s+от\s+(\d{1,3}(?:\s?\d{3})*)', snippet, re.IGNORECASE)
        if m:
            try:
                val = int(m.group(1).replace(' ', ''))
                if 100 < val < 100000:
                    return val, None
            except:
                pass
        return None, None

    # Если только одна цена — это current
    if len(prices) == 1:
        return prices[0], None

    # Если 2+ цены
    p1, p2 = prices[0], prices[1]
    if 'вместо' in snippet.lower():
        # "Цена 1230 ₽ вместо. 1 3596 ₽" — p1=current, p2=old (но 13596 — это мусор, пропустим)
        # Если p2 слишком большая — берём как old только если p2 < 15000
        current_price = p1
        old_price = p2 if p2 < 15000 else None
    else:
        current_price = min(p1, p2)
        old_price = max(p1, p2)

    return current_price, old_price

# Сбор данных
results = []
print(f"{'#':<3} {'Cat':<4} {'Brand':<25} {'Текущая':<10} {'Старая':<10} {'Финал':<10} {'URL':<60}")
print("-" * 130)

for i, p in enumerate(products, 1):
    print(f"\n[{i}/{len(products)}] {p['brand']} {p['product']}")
    # 2 попытки поиска
    data = search_with_snippets(p['query'])
    if not data:
        # Попробуем упрощённый запрос
        data = search_with_snippets(f"{p['brand']} {p['product']} vseinstrumenti")

    product_url = None
    current_price = None
    old_price = None

    for r in data:
        url = r.get('url', '')
        snippet = r.get('snippet', '') or ''
        name = r.get('name', '') or ''

        # Ищем прямой URL товара
        if not product_url and 'vseinstrumenti.ru/product/' in url:
            product_url = url

        # Извлекаем цены
        if not current_price and not old_price:
            cur, old = extract_prices_from_snippet(snippet)
            if cur or old:
                current_price = cur
                old_price = old
                print(f"  URL: {url[:80]}")
                print(f"  snippet: {snippet[:200]}")
                print(f"  → cur={cur}, old={old}")

    # Финальная цена: приоритет old_price (постоянная, без скидки)
    final_price = old_price if old_price else current_price

    print(f"  → ИТОГ: {final_price} ₽ (old={old_price}, cur={current_price})")

    results.append({
        **p,
        'url': product_url or '',
        'current_price': current_price,
        'old_price': old_price,
        'final_price': final_price,
    })

    time.sleep(1.5)  # мягкий rate limit

# Сохраняем
with open('/home/z/my-project/scripts/vsi_prices_final.json', 'w', encoding='utf-8') as f:
    json.dump(results, f, ensure_ascii=False, indent=2)

print(f"\n\n=== ИТОГ ===")
found = sum(1 for r in results if r.get('final_price'))
print(f"Цен найдено: {found} из {len(results)}")
print(f"\n{'#':<3} {'Cat':<4} {'Brand':<25} {'Финал':<10} {'URL':<70}")
print("-" * 110)
for i, r in enumerate(results, 1):
    price = r.get('final_price')
    url = r.get('url', '')[:70]
    print(f"{i:<3} {r['cat']:<4} {r['brand']:<25} {price!s:<10} {url}")
