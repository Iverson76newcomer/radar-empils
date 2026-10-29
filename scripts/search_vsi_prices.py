"""Поиск товаров на vseinstrumenti.ru и извлечение цен."""
import subprocess
import json
import re
import time

# Список продуктов для поиска
products = [
    # Категория 1: Грунт-эмаль 3в1 высокоглянцевая, чёрный RAL 9005
    {'category': 1, 'brand': 'ОРЕОЛ TITAN', 'product': '3в1 глянцевая RAL 9005',
     'query': 'site:vseinstrumenti.ru ОРЕОЛ TITAN грунт-эмаль 3 в 1 RAL 9005 черный 0.75',
     'facovka': '0,75 л'},
    {'category': 1, 'brand': 'Elcon', 'product': '3в1 глянцевая',
     'query': 'site:vseinstrumenti.ru Elcon грунт-эмаль 3 в 1 глянцевая 0.8',
     'facovka': '0,8 кг'},
    {'category': 1, 'brand': 'Dufa CORREX', 'product': '3в1',
     'query': 'site:vseinstrumenti.ru Dufa CORREX грунт-эмаль 3 в 1 0.75',
     'facovka': '0,75 л'},
    {'category': 1, 'brand': 'Dali', 'product': '3в1',
     'query': 'site:vseinstrumenti.ru Dali грунт-эмаль 3 в 1 0.75',
     'facovka': '0,75 л'},
    {'category': 1, 'brand': 'HUSKY KLONDIKE', 'product': '3в1 глянцевая',
     'query': 'site:vseinstrumenti.ru HUSKY KLONDIKE грунт-эмаль 3 в 1 0.9',
     'facovka': '0,9 л'},
    {'category': 1, 'brand': 'Tikkurila METALLISTA', 'product': '3в1',
     'query': 'site:vseinstrumenti.ru Tikkurila METALLISTA грунт-эмаль 0.8',
     'facovka': '0,8 л'},

    # Категория 2: Быстросохнущая матовая, серая RAL 7040
    {'category': 2, 'brand': 'ОРЕОЛ TITAN', 'product': 'быстросохнущая серая',
     'query': 'site:vseinstrumenti.ru ОРЕОЛ TITAN грунт-эмаль быстросохнущая серая 0.75',
     'facovka': '0,75 л (0,9 кг)'},
    {'category': 2, 'brand': 'Elcon', 'product': 'быстросохнущая 3в1',
     'query': 'site:vseinstrumenti.ru Elcon быстросохнущая грунт-эмаль 0.8',
     'facovka': '0,8 кг'},
    {'category': 2, 'brand': 'ЛАКРА PROF IT', 'product': '4в1',
     'query': 'site:vseinstrumenti.ru Лакра PROF IT грунт-эмаль 4 в 1 0.8',
     'facovka': '0,8 кг'},
    {'category': 2, 'brand': 'Новбытхим', 'product': '5в1',
     'query': 'site:vseinstrumenti.ru Новбытхим грунт-эмаль 5 в 1 0.75',
     'facovka': '0,75 л'},

    # Категория 3: Эмаль с молотковым эффектом, золотистая
    {'category': 3, 'brand': 'Лакра', 'product': 'молотковая',
     'query': 'site:vseinstrumenti.ru Лакра эмаль молотковая золотистая 0.8',
     'facovka': '0,8 кг'},
    {'category': 3, 'brand': 'ОРЕОЛ TITAN', 'product': 'молотковая',
     'query': 'site:vseinstrumenti.ru ОРЕОЛ TITAN эмаль молотковая золотистая 0.8',
     'facovka': '0,8 кг'},
    {'category': 3, 'brand': 'Dufa Premium HAMMERLACK', 'product': 'молотковая',
     'query': 'site:vseinstrumenti.ru Dufa Premium HAMMERLACK 750',
     'facovka': '750 мл'},
    {'category': 3, 'brand': 'Dali', 'product': '3в1 молотковая',
     'query': 'site:vseinstrumenti.ru Dali молотковая 3 в 1 0.75',
     'facovka': '0,75 л'},
    {'category': 3, 'brand': 'Elcon', 'product': 'кузнечная золото',
     'query': 'site:vseinstrumenti.ru Elcon кузнечная золото 0.8',
     'facovka': '0,8 кг'},
    {'category': 3, 'brand': 'HUSKY KLONDIKE', 'product': 'золото',
     'query': 'site:vseinstrumenti.ru HUSKY KLONDIKE молотковая золото 0.9',
     'facovka': '0,9 л'},

    # Категория 4: НОВАЯ — Лакра 5 в 1 Грунт-эмаль
    {'category': 4, 'brand': 'Лакра', 'product': '5в1 грунт-эмаль',
     'query': 'site:vseinstrumenti.ru Лакра грунт-эмаль 5 в 1 0.8',
     'facovka': '0,8 кг'},

    # Категория 5: НОВАЯ — Грунт-эмаль TOPGRADE
    {'category': 5, 'brand': 'TOPGRADE', 'product': 'грунт-эмаль',
     'query': 'site:vseinstrumenti.ru TOPGRADE грунт-эмаль 3 в 1',
     'facovka': '—'},
]

print(f"Всего продуктов к поиску: {len(products)}")

def search_product(query):
    """Поиск через z-ai web_search, возвращает первый результат vseinstrumenti.ru"""
    args = json.dumps({"query": query, "num": 8})
    out = subprocess.run(['z-ai', 'function', '-n', 'web_search', '-a', args],
                         capture_output=True, text=True, timeout=60)
    try:
        output = out.stdout
        # z-ai выводит несколько строк логов, затем JSON-массив, затем финальную строку "🎉 Function invocation completed!"
        # Найдём начало JSON-массива (первая '[') и конец (последняя ']')
        json_start = output.find('[')
        # Последняя строка может содержать лишний текст, найдём последнюю ']'
        # Однако проще: ищем финальный ']'
        json_end = output.rfind(']')
        if json_start >= 0 and json_end > json_start:
            json_str = output[json_start:json_end+1]
            data = json.loads(json_str)
            # Ищем первый результат с vseinstrumenti.ru/product/
            for r in data:
                url = r.get('url', '')
                if 'vseinstrumenti.ru/product/' in url:
                    return r
            # Если не нашли /product/, берём любой vseinstrumenti.ru
            for r in data:
                url = r.get('url', '')
                if 'vseinstrumenti.ru' in url:
                    return r
    except Exception as e:
        print(f"  Search error: {e}")
        print(f"  Output (first 500 chars): {out.stdout[:500]}")
    return None

results = []
for p in products:
    print(f"\n[{p['category']}.{p['brand']}] {p['query'][:60]}...")
    r = search_product(p['query'])
    if r:
        url = r.get('url', '')
        name = r.get('name', '')[:80]
        snippet = r.get('snippet', '')[:150]
        print(f"  ✓ {url[:80]}")
        print(f"    name: {name}")
        print(f"    snippet: {snippet}")
        results.append({**p, 'url': url, 'name': name, 'snippet': snippet})
    else:
        print(f"  ✗ Не найдено")
        results.append({**p, 'url': '', 'name': '', 'snippet': ''})
    time.sleep(1)

# Сохраняем
with open('/home/z/my-project/scripts/vsi_products.json', 'w', encoding='utf-8') as f:
    json.dump(results, f, ensure_ascii=False, indent=2)

print(f"\n\n=== НАЙДЕНО {sum(1 for r in results if r['url'])} из {len(results)} ===")
print("Сохранено в vsi_products.json")
