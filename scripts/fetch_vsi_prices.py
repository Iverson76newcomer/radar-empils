"""Загрузка страниц товаров vseinstrumenti.ru и извлечение цен.
   Правило: если цена перечёркнута (распродажа) — берём перечёркнутую (постоянную) цену.
"""
import json
import subprocess
import re
import time

with open('/home/z/my-project/scripts/vsi_products.json', 'r', encoding='utf-8') as f:
    products = json.load(f)

def fetch_page(url):
    """Получить HTML страницы через z-ai page_reader."""
    args = json.dumps({"url": url})
    out = subprocess.run(['z-ai', 'function', '-n', 'page_reader', '-a', args],
                         capture_output=True, text=True, timeout=120)
    try:
        output = out.stdout
        # Ищем JSON-объект (начинается с '{')
        json_start = output.find('{')
        if json_start >= 0:
            json_str = output[json_start:]
            # Ищем последнюю '}'
            json_end = json_str.rfind('}')
            if json_end > 0:
                data = json.loads(json_str[:json_end+1])
                return data.get('data', {}).get('html', '')
    except Exception as e:
        print(f"  Error fetching: {e}")
    return ''

def extract_prices(html):
    """Извлечь цены из HTML страницы vseinstrumenti.ru.
       Приоритет: если есть зачёркнутая цена (старая) — берём её (это постоянная цена без скидки).
       Иначе — берём обычную цену.
    """
    prices = {'old_price': None, 'current_price': None, 'all_prices': []}

    # Ищем все цены в формате "X XXX ₽" или "X ₽"
    # В HTML обычно:
    # <span class="...old">3 972 ₽</span> — перечёркнутая
    # <span class="...price...">2 590 ₽</span> — текущая (со скидкой)

    # Шаблон 1: "old-price" в HTML — зачёркнутая
    # Ищем все совпадения цифр в формате "X XXX ₽" или "XX XXX ₽"
    all_prices_match = re.findall(r'(\d{1,3}(?:\s?\d{3})*)\s*₽', html)
    if all_prices_match:
        prices['all_prices'] = [int(p.replace(' ', '')) for p in all_prices_match if int(p.replace(' ', '')) > 50]

    # Ищем "previousPrice", "oldPrice", "old-price" и т.п.
    # В vseinstrumenti.ru часто используется data-атрибут или класс
    old_price_patterns = [
        r'(?:previousPrice|oldPrice|old-price|price-old|old_price|crossedOut|original|basePrice)["\'>\s:]+(?:\d+\s*)*"?(\d{1,3}(?:\s?\d{3})*)\s*₽?',
        r'(?:data-old-price|data-previous-price|data-original-price)=["\'](\d+)["\']',
        r'<(?:s|del|strike)[^>]*>(\d{1,3}(?:\s?\d{3})*)\s*₽',
        r'(?:class|data-[\w-]+)=["\'][^"\']*old[^"\']*["\'][^>]*>\s*(\d{1,3}(?:\s?\d{3})*)\s*₽',
    ]
    for pattern in old_price_patterns:
        matches = re.findall(pattern, html, re.IGNORECASE)
        if matches:
            for m in matches:
                try:
                    p = int(m.replace(' ', ''))
                    if 100 < p < 100000:  # фильтруем странные значения
                        prices['old_price'] = p
                        break
                except:
                    pass
            if prices['old_price']:
                break

    # Если нашли old_price — это постоянная цена
    # Если нет — нужно найти current_price (текущую регулярную цену)
    # Ищем цену рядом со словом "price"
    current_patterns = [
        r'(?:currentPrice|priceValue|price-value|price-actual|productPrice|price_current)["\'>\s:]+(?:\d+\s*)*"?(\d{1,3}(?:\s?\d{3})*)\s*₽?',
        r'data-price=["\'](\d+)["\']',
        r'data-product-price=["\'](\d+)["\']',
    ]
    for pattern in current_patterns:
        matches = re.findall(pattern, html, re.IGNORECASE)
        if matches:
            for m in matches:
                try:
                    p = int(m.replace(' ', ''))
                    if 100 < p < 100000:
                        prices['current_price'] = p
                        break
                except:
                    pass
            if prices['current_price']:
                break

    return prices

results = []
for i, p in enumerate(products, 1):
    url = p.get('url', '')
    print(f"\n[{i}/{len(products)}] [{p['category']}.{p['brand']}] {p['product']}")
    if not url or 'vseinstrumenti.ru/product/' not in url:
        # URL не ведёт на конкретный товар — попробуем найти через web_search ещё раз
        # или использовать search snippet
        print(f"  ⚠ URL не на товар: {url}")
        # Попробуем найти цену в snippet
        snippet = p.get('snippet', '')
        prices_match = re.findall(r'(\d{1,3}(?:\s?\d{3})*)\s*₽', snippet)
        if prices_match:
            price_val = int(prices_match[0].replace(' ', ''))
            print(f"  ✓ Цена из snippet: {price_val} ₽")
            results.append({**p, 'price': price_val, 'source': 'snippet'})
        else:
            # URL ведёт на главную или категорию — нужно найти прямой URL через web_search
            # Уточним запрос
            args = json.dumps({"query": f"{p['brand']} {p['product']} vseinstrumenti.ru купить", "num": 5})
            out = subprocess.run(['z-ai', 'function', '-n', 'web_search', '-a', args],
                                 capture_output=True, text=True, timeout=60)
            try:
                output = out.stdout
                json_start = output.find('[')
                json_end = output.rfind(']')
                if json_start >= 0 and json_end > json_start:
                    data = json.loads(output[json_start:json_end+1])
                    for r in data:
                        u = r.get('url', '')
                        if 'vseinstrumenti.ru/product/' in u:
                            url = u
                            print(f"  ✓ Найден URL: {url[:80]}")
                            break
            except Exception as e:
                print(f"  Search retry error: {e}")

    # Теперь загружаем страницу товара
    if url and 'vseinstrumenti.ru/product/' in url:
        html = fetch_page(url)
        if html:
            prices = extract_prices(html)
            # Определяем итоговую цену (БЕЗ скидки = old_price если есть, иначе current_price)
            final_price = prices['old_price'] or prices['current_price']
            if final_price is None and prices['all_prices']:
                # Берём самую большую цену как базовую (вероятно это old_price)
                final_price = max(prices['all_prices'])
            print(f"  → Старая цена: {prices['old_price']} ₽ | Текущая: {prices['current_price']} ₽ | Все: {sorted(set(prices['all_prices']), reverse=True)[:5]}")
            print(f"  → ИТОГОВАЯ цена (без скидки): {final_price} ₽")
            results.append({**p, 'url': url, 'price': final_price,
                            'old_price': prices['old_price'], 'current_price': prices['current_price'],
                            'source': 'page'})
        else:
            print(f"  ✗ Не удалось загрузить HTML")
            results.append({**p, 'price': None, 'source': 'no_html'})
    else:
        # Если URL так и не нашли — добавим без цены
        if not any(r.get('brand') == p['brand'] and r.get('category') == p['category'] for r in results):
            results.append({**p, 'price': None, 'source': 'no_url'})

    time.sleep(1)

# Сохраняем
with open('/home/z/my-project/scripts/vsi_prices_final.json', 'w', encoding='utf-8') as f:
    json.dump(results, f, ensure_ascii=False, indent=2)

print(f"\n\n=== ИТОГ ===")
found = sum(1 for r in results if r.get('price'))
print(f"Цен найдено: {found} из {len(results)}")
for r in results:
    price = r.get('price')
    if price:
        print(f"  [{r['category']}] {r['brand']} {r['product']}: {price} ₽ (источник: {r.get('source')})")
    else:
        print(f"  [{r['category']}] {r['brand']} {r['product']}: НЕТ ЦЕНЫ")
