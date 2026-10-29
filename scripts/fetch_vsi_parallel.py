"""Параллельная загрузка всех страниц товаров VSI и извлечение цен."""
import json
import subprocess
import re
import os
from concurrent.futures import ThreadPoolExecutor, as_completed

with open('/home/z/my-project/scripts/vsi_products.json', 'r', encoding='utf-8') as f:
    products = json.load(f)

# Для каждого продукта пере-сформируем query и найдём прямой URL товара через web_search
def find_product_url(brand, product, base_query):
    """Найти прямой URL товара на vseinstrumenti.ru/product/"""
    args = json.dumps({"query": base_query, "num": 10})
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
                    return u, r.get('name', ''), r.get('snippet', '')
    except Exception as e:
        pass
    return None, None, None

def fetch_page(url):
    """Получить HTML страницы через z-ai page_reader."""
    args = json.dumps({"url": url})
    out = subprocess.run(['z-ai', 'function', '-n', 'page_reader', '-a', args],
                         capture_output=True, text=True, timeout=120)
    try:
        output = out.stdout
        json_start = output.find('{')
        if json_start >= 0:
            json_str = output[json_start:]
            json_end = json_str.rfind('}')
            if json_end > 0:
                data = json.loads(json_str[:json_end+1])
                return data.get('data', {}).get('html', '')
    except Exception as e:
        pass
    return ''

def extract_price_from_html(html):
    """Извлечь цену из HTML страницы vseinstrumenti.ru.
       Правило: если есть перечёркнутая (old) цена — берём её (постоянная, без скидки).
       Иначе — берём обычную.
    """
    if not html:
        return None, None, []

    old_price = None
    current_price = None

    # Шаблон 1: ищем зачёркнутую цену через теги <s>, <del>, <strike>
    for tag in ['s', 'del', 'strike']:
        m = re.search(rf'<{tag}[^>]*>([^<]*?)(\d{{1,3}}(?:\s?\d{{3}})*)\s*₽?\s*<', html, re.IGNORECASE)
        if m:
            try:
                p = int(m.group(2).replace(' ', ''))
                if 100 < p < 100000:
                    old_price = p
                    break
            except:
                pass

    # Шаблон 2: классы oldPrice, old-price, price-old, previousPrice
    if not old_price:
        patterns = [
            r'(?:oldPrice|old-price|price-old|previousPrice|crossedOut|originalPrice|basePrice|fullPrice|regularPrice)["\'>\s:]+["\']?(\d{1,3}(?:\s?\d{3})*)\s*₽?',
            r'data-(?:old-price|previous-price|original-price|full-price)=["\'](\d{1,3}(?:\s?\d{3})*)["\']',
            r'class=["\'][^"\']*\b(old|previous|crossed|original)\b[^"\']*["\'][^>]*>\s*(\d{1,3}(?:\s?\d{3})*)\s*₽',
            # JSON внутри HTML
            r'"(?:oldPrice|previousPrice|originalPrice|fullPrice|regularPrice|basePrice)":\s*"?(\d{1,3}(?:\s?\d{3})*)"?'
        ]
        for pattern in patterns:
            matches = re.findall(pattern, html, re.IGNORECASE)
            if matches:
                for m in matches:
                    val = m if isinstance(m, str) else m[-1] if isinstance(m, tuple) else m
                    try:
                        p = int(val.replace(' ', ''))
                        if 100 < p < 100000:
                            old_price = p
                            break
                    except:
                        pass
                if old_price:
                    break

    # Шаблон 3: текущая цена (если нет old_price, берём её)
    patterns_current = [
        r'(?:currentPrice|priceValue|price-value|price-actual|productPrice|price_current|price_current_vat)["\'>\s:]+["\']?(\d{1,3}(?:\s?\d{3})*)\s*₽?',
        r'data-(?:price|product-price|current-price)=["\'](\d{1,3}(?:\s?\d{3})*)["\']',
        r'class=["\'][^"\']*\b(price-current|current-price|actual-price)\b[^"\']*["\'][^>]*>\s*(\d{1,3}(?:\s?\d{3})*)\s*₽',
        r'"(?:currentPrice|priceValue|price|productPrice|current_price)":\s*"?(\d{1,3}(?:\s?\d{3})*)"?'
    ]
    for pattern in patterns_current:
        matches = re.findall(pattern, html, re.IGNORECASE) if False else re.findall(patterns_current[0], html, re.IGNORECASE)
        if matches:
            for m in matches:
                try:
                    p = int(m.replace(' ', ''))
                    if 100 < p < 100000:
                        current_price = p
                        break
                except:
                    pass
            if current_price:
                break

    # Все цены (для отладки)
    all_prices_match = re.findall(r'(\d{1,3}(?:\s?\d{3})*)\s*₽', html)
    all_prices = [int(p.replace(' ', '')) for p in all_prices_match if 100 < int(p.replace(' ', '')) < 100000]

    return old_price, current_price, all_prices

# Шаг 1: находим прямые URL товаров для всех продуктов (параллельно)
print("=== ШАГ 1: Поиск прямых URL товаров ===\n")
product_urls = {}
def find_one(p):
    url, name, snippet = find_product_url(p['brand'], p['product'], p['query'])
    return p['category'], p['brand'], p['product'], url, name, snippet

with ThreadPoolExecutor(max_workers=5) as executor:
    futures = [executor.submit(find_one, p) for p in products]
    for f in as_completed(futures):
        cat, brand, product, url, name, snippet = f.result()
        key = f"{cat}_{brand}_{product}"
        product_urls[key] = {'url': url, 'name': name, 'snippet': snippet}
        if url:
            print(f"  ✓ [{cat}.{brand}] {url[:80]}")
        else:
            print(f"  ✗ [{cat}.{brand}] — URL не найден")

# Сохраняем URL
with open('/home/z/my-project/scripts/vsi_product_urls.json', 'w', encoding='utf-8') as f:
    json.dump(product_urls, f, ensure_ascii=False, indent=2)

print(f"\n=== ШАГ 2: Загрузка страниц и извлечение цен ===\n")

# Шаг 2: загружаем страницы и извлекаем цены (параллельно)
def fetch_one(p):
    key = f"{p['category']}_{p['brand']}_{p['product']}"
    url_info = product_urls.get(key, {})
    url = url_info.get('url')
    if not url or 'vseinstrumenti.ru/product/' not in url:
        return p, None, None, None, None
    html = fetch_page(url)
    old_price, current_price, all_prices = extract_price_from_html(html)
    return p, url, old_price, current_price, all_prices

results = []
with ThreadPoolExecutor(max_workers=5) as executor:
    futures = [executor.submit(fetch_one, p) for p in products]
    for f in as_completed(futures):
        p, url, old_price, current_price, all_prices = f.result()
        if url is None:
            # Если URL не нашли — пробуем извлечь цену из snippet
            key = f"{p['category']}_{p['brand']}_{p['product']}"
            snippet = product_urls.get(key, {}).get('snippet', '')
            prices_match = re.findall(r'(\d{1,3}(?:\s?\d{3})*)\s*₽', snippet)
            if prices_match:
                price_val = int(prices_match[0].replace(' ', ''))
                print(f"  ⚠ [{p['category']}.{p['brand']}] цена из snippet: {price_val} ₽")
                results.append({**p, 'price': price_val, 'source': 'snippet', 'url': product_urls.get(key, {}).get('url', '')})
            else:
                print(f"  ✗ [{p['category']}.{p['brand']}] — нет URL, нет цены")
                results.append({**p, 'price': None, 'source': 'no_data', 'url': ''})
            continue

        # Определяем итоговую цену: приоритет — old_price (постоянная, без скидки)
        final_price = old_price if old_price else current_price
        if final_price is None and all_prices:
            final_price = max(all_prices)  # берём наибольшую (вероятно, постоянную)

        print(f"  → [{p['category']}.{p['brand']}] old={old_price} | current={current_price} | all={sorted(set(all_prices), reverse=True)[:5]} → final: {final_price}")
        results.append({**p, 'url': url, 'price': final_price,
                        'old_price': old_price, 'current_price': current_price,
                        'source': 'page'})

# Сохраняем
with open('/home/z/my-project/scripts/vsi_prices_final.json', 'w', encoding='utf-8') as f:
    json.dump(results, f, ensure_ascii=False, indent=2)

print(f"\n=== ИТОГ ===")
found = sum(1 for r in results if r.get('price'))
print(f"Цен найдено: {found} из {len(results)}")
