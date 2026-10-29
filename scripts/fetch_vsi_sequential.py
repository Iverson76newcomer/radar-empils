"""Последовательный сбор цен через page_reader.
   Сначала для продуктов с уже найденными URL (vsi_products.json).
"""
import json
import subprocess
import re
import time
import os

with open('/home/z/my-project/scripts/vsi_products.json', 'r', encoding='utf-8') as f:
    products = json.load(f)

def fetch_page(url):
    args = json.dumps({"url": url})
    out = subprocess.run(['z-ai', 'function', '-n', 'page_reader', '-a', args],
                         capture_output=True, text=True, timeout=180)
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
        print(f"    fetch error: {e}")
    return ''

def extract_prices(html):
    """Извлечь старую (постоянную) и текущую (со скидкой) цены."""
    if not html:
        return None, None, []

    old_price = None
    current_price = None

    # 1. JSON-LD: <script type="application/ld+json">{"offers":{"price":"1294"}}
    m = re.search(r'<script[^>]*type=["\']application/ld\+json["\'][^>]*>(.*?)</script>',
                  html, re.DOTALL | re.IGNORECASE)
    if m:
        try:
            import json as _json
            data = _json.loads(m.group(1))
            offers = data.get('offers', {})
            if isinstance(offers, dict):
                price_str = offers.get('price', '')
                if price_str:
                    current_price = int(float(price_str))
        except:
            pass

    # 2. Ищем все paternы old_price
    old_patterns = [
        r'(?:oldPrice|previousPrice|originalPrice|fullPrice|regularPrice|basePrice|price_old|price-old)[^"]*"?\s*[:=]\s*"?\s*(\d{1,3}(?:\s?\d{3})*)',
        r'data-(?:old-price|previous-price|original-price|full-price)=["\'](\d{1,3}(?:\s?\d{3})*)["\']',
        r'<(?:s|del|strike)[^>]*>\s*(\d{1,3}(?:\s?\d{3})*)\s*₽\s*</',
        r'class=["\'][^"\']*\b(old|previous|crossed|original|full)\b[^"\']*["\'][^>]*>\s*(\d{1,3}(?:\s?\d{3})*)\s*₽',
    ]
    for pat in old_patterns:
        matches = re.findall(pat, html, re.IGNORECASE)
        if matches:
            for m in matches:
                val = m if isinstance(m, str) else (m[-1] if isinstance(m, tuple) else m)
                try:
                    p = int(val.replace(' ', ''))
                    if 100 < p < 100000:
                        old_price = p
                        break
                except:
                    pass
            if old_price:
                break

    # 3. Текущая цена
    cur_patterns = [
        r'"(?:currentPrice|priceValue|price|productPrice|current_price)":\s*"?(\d{1,3}(?:\s?\d{3})*)"?',
        r'data-(?:price|product-price|current-price)=["\'](\d{1,3}(?:\s?\d{3})*)["\']',
        r'class=["\'][^"\']*\b(price-current|current-price|actual-price|price-value)\b[^"\']*["\'][^>]*>\s*(\d{1,3}(?:\s?\d{3})*)\s*₽',
    ]
    for pat in cur_patterns:
        matches = re.findall(pat, html, re.IGNORECASE)
        if matches:
            for m in matches:
                val = m if isinstance(m, str) else (m[-1] if isinstance(m, tuple) else m)
                try:
                    p = int(val.replace(' ', ''))
                    if 100 < p < 100000:
                        current_price = p
                        break
                except:
                    pass
            if current_price:
                break

    # 4. Все цены (отладка)
    all_prices_match = re.findall(r'(\d{1,3}(?:\s?\d{3})*)\s*₽', html)
    all_prices = [int(p.replace(' ', '')) for p in all_prices_match if 100 < int(p.replace(' ', '')) < 100000]

    return old_price, current_price, all_prices


# Сохраняем промежуточные результаты в файл, чтобы не терять прогресс
output_file = '/home/z/my-project/scripts/vsi_prices_final.json'
results = []
if os.path.exists(output_file):
    with open(output_file) as f:
        results = json.load(f)

processed = {(r.get('category'), r.get('brand'), r.get('product')) for r in results}

print(f"Уже обработано: {len(processed)}")

for i, p in enumerate(products, 1):
    key = (p['category'], p['brand'], p['product'])
    if key in processed:
        continue

    url = p.get('url', '')
    print(f"\n[{i}/{len(products)}] [{p['category']}.{p['brand']}] {p['product']}")

    # Если URL ведёт на главный сайт — нужен повторный поиск
    if not url or 'vseinstrumenti.ru/product/' not in url:
        print(f"  ⚠ URL не на товар, ищем...")
        # Сохраняем без цены
        results.append({**p, 'price': None, 'source': 'no_url'})
        # Записываем промежуточный результат
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(results, f, ensure_ascii=False, indent=2)
        continue

    # Загружаем страницу
    html = fetch_page(url)
    if not html:
        print(f"  ✗ Не удалось загрузить HTML")
        results.append({**p, 'price': None, 'source': 'no_html'})
    else:
        old_price, current_price, all_prices = extract_prices(html)
        # Приоритет: old_price (постоянная, без скидки), иначе current_price
        final_price = old_price if old_price else current_price
        if final_price is None and all_prices:
            final_price = max(all_prices)
        print(f"  old={old_price} | current={current_price} | all={sorted(set(all_prices), reverse=True)[:5]}")
        print(f"  → ИТОГ: {final_price} ₽")
        results.append({**p, 'price': final_price, 'old_price': old_price,
                       'current_price': current_price, 'source': 'page'})

    # Сохраняем после каждого продукта
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(results, f, ensure_ascii=False, indent=2)
    time.sleep(2)

print(f"\n=== ИТОГ ===")
print(f"Цен найдено: {sum(1 for r in results if r.get('price'))} из {len(results)}")
