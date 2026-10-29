"""Извлекаем количество результатов поиска из страниц Яндекса."""
import json, re, os, glob

for fpath in sorted(glob.glob('/tmp/ys_*.json')):
    query = os.path.basename(fpath).replace('ys_', '').replace('.json', '').replace('_', ' ')
    with open(fpath) as f:
        d = json.load(f)
    html = d.get('data', {}).get('html', '')
    # Ищем patterns типа "Нашлось X тыс. результатов" или "Результаты: X"
    # Также ищем meta data
    # Yandex search result page has: "Результатов: X" or "Нашлось X ответов"
    m = re.search(r'нашлось\s+<[^>]*>([\d\s]+)\s*(?:тыс\.|млн\.|млрд\.|ответов|результатов)?', html, re.IGNORECASE)
    if m:
        print(f"{query}: {m.group(0)[:100]}")
    else:
        m2 = re.search(r'(?:result.*?count|нашлось|нашлось\s+ответов)[^<]*?(\d[\d\s]*)', html, re.IGNORECASE)
        if m2:
            print(f"{query} (alt): {m2.group(0)[:100]}")
        else:
            # Look for any number pattern in title or meta
            title_match = re.search(r'<title>([^<]+)</title>', html)
            print(f"{query} | title: {title_match.group(1)[:100] if title_match else 'N/A'}")
