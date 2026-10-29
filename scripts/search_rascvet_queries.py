"""Анализ результатов web_search по 7 запросам 'Расцвет ...'.
   Подсчитываем % результатов, связанных с брендом ЛКМ 'Расцвет' (не с обычным словом 'расцвет').
"""
import json
import subprocess
import time

queries = [
    ("Расцвет краска", "краска"),
    ("Расцвет грунт-эмаль", "грунт-эмаль"),
    ("Расцвет 3 в 1", "3 в 1"),
    ("Расцвет по металлу", "по металлу"),
    ("Расцвет по ржавчине", "по ржавчине"),
    ("Расцвет молотковая", "молотковая"),
    ("Расцвет быстросохнущая", "быстросохнущая"),
]

# Брендовые URL маркеры
BRAND_MARKERS = ['rascvet.pro', 'rastsvet', 'rascvet', 'расцвет', 'empils.ru/brands/rastsvet']

results = {}
for query, key in queries:
    print(f"\n=== {query} ===")
    args = json.dumps({"query": query, "num": 10})
    out = subprocess.run(['z-ai', 'function', '-n', 'web_search', '-a', args],
                          capture_output=True, text=True, timeout=60)
    try:
        # Парсим вывод — z-ai выводит сперва несколько строк логов, затем JSON
        output = out.stdout
        # Найдём начало JSON массива
        json_start = output.find('[')
        if json_start >= 0:
            json_str = output[json_start:]
            data = json.loads(json_str)
        else:
            data = []
    except Exception as e:
        print(f"  ERROR parsing: {e}")
        data = []

    brand_count = 0
    total = len(data)
    print(f"  Total results: {total}")
    for r in data:
        url = r.get('url', '').lower()
        name = r.get('name', '')
        snippet = r.get('snippet', '')
        # Проверяем, связан ли результат с брендом ЛКМ "Расцвет"
        is_brand = ('rascvet' in url or 'rastsvet' in url or
                    'расцвет' in name.lower()[:30] or  # в начале названия
                    'rascvet' in name.lower() or
                    'empils.ru/brands/rastsvet' in url or
                    'vseinstrumenti.ru' in url and 'rastsvet' in url or
                    '12.kz/manufacturers/rastsvet' in url or
                    'rascvet' in snippet.lower()[:100])
        if is_brand:
            brand_count += 1
            print(f"  ✓ {url[:70]} | {name[:60]}")
        else:
            print(f"  ✗ {url[:70]} | {name[:60]}")
    results[key] = {'total': total, 'brand_related': brand_count, 'brand_ratio': brand_count / max(total, 1)}
    time.sleep(1)

# Вывод итогов
print("\n\n=== ИТОГ ===")
print(f"{'Тип запроса':<20} | {'Бренд-результаты':<20} | {'Всего':<6} | {'Доля бренда':<10}")
print("-" * 70)
for key, r in results.items():
    print(f"{key:<20} | {r['brand_related']:<20} | {r['total']:<6} | {r['brand_ratio']:.0%}")

# Сохраняем
with open('/home/z/my-project/scripts/rascvet_search_results.json', 'w', encoding='utf-8') as f:
    json.dump(results, f, ensure_ascii=False, indent=2)
print("\nСохранено в rascvet_search_results.json")
