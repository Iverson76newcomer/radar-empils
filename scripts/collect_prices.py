#!/usr/bin/env python3
"""
Сбор актуальных цен на для 16 SKU из презентации TITAN.
Стратегия: 
1. Web search "site:vseinstrumenti.ru <товар>"
2. Извлечение цены из сниппета/заголовка
3. Если не нашли — попробовать открыть URL через web-reader
"""

import subprocess
import json
import re
import time
from pathlib import Path

# Список товаров (как в Excel)
products = [
    # 3в1 по ржавчине (глянцевая) — слайд 11
    {'n': 1, 'cat': '3в1', 'brand': 'TITAN', 'product': 'Грунт-эмаль по ржавчине Empils ОРЕОЛ TITAN чёрный RAL 9005', 'vol': '0.75 л', 'price_old': 872,
     'query': 'ОРЕОЛ TITAN грунт-эмаль по ржавчине чёрный RAL 9005 0.75 л'},
    {'n': 2, 'cat': '3в1', 'brand': 'Elcon', 'product': 'Грунт-эмаль 3 в 1 Elcon глянцевая чёрная RAL 9005', 'vol': '0.8 кг', 'price_old': 958,
     'query': 'Elcon грунт-эмаль 3 в 1 глянцевая RAL 9005 0.8 кг'},
    {'n': 3, 'cat': '3в1', 'brand': 'Dufa', 'product': 'Грунт-эмаль Dufa CORREX 3 в 1 RAL 9005', 'vol': '0.75 л', 'price_old': 1020,
     'query': 'Dufa CORREX 3 в 1 грунт-эмаль RAL 9005 0.75'},
    {'n': 4, 'cat': '3в1', 'brand': 'Dali', 'product': 'Грунт-эмаль по ржавчине Dali 3 в 1, чёрная RAL 9005', 'vol': '0.75 л', 'price_old': 1055,
     'query': 'Dali 3 в 1 грунт-эмаль по ржавчине RAL 9005 0.75'},
    {'n': 5, 'cat': '3в1', 'brand': 'HUSKY KLONDIKE', 'product': 'Краска по металлу HUSKY KLONDIKE глянцевая, чёрная, RAL 9005', 'vol': '0.9 л', 'price_old': 1153,
     'query': 'HUSKY KLONDIKE краска по металлу глянцевая RAL 9005 0.9'},
    {'n': 6, 'cat': '3в1', 'brand': 'Tikkurila', 'product': 'Краска для металла Tikkurila METALLISTA 3 в 1, чёрная', 'vol': '0.8 л', 'price_old': 2085,
     'query': 'Tikkurila METALLISTA 3 в 1 краска по металлу 0.8 чёрная'},
    # Быстросохнущая (матовая) — слайд 12
    {'n': 7, 'cat': 'Быстросохнущая', 'brand': 'Новбытхим', 'product': 'Грунт-эмаль 5 в 1 Новбытхим, серый, RAL 7040', 'vol': '0.75 л', 'price_old': 765,
     'query': 'Новбытхим грунт-эмаль 5 в 1 RAL 7040 0.75'},
    {'n': 8, 'cat': 'Быстросохнущая', 'brand': 'TITAN', 'product': 'Грунт-эмаль быстросохнущая ОРЕОЛ TITAN матовая серая', 'vol': '0.75 л', 'price_old': 765,
     'query': 'ОРЕОЛ TITAN быстросохнущая матовая серая 0.75'},
    {'n': 9, 'cat': 'Быстросохнущая', 'brand': 'Elcon', 'product': 'Быстросохнущая грунт-эмаль Elcon 3 в 1, матовая, серая', 'vol': '0.8 кг', 'price_old': 838,
     'query': 'Elcon быстросохнущая грунт-эмаль 3 в 1 матовая серая 0.8'},
    {'n': 10, 'cat': 'Быстросохнущая', 'brand': 'Лакра', 'product': 'Грунт-эмаль ЛАКРА PROF IT 4 в 1, быстросохнущая, графит, RAL 7024', 'vol': '0.8 кг', 'price_old': 906,
     'query': 'Лакра PROF IT 4 в 1 быстросохнущая RAL 7024 0.8'},
    # Молотковая — слайд 13
    {'n': 11, 'cat': 'Молотковая', 'brand': 'Dali', 'product': 'Грунт-эмаль Dali 3 в 1 молотковая золотистая', 'vol': '0.75 л', 'price_old': 1057,
     'query': 'Dali молотковая золотистая грунт-эмаль 3 в 1 0.75'},
    {'n': 12, 'cat': 'Молотковая', 'brand': 'Elcon', 'product': 'Кузнечная краска Elcon молотковая золото', 'vol': '0.8 кг', 'price_old': 1097,
     'query': 'Elcon кузнечная краска молотковая золото 0.8'},
    {'n': 13, 'cat': 'Молотковая', 'brand': 'HUSKY KLONDIKE', 'product': 'Краска HUSKY KLONDIKE молотковая золото', 'vol': '0.9 л', 'price_old': 1196,
     'query': 'HUSKY KLONDIKE молотковая золото краска по металлу 0.9'},
    {'n': 14, 'cat': 'Молотковая', 'brand': 'TITAN', 'product': 'Эмаль TITAN молотковая алкидно-стирольная, золотистая', 'vol': '0.8 кг', 'price_old': 1308,
     'query': 'TITAN молотковая алкидно-стирольная золотистая 0.8'},
    {'n': 15, 'cat': 'Молотковая', 'brand': 'Лакра', 'product': 'Эмаль Лакра молотковая золото', 'vol': '0.8 кг', 'price_old': 1341,
     'query': 'Лакра молотковая золото эмаль 0.8'},
    {'n': 16, 'cat': 'Молотковая', 'brand': 'Dufa', 'product': 'Эмаль Dufa Premium HAMMERLACK молотковая, золотая', 'vol': '0.75 л', 'price_old': 1653,
     'query': 'Dufa Premium HAMMERLACK молотковая золотая 750 мл'},
]


def search_web(query, num=8):
    """Поиск через z-ai CLI. Возвращает список результатов."""
    try:
        cmd = ['z-ai', 'function', '-n', 'web_search', '-a', json.dumps({'query': query, 'num': num})]
        r = subprocess.run(cmd, capture_output=True, text=True, timeout=60)
        if r.returncode != 0:
            return []
        output = r.stdout
        # Ищем первый '[' после приглашений z-ai
        idx = output.find('[')
        if idx < 0:
            return []
        # Ищем конец массива — последний ']' в выводе
        end_idx = output.rfind(']')
        if end_idx < idx:
            return []
        json_str = output[idx:end_idx+1]
        try:
            return json.loads(json_str)
        except json.JSONDecodeError:
            # Возможно, в выводе есть trailing символы после JSON
            # Пробуем построчно найти валидный JSON
            try:
                return json.loads(json_str.split(']')[0] + ']')
            except:
                return []
    except Exception as e:
        print(f'  Ошибка: {e}')
        return []


def extract_price(text):
    """Извлекает ВСЕ цены из текста. Возвращает список int (отсортированный по убыванию)."""
    if not text:
        return []
    prices = []
    # Ищем все паттерны цен
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
                if 200 < p < 10000:  # реалистичный диапазон для грунт-эмалей
                    prices.append(p)
            except:
                continue
    # Если нашли несколько цен — часто первая акционная, вторая обычная
    # Возвращаем МАКСИМАЛЬНУЮ (обычная цена, без акции)
    return sorted(set(prices), reverse=True)


def process_product(p):
    """Ищет цену для одного товара."""
    print(f"\n[{p['n']}/16] {p['brand']} — {p['product'][:50]}...")
    
    results = search_web(p['query'], num=10)
    if not results:
        print(f"  ❌ Нет результатов поиска")
        return None
    
    # Ищем в результатах те, что с
    found = []
    for r in results:
        url = r.get('url', '')
        host = r.get('host_name', '')
        if 'vseinstrumenti' in host.lower() or 'vseinstrumenti' in url.lower():
            found.append(r)
    
    # Ищем цену в сниппетах и заголовках
    for r in found[:5]:
        snippet = r.get('snippet', '')
        name = r.get('name', '')
        combined = f"{name} {snippet}"
        
        prices = extract_price(combined)
        if prices:
            # Берём максимальную (обычная цена, не акционная)
            price = prices[0]
            url = r.get('url', '')
            print(f"  ✅ Цена: {price} ₽ (все найденные: {prices})")
            print(f"     URL: {url[:80]}")
            return {'price': price, 'url': url, 'source': 'snippet'}
    
    # Если не нашли в сниппетах — выводим все найденные URL для ручной проверки
    print(f"  ⚠️ Цена не извлечена, найдено URL: {len(found)}")
    for r in found[:3]:
        print(f"     {r.get('host_name', '')[:30]} | {r.get('name', '')[:60]}")
        print(f"       snippet: {r.get('snippet', '')[:120]}")
    return None


# Обработка всех 16 товаров
print("="*80)
print("СБОР ЦЕН С")
print("="*80)

results = []
for p in products:
    res = process_product(p)
    results.append({
        'n': p['n'],
        'brand': p['brand'],
        'product': p['product'],
        'vol': p['vol'],
        'price_old': p['price_old'],
        'price_new': res['price'] if res else None,
        'url': res['url'] if res else '',
        'source': res['source'] if res else 'not_found',
    })
    time.sleep(1)  # пауза, чтобы не перегружать

# Сохранение результатов в JSON для дальнейшей обработки
out_json = '/home/z/my-project/scripts/prices_results.json'
with open(out_json, 'w', encoding='utf-8') as f:
    json.dump(results, f, ensure_ascii=False, indent=2)

print("\n" + "="*80)
print("ИТОГИ СБОРА")
print("="*80)
found_count = sum(1 for r in results if r['price_new'])
print(f"Найдено цен: {found_count} из {len(results)}")
print(f"Результаты сохранены: {out_json}")
print()
print(f"{'N':<3} | {'Бренд':<15} | {'Цена 10.06':<10} | {'Цена акт.':<10} | {'Δ':<8}")
print('-' * 60)
for r in results:
    pn = r['price_new'] if r['price_new'] else '—'
    delta = f"{r['price_new']-r['price_old']:+d}" if r['price_new'] else '—'
    print(f"{r['n']:<3} | {r['brand']:<15} | {r['price_old']:<10} | {str(pn):<10} | {delta:<8}")
