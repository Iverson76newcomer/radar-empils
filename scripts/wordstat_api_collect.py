"""
Yandex Wordstat API — сбор данных по 25 ключевым фразам для TITAN.
3 метода на каждую фразу: TopRequests, RegionsDistribution, Dynamics.
"""
import requests
import json
import time
import os

API_KEY = 'AQVNyBi4wPgqGtaKV9-dvavzCH5TSaUNwBPvAErO'
FOLDER_ID = 'b1g0bo7majef8lq3iii2'
API_BASE = 'https://searchapi.api.cloud.yandex.net'

HEADERS = {
    'Authorization': f'Api-Key {API_KEY}',
    'Content-Type': 'application/json',
    'Accept-Language': 'ru',
}

# 25 ключевых фраз
KEYWORDS = [
    # Брендовые
    'TITAN краска',
    'Ореол TITAN',
    'Эмпилс',
    'краска TITAN по металлу',
    'грунт-эмаль TITAN',
    # Продуктовые
    'грунт-эмаль',
    'грунт-эмаль по ржавчине',
    'грунт-эмаль 3 в 1',
    'краска для металла',
    'краска по металлу',
    'эмаль по металлу',
    'быстросохнущая краска по металлу',
    'молотковая краска',
    'краска по ржавчине',
    'грунт по металлу',
    # Применение
    'краска для ворот',
    'краска для забора',
    'краска для гаража',
    'краска для металлической двери',
    'покрасить металлический забор',
    'чем покрасить металл',
    # Конкурентные
    'Dali краска по металлу',
    'Салван краска',
    'Husky краска',
    'Dufa краска по металлу',
]

def call_api(endpoint, body):
    """Вызывает Wordstat API метод."""
    url = f'{API_BASE}{endpoint}'
    body['folderId'] = FOLDER_ID
    try:
        resp = requests.post(url, headers=HEADERS, json=body, timeout=60)
        if resp.status_code == 200:
            return resp.json()
        else:
            print(f'  ERROR {resp.status_code}: {resp.text[:200]}')
            return None
    except Exception as e:
        print(f'  EXCEPTION: {e}')
        return None

# ---------- Тестовый запрос ----------
print('=== Тестовый запрос ===')
test = call_api('/v2/wordstat/topRequests', {
    'keyword': 'краска для металла',
    'regionIds': [225],  # Вся Россия
    'devices': 'all',
    'numPhrases': 10,
})
if test:
    print(f'Тест успешен! Получено данных: {len(json.dumps(test))} байт')
    # Покажем структуру
    print(f'Ключи ответа: {list(test.keys())}')
    if 'results' in test:
        print(f'Топ запросов ({len(test["results"])} шт.):')
        for i, item in enumerate(test['results'][:5]):
            print(f'  {i+1}. {item}')
    print()
else:
    print('Тест НЕ прошёл. Проверьте API-ключ и биллинг.')
    exit(1)

# ---------- Основной сбор данных ----------
all_data = {}

for idx, keyword in enumerate(KEYWORDS, 1):
    print(f'\n--- [{idx}/25] "{keyword}" ---')
    phrase_data = {'keyword': keyword}

    # 1. TopRequests
    print(f'  TopRequests...', end=' ')
    top = call_api('/v2/wordstat/topRequests', {
        'keyword': keyword,
        'regionIds': [225],
        'devices': 'all',
        'numPhrases': 100,
    })
    if top:
        results = top.get('results', [])
        associations = top.get('associations', [])
        total = top.get('totalCount', 0)
        phrase_data['top_results'] = results
        phrase_data['top_associations'] = associations
        phrase_data['total_count'] = total
        print(f'OK ({len(results)} results, {len(associations)} associations, total={total})')
    else:
        print('FAILED')
        phrase_data['top_results'] = []
        phrase_data['top_associations'] = []
        phrase_data['total_count'] = 0

    time.sleep(1)  # пауза между запросами

    # 2. RegionsDistribution
    print(f'  RegionsDistribution...', end=' ')
    regions = call_api('/v2/wordstat/regionsDistribution', {
        'keyword': keyword,
        'regionMode': 'all',
    })
    if regions:
        regions_list = regions.get('results', [])
        phrase_data['regions'] = regions_list
        print(f'OK ({len(regions_list)} регионов)')
    else:
        print('FAILED')
        phrase_data['regions'] = []

    time.sleep(1)

    # 3. Dynamics
    print(f'  Dynamics...', end=' ')
    dynamics = call_api('/v2/wordstat/dynamics', {
        'keyword': keyword,
        'period': 'monthly',
    })
    if dynamics:
        dynamics_list = dynamics.get('results', [])
        phrase_data['dynamics'] = dynamics_list
        print(f'OK ({len(dynamics_list)} точек)')
    else:
        print('FAILED')
        phrase_data['dynamics'] = []

    all_data[keyword] = phrase_data
    time.sleep(1)

# ---------- Сохранение ----------
out_path = '/home/z/my-project/download/wordstat_titan_raw.json'
with open(out_path, 'w', encoding='utf-8') as f:
    json.dump(all_data, f, ensure_ascii=False, indent=2)

print(f'\n=== Готово! ===')
print(f'Сохранено: {out_path}')
print(f'Размер: {os.path.getsize(out_path)} байт')
print(f'Фраз собрано: {len(all_data)}')

# Краткая сводка
print(f'\n--- Сводка ---')
for kw, data in all_data.items():
    total = data.get('total_count', 0)
    n_regions = len(data.get('regions', []))
    n_dyn = len(data.get('dynamics', []))
    print(f'  "{kw}": total={total}, regions={n_regions}, dynamics={n_dyn}')
