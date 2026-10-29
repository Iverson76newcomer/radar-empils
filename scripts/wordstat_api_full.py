"""
Yandex Wordstat API — полный сбор данных по 25 фразам для TITAN.
3 метода: topRequests, regions, dynamics.
"""
import requests
import json
import time
import os
import calendar
from datetime import datetime

API_KEY = 'AQVNyBi4wPgqGtaKV9-dvavzCH5TSaUNwBPvAErO'
FOLDER_ID = 'b1g0bo7majef8lq3iii2'
API_BASE = 'https://searchapi.api.cloud.yandex.net'
HEADERS = {
    'Authorization': f'Api-Key {API_KEY}',
    'Content-Type': 'application/json',
    'Accept-Language': 'ru',
}

KEYWORDS = [
    'TITAN краска', 'Ореол TITAN', 'Эмпилс', 'краска TITAN по металлу', 'грунт-эмаль TITAN',
    'грунт-эмаль', 'грунт-эмаль по ржавчине', 'грунт-эмаль 3 в 1', 'краска для металла',
    'краска по металлу', 'эмаль по металлу', 'быстросохнущая краска по металлу',
    'молотковая краска', 'краска по ржавчине', 'грунт по металлу',
    'краска для ворот', 'краска для забора', 'краска для гаража',
    'краска для металлической двери', 'покрасить металлический забор', 'чем покрасить металл',
    'Dali краска по металлу', 'Салван краска', 'Husky краска', 'Dufa краска по металлу',
]

def call_api(endpoint, body):
    body['folderId'] = FOLDER_ID
    try:
        resp = requests.post(f'{API_BASE}{endpoint}', headers=HEADERS, json=body, timeout=60)
        if resp.status_code == 200:
            return resp.json()
        else:
            print(f'    ERROR {resp.status_code}: {resp.text[:150]}')
            return None
    except Exception as e:
        print(f'    EXCEPTION: {e}')
        return None

# Даты для динамики
from_date = '2025-09-01T00:00:00Z'
last_day = calendar.monthrange(2026, 8)[1]
to_date = f'2026-08-{last_day}T23:59:59Z'

all_data = {}
total_requests = 0

for idx, keyword in enumerate(KEYWORDS, 1):
    print(f'[{idx}/25] "{keyword}"')
    phrase_data = {'keyword': keyword}

    # 1. TopRequests
    top = call_api('/v2/wordstat/topRequests', {
        'phrase': keyword, 'numPhrases': 100,
    })
    if top:
        phrase_data['top_results'] = top.get('results', [])
        phrase_data['top_associations'] = top.get('associations', [])
        phrase_data['total_count'] = top.get('totalCount', '0')
        total_count = top.get('totalCount', '0')
        print(f'  Top: {len(top.get("results", []))} results, total={total_count}')
    else:
        phrase_data['top_results'] = []
        phrase_data['top_associations'] = []
        phrase_data['total_count'] = '0'
    total_requests += 1
    time.sleep(0.5)

    # 2. Regions
    regions = call_api('/v2/wordstat/regions', {
        'phrase': keyword,
    })
    if regions:
        phrase_data['regions'] = regions.get('results', [])
        print(f'  Regions: {len(regions.get("results", []))} items')
    else:
        phrase_data['regions'] = []
    total_requests += 1
    time.sleep(0.5)

    # 3. Dynamics
    dynamics = call_api('/v2/wordstat/dynamics', {
        'phrase': keyword, 'period': 'PERIOD_MONTHLY',
        'fromDate': from_date, 'toDate': to_date,
    })
    if dynamics:
        phrase_data['dynamics'] = dynamics.get('results', [])
        print(f'  Dynamics: {len(dynamics.get("results", []))} points')
    else:
        phrase_data['dynamics'] = []
    total_requests += 1
    time.sleep(0.5)

    all_data[keyword] = phrase_data

# Сохранение
out_path = '/home/z/my-project/download/wordstat_titan_raw.json'
with open(out_path, 'w', encoding='utf-8') as f:
    json.dump(all_data, f, ensure_ascii=False, indent=2)

print(f'\n=== Готово! ===')
print(f'Запросов к API: {total_requests}')
print(f'Сохранено: {out_path}')
print(f'Размер: {os.path.getsize(out_path):,} байт')

# Сводка
print(f'\n--- Сводка ---')
for kw, data in all_data.items():
    total = data.get('total_count', '0')
    n_reg = len(data.get('regions', []))
    n_dyn = len(data.get('dynamics', []))
    print(f'  "{kw}": total={total}, regions={n_reg}, dynamics={n_dyn}')
