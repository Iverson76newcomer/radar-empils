#!/usr/bin/env python3
"""
Сбор НОВЫХ потенциальных дистрибьюторов ЛКМ по Армении.
Только Армения. Пауза 3 сек между запросами, чтобы не ловить 429.
"""

import subprocess
import json
import re
import time
import os


def search_web(query, num=10):
    try:
        cmd = ['z-ai', 'function', '-n', 'web_search', '-a', json.dumps({'query': query, 'num': num})]
        r = subprocess.run(cmd, capture_output=True, text=True, timeout=60)
        if r.returncode != 0:
            return []
        output = r.stdout
        idx = output.find('[')
        end_idx = output.rfind(']')
        if idx < 0 or end_idx < idx:
            return []
        json_str = output[idx:end_idx+1]
        try:
            return json.loads(json_str)
        except:
            return []
    except Exception as e:
        print(f'  Ошибка: {e}')
        return []


def extract_companies(results):
    companies = []
    for r in results:
        name = r.get('name', '')
        snippet = r.get('snippet', '')
        url = r.get('url', '')
        host = r.get('host_name', '')
        skip_hosts = ['wikipedia', 'youtube', 'facebook', 'instagram', 'yandex', '2gis',
                      'yelp', 'linkedin', 'twitter', 'tiktok', 'telegram', 'tiktok']
        if any(s in host.lower() for s in skip_hosts):
            continue
        phones = []
        for pat in [r'\+\d{1,3}[\s\-]?\(?\d{2,3}\)?[\s\-]?\d{3}[\s\-]?\d{2}[\s\-]?\d{2}',
                    r'\+\d{1,3}[\s\-]?\d{2}[\s\-]?\d{3}[\s\-]?\d{3}',
                    r'\+\d{1,3}[\s\-]?\d{3}[\s\-]?\d{3}[\s\-]?\d{2,3}',
                    r'\+374[\s\-]?\d{2}[\s\-]?\d{2}[\s\-]?\d{2}']:
            phones.extend(re.findall(pat, snippet + ' ' + name))
        emails = re.findall(r'[\w.+-]+@[\w-]+\.[\w.-]+', snippet + ' ' + name)
        site = ''
        if url and url.startswith('http'):
            site = url.split('/')[2] if '//' in url else url
            site = re.sub(r'^www\.', '', site)
        clean_name = re.sub(r'\s*[—–|·].*$', '', name).strip()
        clean_name = re.sub(r'^(Главная|Официальный сайт|Компания|Каталог)\s*', '', clean_name)
        if clean_name and len(clean_name) > 5:
            companies.append({
                'name': clean_name[:120],
                'snippet': snippet[:400],
                'phones': list(set(phones))[:3],
                'emails': list(set(emails))[:2],
                'site': site,
                'url': url,
                'host': host,
            })
    return companies


# Существующие компании Армении — для дедупликации
EXISTING_ARMENIA = [
    'Краун', 'ЛИА-К', 'Lia Ceramics', 'Сард', 'ШЕН-КОНЦЕРН', 'Нарек',
    'Бенджамин Мур', 'Колор Стайл', 'Элит Шинанют', 'Лакме', 'АРТФЕМ',
    'Хачатрян', 'Граф колорс', 'ШЕН ТАНИК',
]


def is_existing(name):
    name_lower = name.lower()
    for ex in EXISTING_ARMENIA:
        if ex.lower() in name_lower:
            return True
    return False


# 6 запросов по Армении
QUERIES = [
    'дистрибьютор ЛКМ Ереван краска оптом',
    'лакокрасочные материалы Армения поставщик импортер',
    'краска для стен оптом Ереван дилер',
    'paint distributor Armenia Yerevan wholesale',
    'строительные материалы краска Армения опт',
    'алкидная эмаль грунт Ереван поставщик',
]

print("="*70)
print("СБОР ПО АРМЕНИИ")
print("="*70)

companies = []
seen_names = set()

for qi, q in enumerate(QUERIES, 1):
    print(f"\n  Q{qi}/{len(QUERIES)}: {q}", flush=True)
    results = search_web(q, num=10)
    print(f"    Результатов: {len(results)}", flush=True)
    
    new_in_this_query = 0
    for c in extract_companies(results):
        name_key = c['name'].lower()[:25]
        if name_key in seen_names:
            continue
        if is_existing(c['name']):
            continue
        seen_names.add(name_key)
        companies.append(c)
        new_in_this_query += 1
    
    print(f"    Новых: {new_in_this_query}", flush=True)
    time.sleep(3)  # пауза, чтобы не ловить 429

# Сохранение
out_json = '/home/z/my-project/scripts/armenia_companies.json'
with open(out_json, 'w', encoding='utf-8') as f:
    json.dump(companies, f, ensure_ascii=False, indent=2)

print(f"\n{'='*70}")
print(f"ИТОГ: найдено {len(companies)} новых компаний по Армении")
print(f"Сохранено: {out_json}")
print('='*70)

# Покажем все найденные
print("\n=== Все найденные новые компании ===")
for i, c in enumerate(companies, 1):
    print(f"\n{i}. {c['name']}")
    print(f"   Сниппет: {c['snippet'][:200]}")
    if c['phones']:
        print(f"   Тел: {c['phones']}")
    if c['emails']:
        print(f"   Email: {c['emails']}")
    if c['site']:
        print(f"   Сайт: {c['site']}")
