#!/usr/bin/env python3
"""
Сбор потенциальных дистрибьюторов ЛКМ по 3 странам: Киргизия, Монголия, Таджикистан.
Стратегия: web_search по 5-7 запросам на страну, извлечение компаний + контактов.
"""

import subprocess
import json
import re
import time
from collections import defaultdict


def search_web(query, num=15):
    """Поиск через z-ai CLI."""
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


def extract_companies(results, country):
    """Извлекает компании из результатов поиска."""
    companies = []
    for r in results:
        name = r.get('name', '')
        snippet = r.get('snippet', '')
        url = r.get('url', '')
        host = r.get('host_name', '')
        
        # Пропускаем информационные сайты
        skip_hosts = ['wikipedia', 'youtube.com', 'facebook.com', 'instagram.com', 
                      'yandex.ru', '2gis', 'yelp.com']
        if any(s in host.lower() for s in skip_hosts):
            continue
        
        # Извлекаем телефоны (международные форматы)
        phones = []
        phone_patterns = [
            r'\+996[\s\-]?\d{3}[\s\-]?\d{2}[\s\-]?\d{2}',  # Киргизия
            r'\+976[\s\-]?\d{4,5}[\s\-]?\d{3,4}',  # Монголия
            r'\+992[\s\-]?\d{2}[\s\-]?\d{3}[\s\-]?\d{2,3}',  # Таджикистан
            r'\+\d{1,3}[\s\-]?\(?\d{2,3}\)?[\s\-]?\d{3}[\s\-]?\d{2}[\s\-]?\d{2}',
        ]
        for pat in phone_patterns:
            phones.extend(re.findall(pat, snippet + ' ' + name))
        
        # Извлекаем email
        emails = re.findall(r'[\w.+-]+@[\w-]+\.[\w.-]+', snippet + ' ' + name)
        
        # Извлекаем сайт
        site = ''
        if url and url.startswith('http'):
            site = url.split('/')[2] if '//' in url else url
        
        # Очищаем имя (часто содержит "...— компания" и т.п.)
        clean_name = re.sub(r'\s*[—–|·].*$', '', name).strip()
        # Убираем "Главная" и т.п.
        clean_name = re.sub(r'^(Главная|Официальный сайт|Компания)\s*', '', clean_name)
        
        if clean_name and len(clean_name) > 3:
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


# === Поиск по 3 странам ===
queries_by_country = {
    'Киргизия': [
        'дистрибьютор ЛКМ краска Бишкек оптом',
        'лакокрасочные материалы оптовая компания Кыргызстан',
        'поставщик краски строительные Бишкек импортер',
        'краска для стен оптом Киргизия дилер',
        'ТД краска Бишкек Ош оптовая торговля',
        'строительные материалы краска Бишкек импорт',
        'алкидная эмаль грунт оптом Кыргызстан',
    ],
    'Монголия': [
        'буу будаг дистрибьютор Улаанбаатар',
        'paint distributor Ulaanbaatar Mongolia',
        'лакокрасочные материалы Монголия Улан-Батор',
        'paint wholesale Mongolia Ulaanbaatar importer',
        'строительные материалы краска Монголия',
        'будаг цавгайтай материал Улаанбаатар',
        'Mongolia paint coatings distributor import',
    ],
    'Таджикистан': [
        'дистрибьютор ЛКМ Душанбе краска оптом',
        'лакокрасочные материалы Таджикистан импортер',
        'поставщик краски строительные Душанбе',
        'краска для стен оптом Таджикистан дилер',
        'алкидная эмаль грунт Душанбе Худжанд',
        'строительные материалы краска Таджикистан импорт',
        'Tajikistan paint coatings distributor wholesale',
    ],
}

all_results = {}

for country, queries in queries_by_country.items():
    print(f"\n{'='*80}")
    print(f"=== {country} ===")
    print('='*80)
    
    country_companies = []
    seen_names = set()
    
    for qi, q in enumerate(queries, 1):
        print(f"\n  Query {qi}/{len(queries)}: {q[:70]}")
        results = search_web(q, num=12)
        print(f"    Найдено результатов: {len(results)}")
        
        companies = extract_companies(results, country)
        for c in companies:
            # Дедупликация по имени (нечёткая)
            name_lower = c['name'].lower()[:30]
            if name_lower not in seen_names:
                seen_names.add(name_lower)
                country_companies.append(c)
        
        time.sleep(1)
    
    all_results[country] = country_companies
    print(f"\n  Всего уникальных компаний по {country}: {len(country_companies)}")

# Сохранение
out_json = '/home/z/my-project/scripts/export_companies.json'
with open(out_json, 'w', encoding='utf-8') as f:
    json.dump(all_results, f, ensure_ascii=False, indent=2)

print(f"\n\n{'='*80}")
print(f"ИТОГИ СБОРА")
print('='*80)
total = 0
for country, companies in all_results.items():
    print(f"  {country}: {len(companies)} компаний")
    total += len(companies)
print(f"  ВСЕГО: {total}")
print(f"\nРезультаты сохранены: {out_json}")

# Вывод по 5 компаний на страну для контроля
for country, companies in all_results.items():
    print(f"\n\n=== {country} — первые 8 компаний ===")
    for i, c in enumerate(companies[:8], 1):
        print(f"\n  {i}. {c['name']}")
        print(f"     Сниппет: {c['snippet'][:200]}")
        if c['phones']:
            print(f"     Телефоны: {c['phones']}")
        if c['emails']:
            print(f"     Email: {c['emails']}")
        if c['site']:
            print(f"     Сайт: {c['site']}")
