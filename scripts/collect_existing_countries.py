#!/usr/bin/env python3
"""
Сбор потенциальных дистрибьюторов ЛКМ по 9 существующим странам.
Цель: дополнить базу, найти НОВЫХ потенциальных партнёров.
Страны: Молдова, Румыния, Сербия, Беларусь, Узбекистан, Грузия, Армения, Казахстан, Туркменистан.
"""

import subprocess
import json
import re
import time
from collections import defaultdict


def search_web(query, num=15):
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
        
        skip_hosts = ['wikipedia', 'youtube.com', 'facebook.com', 'instagram.com', 
                      'yandex.ru', '2gis', 'yelp.com', 'linkedin', 'twitter.com',
                      'tiktok', 'telegram']
        if any(s in host.lower() for s in skip_hosts):
            continue
        
        # Телефоны — универсальный паттерн
        phones = []
        phone_patterns = [
            r'\+\d{1,3}[\s\-]?\(?\d{2,3}\)?[\s\-]?\d{3}[\s\-]?\d{2}[\s\-]?\d{2}',
            r'\+\d{1,3}[\s\-]?\d{2}[\s\-]?\d{3}[\s\-]?\d{3}',
            r'\+\d{1,3}[\s\-]?\d{3}[\s\-]?\d{3}[\s\-]?\d{2,3}',
        ]
        for pat in phone_patterns:
            phones.extend(re.findall(pat, snippet + ' ' + name))
        
        emails = re.findall(r'[\w.+-]+@[\w-]+\.[\w.-]+', snippet + ' ' + name)
        
        site = ''
        if url and url.startswith('http'):
            site = url.split('/')[2] if '//' in url else url
            # Убираем www.
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


# Существующие компании (для дедупликации)
EXISTING_COMPANIES = {
    'Молдова': ['Cheton Grup', 'Unixone', 'Color Smart', 'Supraten', 'SuperStefan', 'Verix-Grup'],
    'Румыния': ['Unixone SRL', 'Color Smart', 'Supraten', 'SuperStefan', 'Verix-Grup'],
    'Сербия': ['Roma', 'Lisa-comerc', 'Hemokolor', 'Beorol', 'Irkom', 'JUB', 'LTV-PROFIL', 'NI-BO COLOR'],
    'Беларусь': ['ПРОМКРАС', 'Proma', 'Диском', 'Лакокраска', 'Лида', 'Ланквитцер', 'Бел-Киндекор',
                 'ТексколорБел', 'Ватра', 'БелПека', 'Мастак', 'Artstroyinvest'],
    'Узбекистан': ['OSCAR Paint', 'STM COLOR', 'WEBER', 'PREMIUM PAINTS', 'Spektr Trade', 'RUMATIN', 'ELEMENT TRADING', 'UNIQ'],
    'Грузия': ['GTS Group', 'Art House', 'PPG Caucasus', 'Betek Profi', 'Mega Colors', 'Eliava', 'GrandMall'],
    'Армения': ['Краун Кемикал', 'ЛИА-К', 'Lia Ceramics', 'Сард', 'ШЕН-КОНЦЕРН', 'Нарек', 'Бенджамин Мур',
                'Колор Стайл', 'Элит Шинанют', 'Лакме', 'АРТФЕМ', 'Генрих Хачатрян', 'Граф колорс', 'ШЕН ТАНИК'],
    'Казахстан': ['Alina Group', 'Алматинские краски', 'Белый Дом', 'Kaizer', 'Remalux', 'ВМП Астана',
                  'BAURAMIX', 'Радуга', 'SMASH', 'GAUDI', 'Walls Decor', 'Казхимсеть', 'SVYDIS', 'Гермес',
                  'Семь Защит', 'Фирма Партнер', 'БелГраА', 'Melcor', '5-БАЛЛОВ', 'Крафт', 'Идеал Трейд',
                  'Alem Color', 'LuxAn', 'Kraska.kz', 'Color Studio'],
    'Туркменистан': ['Ak Gaýa', 'Galkynan', 'Ynamly', 'Berk ýelim', 'Daýanç', 'Merjen', 'TÖWEREK',
                     'ГЕНЧ', 'GENÇ', 'Erco', 'Jotun', 'Kamil', 'BYK', 'Elwan', 'Zok Decor'],
}


def is_existing(name, country):
    """Проверяет, есть ли компания уже в базе."""
    name_lower = name.lower()
    for ex in EXISTING_COMPANIES.get(country, []):
        if ex.lower() in name_lower:
            return True
    return False


# Запросы по 9 странам
queries_by_country = {
    'Молдова': [
        'дистрибьютор ЛКМ Кишинев краска оптом',
        'лакокрасочные материалы Молдова поставщик импортер',
        'краска для стен оптом Кишинев дилер',
        'paint distributor Moldova Chisinau wholesale',
        'строительные материалы краска Молдова',
    ],
    'Румыния': [
        'paint distributor Romania Bucharest wholesale',
        'vopsea distribuitor Romania Bucuresti',
        'лакокрасочные материалы Румыния Бухарест импортер',
        'paint coatings wholesale Romania importer',
        'строительные краски Румыния дистрибьютор',
    ],
    'Сербия': [
        'paint distributor Serbia Belgrade wholesale',
        'boja distributer Srbija Beograd',
        'лакокрасочные материалы Сербия Белград поставщик',
        'paint coatings Serbia importer wholesale',
        'строительные краски Сербия опт',
    ],
    'Беларусь': [
        'дистрибьютор ЛКМ Минск краска оптом',
        'лакокрасочные материалы Беларусь поставщик импортер',
        'краска для стен оптом Минск дилер',
        'paint distributor Belarus Minsk wholesale',
        'алкидная эмаль грунт оптом Беларусь',
    ],
    'Узбекистан': [
        'дистрибьютор ЛКМ Ташкент краска оптом',
        'лакокрасочные материалы Узбекистан поставщик импортер',
        'краска для стен оптом Ташкент дилер',
        'paint distributor Uzbekistan Tashkent wholesale',
        'строительные материалы краска Узбекистан',
    ],
    'Грузия': [
        'дистрибьютор ЛКМ Тбилиси краска оптом',
        'лакокрасочные материалы Грузия поставщик импортер',
        'краска для стен оптом Тбилиси дилер',
        'paint distributor Georgia Tbilisi wholesale',
        'строительные материалы краска Грузия',
    ],
    'Армения': [
        'дистрибьютор ЛКМ Ереван краска оптом',
        'лакокрасочные материалы Армения поставщик импортер',
        'краска для стен оптом Ереван дилер',
        'paint distributor Armenia Yerevan wholesale',
        'строительные материалы краска Армения',
    ],
    'Казахстан': [
        'дистрибьютор ЛКМ Алматы Астана краска оптом',
        'лакокрасочные материалы Казахстан поставщик импортер',
        'краска для стен оптом Алматы дилер',
        'paint distributor Kazakhstan Almaty Astana wholesale',
        'алкидная эмаль грунт оптом Казахстан',
    ],
    'Туркменистан': [
        'дистрибьютор ЛКМ Ашхабад краска оптом',
        'лакокрасочные материалы Туркменистан поставщик импортер',
        'краска для стен оптом Ашхабад дилер',
        'paint distributor Turkmenistan Ashgabat wholesale',
        'строительные материалы краска Туркменистан',
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
        results = search_web(q, num=10)
        print(f"    Найдено результатов: {len(results)}")
        
        companies = extract_companies(results)
        for c in companies:
            # Дедупликация
            name_key = c['name'].lower()[:25]
            if name_key in seen_names:
                continue
            # Пропускаем существующие
            if is_existing(c['name'], country):
                continue
            seen_names.add(name_key)
            country_companies.append(c)
        
        time.sleep(1)
    
    all_results[country] = country_companies
    print(f"\n  Всего НОВЫХ компаний по {country}: {len(country_companies)}")

# Сохранение
out_json = '/home/z/my-project/scripts/export_companies_existing.json'
with open(out_json, 'w', encoding='utf-8') as f:
    json.dump(all_results, f, ensure_ascii=False, indent=2)

print(f"\n\n{'='*80}")
print(f"ИТОГИ СБОРА (ТОЛЬКО НОВЫЕ КОМПАНИИ)")
print('='*80)
total = 0
for country, companies in all_results.items():
    print(f"  {country}: {len(companies)} новых компаний")
    total += len(companies)
print(f"  ВСЕГО НОВЫХ: {total}")
print(f"\nРезультаты: {out_json}")
