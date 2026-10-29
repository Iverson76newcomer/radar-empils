#!/usr/bin/env python3
"""
Сбор компаний по 3 странам (batch). Запускается с аргументом batch=1/2/3.
batch 1: Молдова, Румыния, Сербия
batch 2: Беларусь, Узбекистан, Грузия
batch 3: Армения, Казахстан, Туркменистан
"""

import sys
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
    except:
        return []


def extract_companies(results):
    companies = []
    for r in results:
        name = r.get('name', '')
        snippet = r.get('snippet', '')
        url = r.get('url', '')
        host = r.get('host_name', '')
        skip_hosts = ['wikipedia', 'youtube', 'facebook', 'instagram', 'yandex', '2gis',
                      'yelp', 'linkedin', 'twitter', 'tiktok', 'telegram']
        if any(s in host.lower() for s in skip_hosts):
            continue
        phones = []
        for pat in [r'\+\d{1,3}[\s\-]?\(?\d{2,3}\)?[\s\-]?\d{3}[\s\-]?\d{2}[\s\-]?\d{2}',
                    r'\+\d{1,3}[\s\-]?\d{2}[\s\-]?\d{3}[\s\-]?\d{3}',
                    r'\+\d{1,3}[\s\-]?\d{3}[\s\-]?\d{3}[\s\-]?\d{2,3}']:
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


EXISTING = {
    'Молдова': ['Cheton', 'Unixone', 'Color Smart', 'Supraten', 'SuperStefan', 'Verix'],
    'Румыния': ['Unixone', 'Color Smart', 'Supraten', 'SuperStefan', 'Verix'],
    'Сербия': ['Roma', 'Lisa', 'Hemokolor', 'Beorol', 'Irkom', 'JUB', 'LTV', 'NI-BO'],
    'Беларусь': ['ПРОМКРАС', 'Proma', 'Диском', 'Лакокраска', 'Лида', 'Ланквитцер', 'Бел-Киндекор',
                 'Тексколор', 'Ватра', 'БелПека', 'Мастак', 'Artstroy'],
    'Узбекистан': ['OSCAR', 'STM COLOR', 'WEBER', 'PREMIUM', 'Spektr', 'RUMATIN', 'ELEMENT', 'UNIQ'],
    'Грузия': ['GTS', 'Art House', 'PPG', 'Betek', 'Mega Colors', 'Eliava', 'GrandMall'],
    'Армения': ['Краун', 'ЛИА-К', 'Сард', 'ШЕН', 'Нарек', 'Бенджамин', 'Колор Стайл', 'Элит',
                'Лакме', 'АРТФЕМ', 'Хачатрян', 'Граф колорс'],
    'Казахстан': ['Alina', 'Алматинские', 'Белый Дом', 'Kaizer', 'Remalux', 'ВМП', 'BAURAMIX',
                  'Радуга', 'SMASH', 'GAUDI', 'Walls Decor', 'Казхимсеть', 'SVYDIS', 'Гермес',
                  'Семь Защит', 'Фирма Партнер', 'БелГраА', 'Melcor', '5-БАЛЛОВ', 'Крафт',
                  'Идеал', 'Alem', 'LuxAn', 'Kraska', 'Color Studio'],
    'Туркменистан': ['Ak Gaýa', 'Galkynan', 'Ynamly', 'Berk', 'Daýanç', 'Merjen', 'TÖWEREK',
                     'ГЕНЧ', 'GENÇ', 'Erco', 'Jotun', 'Kamil', 'BYK', 'Elwan', 'Zok Decor'],
}


def is_existing(name, country):
    name_lower = name.lower()
    for ex in EXISTING.get(country, []):
        if ex.lower() in name_lower:
            return True
    return False


QUERIES = {
    'Молдова': [
        'дистрибьютор ЛКМ Кишинев краска оптом',
        'лакокрасочные материалы Молдова поставщик импортер',
        'краска для стен оптом Кишинев дилер',
        'paint distributor Moldova Chisinau wholesale',
    ],
    'Румыния': [
        'paint distributor Romania Bucharest wholesale',
        'vopsea distribuitor Romania Bucuresti',
        'лакокрасочные материалы Румыния Бухарест импортер',
        'paint coatings wholesale Romania importer',
    ],
    'Сербия': [
        'paint distributor Serbia Belgrade wholesale',
        'boja distributer Srbija Beograd',
        'лакокрасочные материалы Сербия Белград поставщик',
        'paint coatings Serbia importer wholesale',
    ],
    'Беларусь': [
        'дистрибьютор ЛКМ Минск краска оптом',
        'лакокрасочные материалы Беларусь поставщик импортер',
        'краска для стен оптом Минск дилер',
        'paint distributor Belarus Minsk wholesale',
    ],
    'Узбекистан': [
        'дистрибьютор ЛКМ Ташкент краска оптом',
        'лакокрасочные материалы Узбекистан поставщик импортер',
        'краска для стен оптом Ташкент дилер',
        'paint distributor Uzbekistan Tashkent wholesale',
    ],
    'Грузия': [
        'дистрибьютор ЛКМ Тбилиси краска оптом',
        'лакокрасочные материалы Грузия поставщик импортер',
        'краска для стен оптом Тбилиси дилер',
        'paint distributor Georgia Tbilisi wholesale',
    ],
    'Армения': [
        'дистрибьютор ЛКМ Ереван краска оптом',
        'лакокрасочные материалы Армения поставщик импортер',
        'краска для стен оптом Ереван дилер',
        'paint distributor Armenia Yerevan wholesale',
    ],
    'Казахстан': [
        'дистрибьютор ЛКМ Алматы краска оптом',
        'лакокрасочные материалы Казахстан поставщик импортер',
        'краска для стен оптом Алматы дилер',
        'paint distributor Kazakhstan Almaty wholesale',
    ],
    'Туркменистан': [
        'дистрибьютор ЛКМ Ашхабад краска оптом',
        'лакокрасочные материалы Туркменистан поставщик',
        'краска для стен оптом Ашхабад дилер',
        'paint distributor Turkmenistan Ashgabat wholesale',
    ],
}

BATCHES = {
    1: ['Молдова', 'Румыния', 'Сербия'],
    2: ['Беларусь', 'Узбекистан', 'Грузия'],
    3: ['Армения', 'Казахстан', 'Туркменистан'],
}

batch_num = int(sys.argv[1]) if len(sys.argv) > 1 else 1
countries = BATCHES[batch_num]

print(f"=== BATCH {batch_num}: {countries} ===\n", flush=True)

results_all = {}
for country in countries:
    print(f"\n=== {country} ===", flush=True)
    companies = []
    seen_names = set()
    for qi, q in enumerate(QUERIES[country], 1):
        print(f"  Q{qi}: {q[:60]}", flush=True)
        results = search_web(q, num=10)
        for c in extract_companies(results):
            name_key = c['name'].lower()[:25]
            if name_key in seen_names:
                continue
            if is_existing(c['name'], country):
                continue
            seen_names.add(name_key)
            companies.append(c)
        time.sleep(5)  # пауза 5 сек, чтобы не попасть под лимит 429
    results_all[country] = companies
    print(f"  → {len(companies)} новых компаний", flush=True)

out_json = f'/home/z/my-project/scripts/batch_{batch_num}.json'
with open(out_json, 'w', encoding='utf-8') as f:
    json.dump(results_all, f, ensure_ascii=False, indent=2)

print(f"\nOK: {out_json}", flush=True)
for c, companies in results_all.items():
    print(f"  {c}: {len(companies)}", flush=True)
