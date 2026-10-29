#!/usr/bin/env python3
"""
Efficient LCM assortment research for remaining clients (groups 3,4,5).
Uses web search + page reader for active clients with websites.
Fills departed/inactive clients from comments data.
"""
import json
import subprocess
import re
import time
import sys
import os

REMAINING_FILE = '/home/z/my-project/scripts/remaining_clients.json'
OUTPUT_FILE = '/home/z/my-project/scripts/results_remaining.json'

# Known LCM brands
LCM_BRANDS = [
    'Лакра', 'Рогнеда', 'Декарт', 'Престиж', 'Элкон', 'Ижсинтез',
    'Ярославские краски', 'Радуга Маляр', 'Радуга Малер', 'Царицыно', 'Царицинские краски',
    'НБХ', 'Проремонт', 'Профилюкс', 'Новоколор', 'Простокрашено', 'Профивуд',
    'Ореол', 'Крафор', 'Славянка', 'Оптимист', 'Светлица', 'Грида', 'Хват',
    'ВГТ', 'VGT', 'Палиж', 'Palizh', 'Неомид', 'Neomid', 'Расцвет',
    'Алия', 'Маршал', 'Marshall', 'Эмпилс', 'Empils', 'Титан', 'Titan',
    'Текс', 'Tex', 'АКЦ', 'Альтакс', 'Altax', 'Колорит', 'Баупроф',
    'Строймарка', 'Кубанские Краски',
    'АВС Farben', 'АВС Фарбен', 'ABC Farben',
    'Тиккурила', 'Tikkurila', 'Dulux', 'AkzoNobel', 'АкзоНобель', 'Акзо', 'Akzo',
    'Alpina', 'Альпина', 'Caparol', 'Капарол', 'Dufa', 'Дюфа',
    'FLUGGER', 'Flugger', 'Флюггер', 'Pinotex', 'Пинотекс',
    'Sadolin', 'Садолин', 'Beckers', 'Беккерс', 'Hammerite', 'Хаммерайт',
    'Jotun', 'Йотун', 'Hempel', 'Хемпел', 'Nippon', 'Kansai',
    'Eskaro', 'Эскаро', 'Sniezka', 'Снежка',
    'Magnat', 'Магнат', 'Oswal', 'Berger', 'Sirca', 'Indigo',
    'Parade', 'Парад', 'Alpa', 'Teknos', 'Teknos', 'Remmers', 'Biofa',
    'Rogner', 'Meffert', 'Feidal', 'Brillux', 'Baufa',
    'Krafor', 'Farbitex', 'Фарбитекс',
]

MAJOR_BRANDS = [
    'Лакра', 'Рогнеда', 'Декарт', 'Текс', 'Новоколор', 'Профилюкс',
    'АВС Farben', 'АВС Фарбен', 'ABC Farben', 'Престиж', 'Элкон',
    'Тиккурила', 'Tikkurila', 'Dulux', 'Caparol', 'Alpina',
    'Царицыно', 'Царицинские краски', 'Ярославские краски',
    'ВГТ', 'VGT', 'Flugger', 'FLUGGER', 'Sadolin',
    'Marshall', 'Sniezka', 'Magnat', 'Эмпилс', 'Титан',
    'Радуга Маляр', 'Радуга Малер', 'Неомид', 'Neomid',
    'Палиж', 'Palizh', 'Крафор', 'Ореол', 'Ижсинтез',
    'Оптимист', 'Славянка', 'Parade', 'Pinotex', 'Hammerite',
    'Dufa', 'Jotun', 'Eskaro', 'Кубанские Краски',
    'Teknos', 'Remmers', 'Meffert', 'Alpa', 'Brillux',
    'Beckers', 'Hempel',
]

def search_web(query, num=5):
    """Search web using z-ai CLI"""
    try:
        cmd = ['z-ai', 'function', '-n', 'web_search', '-a', json.dumps({"query": query, "num": num})]
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=30)
        if result.returncode != 0:
            return []
        data = json.loads(result.stdout)
        return data if isinstance(data, list) else []
    except Exception as e:
        print(f"  Search error: {e}", file=sys.stderr)
        return []

def read_page(url):
    """Read web page using z-ai CLI"""
    try:
        out_file = f'/tmp/page_r_{int(time.time()*1000)}.json'
        cmd = ['z-ai', 'function', '-n', 'page_reader', '-a', json.dumps({"url": url}), '-o', out_file]
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=30)
        if result.returncode != 0:
            return ""
        with open(out_file, 'r') as f:
            data = json.load(f)
        if isinstance(data, dict) and 'data' in data:
            html = data['data'].get('html', '')
        elif isinstance(data, dict):
            html = data.get('html', '')
        else:
            html = ''
        text = re.sub(r'<[^>]+>', ' ', html)
        text = re.sub(r'\s+', ' ', text).strip()
        try:
            os.unlink(out_file)
        except:
            pass
        return text[:8000]
    except Exception as e:
        print(f"  Page read error: {e}", file=sys.stderr)
        return ""

def extract_brands(text):
    """Extract LCM brand names from text"""
    found = []
    text_lower = text.lower()
    for brand in LCM_BRANDS:
        if brand.lower() in text_lower:
            if not any(b.lower() == brand.lower() for b in found):
                found.append(brand)
    return found

def categorize(brands):
    """Split brands into main and additional"""
    if not brands:
        return '', ''
    main = []
    additional = []
    for brand in brands:
        if brand in ['Баупроф', 'Строймарка']:
            additional.append(brand + ' (СТМ)')
        elif brand in MAJOR_BRANDS:
            main.append(brand)
        else:
            additional.append(brand)
    if not main and brands:
        main = brands[:3]
        additional = brands[3:]
    return ', '.join(main), ', '.join(additional)

def is_departed(comments):
    """Check if this is a departed/inactive client"""
    cl = comments.lower()
    if any(p in cl for p in ['ушедший клиент', 'прекратил деятельность', 'компания ликвидирована',
                              'закрылись', 'не занимается больше', 'переименовался в корвон']):
        return True
    return False

def main():
    with open(REMAINING_FILE, 'r', encoding='utf-8') as f:
        clients = json.load(f)
    
    results = []
    print(f"Processing {len(clients)} remaining clients...")
    
    for idx, client in enumerate(clients):
        name = client.get('name', '')
        inn = client.get('inn', '')
        site = client.get('site', '')
        comments = client.get('comments', '')
        
        print(f"\n[{idx+1}/{len(clients)}] {name}")
        
        # Check if departed
        if is_departed(comments):
            brands = extract_brands(comments)
            if brands:
                main, add = categorize(brands)
                source = 'comments'
            else:
                cl = comments.lower()
                if any(p in cl for p in ['прекратил деятельность', 'компания ликвидирована', 'закрылись',
                                          'не занимается больше продажей', 'не занимается больше оптовой']):
                    main = 'Неактивный клиент'
                else:
                    main = 'Нет данных (ушедший клиент)'
                add = ''
                source = 'comments'
            print(f"  DEPARTED -> main={main}")
            results.append({
                'inn': inn, 'name': name,
                'main_assortment': main,
                'additional_assortment': add,
                'source': source
            })
            continue
        
        # Extract brands from comments first
        brands_from_comments = extract_brands(comments)
        print(f"  From comments: {brands_from_comments}")
        
        # Try web search for active clients with sites
        brands_from_web = []
        
        # Clean site URL
        site_clean = ''
        if site and site not in ['-', 'NO SITE', '', None]:
            site_clean = site.strip()
            if '\n' in site_clean:
                site_clean = site_clean.split('\n')[0].strip()
            if ' ' in site_clean and not site_clean.startswith('http'):
                site_clean = site_clean.split(' ')[0].strip()
        
        # Web search
        search_query = f"{name} лакокрасочные материалы бренды краски ассортимент"
        print(f"  Searching: {search_query}")
        search_results = search_web(search_query, num=5)
        
        if search_results:
            all_text = ' '.join([f"{r.get('name', '')} {r.get('snippet', '')}" for r in search_results])
            brands_from_web = extract_brands(all_text)
            print(f"  From search: {brands_from_web}")
            
            # Try to read relevant pages
            for sr in search_results[:2]:
                url = sr.get('url', '')
                if url and '2gis' not in url and 'yandex' not in url and 'google' not in url:
                    print(f"  Reading: {url[:60]}...")
                    page_text = read_page(url)
                    if page_text:
                        page_brands = extract_brands(page_text)
                        if page_brands:
                            brands_from_web = list(set(brands_from_web + page_brands))
                            print(f"  From page: {page_brands}")
        
        # Try direct website
        if site_clean and len(brands_from_web) < 2:
            if not site_clean.startswith('http'):
                site_url = f"https://{site_clean}"
            else:
                site_url = site_clean
            print(f"  Reading direct: {site_url}")
            page_text = read_page(site_url)
            if page_text:
                page_brands = extract_brands(page_text)
                if page_brands:
                    brands_from_web = list(set(brands_from_web + page_brands))
                    print(f"  From website: {page_brands}")
        
        # Combine
        all_brands = list(set(brands_from_comments + brands_from_web))
        
        if all_brands:
            main, add = categorize(all_brands)
            source = 'comments+web' if brands_from_comments and brands_from_web else ('web_search' if brands_from_web else 'comments')
        else:
            main = 'Нет данных'
            add = ''
            source = 'none'
        
        print(f"  RESULT: main={main} | add={add}")
        results.append({
            'inn': inn, 'name': name,
            'main_assortment': main,
            'additional_assortment': add,
            'source': source
        })
        
        # Rate limiting
        time.sleep(0.3)
        
        # Save intermediate results every 15 clients
        if (idx + 1) % 15 == 0:
            with open(OUTPUT_FILE, 'w', encoding='utf-8') as f:
                json.dump(results, f, ensure_ascii=False, indent=2)
            print(f"\n  [Intermediate save at {idx+1} clients]\n")
    
    # Final save
    with open(OUTPUT_FILE, 'w', encoding='utf-8') as f:
        json.dump(results, f, ensure_ascii=False, indent=2)
    
    # Summary
    with_data = sum(1 for r in results if r['main_assortment'] not in ['Нет данных', 'Нет данных (ушедший клиент)', 'Неактивный клиент'])
    without = sum(1 for r in results if r['main_assortment'] in ['Нет данных', 'Нет данных (ушедший клиент)', 'Неактивный клиент'])
    print(f"\nDone! {with_data} with data, {without} without data, total {len(results)}")

if __name__ == '__main__':
    main()
