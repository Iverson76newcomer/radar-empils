#!/usr/bin/env python3
"""
Волна 2: Чтение Rusprofile для чистых данных.
Исправление городов, адресов, сайтов, email.
"""

import openpyxl
import json
import re
import os
import subprocess
import time

SEARCH_DIR = '/home/z/my-project/scripts/search_results'
PAGES_DIR = '/home/z/my-project/scripts/rusprofile_pages'
os.makedirs(PAGES_DIR, exist_ok=True)

# Все INN для обработки
ALL_INNS = [
    ('7811340145', 3), ('2537090274', 4), ('5902816678', 5), ('246600161805', 6),
    ('2465056419', 7), ('5406296559', 8), ('5003123271', 9), ('7806444830', 10),
    ('5752045215', 11), ('9717052425', 12), ('4401025465', 13), ('622700326226', 14),
    ('3827026836', 15), ('7827001116', 16), ('2465254410', 17), ('2511036554', 18),
    ('7816380127', 19), ('550604212050', 20), ('245700918501', 21), ('760400191237', 22),
    ('550400093929', 23), ('246200497430', 24), ('246500188627', 25), ('550405862624', 26),
    ('6670179530', 27), ('6673163835', 28), ('7731564157', 29), ('5506230390', 30),
    ('7203520720', 31), ('7716989966', 32), ('7453271996', 33), ('6625044348', 34),
    ('7452095917', 35), ('6678102667', 36), ('2464236507', 37), ('5906062502', 38),
    ('5904341998', 39), ('1101075340', 40), ('3731030101', 41), ('7017379679', 42),
    ('6623119164', 43), ('2457079148', 44), ('7733251960', 45), ('2464002594', 46),
    ('2462068624', 47), ('3663157837', 48), ('3662061709', 49), ('2904028638', 50),
    ('3442117755', 51), ('4632110238', 52), ('7714996975', 53), ('665200210601', 54),
    ('7728115183', 55), ('9102282914', 56), ('2309181028', 57),
]

def get_rusprofile_url(inn):
    """Получить URL Rusprofile из результатов поиска"""
    for fname in [f'{inn}.json', f'inn_{inn}.json']:
        fpath = os.path.join(SEARCH_DIR, fname)
        if os.path.exists(fpath):
            try:
                with open(fpath, 'r') as f:
                    data = json.load(f)
                if isinstance(data, list):
                    for r in data:
                        url = r.get('url', '')
                        if 'rusprofile.ru' in url:
                            return url
            except:
                pass
    return None

# Читаем Rusprofile для компаний, у которых нет чистых данных
# Приоритет — читаем по 5 компаний за раз, чтобы не упереться в rate limit

print("Поиск Rusprofile URL и чтение страниц...")

for i, (inn, row) in enumerate(ALL_INNS):
    out_file = os.path.join(PAGES_DIR, f'{inn}.json')
    
    if os.path.exists(out_file) and os.path.getsize(out_file) > 100:
        print(f"[{i+1}] {inn} — уже прочитано")
        continue
    
    url = get_rusprofile_url(inn)
    if not url:
        print(f"[{i+1}] {inn} — Rusprofile URL не найден")
        continue
    
    print(f"[{i+1}] {inn} — чтение {url}")
    try:
        result = subprocess.run(
            ['z-ai', 'function', '-n', 'page_reader',
             '-a', json.dumps({"url": url}),
             '-o', out_file],
            capture_output=True, text=True, timeout=45
        )
        if os.path.exists(out_file):
            print(f"  ✅ Сохранено")
        else:
            print(f"  ❌ Файл не создан")
    except subprocess.TimeoutExpired:
        print(f"  ⏰ Timeout")
    except Exception as e:
        print(f"  ❌ Error: {e}")
    
    time.sleep(2)  # Пауза для rate limiting

print("\nГотово!")
