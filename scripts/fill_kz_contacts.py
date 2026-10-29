#!/usr/bin/env python3
"""
Дополнение контактов (email, телефон) для 49 казахстанских дистрибьюторов.
Стратегия: web_search по каждой компании + сайт, извлечение email/телефона.
Пакетная обработка: 8 компаний за запуск, пауза 8 сек.
Прогресс сохраняется после каждой компании.
"""

import subprocess
import json
import re
import time
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from copy import copy
import os

SRC = '/home/z/my-project/upload/Kazachstan_Eksport_potentsialnye_distribyutory_polnaya.xlsx'
OUT = '/home/z/my-project/download/Kazachstan_контакты_дополнено.xlsx'
PROGRESS_FILE = '/home/z/my-project/scripts/kz_contacts_progress.json'


def search_web(query, num=6):
    try:
        cmd = ['z-ai', 'function', '-n', 'web_search', '-a', json.dumps({'query': query, 'num': num})]
        r = subprocess.run(cmd, capture_output=True, text=True, timeout=30)
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


def extract_contacts(results, company_name):
    """Извлекает email и телефон из результатов поиска."""
    info = {'email': '', 'phone': '', 'site': ''}
    
    # Плохие сайты — доски объявлений
    bad_hosts = ['wikipedia', 'yelp', 'linkedin', 'facebook', 'instagram', 'youtube',
                 'tiktok', 'telegram', 'pulscen', 'satu.kz', 'olx', 'flagma', 'all.biz',
                 'bizorg', 'optlist', 'yellowpages', 'kompass', 'infobel', 'eponuda',
                 'rentechdigital', 'optoviki', 'centralasiacoatingshow']
    
    # Собираем текст из всех результатов
    all_text = ''
    for r in results:
        host = r.get('host_name', '').lower()
        url = r.get('url', '').lower()
        if any(bad in host or bad in url for bad in bad_hosts):
            continue
        snippet = r.get('snippet', '')
        name = r.get('name', '')
        all_text += ' ' + name + ' ' + snippet
        # Запоминаем первый хороший сайт
        if not info['site'] and host:
            site = host.replace('www.', '')
            if not any(bad in site for bad in bad_hosts):
                info['site'] = site
    
    # Email — ищем все
    emails = re.findall(r'[\w.+-]+@[\w-]+\.[\w.-]+', all_text)
    clean_emails = [e for e in emails if not e.endswith('.') 
                   and 'sentry' not in e and 'wixpress' not in e
                   and 'example.com' not in e and 'noreply' not in e
                   and 'gmail' not in e.lower() or 'info@' in e.lower() or 'sales@' in e.lower()]
    # Предпочитаем корпоративные email
    corp_emails = [e for e in clean_emails if 'gmail' not in e.lower() and 'mail.ru' not in e.lower()]
    if corp_emails:
        info['email'] = corp_emails[0]
    elif clean_emails:
        info['email'] = clean_emails[0]
    
    # Телефоны — Казахстан (+7) и международные
    phones = []
    phone_patterns = [
        r'\+7[\s\-]?\(?\d{3}\)?[\s\-]?\d{3}[\s\-]?\d{2}[\s\-]?\d{2}',
        r'\+7[\s\-]?\d{3}[\s\-]?\d{3}[\s\-]?\d{2}[\s\-]?\d{2}',
        r'\+7[\s\-]?\d{2}[\s\-]?\d{3}[\s\-]?\d{3}',
        r'8[\s\-]?\(?\d{3}\)?[\s\-]?\d{3}[\s\-]?\d{2}[\s\-]?\d{2}',
    ]
    for pat in phone_patterns:
        phones.extend(re.findall(pat, all_text))
    if phones:
        info['phone'] = phones[0]
    
    return info


# === ЗАГРУЗКА EXCEL ===
wb = openpyxl.load_workbook(SRC)
ws = wb['Лист1']

# Загружаем прогресс
progress = {}
if os.path.exists(PROGRESS_FILE):
    with open(PROGRESS_FILE, 'r', encoding='utf-8') as f:
        progress = json.load(f)
    print(f"Прогресс: {len(progress)} компаний уже обработано")

# Список компаний (строки 4-52)
companies = []
for ri in range(4, 53):
    name = ws.cell(ri, 3).value
    if name:
        site = ws.cell(ri, 10).value or ''
        companies.append((ri, str(name).strip(), site))

print(f"Всего компаний: {len(companies)}")
print(f"Осталось обработать: {len([c for c in companies if str(c[0]) not in progress])}")

MAX_PER_RUN = 8
processed = 0

for ri, name, site in companies:
    if str(ri) in progress:
        continue
    if processed >= MAX_PER_RUN:
        print(f"\nДостигнут лимит {MAX_PER_RUN} компаний. Сохраняю и завершаюсь.")
        break
    
    # Формируем запрос
    site_clean = site.replace('www.', '').strip() if site else ''
    if site_clean:
        query = f'{name} Казахстан краска ЛКМ контакты email {site_clean}'
    else:
        query = f'{name} Казахстан ЛКМ дистрибьютор контакты email телефон'
    
    print(f"\n[{processed+1}] {name[:50]} (строка {ri})")
    print(f"  Query: {query[:70]}")
    
    results = search_web(query, num=6)
    print(f"  Результатов: {len(results)}")
    
    info = extract_contacts(results, name)
    
    if info['email']:
        print(f"  ✅ Email: {info['email']}")
    else:
        print(f"  ❌ Email не найден")
    if info['phone']:
        print(f"  ✅ Тел: {info['phone']}")
    if info['site']:
        print(f"  Сайт: {info['site']}")
    
    progress[str(ri)] = info
    
    # Сохраняем прогресс
    with open(PROGRESS_FILE, 'w', encoding='utf-8') as f:
        json.dump(progress, f, ensure_ascii=False, indent=2)
    
    processed += 1
    time.sleep(8)

# === ЗАПИСЬ В EXCEL ===
print(f"\n{'='*70}")
print(f"Записываю {len(progress)} контактов в Excel...")

thin = Side(style='thin', color='D1D5DB')
border = Border(left=thin, right=thin, top=thin, bottom=thin)
yellow_fill = PatternFill('solid', fgColor='FEF3C7')

for ri_str, info in progress.items():
    ri = int(ri_str)
    if info.get('phone'):
        cell = ws.cell(ri, 8)
        if not cell.value or not str(cell.value).strip():
            cell.value = info['phone']
            cell.font = Font(name='Arial', size=10, bold=True, color='16A34A')
            cell.fill = yellow_fill
            cell.alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)
    if info.get('email'):
        cell = ws.cell(ri, 9)
        if not cell.value or not str(cell.value).strip():
            cell.value = info['email']
            cell.font = Font(name='Arial', size=10, bold=True, color='16A34A')
            cell.fill = yellow_fill
            cell.alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)
    if info.get('site'):
        cell = ws.cell(ri, 10)
        if not cell.value or not str(cell.value).strip():
            cell.value = info['site']
            cell.font = Font(name='Arial', size=10, color='1A1F36')
            cell.fill = yellow_fill
            cell.alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)

wb.save(OUT)
print(f"OK: {OUT}")
print(f"Размер: {os.path.getsize(OUT) / 1024:.1f} KB")

# Статистика
filled_email = sum(1 for v in progress.values() if v.get('email'))
filled_phone = sum(1 for v in progress.values() if v.get('phone'))
print(f"\nОбработано: {len(progress)}")
print(f"  Email найден: {filled_email}")
print(f"  Телефон найден: {filled_phone}")
print(f"  Осталось: {len(companies) - len(progress)}")
