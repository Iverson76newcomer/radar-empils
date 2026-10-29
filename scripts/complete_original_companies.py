#!/usr/bin/env python3
"""
Дополнение контактов (email, телефон, сайт, адрес, ассортимент) для 87 исходных компаний.
Стратегия: web_search по каждой компании с указанием страны, извлечение контактов из сниппетов.
Пауза 5 сек между запросами, чтобы не ловить 429.
"""

import subprocess
import json
import re
import time
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from copy import copy

SRC_EXCEL = '/home/z/my-project/download/Экспорт_потенциальные_дистрибьюторы_с_Сербией.xlsx'
OUT = '/home/z/my-project/download/Экспорт_потенциальные_дистрибьюторы_полная.xlsx'
PROGRESS_FILE = '/home/z/my-project/scripts/complete_progress.json'


def search_web(query, num=5):
    try:
        cmd = ['z-ai', 'function', '-n', 'web_search', '-a', json.dumps({'query': query, 'num': num})]
        r = subprocess.run(cmd, capture_output=True, text=True, timeout=30)
        if r.returncode != 0:
            print(f'    search_web error: rc={r.returncode}', flush=True)
            return []
        output = r.stdout
        idx = output.find('[')
        end_idx = output.rfind(']')
        if idx < 0 or end_idx < idx:
            return []
        json_str = output[idx:end_idx+1]
        try:
            return json.loads(json_str)
        except Exception as e:
            print(f'    JSON parse error: {e}', flush=True)
            return []
    except subprocess.TimeoutExpired:
        print(f'    search_web timeout', flush=True)
        return []
    except Exception as e:
        print(f'    search_web exception: {e}', flush=True)
        return []


def extract_contacts(name, country, results):
    """Извлекает email, телефон, сайт, адрес, ассортимент из результатов поиска."""
    info = {
        'email': '', 'phone': '', 'site': '', 'address': '', 
        'assortment': '', 'features': '', 'about': ''
    }
    
    if not results:
        return info
    
    # Список "плохих" сайтов — доски объявлений, википедия, справочники
    bad_hosts = ['wikipedia', 'eponuda', 'infobel', 'yelp', 'yellowpages', 'kompass',
                 'linkedin', 'facebook', 'instagram', 'youtube', 'tiktok', 'telegram',
                 'rabota', 'hh.ru', 'glassdoor', 'tripadvisor', 'stroyka.md', 'all.biz',
                 'bizorg', 'flagma', 'optlist', 'satu.kz', 'pulscen', 'olx', 'lalafo']
    
    # Сначала отфильтровать плохие
    good_results = []
    for r in results:
        host = r.get('host_name', '').lower()
        url = r.get('url', '').lower()
        if not any(bad in host or bad in url for bad in bad_hosts):
            good_results.append(r)
    
    # Если все отфильтровали — берём все
    if not good_results:
        good_results = results
    
    # Берём лучший результат (первый с совпадением по имени)
    best = None
    name_lower = name.lower()
    name_words = [w for w in name_lower.split() if len(w) > 3]
    
    for r in good_results:
        title = r.get('name', '').lower()
        snippet = r.get('snippet', '').lower()
        matches = sum(1 for w in name_words if w in title or w in snippet)
        if matches >= 1:
            best = r
            break
    if not best:
        best = good_results[0]
    
    combined_all = ''  # все сниппеты для извлечения
    for r in good_results[:3]:  # берём первые 3
        combined_all += ' ' + r.get('name', '') + ' ' + r.get('snippet', '')
    
    combined = (best.get('name', '') + ' ' + best.get('snippet', '')).strip()
    url = best.get('url', '')
    host = best.get('host_name', '')
    
    # Email — ищем во всех сниппетах
    emails = re.findall(r'[\w.+-]+@[\w-]+\.[\w.-]+', combined_all)
    if emails:
        clean_emails = [e for e in emails if not e.endswith('.') 
                       and 'sentry' not in e and 'wixpress' not in e
                       and 'example.com' not in e and 'sentry.io' not in e
                       and 'noreply' not in e]
        if clean_emails:
            info['email'] = clean_emails[0]
    
    # Телефон — ищем во всех сниппетах
    phones = []
    phone_patterns = [
        r'\+\d{1,3}[\s\-]?\(?\d{2,3}\)?[\s\-]?\d{3}[\s\-]?\d{2}[\s\-]?\d{2}',
        r'\+\d{1,3}[\s\-]?\d{2}[\s\-]?\d{3}[\s\-]?\d{3}',
        r'\+\d{1,3}[\s\-]?\d{3}[\s\-]?\d{3}[\s\-]?\d{2,3}',
    ]
    for pat in phone_patterns:
        phones.extend(re.findall(pat, combined_all))
    if phones:
        info['phone'] = phones[0]
    
    # Сайт — только официальный
    if url and url.startswith('http'):
        site = url.split('/')[2] if '//' in url else url
        site = re.sub(r'^www\.', '', site)
        # Если это доска — пропускаем
        if not any(bad in site.lower() for bad in bad_hosts):
            info['site'] = site
    if not info['site'] and host and not any(bad in host.lower() for bad in bad_hosts):
        info['site'] = re.sub(r'^www\.', '', host)
    
    # Адрес
    addr_patterns = [
        r'(?:ул\.|улица|str\.|street|адрес|address)[^.,;]{5,80}',
        r'(?:белград|кишинев|bucharest|bucurești|bucuresti|минск|novi sad|нови-сад)[^.,;]{0,80}',
    ]
    for pat in addr_patterns:
        m = re.search(pat, combined_all, re.IGNORECASE)
        if m:
            info['address'] = m.group(0).strip()[:120]
            break
    
    # Ассортимент — ищем во всех
    text_lower = combined_all.lower()
    products = []
    if any(w in text_lower for w in ['акрил', 'acrilic', 'acrylic']):
        products.append('акриловые краски')
    if any(w in text_lower for w in ['алкид', 'alkyd', 'эмаль', 'email']):
        products.append('алкидные эмали')
    if any(w in text_lower for w in ['лак', 'lak', 'lac', 'varnish']):
        products.append('лаки')
    if any(w in text_lower for w in ['грунт', 'grund', 'primer']):
        products.append('грунтовки')
    if any(w in text_lower for w in ['фасад', 'facad', 'facade']):
        products.append('фасадные')
    if any(w in text_lower for w in ['интерьер', 'interior', 'внутренн']):
        products.append('интерьерные')
    if any(w in text_lower for w in ['декоратив', 'decorativ']):
        products.append('декоративные')
    if any(w in text_lower for w in ['молотк', 'hammer']):
        products.append('молотковые')
    if any(w in text_lower for w in ['маслян', 'oil based', 'ulei']):
        products.append('масляные')
    if any(w in text_lower for w in ['водоэмульс', 'water', 'latex']):
        products.append('водоэмульсионные')
    if not products:
        products.append('ЛКМ (общий ассортимент)')
    info['assortment'] = ', '.join(products[:5])
    
    # Особенности
    features = []
    if any(w in text_lower for w in ['опт', 'wholesale', 'veleprodaja']):
        features.append('опт')
    if any(w in text_lower for w in ['розниц', 'retail', 'maloprodaja']):
        features.append('розница')
    if any(w in text_lower for w in ['импорт', 'import']):
        features.append('импортёр')
    if any(w in text_lower for w in ['доставк', 'delivery', 'isporuka']):
        features.append('доставка')
    if any(w in text_lower for w in ['колеровк', 'colorant']):
        features.append('колеровка')
    info['features'] = ', '.join(features) if features else 'опт'
    
    # О компании (берём сниппет)
    snippet = best.get('snippet', '')
    snippet = re.sub(r'\s+', ' ', snippet).strip()
    if snippet:
        info['about'] = snippet[:250]
    
    return info


# === ОТКРЫВАЕМ EXCEL ===
print("Открываю Excel...")
wb = openpyxl.load_workbook(SRC_EXCEL)
ws = wb['Лист1']

# Загружаем прогресс (если есть)
progress = {}
import os
if os.path.exists(PROGRESS_FILE):
    with open(PROGRESS_FILE, 'r', encoding='utf-8') as f:
        progress = json.load(f)
    print(f"Загружен прогресс: {len(progress)} компаний уже обработано")

# Список исходных компаний с пропусками
col_names = {5: 'О компании', 6: 'Ключевые особенности', 7: 'Ассортимент', 
             8: 'Адрес', 9: 'Телефоны', 10: 'Email', 11: 'Сайт'}

companies_to_complete = []
for ri in range(4, 92):
    country = ws.cell(ri, 2).value
    name = ws.cell(ri, 4).value
    if not name:
        continue
    missing = []
    for ci in [5, 6, 7, 8, 9, 10, 11]:
        val = ws.cell(ri, ci).value
        if not val or not str(val).strip():
            missing.append(ci)
    if missing:
        companies_to_complete.append((ri, country, name, missing))

print(f"Компаний для дополнения: {len(companies_to_complete)}")
print(f"Из них уже обработано: {len(progress)}")
print()

# Обработка
processed = 0
MAX_PER_RUN = 8  # максимум компаний за один запуск (защита от OOM)
for ri, country, name, missing_cols in companies_to_complete:
    if processed >= MAX_PER_RUN:
        print(f"\nДостигнут лимит {MAX_PER_RUN} компаний за запуск. Сохраняю и завершаюсь.", flush=True)
        break
    name_clean = str(name).strip()
    # Пропускаем уже обработанные
    if str(ri) in progress:
        continue
    
    # Строим запрос
    country_short = str(country).strip()
    query = f'"{name_clean}" {country_short} краска ЛКМ контакты'
    if country in ['Сербия', 'Румыния', 'Молдова']:
        query = f'"{name_clean}" paint distributor {country_short} contacts'
    
    print(f"[{ri-3}/{len(companies_to_complete)}] {name_clean[:50]} ({country_short})", flush=True)
    
    results = search_web(query, num=5)
    info = extract_contacts(name_clean, country, results)
    
    # Сохраняем что нашли
    progress[str(ri)] = {
        'name': name_clean,
        'country': country_short,
        'info': info,
        'missing_cols': missing_cols,
    }
    
    # Пишем в Excel только то, что было пустым
    col_map = {5: 'about', 6: 'features', 7: 'assortment', 8: 'address', 
               9: 'phone', 10: 'email', 11: 'site'}
    for ci in missing_cols:
        field = col_map[ci]
        val = info.get(field, '')
        if val:
            cell = ws.cell(ri, ci, val)
            # Копируем стиль из первой ячейки той же строки
            sample_cell = ws.cell(ri, 2)
            cell.font = copy(sample_cell.font)
            cell.fill = copy(sample_cell.fill)
            cell.alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)
            thin = Side(style='thin', color='D1D5DB')
            cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)
    
    # Сохраняем прогресс после каждой компании (надёжнее)
    processed += 1
    with open(PROGRESS_FILE, 'w', encoding='utf-8') as f:
        json.dump(progress, f, ensure_ascii=False, indent=2)
    if processed % 5 == 0:
        wb.save(OUT)
        print(f"    → сохранено (прогресс: {processed}/{len(companies_to_complete)})", flush=True)
    
    time.sleep(8)  # пауза 8 сек, чтобы не ловить 429

# Финальное сохранение
wb.save(OUT)
with open(PROGRESS_FILE, 'w', encoding='utf-8') as f:
    json.dump(progress, f, ensure_ascii=False, indent=2)

print(f"\n{'='*70}")
print(f"ГОТОВО")
print(f"{'='*70}")
print(f"Обработано: {processed} компаний")
print(f"Файл: {OUT}")

# Статистика
filled_count = 0
for ri_str, data in progress.items():
    info = data.get('info', {})
    if info.get('email') or info.get('phone') or info.get('site'):
        filled_count += 1
print(f"Дополнено контактов: {filled_count} из {len(progress)}")
