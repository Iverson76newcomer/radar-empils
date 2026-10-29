#!/usr/bin/env python3
"""
Пакетный web-поиск для оставшихся компаний с улучшенным фильтром.
Читает /tmp/retry_list.json и обрабатывает по 5 компаний за раз.
"""

import os, json, subprocess, time, sys
from urllib.parse import urlparse
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

DST = "/home/z/my-project/download/Distributors_2026_с_сайтами.xlsx"
LIST_FILE = "/tmp/retry_list.json"

# Загружаем список компаний
with open(LIST_FILE, encoding="utf-8") as f:
    ALL_CANDIDATES = json.load(f)

# Параметры
start_idx = int(sys.argv[1]) if len(sys.argv) > 1 else 0
end_idx = int(sys.argv[2]) if len(sys.argv) > 2 else len(ALL_CANDIDATES)

batch = ALL_CANDIDATES[start_idx:end_idx]
print(f"=== Обработка {len(batch)} компаний (индексы {start_idx}..{end_idx-1}) ===\n")

def web_search(query, num=3):
    try:
        cmd = ["z-ai", "function", "-n", "web_search", "-a", json.dumps({"query": query, "num": num})]
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=60)
        out = result.stdout
        idx = out.find('[')
        if idx == -1:
            return []
        json_str = out[idx:]
        try:
            return json.loads(json_str)
        except json.JSONDecodeError:
            end = json_str.rfind(']')
            if end > 0:
                try:
                    return json.loads(json_str[:end+1])
                except:
                    return []
            return []
    except Exception:
        return []

def extract_base_url(url):
    if not url:
        return None
    s = url.strip()
    if not s.startswith(('http://', 'https://')):
        s = 'http://' + s
    try:
        p = urlparse(s)
        if not p.netloc:
            return None
        return f"{p.scheme}://{p.netloc}"
    except Exception:
        return None

SKIP_DOMAINS = [
    'yandex.', 'google.', 'youtube.', 'facebook.com', 'instagram.com',
    'vk.com', 'olx.', '2gis.', 'linkedin.', 'twitter.com', 'wikipedia.org',
    'kp.ru', 'forum.', 'news.', 'yellowpages', 'allbiz', 'kompass',
    'rusprofile', 'list-', 'safework.ru', 'findpatent', 'maps.yandex',
    'flamp.ru', 'zoon.ru', 'otzovik.com', 'irecommend.ru', 'patee.ru',
    'orgpage.ru', 'firma.gov.ru', 'egrul.ru', '1cont.ru', 'be-informed.ru',
    'sbis.ru', 'kontur.ru', 'nalog.ru', 'rosstat', 'rugion.ru',
    'spark-interfax', 'companies.rbc', 'rbc.ru', 'domclick.ru', 'vsem-podryad',
    'companium.ru', 'agrobase.ru', 'vestnikstroy.ru', 'mir36.ru', 'b2b-',
    'rusbanki.ru', 'banki.ru', 'otzovut', 'otzyvy', 'profi.ru',
    'cataloxy.ru', 'tyndex.ru', 'rusfirm.ru', 'firmsbase.ru',
    'spravker.ru', 'orgspro.ru', 'clickberry.ru', 'fasie.ru', 'shopdel.ru',
    'ruprosvet.ru', 'exibition', 'proportal', 'rusprofile.ru',
    'seldon.ru', 'b2b-center.ru', 'e-disclosure.ru', 'fedresurs.ru',
    'glavbukh.ru', 'tbk.ru', 'torgbase.ru', 'b2brus.ru', 'regionfirm.ru',
    'rus-firm.ru', 'tpprf.ru', 'abirus.ru', 'uchetfin.ru', 'tass.ru',
    'tjournal.ru', 'vc.ru', 'habr.com', 'pikabu.ru', 'livejournal.com',
    'drom.ru', 'auto.ru', 'avito.ru', 'cian.ru', 'pinterest.com',
    't.me', 'telegram.me', 'whatsapp.com', 'wechat.com',
]

def is_likely_real_site(base_url, company_name, snippet):
    """Дополнительная проверка, что сайт реальный, а не справочник."""
    base_lower = base_url.lower()
    snippet_lower = (snippet or '').lower()
    name_lower = company_name.lower()

    # Если домен содержит часть названия компании — хорошо
    # Извлечём main part домена (без www. и zone)
    import re
    m = re.match(r'https?://(?:www\.)?([a-z0-9\-_]+)', base_lower)
    if not m:
        return False
    domain_main = m.group(1)
    # Проверим, что имя компании содержит main часть домена или наоборот
    # (для русских компаний — транслитерация)
    # Простая проверка: если snippet содержит название компании целиком или его часть
    name_tokens = re.split(r'[\s\-,()\.]+', name_lower)
    name_tokens = [t for t in name_tokens if len(t) >= 4 and t not in ('ооо', 'тк', 'оао', 'зао', 'ип', 'тд', 'компания', 'санкт-петербург', 'строительный', 'холдинг', 'опт')]

    # Если хотя бы 2 токена из имени в snippet — считаем реальным
    matched_tokens = sum(1 for t in name_tokens if t in snippet_lower)
    if matched_tokens >= 2:
        return True
    # Если домен содержит транслит имени — OK
    # Простые соответствия
    translit_map = {
        'строй': 'stroy', 'стройка': 'stroyka', 'краска': 'kraska', 'краски': 'kraski',
        'дом': 'dom', 'торг': 'torg', 'опт': 'opt', 'база': 'baza',
        'спец': 'spec', 'пром': 'prom', 'промышленность': 'prom',
    }
    for ru, en in translit_map.items():
        if ru in name_lower and en in domain_main:
            return True
    return False

wb = load_workbook(DST)
ws = wb["Лист1"]

fill_search = PatternFill("solid", fgColor="BBDEFB")
fill_unmatched = PatternFill("solid", fgColor="FFCDD2")
font_url = Font(name="Calibri", size=10, color="0563C1", underline="single")
font_notfound = Font(name="Calibri", size=10, color="999999", italic=True)

found = 0
not_found = 0

for row_idx_str, name, query in batch:
    row_idx = int(row_idx_str)
    # Проверим, не заполнено ли уже
    existing = ws.cell(row=row_idx, column=5).value
    if existing and str(existing).strip() and str(existing).strip() != "не найдено":
        print(f"  ⏭️  R{row_idx}: {name} — уже есть: {existing}")
        continue

    print(f"  🔍 R{row_idx}: {name}")
    # Очистим если было "не найдено"
    if existing and str(existing).strip() == "не найдено":
        ws.cell(row=row_idx, column=5).value = None

    results_list = web_search(query, num=5)
    found_url = None
    for r in results_list[:5]:
        if isinstance(r, dict):
            url = r.get('url', '')
            snippet = r.get('snippet', '') or ''
            if not url:
                continue
            base = extract_base_url(url)
            if not base:
                continue
            # Пропускаем нежелательные домены
            if any(skip in base.lower() for skip in SKIP_DOMAINS):
                continue
            # Проверим, что это реальный сайт компании
            if is_likely_real_site(base, name, snippet):
                found_url = base
                break

    if found_url:
        cell = ws.cell(row=row_idx, column=5)
        cell.value = found_url
        cell.font = font_url
        cell.fill = fill_search
        cell.alignment = Alignment(horizontal="left", vertical="center")
        print(f"     ✅ {found_url}")
        found += 1
    else:
        cell = ws.cell(row=row_idx, column=5)
        cell.value = "не найдено"
        cell.font = font_notfound
        cell.fill = fill_unmatched
        cell.alignment = Alignment(horizontal="left", vertical="center")
        print(f"     ❌ не найдено")
        not_found += 1

    wb.save(DST)
    time.sleep(7)

print(f"\n{'='*60}")
print(f"  ✅ Найдено: {found}")
print(f"  ❌ Не найдено: {not_found}")
