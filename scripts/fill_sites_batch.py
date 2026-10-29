#!/usr/bin/env python3
"""
ЭТАП 2 (v2): Web-поиск с увеличенными паузами.
Запускаем по 5 компаний за раз, пауза 8 секунд между запросами.
"""

import os
import re
import json
import subprocess
import time
import sys
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

DST = "/home/z/my-project/download/Distributors_2026_с_сайтами.xlsx"

# Список для web-поиска — только те, что точно не из KNOWN_SITES
# (KNOWN_SITES уже применены на предыдущем шаге)
# Разбиваю на порции по 10 компаний
SEARCH_CANDIDATES = [
    # Российские ООО
    (104, "СибХимТорг-Т ООО", "СибХимТорг-Т ООО Омск сайт"),
    (105, "Большая перемена-Исилькуль ООО", "Большая перемена Исилькуль ООО магазин"),
    (106, "ВелКом Холдинг ООО", "ВелКом Холдинг Сибирская база ООО сайт"),
    (112, "Сампо ТК ООО", "Сампо ТК ООО Санкт-Петербург строительные"),
    (220, "НАМЕСТНИК ООО", "НАМЕСТНИК ООО Краснодар строительные"),
    (222, "МСМ ООО (СК ООО)", "МСМ СК ООО Москва строительные"),
    (223, "Деловой контакт ООО", "Деловой контакт ООО Москва стройматериалы"),
    (224, "ЭЛЕКТРУМ ООО", "ЭЛЕКТРУМ ООО строительные материалы"),
    (226, "Главхозторг ООО", "Главхозторг ООО опт стройматериалы"),
    (228, "Оптовый центр ООО", "Оптовый центр ООО стройматериалы"),
    (229, "Тренд ООО (Импульс)", "Тренд Импульс ООО стройматериалы"),
    (230, "Котласское ССП АО", "Котласское ССП АО Архангельск"),
    (232, "ПТЗОПТСТРОЙТОРГ ООО", "ПТЗОПТСТРОЙТОРГ ООО стройматериалы"),
    (233, "Сигма ООО", "Сигма ООО стройматериалы опт"),
    (242, "Велес-быт-сервис ООО", "Велес-быт-сервис ООО стройматериалы"),
    (245, "КорТрейд ООО", "КорТрейд ООО стройматериалы"),
    (246, "Сити Строй ООО", "Сити Строй ООО стройматериалы"),
    (316, "СтройГарант ООО", "СтройГарант ООО стройматериалы"),
    (362, "ЮМКОМ ТПК ООО", "ЮМКОМ ТПК ООО торговая компания"),
    (363, "ДКТ ООО", "ДКТ ООО стройматериалы"),
    (364, "ОНК ООО (ООО Арсенал)", "ОНК Арсенал ООО стройматериалы"),
    (365, "Радуга ДОН ООО", "Радуга ДОН ООО краски"),
    (366, "РАДУГА ЛКМ ООО", "РАДУГА ЛКМ ООО лакокрасочные"),
    (367, "РостовПромПокрытия ООО", "РостовПромПокрытия ООО"),
    (368, "СНАБГАРАНТ ООО", "СНАБГАРАНТ ООО стройматериалы"),
    (369, "Стройторг ООО", "Стройторг ООО стройматериалы"),
    (370, "СТУДИЯ ЦВЕТА РОСТОВ ООО", "СТУДИЯ ЦВЕТА РОСТОВ краски колеровка"),
    (372, "Лекма холод ООО", "Лекма холод ООО"),
    (373, "Морской Воздух ООО", "Морской Воздух ООО краски"),
    (374, "Палитра-Юг ООО", "Палитра-Юг ООО краски"),
    (375, "Строймаркет-05 ООО", "Строймаркет-05 ООО магазин"),
    # Казахстанские ТОО
    (336, "AJR & Company ТОО", "AJR Company ТОО Казахстан стройматериалы"),
    (337, "Graffiti Co ТОО", "Graffiti Co ТОО Казахстан"),
    (338, "Graffiti-Юг ТОО", "Graffiti-Юг ТОО Казахстан"),
    (340, "PAKO TRADE ТОО", "PAKO TRADE ТОО Казахстан"),
    (341, "Stroy Depot ТОО", "Stroy Depot ТОО Казахстан стройматериалы"),
    (342, "USTA MASTER ООО", "USTA MASTER Казахстан стройматериалы"),
    (343, "UZCHEM", "UZCHEM Узбекистан химия"),
    (344, "ВАГ ИМПОРТ ООО", "ВАГ ИМПОРТ ООО стройматериалы"),
    (345, "ГолдДекор ООО", "ГолдДекор ООО стройматериалы"),
    (347, "Держава Трейд ТОО", "Держава Трейд ТОО Казахстан"),
    (349, "КОРВОН 2011 ООО", "КОРВОН 2011 ООО Казахстан"),
    (352, "СпецПром INVEST ТОО", "СпецПром INVEST ТОО Казахстан"),
    (353, "Стройинвест-К ТОО", "Стройинвест-К ТОО Казахстан"),
    (354, "Стройлюкс-1 ТОО", "Стройлюкс-1 ТОО Казахстан"),
    (355, "Транзит-М ООО", "Транзит-М ООО стройматериалы"),
    (358, "Памир Фирма ООО", "Памир Фирма ООО стройматериалы"),
    (360, "СТРОЙКАР-13 ООО", "СТРОЙКАР-13 ООО Казахстан"),
    # Пропущенные/крупные
    (235, "Брозэкс ТК ООО", "Брозэкс Групп Екатеринбург сайт"),
    (271, "Леруа Мерлен Восток ООО", "Леруа Мерлен Россия сайт"),
    (270, "ВсеИнструменты.ру ООО", "ВсеИнструменты.ру сайт"),
    (272, "МАКСИДОМ ООО", "Максидом строительный магазин сайт"),
    (287, "Петрович Строительный ТД Санкт-Петербург ООО", "Петрович строительный торговый дом сайт"),
    # Прочие
    (60, "Роща ООО", "Роща ООО Иркутск стройматериалы"),
    (61, "Символ ООО", "Символ ООО Иркутск стройматериалы метр"),
    (81, "Крипт ТК ООО", "Крипт ТК ООО стройматериалы"),
    # ИП (попробуем некоторые — у которых есть trademark-имя)
    # Пропускаем ИП — у них обычно нет сайтов
]

def web_search(query, num=2):
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
    except Exception as e:
        return []

def extract_base_url(url):
    if not url:
        return None
    s = url.strip()
    if not s.startswith(('http://', 'https://')):
        s = 'http://' + s
    try:
        from urllib.parse import urlparse
        p = urlparse(s)
        if not p.netloc:
            return None
        return f"{p.scheme}://{p.netloc}"
    except Exception:
        return None

SKIP_DOMAINS = ['yandex.', 'google.', 'youtube.', 'facebook.com', 'instagram.com',
                'vk.com', 'olx.', '2gis.', 'linkedin.', 'twitter.com', 'wikipedia.org',
                'kp.ru', 'forum.', 'news.', 'yellowpages', 'allbiz', 'kompass',
                'rusprofile', 'list-', 'safework.ru', 'findpatent', 'maps.yandex',
                'flamp.ru', 'zoon.ru', 'otzovik.com', 'irecommend.ru', 'patee.ru',
                'orgpage.ru', 'firma.gov.ru', 'egrul.ru', '1cont.ru', 'be-informed.ru',
                'sbis.ru', 'kontur.ru', 'nalog.ru', 'rosstat', 'rugion.ru',
                # Каталоги и справочники (часто ложные совпадения)
                'spark-interfax', 'companies.rbc', 'rbc.ru', 'domclick.ru', 'vsem-podryad',
                'companium.ru', 'agrobase.ru', 'vestnikstroy.ru', 'mir36.ru', 'b2b-',
                'rusbanki.ru', 'banki.ru', 'vseinstrumenti.ru/', 'otzovut', 'otzyvy',
                'profi.ru', 'cataloxy.ru', 'tyndex.ru', 'rusfirm.ru', 'firmsbase.ru',
                'moskva.spravker.ru', 'spravker.ru', 'orgspro.ru', 'clickberry.ru',
                'fasie.ru', 'digital.gov.ru', 'shopdel.ru', 'ruprosvet.ru']

# Параметры: какие компании обрабатывать (передаются как аргументы)
# Пример: python script.py 0 15  → обработать индексы 0..14
start_idx = int(sys.argv[1]) if len(sys.argv) > 1 else 0
end_idx = int(sys.argv[2]) if len(sys.argv) > 2 else len(SEARCH_CANDIDATES)

batch = SEARCH_CANDIDATES[start_idx:end_idx]
print(f"=== Обработка индексов {start_idx}..{end_idx-1} ({len(batch)} компаний) ===\n")

wb = load_workbook(DST)
ws = wb["Лист1"]

fill_search = PatternFill("solid", fgColor="BBDEFB")
fill_unmatched = PatternFill("solid", fgColor="FFCDD2")
font_url = Font(name="Calibri", size=10, color="0563C1", underline="single")
font_notfound = Font(name="Calibri", size=10, color="999999", italic=True)

found = 0
not_found = 0

for row_idx, name, query in batch:
    # Проверим, не заполнено ли уже
    existing = ws.cell(row=row_idx, column=5).value
    if existing and str(existing).strip() and str(existing).strip() != "не найдено":
        print(f"  ⏭️  R{row_idx}: {name} — уже есть: {existing}")
        continue

    print(f"  🔍 R{row_idx}: {name}")

    # Очистим ячейку если там "не найдено"
    if existing and str(existing).strip() == "не найдено":
        ws.cell(row=row_idx, column=5).value = None

    results_list = web_search(query, num=3)
    found_url = None
    for r in results_list[:3]:
        if isinstance(r, dict):
            url = r.get('url', '')
            snippet = (r.get('snippet', '') or '').lower()
            name_lower = name.lower()

            base = extract_base_url(url)
            if not base:
                continue

            # Пропускаем нежелательные домены
            if any(skip in base.lower() for skip in SKIP_DOMAINS):
                continue

            # Дополнительная проверка: snippet должен содержать часть названия компании
            # (не строго — хотя бы одно значимое слово)
            name_tokens = [t for t in re.split(r'[\s\-,()\.]+', name_lower) if len(t) >= 4 and t not in ('ооо', 'тк', 'оао', 'зао', 'ип', 'тд', 'компания', 'санкт-петербург', 'строительный')]
            if name_tokens:
                matched_token = any(t in snippet for t in name_tokens)
                if not matched_token:
                    continue

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

    # Сохраняем после каждой компании (для безопасности)
    wb.save(DST)
    time.sleep(6)  # пауза 6 сек между запросами

print(f"\n{'='*60}")
print(f"=== ИТОГ ПАРТИИ {start_idx}-{end_idx-1} ===")
print(f"  ✅ Найдено: {found}")
print(f"  ❌ Не найдено: {not_found}")
print(f"  Файл сохранён: {DST}")
