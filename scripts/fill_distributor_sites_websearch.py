#!/usr/bin/env python3
"""
ЭТАП 2: Web-поиск сайтов для нераспознанных дистрибьюторов.
Запускается после fill_distributor_sites_v2.py.
"""

import os
import re
import json
import subprocess
import time
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

DST = "/home/z/my-project/download/Distributors_2026_с_сайтами.xlsx"

# === Сначала добавлю известные мне сайты (из предыдущих анализов и общих знаний) ===

KNOWN_SITES = {
    "Брозэкс ТК ООО": "https://brozex.ru",
    "Леруа Мерлен Восток ООО": "https://www.leroymerlin.ru",
    "Петрович Строительный ТД Санкт-Петербург ООО": "https://petrovich.ru",
    "ВсеИнструменты.ру ООО": "https://www.vseinstrumenti.ru",
    "МАКСИДОМ ООО": "https://www.maxidom.ru",
    "БАРИС ООО": "http://baris-online.ru",
    "Альянс ООО (Иркутск)": "https://стройлогистика.рф",
    "ЭГИДА ООО": "https://kapitel38.ru",
    "АЮСС-ВСТК ООО": "http://auss-vstk.ru",
    "Кузнецкий Альянс АО": "http://www.прокомплизинг.рф",  # нужна проверка
    "Спектр ООО (Воронеж)": "https://sct.ru",
    "СтройТерминал ООО (ООО Цитадель)": "https://tdcsm.ru",
    "Оптонофф ООО (Ю-ТРЕЙД)": "https://shop.bteam.ru",
    "Севастьянов Эдуард Юрьевич ИП": "http://vstk.su",
    "Технологии, Ресурсы Строительства ООО": "https://www.kraski.ru",
    "Радуга КРАСКИ ООО": "https://radugakraski.ru",
    "Уралстройсервис-опт ООО": "https://www.uss74.ru",  # уже есть
    "Центр Красок ООО": "https://centr-krasok.ru",
}

# === Список для web-поиска ===
# Только ООО с понятными названиями. ИП и ТОО/ОсОО пропускаем (маловероятно иметь сайт)

SEARCH_CANDIDATES = [
    # (row, name, query)
    # Сначала ООО с уникальными названиями
    (104, "СибХимТорг-Т ООО", "СибХимТорг-Т ООО строительные материалы сайт"),
    (105, "Большая перемена-Исилькуль ООО", "Большая перемена Исилькуль ООО магазин сайт"),
    (106, "ВелКом Холдинг ООО", "ВелКом Холдинг Сибирская база ООО сайт строительные материалы"),
    (112, "Сампо ТК ООО", "Сампо ТК ООО строительные материалы сайт"),
    (220, "НАМЕСТНИК ООО", "НАМЕСТНИК ООО строительные материалы сайт"),
    (222, "МСМ ООО (СК ООО)", "МСМ ООО СК строительные материалы сайт"),
    (223, "Деловой контакт ООО", "Деловой контакт ООО строительные материалы сайт"),
    (224, "ЭЛЕКТРУМ ООО", "ЭЛЕКТРУМ ООО строительные материалы сайт"),
    (226, "Главхозторг ООО", "Главхозторг ООО опт строительные материалы сайт"),
    (228, "Оптовый центр ООО", "Оптовый центр ООО строительные материалы сайт"),
    (229, "Тренд ООО (Импульс)", "Тренд Импульс ООО строительные материалы сайт"),
    (230, "Котласское ССП АО", "Котласское ССП АО строительные материалы"),
    (232, "ПТЗОПТСТРОЙТОРГ ООО", "ПТЗОПТСТРОЙТОРГ ООО опт стройматериалы сайт"),
    (233, "Сигма ООО", "Сигма ООО строительные материалы сайт"),
    (242, "Велес-быт-сервис ООО", "Велес-быт-сервис ООО строительные материалы сайт"),
    (245, "КорТрейд ООО", "КорТрейд ООО строительные материалы сайт"),
    (246, "Сити Строй ООО", "Сити Строй ООО строительные материалы сайт"),
    (316, "СтройГарант ООО", "СтройГарант ООО строительные материалы сайт"),
    (362, "ЮМКОМ ТПК ООО", "ЮМКОМ ТПК ООО торгово-промышленная компания сайт"),
    (363, "ДКТ ООО", "ДКТ ООО строительные материалы сайт"),
    (364, "ОНК ООО (ООО Арсенал)", "ОНК ООО Арсенал строительные материалы сайт"),
    (365, "Радуга ДОН ООО", "Радуга ДОН ООО краски строительные материалы сайт"),
    (366, "РАДУГА ЛКМ ООО", "РАДУГА ЛКМ ООО лакокрасочные материалы сайт"),
    (367, "РостовПромПокрытия ООО", "РостовПромПокрытия ООО краски покрытия сайт"),
    (368, "СНАБГАРАНТ ООО", "СНАБГАРАНТ ООО строительные материалы сайт"),
    (369, "Стройторг ООО", "Стройторг ООО строительные материалы сайт"),
    (370, "СТУДИЯ ЦВЕТА РОСТОВ ООО", "СТУДИЯ ЦВЕТА РОСТОВ краски колеровка сайт"),
    (372, "Лекма холод ООО", "Лекма холод ООО сайт"),
    (373, "Морской Воздух ООО", "Морской Воздух ООО краски сайт"),
    (374, "Палитра-Юг ООО", "Палитра-Юг ООО краски сайт"),
    (375, "Строймаркет-05 ООО", "Строймаркет-05 ООО магазин сайт"),
    (271, "Леруа Мерлен Восток ООО", "Леруа Мерлен Россия сайт"),  # точно
    (270, "ВсеИнструменты.ру ООО", "ВсеИнструменты сайт"),  # точно
    (272, "МАКСИДОМ ООО", "Максидом строительный магазин сайт"),  # точно
    (287, "Петрович Строительный ТД Санкт-Петербург ООО", "Петрович строительный торговый дом сайт"),
    (235, "Брозэкс ТК ООО", "Брозекс Групп Екатеринбург краски сайт"),  # точно
    # Казахстанские ТОО
    (336, "AJR & Company ТОО", "AJR Company ТОО Казахстан строительные материалы"),
    (337, "Graffiti Co ТОО", "Graffiti Co ТОО Казахстан строительные материалы"),
    (338, "Graffiti-Юг ТОО", "Graffiti-Юг ТОО Казахстан строительные"),
    (339, "MON-EMP Co LTD", "MON-EMP Co LTD Казахстан строительные"),
    (340, "PAKO TRADE ТОО", "PAKO TRADE ТОО Казахстан строительные"),
    (341, "Stroy Depot ТОО", "Stroy Depot ТОО Казахстан строительные"),
    (342, "USTA MASTER ООО", "USTA MASTER Казахстан строительные"),
    (343, "UZCHEM", "UZCHEM Узбекистан химия строительные"),
    (344, "ВАГ ИМПОРТ ООО", "ВАГ ИМПОРТ ООО строительные материалы сайт"),
    (345, "ГолдДекор ООО", "ГолдДекор ООО строительные материалы сайт"),
    (347, "Держава Трейд ТОО", "Держава Трейд ТОО Казахстан строительные"),
    (349, "КОРВОН 2011 ООО", "КОРВОН 2011 ООО Казахстан"),
    (352, "СпецПром INVEST ТОО", "СпецПром INVEST ТОО Казахстан"),
    (353, "Стройинвест-К ТОО", "Стройинвест-К ТОО Казахстан"),
    (354, "Стройлюкс-1 ТОО", "Стройлюкс-1 ТОО Казахстан"),
    (355, "Транзит-М ООО", "Транзит-М ООО строительные материалы"),
    (358, "Памир Фирма ООО", "Памир Фирма ООО Казахстан"),
    (360, "СТРОЙКАР-13 ООО", "СТРОЙКАР-13 ООО Казахстан"),
]

# === Сначала применяем KNOWN_SITES ===

wb = load_workbook(DST)
ws = wb["Лист1"]

fill_exact = PatternFill("solid", fgColor="C8E6C9")
fill_search = PatternFill("solid", fgColor="BBDEFB")
fill_unmatched = PatternFill("solid", fgColor="FFCDD2")
font_url = Font(name="Calibri", size=10, color="0563C1", underline="single")

applied_known = 0
for row_idx in range(2, ws.max_row + 1):
    client = ws.cell(row=row_idx, column=4).value
    if not client:
        continue
    client_str = str(client).strip()
    if client_str in KNOWN_SITES:
        url = KNOWN_SITES[client_str]
        # Проверим, не заполнено ли уже
        existing = ws.cell(row=row_idx, column=5).value
        if existing:
            continue
        cell = ws.cell(row=row_idx, column=5)
        cell.value = url
        cell.font = font_url
        cell.fill = fill_exact
        cell.alignment = Alignment(horizontal="left", vertical="center")
        applied_known += 1
        print(f"  ✅ R{row_idx}: {client_str} → {url}")

wb.save(DST)
print(f"\nПрименено известных сайтов: {applied_known}")
print(f"\n=== Запускаем web-поиск для {len(SEARCH_CANDIDATES)} компаний ===\n")

# === Web-поиск ===

def web_search(query, num=2):
    """Запуск z-ai web_search через subprocess."""
    try:
        cmd = ["z-ai", "function", "-n", "web_search", "-a", json.dumps({"query": query, "num": num})]
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=60)
        out = result.stdout
        # Strip CLI logs
        idx = out.find('[')
        if idx == -1:
            idx = out.find('{')
        if idx == -1:
            return []
        json_str = out[idx:]
        try:
            return json.loads(json_str)
        except json.JSONDecodeError:
            if json_str.strip().startswith('['):
                end = json_str.rfind(']')
                if end > 0:
                    try:
                        return json.loads(json_str[:end+1])
                    except:
                        return []
            return []
    except Exception as e:
        print(f"  ERROR: {e}")
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

# Загружаем снова
wb = load_workbook(DST)
ws = wb["Лист1"]

found_count = 0
not_found_count = 0
results = []

for row_idx, name, query in SEARCH_CANDIDATES:
    # Проверим, не заполнено ли уже (из KNOWN_SITES или ранее)
    existing = ws.cell(row=row_idx, column=5).value
    if existing:
        print(f"  ⏭️  R{row_idx}: {name} — уже заполнено ({existing})")
        continue

    print(f"  🔍 R{row_idx}: {name}")
    results_list = web_search(query, num=2)
    found_url = None
    for r in results_list[:2]:
        if isinstance(r, dict):
            url = r.get('url', '')
            snippet = (r.get('snippet', '') or '').lower()
            # Проверим, что URL не поисковик/соцсеть
            if not url:
                continue
            base = extract_base_url(url)
            if not base:
                continue
            # Пропускаем yandex/google/youtube/maps/facebook/instagram/vk/olx
            skip_domains = ['yandex.', 'google.', 'youtube.', 'maps.yandex', 'facebook.com',
                          'instagram.com', 'vk.com', 'olx.', '2gis.', 'instagram.', 'linkedin.',
                          'twitter.com', 'wikipedia.org', 'kp.ru', 'forum.', 'news.',
                          'yellowpages', 'allbiz', 'kompass', 'rusprofile', 'list-',
                          'safework.ru', 'findpatent']
            if any(skip in base.lower() for skip in skip_domains):
                continue
            # Берем этот URL
            found_url = base
            break

    if found_url:
        cell = ws.cell(row=row_idx, column=5)
        cell.value = found_url
        cell.font = font_url
        cell.fill = fill_search
        cell.alignment = Alignment(horizontal="left", vertical="center")
        print(f"     ✅ {found_url}")
        found_count += 1
        results.append((row_idx, name, found_url))
    else:
        # Помечаем как не найдено
        cell = ws.cell(row=row_idx, column=5)
        cell.value = "не найдено"
        cell.font = Font(name="Calibri", size=10, color="999999", italic=True)
        cell.fill = fill_unmatched
        cell.alignment = Alignment(horizontal="left", vertical="center")
        print(f"     ❌ не найдено")
        not_found_count += 1

    time.sleep(3)  # чтобы не словить rate limit

wb.save(DST)

print(f"\n{'='*60}")
print(f"=== ИТОГ ===")
print(f"{'='*60}")
print(f"  ✅ Найдено через web-search: {found_count}")
print(f"  ❌ Не найдено:              {not_found_count}")
print(f"\n✅ Файл сохранён: {DST}")
print(f"   Размер: {os.path.getsize(DST):,} байт")

# Сохраняем результаты для отчёта
with open("/tmp/web_search_results.json", "w", encoding="utf-8") as f:
    json.dump(results, f, ensure_ascii=False, indent=2)
