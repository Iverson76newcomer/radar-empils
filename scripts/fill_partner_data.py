#!/usr/bin/env python3
"""
Заполнение белых пятен в таблице партнёров Эмпилс.
Поиск по ИНН и названию компании через интернет.
"""

import openpyxl
import subprocess
import json
import re
import time
import os

# === Маппинг городов/регионов → Федеральные округа ===
CITY_TO_FO = {
    # Северо-Западный ФО
    'санкт-петербург': 'Северо-Западный',
    'петрозаводск': 'Северо-Западный',
    'мурманск': 'Северо-Западный',
    'калуга': 'Центральный',
    'псков': 'Северо-Западный',
    'новгород': 'Северо-Западный',
    'великий новгород': 'Северо-Западный',
    'сыктывкар': 'Северо-Западный',
    'архангельск': 'Северо-Западный',
    'вологда': 'Северо-Западный',
    'череповец': 'Северо-Западный',
    'котлас': 'Северо-Западный',
    'выборг': 'Северо-Западный',
    'калининград': 'Северо-Западный',
    # Центральный ФО
    'москва': 'Центральный',
    'воронеж': 'Центральный',
    'ярославль': 'Центральный',
    'тверь': 'Центральный',
    'тула': 'Центральный',
    'рязань': 'Центральный',
    'брянск': 'Центральный',
    'смоленск': 'Центральный',
    'калуга': 'Центральный',
    'иваново': 'Центральный',
    'владимир': 'Центральный',
    'кострома': 'Центральный',
    'орёл': 'Центральный',
    'курск': 'Центральный',
    'липецк': 'Центральный',
    'белгород': 'Центральный',
    'тамббов': 'Центральный',
    'тамбов': 'Центральный',
    'подольск': 'Центральный',
    'красногорск': 'Центральный',
    # Уральский ФО
    'екатеринбург': 'Уральский',
    'челябинск': 'Уральский',
    'тюмень': 'Уральский',
    'нижний тагил': 'Уральский',
    'магнитогорск': 'Уральский',
    'сургут': 'Уральский',
    'нижневартовск': 'Уральский',
    # Сибирский ФО
    'новосибирск': 'Сибирский',
    'красноярск': 'Сибирский',
    'иркутск': 'Сибирский',
    'омск': 'Сибирский',
    'барнаул': 'Сибирский',
    'новокузнецк': 'Сибирский',
    'кемерово': 'Сибирский',
    'томск': 'Сибирский',
    'норильск': 'Сибирский',
    'братск': 'Сибирский',
    'абакан': 'Сибирский',
    # Дальневосточный ФО
    'владивосток': 'Дальневосточный',
    'хабаровск': 'Дальневосточный',
    'благовещенск': 'Дальневосточный',
    'комсомольск-на-амуре': 'Дальневосточный',
    'южно-сахалинск': 'Дальневосточный',
    'петропавловск-камчатский': 'Дальневосточный',
    'магадан': 'Дальневосточный',
    'якутск': 'Дальневосточный',
    'нахodka': 'Дальневосточный',
    'находка': 'Дальневосточный',
    # Приволжский ФО
    'нижний новгород': 'Приволжский',
    'казань': 'Приволжский',
    'самара': 'Приволжский',
    '.ufa': 'Приволжский',
    'уфа': 'Приволжский',
    'пермь': 'Приволжский',
    'чебоксары': 'Приволжский',
    'йошкар-ола': 'Приволжский',
    'саранск': 'Приволжский',
    'пенза': 'Приволжский',
    'ульяновск': 'Приволжский',
    'киров': 'Приволжский',
    'оренбург': 'Приволжский',
    'саратов': 'Приволжский',
    'волгоград': 'Приволжский',
    'астрахань': 'Приволжский',
    'набережные челны': 'Приволжский',
    'альметьевск': 'Приволжский',
    'тольятти': 'Приволжский',
    # Южный ФО
    'ростов-на-дону': 'Южный',
    'краснодар': 'Южный',
    'ставрополь': 'Южный',
    'сочи': 'Южный',
    'волгоград': 'Южный',
    'астрахань': 'Южный',
    'майкоп': 'Южный',
    'владикавказ': 'Южный',
    'махачкала': 'Южный',
    # Северо-Кавказский ФО
    'махачкала': 'Северо-Кавказский',
    'ставрополь': 'Северо-Кавказский',
    'владикавказ': 'Северо-Кавказский',
    'грозный': 'Северо-Кавказский',
    'нальчик': 'Северо-Кавказский',
    'пятигорск': 'Северо-Кавказский',
}

def get_federal_district(city):
    """Определить федеральный округ по городу"""
    if not city:
        return None
    city_lower = city.strip().lower()
    for key, fo in CITY_TO_FO.items():
        if key in city_lower:
            return fo
    return None


def web_search(query, num=5):
    """Выполнить веб-поиск через z-ai CLI"""
    output_file = f'/tmp/search_{int(time.time()*1000)}.json'
    try:
        result = subprocess.run(
            ['z-ai', 'function', '-n', 'web_search',
             '-a', json.dumps({"query": query, "num": num}),
             '-o', output_file],
            capture_output=True, text=True, timeout=30
        )
        if os.path.exists(output_file):
            with open(output_file, 'r') as f:
                data = json.load(f)
            os.remove(output_file)
            return data
    except Exception as e:
        print(f"  Ошибка поиска: {e}")
    return []


def read_page(url):
    """Прочитать веб-страницу через z-ai CLI"""
    output_file = f'/tmp/page_{int(time.time()*1000)}.json'
    try:
        result = subprocess.run(
            ['z-ai', 'function', '-n', 'page_reader',
             '-a', json.dumps({"url": url}),
             '-o', output_file],
            capture_output=True, text=True, timeout=45
        )
        if os.path.exists(output_file):
            with open(output_file, 'r') as f:
                data = json.load(f)
            os.remove(output_file)
            if isinstance(data, dict) and 'data' in data:
                return data['data']
            return data
    except Exception as e:
        print(f"  Ошибка чтения страницы: {e}")
    return None


def extract_city_from_text(text):
    """Извлечь город из текста"""
    if not text:
        return None
    
    # Ищем паттерн "г. Название" или "город Название"
    patterns = [
        r'г\.\s*([А-ЯЁ][а-яё]+(?:-[а-яё]+)*)',
        r'город\s+([А-ЯЁ][а-яё]+(?:-[а-яё]+)*)',
        r'г\s+([А-ЯЁ][а-яё]+(?:-[а-яё]+)*)',
    ]
    for p in patterns:
        m = re.search(p, text)
        if m:
            return m.group(1)
    
    # Проверяем известные города
    known_cities = [
        'Санкт-Петербург', 'Москва', 'Новосибирск', 'Екатеринбург',
        'Красноярск', 'Иркутск', 'Владивосток', 'Хабаровск',
        'Нижний Новгород', 'Казань', 'Самара', 'Челябинск',
        'Омск', 'Ростов-на-Дону', 'Уфа', 'Волгоград',
        'Пермь', 'Краснодар', 'Воронеж', 'Тюмень',
        'Сыктывкар', 'Ярославль', 'Рязань', 'Барнаул',
        'Кемерово', 'Томск', 'Норильск', 'Киров',
        'Йошкар-Ола', 'Благовещенск', 'Мурманск',
        'Петрозаводск', 'Калининград', 'Псков',
        'Архангельск', 'Вологда', 'Череповец',
        'Владимир', 'Иваново', 'Кострома', 'Тверь',
        'Тула', 'Брянск', 'Смоленск', 'Калуга',
        'Орёл', 'Курск', 'Липецк', 'Белгород',
        'Тамбов', 'Сургут', 'Магнитогорск',
        'Подольск', 'Красногорск', 'Котлас',
        'Набережные Челны', 'Тольятти', 'Альметьевск',
    ]
    text_lower = text.lower()
    for city in known_cities:
        if city.lower() in text_lower:
            return city
    
    return None


def extract_address_from_text(text):
    """Извлечь почтовый адрес из текста"""
    if not text:
        return None
    
    # Ищем паттерн с индексом
    patterns = [
        r'(\d{6},?\s*[^,]+(?:г\.|город)[^,]*(?:,?[^,]*)*)',
        r'(\d{6}[^<]{10,150})',
        r'(г\.[^,]+(?:,?[^,]+){1,5})',
        r'(?:адрес[у:]?\s*:?\s*)(.{20,200})',
        r'(?:Юр\.?\s*адрес[у:]?\s*:?\s*)(.{20,200})',
        r'(?:Юридический\s+адрес[у:]?\s*:?\s*)(.{20,200})',
    ]
    
    for p in patterns:
        m = re.search(p, text, re.IGNORECASE)
        if m:
            addr = m.group(1).strip()
            # Очистка
            addr = re.sub(r'\s+', ' ', addr)
            addr = addr.rstrip('.,; ')
            if len(addr) > 10:
                return addr
    
    return None


def extract_email_from_text(text):
    """Извлечь email из текста"""
    if not text:
        return None
    pattern = r'[\w\.-]+@[\w\.-]+\.\w{2,10}'
    matches = re.findall(pattern, text)
    # Фильтруем технические email
    for email in matches:
        lower = email.lower()
        if any(skip in lower for skip in ['example', 'test', 'email]', 'yoursite', 'domain']):
            continue
        return email
    return matches[0] if matches else None


def extract_website_from_results(results, company_name):
    """Извлечь сайт компании из результатов поиска"""
    if not results:
        return None
    
    # Список доменов-справочников, которые нам не нужны
    skip_domains = [
        'rusprofile.ru', 'list-org.com', 'spark-interfax.ru',
        'companies.rbc.ru', 'zachestnyibiznes.ru', 'sbis.ru',
        'org-info.com', 'www.org-info.com', 'venyoo.ru',
        'find-org.com', 'orglist.ru', 'regfile.ru',
        'ruscompany.com', 'inc-legal.ru', 'buhschetchet.ru',
        'e-grfn.ru', 'taxcom.ru', 'kontragent.skbkontur.ru',
        'focus.kontur.ru', 'ul.company', 'glayl.com',
    ]
    
    for r in results:
        url = r.get('url', '')
        host = r.get('host_name', '')
        
        # Пропускаем справочники
        if any(d in host for d in skip_domains):
            continue
        
        # Ищем URL, который выглядит как сайт компании
        if url.startswith('http'):
            # Убираем пути вроде /contacts, /about и т.д.
            domain_match = re.match(r'https?://(?:www\.)?([^/]+)', url)
            if domain_match:
                domain = domain_match.group(1)
                # Проверяем что это не поисковик/справочник
                if not any(d in domain for d in ['google', 'yandex', 'bing', '2gis', 'avito']):
                    return f"https://{domain}"
    
    return None


def search_company_info(company_name, inn):
    """Поиск информации о компании через интернет"""
    info = {
        'federal_district': None,
        'city': None,
        'email': None,
        'website': None,
        'postal_address': None,
    }
    
    # Формируем поисковый запрос
    query = f"ИНН {inn}"
    print(f"  Поиск: {query}")
    
    results = web_search(query, num=8)
    if not results:
        # Пробуем по названию
        query = f'"{company_name}" сайт контакты адрес'
        print(f"  Повторный поиск: {query}")
        results = web_search(query, num=5)
    
    if not results:
        return info
    
    # Собираем весь текст из сниппетов
    all_text = ' '.join([r.get('snippet', '') + ' ' + r.get('name', '') for r in results])
    
    # Извлекаем город
    city = extract_city_from_text(all_text)
    if city:
        info['city'] = city
        fo = get_federal_district(city)
        if fo:
            info['federal_district'] = fo
    
    # Извлекаем адрес из сниппетов
    address = extract_address_from_text(all_text)
    if address:
        info['postal_address'] = address
    
    # Извлекаем сайт
    website = extract_website_from_results(results, company_name)
    if website:
        info['website'] = website
    
    # Если email не найден, пробуем найти сайт компании и прочитать контакты
    # Сначала проверяем сниппеты
    email = extract_email_from_text(all_text)
    if email:
        info['email'] = email
    
    # Если нет email или адреса, пытаемся прочитать Rusprofile
    rusprofile_url = None
    for r in results:
        if 'rusprofile.ru' in r.get('url', ''):
            rusprofile_url = r.get('url')
            break
    
    if rusprofile_url and (not info['email'] or not info['postal_address'] or not info['city']):
        print(f"  Чтение Rusprofile: {rusprofile_url}")
        page_data = read_page(rusprofile_url)
        if page_data:
            html = page_data.get('html', '') or page_data.get('text', '') or ''
            if html:
                # Убираем HTML-теги
                text = re.sub(r'<[^>]+>', ' ', html)
                text = re.sub(r'\s+', ' ', text)
                
                if not info['email']:
                    info['email'] = extract_email_from_text(text)
                
                if not info['postal_address']:
                    addr = extract_address_from_text(text)
                    if addr:
                        info['postal_address'] = addr
                
                if not info['city']:
                    city = extract_city_from_text(text)
                    if city:
                        info['city'] = city
                        fo = get_federal_district(city)
                        if fo:
                            info['federal_district'] = fo
    
    # Если нет email, пробуем сайт компании
    if not info['email'] and info['website']:
        contact_url = info['website'].rstrip('/') + '/contacts'
        print(f"  Чтение контактов: {contact_url}")
        page_data = read_page(contact_url)
        if page_data:
            html = page_data.get('html', '') or page_data.get('text', '') or ''
            text = re.sub(r'<[^>]+>', ' ', html)
            text = re.sub(r'\s+', ' ', text)
            email = extract_email_from_text(text)
            if email:
                info['email'] = email
    
    return info


def main():
    src_file = '/home/z/my-project/upload/Задание на дополнение таблицы белые пятна.xlsx'
    dst_file = '/home/z/my-project/download/Задание_на_дополнение_таблицы_заполнено.xlsx'
    
    # Загружаем файл
    wb = openpyxl.load_workbook(src_file)
    ws = wb['Лист1']
    
    # Строки 3-57 содержат данные о партнёрах
    companies = []
    for row_num in range(3, ws.max_row + 1):
        client = ws.cell(row=row_num, column=4).value  # D - Клиент
        inn = ws.cell(row=row_num, column=5).value  # E - ИНН
        existing_city = ws.cell(row=row_num, column=3).value  # C - Город
        existing_fo = ws.cell(row=row_num, column=2).value  # B - ФО
        
        if client and inn:
            companies.append({
                'row': row_num,
                'client': str(client).strip(),
                'inn': str(inn).strip(),
                'existing_city': str(existing_city).strip() if existing_city else None,
                'existing_fo': str(existing_fo).strip() if existing_fo else None,
            })
    
    print(f"Найдено {len(companies)} компаний для обработки\n")
    
    # Обрабатываем каждую компанию
    for i, comp in enumerate(companies):
        print(f"\n[{i+1}/{len(companies)}] {comp['client']} (ИНН: {comp['inn']})")
        
        # Ищем информацию
        info = search_company_info(comp['client'], comp['inn'])
        
        # Заполняем данные
        row = comp['row']
        
        # Федеральный округ (столбец B)
        if info['federal_district'] and not comp['existing_fo']:
            ws.cell(row=row, column=2, value=info['federal_district'])
            print(f"  ✅ ФО: {info['federal_district']}")
        elif not info['federal_district'] and not comp['existing_fo']:
            print(f"  ❌ ФО: не найден")
        
        # Город (столбец C)
        if info['city'] and not comp['existing_city']:
            ws.cell(row=row, column=3, value=info['city'])
            print(f"  ✅ Город: {info['city']}")
        elif not info['city'] and not comp['existing_city']:
            print(f"  ❌ Город: не найден")
        
        # Email (столбец M)
        if info['email']:
            ws.cell(row=row, column=13, value=info['email'])
            print(f"  ✅ Email: {info['email']}")
        else:
            print(f"  ❌ Email: не найден")
        
        # Сайт (столбец O)
        if info['website']:
            ws.cell(row=row, column=15, value=info['website'])
            print(f"  ✅ Сайт: {info['website']}")
        else:
            print(f"  ❌ Сайт: не найден")
        
        # Почтовый адрес (столбец P)
        if info['postal_address']:
            ws.cell(row=row, column=16, value=info['postal_address'])
            print(f"  ✅ Адрес: {info['postal_address'][:80]}...")
        else:
            print(f"  ❌ Адрес: не найден")
        
        # Небольшая пауза между запросами
        time.sleep(1)
    
    # Сохраняем
    wb.save(dst_file)
    print(f"\n\n✅ Файл сохранён: {dst_file}")
    
    # Статистика
    filled = {'fo': 0, 'city': 0, 'email': 0, 'website': 0, 'address': 0}
    total = len(companies)
    for row_num in range(3, ws.max_row + 1):
        if ws.cell(row=row_num, column=2).value: filled['fo'] += 1
        if ws.cell(row=row_num, column=3).value: filled['city'] += 1
        if ws.cell(row=row_num, column=13).value: filled['email'] += 1
        if ws.cell(row=row_num, column=15).value: filled['website'] += 1
        if ws.cell(row=row_num, column=16).value: filled['address'] += 1
    
    print(f"\n📊 Статистика заполнения:")
    print(f"  Федеральный округ: {filled['fo']}/{total}")
    print(f"  Город: {filled['city']}/{total}")
    print(f"  Email: {filled['email']}/{total}")
    print(f"  Сайт: {filled['website']}/{total}")
    print(f"  Почтовый адрес: {filled['address']}/{total}")


if __name__ == '__main__':
    main()
