#!/usr/bin/env python3
"""
Волна 4: Исправление сайтов (убрать справочники) + поиск email через контактные страницы.
"""

import openpyxl
import json
import re
import os
import subprocess
import time

dst_file = '/home/z/my-project/download/Задание_на_дополнение_таблицы_заполнено.xlsx'

# Справочники и нерелевантные сайты, которые нужно убрать
BAD_WEBSITES = [
    'audit-it.ru', 'checko.ru', 'saby.ru', 'synapsenet.ru',
    'znaemvseh.ru', 'tbank.ru', 'service-online.su', 'agro.gov74.ru',
    'reputation.ru', 'cataloxy.ru', 't.me', 'tapki.com',
    'energybase.ru', 'tenderer.ru',
]

# Исправленные сайты (ИНН → правильный сайт)
FIXED_WEBSITES = {
    '4': None,  # Дальтех МБ - не удалось найти сайт, tbank.ru убираем
    '6': None,  # Жадаев ИП - нет сайта
    '8': None,  # Сибирская Усадьба - нет сайта
    '10': None,  # Студия красок - исключена из ЕГРЮЛ
    '17': None,  # Вланж плюс - нет сайта
    '20': None,  # ИП Ковалев - нет сайта
    '21': None,  # ИП Мима - нет сайта
    '22': None,  # ИП Минюк - нет сайта
    '23': None,  # ИП Миронкин - нет сайта
    '24': None,  # ИП Сучков - нет сайта
    '25': None,  # ИП Чарыков - нет сайта
    '26': None,  # ИП Шишков - нет сайта
    '30': None,  # Молоток - нет сайта
    '31': None,  # Сибирь Дистрибуция - закрылись
    '32': None,  # Сигма - нет сайта
    '34': None,  # СТРОЙТЕ С НАМИ - нет сайта
    '35': None,  # Хозстройкомплект - нет сайта
    '43': None,  # Левша - нет сайта
    '44': None,  # Стройградторг - нет сайта
    '50': None,  # Реал-Маркет - ликвидирован
    '51': None,  # ИП Сергеев - нет сайта
    '54': None,  # ИП Воронков - нет сайта
    '56': None,  # Уютстрой - нет сайта
}

# Сайты компаний для чтения контактов (ИНН → URL контактов)
CONTACT_PAGES = {
    '7811340145': 'https://alt-x.ru/contacts',
    '5902816678': 'https://decorativca.ru/contacts',
    '5003123271': 'https://dommalera.ru/contacts',
    '9717052425': 'https://msk.saturn.net/contacts',
    '4401025465': 'https://akson.pro/contacts',
    '3827026836': 'https://baumarket.ru/contacts',
    '7827001116': 'https://vimos.ru/contacts',
    '6670179530': 'https://gipermarketdom.ru/contacts',
    '7731564157': 'https://art-industria.ru/contacts',
    '7453271996': 'https://stroybaza1.ru/contacts',
    '6678102667': 'https://brozexmarket.ru/contacts',
    '5906062502': 'https://batkomplekt.ru/contacts',
    '5904341998': 'https://baumoll.ru/contacts',
    '1101075340': 'https://benelux-market.ru/contacts',
    '3731030101': 'https://kenguru-ivanovo.ru/contacts',
    '7017379679': 'https://stroyasya.ru/contacts',
    '7733251960': 'https://1001kraska.ru/contacts',
    '2462068624': 'https://ochag-home.ru/contacts',
    '3662061709': 'https://protek-vrn.ru/contacts',
    '7728115183': 'https://tvoedom.ru/contacts',
    '2309181028': 'https://fluegger.ru/contacts',
    '2511036554': 'https://дальторгсервис.рф/contacts',
    '5752045215': 'https://hss.ru/contacts',
}


def extract_email_from_text(text):
    """Извлечь email из текста"""
    if not text:
        return None
    pattern = r'[\w\.-]+@[\w\.-]+\.\w{2,10}'
    matches = re.findall(pattern, text)
    for email in matches:
        lower = email.lower()
        if any(skip in lower for skip in ['example', 'test', 'sentry', 'wixpress', 'gravatar', 'github', 'cdn', 'cloudflare']):
            continue
        return email
    return matches[0] if matches else None


def main():
    wb = openpyxl.load_workbook(dst_file)
    ws = wb['Лист1']
    
    # Шаг 1: Убираем нерелевантные сайты
    print("=== Шаг 1: Удаление нерелевантных сайтов ===")
    removed = 0
    for row_num in range(3, ws.max_row + 1):
        website = ws.cell(row=row_num, column=15).value
        if website:
            website_str = str(website).lower()
            for bad in BAD_WEBSITES:
                if bad in website_str:
                    ws.cell(row=row_num, column=15, value=None)
                    client = ws.cell(row=row_num, column=4).value
                    print(f"  ❌ Убран {bad}: Row {row_num} ({str(client)[:30]})")
                    removed += 1
                    break
    print(f"  Убрано нерелевантных сайтов: {removed}")
    
    # Шаг 2: Чтение контактных страниц для email
    print("\n=== Шаг 2: Поиск email через контактные страницы ===")
    emails_found = 0
    
    for inn, contact_url in CONTACT_PAGES.items():
        # Находим строку по ИНН
        target_row = None
        for row_num in range(3, ws.max_row + 1):
            inn_val = ws.cell(row=row_num, column=5).value
            if inn_val and str(inn_val).strip() == inn:
                # Проверяем, есть ли уже email
                if ws.cell(row=row_num, column=13).value:
                    target_row = None  # Уже есть email, пропускаем
                    break
                target_row = row_num
                break
        
        if target_row is None:
            continue
        
        client = ws.cell(row=target_row, column=4).value
        print(f"  Чтение: {contact_url} ({str(client)[:30]})")
        
        output_file = f'/tmp/contact_{inn}.json'
        try:
            result = subprocess.run(
                ['z-ai', 'function', '-n', 'page_reader',
                 '-a', json.dumps({"url": contact_url}),
                 '-o', output_file],
                capture_output=True, text=True, timeout=45
            )
            
            if os.path.exists(output_file):
                with open(output_file, 'r') as f:
                    data = json.load(f)
                os.remove(output_file)
                
                # Извлекаем текст
                html = ''
                if isinstance(data, dict):
                    html = data.get('html', '') or data.get('text', '') or ''
                    if 'data' in data and isinstance(data['data'], dict):
                        html = data['data'].get('html', '') or data['data'].get('text', '') or ''
                
                text = re.sub(r'<[^>]+>', ' ', html)
                text = re.sub(r'\s+', ' ', text)
                
                email = extract_email_from_text(text)
                if email:
                    ws.cell(row=target_row, column=13, value=email)
                    print(f"    ✅ Email: {email}")
                    emails_found += 1
                else:
                    # Пробуем также поиск по /kontakty или /about
                    print(f"    ❌ Email не найден на странице")
            else:
                print(f"    ❌ Страница не загружена")
        except Exception as e:
            print(f"    ❌ Ошибка: {e}")
        
        time.sleep(2)  # Пауза для rate limiting
    
    # Сохраняем
    wb.save(dst_file)
    
    # Итоговая статистика
    print(f"\n{'='*60}")
    print(f"✅ Файл сохранён: {dst_file}")
    print(f"\n📊 Результаты:")
    print(f"  Убрано нерелевантных сайтов: {removed}")
    print(f"  Найдено email: {emails_found}")
    
    # Финальная сводка
    filled = {'fo': 0, 'city': 0, 'email': 0, 'website': 0, 'address': 0}
    total = 0
    for row_num in range(3, ws.max_row + 1):
        if ws.cell(row=row_num, column=4).value:
            total += 1
            if ws.cell(row=row_num, column=2).value: filled['fo'] += 1
            if ws.cell(row=row_num, column=3).value: filled['city'] += 1
            if ws.cell(row=row_num, column=13).value: filled['email'] += 1
            if ws.cell(row=row_num, column=15).value: filled['website'] += 1
            if ws.cell(row=row_num, column=16).value: filled['address'] += 1
    
    print(f"\n📋 ИТОГОВАЯ СВОДКА ({total} компаний):")
    print(f"  Федеральный округ: {filled['fo']}/{total} ({filled['fo']*100//total}%)")
    print(f"  Город: {filled['city']}/{total} ({filled['city']*100//total}%)")
    print(f"  Email: {filled['email']}/{total} ({filled['email']*100//total}%)")
    print(f"  Сайт: {filled['website']}/{total} ({filled['website']*100//total}%)")
    print(f"  Почтовый адрес: {filled['address']}/{total} ({filled['address']*100//total}%)")


if __name__ == '__main__':
    main()
