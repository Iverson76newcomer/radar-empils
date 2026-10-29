#!/usr/bin/env python3
"""
Волна 3: Очистка и корректировка данных.
Исправление городов, сайтов, адресов, email.
"""

import openpyxl
import json
import re
import os
import subprocess
import time

SEARCH_DIR = '/home/z/my-project/scripts/search_results'
src_file = '/home/z/my-project/download/Задание_на_дополнение_таблицы_заполнено.xlsx'
dst_file = '/home/z/my-project/download/Задание_на_дополнение_таблицы_заполнено.xlsx'

# Справочник: ИНН → правильные данные (собранные из результатов поиска)
MANUAL_FIXES = {
    # Row 3: Альт-Икс
    '7811340145': {
        'city': 'Санкт-Петербург',
        'fo': 'Северо-Западный',
        'website': 'https://alt-x.ru',
        'address': '194292, г. Санкт-Петербург, 1-й Верхний пер., д. 10, к. 2, стр. 1, офис 3.4',
    },
    # Row 4: Дальтех МБ / CarSan
    '2537090274': {
        'city': 'Владивосток',
        'fo': 'Дальневосточный',
        'address': 'Приморский край, г. Владивосток, ул. 50 лет ВЛКСМ, д. 26',
    },
    # Row 5: Декор-Центр
    '5902816678': {
        'city': 'Пермь',
        'fo': 'Приволжский',
        'website': 'https://decorativca.ru',
    },
    # Row 6: Жадаев / Колор студия / КЛК - ИП в Красноярском крае
    '246600161805': {
        'city': 'Красноярск',
        'fo': 'Сибирский',
        'website': None,  # нет сайта
        'address': None,  # ИП — нет юрадреса
    },
    # Row 7: Красхимресурс
    '2465056419': {
        'city': 'Красноярск',
        'fo': 'Сибирский',
        'website': 'https://mirkrasok24.ru',
        'email': 'zakaz@mirkrasok24.ru',
        'address': '660077, Красноярский край, г. Красноярск, ул. Молокова, д. 40, помещ. 330',
    },
    # Row 8: Сибирская Усадьба
    '5406296559': {
        'city': 'Новосибирск',
        'fo': 'Сибирский',
        'website': None,
    },
    # Row 9: Склад красок / Дом Маляра
    '5003123271': {
        'city': 'Люберцы',
        'fo': 'Центральный',
        'website': 'https://dommalera.ru',
        'address': '140060, Московская обл., Люберецкий р-н, г. Люберцы',
    },
    # Row 10: Студия красок - исключена из ЕГРЮЛ, была в СПб
    '7806444830': {
        'city': 'Санкт-Петербург',
        'fo': 'Северо-Западный',
        'website': None,
    },
    # Row 11: Химснабсервис
    '5752045215': {
        'city': 'Орёл',
        'fo': 'Центральный',
        'website': 'https://hss.ru',
        'email': 'info@hss.ru',
    },
    # Row 12: САТУРН СТРОЙМАРКЕТ ЦЕНТР
    '9717052425': {
        'city': 'Санкт-Петербург',
        'fo': 'Северо-Западный',  # Главный офис в СПб
        'website': 'https://msk.saturn.net',
    },
    # Row 13: Аксон
    '4401025465': {
        'city': 'Кострома',
        'fo': 'Центральный',
        'website': 'https://akson.pro',
    },
    # Row 14: Апельсин
    '622700326226': {
        'city': 'Рязань',
        'fo': 'Центральный',
        'website': 'https://apelsin.ru',
        'email': 'mail@apelsin.ru',
    },
    # Row 15: Баумаркет
    '3827026836': {
        'city': 'Иркутск',
        'fo': 'Сибирский',
        'website': 'https://baumarket.ru',
    },
    # Row 16: Вимос
    '7827001116': {
        'city': 'Всеволожск',
        'fo': 'Северо-Западный',
        'website': 'https://vimos.ru',
    },
    # Row 17: Вланж плюс / Дедал
    '2465254410': {
        'city': 'Красноярск',
        'fo': 'Сибирский',
        'website': None,
        'address': '660004, Красноярский край, г. Красноярск, пр-кт Им.Газеты "Красноярский Рабочий", д. 27, офис 43',
    },
    # Row 18: Дальторгсервис
    '2511036554': {
        'city': 'Хабаровск',
        'fo': 'Дальневосточный',
        'website': 'https://дальторгсервис.рф',
    },
    # Row 19: Домовой / УК СТАРТ
    '7816380127': {
        'city': 'Санкт-Петербург',
        'fo': 'Северо-Западный',
        'website': 'https://startonline.ru',
        'email': 'info@startonline.ru',
        'address': '192283, г. Санкт-Петербург, Бухарестская ул., д. 89, литера А, помещ. 12н',
    },
    # Row 20: ИП Ковалев
    '550604212050': {
        'city': 'Омск',
        'fo': 'Сибирский',
        'website': None,
    },
    # Row 21: ИП Мима
    '245700918501': {
        'city': 'Норильск',
        'fo': 'Сибирский',
        'website': None,
    },
    # Row 22: ИП Минюк / Оранж
    '760400191237': {
        'city': 'Ярославль',
        'fo': 'Центральный',
        'website': None,
        'address': '150003, г. Ярославль, Кооперативная ул., 11',
    },
    # Row 23: ИП Миронкин
    '550400093929': {
        'city': 'Омск',
        'fo': 'Сибирский',
        'website': None,
        'address': '646710, Омская обл., Шербакульский р-н, с. Борисовское',
    },
    # Row 24: ИП Сучков / Очаг
    '246200497430': {
        'city': 'Красноярск',
        'fo': 'Сибирский',
        'website': None,
    },
    # Row 25: ИП Чарыков / Ваш Дом
    '246500188627': {
        'city': 'Красноярск',
        'fo': 'Сибирский',
        'website': None,
    },
    # Row 26: ИП Шишков
    '550405862624': {
        'city': 'Омск',
        'fo': 'Сибирский',
        'website': None,
    },
    # Row 27: ДОМ КорТрейд
    '6670179530': {
        'city': 'Екатеринбург',
        'fo': 'Уральский',
        'website': 'https://gipermarketdom.ru',
        'address': '620014, Свердловская обл., г. Екатеринбург, ул. 8 Марта, д. 179',
    },
    # Row 28: ДОМ Сити строй
    '6673163835': {
        'city': 'Екатеринбург',
        'fo': 'Уральский',
        'website': 'https://gipermarketdom.ru',
        'address': '620014, Свердловская обл., г. Екатеринбург, ул. 8 Марта, д. 179',
    },
    # Row 29: АРТ индустрия
    '7731564157': {
        'city': 'Москва',
        'fo': 'Центральный',
        'website': 'https://art-industria.ru',
        'address': '121354, г. Москва, ул. Дорогобужская, 14 / стр. 7, помещ. 3',
    },
    # Row 30: Молоток
    '5506230390': {
        'city': 'Омск',
        'fo': 'Сибирский',
        'website': None,
        'address': '644007, г. Омск, ул. учебная, д. 81',
    },
    # Row 31: Сибирь Дистрибуция (закрылись)
    '7203520720': {
        'city': 'Тюмень',
        'fo': 'Уральский',
        'website': None,
        'address': '625002, Тюменская обл., г. Тюмень, Профсоюзная ул., д. 30, кв. 179',
    },
    # Row 32: Сигма
    '7716989966': {
        'city': 'Москва',
        'fo': 'Центральный',
        'website': None,
        'address': '129327, г. Москва, ул. Лазо, д. 9, помещ. 3/2',
    },
    # Row 33: Стройбаза №1
    '7453271996': {
        'city': 'Челябинск',
        'fo': 'Уральский',
        'website': 'https://stroybaza1.ru',
    },
    # Row 34: СТРОЙТЕ С НАМИ
    '6625044348': {
        'city': 'Екатеринбург',
        'fo': 'Уральский',
        'website': None,
    },
    # Row 35: Хозстройкомплект
    '7452095917': {
        'city': 'Челябинск',
        'fo': 'Уральский',
        'website': None,
        'address': '454085, Челябинская обл., г. Челябинск, Шуменская ул., д. 31, 274',
    },
    # Row 36: РИТЕЙЛ / Домострой / Brozexmarket
    '6678102667': {
        'city': 'Берёзовский',
        'fo': 'Уральский',
        'website': 'https://brozexmarket.ru',
        'address': '623701, Свердловская обл., г. Берёзовский, ул. Гагарина, д. 6',
    },
    # Row 37: Аксиома
    '2464236507': {
        'city': 'Красноярск',
        'fo': 'Сибирский',
        'website': 'https://aksioma-krsk.ru',
        'email': 'axioma@stp.su',
        'address': '660064, Красноярский край, г. Красноярск, ул. Академика Вавилова, зд. 1г/1',
    },
    # Row 38: Баткомплект
    '5906062502': {
        'city': 'Пермь',
        'fo': 'Приволжский',
        'website': 'https://batkomplekt.ru',
        'address': '614014, Пермский край, г. Пермь',
    },
    # Row 39: Баумолл - Трейд
    '5904341998': {
        'city': 'Пермь',
        'fo': 'Приволжский',
        'website': 'https://baumoll.ru',
        'address': '614025, Пермский край, г. Пермь, ул. Героев Хасана, зд. 109, каб. 218',
    },
    # Row 40: Бенелюкс Сыктывкар
    '1101075340': {
        'city': 'Сыктывкар',
        'fo': 'Северо-Западный',
        'website': 'https://benelux-market.ru',
    },
    # Row 41: Кенгуру
    '3731030101': {
        'city': 'Иваново',
        'fo': 'Центральный',
        'website': 'https://kenguru-ivanovo.ru',
        'address': '153011, Ивановская обл., г. Иваново, ул. Красных Зорь, д. 34',
    },
    # Row 42: Комплекс / Стройся
    '7017379679': {
        'city': 'Томск',
        'fo': 'Сибирский',
        'website': 'https://stroyasya.ru',
    },
    # Row 43: Левша
    '6623119164': {
        'city': 'Нижний Тагил',
        'fo': 'Уральский',
        'website': None,
        'address': '622005, Свердловская обл., г. Нижний Тагил, ул. Индустриальная, д. 28',
    },
    # Row 44: Стройградторг / Стройград Норильск
    '2457079148': {
        'city': 'Норильск',
        'fo': 'Сибирский',
        'website': None,
    },
    # Row 45: 1001 КРАСКА
    '7733251960': {
        'city': 'Москва',
        'fo': 'Центральный',
        'website': 'https://1001kraska.ru',
    },
    # Row 46: ЭПФ Пилон
    '2464002594': {
        'city': 'Красноярск',
        'fo': 'Сибирский',
        'website': 'https://pilon24.ru',
    },
    # Row 47: Очаг Хоум
    '2462068624': {
        'city': 'Красноярск',
        'fo': 'Сибирский',
        'website': 'https://ochag-home.ru',
        'address': '660013, Красноярский край, г. Красноярск, Томская ул., зд. 4, к. 2, стр. 33',
    },
    # Row 48: Порядок / Спектр
    '3663157837': {
        'city': 'Воронеж',
        'fo': 'Центральный',
        'website': 'https://poryadok.ru',
        'email': 'spektropt1@yandex.ru',
        'address': '394028, г. Воронеж',
    },
    # Row 49: Протэк
    '3662061709': {
        'city': 'Воронеж',
        'fo': 'Центральный',
        'website': 'https://protek-vrn.ru',
    },
    # Row 50: Реал-Маркет / Стар Трейд (Котлас)
    '2904028638': {
        'city': 'Котлас',
        'fo': 'Северо-Западный',
        'website': None,
        'address': '165300, Архангельская обл., г. Котлас',
    },
    # Row 51: Сергеев / Лаки Краски
    '3442117755': {
        'city': 'Волгоград',
        'fo': 'Южный',
        'website': None,
    },
    # Row 52: Стройгигант / Курс-Логистика
    '4632110238': {
        'city': 'Курск',
        'fo': 'Центральный',
        'website': 'https://stroygigant.ru',
        'email': 'n.guseva@stroygigant.ru',
        'address': '305038, Курская обл., г. Курск, ул. Косухина, д. 51',
    },
    # Row 53: СтройДепо / НЭКС
    '7714996975': {
        'city': 'Москва',
        'fo': 'Центральный',
        'website': 'https://stroydepo.ru',
        'email': 'ooo-neks@rambler.ru',
        'address': '119192, г. Москва, ул. Мосфильмовская, 74Б, офис 219',
    },
    # Row 54: Стройсервис АСМ / ИП Воронков
    '665200210601': {
        'city': 'Арамиль',
        'fo': 'Уральский',
        'website': None,
    },
    # Row 55: Твой дом / Крокус
    '7728115183': {
        'city': 'Красногорск',
        'fo': 'Центральный',
        'website': 'https://tvoedom.ru',
        'address': '143403, Московская обл., г. Красногорск',
    },
    # Row 56: Уютстрой / Стройцентр
    '9102282914': {
        'city': 'Симферополь',
        'fo': 'Южный',  # Крым — Южный ФО
        'website': None,
    },
    # Row 57: Хорошиекраски / FLUGGER
    '2309181028': {
        'city': 'Краснодар',
        'fo': 'Южный',
        'website': 'https://fluegger.ru',
    },
}


def web_search(query, num=5):
    """Веб-поиск через z-ai CLI"""
    output_file = f'/tmp/fix_search_{int(time.time()*1000)}.json'
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
    except:
        pass
    return None


def extract_email_from_text(text):
    """Извлечь email из текста"""
    if not text:
        return None
    pattern = r'[\w\.-]+@[\w\.-]+\.\w{2,10}'
    matches = re.findall(pattern, text)
    for email in matches:
        lower = email.lower()
        if any(skip in lower for skip in ['example', 'test', 'sentry', 'wixpress', 'gravatar', 'github']):
            continue
        return email
    return matches[0] if matches else None


def main():
    wb = openpyxl.load_workbook(src_file)
    ws = wb['Лист1']
    
    stats = {'fo': 0, 'city': 0, 'email': 0, 'website': 0, 'address': 0, 'fixed': 0}
    
    for row_num in range(3, ws.max_row + 1):
        client = ws.cell(row=row_num, column=4).value
        inn_val = ws.cell(row=row_num, column=5).value
        if not client or not inn_val:
            continue
        
        inn_str = str(inn_val).strip()
        fixes = MANUAL_FIXES.get(inn_str, {})
        if not fixes:
            continue
        
        changed = False
        
        # Федеральный округ (B)
        current_fo = ws.cell(row=row_num, column=2).value
        if fixes.get('fo') and (not current_fo or current_fo != fixes['fo']):
            ws.cell(row=row_num, column=2, value=fixes['fo'])
            stats['fo'] += 1
            changed = True
        
        # Город (C)
        current_city = ws.cell(row=row_num, column=3).value
        if fixes.get('city') and (not current_city or current_city != fixes['city']):
            ws.cell(row=row_num, column=3, value=fixes['city'])
            stats['city'] += 1
            changed = True
        
        # Email (M)
        current_email = ws.cell(row=row_num, column=13).value
        if fixes.get('email') and not current_email:
            ws.cell(row=row_num, column=13, value=fixes['email'])
            stats['email'] += 1
            changed = True
        
        # Сайт (O)
        current_website = ws.cell(row=row_num, column=15).value
        fix_website = fixes.get('website')
        if fix_website is not None:
            if not current_website or current_website != fix_website:
                ws.cell(row=row_num, column=15, value=fix_website if fix_website else None)
                stats['website'] += 1
                changed = True
        
        # Адрес (P)
        current_address = ws.cell(row=row_num, column=16).value
        if fixes.get('address') and (not current_address or current_address != fixes['address']):
            ws.cell(row=row_num, column=16, value=fixes['address'])
            stats['address'] += 1
            changed = True
        
        if changed:
            stats['fixed'] += 1
            print(f"✅ Row {row_num}: {str(client)[:40]} — исправлено")
    
    # Сохраняем
    wb.save(dst_file)
    
    # Итоговая статистика
    print(f"\n{'='*60}")
    print(f"✅ Файл сохранён: {dst_file}")
    print(f"\n📊 Исправлено:")
    print(f"  Федеральный округ: {stats['fo']}")
    print(f"  Город: {stats['city']}")
    print(f"  Email: {stats['email']}")
    print(f"  Сайт: {stats['website']}")
    print(f"  Почтовый адрес: {stats['address']}")
    print(f"  Компаний с изменениями: {stats['fixed']}")
    
    # Финальная сводка
    print(f"\n📋 ИТОГОВАЯ СВОДКА:")
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
    
    print(f"  Федеральный округ: {filled['fo']}/{total}")
    print(f"  Город: {filled['city']}/{total}")
    print(f"  Email: {filled['email']}/{total}")
    print(f"  Сайт: {filled['website']}/{total}")
    print(f"  Почтовый адрес: {filled['address']}/{total}")


if __name__ == '__main__':
    main()
