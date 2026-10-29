#!/usr/bin/env python3
"""
ФИНАЛЬНАЯ ОЧИСТКА — полная корректировка всех данных.
"""

import openpyxl

dst_file = '/home/z/my-project/download/Задание_на_дополнение_таблицы_заполнено.xlsx'

wb = openpyxl.load_workbook(dst_file)
ws = wb['Лист1']

# === ПОЛНАЯ ТАБЛИЦА КОРРЕКТИРОВОК ===
# Row: {fo, city, email, website, address}
CORRECTIONS = {
    3: {  # Альт-Икс
        'fo': 'Северо-Западный',
        'city': 'Санкт-Петербург',
        'email': 'opt@alt-x.ru',
        'website': 'https://alt-x.ru',
        'address': '194292, г. Санкт-Петербург, 1-й Верхний пер., д. 10, к. 2, стр. 1, офис 3.4',
    },
    4: {  # Дальтех МБ / CarSan
        'fo': 'Дальневосточный',
        'city': 'Владивосток',
        'email': None,
        'website': None,
        'address': '690091, Приморский край, г. Владивосток, ул. 50 лет ВЛКСМ, д. 26',
    },
    5: {  # Декор-Центр
        'fo': 'Приволжский',
        'city': 'Пермь',
        'email': 'info@decor-centre.ru',
        'website': 'https://decorativca.ru',
        'address': '614113, Пермский край, г. Пермь',
    },
    6: {  # Жадаев / Колор студия / КЛК
        'fo': 'Сибирский',
        'city': 'Красноярск',
        'email': None,
        'website': None,
        'address': None,  # ИП — нет юрадреса в открытом доступе
    },
    7: {  # Красхимресурс
        'fo': 'Сибирский',
        'city': 'Красноярск',
        'email': 'zakaz@mirkrasok24.ru',
        'website': 'https://mirkrasok24.ru',
        'address': '660077, Красноярский край, г. Красноярск, ул. Молокова, д. 40, помещ. 330',
    },
    8: {  # Сибирская Усадьба
        'fo': 'Сибирский',
        'city': 'Новосибирск',
        'email': None,
        'website': None,
        'address': '630005, Новосибирская обл., г. Новосибирск',
    },
    9: {  # Склад красок / Дом Маляра
        'fo': 'Центральный',
        'city': 'Люберцы',
        'email': 'client@dommalera.ru',
        'website': 'https://dommalera.ru',
        'address': '140060, Московская обл., г. Люберцы',
    },
    10: {  # Студия красок — исключена из ЕГРЮЛ
        'fo': 'Северо-Западный',
        'city': 'Санкт-Петербург',
        'email': None,
        'website': None,
        'address': 'г. Санкт-Петербург (ликвидирована 14.03.2025)',
    },
    11: {  # Химснабсервис
        'fo': 'Центральный',
        'city': 'Орёл',
        'email': 'info@hss.ru',
        'website': 'https://hss.ru',
        'address': '302028, Орловская обл., г. Орёл',
    },
    12: {  # САТУРН СТРОЙМАРКЕТ ЦЕНТР
        'fo': 'Северо-Западный',
        'city': 'Санкт-Петербург',
        'email': 'cs@msk.saturn.net',
        'website': 'https://msk.saturn.net',
        'address': '196128, г. Санкт-Петербург, Варшавская ул., д. 5А',
    },
    13: {  # Аксон
        'fo': 'Центральный',
        'city': 'Кострома',
        'email': 'admin@akson.pro',
        'website': 'https://akson.pro',
        'address': '156013, Костромская обл., г. Кострома',
    },
    14: {  # Апельсин
        'fo': 'Центральный',
        'city': 'Рязань',
        'email': 'mail@apelsin.ru',
        'website': 'https://apelsin.ru',
        'address': '390044, Рязанская обл., г. Рязань',
    },
    15: {  # Баумаркет
        'fo': 'Сибирский',
        'city': 'Иркутск',
        'email': None,
        'website': 'https://baumarket.ru',
        'address': '664007, Иркутская обл., г. Иркутск',
    },
    16: {  # Вимос
        'fo': 'Северо-Западный',
        'city': 'Всеволожск',
        'email': 'zakaz@vimos.ru',
        'website': 'https://vimos.ru',
        'address': '188640, Ленинградская обл., г. Всеволожск',
    },
    17: {  # Вланж плюс / Дедал
        'fo': 'Сибирский',
        'city': 'Красноярск',
        'email': None,
        'website': None,
        'address': '660004, Красноярский край, г. Красноярск, пр-кт Газеты «Красноярский Рабочий», д. 27, офис 43',
    },
    18: {  # Дальторгсервис
        'fo': 'Дальневосточный',
        'city': 'Хабаровск',
        'email': 'shop@dtservice.ru',
        'website': 'https://дальторгсервис.рф',
        'address': '680000, Хабаровский край, г. Хабаровск',
    },
    19: {  # Домовой / УК СТАРТ
        'fo': 'Северо-Западный',
        'city': 'Санкт-Петербург',
        'email': 'info@startonline.ru',
        'website': 'https://startonline.ru',
        'address': '192283, г. Санкт-Петербург, Бухарестская ул., д. 89, литера А, помещ. 12н',
    },
    20: {  # ИП Ковалев
        'fo': 'Сибирский',
        'city': 'Омск',
        'email': None,
        'website': None,
        'address': 'г. Омск',
    },
    21: {  # ИП Мима
        'fo': 'Сибирский',
        'city': 'Норильск',
        'email': None,
        'website': None,
        'address': '663300, Красноярский край, г. Норильск',
    },
    22: {  # ИП Минюк / Оранж
        'fo': 'Центральный',
        'city': 'Ярославль',
        'email': None,
        'website': None,
        'address': '150003, г. Ярославль, Кооперативная ул., д. 11',
    },
    23: {  # ИП Миронкин
        'fo': 'Сибирский',
        'city': 'Омск',
        'email': None,
        'website': None,
        'address': '646710, Омская обл., Шербакульский р-н, с. Борисовское',
    },
    24: {  # ИП Сучков / Очаг
        'fo': 'Сибирский',
        'city': 'Красноярск',
        'email': None,
        'website': None,
        'address': 'г. Красноярск',
    },
    25: {  # ИП Чарыков / Ваш Дом
        'fo': 'Сибирский',
        'city': 'Красноярск',
        'email': None,
        'website': None,
        'address': 'г. Красноярск',
    },
    26: {  # ИП Шишков
        'fo': 'Сибирский',
        'city': 'Омск',
        'email': None,
        'website': None,
        'address': 'г. Омск',
    },
    27: {  # ДОМ КорТрейд
        'fo': 'Уральский',
        'city': 'Екатеринбург',
        'email': 'info@gipermarketdom.ru',
        'website': 'https://gipermarketdom.ru',
        'address': '620014, Свердловская обл., г. Екатеринбург, ул. 8 Марта, д. 179',
    },
    28: {  # ДОМ Сити строй
        'fo': 'Уральский',
        'city': 'Екатеринбург',
        'email': 'info@gipermarketdom.ru',
        'website': 'https://gipermarketdom.ru',
        'address': '620014, Свердловская обл., г. Екатеринбург, ул. 8 Марта, д. 179',
    },
    29: {  # АРТ индустрия
        'fo': 'Центральный',
        'city': 'Москва',
        'email': 'info@art-industria.ru',
        'website': 'https://art-industria.ru',
        'address': '121354, г. Москва, ул. Дорогобужская, д. 14, стр. 7, помещ. 3',
    },
    30: {  # Молоток
        'fo': 'Сибирский',
        'city': 'Омск',
        'email': None,
        'website': None,
        'address': '644007, г. Омск',
    },
    31: {  # Сибирь Дистрибуция (закрылись)
        'fo': 'Уральский',
        'city': 'Тюмень',
        'email': None,
        'website': None,
        'address': '625002, Тюменская обл., г. Тюмень, Профсоюзная ул., д. 30',
    },
    32: {  # Сигма
        'fo': 'Центральный',
        'city': 'Москва',
        'email': None,
        'website': None,
        'address': '129327, г. Москва, ул. Лазо, д. 9, помещ. 3/2',
    },
    33: {  # Стройбаза №1
        'fo': 'Уральский',
        'city': 'Челябинск',
        'email': 'shop@s-b-1.ru',
        'website': 'https://stroybaza1.ru',
        'address': '454084, Челябинская обл., г. Челябинск',
    },
    34: {  # СТРОЙТЕ С НАМИ
        'fo': 'Уральский',
        'city': 'Екатеринбург',
        'email': None,
        'website': None,
        'address': 'Свердловская обл., г. Екатеринбург',
    },
    35: {  # Хозстройкомплект
        'fo': 'Уральский',
        'city': 'Челябинск',
        'email': None,
        'website': None,
        'address': '454085, Челябинская обл., г. Челябинск, Шуменская ул., д. 31, 274',
    },
    36: {  # РИТЕЙЛ / Brozexmarket
        'fo': 'Уральский',
        'city': 'Берёзовский',
        'email': 'brozex@brozex.ru',
        'website': 'https://brozexmarket.ru',
        'address': '623701, Свердловская обл., г. Берёзовский, ул. Гагарина, д. 6',
    },
    37: {  # Аксиома
        'fo': 'Сибирский',
        'city': 'Красноярск',
        'email': 'axioma@stp.su',
        'website': 'https://aksioma-krsk.ru',
        'address': '660064, Красноярский край, г. Красноярск, ул. Академика Вавилова, зд. 1г/1',
    },
    38: {  # Баткомплект
        'fo': 'Приволжский',
        'city': 'Пермь',
        'email': 'tdm@batkomplekt.ru',
        'website': 'https://batkomplekt.ru',
        'address': '614014, Пермский край, г. Пермь',
    },
    39: {  # Баумолл - Трейд
        'fo': 'Приволжский',
        'city': 'Пермь',
        'email': 'arenda@baumall.ru',
        'website': 'https://baumoll.ru',
        'address': '614025, Пермский край, г. Пермь, ул. Героев Хасана, зд. 109, каб. 218',
    },
    40: {  # Бенелюкс Сыктывкар
        'fo': 'Северо-Западный',
        'city': 'Сыктывкар',
        'email': None,
        'website': 'https://benelux-market.ru',
        'address': '167000, Республика Коми, г. Сыктывкар',
    },
    41: {  # Кенгуру
        'fo': 'Центральный',
        'city': 'Иваново',
        'email': 'info@kenguru.ru',
        'website': 'https://kenguru-ivanovo.ru',
        'address': '153011, Ивановская обл., г. Иваново, ул. Красных Зорь, д. 34',
    },
    42: {  # Комплекс / Стройся
        'fo': 'Сибирский',
        'city': 'Томск',
        'email': 'stroysaopt@elc.tomsk.ru',
        'website': 'https://stroyasya.ru',
        'address': '634009, Томская обл., г. Томск',
    },
    43: {  # Левша
        'fo': 'Уральский',
        'city': 'Нижний Тагил',
        'email': None,
        'website': None,
        'address': '622005, Свердловская обл., г. Нижний Тагил, ул. Индустриальная, д. 28',
    },
    44: {  # Стройградторг / Стройград Норильск
        'fo': 'Сибирский',
        'city': 'Норильск',
        'email': None,
        'website': None,
        'address': '663300, Красноярский край, г. Норильск',
    },
    45: {  # 1001 КРАСКА
        'fo': 'Центральный',
        'city': 'Москва',
        'email': 'info@1001kraska.ru',
        'website': 'https://1001kraska.ru',
        'address': '121354, г. Москва',
    },
    46: {  # ЭПФ Пилон
        'fo': 'Сибирский',
        'city': 'Красноярск',
        'email': None,
        'website': 'https://pilon24.ru',
        'address': '660049, Красноярский край, г. Красноярск',
    },
    47: {  # Очаг Хоум
        'fo': 'Сибирский',
        'city': 'Красноярск',
        'email': 'info@ochag-home.ru',
        'website': 'https://ochag-home.ru',
        'address': '660013, Красноярский край, г. Красноярск, Томская ул., зд. 4, к. 2, стр. 33',
    },
    48: {  # Порядок / Спектр
        'fo': 'Центральный',
        'city': 'Воронеж',
        'email': 'spektropt1@yandex.ru',
        'website': 'https://poryadok.ru',
        'address': '394028, г. Воронеж',
    },
    49: {  # Протэк
        'fo': 'Центральный',
        'city': 'Воронеж',
        'email': 'protek@protek-vrn.ru',
        'website': 'https://protek-vrn.ru',
        'address': '394018, г. Воронеж',
    },
    50: {  # Реал-Маркет / Стар Трейд (ликвидирован)
        'fo': 'Северо-Западный',
        'city': 'Котлас',
        'email': None,
        'website': None,
        'address': '165300, Архангельская обл., г. Котлас',
    },
    51: {  # ИП Сергеев / Лаки Краски
        'fo': 'Южный',
        'city': 'Волгоград',
        'email': None,
        'website': None,
        'address': 'г. Волгоград',
    },
    52: {  # Стройгигант / Курс-Логистика
        'fo': 'Центральный',
        'city': 'Курск',
        'email': 'n.guseva@stroygigant.ru',
        'website': 'https://stroygigant.ru',
        'address': '305038, Курская обл., г. Курск, ул. Косухина, д. 51',
    },
    53: {  # СтройДепо / НЭКС
        'fo': 'Центральный',
        'city': 'Москва',
        'email': 'ooo-neks@rambler.ru',
        'website': 'https://stroydepo.ru',
        'address': '119192, г. Москва, ул. Мосфильмовская, д. 74Б, офис 219',
    },
    54: {  # Стройсервис АСМ / ИП Воронков
        'fo': 'Уральский',
        'city': 'Арамиль',
        'email': None,
        'website': None,
        'address': '623000, Свердловская обл., г. Арамиль',
    },
    55: {  # Твой дом / Крокус
        'fo': 'Центральный',
        'city': 'Красногорск',
        'email': 'sp@tvoydom.ru',
        'website': 'https://tvoedom.ru',
        'address': '143403, Московская обл., г. Красногорск',
    },
    56: {  # Уютстрой / Стройцентр
        'fo': 'Южный',
        'city': 'Симферополь',
        'email': None,
        'website': None,
        'address': '295000, Республика Крым, г. Симферополь',
    },
    57: {  # Хорошиекраски / FLUGGER
        'fo': 'Южный',
        'city': 'Краснодар',
        'email': 'flugger64@bk.ru',
        'website': 'https://fluegger.ru',
        'address': '350000, Краснодарский край, г. Краснодар',
    },
}

# Применяем все корректировки
for row, data in CORRECTIONS.items():
    for col, val in data.items():
        col_map = {'fo': 2, 'city': 3, 'email': 13, 'website': 15, 'address': 16}
        ws.cell(row=row, column=col_map[col], value=val)

wb.save(dst_file)
print("✅ Все данные скорректированы!")

# Финальная статистика
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

print(f"\n📋 ФИНАЛЬНАЯ СВОДКА ({total} компаний):")
labels = {'fo': 'Федеральный округ', 'city': 'Город', 'email': 'Email', 'website': 'Сайт компании', 'address': 'Почтовый адрес'}
for k, v in filled.items():
    pct = v*100//total
    print(f"  {labels[k]}: {v}/{total} ({pct}%)")

# Детально по пустым email
print(f"\n📧 Компании без email ({total - filled['email']}):")
for row_num in range(3, ws.max_row + 1):
    client = ws.cell(row=row_num, column=4).value
    email = ws.cell(row=row_num, column=13).value
    if client and not email:
        print(f"  Row {row_num}: {str(client)[:50]}")
