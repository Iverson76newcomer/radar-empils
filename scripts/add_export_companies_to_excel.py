#!/usr/bin/env python3
"""
Фильтрация собранных компаний и добавление в Excel-файл.
Убираем доски объявлений, классифицируем, добавляем в лист 1 после строки 105.
"""

import json
import re
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from copy import copy

SRC_EXCEL = '/home/z/my-project/upload/Экспорт потенциальтные дистрибьюторы.xlsx'
SRC_JSON = '/home/z/my-project/scripts/export_companies.json'
OUT = '/home/z/my-project/download/Экспорт_потенциальные_дистрибьюторы_дополнено.xlsx'

# Загрузка собранных данных
with open(SRC_JSON, 'r', encoding='utf-8') as f:
    all_data = json.load(f)

# === ФИЛЬТРАЦИЯ ===
# Сайты-доски объявлений, которые надо исключить
BAD_HOSTS = [
    'lalafo', 'somon.tj', 'bizorg', 'optlist', 'all.biz', '2sp.me', 'opendatalab',
    'linkedin', 'facebook', 'instagram', 'youtube', 'wikipedia', 'yandex', '2gis',
    'yelp', 'mino.tj', 'flagma', 'metaprom', 'znaki154', 'nekei',
    'wildberries', 'ozon', 'avito', 'orgpage', 'rusbase', 'spr.ru',
    'yell', 'yellowpages', 'rusprofile',
]

# Плохие слова в названии
BAD_NAME_KEYWORDS = [
    'оптом краски', 'купить', 'заказать', 'цена', 'отзывы', 'каталог', 'доска',
    'объявлени', 'услуги', 'найти', 'просмотр', 'закрыто', 'классиф',
]

def is_bad_company(c):
    """Возвращает True, если компания — это доска объявлений или информационный сайт."""
    host = c.get('host', '').lower()
    name = c.get('name', '').lower()
    site = c.get('site', '').lower()
    
    # Доски объявлений
    for bad in BAD_HOSTS:
        if bad in host or bad in site:
            return True
    
    # Плохие слова в названии
    for bad in BAD_NAME_KEYWORDS:
        if bad in name:
            return True
    
    # Слишком короткое название
    if len(c.get('name', '')) < 5:
        return True
    
    # Только цифры/символы
    if not re.search(r'[а-яА-Яa-zA-Z]', c.get('name', '')):
        return True
    
    return False


def classify_company(c, country):
    """Определяет категорию компании на основе сниппета."""
    text = (c.get('name', '') + ' ' + c.get('snippet', '')).lower()
    
    if any(w in text for w in ['производств', 'завод', 'производит', 'выпускает', 'factory', 'manufacturer']):
        return 'Производитель'
    if any(w in text for w in ['дистрибьют', 'импортер', 'официальн', 'дилер', 'distributor', 'import']):
        return 'Дистрибьютор'
    if any(w in text for w in ['розничн', 'магазин', 'сеть ', 'торгов', 'retail', 'shop']):
        return 'Розничные сети и магазины'
    if any(w in text for w in ['опт', 'оптов', 'wholesale', 'крупный опт']):
        return 'Дистрибьютор'
    return 'Дистрибьютор'  # по умолчанию — для экспортной базы


def extract_address(snippet, country):
    """Пытается извлечь адрес из сниппета."""
    # Ищем город
    cities = {
        'Киргизия': ['Бишкек', 'Ош', 'Каракол', 'Джалал-Абад', 'Талас'],
        'Монголия': ['Улаанбаатар', 'Улан-Батор', 'Эрдэнэт', 'Дархан'],
        'Таджикистан': ['Душанбе', 'Худжанд', 'Бохтар', 'Куляб'],
    }
    for city in cities.get(country, []):
        if city.lower() in snippet.lower():
            return f'г. {city}'
    return ''


def extract_about(name, snippet):
    """Формирует короткое описание компании."""
    # Берём первое осмысленное предложение из сниппета
    s = snippet[:300].strip()
    # Убираем "Missing:" и пр.
    s = re.sub(r'Missing:.*$', '', s, flags=re.IGNORECASE)
    s = re.sub(r'Show more.*$', '', s, flags=re.IGNORECASE)
    return s[:250]


def extract_features(snippet):
    """Извлекает ключевые особенности."""
    text = snippet[:400]
    features = []
    if 'акция' in text.lower() or 'скидк' in text.lower():
        features.append('акции')
    if 'доставк' in text.lower():
        features.append('доставка')
    if 'опт' in text.lower():
        features.append('опт')
    if 'розниц' in text.lower():
        features.append('розница')
    if 'импорт' in text.lower() or 'официальн' in text.lower():
        features.append('импортёр')
    return ', '.join(features)


def extract_assortment(snippet, name):
    """Извлекает ассортимент."""
    text = (name + ' ' + snippet).lower()
    products = []
    if 'акрил' in text:
        products.append('акриловые краски')
    if 'алкид' in text or 'эмаль' in text:
        products.append('алкидные эмали')
    if 'маслян' in text:
        products.append('масляные краски')
    if 'водоэмульс' in text or 'вододисперс' in text:
        products.append('водоэмульсионные')
    if 'лак' in text:
        products.append('лаки')
    if 'грунт' in text:
        products.append('грунтовки')
    if 'декоратив' in text:
        products.append('декоративные')
    if 'фасад' in text:
        products.append('фасадные')
    if 'молотк' in text:
        products.append('молотковые')
    if not products:
        products.append('ЛКМ (общий ассортимент)')
    return ', '.join(products[:5])


# === ФИЛЬТРУЕМ И ПОДГОТАВЛИВАЕМ ДАННЫЕ ===
print("=== Фильтрация и подготовка ===\n")

clean_data = {}
for country, companies in all_data.items():
    clean = []
    seen_names = set()
    for c in companies:
        if is_bad_company(c):
            continue
        # Дедупликация по имени
        name_key = c['name'].lower()[:25]
        if name_key in seen_names:
            continue
        seen_names.add(name_key)
        
        c['category'] = classify_company(c, country)
        c['address'] = extract_address(c['snippet'], country)
        c['about'] = extract_about(c['name'], c['snippet'])
        c['features'] = extract_features(c['snippet'])
        c['assortment'] = extract_assortment(c['snippet'], c['name'])
        c['phones_str'] = ', '.join(c.get('phones', []))
        c['emails_str'] = ', '.join(c.get('emails', []))
        clean.append(c)
    
    clean_data[country] = clean
    print(f"  {country}: было {len(companies)}, после фильтра {len(clean)}")

# === ОТКРЫВАЕМ ИСХОДНЫЙ EXCEL И ДОБАВЛЯЕМ СТРОКИ ===
print("\n=== Добавление в Excel ===")

wb = openpyxl.load_workbook(SRC_EXCEL)
ws = wb['Лист1']

# Стиль для новых строк (как в существующих)
thin = Side(style='thin', color='D1D5DB')
border = Border(left=thin, right=thin, top=thin, bottom=thin)
normal_font = Font(name='Arial', size=10)
bold_font = Font(name='Arial', size=10, bold=True)

# Цвета по странам (заливка)
COUNTRY_FILLS = {
    'Киргизия': PatternFill('solid', fgColor='DBEAFE'),  # голубой
    'Монголия': PatternFill('solid', fgColor='FEE2E2'),  # розовый
    'Таджикистан': PatternFill('solid', fgColor='DCFCE7'),  # зелёный
}

# Находим последнюю заполненную строку
last_row = 4
for r in range(4, ws.max_row + 1):
    if ws.cell(r, 2).value:
        last_row = r

print(f"  Последняя заполненная строка в файле: {last_row}")

# Добавляем строки для 3 новых стран
current_row = last_row + 1
added = 0

for country in ['Киргизия', 'Монголия', 'Таджикистан']:
    companies = clean_data[country]
    fill = COUNTRY_FILLS[country]
    
    for c in companies:
        # Колонки: 2=Страна, 3=Категория, 4=Компания, 5=О компании, 6=Ключевые особенности,
        # 7=Ассортимент, 8=Адрес, 9=Телефоны, 10=Email, 11=Сайт
        cells_data = [
            '',  # col 1 пустая (как в исходнике)
            country,
            c['category'],
            c['name'],
            c['about'],
            c['features'],
            c['assortment'],
            c['address'],
            c['phones_str'],
            c['emails_str'],
            c.get('site', ''),
        ]
        
        for ci, val in enumerate(cells_data, 1):
            cell = ws.cell(current_row, ci, val)
            cell.font = bold_font if ci in [2, 4] else normal_font
            cell.fill = fill
            cell.alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)
            cell.border = border
        
        current_row += 1
        added += 1

print(f"  Добавлено строк: {added}")
print(f"  Конечная строка: {current_row - 1}")

# === ШИРИНА КОЛОНОК ===
# Возьмём ширину из исходного файла, но проверим/поправим
col_widths = {2: 14, 3: 22, 4: 32, 5: 45, 6: 35, 7: 32, 8: 28, 9: 22, 10: 22, 11: 20}
for ci, w in col_widths.items():
    cur = ws.column_dimensions[get_column_letter(ci)].width
    if not cur or cur < w:
        ws.column_dimensions[get_column_letter(ci)].width = w

# Сохраняем
wb.save(OUT)

import os
print(f"\nOK: {OUT}")
print(f"Размер: {os.path.getsize(OUT) / 1024:.1f} KB")

# === ИТОГОВАЯ СТАТИСТИКА ===
print(f"\n=== ИТОГОВАЯ СТАТИСТИКА ===")
print(f"Киргизия: {len(clean_data['Киргизия'])} компаний")
print(f"Монголия: {len(clean_data['Монголия'])} компаний")
print(f"Таджикистан: {len(clean_data['Таджикистан'])} компаний")
print(f"ВСЕГО добавлено: {added} компаний")
print(f"Исходно в файле: 88 компаний")
print(f"Итого в файле: {88 + added} компаний по 12 странам")
