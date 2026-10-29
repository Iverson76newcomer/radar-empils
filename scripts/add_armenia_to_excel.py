#!/usr/bin/env python3
"""
Фильтрация армянских компаний (ручная чистка) + добавление в Excel.
Сохраняем исходное форматирование — копируем стиль строки-образца.
"""

import json
import re
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from copy import copy

SRC_EXCEL = '/home/z/my-project/upload/Экспорт_потенциальные_дистрибьюторы_дополнено.xlsx'
SRC_JSON = '/home/z/my-project/scripts/armenia_companies.json'
OUT = '/home/z/my-project/download/Экспорт_потенциальные_дистрибьюторы_с_Арменией.xlsx'


# === РУЧНАЯ ЧИСТКА — оставляем только реальные армянские компании ===
# На основе просмотра 12 отфильтрованных

CLEAN_COMPANIES = [
    {
        'name': 'ШЕН ДЕКОР (салон-магазин)',
        'category': 'Розничные сети и магазины',
        'about': 'Салон-магазин ЛКМ в Ереване. Опт и розница.',
        'features': 'опт, розница',
        'assortment': 'ЛКМ (общий ассортимент)',
        'address': 'г. Ереван',
        'phones': '',
        'emails': '',
        'site': '1285-am.all.biz',
    },
    {
        'name': 'Yerang Shop (Еранг)',
        'category': 'Импортер / Дистрибьютор',
        'about': 'Импорт лакокрасочных материалов из США и Канады в Армению. Широкий ассортимент архитектурных и промышленных красок, эмалей, грунтовок.',
        'features': 'импортёр, опт',
        'assortment': 'акриловые краски, алкидные эмали, грунтовки, лаки',
        'address': 'г. Ереван',
        'phones': '',
        'emails': '',
        'site': 'construction.am',
    },
    {
        'name': 'Berd Paints / TSK Group LLC',
        'category': 'Дистрибьютор',
        'about': 'Розничная и оптовая торговля лакокрасочными материалами. Финишные краски.',
        'features': 'опт, розница',
        'assortment': 'финишные краски, ЛКМ',
        'address': 'г. Ереван, ул. Маро Маргарян, 1/6 (0051)',
        'phones': '+37491 262611',
        'emails': '',
        'site': 'pages.am',
    },
    {
        'name': 'MIX PAINTS LLC (Mix Group)',
        'category': 'Дистрибьютор',
        'about': 'Дистрибьютор лакокрасочных материалов. Адрес: Masis Street 3rd Ln 25, Армения, 0018, Ереван.',
        'features': 'опт',
        'assortment': 'ЛКМ (общий ассортимент)',
        'address': 'г. Ереван, Masis Street 3rd Ln 25',
        'phones': '',
        'emails': 'info@mixgroup.am',
        'site': 'mixgroup.am',
    },
    {
        'name': 'PaintPro Armenia',
        'category': 'Розничные сети и магазины',
        'about': 'Магазин красок и покрытий в Армении. Помощь в выборе, экспертные консультации.',
        'features': 'розница',
        'assortment': 'краски, покрытия',
        'address': 'г. Ереван',
        'phones': '',
        'emails': '',
        'site': 'paintpro.am',
    },
    {
        'name': 'Expert1.am',
        'category': 'Розничные сети и магазины',
        'about': 'Интернет-магазин ЛКМ в Армении. Алкидные эмали, латексные краски, грунтовки.',
        'features': 'розница',
        'assortment': 'алкидные эмали, латексные краски, грунтовки',
        'address': 'г. Ереван',
        'phones': '+374 55 409900',
        'emails': '',
        'site': 'expert1.am',
    },
    {
        'name': 'Импортёры стройматериалов (Арами 64)',
        'category': 'Импортер / Дистрибьютор',
        'about': 'Импортёры стройматериалов, мебели и строительной техники. Окрасочные материалы.',
        'features': 'импортёр',
        'assortment': 'стройматериалы, ЛКМ',
        'address': 'г. Ереван, Арами 64',
        'phones': '+37496 405043',
        'emails': '',
        'site': 'pages.am',
    },
    {
        'name': 'АРЕГ (АБРО салон красок)',
        'category': 'Розничные сети и магазины',
        'about': 'Салон красок АБРО. Розничная продажа ЛКМ.',
        'features': 'розница',
        'assortment': 'ЛКМ (общий ассортимент)',
        'address': 'г. Ереван',
        'phones': '',
        'emails': '',
        'site': 'areg.am',
    },
    {
        'name': 'ShinShin (интернет-магазин)',
        'category': 'Розничные сети и магазины',
        'about': 'Интернет-магазин строительных материалов в Ереване. Краски, инструменты.',
        'features': 'розница',
        'assortment': 'стройматериалы, краски, инструменты',
        'address': 'г. Ереван',
        'phones': '',
        'emails': '',
        'site': 'shinshin.am',
    },
    {
        'name': 'EFFECT GROUP',
        'category': 'Дистрибьютор',
        'about': 'Дистрибьютор алкидных эмалей и лакокрасочных материалов в Армении.',
        'features': 'опт',
        'assortment': 'алкидные эмали (цветные)',
        'address': 'г. Ереван',
        'phones': '',
        'emails': 'info@effectgroup.am',
        'site': 'effectgroup.am',
    },
    {
        'name': 'LTD TOR (дистрибьютор Tikkurila)',
        'category': 'Официальный дистрибьютор (Tikkurila)',
        'about': 'Официальный дистрибьютор Tikkurila в Армении. Покрытия по металлу, напольные системы.',
        'features': 'официальный дилер',
        'assortment': 'Tikkurila, покрытия по металлу, напольные системы',
        'address': 'г. Ереван, пр. Саят-Нова 35 (0025)',
        'phones': '+374 10 55 22 73',
        'emails': '',
        'site': 'tikkurila.com',
    },
    {
        'name': 'BELLA CASA',
        'category': 'Импортер / Дистрибьютор',
        'about': 'Сеть магазинов. Импорт, опт и розница декоративных материалов высокого качества.',
        'features': 'импортёр, опт, розница',
        'assortment': 'декоративные краски, ЛКМ',
        'address': 'г. Ереван',
        'phones': '',
        'emails': '',
        'site': 'lu.kompass.com',
    },
]

print(f"Подготовлено компаний для добавления: {len(CLEAN_COMPANIES)}")

# === ОТКРЫВАЕМ EXCEL ===
wb = openpyxl.load_workbook(SRC_EXCEL)
ws = wb['Лист1']

# Находим последнюю заполненную строку
last_row = 4
for r in range(4, ws.max_row + 1):
    if ws.cell(r, 2).value:
        last_row = r

print(f"Последняя заполненная строка: {last_row}")

# === СОХРАНЯЕМ ИСХОДНОЕ ФОРМАТИРОВАНИЕ ===
# Берём строку-образец — это первая строка с Арменией в исходном файле (была в загрузке)
# Найдём любую строку с Арменией (последнюю, чтобы был тот же цвет)
sample_row = None
for r in range(4, ws.max_row + 1):
    if ws.cell(r, 2).value == 'Армения':
        sample_row = r

print(f"Строка-образец для форматирования: {sample_row}")

# Создаём заливку как у существующих строк Армении
# Возьмём фактический fill из sample_row
sample_fill = None
if sample_row:
    sample_cell = ws.cell(sample_row, 2)
    if sample_cell.fill and sample_cell.fill.fgColor:
        sample_fill = copy(sample_cell.fill)

# Если не получилось — берём нейтральный голубой
if not sample_fill:
    sample_fill = PatternFill('solid', fgColor='FEE2E2')  # светло-розовый

# Стиль
thin = Side(style='thin', color='D1D5DB')
border = Border(left=thin, right=thin, top=thin, bottom=thin)

# Добавляем строки
current_row = last_row + 1
added = 0

for c in CLEAN_COMPANIES:
    # Колонки: 1=пусто, 2=Страна, 3=Категория, 4=Компания, 5=О компании,
    # 6=Ключевые особенности, 7=Ассортимент, 8=Адрес, 9=Телефоны, 10=Email, 11=Сайт
    cells_data = [
        '',
        'Армения',
        c['category'],
        c['name'],
        c['about'],
        c['features'],
        c['assortment'],
        c['address'],
        c['phones'],
        c['emails'],
        c['site'],
    ]
    
    for ci, val in enumerate(cells_data, 1):
        cell = ws.cell(current_row, ci, val)
        # Шрифт как в строке-образце
        if sample_row:
            sample_cell = ws.cell(sample_row, ci)
            cell.font = copy(sample_cell.font)
        else:
            cell.font = Font(name='Arial', size=10, bold=(ci in [2, 4]))
        cell.fill = copy(sample_fill)
        cell.alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)
        cell.border = border
    
    current_row += 1
    added += 1

print(f"Добавлено строк: {added}")
print(f"Конечная строка: {current_row - 1}")

# === ПРОВЕРКА ШИРИНЫ КОЛОНОК ===
col_widths = {2: 14, 3: 22, 4: 32, 5: 50, 6: 35, 7: 32, 8: 28, 9: 22, 10: 22, 11: 20}
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
wb2 = openpyxl.load_workbook(OUT, data_only=True)
ws2 = wb2['Лист1']
from collections import Counter
countries = []
for ri in range(4, ws2.max_row + 1):
    c = ws2.cell(ri, 2).value
    if c:
        countries.append(str(c).strip())
print(f"\n=== Итоговая статистика ===")
for c, cnt in Counter(countries).most_common():
    marker = ' ← +' if c == 'Армения' else ''
    print(f"  {c}: {cnt} компаний{marker}")
print(f"Всего: {sum(Counter(countries).values())} компаний")

# Покажем добавленные Армении
print(f"\n=== Добавленные Армении ===")
n = 0
for ri in range(last_row + 1, ws2.max_row + 1):
    if ws2.cell(ri, 2).value == 'Армения':
        n += 1
        print(f"  {n}. {ws2.cell(ri, 4).value} | {ws2.cell(ri, 3).value} | {ws2.cell(ri, 11).value}")
