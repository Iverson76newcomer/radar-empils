#!/usr/bin/env python3
"""
Фильтрация и добавление сербских компаний в Excel.
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from copy import copy

SRC_EXCEL = '/home/z/my-project/download/Экспорт_потенциальные_дистрибьюторы_с_Молдовой.xlsx'
OUT = '/home/z/my-project/download/Экспорт_потенциальные_дистрибьюторы_с_Сербией.xlsx'


CLEAN_COMPANIES = [
    {
        'name': 'PROHEMA SRB',
        'category': 'Дистрибьютор',
        'about': 'Дистрибуция оборудования, красок, промышленных покрытий, материалов, разбавителей и другой продукции для строительства и химии.',
        'features': 'дистрибьютор, опт',
        'assortment': 'краски, промышленные покрытия, разбавители',
        'address': 'Сербия',
        'phones': '',
        'emails': '',
        'site': 'prohema.rs',
    },
    {
        'name': 'Credicom International',
        'category': 'Импортер / Дистрибьютор',
        'about': 'Импорт и дистрибуция сырья для производства красок и лаков, огнеупорных материалов, отделочных строительных работ.',
        'features': 'импортёр, опт',
        'assortment': 'сырьё для ЛКМ, огнеупорные материалы',
        'address': 'Сербия',
        'phones': '',
        'emails': '',
        'site': 'credicom.rs',
    },
    {
        'name': 'Tritonex (Prvako Group)',
        'category': 'Производитель / Дистрибьютор',
        'about': 'Часть группы Prvako. Производство, опт, экспорт, розница и дистрибуция строительных материалов.',
        'features': 'производитель, опт, экспорт',
        'assortment': 'стройматериалы, ЛКМ',
        'address': 'Сербия',
        'phones': '',
        'emails': '',
        'site': 'tritonex.rs',
    },
    {
        'name': 'BTA — Balkan Technical Aerosol',
        'category': 'Производитель / Дистрибьютор',
        'about': 'Специализация на поставках аэрозольных красок, монтажных пен, автохимии, клеев и герметиков для строительства и ремонта.',
        'features': 'производитель, опт',
        'assortment': 'аэрозольные краски, пена, автохимия, клеи, герметики',
        'address': 'Сербия',
        'phones': '',
        'emails': '',
        'site': 'eng.bta.rs',
    },
    {
        'name': 'ILI BOJA (Farbara Novi Sad)',
        'category': 'Розничные сети и магазины',
        'about': 'Магазин красок и лаков в Нови-Саде. 3 номера телефона, отдел продаж.',
        'features': 'розница, опт',
        'assortment': 'ЛКМ (общий ассортимент)',
        'address': 'г. Нови-Сад',
        'phones': '+381 21 325 11 61, +381 62 1504 676, +381 63 504 676',
        'emails': 'info@iliboja.rs, prodaja@iliboja.rs',
        'site': 'iliboja.rs',
    },
    {
        'name': 'Alexandar Cosmetics (Novi Sad)',
        'category': 'Розничные сети и магазины',
        'about': 'Магазин в Нови-Саде, ул. Цара Душана 69. Пн-Пт 8:00-16:00.',
        'features': 'розница',
        'assortment': 'ЛКМ, косметика',
        'address': 'г. Нови-Сад, ул. Цара Душана 69',
        'phones': '+381 21 455 558',
        'emails': '',
        'site': 'alexandar-cosmetics.com',
    },
    {
        'name': 'Hemiprodukt boje i lakovi',
        'category': 'Дистрибьютор',
        'about': 'Магазин красок и лаков в Нови-Саде. Ул. Масарикова 1. Опт и розница.',
        'features': 'опт, розница',
        'assortment': 'ЛКМ (общий ассортимент)',
        'address': 'г. Нови-Сад, ул. Масарикова 1 (21000)',
        'phones': '+381 21 2998 420, +381 60 531 8220, +381 60 531 7440',
        'emails': 'hemiboje@gmail.com',
        'site': 'hemiprodukt.co.rs',
    },
    {
        'name': 'Auto Boje Telming (Novi Sad)',
        'category': 'Дистрибьютор',
        'about': 'Продажа автоэмалей и лакокрасочных материалов в Нови-Саде. Качественные материалы для авторемонта.',
        'features': 'опт, розница',
        'assortment': 'автоэмали, лаки',
        'address': 'г. Нови-Сад',
        'phones': '+381 21 6618 180, +381 63 550 819',
        'emails': 'info@autobojetelming.rs, office@autobojetelming.rs',
        'site': 'autobojetelming.rs',
    },
    {
        'name': 'Dekorex (veleprodaja)',
        'category': 'Дистрибьютор',
        'about': 'Оптовая торговля. В ассортименте крупные производители красок и лаков: AkzoNobel, Trilak, Polifarbe, Köber.',
        'features': 'опт',
        'assortment': 'AkzoNobel, Trilak, Polifarbe, Köber',
        'address': 'Сербия',
        'phones': '',
        'emails': '',
        'site': 'dekorex.rs',
    },
    {
        'name': 'BOLAGO-M',
        'category': 'Розничные сети и магазины',
        'about': 'Магазин красок и лаков в Сербии. Пон-Пт 7:00-16:00, Сб 7:00-14:00.',
        'features': 'розница',
        'assortment': 'ЛКМ (общий ассортимент)',
        'address': 'Сербия',
        'phones': '+381 32 5371 476, +381 63 695 272, +381 64 703 9390',
        'emails': 'prodaja@bolago-m.rs',
        'site': 'bolago-m.rs',
    },
    {
        'name': 'Rotor Boje i Lakovi (Novi Sad)',
        'category': 'Розничные сети и магазины',
        'about': 'Магазин красок и лаков в Нови-Саде (по каталогу Planplus).',
        'features': 'розница',
        'assortment': 'ЛКМ (общий ассортимент)',
        'address': 'г. Нови-Сад, Трг Республики 11 (21000)',
        'phones': '',
        'emails': '',
        'site': 'rotorns.co.rs',
    },
    {
        'name': 'Boje i lakovi Royal Z&G',
        'category': 'Розничные сети и магазины',
        'about': 'Магазин красок и лаков в Нови-Саде, ул. Мичуринова 42.',
        'features': 'розница',
        'assortment': 'ЛКМ (общий ассортимент)',
        'address': 'г. Нови-Сад, ул. Мичуринова 42',
        'phones': '',
        'emails': '',
        'site': 'planplus.rs',
    },
    {
        'name': 'Color Star (boje i lakovi)',
        'category': 'Розничные сети и магазины',
        'about': 'Магазин красок и лаков в Нови-Саде, ул. Бранимира.',
        'features': 'розница',
        'assortment': 'ЛКМ (общий ассортимент)',
        'address': 'г. Нови-Сад, ул. Бранимира',
        'phones': '',
        'emails': '',
        'site': 'planplus.rs',
    },
    {
        'name': 'Farbe i lakovi Novi Sad',
        'category': 'Розничные сети и магазины',
        'about': 'Магазин в Нови-Саде. Автоэмали, лаки, покрытия, материалы для гидроизоляции, декоративные техники, покрытия для дерева и металла.',
        'features': 'розница',
        'assortment': 'автоэмали, лаки, гидроизоляция, покрытия для дерева/металла',
        'address': 'г. Нови-Сад',
        'phones': '',
        'emails': '',
        'site': 'globalmediagroup.rs',
    },
    {
        'name': 'UC Partizan',
        'category': 'Производитель',
        'about': 'Производитель стеновых красок. Белград, 11147. Ул. (нет данных).',
        'features': 'производитель, опт',
        'assortment': 'стеновые краски',
        'address': 'г. Белград (11147)',
        'phones': '+381 11 2393295',
        'emails': 'partizan.uc@gmail.com',
        'site': 'ucpartizan.com',
    },
    {
        'name': 'ForkSide (российско-сербское производство)',
        'category': 'Производитель',
        'about': 'Российско-сербская производственная компания по производству инновационных лакокрасочных материалов и смол с локализацией на территории РФ.',
        'features': 'производитель',
        'assortment': 'инновационные ЛКМ, смолы',
        'address': 'Сербия / Россия',
        'phones': '',
        'emails': '',
        'site': 'forkside.ru',
    },
    {
        'name': 'PANZER (сербский производитель)',
        'category': 'Производитель',
        'about': 'Сербский производитель лакокрасочных материалов. Молотковые краски для внутренних и наружных работ.',
        'features': 'производитель, опт',
        'assortment': 'молотковые краски',
        'address': 'Сербия',
        'phones': '',
        'emails': '',
        'site': '1-str.ru',
    },
    {
        'name': 'KB Armature',
        'category': 'Дистрибьютор',
        'about': 'Дистрибьютор. Белград, ул. Радничка 5 (11030).',
        'features': 'дистрибьютор, опт',
        'assortment': 'ЛКМ, арматура',
        'address': 'г. Белград, ул. Радничка 5 (11030)',
        'phones': '+381 11 3557 541',
        'emails': 'office@kbarmature.rs',
        'site': 'kbarmature.rs',
    },
    {
        'name': 'Kompass Serbia (paint manufacturers)',
        'category': 'Производитель',
        'about': 'Каталог производителей красок, лаков и подобных покрытий в Сербии. Опт.',
        'features': 'производитель, опт',
        'assortment': 'ЛКМ',
        'address': 'г. Белград',
        'phones': '+381 11 2085505',
        'emails': '',
        'site': 'kompass.com',
    },
]

print(f"Подготовлено компаний для добавления: {len(CLEAN_COMPANIES)}")

wb = openpyxl.load_workbook(SRC_EXCEL)
ws = wb['Лист1']

last_row = 4
for r in range(4, ws.max_row + 1):
    if ws.cell(r, 2).value:
        last_row = r

print(f"Последняя заполненная строка: {last_row}")

sample_row = None
for r in range(4, ws.max_row + 1):
    if ws.cell(r, 2).value == 'Сербия':
        sample_row = r
        break

print(f"Строка-образец для форматирования: {sample_row}")

sample_fill = None
if sample_row:
    sample_cell = ws.cell(sample_row, 2)
    if sample_cell.fill and sample_cell.fill.fgColor:
        sample_fill = copy(sample_cell.fill)
if not sample_fill:
    sample_fill = PatternFill('solid', fgColor='FEF3C7')

thin = Side(style='thin', color='D1D5DB')
border = Border(left=thin, right=thin, top=thin, bottom=thin)

current_row = last_row + 1
added = 0

for c in CLEAN_COMPANIES:
    cells_data = [
        '',
        'Сербия',
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

col_widths = {2: 14, 3: 28, 4: 36, 5: 50, 6: 35, 7: 32, 8: 32, 9: 28, 10: 28, 11: 22}
for ci, w in col_widths.items():
    cur = ws.column_dimensions[get_column_letter(ci)].width
    if not cur or cur < w:
        ws.column_dimensions[get_column_letter(ci)].width = w

wb.save(OUT)
import os
print(f"\nOK: {OUT}")
print(f"Размер: {os.path.getsize(OUT) / 1024:.1f} KB")

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
    marker = ' ← +' if c == 'Сербия' else ''
    print(f"  {c}: {cnt} компаний{marker}")
print(f"Всего: {sum(Counter(countries).values())} компаний")
