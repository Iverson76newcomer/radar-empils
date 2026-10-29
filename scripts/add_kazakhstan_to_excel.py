#!/usr/bin/env python3
"""
Фильтрация и добавление казахских компаний в Excel.
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from copy import copy

SRC_EXCEL = '/home/z/my-project/download/Экспорт_потенциальные_дистрибьюторы_с_Узбекистаном.xlsx'
OUT = '/home/z/my-project/download/Экспорт_потенциальные_дистрибьюторы_с_Казахстаном.xlsx'


# === РУЧНАЯ ЧИСТКА — реальные казахские компании ===
CLEAN_COMPANIES = [
    {
        'name': 'ТОО КазПолимерКом (ЛКЗ)',
        'category': 'Производитель',
        'about': 'Лакокрасочный завод «КАЗПОЛИМЕРКОМ» с 2014 года. Поставщик краски: водоэмульсии, грунтовки, эмали для внутренних и наружных работ.',
        'features': 'производитель, опт',
        'assortment': 'водоэмульсии, грунтовки, эмали',
        'address': 'г. Алматы',
        'phones': '',
        'emails': '',
        'site': 'kazpolymercom.kz',
    },
    {
        'name': 'Irbis Color (ТОО Melcor — новое направление)',
        'category': 'Дистрибьютор',
        'about': 'Динамично развивающаяся компания. Поставки лакокрасочных материалов, красок, лаков и эмалей.',
        'features': 'дистрибьютор, опт',
        'assortment': 'краски, лаки, эмали',
        'address': 'г. Алматы',
        'phones': '',
        'emails': '',
        'site': 'irbiscolor.kz',
    },
    {
        'name': 'Madeira',
        'category': 'Дистрибьютор',
        'about': 'Поставщик качественных лакокрасочных материалов для любых задач.',
        'features': 'опт, розница',
        'assortment': 'ЛКМ (общий ассортимент)',
        'address': 'г. Астана, ул. Шары Жиенкуловой 15/1',
        'phones': '+7 777 815-53-33',
        'emails': '',
        'site': 'madeira.kz',
    },
    {
        'name': 'Aqua Paint (ТОО Казкомплектстрой)',
        'category': 'Дистрибьютор',
        'about': 'С 1999 года реализация лакокрасочной продукции на строительном рынке Астаны.',
        'features': 'дистрибьютор, опт',
        'assortment': 'ЛКМ (общий ассортимент)',
        'address': 'г. Астана',
        'phones': '',
        'emails': '',
        'site': 'aqua-paint.kz',
    },
    {
        'name': 'ЛКЗ «Краски НИПОЛ»',
        'category': 'Производитель',
        'about': 'Производство и продажа ЛКМ. Промышленные лакокрасочные материалы от эконом до премиум. ISO 9001-2015.',
        'features': 'производитель, опт',
        'assortment': 'промышленные ЛКМ, антикоррозийные',
        'address': 'г. Алматы',
        'phones': '',
        'emails': '',
        'site': 'nipol.kz',
    },
    {
        'name': 'ПО КазМетСервис (kzmc.kz)',
        'category': 'Дистрибьютор',
        'about': 'Реализация лакокрасочных материалов в Республике Казахстан. Опт от 100 кг. Доставка по РК и СНГ. Филиалы: Астана, Караганда.',
        'features': 'опт, доставка по СНГ',
        'assortment': 'ЛКМ, эмали, грунт-эмали',
        'address': 'г. Астана; г. Караганда',
        'phones': '+7 7212 940-339 (Караганда)',
        'emails': '',
        'site': 'astana.kzmc.kz',
    },
    {
        'name': 'Boyau House Astana',
        'category': 'Дистрибьютор',
        'about': 'Краски, грунтовки, эмали, растворители, декоративные штукатурки на складе. Доставка по Астане и РК. Колеровка в 20 000 оттенков.',
        'features': 'опт, розница, колеровка',
        'assortment': 'краски, грунтовки, эмали, растворители, декоративные штукатурки',
        'address': 'г. Астана',
        'phones': '',
        'emails': '',
        'site': 'boyauhouse.kz',
    },
    {
        'name': 'ТОО Best Home',
        'category': 'Производитель',
        'about': 'Производитель фасадных красок (по каталогу Optoviki.kz).',
        'features': 'производитель',
        'assortment': 'фасадные краски',
        'address': 'г. Алматы',
        'phones': '',
        'emails': '',
        'site': 'optoviki.kz',
    },
    {
        'name': 'ТОО PIK',
        'category': 'Производитель',
        'about': 'Производитель фасадных красок в Казахстане.',
        'features': 'производитель',
        'assortment': 'фасадные краски',
        'address': 'г. Алматы',
        'phones': '',
        'emails': '',
        'site': 'optoviki.kz',
    },
    {
        'name': 'ЖАНАЖОЛ ПЛЮС',
        'category': 'Дистрибьютор',
        'about': 'Оптовая продажа красок в Астане. Сибирский лакокрасочный завод.',
        'features': 'опт',
        'assortment': 'краски Сибирского ЛКЗ',
        'address': 'г. Астана',
        'phones': '',
        'emails': '',
        'site': 'astana.pulscen.kz',
    },
    {
        'name': 'Almaty Paints LLP',
        'category': 'Производитель / Дистрибьютор',
        'about': 'Производитель лакокрасочных материалов и дорожной разметки. Официальный дистрибьютор красок и декоративных покрытий.',
        'features': 'производитель, дистрибьютор',
        'assortment': 'ЛКМ, краски для дорожной разметки, декоративные покрытия',
        'address': 'г. Алматы',
        'phones': '',
        'emails': '',
        'site': 'centralasiacoatingshow.com',
    },
    {
        'name': 'NEOCHIM',
        'category': 'Производитель',
        'about': 'Производитель ЛКМ в Алматы.',
        'features': 'производитель',
        'assortment': 'ЛКМ',
        'address': 'г. Алматы',
        'phones': '',
        'emails': '',
        'site': 'rentechdigital.com',
    },
    {
        'name': 'Nanotech Paints',
        'category': 'Производитель',
        'about': 'Производитель ЛКМ в Алматы (нанотехнологии).',
        'features': 'производитель',
        'assortment': 'ЛКМ (нанотехнологии)',
        'address': 'г. Алматы',
        'phones': '',
        'emails': '',
        'site': 'rentechdigital.com',
    },
    {
        'name': 'ТОО Kazimport (Baiparts)',
        'category': 'Импортер / Дистрибьютор',
        'about': 'Один из ведущих производителей и импортёров красок в Казахстане. Работают по всей стране.',
        'features': 'импортёр, опт',
        'assortment': 'ЛКМ (общий ассортимент)',
        'address': 'г. Алматы',
        'phones': '+7 707 104 09 02, +7 707 300 81 08',
        'emails': 'info@baiparts.kz',
        'site': 'kazimport.kz',
    },
    {
        'name': 'QORGAU KZ',
        'category': 'Производитель / Дистрибьютор',
        'about': 'Огнезащитные краски и покрытия для промышленных и строительных компаний Казахстана.',
        'features': 'производитель, опт',
        'assortment': 'огнезащитные краски, покрытия',
        'address': 'г. Алматы',
        'phones': '+7 706 647 8999',
        'emails': '',
        'site': 'qorgau.kz',
    },
    {
        'name': 'iGateway (International Paint)',
        'category': 'Официальный дистрибьютор (International Paint)',
        'about': 'Авторизованный дистрибьютор International Paint в Казахстане. Высокоэффективные морские и защитные покрытия.',
        'features': 'официальный дилер',
        'assortment': 'International Paint, морские, защитные покрытия',
        'address': 'г. Алматы',
        'phones': '+7 727 356 1000',
        'emails': 'info@igateway',
        'site': 'igateway.ltd',
    },
    {
        'name': 'Amanat Astana',
        'category': 'Дистрибьютор',
        'about': 'Ведущий поставщик красок и покрытий для промышленных и строительных проектов в Казахстане.',
        'features': 'дистрибьютор, опт',
        'assortment': 'промышленные покрытия, ЛКМ',
        'address': 'г. Астана',
        'phones': '',
        'emails': '',
        'site': 'amanatastana.kz',
    },
    {
        'name': 'ZENITH (Industrial Painting)',
        'category': 'Официальный дистрибьютор',
        'about': 'Официальный дистрибьютор ведущих производителей красок. Промышленная окраска и защита крупных объектов и стали.',
        'features': 'официальный дилер, опт',
        'assortment': 'промышленные краски, покрытия',
        'address': 'г. Алматы, ул. Чаплина 71/66, офис 708 (050019)',
        'phones': '+7 701 800-46-26',
        'emails': '',
        'site': 'zenith.kz',
    },
    {
        'name': 'Nasyr Stroy',
        'category': 'Дистрибьютор',
        'about': 'Ведущий поставщик строительных материалов, водоэмульсий и полиэтиленовых труб в Шымкенте.',
        'features': 'дистрибьютор, опт',
        'assortment': 'водоэмульсии, стройматериалы',
        'address': 'г. Шымкент',
        'phones': '',
        'emails': '',
        'site': 'nasyrstroy.kz',
    },
    {
        'name': 'LKM.com.kz (итальянские краски)',
        'category': 'Импортер / Дистрибьютор',
        'about': 'Поставка лакокрасочных материалов и инструмента от мировых брендов. Доставка в любую точку Казахстана. Итальянские краски и лаки для мебели.',
        'features': 'импортёр, опт',
        'assortment': 'итальянские краски, лаки для мебели',
        'address': 'г. Шымкент',
        'phones': '',
        'emails': '',
        'site': 'lkm.com.kz',
    },
    {
        'name': 'Vesta Company (ТОО)',
        'category': 'Дистрибьютор',
        'about': 'Оптовая продажа красок в Шымкенте. Краска Престиж, кровельные краски.',
        'features': 'опт',
        'assortment': 'краска Престиж, кровельные краски',
        'address': 'г. Шымкент',
        'phones': '',
        'emails': '',
        'site': 'shyimkent.satu.kz',
    },
    {
        'name': 'Импульс-КЗ',
        'category': 'Дистрибьютор',
        'about': 'Оптовая продажа алкидных эмалей ПФ-115 ГОСТ в Караганде. Цвета: белая, серая. Фасовки 20, 25, 30 кг.',
        'features': 'опт',
        'assortment': 'алкидные эмали ПФ-115',
        'address': 'г. Караганда',
        'phones': '',
        'emails': '',
        'site': 'impuls-kz.kz',
    },
    {
        'name': 'МГК Авангард',
        'category': 'Производитель / Поставщик',
        'about': 'Поставщик и производитель фасадных красок в Казахстане (по каталогу Optoviki.kz).',
        'features': 'производитель, опт',
        'assortment': 'фасадные краски',
        'address': 'г. Алматы',
        'phones': '',
        'emails': '',
        'site': 'optoviki.kz',
    },
    {
        'name': 'Rainbow Paints',
        'category': 'Производитель',
        'about': 'Производитель ЛКМ в Алматы.',
        'features': 'производитель',
        'assortment': 'ЛКМ',
        'address': 'г. Алматы',
        'phones': '',
        'emails': '',
        'site': 'rentechdigital.com',
    },
]

print(f"Подготовлено компаний для добавления: {len(CLEAN_COMPANIES)}")

# === ОТКРЫВАЕМ EXCEL ===
wb = openpyxl.load_workbook(SRC_EXCEL)
ws = wb['Лист1']

last_row = 4
for r in range(4, ws.max_row + 1):
    if ws.cell(r, 2).value:
        last_row = r

print(f"Последняя заполненная строка: {last_row}")

# Строка-образец (Казахстан)
sample_row = None
for r in range(4, ws.max_row + 1):
    if ws.cell(r, 2).value == 'Казахстан':
        sample_row = r
        break

print(f"Строка-образец для форматирования: {sample_row}")

sample_fill = None
if sample_row:
    sample_cell = ws.cell(sample_row, 2)
    if sample_cell.fill and sample_cell.fill.fgColor:
        sample_fill = copy(sample_cell.fill)
if not sample_fill:
    sample_fill = PatternFill('solid', fgColor='DCFCE7')

thin = Side(style='thin', color='D1D5DB')
border = Border(left=thin, right=thin, top=thin, bottom=thin)

current_row = last_row + 1
added = 0

for c in CLEAN_COMPANIES:
    cells_data = [
        '',
        'Казахстан',
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
    marker = ' ← +' if c == 'Казахстан' else ''
    print(f"  {c}: {cnt} компаний{marker}")
print(f"Всего: {sum(Counter(countries).values())} компаний")
