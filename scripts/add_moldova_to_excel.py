#!/usr/bin/env python3
"""
Фильтрация и добавление молдавских компаний в Excel.
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from copy import copy

SRC_EXCEL = '/home/z/my-project/download/Экспорт_потенциальные_дистрибьюторы_с_Туркменистаном.xlsx'
OUT = '/home/z/my-project/download/Экспорт_потенциальные_дистрибьюторы_с_Молдовой.xlsx'


CLEAN_COMPANIES = [
    {
        'name': 'SUPRATEN.md (строительный магазин №1)',
        'category': 'Розничные сети и магазины',
        'about': 'Строительный магазин №1 в Молдове. Большой выбор красок, водоэмульсионных, универсальных красок. Доставка по стране.',
        'features': 'розница, опт',
        'assortment': 'универсальные, водоэмульсионные краски',
        'address': 'г. Кишинев',
        'phones': '+373 22 268 000',
        'emails': '',
        'site': 'supraten.md',
    },
    {
        'name': 'Color.md',
        'category': 'Розничные сети и магазины',
        'about': 'Магазин красок в Кишиневе. 3 адреса: бул. Дечебал 91, ул. Л. Делану 14D, ул. Каля.',
        'features': 'розница',
        'assortment': 'ЛКМ (общий ассортимент)',
        'address': 'г. Кишинев (3 точки)',
        'phones': '+373 68 949 466, +373 68 946 644',
        'emails': '',
        'site': 'color.md',
    },
    {
        'name': 'Colorista.md',
        'category': 'Дистрибьютор',
        'about': 'Дистрибьютор красок в Кишиневе: Sniezka, Aura, Eskaro, Magnat и 65+ брендов. 2500+ продуктов, грунты, инструменты. Доставка.',
        'features': 'дистрибьютор, опт, розница',
        'assortment': 'Sniezka, Aura, Eskaro, Magnat, грунты',
        'address': 'г. Кишинев',
        'phones': '+373 69 065 173',
        'emails': '',
        'site': 'colorista.md',
    },
    {
        'name': 'Caparol Moldova',
        'category': 'Официальный дистрибьютор (Caparol)',
        'about': 'Официальный дистрибьютор Caparol в Молдове. Внутренние и внешние краски, грунты, структурные штукатурки, лаки.',
        'features': 'официальный дилер, опт',
        'assortment': 'Caparol, интерьерные/фасадные краски, штукатурки, лаки',
        'address': 'г. Кишинев',
        'phones': '',
        'emails': '',
        'site': 'caparol.md',
    },
    {
        'name': 'Boero Moldova',
        'category': 'Официальный дистрибьютор (Boero)',
        'about': 'Официальный дистрибьютор Boero в Молдове. Кишинев, ул. Гренобль 41.',
        'features': 'официальный дилер, опт, розница',
        'assortment': 'Boero, декоративные краски',
        'address': 'г. Кишинев, ул. Гренобль 41',
        'phones': '+373 60 502 502',
        'emails': 'info@boero.md',
        'site': 'boero.md',
    },
    {
        'name': 'Renner Moldova (Elite Vernici Srl)',
        'category': 'Официальный дистрибьютор (Renner)',
        'about': 'Первый и единственный дилер красок Renner по дереву в Молдове. Ответственен за дистрибуцию в Восточной Европе.',
        'features': 'официальный дилер, опт',
        'assortment': 'Renner, краски по дереву',
        'address': 'г. Кишинев',
        'phones': '',
        'emails': '',
        'site': 'rennermoldova.com',
    },
    {
        'name': 'Werker',
        'category': 'Производитель / Дистрибьютор',
        'about': 'Производитель моющихся, декоративных красок, грунтов и профессиональных растворов. Доставка по Кишиневу и Молдове.',
        'features': 'производитель, опт',
        'assortment': 'моющиеся, декоративные краски, грунты',
        'address': 'г. Кишинев',
        'phones': '',
        'emails': '',
        'site': 'werker.md',
    },
    {
        'name': 'VOPSELE.MD',
        'category': 'Импортер / Дистрибьютор',
        'about': 'Дистрибьютор красок и лаков из ЕС в Республике Молдова. Эмали по металлу и дереву.',
        'features': 'импортёр, опт',
        'assortment': 'эмали по металлу и дереву, ЛКМ из ЕС',
        'address': 'г. Кишинев',
        'phones': '+373 760 016-70',
        'emails': '',
        'site': 'lacheie.md',
    },
    {
        'name': 'SuperStefan (Tikkurila)',
        'category': 'Официальный дистрибьютор (Tikkurila)',
        'about': 'Магазин Tikkurila в Кишиневе. Высокопрочные моющиеся краски для стен и потолков, атмосферостойкие фасадные, адгезионные грунтовки.',
        'features': 'официальный дилер, розница',
        'assortment': 'Tikkurila, моющиеся, фасадные краски, грунтовки',
        'address': 'г. Кишинев',
        'phones': '',
        'emails': '',
        'site': 'superstefan.md',
    },
    {
        'name': 'HomeDecor (акриловые краски из Италии)',
        'category': 'Импортер / Дистрибьютор',
        'about': 'Импортёр итальянских акриловых красок. 24 000 оттенков. Кишинев, ул. Василе Александри 106.',
        'features': 'импортёр, опт, розница',
        'assortment': 'акриловые краски из Италии',
        'address': 'г. Кишинев, ул. Василе Александри 106',
        'phones': '',
        'emails': '',
        'site': 'homedecor.md',
    },
    {
        'name': 'RichColors',
        'category': 'Розничные сети и магазины',
        'about': 'Премиальные краски для интерьера и экстерьера. Проверенные бренды, высокая укрывистость. Доставка по Кишиневу.',
        'features': 'розница',
        'assortment': 'премиальные интерьерные/фасадные краски',
        'address': 'г. Кишинев',
        'phones': '+373 79',
        'emails': '',
        'site': 'richcolors.md',
    },
    {
        'name': 'VERIX-GRUP',
        'category': 'Импортер / Дистрибьютор',
        'about': 'Один из крупнейших импортёров стройматериалов: керамическая плитка, ламинат, линолеум, смесители, аксессуары.',
        'features': 'импортёр, опт',
        'assortment': 'стройматериалы, ЛКМ',
        'address': 'Молдова (Кишинев, Бельцы)',
        'phones': '',
        'emails': '',
        'site': 'verix.md',
    },
    {
        'name': 'ICA (Prolomilux SRL)',
        'category': 'Дистрибьютор',
        'about': 'Оптовая и розничная торговля лаками, красками, морилками. Кишинев.',
        'features': 'опт, розница',
        'assortment': 'лаки, краски, морилки',
        'address': 'г. Кишинев',
        'phones': '',
        'emails': '',
        'site': 'kompass.com',
    },
    {
        'name': 'Lizarin Srl',
        'category': 'Дистрибьютор',
        'about': 'Оптовая продажа красок на водной и масляной основе, покрытий. Кишинев.',
        'features': 'опт',
        'assortment': 'водные/масляные краски, покрытия',
        'address': 'г. Кишинев (MD2002)',
        'phones': '',
        'emails': '',
        'site': 'kompass.com',
    },
    {
        'name': 'Ecolor',
        'category': 'Дистрибьютор',
        'about': 'Оптовая продажа древесины, стройматериалов и сантехнического оборудования.',
        'features': 'опт',
        'assortment': 'стройматериалы, ЛКМ',
        'address': 'г. Кишинев',
        'phones': '',
        'emails': '',
        'site': 'kr.kompass.com',
    },
    {
        'name': 'PoVladar (Бельцы)',
        'category': 'Розничные сети и магазины',
        'about': 'Интернет-магазин красок в Бельцах. Краски для всех видов поверхностей: матовые, глянцевые финиш.',
        'features': 'розница',
        'assortment': 'ЛКМ (общий ассортимент)',
        'address': 'г. Бельцы',
        'phones': '',
        'emails': '',
        'site': 'povladar.md',
    },
    {
        'name': 'StroyLux (Бельцы)',
        'category': 'Розничные сети и магазины',
        'about': 'Интернет-магазин стройматериалов в Молдове (Бельцы). Отделочные материалы, краски.',
        'features': 'розница',
        'assortment': 'отделочные материалы, краски',
        'address': 'г. Бельцы',
        'phones': '',
        'emails': '',
        'site': 'stroylux.md',
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

# Молдова — единственная компания в строке 4
sample_row = None
for r in range(4, ws.max_row + 1):
    if ws.cell(r, 2).value == 'Молдова':
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
        'Молдова',
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
    marker = ' ← +' if c == 'Молдова' else ''
    print(f"  {c}: {cnt} компаний{marker}")
print(f"Всего: {sum(Counter(countries).values())} компаний")
