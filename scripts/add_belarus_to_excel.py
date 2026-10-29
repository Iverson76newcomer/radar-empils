#!/usr/bin/env python3
"""
Фильтрация и добавление белорусских компаний в Excel.
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from copy import copy

SRC_EXCEL = '/home/z/my-project/download/Экспорт_потенциальные_дистрибьюторы_с_Казахстаном.xlsx'
OUT = '/home/z/my-project/download/Экспорт_потенциальные_дистрибьюторы_с_Беларусью.xlsx'


CLEAN_COMPANIES = [
    {
        'name': 'Эталон Фарб (OPTIMA)',
        'category': 'Производитель',
        'about': 'Производитель акриловых красок для стен и потолков. Торговая марка OPTIMA.',
        'features': 'производитель, опт',
        'assortment': 'акриловые краски OPTIMA',
        'address': 'Беларусь',
        'phones': '',
        'emails': '',
        'site': 'etalonfarb.by',
    },
    {
        'name': 'Tikkurila-shop.by (дистрибьютор Tikkurila)',
        'category': 'Официальный дистрибьютор (Tikkurila)',
        'about': 'Дистрибьютор лаков и красок Tikkurila в Беларуси. Декоративные покрытия, интерьерные и фасадные краски.',
        'features': 'официальный дилер, опт, розница',
        'assortment': 'Tikkurila, декоративные покрытия',
        'address': 'г. Минск, оф. 1028',
        'phones': '+375 29 152 55 55',
        'emails': 'info@tikkurila-shop.by',
        'site': 'tikkurila-shop.by',
    },
    {
        'name': 'Limex (ООО Колоритон)',
        'category': 'Дистрибьютор',
        'about': 'Сеть дилеров по Беларуси: Брестская, Гомельская, Гродненская, Минская области. Также работает в России (Москва).',
        'features': 'дистрибьютор, опт',
        'assortment': 'ЛКМ (общий ассортимент)',
        'address': 'Беларусь (Брест, Гомель, Гродно, Минск) + Россия',
        'phones': '+375 919 765-40-05, +7 495 151-51-84',
        'emails': '',
        'site': 'limex.by',
    },
    {
        'name': 'VseKraski (всекраски.by)',
        'category': 'Розничные сети и магазины',
        'about': 'Магазин дизайнерских красок и декоративных штукатурок в Минске. Интерьерные, фасадные краски, грунты, праймеры.',
        'features': 'розница',
        'assortment': 'интерьерные, фасадные краски, грунты, декоративные штукатурки',
        'address': 'г. Минск',
        'phones': '',
        'emails': '',
        'site': 'vsekraski.by',
    },
    {
        'name': 'MAV (MAV.PAINT CENTER)',
        'category': 'Производитель',
        'about': 'Семейная компания. Производство красок на водной и сольвентной основе, порошковых покрытий, аэрозольных красок, продуктов для DIY-рынка. Порошковые краски AMIKA.',
        'features': 'производитель, опт',
        'assortment': 'водные/сольвентные краски, порошковые AMIKA, аэрозоли',
        'address': 'Минская обл., г. Дзержинск, 222720',
        'phones': '+375 1716 26 000, +375 29 660 10 39',
        'emails': 'amika@mav.by',
        'site': 'mav.by',
    },
    {
        'name': 'Belhim',
        'category': 'Импортер / Дистрибьютор',
        'about': 'Основан в 1944 году в Минске. Крупнейший поставщик химической и пищевой продукции на рынок Республики Беларусь.',
        'features': 'импортёр, опт',
        'assortment': 'химическая продукция, сырьё для ЛКМ',
        'address': 'г. Минск',
        'phones': '',
        'emails': '',
        'site': 'belhim.com',
    },
    {
        'name': 'Lakokraska JSC',
        'category': 'Производитель',
        'about': 'Крупнейший белорусский производитель лакокрасочных материалов. Ассортимент — более 2000 наименований. Резидент СЭЗ Гродноинвест.',
        'features': 'производитель, опт',
        'assortment': 'ЛКМ (более 2000 наименований)',
        'address': 'г. Гродно',
        'phones': '',
        'emails': '',
        'site': 'grodnoinvest.by',
    },
    {
        'name': 'ООО Colorfaktori (ARMOR Chemie)',
        'category': 'Официальный дистрибьютор (ARMOR Chemie)',
        'about': 'Генеральный дистрибьютор ARMOR Chemie в Беларуси.',
        'features': 'официальный дилер, опт',
        'assortment': 'ARMOR Chemie, ЛКМ',
        'address': 'Беларусь',
        'phones': '',
        'emails': 'info@armor-chemie.com',
        'site': 'armor-chemie.com',
    },
    {
        'name': 'ОДО Баскет Групп (ЛКМ.бел)',
        'category': 'Дистрибьютор',
        'about': 'Краски и клеи с доставкой по Минску и Беларуси. ЛКМ.бел — онлайн-направление компании.',
        'features': 'дистрибьютор, опт, розница',
        'assortment': 'краски, клеи',
        'address': 'г. Минск',
        'phones': '',
        'emails': '',
        'site': 'brest.deal.by',
    },
    {
        'name': 'ООО Колорика',
        'category': 'Импортер / Дистрибьютор',
        'about': 'Официальный импортёр грунт-эмалей. Доставка по Беларуси. Опт, все документы.',
        'features': 'импортёр, опт',
        'assortment': 'грунт-эмали алкидные',
        'address': 'г. Минск',
        'phones': '+375 33 992-64-65, +375 29 325-74-31',
        'emails': 'zakaz@kolorika.by',
        'site': 'kolorika.by',
    },
    {
        'name': 'ООО ГринБи (s-w.by)',
        'category': 'Дистрибьютор',
        'about': 'Алкидные грунты, эмали и грунт-эмали в Минске. Краски по металлу.',
        'features': 'дистрибьютор, опт',
        'assortment': 'алкидные грунты, эмали, грунт-эмали, краски по металлу',
        'address': 'г. Минск',
        'phones': '+375 29 629-25-57',
        'emails': '',
        'site': 's-w.by',
    },
    {
        'name': 'SDM-Kraska',
        'category': 'Производитель',
        'about': 'Производство лакокрасочной продукции в Беларуси. Грунт-эмаль 3 в 1 для антикоррозийной защиты металлических поверхностей.',
        'features': 'производитель, опт',
        'assortment': 'грунт-эмаль 3 в 1, антикоррозийные ЛКМ',
        'address': 'Беларусь',
        'phones': '+375 29 305-49',
        'emails': '',
        'site': 'sdm-kraska.by',
    },
    {
        'name': 'DFarb',
        'category': 'Производитель',
        'about': 'Производитель эмалей в Минске. Большой выбор, доступные цены.',
        'features': 'производитель, опт',
        'assortment': 'эмали',
        'address': 'г. Минск',
        'phones': '',
        'emails': '',
        'site': 'dfarb.by',
    },
    {
        'name': 'Izima Group',
        'category': 'Импортер / Дистрибьютор',
        'about': 'Поставка цветных металлов, сырья для лакокрасочной промышленности и оборудования.',
        'features': 'импортёр, опт',
        'assortment': 'сырьё для ЛКМ, цветные металлы',
        'address': 'г. Минск, 10А',
        'phones': '+375 44 715-08-53',
        'emails': '',
        'site': 'izima-group.com',
    },
    {
        'name': 'Строймаркет KSK',
        'category': 'Розничные сети и магазины',
        'about': 'Интернет-магазин стройматериалов в Гомеле. Колеровка красок, гибка металла.',
        'features': 'розница, колеровка',
        'assortment': 'стройматериалы, ЛКМ',
        'address': 'г. Гомель',
        'phones': '',
        'emails': '',
        'site': 'ksk.by',
    },
    {
        'name': 'Сеть «Гермес» (BSP-Stroy)',
        'category': 'Розничные сети и магазины',
        'about': 'Сеть магазинов стройматериалов в Гомеле. Акриловые и другие краски.',
        'features': 'розница',
        'assortment': 'акриловые краски, ЛКМ',
        'address': 'г. Гомель',
        'phones': '+375 44 786-58-00',
        'emails': '',
        'site': 'bsp-stroy.by',
    },
    {
        'name': 'Сеть «Материк» (DIY.by)',
        'category': 'Розничные сети и магазины',
        'about': 'Сеть строительных гипермаркетов в Бресте. Кровельные краски, ЛКМ.',
        'features': 'розница',
        'assortment': 'кровельные краски, ЛКМ',
        'address': 'г. Брест',
        'phones': '',
        'emails': '',
        'site': 'diy.by',
    },
    {
        'name': 'Техноизолстрой',
        'category': 'Розничные сети и магазины',
        'about': 'Магазин красок в Гомеле. Доставка по городу, консультации специалистов.',
        'features': 'розница',
        'assortment': 'ЛКМ',
        'address': 'г. Гомель',
        'phones': '+375 44 701-43-61',
        'emails': '',
        'site': 'tehnoizol.by',
    },
    {
        'name': 'Bild-Shop (Гомель)',
        'category': 'Розничные сети и магазины',
        'about': 'Интернет-магазин ЛКМ в Гомеле. Краски Diamant, ВД-АК-2180.',
        'features': 'розница',
        'assortment': 'Diamant, ВД-АК-2180',
        'address': 'г. Гомель',
        'phones': '',
        'emails': '',
        'site': 'gomel.bild-shop.by',
    },
    {
        'name': 'LunaStroy',
        'category': 'Розничные сети и магазины',
        'about': 'Магазин отделочных материалов. Цветогамма-акрил 201 для интерьерной окраски.',
        'features': 'розница',
        'assortment': 'акриловые краски (Цветогамма-акрил)',
        'address': 'Беларусь',
        'phones': '',
        'emails': '',
        'site': 'lunastroy.by',
    },
    {
        'name': 'Palas',
        'category': 'Розничные сети и магазины',
        'about': 'Магазин отделочных материалов в Гомеле. Краски, стройматериалы. Оверлок ковровых покрытий.',
        'features': 'розница',
        'assortment': 'краски, стройматериалы',
        'address': 'г. Гомель, ул. Советская 97/2',
        'phones': '',
        'emails': '',
        'site': 'palas.by',
    },
    {
        'name': 'Serpa',
        'category': 'Розничные сети и магазины',
        'about': 'Интернет-магазин стройматериалов в Бресте. Бесплатная доставка, система скидок. Работают частные и государственные строительные организации.',
        'features': 'розница, опт',
        'assortment': 'стройматериалы, ЛКМ',
        'address': 'г. Брест',
        'phones': '',
        'emails': '',
        'site': 'serpa.by',
    },
    {
        'name': 'KraskiTut (ООО Колорика — направление)',
        'category': 'Дистрибьютор',
        'about': 'Грунт-эмаль алкидная. Официальный импортёр. Отдел продаж.',
        'features': 'импортёр, опт',
        'assortment': 'грунт-эмаль алкидная',
        'address': 'Беларусь',
        'phones': '+375 33 992-64-65',
        'emails': '',
        'site': 'kraskitut.by',
    },
    {
        'name': 'ОМА (b2b.oma.by)',
        'category': 'Розничные сети и магазины',
        'about': 'Краски оптом в интернет-магазине. Оплата по безналу, доставка с собственного склада по всей Беларуси.',
        'features': 'опт, розница',
        'assortment': 'ЛКМ (общий ассортимент)',
        'address': 'Беларусь',
        'phones': '',
        'emails': '',
        'site': 'b2b.oma.by',
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
    if ws.cell(r, 2).value == 'Беларусь':
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
        'Беларусь',
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
        ws.column_dimensions[get_column_letter(i)].width = w

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
    marker = ' ← +' if c == 'Беларусь' else ''
    print(f"  {c}: {cnt} компаний{marker}")
print(f"Всего: {sum(Counter(countries).values())} компаний")
