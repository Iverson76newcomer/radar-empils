#!/usr/bin/env python3
"""
Фильтрация и добавление грузинских компаний в Excel.
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from copy import copy

SRC_EXCEL = '/home/z/my-project/download/Экспорт_потенциальные_дистрибьюторы_с_Арменией.xlsx'
OUT = '/home/z/my-project/download/Экспорт_потенциальные_дистрибьюторы_с_Грузией.xlsx'


# === РУЧНАЯ ЧИСТКА — оставляем реальные грузинские компании ===
CLEAN_COMPANIES = [
    {
        'name': 'КАПАРОЛ ДЖОРДЖИЯ (Caparol Center Tbilisi)',
        'category': 'Официальный дистрибьютор (Caparol)',
        'about': 'Официальный дилер Caparol в Грузии. Центральный магазин в Тбилиси.',
        'features': 'официальный дилер, розница',
        'assortment': 'Caparol, краски, штукатурки',
        'address': 'г. Тбилиси, ул. Панаскертел-Цицишвили 1а (0194)',
        'phones': '+995 598 047 777',
        'emails': 'caparolcenter@caparol.ge',
        'site': 'caparol.ge',
    },
    {
        'name': 'Flügger (через GTS Group)',
        'category': 'Официальный дистрибьютор (Flügger)',
        'about': 'Официальный дистрибьютор датских экологичных красок Flügger в Грузии.',
        'features': 'официальный дилер',
        'assortment': 'Flügger, экологичные краски',
        'address': 'г. Тбилиси',
        'phones': '',
        'emails': '',
        'site': 'gts-group.ge',
    },
    {
        'name': 'LTD Danika (Tikkurila)',
        'category': 'Официальный дистрибьютор (Tikkurila)',
        'about': 'Официальный дистрибьютор Tikkurila в Грузии. Покрытия по металлу, напольные системы.',
        'features': 'официальный дилер',
        'assortment': 'Tikkurila, покрытия по металлу, напольные системы',
        'address': 'г. Тбилиси, ул. Саакадзе, I линия N1',
        'phones': '+995 322 38',
        'emails': '',
        'site': 'tikkurila.com',
    },
    {
        'name': 'BOERO Georgia (Art House)',
        'category': 'Импортер / Дистрибьютор',
        'about': 'Импортёр лакокрасочных материалов BOERO. Офисы в Батуми и Тбилиси.',
        'features': 'импортёр, розница',
        'assortment': 'BOERO, декоративные краски',
        'address': 'г. Батуми; г. Тбилиси, ул. Цацхвели 14',
        'phones': '+995 550 006 999',
        'emails': 'arthousebatumi2020@gmail.com',
        'site': 'boero.ge',
    },
    {
        'name': 'Baril Coatings (Baril Caucasia)',
        'category': 'Дистрибьютор',
        'about': 'Дистрибьютор покрытий Baril Coatings в Грузии и на Кавказе.',
        'features': 'дистрибьютор',
        'assortment': 'Baril, Candi Paints, Biocoat',
        'address': 'г. Тбилиси',
        'phones': '+995 455 07 44',
        'emails': 'info@candipaints.com, info@barilcoatings.nl',
        'site': 'barilcoatings.com',
    },
    {
        'name': 'RVM International',
        'category': 'Импортер / Дистрибьютор',
        'about': 'Штаб-квартира в Тбилиси. Организация дистрибуции и оптовой продажи товаров на Кавказе.',
        'features': 'импортёр, опт',
        'assortment': 'ЛКМ (общий ассортимент)',
        'address': 'г. Тбилиси',
        'phones': '+995 591 817 388',
        'emails': '',
        'site': 'rvm-international.com',
    },
    {
        'name': 'Chemicals Coating Trading',
        'category': 'Дистрибьютор',
        'about': 'Торговля и дистрибуция лакокрасочных покрытий. Работает в Тбилиси, Рустави, Гори, Телави, Батуми, Кутаиси.',
        'features': 'дистрибьютор, опт',
        'assortment': 'ЛКМ (общий ассортимент)',
        'address': 'г. Тбилиси (филиалы: Рустави, Гори, Телави, Батуми, Кутаиси)',
        'phones': '',
        'emails': '',
        'site': 'chemicalscoating.ge',
    },
    {
        'name': 'Green Lab (Kobra Paint)',
        'category': 'Официальный дистрибьютор (Kobra)',
        'about': 'Эксклюзивный продавец Kobra Paint в Грузии. Импортёр аэрозольных красок Kobra.',
        'features': 'импортёр, эксклюзивный дилер',
        'assortment': 'Kobra, аэрозольные краски',
        'address': 'г. Тбилиси',
        'phones': '',
        'emails': '',
        'site': 'greenlab.ge',
    },
    {
        'name': 'Almet Trading Georgia',
        'category': 'Импортер / Дистрибьютор',
        'about': 'Работает из Тбилиси с 2021 года. Снабжение, оптовые поставки, распределительный склад.',
        'features': 'импортёр, опт',
        'assortment': 'стройматериалы, металлопрокат',
        'address': 'г. Тбилиси',
        'phones': '',
        'emails': '',
        'site': 'almettrading.com',
    },
    {
        'name': 'Italica',
        'category': 'Производитель / Дистрибьютор',
        'about': 'Производитель декоративных штукатурок и красок. Венецианские и марокканские штукатурки. Бесплатные образцы.',
        'features': 'производитель, дистрибьютор',
        'assortment': 'декоративные штукатурки, краски',
        'address': 'г. Тбилиси',
        'phones': '',
        'emails': '',
        'site': 'italica.com.ge',
    },
    {
        'name': 'Domino (домино.com.ge)',
        'category': 'Розничные сети и магазины',
        'about': 'Магазин ЛКМ в Грузии. Tikkurila, Eskaro, Alpina, дисперсионные краски.',
        'features': 'розница',
        'assortment': 'Tikkurila, Eskaro, Alpina, дисперсионные краски',
        'address': 'г. Тбилиси',
        'phones': '',
        'emails': '',
        'site': 'domino.com.ge',
    },
    {
        'name': 'Nika (трейдер полимеров)',
        'category': 'Дистрибьютор',
        'about': 'Поставщик красителей и пигментных гранул для промышленности.',
        'features': 'опт',
        'assortment': 'красители, пигменты',
        'address': 'г. Тбилиси',
        'phones': '+995 32 243 56',
        'emails': '',
        'site': 'nika.ge',
    },
    {
        'name': 'ЭКОВУД',
        'category': 'Розничные сети и магазины',
        'about': 'Магазин ЛКМ в Тбилиси (Самгори, Авчала).',
        'features': 'розница',
        'assortment': 'ЛКМ (общий ассортимент)',
        'address': 'г. Тбилиси, Авчала, ул. Цкалсадени 21',
        'phones': '',
        'emails': '',
        'site': 'yell.ge',
    },
    {
        'name': 'ДЕРУФА ГРУЗИЯ',
        'category': 'Розничные сети и магазины',
        'about': 'Магазин ЛКМ в Батуми.',
        'features': 'розница',
        'assortment': 'Dufa, ЛКМ',
        'address': 'г. Батуми, ул. Пушкина 37',
        'phones': '+995 596 70 77 99',
        'emails': '',
        'site': 'yp.com.ge',
    },
    {
        'name': 'ЕВРОМАКС КОЛОРЗ',
        'category': 'Розничные сети и магазины',
        'about': 'Магазин ЛКМ в Кутаиси.',
        'features': 'розница',
        'assortment': 'ЛКМ (общий ассортимент)',
        'address': 'г. Кутаиси',
        'phones': '',
        'emails': '',
        'site': 'yp.com.ge',
    },
    {
        'name': 'Lumino',
        'category': 'Розничные сети и магазины',
        'about': 'Крупнейший в Грузии магазин художественных товаров. Краски, холсты, кисти.',
        'features': 'розница',
        'assortment': 'художественные краски, кисти',
        'address': 'г. Тбилиси',
        'phones': '',
        'emails': '',
        'site': 'lumino.ge',
    },
    {
        'name': 'Moedani.online',
        'category': 'Розничные сети и магазины',
        'about': 'Интернет-магазин ЛКМ и строительных материалов. Knauf, Caparol, Capadin.',
        'features': 'розница',
        'assortment': 'Knauf, Caparol, грунтовки',
        'address': 'г. Тбилиси',
        'phones': '',
        'emails': '',
        'site': 'moedani.online',
    },
    {
        'name': 'Matemar',
        'category': 'Дистрибьютор',
        'about': 'Оптовая продажа стройматериалов в Тбилиси. Краски и лакокрасочные материалы.',
        'features': 'опт',
        'assortment': 'стройматериалы, ЛКМ',
        'address': 'г. Тбилиси, ул. Григола Лорткипанидзе 112',
        'phones': '',
        'emails': '',
        'site': 'madloba.info',
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

# Строка-образец (Грузия)
sample_row = None
for r in range(4, ws.max_row + 1):
    if ws.cell(r, 2).value == 'Грузия':
        sample_row = r
        break

print(f"Строка-образец для форматирования: {sample_row}")

sample_fill = None
if sample_row:
    sample_cell = ws.cell(sample_row, 2)
    if sample_cell.fill and sample_cell.fill.fgColor:
        sample_fill = copy(sample_cell.fill)

if not sample_fill:
    sample_fill = PatternFill('solid', fgColor='DBEAFE')

thin = Side(style='thin', color='D1D5DB')
border = Border(left=thin, right=thin, top=thin, bottom=thin)

current_row = last_row + 1
added = 0

for c in CLEAN_COMPANIES:
    cells_data = [
        '',
        'Грузия',
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

# Ширина колонок
col_widths = {2: 14, 3: 28, 4: 36, 5: 50, 6: 35, 7: 32, 8: 32, 9: 24, 10: 32, 11: 22}
for ci, w in col_widths.items():
    cur = ws.column_dimensions[get_column_letter(ci)].width
    if not cur or cur < w:
        ws.column_dimensions[get_column_letter(ci)].width = w

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
    marker = ' ← +' if c == 'Грузия' else ''
    print(f"  {c}: {cnt} компаний{marker}")
print(f"Всего: {sum(Counter(countries).values())} компаний")

# Добавленные Грузии
print(f"\n=== Добавленные Грузии ===")
n = 0
for ri in range(last_row + 1, ws2.max_row + 1):
    if ws2.cell(ri, 2).value == 'Грузия':
        n += 1
        print(f"  {n}. {ws2.cell(ri, 4).value} | {ws2.cell(ri, 3).value} | {ws2.cell(ri, 11).value}")
