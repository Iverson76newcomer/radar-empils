#!/usr/bin/env python3
"""
Фильтрация и добавление узбекских компаний в Excel.
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from copy import copy

SRC_EXCEL = '/home/z/my-project/download/Экспорт_потенциальные_дистрибьюторы_с_Грузией.xlsx'
OUT = '/home/z/my-project/download/Экспорт_потенциальные_дистрибьюторы_с_Узбекистаном.xlsx'


# === РУЧНАЯ ЧИСТКА — реальные узбекские компании ===
CLEAN_COMPANIES = [
    {
        'name': 'Rangli Bo\'yoq',
        'category': 'Производитель / Дистрибьютор',
        'about': 'Широкий ассортимент ЛКМ: краски, лаки, эмали, грунтовки, декоративные штукатурки.',
        'features': 'производитель, опт, розница',
        'assortment': 'краски, лаки, эмали, грунтовки, декоративные штукатурки',
        'address': 'г. Ташкент',
        'phones': '',
        'emails': '',
        'site': 'ranglibuyoq.uz',
    },
    {
        'name': 'Sobsan / Tytan (Узбекистан)',
        'category': 'Производитель',
        'about': 'Производитель ЛКМ под марками Sobsan и Tytan в Узбекистане. Строительные краски и лаки.',
        'features': 'производитель',
        'assortment': 'Sobsan, Tytan, строительные краски, лаки, эмульсии',
        'address': 'г. Ташкент',
        'phones': '',
        'emails': '',
        'site': 'goldenpages.uz',
    },
    {
        'name': 'ООО Stroy-Kontinent',
        'category': 'Дистрибьютор',
        'about': 'Широкий ассортимент лакокрасочных материалов и растворителей для строительных, ремонтных и отделочных работ.',
        'features': 'опт',
        'assortment': 'ЛКМ, растворители',
        'address': 'г. Ташкент',
        'phones': '',
        'emails': '',
        'site': 'stroykontinent.uz',
    },
    {
        'name': 'KZMC (оптовые поставки красок)',
        'category': 'Дистрибьютор',
        'about': 'Оптовые поставки фасадных, интерьерных, антикоррозионных и специальных красок для металла, бетона и дерева по всему Узбекистану.',
        'features': 'опт',
        'assortment': 'фасадные, интерьерные, антикоррозийные, специальные краски',
        'address': 'г. Ташкент',
        'phones': '',
        'emails': '',
        'site': 'kzmc.uz',
    },
    {
        'name': 'ALEM (завод красок)',
        'category': 'Производитель',
        'about': 'Завод ALEM: водно-дисперсионные, масляные, нитро-краски. Оптовые цены, доставка по Ташкенту и Узбекистану.',
        'features': 'производитель, опт',
        'assortment': 'водно-дисперсионные, масляные, нитро-краски',
        'address': 'г. Ташкент',
        'phones': '+998 95 022-01-01',
        'emails': 'info@alem.uz',
        'site': 'alem.uz',
    },
    {
        'name': 'STROY MIX IMPORT',
        'category': 'Импортер / Дистрибьютор',
        'about': 'Оптовая поставка водоэмульсионных красок в Ташкенте.',
        'features': 'импортёр, опт',
        'assortment': 'водоэмульсионные краски',
        'address': 'г. Ташкент',
        'phones': '+998 99 828-55-67, +998 77 052-84-23',
        'emails': '',
        'site': 'stroyka.uz',
    },
    {
        'name': 'KRASKO (интернет-магазин)',
        'category': 'Розничные сети и магазины',
        'about': 'Интернет-магазин красок, штукатурок, лаков и инструментов. Колеровка более 10 000 оттенков.',
        'features': 'розница, колеровка',
        'assortment': 'краски, штукатурки, лаки, инструменты',
        'address': 'г. Ташкент',
        'phones': '',
        'emails': '',
        'site': 'krasko.uz',
    },
    {
        'name': 'Silkcoat Uzbekistan',
        'category': 'Официальный дистрибьютор (Silkcoat)',
        'about': 'Международный производитель декоративных и лакокрасочных покрытий с 1993 года. Интерьерные, фасадные и эффектные краски.',
        'features': 'производитель, опт',
        'assortment': 'Silkcoat, интерьерные, фасадные, эффектные краски',
        'address': 'г. Ташкент',
        'phones': '+998 90 983-48',
        'emails': '',
        'site': 'silkcoat.uz',
    },
    {
        'name': 'Uz-DongJu Paint Co. (ROYAL PAINT)',
        'category': 'Производитель',
        'about': 'Завод лакокрасочных материалов. Более 25 лет на рынке. Производство и оптовая продажа.',
        'features': 'производитель, опт',
        'assortment': 'лакокрасочные материалы ROYAL PAINT',
        'address': 'г. Ташкент',
        'phones': '',
        'emails': '',
        'site': 'uzdongju.uz',
    },
    {
        'name': 'Tashkent Chemicals',
        'category': 'Импортер / Дистрибьютор',
        'about': 'Поставщик сырья для ЛКМ и строительной химии из Турции и соседних стран.',
        'features': 'импортёр, опт',
        'assortment': 'сырьё для ЛКМ, строительная химия',
        'address': 'г. Ташкент',
        'phones': '',
        'emails': '',
        'site': 'tashkentchemicals.uz',
    },
    {
        'name': 'Luna Tones (турецкий экспортёр)',
        'category': 'Импортер / Дистрибьютор',
        'about': 'Турецкий экспортёр красок. Поставки в Узбекистан: интерьерные, фасадные, потолочные краски, грунтовки. ISO 9001:2015, CE.',
        'features': 'импортёр, опт',
        'assortment': 'интерьерные, фасадные, потолочные краски, грунтовки',
        'address': 'г. Ташкент',
        'phones': '',
        'emails': '',
        'site': 'lunatones.com',
    },
    {
        'name': 'RuslanBaraka (магазин стройматериалов)',
        'category': 'Розничные сети и магазины',
        'about': 'Надёжный поставщик лакокрасочной продукции и строительных материалов в больших объёмах.',
        'features': 'опт, розница',
        'assortment': 'ЛКМ, стройматериалы',
        'address': 'г. Ташкент',
        'phones': '+998 99 769-13-39',
        'emails': '',
        'site': 'ruslanbaraka.uz',
    },
    {
        'name': 'Leonis',
        'category': 'Дистрибьютор',
        'about': 'Оптовая продажа ЛКМ: краски, лаки, грунтовки, эмали для внутренних и наружных работ.',
        'features': 'опт',
        'assortment': 'краски, лаки, грунтовки, эмали',
        'address': 'г. Ташкент',
        'phones': '',
        'emails': '',
        'site': 'leonis.uz',
    },
    {
        'name': 'Spec-Emal (Элконтермм)',
        'category': 'Производитель',
        'about': 'Химстойкие эмали, химостойкие грунты, лаки. Готовые покрытия для промышленных задач.',
        'features': 'производитель, опт',
        'assortment': 'химстойкие эмали, грунты, лаки',
        'address': 'г. Ташкент, Алмазарский р-н, ул. Джами 5',
        'phones': '+998 78 555-6-555',
        'emails': 'elkonterm@mail.ru',
        'site': 'spec-emal.uz',
    },
    {
        'name': 'ALEKS-TRADE ООО',
        'category': 'Дистрибьютор',
        'about': 'Дистрибьютор эмалей и лакокрасочных материалов в Ташкенте.',
        'features': 'опт',
        'assortment': 'эмали, ЛКМ',
        'address': 'г. Ташкент',
        'phones': '',
        'emails': '',
        'site': 'prom.uz',
    },
    {
        'name': 'PRIME STROY SHOP 777',
        'category': 'Розничные сети и магазины',
        'about': 'Магазин универсальных эмалей Uzcolor и лакокрасочных материалов в Ташкенте.',
        'features': 'розница',
        'assortment': 'Uzcolor, универсальные эмали',
        'address': 'г. Ташкент',
        'phones': '',
        'emails': '',
        'site': 'prom.uz',
    },
    {
        'name': 'Dunyabunya',
        'category': 'Дистрибьютор',
        'about': 'Поставщик грунт-эмали по ржавчине «3 в 1» алкидно-уретановой в Ташкенте.',
        'features': 'опт',
        'assortment': 'грунт-эмаль 3 в 1, алкидно-уретановые',
        'address': 'г. Ташкент',
        'phones': '',
        'emails': '',
        'site': 'stroyka.uz',
    },
    {
        'name': 'AISHAH PRO BUSINESS ООО',
        'category': 'Производитель',
        'about': 'Лакокрасочное производство в Узбекистане (по каталогу Yellowpages.uz).',
        'features': 'производитель',
        'assortment': 'ЛКМ',
        'address': 'г. Ташкент',
        'phones': '',
        'emails': '',
        'site': 'yellowpages.uz',
    },
    {
        'name': 'ALAZOR ООО',
        'category': 'Производитель',
        'about': 'Лакокрасочное производство в Узбекистане.',
        'features': 'производитель',
        'assortment': 'ЛКМ',
        'address': 'г. Ташкент',
        'phones': '',
        'emails': '',
        'site': 'yellowpages.uz',
    },
    {
        'name': 'A.Sh.A O\'',
        'category': 'Производитель',
        'about': 'Лакокрасочное производство в Узбекистане.',
        'features': 'производитель',
        'assortment': 'ЛКМ',
        'address': 'г. Ташкент',
        'phones': '',
        'emails': '',
        'site': 'yellowpages.uz',
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

# Строка-образец (Узбекистан)
sample_row = None
for r in range(4, ws.max_row + 1):
    if ws.cell(r, 2).value == 'Узбекистан':
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
        'Узбекистан',
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

# Итог
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
    marker = ' ← +' if c == 'Узбекистан' else ''
    print(f"  {c}: {cnt} компаний{marker}")
print(f"Всего: {sum(Counter(countries).values())} компаний")
