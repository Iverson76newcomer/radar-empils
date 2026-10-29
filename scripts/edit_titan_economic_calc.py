"""
Редактирование Titan economic calculation.docx через прямой XML-обход.
Заменяем цены TITAN: 700 → 860, 1 400 → 1 720, 2 900 → 3 220, 290 → 322,
экономия 5350 → 5 030, 8250 → 8 250, 2900 → 3 220.
Текст вступления и выводы остаются как есть.
"""
import shutil
import re
from docx import Document
from docx.oxml.ns import qn
from lxml import etree

SRC = '/home/z/my-project/upload/Titan economic calculation.docx'
OUT = '/home/z/my-project/download/Titan_economic_calculation.docx'

shutil.copy(SRC, OUT)

doc = Document(OUT)

# Карта замен: точные подстроки
# ВАЖНО: длинные/уникальные строки заменяем раньше коротких,
# чтобы избежать двойной замены.
replacements = [
    # Вступительный текст
    ('TITAN 0,9 кг стоит ~700 ₽',  'TITAN 0,9 кг стоит ~860 ₽'),
    # Главный вывод в плашке: «5350 ₽ (8250 ₽ – 2900 ₽)» → «5 030 ₽ (8 250 ₽ – 3 220 ₽)»
    # Текст мог быть разбит; возьмём как есть.
    ('5350 ₽',       '5 030 ₽'),
    ('8250 ₽',       '8 250 ₽'),
    ('2900 ₽',       '3 220 ₽'),
    # Таблица — значения TITAN
    # Сначала заменяем конкретные ячейки с большой вероятностью уникальности
    # Цена 700 — нужно заменить только TITAN, не задев ПФ-115 (тоже 250 — не заденет, 700 только у TITAN)
    # 1 400 — только TITAN (материалы)
    # 2 900 — только TITAN (итог 1 перекраски)
    # 290 — только TITAN (ССВ/год)
    # Также обновим 5 350 (если где-то ещё встречается)
    ('5 350',        '5 030'),
    ('5350',         '5 030'),
    # Цены в таблице — заменяем «700» только в контексте цены TITAN
    # «290» — тоже только TITAN. Других 700 и 290 в документе нет.
    ('700',           '860'),       # цена банки TITAN
    ('1 400',         '1 720'),     # материалы TITAN
    ('2 900',         '3 220'),     # итог 1 перекраски TITAN + ССВ 10 лет TITAN
    ('290',           '322'),       # ССВ/год TITAN
]

# Подсчёт
counts = {}

def replace_text_in_element(elem, old, new):
    """Заменяем текст во всех <w:t> элементах внутри elem.
    Текст может быть разбит по нескольким w:t — собираем, заменяем, перераспределяем."""
    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    t_nodes = elem.findall('.//w:t', ns)
    if not t_nodes:
        return 0
    # Собираем полный текст
    full_text = ''.join(t.text or '' for t in t_nodes)
    if old not in full_text:
        return 0
    count = full_text.count(old)
    new_text = full_text.replace(old, new)
    # Простое перераспределение: весь текст в первый w:t, остальные пустыми
    # Это сохранит форматирование первого run-а, но потеряет внутреннее — для ячеек таблицы
    # с одним форматом это безопасно.
    if t_nodes:
        t_nodes[0].text = new_text
        for t in t_nodes[1:]:
            t.text = ''
    return count

# Применяем замены ко всем параграфам и ячейкам
for old, new in replacements:
    counts[old] = 0
    # Все параграфы верхнего уровня
    for p in doc.paragraphs:
        counts[old] += replace_text_in_element(p._p, old, new)
    # Все таблицы (включая вложенные ячейки)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                # В каждой ячейке — свои параграфы
                for p in cell.paragraphs:
                    counts[old] += replace_text_in_element(p._p, old, new)
                # Также могут быть другие таблицы внутри ячейки (вложенные)
                for nested_tbl in cell.tables:
                    for nested_row in nested_tbl.rows:
                        for nested_cell in nested_row.cells:
                            for p in nested_cell.paragraphs:
                                counts[old] += replace_text_in_element(p._p, old, new)

print('Замены:')
for old, new in replacements:
    c = counts.get(old, 0)
    print(f'  {old!r} → {new!r}: {c} замен')

doc.save(OUT)
print(f'\nOK saved: {OUT}')
import os
print(f'Size: {os.path.getsize(OUT)} bytes')
