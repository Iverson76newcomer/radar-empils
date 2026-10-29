#!/usr/bin/env python3
"""
Модифицировать PPTX пользователя:
- Добавить строку с Ираном в таблицу слайда 5 (между Туркменистаном и Сербией)
- Добавить строку с Казахстаном в таблицу слайда 5 (после Ирана)
- Обновить подзаголовок (10 стран вместо 9)
- Обновить нижний блок (добавить Иран и Казахстан в плашки)
"""

import shutil
from copy import deepcopy
from pptx import Presentation
from pptx.util import Pt, Emu
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn

SRC = "/home/z/my-project/upload/Kazbuild_2026_Эмпилс_Стратегия.pptx"
DST = "/home/z/my-project/download/Kazbuild_2026_Эмпилс_Стратегия_v4.pptx"
shutil.copy(SRC, DST)

prs = Presentation(DST)
slide = prs.slides[4]  # slide 5 (0-indexed)

# Найдём таблицу на слайде
table_shape = None
for shape in slide.shapes:
    if shape.has_table:
        table_shape = shape
        break

if not table_shape:
    print("ERROR: Таблица не найдена на слайде 5")
    exit(1)

table = table_shape.table
print(f"Найдена таблица: {len(table.rows)} строк, {len(table.columns)} колонок")
print("\nТекущие строки:")
for i, row in enumerate(table.rows):
    cells = [cell.text.strip()[:30] for cell in row.cells]
    print(f"  R{i}: {cells}")

# Нужно найти индекс строки "Туркменистан" и вставить после неё 2 новые строки (Иран, Казахстан) перед Сербией
turkmen_idx = None
serbia_idx = None
for i, row in enumerate(table.rows):
    if "Туркменистан" in row.cells[0].text:
        turkmen_idx = i
    if "Сербия" in row.cells[0].text:
        serbia_idx = i

print(f"\nИндекс Туркменистана: {turkmen_idx}")
print(f"Индекс Сербии: {serbia_idx}")

if turkmen_idx is None or serbia_idx is None:
    print("ERROR: не найдены строки")
    exit(1)

# python-pptx не умеет напрямую вставлять строки, нужно работать с XML
# Скопируем XML строки Туркменистана как шаблон и вставим 2 новые после неё

tbl = table._tbl
rows_xml = tbl.findall(qn('a:tr'))
print(f"\nНайдено XML строк: {len(rows_xml)}")

# Шаблон — строка Туркменистана
template_tr = rows_xml[turkmen_idx]

# Новые данные для строк
new_rows_data = [
    # (страна, дефицит, рост, потребл, рекомендация)
    ("🇮🇷 Иран", "17%", "2,85–8% / год", "6,45 кг/чел",
     "Большой потенциал, но санкции. Кейс «Алдан.Экспорт» → Иран работает САНКЦИИ"),
    ("🇰🇿 Казахстан", "25–30%", "5–7% / год", "7,0 кг/чел",
     "Ключевой текущий рынок Эмпилс. 100+ дистрибьюторов, 6/14 обл. Цель — углубление В РАБОТЕ"),
]

# Вставляем новые строки после Туркменистана (перед Сербией)
# Создаём копии template_tr и подменяем текст
for i, (country, deficit, growth, percap, reco) in enumerate(new_rows_data):
    new_tr = deepcopy(template_tr)
    # Получаем ячейки (a:tc) в новой строке
    tcs = new_tr.findall(qn('a:tc'))
    print(f"\nСоздаю строку {i+1}: {country}")
    print(f"  Ячеек в строке: {len(tcs)}")

    # Очищаем текст в каждой ячейке и записываем новый
    for tc_idx, tc in enumerate(tcs):
        # Находим все a:p (параграфы) в ячейке
        # Берём только первый параграф, остальные удаляем
        ps = tc.findall(qn('a:p'))
        if not ps:
            continue

        # Оставляем только первый параграф
        for p in ps[1:]:
            tc.remove(p)

        first_p = ps[0]
        # Удаляем все a:r (runs) кроме первого
        rs = first_p.findall(qn('a:r'))
        if not rs:
            # Создаём новый run
            r_elem = first_p.makeelement(qn('a:r'), {})
            rPr = r_elem.makeelement(qn('a:rPr'), {})
            r_elem.append(rPr)
            t_elem = r_elem.makeelement(qn('a:t'), {})
            r_elem.append(t_elem)
            first_p.append(r_elem)
            rs = [r_elem]
        else:
            # Удаляем все runs кроме первого
            for r in rs[1:]:
                first_p.remove(r)
            rs = rs[:1]

        # Меняем текст в первом run
        t_elem = rs[0].find(qn('a:t'))
        if t_elem is None:
            t_elem = rs[0].makeelement(qn('a:t'), {})
            rs[0].append(t_elem)

        # Записываем текст в зависимости от индекса ячейки
        if tc_idx == 0:
            t_elem.text = country
        elif tc_idx == 1:
            t_elem.text = deficit
        elif tc_idx == 2:
            t_elem.text = growth
        elif tc_idx == 3:
            t_elem.text = percap
        elif tc_idx == 4:
            t_elem.text = reco

    # Вставляем новую строку после строки Туркменистана
    # Находим текущую строку Туркменистана (или Ирана для второй вставки)
    current_rows = tbl.findall(qn('a:tr'))
    insert_after_idx = None
    target_text_1 = "Туркменистан" if i == 0 else "Иран"
    target_text_2 = "Иран" if i == 0 else None  # для второй вставки

    for j, tr in enumerate(current_rows):
        # Получаем текст первой ячейки — объединяем все runs
        tcs_check = tr.findall(qn('a:tc'))
        if not tcs_check:
            continue
        first_tc = tcs_check[0]
        text = ""
        # Все параграфы, все runs
        for p in first_tc.findall(qn('a:p')):
            for r in p.findall(qn('a:r')):
                t = r.find(qn('a:t'))
                if t is not None and t.text:
                    text += t.text

        # Простая проверка вхождения подстроки
        if i == 0 and "Туркменистан" in text:
            insert_after_idx = j
            print(f"  Найден Туркменистан в строке {j}, текст: '{text[:50]}'")
            break
        elif i == 1 and "Иран" in text:
            insert_after_idx = j
            print(f"  Найден Иран в строке {j}, текст: '{text[:50]}'")
            break

    if insert_after_idx is not None:
        ref_tr = current_rows[insert_after_idx]
        ref_tr.addnext(new_tr)
        print(f"  → вставлена после строки {insert_after_idx}")
    else:
        # Fallback: если не нашли по тексту, используем индекс
        # Для Ирана (i=0): после Туркменистана (turkmen_idx=6)
        # Для Казахстана (i=1): после Ирана (turkmen_idx + 1 = 7, но строка уже вставлена)
        fallback_idx = turkmen_idx + i + 1  # 6 + 0 + 1 = 7 для Ирана; 6 + 1 + 1 = 8 для Казахстана
        current_rows = tbl.findall(qn('a:tr'))
        if fallback_idx < len(current_rows):
            ref_tr = current_rows[fallback_idx]
            # Вставляем ПЕРЕД этой строкой (используем addprevious)
            ref_tr.addprevious(new_tr)
            print(f"  → вставлена перед строкой {fallback_idx} (fallback)")

# Сохраняем
prs.save(DST)
print(f"\n✅ Файл сохранён: {DST}")
