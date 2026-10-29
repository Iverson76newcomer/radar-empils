#!/usr/bin/env python3
"""
Модифицировать PPTX пользователя — добавить Иран и Казахстан в таблицу слайда 5.
Используем прямое XML-редактирование через deepcopy + замену текста в <a:t>.
"""

import shutil
from copy import deepcopy
from pptx import Presentation
from pptx.oxml.ns import qn

SRC = "/home/z/my-project/upload/Kazbuild_2026_Эмпилс_Стратегия.pptx"
DST = "/home/z/my-project/download/Kazbuild_2026_Эмпилс_Стратегия_v4.pptx"
shutil.copy(SRC, DST)

prs = Presentation(DST)
slide = prs.slides[4]

# Найдём таблицу
for shape in slide.shapes:
    if shape.has_table:
        table = shape.table
        break

tbl = table._tbl

# Получаем все строки
rows = tbl.findall(qn('a:tr'))
print(f"Текущих строк: {len(rows)}")

# Найдём индекс строки Туркменистана (по text ячейки)
def get_row_text(tr):
    """Получить текст первой ячейки строки (объединяя все a:t)."""
    tcs = tr.findall(qn('a:tc'))
    if not tcs:
        return ""
    # Все <a:t> элементы в первой ячейке
    t_elems = tcs[0].iter(qn('a:t'))
    return "".join(t.text or '' for t in t_elems)

turkmen_idx = None
for i, tr in enumerate(rows):
    text = get_row_text(tr)
    if "Туркменистан" in text:
        turkmen_idx = i
        print(f"Туркменистан найден в R{i}: '{text[:30]}'")
        break

if turkmen_idx is None:
    print("ERROR: Туркменистан не найден")
    exit(1)

template_tr = rows[turkmen_idx]

# Данные для новых строк
new_rows_data = [
    # (country, deficit, growth, percap, reco)
    ("🇮🇷 Иран", "17%", "2,85–8% / год", "6,45 кг/чел",
     "Большой потенциал, но санкции. Кейс «Алдан.Экспорт» → Иран работает  САНКЦИИ"),
    ("🇰🇿 Казахстан", "25–30%", "5–7% / год", "7,0 кг/чел",
     "Ключевой текущий рынок Эмпилс. 100+ дистрибьюторов, 6/14 обл. Цель — углубление  В РАБОТЕ"),
]

# Создаём новые строки
for i, (country, deficit, growth, percap, reco) in enumerate(new_rows_data):
    print(f"\nСоздаю строку {i+1}: {country}")

    # Глубокая копия строки Туркменистана
    new_tr = deepcopy(template_tr)

    # Получаем ячейки новой строки
    tcs = new_tr.findall(qn('a:tc'))
    print(f"  Ячеек в строке: {len(tcs)}")

    # Тексты для подмены в каждой ячейке
    cell_texts = [country, deficit, growth, percap, reco]

    for tc_idx, tc in enumerate(tcs):
        if tc_idx >= len(cell_texts):
            break

        new_text = cell_texts[tc_idx]
        print(f"  Ячейка {tc_idx}: '{new_text[:40]}'")

        # Найдём ВСЕ <a:t> элементы в ячейке
        t_elems = list(tc.iter(qn('a:t')))

        if t_elems:
            # Берём первый a:t и записываем в него текст
            first_t = t_elems[0]
            first_t.text = new_text
            print(f"    → первый a:t обновлён: '{new_text[:30]}'")

            # Удаляем текст из остальных a:t (или удаляем сами runs)
            for t in t_elems[1:]:
                # Найдём родительский <a:r> и удалим
                parent_r = t.getparent()  # это <a:r>
                if parent_r is not None:
                    parent_p = parent_r.getparent()  # это <a:p>
                    if parent_p is not None:
                        parent_p.remove(parent_r)
            print(f"    → очищено {len(t_elems)-1} доп. runs")
        else:
            print(f"    WARNING: a:t не найден в ячейке {tc_idx}")

    # Вставляем новую строку после template_tr (или после предыдущей вставленной)
    if i == 0:
        # Первая вставка — после Туркменистана
        template_tr.addnext(new_tr)
        print(f"  → вставлена после R{turkmen_idx} (Туркменистан)")
    else:
        # Вторая вставка — после первой новой строки (Иран)
        # Нужно найти свежую строку с Ираном
        current_rows = tbl.findall(qn('a:tr'))
        for j, tr in enumerate(current_rows):
            text = get_row_text(tr)
            if "Иран" in text:
                tr.addnext(new_tr)
                print(f"  → вставлена после R{j} (Иран)")
                break

# Сохраняем
prs.save(DST)
print(f"\n✅ Файл сохранён: {DST}")

# Проверяем
print(f"\n=== Проверка результата ===")
prs2 = Presentation(DST)
slide2 = prs2.slides[4]
for shape in slide2.shapes:
    if shape.has_table:
        table2 = shape.table
        for i, row in enumerate(table2.rows):
            cells = [cell.text.strip()[:50] for cell in row.cells]
            print(f"  R{i}: {cells}")
        break
