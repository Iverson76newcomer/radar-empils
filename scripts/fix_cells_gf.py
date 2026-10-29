"""Точечная корректировка отдельных ячеек таблиц — добавление ГФ-021."""
from docx import Document

doc = Document('/home/z/my-project/download/TITAN_ССВ_расчёт_пояснительная_справка.docx')

# Функция для замены текста в отдельной ячейке
def set_cell(table_idx, row_idx, col_idx, new_text):
    cell = doc.tables[table_idx].rows[row_idx].cells[col_idx]
    # Очищаем
    for p in cell.paragraphs:
        for run in p.runs:
            run.text = ''
    # Записываем в первый run первого параграфа
    if cell.paragraphs and cell.paragraphs[0].runs:
        cell.paragraphs[0].runs[0].text = new_text
    else:
        cell.paragraphs[0].text = new_text

# Функция для замены текста в параграфе
def update_para(para_idx, old, new):
    p = doc.paragraphs[para_idx]
    if len(p.runs) == 1:
        p.runs[0].text = p.runs[0].text.replace(old, new)
    else:
        full = ''.join(r.text for r in p.runs)
        full = full.replace(old, new)
        if p.runs:
            p.runs[0].text = full
            for r in p.runs[1:]:
                r.text = ''

# === Таблица 1 (индекс 0) ===
# R1: "320" → "320 + 232*"
set_cell(0, 1, 1, "320 + 232*")
# R1: заголовок "Цена банки 0,75 л, ₽" → "Цена: ПФ-115 + ГФ-021, ₽"
set_cell(0, 1, 0, "Цена: ПФ-115 (0,75 л) + ГФ-021 (0,9 кг), ₽")

# R4: "4" → "4 ПФ-115 + 1 ГФ-021"
set_cell(0, 4, 1, "4 ПФ-115 + 1 ГФ-021")
# R4: заголовок "Кол-во банок на 1 перекраску (10 м²)" → "Кол-во банок на 1 перекраску"
set_cell(0, 4, 0, "Кол-во банок на 1 перекраску (10 м²)")

# R5: формула "Цена × кол-во банок" — обновить источник
# T1 R5 уже показывает 1 512 — OK

# === Таблица 2 (индекс 1) — Шаг 1 ===
# R1: "320 ₽" → "320 ₽ + 232 ₽"
set_cell(1, 1, 1, "320 ₽ + 232 ₽")
# R0: "Цена банки 0,75 л" → "Цена: ПФ-115 + ГФ-021"
set_cell(1, 0, 1, "Цена: ПФ-115 (0,75 л) + ГФ-021 (0,9 кг)")

# === Таблица 5 (индекс 4) — Шаг 4 ===
# R1: "4 банки" → "4 банки ПФ-115 + 1 банка ГФ-021"
set_cell(4, 1, 1, "4 банки ПФ-115 + 1 банка ГФ-021")
# R0: "Кол-во банок 0,75 л на 1 перекраску" — OK
set_cell(4, 0, 1, "Кол-во банок на 1 перекраску")

# === Таблица 14 (индекс 13) — Сводная ===
# R1: "Цена банки 0,75 л" → "Цена: ПФ-115 + ГФ-021"
set_cell(13, 1, 1, "Цена: ПФ-115 + ГФ-021")
# R1: "320 ₽" → "320 ₽ + 232 ₽"
set_cell(13, 1, 2, "320 ₽ + 232 ₽")

# R4: "4 шт." → "4 ПФ-115 + 1 ГФ-021"
set_cell(13, 4, 2, "4 ПФ + 1 ГФ")

# R5: формула "Цена × кол-во банок" → "ПФ: 320×4 + ГФ: 232×1"
set_cell(13, 5, 4, "ПФ: 320×4 + ГФ: 232×1")

# === Параграф 35 (Шаг 5) — формула ===
update_para(35, "320 ₽ × 4 = 1 512 ₽", "320 ₽ × 4 + 232 ₽ × 1 = 1 512 ₽")

# === Параграф 16 (Шаг 1) — почистить двойное "средняя цена" ===
update_para(16, "ПФ-115 — 320 ₽ (средняя цена по 5 брендам). ГФ-021 — 232 ₽ за 0,9 кг (банка). (Лакра, Ореол, Новбытхим, Текс, VGT)", 
            "ПФ-115 — 320 ₽ (средняя цена по 5 брендам: Лакра, Ореол, Новбытхим, Текс, VGT). ГФ-021 — 232 ₽ за 0,9 кг (банка)")

# Сохраняем
doc.save('/home/z/my-project/download/TITAN_ССВ_расчёт_пояснительная_справка.docx')
print('✅ Точечные правки применены')

# Проверка
doc2 = Document('/home/z/my-project/download/TITAN_ССВ_расчёт_пояснительная_справка.docx')
print('\n=== ТАБЛИЦА 1 ===')
for ri, row in enumerate(doc2.tables[0].rows):
    cells = [c.text.strip()[:60] for c in row.cells]
    print(f'  R{ri}: {" | ".join(cells)}')

print('\n=== ТАБЛИЦА 2 ===')
for ri, row in enumerate(doc2.tables[1].rows):
    cells = [c.text.strip()[:60] for c in row.cells]
    print(f'  R{ri}: {" | ".join(cells)}')

print('\n=== ТАБЛИЦА 5 ===')
for ri, row in enumerate(doc2.tables[4].rows):
    cells = [c.text.strip()[:60] for c in row.cells]
    print(f'  R{ri}: {" | ".join(cells)}')

print('\n=== ТАБЛИЦА 14 (сводная) ===')
for ri, row in enumerate(doc2.tables[13].rows):
    cells = [c.text.strip()[:60] for c in row.cells]
    print(f'  R{ri}: {" | ".join(cells)}')

print('\n=== ПАРАГРАФ 35 (шаг 5) ===')
print(doc2.paragraphs[35].text[:200])

print('\n=== ПАРАГРАФ 16 (шаг 1) ===')
print(doc2.paragraphs[16].text[:300])
