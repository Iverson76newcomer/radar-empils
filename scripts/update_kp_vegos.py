"""
Обновить КП Вегос-М с новыми ТО: Год 1 = 13 млн, Год 2 = 25 млн, Год 3 = 50 млн.
Пересчитать все зависимые цифры.
"""
from docx import Document
from docx.shared import Pt, RGBColor
import os, shutil

SRC = "/home/z/my-project/upload/КП_Вегос-М_коммерческое_предложение.docx"
OUT = "/home/z/my-project/download/КП_Вегос-М_коммерческое_предложение_v2.docx"

shutil.copy(SRC, OUT)
doc = Document(OUT)

# === НОВЫЕ ДАННЫЕ ===
# Год 1: ТО=13, Год 2: ТО=25, Год 3: ТО=50
# Наценка 35%: 5, 9, 18
# Бонусы: 1,17 (9%), 2,5 (10%), 5,5 (11%)
# Доход: 6,2, 11,5, 23,5
# ИТОГО: ТО=88, наценка=32, бонус=9,2, доход=41,2

def set_cell(cell, text):
    """Обновить текст ячейки, сохраняя стиль."""
    # Сохраняем формат первого run
    p = cell.paragraphs[0]
    if p.runs:
        first_run = p.runs[0]
        font_name = first_run.font.name
        font_size = first_run.font.size
        font_bold = first_run.font.bold
        font_color = first_run.font.color.rgb if first_run.font.color and first_run.font.color.rgb else None
        
        # Очистить все runs
        for run in p.runs:
            run.text = ''
        
        # Установить новый текст в первый run
        lines = text.split('\n')
        p.runs[0].text = lines[0]
        for i, line in enumerate(lines[1:], 1):
            run = p.add_run()
            run.add_break()
            run.text = line
            if font_name:
                run.font.name = font_name
            if font_size:
                run.font.size = font_size
            if font_bold is not None:
                run.font.bold = font_bold
            if font_color:
                run.font.color.rgb = font_color
    else:
        cell.text = text

# === Таблица 1: ЗАРАБОТОК ПАРТНЁРА (сводная) ===
t = doc.tables[0]
# Row 2: Год 1
t.rows[1].cells[1].text = '13 млн ₽'
t.rows[1].cells[2].text = '5 млн ₽'
t.rows[1].cells[3].text = '1,2 млн ₽'
t.rows[1].cells[4].text = '6,2 млн ₽'
t.rows[1].cells[5].text = '—'
# Row 3: Год 2
t.rows[2].cells[1].text = '25 млн ₽'
t.rows[2].cells[2].text = '9 млн ₽'
t.rows[2].cells[3].text = '2,5 млн ₽'
t.rows[2].cells[4].text = '11,5 млн ₽'
t.rows[2].cells[5].text = '×1,9'
# Row 4: Год 3
t.rows[3].cells[1].text = '50 млн ₽'
t.rows[3].cells[2].text = '18 млн ₽'
t.rows[3].cells[3].text = '5,5 млн ₽'
t.rows[3].cells[4].text = '23,5 млн ₽'
t.rows[3].cells[5].text = '×2,0'
# Row 5: ИТОГО
t.rows[4].cells[1].text = '88 млн ₽'
t.rows[4].cells[2].text = '32 млн ₽'
t.rows[4].cells[3].text = '9,2 млн ₽'
t.rows[4].cells[4].text = '41,2 млн ₽'
t.rows[4].cells[5].text = ''

# === Таблица 2: ТОВАРООБОРОТ И ЗАРАБОТОК ПО ГОДАМ ===
t = doc.tables[1]
t.rows[1].cells[1].text = '13 млн ₽'
t.rows[1].cells[2].text = '5 млн ₽'
t.rows[2].cells[1].text = '25 млн ₽'
t.rows[2].cells[2].text = '9 млн ₽'
t.rows[4].cells[1].text = '88 млн ₽'
t.rows[4].cells[2].text = '32 млн ₽'

# === Таблица 3: СИСТЕМА БОНУСИРОВАНИЯ ===
t = doc.tables[2]
# Год 1
t.rows[1].cells[1].text = '13 млн ₽'
t.rows[1].cells[2].text = '0,52 млн ₽'
t.rows[1].cells[3].text = '0,52 млн ₽'
t.rows[1].cells[4].text = '0,13 млн ₽ (1%)'
t.rows[1].cells[5].text = '1,17 млн ₽ (9%)'
t.rows[1].cells[6].text = '6,2 млн ₽'
# Год 2
t.rows[2].cells[1].text = '25 млн ₽'
t.rows[2].cells[2].text = '1,0 млн ₽'
t.rows[2].cells[3].text = '1,0 млн ₽'
t.rows[2].cells[4].text = '0,5 млн ₽ (2%)'
t.rows[2].cells[5].text = '2,5 млн ₽ (10%)'
t.rows[2].cells[6].text = '11,5 млн ₽'
# Год 3 — без изменений (50 млн)
t.rows[3].cells[5].text = '5,5 млн ₽ (11%)'
# ИТОГО
t.rows[4].cells[1].text = '88 млн ₽'
t.rows[4].cells[2].text = '3,52 млн ₽'
t.rows[4].cells[3].text = '3,52 млн ₽'
t.rows[4].cells[4].text = '2,13 млн ₽'
t.rows[4].cells[5].text = '9,17 млн ₽'
t.rows[4].cells[6].text = '41,2 млн ₽'

# === Таблица 11: ОРГАНИЗАЦИЯ ПРОДАЖ — ЭТП ===
t = doc.tables[10]
t.rows[2].cells[1].text = '8 млн ₽/год'  # ТО через ЭТП
t.rows[3].cells[1].text = '5 млн ₽/год'  # ТО через ТП
t.rows[4].cells[1].text = '13 млн ₽'    # ИТОГО год 1

# === Таблица 12: ПРОГНОЗ ТОВАРООБОРОТА ===
t = doc.tables[11]
t.rows[1].cells[1].text = '13 млн ₽'  # Год 1
t.rows[2].cells[1].text = '25 млн ₽'  # Год 2
t.rows[2].cells[2].text = '×1,9'       # Рост
# Год 3 — без изменений (50 млн)
t.rows[3].cells[2].text = '×2,0'       # Рост

# === Обновить текстовые абзацы ===
# Абзац 5: «50 млн ₽ — целевой товарооборот... Заработок партнёра: 18 млн ₽/год...»
for p in doc.paragraphs:
    if '50 млн ₽ — целевой товарооборот' in p.text:
        for run in p.runs:
            run.text = ''
        doc.paragraphs[5].runs[0].text = '50 млн ₽ — целевой товарооборот Эмпилс в Год 3 = 8% потенциала региона. Заработок партнёра: 18 млн ₽/год при наценке 35%.'
        break

# === Сохранить ===
doc.save(OUT)
print(f"✓ Файл сохранён: {OUT}")
print(f"  Размер: {os.path.getsize(OUT):,} байт")

# === Проверка ===
doc2 = Document(OUT)
print("\n=== ПРОВЕРКА ОБНОВЛЁННЫХ ТАБЛИЦ ===\n")

print("Таблица 1 (Заработок партнёра):")
for row in doc2.tables[0].rows:
    print(f"  {[c.text[:25] for c in row.cells]}")

print("\nТаблица 2 (ТО по годам):")
for row in doc2.tables[1].rows:
    print(f"  {[c.text[:30] for c in row.cells]}")

print("\nТаблица 3 (Бонусы):")
for row in doc2.tables[2].rows:
    print(f"  {[c.text[:25] for c in row.cells]}")

print("\nТаблица 11 (ЭТП):")
for row in doc2.tables[10].rows:
    print(f"  {[c.text[:30] for c in row.cells]}")

print("\nТаблица 12 (Прогноз):")
for row in doc2.tables[11].rows:
    print(f"  {[c.text[:30] for c in row.cells]}")
