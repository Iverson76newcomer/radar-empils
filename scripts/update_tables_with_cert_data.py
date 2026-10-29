"""
Update Tables #2, #5, and #6 in the document with real certificate data
Tables:
- Table #2 (index 2): Products and their fire classes - add РП1, certificate numbers, correct classes
- Table #5 (index 5): Zone/height matrix with recommended products
- Table #6 (index 6): Matrix of products × fire classes
"""
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

INPUT_PATH = '/home/z/my-project/download/Эмпилс_ЛКМ_пожароопасность_жилье_с_сертификатами.docx'
OUTPUT_PATH = '/home/z/my-project/download/Эмпилс_ЛКМ_пожароопасность_жилье_с_сертификатами.docx'

doc = Document(INPUT_PATH)

# Helper to update a cell's text while preserving formatting
def update_cell_text(cell, new_text):
    """Update cell text preserving basic formatting"""
    # Clear existing paragraphs
    for p in cell.paragraphs:
        for run in p.runs:
            run.text = ''
    # Set new text in first paragraph
    if cell.paragraphs:
        if cell.paragraphs[0].runs:
            cell.paragraphs[0].runs[0].text = new_text
        else:
            cell.paragraphs[0].add_run(new_text)
    else:
        cell.text = new_text

# ============================================================
# TABLE #2 (index 2): Products and their fire classes
# Update with real certificate data, add РП1, certificate numbers
# ============================================================
table2 = doc.tables[2]
print("=== Обновление Таблицы #2 (продукты Эмпилс) ===")
print(f"Было: {len(table2.rows)} строк")
for r_idx, row in enumerate(table2.rows):
    cells = [cell.text.strip()[:50] for cell in row.cells]
    print(f"  R{r_idx}: {' | '.join(cells)}")

# Update each row with corrected data based on certificates
# Original headers: № | Продукт | Тип | Применение | Класс пожароопасности*
# Need to add: certificate number column - but easier to just update class text with cert info

# Row 1: NEWTONE - was "Г1, В1, Д2, Т2", add РП1 and cert numbers
update_cell_text(table2.rows[1].cells[4],
    'Г1, В1, Д2, Т2, РП1\n'
    'Сертификаты:\n'
    '• RU Д-RU.PA01.B.00924/23 (NEWTONE INTERIO)\n'
    '• RU Д-RU.PA01.B.00925/23 (NEWTONE LOTUS 3)\n'
    '• RU Д-RU.PA01.B.00923/23 (NEWTONE LOTUS 20)\n'
    '• RU Д-RU.RA01.B.00926/23 (NEWTONE LOTUS 7)\n'
    'Действует до: 19.06.2028'
)
# Set font size
for p in table2.rows[1].cells[4].paragraphs:
    for r in p.runs:
        r.font.size = Pt(8)

# Row 2: ОРЕОЛ ДИСКОНТ интерьерная - upgrade to КМ1
update_cell_text(table2.rows[2].cells[4],
    'КМ1 (Г1, В1, Д1, Т1, РП1) ⬆️\n'
    'Сертификат: RU Д-RU.PA01.B.00164/23\n'
    '(ОРЕОЛ белая моющаяся ВД полиакриловая)\n'
    'Действует до: 05.02.2028'
)
for p in table2.rows[2].cells[4].paragraphs:
    for r in p.runs:
        r.font.size = Pt(8)

# Row 3: ОРЕОЛ ДИСКОНТ моющаяся - upgrade to КМ1
update_cell_text(table2.rows[3].cells[4],
    'КМ1 (Г1, В1, Д1, Т1, РП1) ⬆️\n'
    'Сертификаты:\n'
    '• RU Д-RU.PA01.B.00164/23 (ОРЕОЛ белая моющаяся)\n'
    '• RU Д-RU.RA01.B.03124/21 (Краска интерьерная моющаяся)\n'
    'Действует до: 05.02.2028 / до 2026 г.'
)
for p in table2.rows[3].cells[4].paragraphs:
    for r in p.runs:
        r.font.size = Pt(8)

# Row 4: ОРЕОЛ ДИСКОНТ полиакрил потолок - upgrade to КМ1
update_cell_text(table2.rows[4].cells[4],
    'КМ1 (Г1, В1, Д1, Т1, РП1) ⬆️\n'
    'Сертификат: RU Д-RU.PA01.B.00164/23\n'
    '(ОРЕОЛ белая моющаяся ВД полиакриловая)\n'
    'Действует до: 05.02.2028'
)
for p in table2.rows[4].cells[4].paragraphs:
    for r in p.runs:
        r.font.size = Pt(8)

# Row 5: ОРЕОЛ ДИСКОНТ фасадная - WAS В2, NOW В1 (better!)
update_cell_text(table2.rows[5].cells[4],
    'КМ1 (Г1, В1, Д2, Т2, РП1) ⬆️\n'
    '⚠️ В2 → В1 (улучшено по сертификату)\n'
    'Сертификат: RU Д-RU.RA01.B.03123/21\n'
    '(Краска фасадная СТО 135-2018)\n'
    'Действует до: 2026 г.'
)
for p in table2.rows[5].cells[4].paragraphs:
    for r in p.runs:
        r.font.size = Pt(8)

# Row 6: Эмаль для радиаторов акриловая - no specific cert, keep as before
update_cell_text(table2.rows[6].cells[4],
    'Г1, В1, Д2, Т2, РП1*\n'
    '* По типичным показателям ВД-эмалей.\n'
    '⚠️ В наборе сертификатов нет отдельного документа на эмаль для радиаторов.\n'
    'Запросите сертификат у производителя на конкретный артикул.'
)
for p in table2.rows[6].cells[4].paragraphs:
    for r in p.runs:
        r.font.size = Pt(8)

# Row 7: Эмаль акриловая универсальная - mixed: matte has В2 (WORSE!)
update_cell_text(table2.rows[7].cells[4],
    '⚠️ Различается по модификациям:\n'
    '• Матовая: Г1, В2, Д2, Т2 (RU Д-RU.RA01.B.00950/26)\n'
    '• Глянцевая: Г1, В1, Д1, Т1 (RU Д-RU.PA01.B.01134/26)\n'
    '• СТО 136-2018: Г1, В2, Д2, Т2 (RU Д-RU.RA01.B.00723/24)\n'
    'Действует до: 20.05.2031 / 24.05.2031 / 02.04.2029'
)
for p in table2.rows[7].cells[4].paragraphs:
    for r in p.runs:
        r.font.size = Pt(8)

# Add note paragraph after Table #2 about new products
# Find the paragraph after Table #2
# Add 2 new rows for new products discovered via certificates
# Actually let me add new rows to the table

# Add row for ББ-11 (new product!)
new_row1 = table2.add_row()
update_cell_text(new_row1.cells[0], '8')
update_cell_text(new_row1.cells[1], 'Раствор водный ББ-11 (1 и 2 гр) — НОВЫЙ ПРОДУКТ')
update_cell_text(new_row1.cells[2], 'Огнезащитная пропитка для древесины')
update_cell_text(new_row1.cells[3], 'Огнезащитная обработка деревянных поверхностей (ДСП, МДФ, массив) под ВД-эмаль. Повышает класс пожаробезопасности древесины с Г3/В2 до Г1/В1.')
update_cell_text(new_row1.cells[4],
    '🟢 Г1, В1, Д1, Т1, РП1 (САМЫЙ БЕЗОПАСНЫЙ)\n'
    'Сертификат: ЕАЭС RU C-RU.ПБ68.В.02464/26\n'
    '(ТР ЕАЭС 043-2017)\n'
    'Действует до: 29.03.2031'
)
for cell in new_row1.cells:
    for p in cell.paragraphs:
        for r in p.runs:
            r.font.size = Pt(9 if cell == new_row1.cells[0] else 8)

# Add row for EMPILS E 608 (new product!)
new_row2 = table2.add_row()
update_cell_text(new_row2.cells[0], '9')
update_cell_text(new_row2.cells[1], 'Компаунд эпоксидный «EMPILS» E 608 — НОВЫЙ ПРОДУКТ')
update_cell_text(new_row2.cells[2], 'Эпоксидное покрытие для полов')
update_cell_text(new_row2.cells[3], 'Покрытие полов в жилых и общественных зданиях. Высокая износостойкость, водостойкость.')
update_cell_text(new_row2.cells[4],
    '🟢 Г1, В1, Д1, Т1, РП1 (САМЫЙ БЕЗОПАСНЫЙ)\n'
    'Сертификат: № RU C-RU.ПБ68.В.02055/24\n'
    '(ТР № 123-ФЗ)\n'
    'Действует до: 27.10.2029'
)
for cell in new_row2.cells:
    for p in cell.paragraphs:
        for r in p.runs:
            r.font.size = Pt(9 if cell == new_row2.cells[0] else 8)

# ============================================================
# TABLE #6 (index 6): Matrix of products × fire classes
# Update checks/crosses based on actual certificate data
# ============================================================
table6 = doc.tables[6]
print(f"\n=== Обновление Таблицы #6 (матрица продуктов) ===")

# Update header - РП2 should be РП1 based on real data
update_cell_text(table6.rows[0].cells[5], 'РП1\n(не распр.)')

# Update each row's fire class columns (indices 1-5) based on real certificates

# Row 1: NEWTONE - Г1, В1, Д2, Т2, РП1 - was already correct
# No changes needed, all marked ✓

# Row 2: ОРЕОЛ ДИСКОНТ интерьерная - upgrade to КМ1: Г1, В1, Д1, Т1, РП1
# Was: ✓ ✓ ✓ ✓ ✓ - need to update column descriptions
# Header is Г1, В1, Д2, Т2, РП2 - so Д1 means BETTER than Д2, T1 means BETTER than T2, РП1 means BETTER than РП2
# We need to show "✓+" for better than required
for col_idx in range(1, 6):
    cell = table6.rows[2].cells[col_idx]
    update_cell_text(cell, '✓+')
# Update the "Подходит для" column
update_cell_text(table6.rows[2].cells[6],
    '✅ КМ1 — Все жилые здания, до 17 этажей, залы 15-800 чел.\n'
    '(ЛУЧШЕ чем требовалось: Д1/Т1/РП1 вместо Д2/Т2/РП2)'
)

# Row 3: ОРЕОЛ ДИСКОНТ моющаяся - upgrade to КМ1
for col_idx in range(1, 6):
    cell = table6.rows[3].cells[col_idx]
    update_cell_text(cell, '✓+')
update_cell_text(table6.rows[3].cells[6],
    '✅ КМ1 — Кухни, прихожие, коридоры, ванные\n'
    '(ЛУЧШЕ чем требовалось: Д1/Т1/РП1)'
)

# Row 4: ОРЕОЛ ДИСКОНТ полиакрил потолок - upgrade to КМ1
for col_idx in range(1, 6):
    cell = table6.rows[4].cells[col_idx]
    update_cell_text(cell, '✓+')
update_cell_text(table6.rows[4].cells[6],
    '✅ КМ1 — Потолки во всех типах зданий\n'
    '(ЛУЧШЕ чем требовалось: Д1/Т1/РП1)'
)

# Row 5: ОРЕОЛ ДИСКОНТ фасадная - was В2, now В1!
# Header В1, so now it's ✓
update_cell_text(table6.rows[5].cells[1], '✓')  # Г1
update_cell_text(table6.rows[5].cells[2], '✓+')  # В1 - upgraded!
update_cell_text(table6.rows[5].cells[3], '✓')  # Д2 (matches header)
update_cell_text(table6.rows[5].cells[4], '✓')  # Т2
update_cell_text(table6.rows[5].cells[5], '✓+')  # РП1 - upgraded!
update_cell_text(table6.rows[5].cells[6],
    '✅ Только наружные фасады\n'
    '⚠️ В2 → В1 (улучшено по сертификату)\n'
    '⚠️ РП2 → РП1 (улучшено)'
)

# Row 6: Эмаль для радиаторов - keep as is, no specific cert
# But add РП1 (was РП2)
update_cell_text(table6.rows[6].cells[5], '✓*')  # РП1 with note
update_cell_text(table6.rows[6].cells[6],
    '✅ Радиаторы, трубы (везде)\n'
    '* По типичным показателям ВД-эмалей\n'
    'Запросите сертификат на конкретный артикул'
)

# Row 7: Эмаль акриловая универсальная - MIXED!
# Mat: Г1 ✓, В2 ✗, Д2 ✓, Т2 ✓, РП1 ✓+ (better than РП2 header)
# Glossy: Г1 ✓, В1 ✓, Д1 ✓+, Т1 ✓+, РП1 ✓+
update_cell_text(table6.rows[7].cells[1], '✓')  # Г1 (both)
update_cell_text(table6.rows[7].cells[2], '✓/✗')  # В1 - glossy yes, matte NO
update_cell_text(table6.rows[7].cells[3], '✓/✓+')  # Д2/Д1
update_cell_text(table6.rows[7].cells[4], '✓/✓+')  # Т2/Т1
update_cell_text(table6.rows[7].cells[5], '✓+')  # РП1 (better than РП2 header)
update_cell_text(table6.rows[7].cells[6],
    '⚠️ Дерево, металл внутри зданий\n'
    '• Глянцевая: Г1, В1, Д1, Т1 (✅ подходит везде)\n'
    '• Матовая: Г1, В2, Д2, Т2 (⚠️ только для квартир,\n'
    '  НЕ для путей эвакуации из-за В2)'
)

# Add new rows for new products
new_row1_m = table6.add_row()
update_cell_text(new_row1_m.cells[0], 'Раствор ББ-11 (новый)')
for col_idx in range(1, 6):
    update_cell_text(new_row1_m.cells[col_idx], '✓+')  # All better than headers
update_cell_text(new_row1_m.cells[6],
    '🟢 Огнезащита древесины — повышает класс\n'
    'с Г3/В2 до Г1/В1. ТР ЕАЭС 043-2017'
)

new_row2_m = table6.add_row()
update_cell_text(new_row2_m.cells[0], 'EMPILS E 608 (новый)')
for col_idx in range(1, 6):
    update_cell_text(new_row2_m.cells[col_idx], '✓+')
update_cell_text(new_row2_m.cells[6],
    '🟢 Эпоксидное покрытие полов\n'
    'КМ1 — Г1, В1, Д1, Т1, РП1'
)

# Set font size for all cells in table6
for row in table6.rows:
    for cell in row.cells:
        for p in cell.paragraphs:
            for r in p.runs:
                r.font.size = Pt(8)

# ============================================================
# TABLE #5 (index 5): Zone/height matrix with recommended products
# Update notes about new products
# ============================================================
table5 = doc.tables[5]
print(f"\n=== Обновление Таблицы #5 (матрица зон/этажности) ===")

# Update row 6 (Деревянные элементы, ДСП, МДФ) - add ББ-11
update_cell_text(table5.rows[6].cells[5],
    'Эмаль акриловая универсальная ОРЕОЛ\n'
    '🆕 + Раствор ББ-11 (огнезащитная пропитка)'
)
for p in table5.rows[6].cells[5].paragraphs:
    for r in p.runs:
        r.font.size = Pt(9)

# Update row 3 (Вестибюли, лестн. клетки) - clarify which products are OK
update_cell_text(table5.rows[3].cells[5],
    'NEWTONE (Г1, В1, Д2, Т2, РП1) — для зданий до 17 этажей\n'
    '🆕 EMPILS E 608 (для полов, КМ1)'
)
for p in table5.rows[3].cells[5].paragraphs:
    for r in p.runs:
        r.font.size = Pt(9)

# Save document
doc.save(OUTPUT_PATH)

import os
file_size = os.path.getsize(OUTPUT_PATH)
print(f'\n✅ Файл сохранён: {OUTPUT_PATH}')
print(f'Размер: {file_size} байт ({file_size/1024:.1f} KB)')

# Verify changes
doc2 = Document(OUTPUT_PATH)
print(f'\n=== Проверка обновлений ===')
print(f'\nТаблица #2 (продукты):')
table2_new = doc2.tables[2]
print(f'  Строк: {len(table2_new.rows)} (было 8, стало {len(table2_new.rows)})')
for r_idx in [1, 2, 5, 7, 8, 9]:
    if r_idx < len(table2_new.rows):
        cell_text = table2_new.rows[r_idx].cells[4].text[:100].replace('\n', ' | ')
        print(f'  R{r_idx}: {cell_text}')

print(f'\nТаблица #6 (матрица):')
table6_new = doc2.tables[6]
print(f'  Строк: {len(table6_new.rows)} (было 8, стало {len(table6_new.rows)})')
for r_idx in [2, 5, 7, 8, 9]:
    if r_idx < len(table6_new.rows):
        row_text = ' | '.join([c.text[:15].replace('\n', ' ') for c in table6_new.rows[r_idx].cells])
        print(f'  R{r_idx}: {row_text}')
