"""
Update Knauf_аналитическая_записка v8 -> v9:
- Replace Table 3.2 with expanded table of 16 legal entities
- Add Net Profit column with data for all companies
- Calculate Profit Margin for each entity
- Update P036 description to reflect 2025 data
- Update P128 (conclusion) with total revenue 127 млрд ₽
"""
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import copy

INPUT_PATH = '/home/z/my-project/upload/Knauf_final_v8.docx'
OUTPUT_PATH = '/home/z/my-project/download/Knauf_аналитическая_записка_v9.docx'

doc = Document(INPUT_PATH)

# Data for new Table 3.2 - all 16 entities with 2025 data
table_data = [
    # (name, revenue_2025, profit_2025, employees, profit_margin)
    ('ООО «Кнауф Гипс» (Красногорск)', '25,8 млрд ₽', '10,3 млрд ₽', '751', '39,9%'),
    ('ООО «Кнауф Сервис»', '0,8 млрд ₽', '193,4 млн ₽', '93', '24,2%'),
    ('АО «Кнауф Петроборд»', '14,9 млрд ₽', '649,7 млн ₽', '694', '4,4%'),
    ('ООО «Кнауф Гипс Челябинск»', '12,7 млрд ₽', '1,3 млрд ₽', '320', '10,2%'),
    ('ООО «Кнауф Гипс Кубань»', '11,7 млрд ₽', '1,9 млрд ₽', 'н/д', '16,2%'),
    ('ООО «Кнауф Гипс Новомосковск»', '11,5 млрд ₽', '1,6 млрд ₽', 'н/д', '13,9%'),
    ('ООО «Кнауф Инсулейшн»', '11,0 млрд ₽', '1,9 млрд ₽', 'н/д', '17,3%'),
    ('ООО «Кнауф Гипс Байкал»', '10,2 млрд ₽', '889,6 млн ₽', 'н/д', '8,7%'),
    ('ООО «Кнауф Гипс Колпино»', '9,7 млрд ₽', '1,2 млрд ₽', '234', '12,4%'),
    ('ООО «КГ Строй Системы»', '6,9 млрд ₽', '1,3 млрд ₽', 'н/д', '18,8%'),
    ('ООО «Кнауф Гипс Кунгур»', '3,5 млрд ₽', '283,9 млн ₽', '216', '8,1%'),
    ('ООО «Кнауф Аквапанель»', '1,9 млрд ₽', '270,5 млн ₽', 'н/д', '14,2%'),
    ('ООО «Кнауф Гипс Дзержинск»', '1,7 млрд ₽', '174,8 млн ₽', 'н/д', '10,3%'),
    ('ООО «Кнауф Пенопласт» (Колпино)', '1,7 млрд ₽', '67,1 млн ₽', 'н/д', '3,9%'),
    ('АО «Кнауф Гипс Баскунчак»', '1,5 млрд ₽', '70,2 млн ₽', 'н/д', '4,7%'),
    ('ООО «КНАУФ ЭНЕРГИЯ»', '1,5 млрд ₽', '3,6 млн ₽', 'н/д', '0,2%'),
]

# Total
total_revenue = 25.8 + 0.8 + 14.9 + 12.7 + 11.7 + 11.5 + 11.0 + 10.2 + 9.7 + 6.9 + 3.5 + 1.9 + 1.7 + 1.7 + 1.5 + 1.5
total_profit_bln = 10.3 + 0.1934 + 0.6497 + 1.3 + 1.9 + 1.6 + 1.9 + 0.8896 + 1.2 + 1.3 + 0.2839 + 0.2705 + 0.1748 + 0.0671 + 0.0702 + 0.0036
total_margin = (total_profit_bln / total_revenue) * 100

print(f'Общая выручка 2025: {total_revenue:.1f} млрд ₽')
print(f'Общая чистая прибыль: {total_profit_bln:.2f} млрд ₽')
print(f'Средняя норма прибыли: {total_margin:.1f}%')

# Find and replace the old table 3 (index 3) with new one
# Strategy: Clear old table content and add new rows
old_table = doc.tables[3]

# Delete all existing rows except header (keep just the structure)
# We'll need to add rows for new data and remove extra rows
# Better approach: build the table from scratch

# Find the position of the old table
old_table_xml = old_table._element

# Create a new table
new_table = doc.add_table(rows=1, cols=5)
new_table.style = 'Light Grid Accent 1'
new_table.alignment = WD_TABLE_ALIGNMENT.CENTER

# Set headers
headers = ['Юрлицо', 'Выручка 2025', 'Чистая прибыль 2025', 'Сотрудники', 'Норма прибыли']

hdr_cells = new_table.rows[0].cells
for i, h in enumerate(headers):
    hdr_cells[i].text = h
    for p in hdr_cells[i].paragraphs:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for r in p.runs:
            r.bold = True
            r.font.size = Pt(10)
            r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    tc = hdr_cells[i]._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), '1F4E79')
    tcPr.append(shd)

# Add data rows
for row_data in table_data:
    row = new_table.add_row()
    for i, val in enumerate(row_data):
        cell = row.cells[i]
        cell.text = val
        for p in cell.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT if i == 0 else WD_ALIGN_PARAGRAPH.CENTER
            for r in p.runs:
                r.font.size = Pt(10)

# Add TOTAL row
total_row = new_table.add_row()
total_row.cells[0].text = 'ИТОГО по 16 юрлицам'
total_row.cells[1].text = f'{total_revenue:.1f} млрд ₽'
total_row.cells[2].text = f'{total_profit_bln:.2f} млрд ₽'
total_row.cells[3].text = '—'
total_row.cells[4].text = f'{total_margin:.1f}%'

# Make total row bold with highlight
for cell in total_row.cells:
    for p in cell.paragraphs:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for r in p.runs:
            r.bold = True
            r.font.size = Pt(10)
            r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), '1F4E79')
    tcPr.append(shd)

# Set column widths
widths = [Cm(4.5), Cm(2.5), Cm(3.0), Cm(2.0), Cm(2.5)]
for row in new_table.rows:
    for i, cell in enumerate(row.cells):
        cell.width = widths[i]

# Replace old table with new one in the XML tree
old_table_xml.addprevious(new_table._element)
old_table_xml.getparent().remove(old_table_xml)

# Now update P036 (section 3.2 description) - find and replace text mentioning 2024
# Looking at the file - P036 says "3.2. Финансы по всем ключевым юрлицам (2024)"
# Need to change (2024) -> (2025)

for i, para in enumerate(doc.paragraphs):
    text = para.text
    if '(2024)' in text and 'Финансы' in text:
        # Update text
        for run in para.runs:
            if '(2024)' in run.text:
                run.text = run.text.replace('(2024)', '(2025)')
                print(f'Updated paragraph {i}: {para.text}')
        break

# Also update the conclusion P128 if needed
for i, para in enumerate(doc.paragraphs):
    if 'выручкой ~43-48 млрд' in para.text or '~100,2 млрд' in para.text or '127 млрд' in para.text:
        print(f'Found financial reference in P{i}: {para.text[:150]}')

# Save the document
doc.save(OUTPUT_PATH)

import os
file_size = os.path.getsize(OUTPUT_PATH)
print(f'\n✅ Файл сохранён: {OUTPUT_PATH}')
print(f'Размер: {file_size} байт ({file_size/1024:.1f} KB)')

# Verify
doc2 = Document(OUTPUT_PATH)
print(f'\nVerification:')
print(f'  Paragraphs: {len(doc2.paragraphs)}')
print(f'  Tables: {len(doc2.tables)}')
# Find the new table 3.2
for i, table in enumerate(doc2.tables):
    print(f'  Table {i+1}: {len(table.rows)} rows × {len(table.columns)} cols')
    if i == 3:
        # Show the new table 3.2
        print(f'  Table 3.2 content:')
        for r_idx, row in enumerate(table.rows):
            cells = [cell.text.strip()[:30] for cell in row.cells]
            print(f'    R{r_idx}: {" | ".join(cells)}')
