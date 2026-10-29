"""
Update Knauf_аналитическая_записка v9 -> v10:
- Replace Table 3.2 with ORIGINAL 6 entities (as in user's request)
- Calculate Profit Margin only for entities WITH net profit data (3 of 6)
- For others, leave "-" or "н/д"
- Update P036 description to reflect 2024 year (as per user's table)
"""
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

INPUT_PATH = '/home/z/my-project/download/Knauf_аналитическая_записка_v9.docx'
OUTPUT_PATH = '/home/z/my-project/download/Knauf_аналитическая_записка_v10.docx'

doc = Document(INPUT_PATH)

# Original 6 entities as in user's v8 table
# Calculate margin only for those with net profit data
table_data = [
    # (name, revenue, profit, employees, profit_margin)
    ('ООО «Кнауф Гипс» (Красногорск)', '24,21 млрд ₽', '~11,4 млрд ₽', '751', '46,9% (testfirm.ru, на основе ФНС)'),
    ('ООО «Кнауф Гипс Кунгур»', '3,54 млрд ₽', '~240 млн ₽', '216', '6,8%'),
    ('ООО «Кнауф Гипс Колпино»', '9,9 млрд ₽', 'н/д', '234', 'н/д'),
    ('АО «Кнауф Петроборд»', '15,23 млрд ₽', '1,33 млрд ₽', '694', '8,7%'),
    ('ООО «Кнауф Инсулейшн»', '12,8 млрд ₽', 'н/д', '—', '—'),
    ('ООО «Кнауф Гипс Кубань»', '11,49 млрд ₽', 'н/д', '—', '—'),
    ('ООО «Кнауф Гипс Дзержинск»', 'н/д', 'н/д', '—', '—'),
]

# Find the old Table 3.2 (now index 3 after we updated v9)
old_table = doc.tables[3]

# Create new table
new_table = doc.add_table(rows=1, cols=5)
new_table.style = 'Light Grid Accent 1'
new_table.alignment = WD_TABLE_ALIGNMENT.CENTER

# Headers
headers = ['Юрлицо', 'Выручка 2024', 'Чистая прибыль', 'Сотрудники', 'Норма прибыли']
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

# Data rows
for row_data in table_data:
    row = new_table.add_row()
    for i, val in enumerate(row_data):
        cell = row.cells[i]
        cell.text = val
        for p in cell.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT if i == 0 else WD_ALIGN_PARAGRAPH.CENTER
            for r in p.runs:
                r.font.size = Pt(10)

# Set column widths
widths = [Cm(4.5), Cm(2.5), Cm(3.0), Cm(2.0), Cm(3.5)]
for row in new_table.rows:
    for i, cell in enumerate(row.cells):
        cell.width = widths[i]

# Replace old table with new one
old_table._element.addprevious(new_table._element)
old_table._element.getparent().remove(old_table._element)

# Update P036 (section 3.2 description) - back to (2024)
for i, para in enumerate(doc.paragraphs):
    text = para.text
    if '(2025)' in text and 'Финансы' in text:
        for run in para.runs:
            if '(2025)' in run.text:
                run.text = run.text.replace('(2025)', '(2024)')
                print(f'Updated paragraph {i}: {para.text}')
        break

# Save
doc.save(OUTPUT_PATH)

import os
file_size = os.path.getsize(OUTPUT_PATH)
print(f'\n✅ Файл сохранён: {OUTPUT_PATH}')
print(f'Размер: {file_size} байт ({file_size/1024:.1f} KB)')

# Verify
doc2 = Document(OUTPUT_PATH)
print(f'\nVerification - Table 3.2 (table 4):')
table = doc2.tables[3]
for r_idx, row in enumerate(table.rows):
    cells = [cell.text.strip()[:40] for cell in row.cells]
    print(f'  R{r_idx}: {" | ".join(cells)}')
