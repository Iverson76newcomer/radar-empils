"""Дополнительные правки: 322→360 в таблице 1, 2,8→2,5."""
from docx import Document

doc = Document('/home/z/my-project/download/TITAN_ССВ_расчёт_пояснительная_справка.docx')

# Дополнительные замены
replacements = [
    ("322", "360"),       # В таблице 1 R11 значение без ₽
    ("2,8 раза", "2,5 раза"),
    ("2.8 раза", "2.5 раза"),
]

def replace_in_runs(paragraph):
    full_text = paragraph.text
    has_match = any(old in full_text for old, _ in replacements)
    if not has_match:
        return False
    
    if len(paragraph.runs) == 1:
        text = paragraph.runs[0].text
        for old, new in replacements:
            text = text.replace(old, new)
        paragraph.runs[0].text = text
        return True
    
    all_text = ''
    for run in paragraph.runs:
        all_text += run.text
    for old, new in replacements:
        all_text = all_text.replace(old, new)
    if paragraph.runs:
        paragraph.runs[0].text = all_text
        for run in paragraph.runs[1:]:
            run.text = ''
    return True

changed = 0
for para in doc.paragraphs:
    if replace_in_runs(para):
        changed += 1
for table in doc.tables:
    for row in table.rows:
        for cell in row.cells:
            for para in cell.paragraphs:
                if replace_in_runs(para):
                    changed += 1

doc.save('/home/z/my-project/download/TITAN_ССВ_расчёт_пояснительная_справка.docx')
print(f'✅ Доп. правки: изменено {changed} элементов')

# Проверка
doc2 = Document('/home/z/my-project/download/TITAN_ССВ_расчёт_пояснительная_справка.docx')
print('\n=== ПРОВЕРКА ===')
for ti, table in enumerate(doc2.tables):
    for ri, row in enumerate(table.rows):
        cells = [c.text.strip()[:60] for c in row.cells]
        line = ' | '.join(cells)
        if '322' in line or '360' in line or '2,8' in line or '2,5' in line or '2.8' in line or '2.5' in line:
            print(f'  Т{ti+1} R{ri}: {line}')
