"""Точное обновление TITAN_Региональные_блогеры_коротко.docx
   Заменяем только «через YT поиск» в колонке «Канал» (cell 1).
   ИДЕИ ДЛЯ ДОМА → https://www.youtube.com/channel/UCxRuH8DPKMl7lKOrWw7KWiQ
   Деревенский кузнец → https://www.youtube.com/@ruralsmith
   Художественная ковка металла → https://rutube.ru/channel/24685954/
"""
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

doc = Document('/home/z/my-project/upload/TITAN_Региональные_блогеры_коротко.docx')

updates = {
    'ИДЕИ ДЛЯ ДОМА': 'https://www.youtube.com/channel/UCxRuH8DPKMl7lKOrWw7KWiQ',
    'Деревенский кузнец': 'https://www.youtube.com/@ruralsmith',
    'Художественная ковка металла': 'https://rutube.ru/channel/24685954/',
}

def set_cell_text(cell, new_text):
    """Полная замена текста ячейки с сохранением форматирования первого run."""
    paragraphs = cell.paragraphs
    if len(paragraphs) > 1:
        for p in paragraphs[1:]:
            p._p.getparent().remove(p._p)
    p = paragraphs[0]
    runs = p.runs
    if len(runs) > 1:
        for r in runs[1:]:
            r._r.getparent().remove(r._r)
    if runs:
        runs[0].text = new_text
    else:
        r = OxmlElement('w:r')
        t = OxmlElement('w:t')
        t.text = new_text
        r.append(t)
        p._p.append(r)

for table in doc.tables:
    for row in table.rows:
        cells = row.cells
        if len(cells) > 1:
            cell1_text = cells[1].text.strip()
            for channel_name, new_url in updates.items():
                if channel_name in cell1_text:
                    print(f"Найден: {channel_name}")
                    new_cell1_text = f"{channel_name}\n{new_url}"
                    set_cell_text(cells[1], new_cell1_text)
                    print(f"  cell 1 → {new_url}")

# Сохраняем
out_path = '/home/z/my-project/download/TITAN_Региональные_блогеры_коротко.docx'
doc.save(out_path)
print(f'\n✅ Сохранено: {out_path}')

# Проверка
doc2 = Document(out_path)
print('\n=== ПРОВЕРКА ===')
for table in doc2.tables:
    for ri, row in enumerate(table.rows):
        cells = row.cells
        if len(cells) > 1:
            text = cells[1].text
            for channel_name in updates.keys():
                if channel_name in text:
                    print(f"\n{channel_name}:")
                    print(f"  Канал: {cells[1].text.strip()}")
                    print(f"  Контакт: {cells[5].text.strip()}")
