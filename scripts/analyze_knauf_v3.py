"""Полный анализ файла Knauf v3 на ошибки, противоречия и несоответствия."""
from docx import Document

doc = Document('/home/z/my-project/upload/Knauf_Россия_аналитическая_справка v3.docx')

print("=" * 100)
print("ПОЛНЫЙ ТЕКСТ ДОКУМЕНТА")
print("=" * 100)
for i, p in enumerate(doc.paragraphs):
    if p.text.strip():
        print(f"[{i}] {p.text}")

print(f"\n\n{'=' * 100}")
print(f"ТАБЛИЦЫ ({len(doc.tables)})")
print("=" * 100)
for ti, t in enumerate(doc.tables):
    print(f"\n--- ТАБЛИЦА {ti+1}: {len(t.rows)}×{len(t.columns)} ---")
    for ri, r in enumerate(t.rows):
        cells = [c.text.strip().replace('\n', ' | ')[:80] for c in r.cells]
        print(f"  R{ri}: {' || '.join(cells)}")
