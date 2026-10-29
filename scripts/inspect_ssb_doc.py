"""Анализ документа TITAN_ССВ_расчёт_пояснительная_справка.docx и замена цен TITAN на 1050 руб."""
from docx import Document
import re

doc = Document('/home/z/my-project/upload/TITAN_ССВ_расчёт_пояснительная_справка.docx')

print("=" * 80)
print("ПАРАГРАФЫ")
print("=" * 80)
for i, p in enumerate(doc.paragraphs):
    if p.text.strip():
        print(f"[{i}] {p.text[:200]}")

print("\n\n" + "=" * 80)
print(f"ТАБЛИЦЫ ({len(doc.tables)})")
print("=" * 80)
for ti, t in enumerate(doc.tables):
    print(f"\n--- ТАБЛИЦА {ti+1}: {len(t.rows)} строк × {len(t.columns)} колонок ---")
    for ri, r in enumerate(t.rows):
        cells = [c.text.strip().replace('\n', ' | ')[:80] for c in r.cells]
        print(f"  R{ri}: {' || '.join(cells)}")
