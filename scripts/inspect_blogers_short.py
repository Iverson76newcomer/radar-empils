"""Извлечь все YouTube-каналы из файла и найти те, у которых URL неполный или отсутствует."""
from docx import Document
import re

doc = Document('/home/z/my-project/upload/TITAN_Региональные_блогеры_коротко.docx')

print("=== ПАРАГРАФЫ (поиск упоминаний YouTube) ===")
for i, p in enumerate(doc.paragraphs):
    if 'youtube' in p.text.lower() or 'ютуб' in p.text.lower() or '/@' in p.text:
        print(f"[{i}] {p.text[:200]}")

print("\n\n=== ТАБЛИЦЫ ===")
for ti, t in enumerate(doc.tables):
    print(f"\n--- ТАБЛИЦА {ti+1} ({len(t.rows)} строк × {len(t.columns)} колонок) ---")
    for ri, r in enumerate(t.rows):
        cells = [c.text.strip().replace('\n', ' | ')[:120] for c in r.cells]
        # Ищем упоминания YouTube
        line = ' | '.join(cells)
        if 'youtube' in line.lower() or '/@' in line or 'покраска' in line.lower() or 'diy' in line.lower():
            print(f"  R{ri}: {line}")
