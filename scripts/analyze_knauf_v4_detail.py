"""
Print specific paragraphs around the financial section, market section, etc.
to analyze cross-references and contradictions
"""
from docx import Document

DOC_PATH = "/home/z/my-project/upload/Knauf_Россия_аналитическая_справка v4.docx"
doc = Document(DOC_PATH)

print("ВСЕ ПАРАГРАФЫ ДОКУМЕНТА (v4):")
print("=" * 100)
for i, para in enumerate(doc.paragraphs):
    text = para.text.strip()
    if not text:
        continue
    style = para.style.name if para.style else "Normal"
    print(f"[P{i:03d}][{style}] {text}")

print("\n\n" + "=" * 100)
print("ПОИСК КЛЮЧЕВЫХ ТЕМ:")
print("=" * 100)

# Look for specific keywords and print surrounding paragraphs
keywords = [
    "100,2 млрд", "100 млрд", "100,2", "100",
    "Forbes", "Форбс", "RuWiki", "РуВики",
    "3 785", "3785", "1 500", "1500",
    "1 400", "1400", "1 450", "1450", "1 480", "1480",
    "Колпино", "Кунгур",
    "Петроборд", "1,8 млрд", "15,23 млрд", "176 млн", "1,33 млрд",
    "46,9", "46.9", "testfirm", "тестфирм",
    "2,13 млн", "2-13 млн", "2 - 13 млн", "2,13",
    "4 500", "4500", "сотрудник", "Численность",
    "bo.nalog", "audit-it", "checko",
    "Инсулейшн", "Кубань", "Дзержинск",
    "н/д", "нет данных",
]

paras = [p.text for p in doc.paragraphs]
# Also include table cells
all_text_items = list(paras)
for t_idx, table in enumerate(doc.tables):
    for r_idx, row in enumerate(table.rows):
        for c_idx, cell in enumerate(row.cells):
            all_text_items.append(cell.text)

for kw in keywords:
    matches = []
    for idx, txt in enumerate(all_text_items):
        if kw in txt:
            source = "P" if idx < len(paras) else f"T_cell"
            matches.append((idx, txt))
    if matches:
        print(f"\n>>> '{kw}': {len(matches)} совпадений")
        for idx, txt in matches[:5]:
            source = f"P{idx}" if idx < len(paras) else f"T{idx}"
            print(f"   [{source}]: {txt[:200]}")
