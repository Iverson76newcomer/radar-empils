"""Замена цены TITAN с 860 на 1050 руб. во всём документе с пересчётом зависимых значений."""
from docx import Document
import copy

doc = Document('/home/z/my-project/upload/TITAN_ССВ_расчёт_пояснительная_справка.docx')

# Все замены (старое → новое) — порядок важен!
replacements = [
    # Цена банки
    ("860*", "1 050*"),
    ("860 ₽", "1 050 ₽"),
    ("860", "1 050"),  # остальные случаи (без ₽ и *)
    
    # Стоимость материалов 1 перекраски: 1050 × 2 = 2100
    ("1 720 ₽", "2 100 ₽"),
    ("1 720", "2 100"),
    
    # ИТОГО 1 перекраска: 2100 + 1500 = 3600
    ("3 220 ₽", "3 600 ₽"),
    ("3 220", "3 600"),
    
    # ССВ на 1 год: 3600 ÷ 10 = 360
    ("322 ₽/год", "360 ₽/год"),
    ("322 ₽", "360 ₽"),
    ("= 322", "= 360"),
    
    # Экономия за 10 лет: 9174 - 3600 = 5574
    ("5 954 ₽", "5 574 ₽"),
    ("5 954", "5 574"),
    
    # Во сколько дороже банка: 1050 / 320 = 3.28 ≈ 3,3
    ("2,7 раза", "3,3 раза"),
    ("2.7 раза", "3.3 раза"),
    
    # На 1 перекраске дороже на: 2100 - 1280 = 820
    ("440 ₽", "820 ₽"),
    ("на 440", "на 820"),
    
    # Формулы с пересчётом
    ("860 ₽ × 2 = 1 720 ₽", "1 050 ₽ × 2 = 2 100 ₽"),
    ("1 720 ₽ + 1 500 ₽ = 3 220 ₽", "2 100 ₽ + 1 500 ₽ = 3 600 ₽"),
    ("3 220 ₽ × 1 = 3 220 ₽", "3 600 ₽ × 1 = 3 600 ₽"),
    ("3 220 ₽ ÷ 10 = 322 ₽/год", "3 600 ₽ ÷ 10 = 360 ₽/год"),
    ("9 174 − 3 220", "9 174 − 3 600"),
    ("9 174 ₽ – 3 220 ₽", "9 174 ₽ – 3 600 ₽"),
    ("9 174 − 3 220", "9 174 − 3 600"),
    
    # Процент экономии: 5574 / 9174 = 60.8% ≈ 61%
    ("(65%)", "(61%)"),
    ("65%", "61%"),
]

def replace_in_runs(paragraph):
    """Замена текста в параграфе с сохранением форматирования runs."""
    # Сначала проверим — есть ли совпадения в полном тексте параграфа
    full_text = paragraph.text
    has_match = any(old in full_text for old, _ in replacements)
    if not has_match:
        return False
    
    # Если в параграфе только один run — простая замена
    if len(paragraph.runs) == 1:
        text = paragraph.runs[0].text
        for old, new in replacements:
            text = text.replace(old, new)
        paragraph.runs[0].text = text
        return True
    
    # Если несколько runs — объединяем текст, заменяем, кладём в первый run
    # Сохраняем формат первого run
    all_text = ''
    for run in paragraph.runs:
        all_text += run.text
    
    for old, new in replacements:
        all_text = all_text.replace(old, new)
    
    # Кладём результат в первый run, остальные очищаем
    if paragraph.runs:
        paragraph.runs[0].text = all_text
        for run in paragraph.runs[1:]:
            run.text = ''
    
    return True

# Обрабатываем все параграфы
changed_count = 0
for para in doc.paragraphs:
    if replace_in_runs(para):
        changed_count += 1

# Обрабатываем все таблицы
for table in doc.tables:
    for row in table.rows:
        for cell in row.cells:
            for para in cell.paragraphs:
                if replace_in_runs(para):
                    changed_count += 1

# Сохраняем
out_path = '/home/z/my-project/download/TITAN_ССВ_расчёт_пояснительная_справка.docx'
doc.save(out_path)

print(f'✅ Сохранено: {out_path}')
print(f'Изменено параграфов/ячеек: {changed_count}')

# Проверка — выведем ключевые значения
print('\n=== ПРОВЕРКА КЛЮЧЕВЫХ ЗНАЧЕНИЙ ===')
doc2 = Document(out_path)
for table in doc2.tables:
    for row in table.rows:
        cells = [c.text.strip() for c in row.cells]
        line = ' | '.join(cells)
        if '860' in line or '1 050' in line or '1 720' in line or '2 100' in line or '3 220' in line or '3 600' in line or '5 954' in line or '5 574' in line or '322' in line or '360' in line:
            print(f'  {line}')
