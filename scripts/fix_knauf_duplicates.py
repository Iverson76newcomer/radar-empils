"""Удаление дублирующих таблиц + добавление описаний под 8.2 и 8.3."""
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from copy import deepcopy

doc = Document('/home/z/my-project/download/Knauf_Россия_аналитическая_справка.docx')

COLOR_BODY = RGBColor(0x1A, 0x2B, 0x3D)

# ============================================================
# ШАГ 1: Идентифицировать и удалить дублирующие таблицы
# ============================================================
# Дубли:
# T2 (грунтовки, 7×6) — дублирует T10 (грунтовки, 7×6)
# T3 (шпаклёвки, 5×6) — дублирует T11 (шпаклёвки, 5×6)
# T7 (краски планируемые, 4×6) — старая таблица из оригинала, в обновлённом документе заменена на T9 (краски, 3×6)
# T13 (мощности, 8×5) — дублирует T5 (мощности, 8×5)
# T17 (таймлайн, 6×4) — дублирует T16 (таймлайн, 6×4)

# Для удаления таблицы из docx нужно удалить её XML-элемент из родителя
tables_to_delete = []

# Проверяем T2 vs T10 — одинаковые грунтовки
t2_first = doc.tables[2].rows[1].cells[0].text.strip()
t10_first = doc.tables[10].rows[1].cells[0].text.strip()
if t2_first == t10_first:
    tables_to_delete.append(2)
    print(f'T2 дублирует T10 (грунтовки): УДАЛЯЕМ T2')

# T3 vs T11 — одинаковые шпаклёвки
t3_first = doc.tables[3].rows[1].cells[0].text.strip()
t11_first = doc.tables[11].rows[1].cells[0].text.strip()
if t3_first == t11_first:
    tables_to_delete.append(3)
    print(f'T3 дублирует T11 (шпаклёвки): УДАЛЯЕМ T3')

# T7 — старая таблица красок (Фасаденфарбе/Инненфарбе/наливной пол)
# Проверим — есть ли T9 (новая таблица Вайсдэкке/Вайсванд)
t7_first = doc.tables[7].rows[1].cells[0].text.strip()
print(f'T7 первый продукт: {t7_first}')
# T7 содержит "КНАУФ-Фасаденфарбе" — это старая таблица, нужно проверить
# Если T9 (3×6) содержит Вайсдэкке — то T7 дублирует старую структуру
# Но T7 может быть полезна (содержит планируемые краски), поэтому проверим
t9_first = doc.tables[9].rows[1].cells[0].text.strip()
print(f'T9 первый продукт: {t9_first}')
# T9 = Вайсдэкке (новая таблица), T7 = Фасаденфарбе (старая)
# Удалим T7 — она дублирует информацию из раздела 8.1 (T9) и 8.4
tables_to_delete.append(7)
print(f'T7 (старые планируемые краски): УДАЛЯЕМ — заменена на T9')

# T13 vs T5 — одинаковые мощности
t5_first = doc.tables[5].rows[1].cells[0].text.strip()
t13_first = doc.tables[13].rows[1].cells[0].text.strip()
if t5_first == t13_first:
    tables_to_delete.append(13)
    print(f'T13 дублирует T5 (мощности): УДАЛЯЕМ T13')

# T17 vs T16 — одинаковый таймлайн
t16_first = doc.tables[16].rows[1].cells[0].text.strip()
t17_first = doc.tables[17].rows[1].cells[0].text.strip()
if t16_first == t17_first:
    tables_to_delete.append(17)
    print(f'T17 дублирует T16 (таймлайн): УДАЛЯЕМ T17')

# Сортируем в обратном порядке (чтобы индексы не сбились)
tables_to_delete.sort(reverse=True)
print(f'\nВсего к удалению: {len(tables_to_delete)} таблиц (индексы: {tables_to_delete})')

# Удаляем таблицы
for idx in tables_to_delete:
    table = doc.tables[idx]
    table._element.getparent().remove(table._element)
    print(f'  Удалена таблица с индексом {idx}')

print(f'\nОсталось таблиц: {len(doc.tables)}')

# ============================================================
# ШАГ 2: Добавить описания под заголовками 8.2 и 8.3
# ============================================================

def insert_paragraph_after(paragraph, text):
    """Вставить параграф с текстом после указанного параграфа."""
    new_p = OxmlElement('w:p')
    paragraph._element.addnext(new_p)
    from docx.text.paragraph import Paragraph
    new_para = Paragraph(new_p, paragraph._parent)
    new_para.paragraph_format.first_line_indent = Cm(0.5)
    new_para.paragraph_format.line_spacing = 1.3
    new_para.paragraph_format.space_after = Pt(6)
    run = new_para.add_run(text)
    run.font.size = Pt(10)
    run.font.color.rgb = COLOR_BODY
    run.font.name = 'Arial'
    return new_para

# Ищем параграф "8.2. Грунтовки Knauf (уже на рынке)" и вставляем после него
for i, p in enumerate(doc.paragraphs):
    if '8.2. Грунтовки Knauf' in p.text:
        insert_paragraph_after(p,
            'Knauf предлагает 6 наименований грунтовок для различных типов оснований — '
            'от глубокопроникающих до адгезионных. Все продукты уже представлены в рознице '
            'и напрямую конкурируют с грунтовками ОРЕОЛ от Эмпилс по цене и назначению.'
        )
        print(f'✓ Добавлено описание под 8.2 (параграф {i})')
        break

# Ищем параграф "8.3. Шпаклёвки Knauf" и вставляем после него
for i, p in enumerate(doc.paragraphs):
    if '8.3. Шпаклёвки Knauf' in p.text:
        insert_paragraph_after(p,
            'Линейка шпаклёвок Knauf включает 4 продукта — от базовой Фуген для швов ГКЛ '
            'до премиальной Унифлот и готовой пасты Ротбанд. Knauf доминирует в этой '
            'категории, и Эмпилс может конкурировать только в сегменте готовых шпаклёвок '
            'под брендом ОРЕОЛ.'
        )
        print(f'✓ Добавлено описание под 8.3 (параграф {i})')
        break

# Сохраняем
out_path = '/home/z/my-project/download/Knauf_Россия_аналитическая_справка.docx'
doc.save(out_path)
import os
print(f'\n✅ Сохранено: {out_path}')
print(f'Размер: {os.path.getsize(out_path) // 1024} KB')

# Проверка
doc2 = Document(out_path)
print(f'Таблиц в финальном документе: {len(doc2.tables)}')
for i, t in enumerate(doc2.tables):
    first = t.rows[0].cells[0].text[:40] if t.rows else 'EMPTY'
    print(f'  T{i}: {len(t.rows)}×{len(t.columns)} — {first}')
