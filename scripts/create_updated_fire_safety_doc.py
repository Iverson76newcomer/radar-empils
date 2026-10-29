"""
Create updated fire safety document with:
1. Main document structure (from fire_safety_main.docx)
2. Updated tables with REAL certificate data
3. Knowledge base analysis (12 ДС + 2 Сертификата)
4. Passports of safety (ПБ) as supplementary info
5. Note about missing "Эмаль акриловая матовая" ДС
"""
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import os

# Start from the main document
INPUT_PATH = '/home/z/my-project/upload/fire_safety_main.docx'
OUTPUT_PATH = '/home/z/my-project/download/Эмпилс_ЛКМ_пожароопасность_актуализированный.docx'

doc = Document(INPUT_PATH)

# Helper to update cell text
def update_cell_text(cell, new_text):
    for p in cell.paragraphs:
        for run in p.runs:
            run.text = ''
    if cell.paragraphs:
        if cell.paragraphs[0].runs:
            cell.paragraphs[0].runs[0].text = new_text
        else:
            cell.paragraphs[0].add_run(new_text)
    else:
        cell.text = new_text

# ============================================================
# UPDATE TABLE #2 (products) with REAL certificate data
# ============================================================
table2 = doc.tables[2]

# Row 1: NEWTONE - Г1, В1, Д2, Т2, РП1
update_cell_text(table2.rows[1].cells[4],
    'Г1, В1, Д2, Т2, РП1\n'
    '✅ 4 сертификата:\n'
    '• RU Д-RU.PA01.B.00924/23 (NEWTONE INTERIO)\n'
    '• RU Д-RU.PA01.B.00925/23 (NEWTONE LOTUS 3)\n'
    '• RU Д-RU.PA01.B.00923/23 (NEWTONE LOTUS 20)\n'
    '• RU Д-RU.RA01.B.00926/23 (NEWTONE LOTUS 7)\n'
    'Действует до: 19.06.2028')

# Row 2: ОРЕОЛ ДИСКОНТ интерьерная - КМ1 (улучшено!)
update_cell_text(table2.rows[2].cells[4],
    'КМ1 (Г1, В1, Д1, Т1, РП1) ⬆️\n'
    '✅ Сертификат: RU Д-RU.PA01.B.00164/23\n'
    '(ОРЕОЛ белая моющаяся ВД полиакриловая)\n'
    'Действует до: 05.02.2028')

# Row 3: ОРЕОЛ ДИСКОНТ моющаяся - КМ1
update_cell_text(table2.rows[3].cells[4],
    'КМ1 (Г1, В1, Д1, Т1, РП1) ⬆️\n'
    '✅ 2 сертификата:\n'
    '• RU Д-RU.PA01.B.00164/23 (ОРЕОЛ белая моющаяся)\n'
    '• RU Д-RU.RA01.B.03124/21 (Краска интерьерная моющаяся)\n'
    'Действует до: 05.02.2028 / 2026 г.')

# Row 4: ОРЕОЛ ДИСКОНТ полиакрил потолок - КМ1
update_cell_text(table2.rows[4].cells[4],
    'КМ1 (Г1, В1, Д1, Т1, РП1) ⬆️\n'
    '✅ Сертификат: RU Д-RU.PA01.B.00164/23\n'
    'Действует до: 05.02.2028')

# Row 5: ОРЕОЛ ДИСКОНТ фасадная - В2→В1 (улучшено!)
update_cell_text(table2.rows[5].cells[4],
    'КМ1 (Г1, В1, Д2, Т2, РП1) ⬆️\n'
    '⚠️ В2 → В1 (улучшено по сертификату!)\n'
    '✅ Сертификат: RU Д-RU.RA01.B.03123/21\n'
    '(Краска фасадная СТО 135-2018)\n'
    'Действует до: 2026 г.')

# Row 6: Эмаль для радиаторов - нет отдельного сертификата
update_cell_text(table2.rows[6].cells[4],
    'Г1, В1, Д2, Т2, РП1*\n'
    '* По типичным показателям ВД-эмалей\n'
    '⚠️ В базе знаний нет отдельного ДС на эмаль для радиаторов\n'
    'Есть ПБ: «ПБ Эмаль для радиаторов СТО 012-2011.pdf»\n'
    'Запросите ДС у производителя на конкретный артикул')

# Row 7: Эмаль акриловая универсальная - матовая В2, глянцевая В1!
update_cell_text(table2.rows[7].cells[4],
    '⚠️ Различается по модификациям:\n'
    '• Матовая: Г1, В2, Д2, Т2 (ДС был в первой партии, но ОТСУТСТВУЕТ в базе знаний)\n'
    '• Глянцевая: Г1, В1, Д1, Т1 ✅ (RU Д-RU.PA01.B.01134/26, до 24.05.2031)\n'
    '• СТО 136-2018: Г1, В2, Д2, Т2 (RU Д-RU.RA01.B.00723/24, до 02.04.2029)\n'
    'ПБ: «ПБ Эмаль акриловая СТО 062-2010.pdf» (паспорт безопасности)')

# Set font size for updated cells
for row in table2.rows:
    for cell in row.cells:
        for p in cell.paragraphs:
            for r in p.runs:
                r.font.size = Pt(8)

# ============================================================
# UPDATE TABLE #6 (matrix of products × fire classes)
# ============================================================
table6 = doc.tables[6]

# Update header: РП2 → РП1
update_cell_text(table6.rows[0].cells[5], 'РП1\n(не распр.)')

# Row 2: ОРЕОЛ ДИСКОНТ интерьерная - КМ1 (all ✓+)
for col_idx in range(1, 6):
    update_cell_text(table6.rows[2].cells[col_idx], '✓+')
update_cell_text(table6.rows[2].cells[6],
    '✅ КМ1 — Все жилые здания, до 17 этажей\n'
    '(ЛУЧШЕ: Д1/Т1/РП1 вместо Д2/Т2/РП2)')

# Row 3: ОРЕОЛ ДИСКОНТ моющаяся - КМ1
for col_idx in range(1, 6):
    update_cell_text(table6.rows[3].cells[col_idx], '✓+')
update_cell_text(table6.rows[3].cells[6],
    '✅ КМ1 — Кухни, прихожие, коридоры\n'
    '(ЛУЧШЕ: Д1/Т1/РП1)')

# Row 4: ОРЕОЛ ДИСКОНТ полиакрил потолок - КМ1
for col_idx in range(1, 6):
    update_cell_text(table6.rows[4].cells[col_idx], '✓+')
update_cell_text(table6.rows[4].cells[6],
    '✅ КМ1 — Потолки во всех типах зданий\n'
    '(ЛУЧШЕ: Д1/Т1/РП1)')

# Row 5: ОРЕОЛ ДИСКОНТ фасадная - В2→В1!
update_cell_text(table6.rows[5].cells[1], '✓')
update_cell_text(table6.rows[5].cells[2], '✓+ (В2→В1)')
update_cell_text(table6.rows[5].cells[3], '✓')
update_cell_text(table6.rows[5].cells[4], '✓')
update_cell_text(table6.rows[5].cells[5], '✓+ (РП2→РП1)')
update_cell_text(table6.rows[5].cells[6],
    '✅ Наружные фасады\n'
    '⚠️ В2→В1, РП2→РП1 (улучшено по ДС)')

# Row 6: Эмаль для радиаторов - нет ДС
update_cell_text(table6.rows[6].cells[5], '✓*')
update_cell_text(table6.rows[6].cells[6],
    '✅ Радиаторы, трубы (везде)\n'
    '* Нет ДС в базе знаний, есть ПБ\n'
    'Запросите ДС у производителя')

# Row 7: Эмаль акриловая универсальная
update_cell_text(table6.rows[7].cells[1], '✓')
update_cell_text(table6.rows[7].cells[2], '✓/✗')
update_cell_text(table6.rows[7].cells[3], '✓/✓+')
update_cell_text(table6.rows[7].cells[4], '✓/✓+')
update_cell_text(table6.rows[7].cells[5], '✓+')
update_cell_text(table6.rows[7].cells[6],
    '⚠️ Дерево, металл внутри зданий\n'
    '• Глянцевая: Г1, В1, Д1, Т1 (✅ везде)\n'
    '• Матовая: Г1, В2, Д2, Т2 (⚠️ только квартиры)\n'
    '  ДС отсутствует в базе знаний')

# Set font size for all cells in table6
for row in table6.rows:
    for cell in row.cells:
        for p in cell.paragraphs:
            for r in p.runs:
                r.font.size = Pt(8)

# ============================================================
# UPDATE TABLE #5 (zone/height matrix with recommended products)
# ============================================================
table5 = doc.tables[5]

# Update row 6 (Деревянные элементы) - add ББ-11
update_cell_text(table5.rows[6].cells[5],
    'Эмаль акриловая универсальная ОРЕОЛ\n'
    '🆕 + Раствор ББ-11 (огнезащитная пропитка)')
for p in table5.rows[6].cells[5].paragraphs:
    for r in p.runs:
        r.font.size = Pt(9)

# Update row 3 (Вестибюли) - add EMPILS E 608
update_cell_text(table5.rows[3].cells[5],
    'NEWTONE (Г1, В1, Д2, Т2, РП1) — до 17 этажей\n'
    '🆕 EMPILS E 608 (для полов, КМ1)')
for p in table5.rows[3].cells[5].paragraphs:
    for r in p.runs:
        r.font.size = Pt(9)

# ============================================================
# ADD NEW SECTION: Knowledge Base Analysis
# ============================================================
doc.add_page_break()

title = doc.add_heading('Приложение Б. Сверка с базой знаний сертификатов и паспортов безопасности', level=1)

intro = doc.add_paragraph(
    'Проведена полная сверка основного документа с базой знаний Эмпилс, '
    'содержащей 1693 файла (сертификаты соответствия, декларации о соответствии, '
    'паспорта безопасности). Ниже — результаты сверки и выявленные расхождения.'
)
intro.paragraph_format.first_line_indent = Cm(0)

# Section 1: Summary of knowledge base
doc.add_heading('Б.1. Состав базы знаний', level=2)

kb_data = [
    ('Всего файлов в базе', '1693'),
    ('PDF-файлов (сертификаты, паспорта)', '471'),
    ('JPG-файлов (сканы паспортов)', '1043'),
    ('Деклараций о соответствии (ДС) по пожарной безопасности', '11'),
    ('Сертификатов соответствия (ТР ЕАЭС 043-2017 и ТР 123-ФЗ)', '2'),
    ('Паспортов безопасности (ПБ)', '~450'),
    ('ИТОГО документов по пожарной безопасности', '13'),
]

table_kb = doc.add_table(rows=1, cols=2)
table_kb.style = 'Light Grid Accent 1'
table_kb.alignment = WD_TABLE_ALIGNMENT.CENTER

hdr = table_kb.rows[0].cells
hdr[0].text = 'Параметр'
hdr[1].text = 'Значение'
for cell in hdr:
    for p in cell.paragraphs:
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

for param, val in kb_data:
    row = table_kb.add_row()
    row.cells[0].text = param
    row.cells[1].text = val
    for cell in row.cells:
        for p in cell.paragraphs:
            for r in p.runs:
                r.font.size = Pt(10)

doc.add_paragraph()

# Section 2: Comparison table
doc.add_heading('Б.2. Сверка сертификатов: основной документ vs база знаний', level=2)

comp_data = [
    ('1', 'Компаунд эпоксидный EMPILS E 608', '✅ Найден', 'Сертификат ТР № 123-ФЗ', 'Г1, В1, Д1, Т1, РП1'),
    ('2', 'Раствор водный ББ-11', '✅ Найден', 'Сертификат ТР ЕАЭС 043-2017', 'Г1, В1, Д1, Т1, РП1'),
    ('3', 'Эмаль акриловая матовая СТО 062-2010', '❌ НЕ найден', 'Только ПБ (паспорт безопасности)', 'Г1, В2, Д2, Т2 (из предыдущей сессии)'),
    ('4', 'Эмаль акриловая глянцевая СТО 062-2010', '✅ Найден', 'ДС (пожарн.)', 'Г1, В1, Д1, Т1'),
    ('5', 'Эмаль СТО 136-2018', '✅ Найден', 'ДС (пожарн.)', 'Г1, В2, Д2, Т2'),
    ('6', 'Краска ОРЕОЛ белая моющаяся ВД полиакриловая', '✅ Найден', 'ДС (пожарн.)', 'КМ1 (Г1, В1, Д1, Т1, РП1)'),
    ('7', 'Эмаль для пола акриловая СТО 060-2010', '✅ Найден', 'ДС (пожарн.)', 'КМ1 (Г1, В1, Д1, Т1)'),
    ('8', 'NEWTONE INTERIO СТО 096-2013', '✅ Найден', 'ДС (пожарн.)', 'Г1, В1, Д2, Т2, РП1'),
    ('9', 'NEWTONE LOTUS 3 СТО 096-2013', '✅ Найден', 'ДС (пожарн.)', 'Г1, В1, Д2, Т2, РП1'),
    ('10', 'NEWTONE LOTUS 20 СТО 096-2013', '✅ Найден', 'ДС (пожарн.)', 'Г1, В1, Д2, Т2, РП1'),
    ('11', 'NEWTONE LOTUS 7 СТО 096-2013', '✅ Найден', 'ДС (пожарн.)', 'Г1, В1, Д2, Т2, РП1'),
    ('12', 'Краска фасадная СТО 135-2018', '✅ Найден', 'ДС', 'Г1, В1, Д2, Т2, РП1 (КМ1)'),
    ('13', 'Краска интерьерная моющаяся СТО 134-2018', '✅ Найден', 'ДС', 'КМ1 (Г1, В1, Д2, Т2, РП1)'),
]

table_comp = doc.add_table(rows=1, cols=6)
table_comp.style = 'Light Grid Accent 1'
table_comp.alignment = WD_TABLE_ALIGNMENT.CENTER

headers = ['№', 'Продукт', 'В базе знаний?', 'Тип документа', 'Классы пожароопасности']
hdr = table_comp.rows[0].cells
for i, h in enumerate(headers[:5]):
    hdr[i].text = h
    for p in hdr[i].paragraphs:
        for r in p.runs:
            r.bold = True
            r.font.size = Pt(8)
            r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    tc = hdr[i]._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), '1F4E79')
    tcPr.append(shd)

for row_data in comp_data:
    row = table_comp.add_row()
    for i, val in enumerate(row_data[:5]):
        cell = row.cells[i]
        cell.text = val
        for p in cell.paragraphs:
            for r in p.runs:
                r.font.size = Pt(8)

doc.add_paragraph()

# Section 3: Key findings
doc.add_heading('Б.3. Ключевые выводы сверки', level=2)

findings = """
🔍 РЕЗУЛЬТАТЫ СВЕРКИ:

1. СОВПАДЕНИЕ: 12 из 13 сертификатов найдены в базе знаний ✅
   База знаний содержит все ключевые ДС по пожарной безопасности.

2. ОТСУТСТВУЕТ: «ДС (пожарн.) Эмаль акриловая матовая СТО 05744283-062-2010» ❌
   В базе знаний есть:
   • «ДС (пожарн.) Эмаль акриловая глянцевая СТО 05744283-062-2010» (глянцевая версия)
   • «ПБ Эмаль акриловая СТО 05744283-062-2010.pdf» (паспорт безопасности, не ДС)
   Рекомендация: запросить у производителя отдельный ДС на матовую модификацию.

3. УЛУЧШЕНИЯ ПО КЛАССАМ (подтверждено сертификатами):
   • ОРЕОЛ ДИСКОНТ интерьерная: Г1, В1, Д2, Т2 → КМ1 (Г1, В1, Д1, Т1, РП1) — ЛУЧШЕ на 2 класса
   • ОРЕОЛ ДИСКОНТ моющаяся: Г1, В1, Д2, Т2 → КМ1 — ЛУЧШЕ на 2 класса
   • ОРЕОЛ ДИСКОНТ полиакрил: Г1, В1, Д2, Т2 → КМ1 — ЛУЧШЕ на 2 класса
   • ОРЕОЛ ДИСКОНТ фасадная: В2 → В1 (улучшено!), РП2 → РП1 (улучшено!)

4. УХУДШЕНИЯ (требуют внимания):
   • Эмаль акриловая матовая: В1 → В2 (ХУЖЕ!) — НЕ подходит для путей эвакуации
   • Эмаль СТО 136-2018: В1 → В2 — НЕ подходит для путей эвакуации

5. НОВЫЕ ПРОДУКТЫ (не были в основном документе ранее):
   • Раствор водный ББ-11 — огнезащитная пропитка (Г1, В1, Д1, Т1, РП1) — ТР ЕАЭС 043-2017
   • Компаунд эпоксидный EMPILS E 608 — покрытие полов (Г1, В1, Д1, Т1, РП1) — ТР 123-ФЗ

6. ПАСПОРТА БЕЗОПАСНОСТИ (ПБ):
   В базе знаний ~450 ПБ. ПБ не заменяют ДС, но содержат информацию о:
   • Классе опасности вещества (по ГОСТ 12.1.007)
   • ПДК и ОБУВ
   • Мерах пожарной безопасности при работе с продуктом
   ПБ используются для СИЗ и техники безопасности, но НЕ для проектирования зданий.
"""

p = doc.add_paragraph(findings)
p.paragraph_format.first_line_indent = Cm(0)
for r in p.runs:
    r.font.size = Pt(11)

# Section 4: Recommendations
doc.add_heading('Б.4. Рекомендации по актуализации', level=2)

recommendations = """
1. ✅ ОБНОВЛЕНЫ ТАБЛИЦЫ #2 и #6 — классы пожароопасности заменены на подтверждённые сертификатами
2. ✅ ДОБАВЛЕН РП1 ко всем продуктам (где подтверждено сертификатом)
3. ✅ УТОЧНЕНЫ КЛАССЫ:
   • ОРЕОЛ ДИСКОНТ интерьерная/моющаяся/полиакрил: КМ1 (вместо Г1, В1, Д2, Т2)
   • ОРЕОЛ ДИСКОНТ фасадная: В1 (вместо В2)
4. ⚠️ Эмаль акриловая матовая: В2 (НЕ В1!) — исключить из рекомендаций для путей эвакуации
5. 🆕 ДОБАВЛЕНЫ 2 НОВЫХ ПРОДУКТА: ББ-11 и EMPILS E 608
6. ❓ ЗАПРОСИТЬ У ПРОИЗВОДИТЕЛЯ:
   • ДС на «Эмаль акриловую матовую» (отсутствует в базе знаний)
   • ДС на «Эмаль для радиаторов акриловую» (отсутствует в базе знаний)
   • Подтверждение срока действия ДС на Краску фасадную (до 2026 г. — конкретная дата?)
"""

p = doc.add_paragraph(recommendations)
p.paragraph_format.first_line_indent = Cm(0)
for r in p.runs:
    r.font.size = Pt(11)

# Save
doc.save(OUTPUT_PATH)

file_size = os.path.getsize(OUTPUT_PATH)
print(f'✅ Файл сохранён: {OUTPUT_PATH}')
print(f'Размер: {file_size} байт ({file_size/1024:.1f} KB)')
print(f'Параграфов: {len(doc.paragraphs)}')
print(f'Таблиц: {len(doc.tables)}')
