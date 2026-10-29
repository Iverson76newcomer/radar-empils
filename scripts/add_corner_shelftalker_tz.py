#!/usr/bin/env python3
"""
Добавить в ТЗ описание углового шелфтокера (под банку, информацией вверх).
Вставляется после секции "1. Шелфтокер" и перед "2. Воблер".
"""

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import copy

INPUT = '/home/z/my-project/upload/ТЗ_POSM_TITAN_шелфтокер_воблер_ценник.docx'
OUTPUT = '/home/z/my-project/download/ТЗ_POSM_TITAN_шелфтокер_воблер_ценник.docx'

RED = RGBColor(0xC6, 0x28, 0x28)
DARK = RGBColor(0x1A, 0x1F, 0x36)
GRAY = RGBColor(0x60, 0x60, 0x60)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)

# Читаем исходный документ
doc = Document(INPUT)

# Найдём индекс параграфа "2. Воблер" — перед ним вставим новую секцию
insert_before_idx = None
for i, p in enumerate(doc.paragraphs):
    if '2. Воблер' in p.text:
        insert_before_idx = i
        break

print(f'Вставка перед параграфом {insert_before_idx}: "{doc.paragraphs[insert_before_idx].text[:50]}"')


# === Создаём содержимое новой секции ===
# Используем XML для вставки параграфов и таблиц перед нужным параграфом

def insert_paragraph_before(doc, before_idx, text, size=11, bold=False, color=DARK,
                            align=WD_ALIGN_PARAGRAPH.LEFT, italic=False, space_after=6, space_before=0):
    """Вставить параграф перед указанным индексом."""
    ref_para = doc.paragraphs[before_idx]
    new_para = ref_para.insert_paragraph_before('')
    
    pf = new_para.paragraph_format
    pf.space_after = Pt(space_after)
    pf.space_before = Pt(space_before)
    pf.line_spacing = 1.3
    new_para.alignment = align
    
    if text:
        run = new_para.add_run(text)
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.italic = italic
        run.font.color.rgb = color
        run.font.name = 'Arial'
    
    return new_para


def insert_table_before(doc, before_idx, headers, rows, col_widths=None):
    """Вставить таблицу перед указанным параграфом."""
    ref_para = doc.paragraphs[before_idx]
    
    # Создаём таблицу
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.style = 'Table Grid'
    
    # Заголовки
    for ci, h in enumerate(headers):
        cell = table.cell(0, ci)
        cell.text = ''
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(h)
        run.font.size = Pt(10)
        run.font.bold = True
        run.font.color.rgb = WHITE
        run.font.name = 'Arial'
        shading = OxmlElement('w:shd')
        shading.set(qn('w:val'), 'clear')
        shading.set(qn('w:fill'), '1E3A5F')
        cell._element.get_or_add_tcPr().append(shading)
    
    # Данные
    for ri, row in enumerate(rows):
        for ci, val in enumerate(row):
            cell = table.cell(ri + 1, ci)
            cell.text = ''
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = p.add_run(str(val))
            run.font.size = Pt(10)
            run.font.name = 'Arial'
            run.font.color.rgb = DARK
            if ci == 0:
                run.font.bold = True
    
    # Перемещаем таблицу перед нужным параграфом
    # python-pptx не поддерживает прямое перемещение таблиц,
    # но мы можем переместить XML элемент
    ref_para._element.addprevious(table._tbl)
    
    return table


# === Заголовок секции ===
insert_paragraph_before(doc, insert_before_idx, '', space_after=4)

p = doc.paragraphs[insert_before_idx]
# Пустой параграф для разделителя
new_p = p.insert_paragraph_before('1.б. Шелфтокер угловой (под банку, информацией вверх) — техническое задание')
new_p.paragraph_format.space_before = Pt(12)
new_p.paragraph_format.space_after = Pt(6)
for run in new_p.runs:
    run.font.size = Pt(12)
    run.font.bold = True
    run.font.color.rgb = RED
    run.font.name = 'Arial'

# === Назначение ===
insert_paragraph_before(doc, insert_before_idx, 'Назначение', size=12, bold=True, color=RED, space_before=8, space_after=4)

insert_paragraph_before(doc, insert_before_idx,
    'Угловой шелфтокер устанавливается непосредственно под банку TITAN на полке. '
    'Нижняя часть (основание) подсовывается под дно банки — банка стоит на ней и фиксирует шелфтокер. '
    'Верхняя часть (лицевая панель) направлена вертикально вверх и возвышается над банкой, '
    'демонстрируя слоган и ключевую информацию покупателю. '
    'Не требует двустороннего скотча — держится за счёт веса самой банки.',
    size=11, space_after=6)

# === Текстовая информация ===
insert_paragraph_before(doc, insert_before_idx, 'Текстовая информация на макете', size=12, bold=True, color=RED, space_before=8, space_after=4)

insert_table_before(doc, insert_before_idx,
    ['Зона', 'Текст', 'Шрифт / Стиль'],
    [
        ['Логотип', 'TITAN', 'В верхней части лицевой панели. Белым на красном фоне.'],
        ['Слоган', 'Красишь 1 раз\nслужит 10 лет', 'Крупно, жирно. В 2 строки. Белым или жёлтым на красном.'],
        ['Доп. текст', 'Грунт-эмаль 3в1 по ржавчине', 'Мелким шрифтом, под слоганом.'],
        ['QR-код', 'QR-код → www.titanpaints.ru', 'В нижней части лицевой панели. Размер квадрата — не менее 20×20 мм.'],
    ],
    col_widths=[3.5, 7, 6]
)

# === Визуальные элементы ===
insert_paragraph_before(doc, insert_before_idx, 'Визуальные элементы', size=12, bold=True, color=RED, space_before=8, space_after=4)

elements = [
    '• Фон лицевой панели — красный (#C62828)',
    '• Основание (под банкой) — белое или красное, без текста',
    '• Логотип TITAN — в верхней части лицевой панели',
    '• Слоган — центр лицевой панели, крупно',
    '• QR-код — в нижней части лицевой панели, на белом квадрате-подложке',
    '• Жёлтая полоса (#FBBF24) — разделитель под логотипом',
    '• Форма — Г-образная (L-shape): основание горизонтальное + лицевая панель вертикальная',
    '• Угол между основанием и лицевой панелью — 90° (прямой)',
]
for el in elements:
    insert_paragraph_before(doc, insert_before_idx, el, size=11, space_after=2)

# === Технические параметры ===
insert_paragraph_before(doc, insert_before_idx, 'Технические параметры для типографии', size=12, bold=True, color=RED, space_before=8, space_after=4)

insert_table_before(doc, insert_before_idx,
    ['Параметр', 'Значение'],
    [
        ['Тип', 'Шелфтокер угловой (L-образный) — основание под банку, лицевая панель вверх'],
        ['Конструкция', 'Г-образная: горизонтальное основание + вертикальная лицевая панель (угол 90°)'],
        ['Лицевая панель (Ш × В)', '100 × 80 мм (примерно)'],
        ['Основание (Ш × Г)', '100 × 50 мм (подсунуть под банку 0.75 л)'],
        ['Высота над банкой', 'Лицевая панель выступает на 50–60 мм над верхней крышкой банки'],
        ['Материал', 'Картон мелованный 350 г/м² (жёсткий, держит форму)'],
        ['Покрытие', 'Ламинация глянцевая (лицевая сторона)'],
        ['Печать', '4+0 (полноцвет, лицевая сторона)'],
        ['Биговка', 'Одна линия биговки на стыке основания и лицевой панели (для сгиба 90°)'],
        ['Вырубка', 'По контуру Г-образной формы'],
        ['Крепление', 'Не требуется — основание фиксируется весом банки (0.75–2 л)'],
        ['Совместимость', 'Подходит для банок 0.75 л и 2 л (основание шириной 100 мм)'],
        ['Тираж (ориентир.)', '500 шт (пилот) → 2000+ шт (масштаб)'],
        ['Срок изготовления', '7-10 рабочих дней после утверждения макета'],
    ],
    col_widths=[5, 12]
)

# Пустая строка перед следующей секцией
insert_paragraph_before(doc, insert_before_idx, '', space_after=4)

# Сохраняем
doc.save(OUTPUT)
print(f'OK: {OUTPUT}')
