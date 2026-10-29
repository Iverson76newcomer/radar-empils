"""Добавление объекта «Тюрьма Черный дельфин» в существующий DOCX."""
import os
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from copy import deepcopy

SRC = '/home/z/my-project/upload/TITAN_Портфолио_объектов_18.09.2026.docx'
OUT = '/home/z/my-project/download/TITAN_Портфолио_объектов_18.09.2026.docx'

doc = Document(SRC)

COLOR_PRIMARY = RGBColor(0x0B, 0x1C, 0x2C)
COLOR_ACCENT = RGBColor(0xF4, 0x76, 0x29)
COLOR_BODY = RGBColor(0x1A, 0x2B, 0x3D)
COLOR_SECONDARY = RGBColor(0x5B, 0x6B, 0x7D)
COLOR_GREEN = RGBColor(0x16, 0xA3, 0x4A)
COLOR_WHITE = RGBColor(0xFF, 0xFF, 0xFF)

def shade_cell(cell, hex_color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    tcPr.append(shd)

# === ШАГ 1: Добавить строку в сводную таблицу ===
# Сводная таблица — первая таблица в документе
summary_table = doc.tables[0]

# Добавляем новую строку с данными о Чёрном дельфине
new_row = summary_table.add_row()
cells = new_row.cells
# №
cells[0].text = ''
p = cells[0].paragraphs[0]
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run('6')
run.font.bold = True
run.font.size = Pt(11)
run.font.color.rgb = COLOR_ACCENT
# Название
cells[1].text = ''
p = cells[1].paragraphs[0]
run = p.add_run('Тюрьма «Чёрный дельфин»')
run.font.bold = True
run.font.size = Pt(10)
# Локация
cells[2].text = ''
p = cells[2].paragraphs[0]
run = p.add_run('Соль-Илецк, Оренбургская область')
run.font.size = Pt(9)
# Материалы
cells[3].text = ''
p = cells[3].paragraphs[0]
run = p.add_run('Чёрная молотковая эмаль TITAN (металлоконструкции, ограждения, корпуса)')
run.font.size = Pt(9)
run.font.color.rgb = COLOR_GREEN

# Скопируем ширины колонок из предыдущей строки
prev_row = summary_table.rows[-2]
for i, cell in enumerate(prev_row.cells):
    if i < len(cells):
        cells[i].width = cell.width

# === ШАГ 2: Добавить полноценный раздел объекта в конец документа ===
# Сначала добавим page break
p = doc.add_paragraph()
run = p.add_run()
run.add_break()

# Заголовок объекта
p = doc.add_paragraph()
run = p.add_run('Объект 6. Тюрьма «Чёрный дельфин»')
run.font.size = Pt(16)
run.font.bold = True
run.font.color.rgb = COLOR_PRIMARY

# Локация
p = doc.add_paragraph()
run = p.add_run('📍 ')
run.font.size = Pt(11)
run = p.add_run('Соль-Илецк, Оренбургская область')
run.font.size = Pt(11)
run.font.bold = True
run.font.color.rgb = COLOR_ACCENT

# Источник
p = doc.add_paragraph()
run = p.add_run('🔗 ')
run.font.size = Pt(10)
run = p.add_run('https://ru.wikipedia.org/wiki/Чёрный_дельфин')
run.font.size = Pt(9)
run.font.color.rgb = COLOR_SECONDARY
run.font.italic = True

# Материалы
p = doc.add_paragraph()
run = p.add_run('🎨 Применённые материалы: ')
run.font.bold = True
run.font.size = Pt(11)
run.font.color.rgb = COLOR_PRIMARY
run = p.add_run('Чёрная молотковая эмаль TITAN (металлоконструкции, ограждения, периметр, корпуса)')
run.font.size = Pt(11)
run.font.color.rgb = COLOR_GREEN
run.font.bold = True

# Описание
p = doc.add_paragraph()
run = p.add_run('Описание объекта')
run.font.bold = True
run.font.size = Pt(11)
run.font.color.rgb = COLOR_PRIMARY

description = (
    'Исправительная колония ФКУ ИК-6 УФСИН России по Оренбургской области, известная как '
    '«Чёрный дельфин» — одна из крупнейших колоний особого режима для пожизненно осуждённых '
    'в России. Расположена в городе Соль-Илецк Оренбургской области. Лимит наполнения — '
    '1600 человек. На объекте применена чёрная молотковая эмаль TITAN для антикоррозийной '
    'защиты металлических конструкций: ограждений периметра, решёток на окнах, металлических '
    'дверей, ворот, внутренних металлоконструкций корпусов. Молотковый эффект краски '
    'обеспечивает не только защитные, но и эстетические свойства — маскирует неровности '
    'поверхности и придаёт благородный матовый чёрный оттенок.'
)
p = doc.add_paragraph(description)
p.paragraph_format.first_line_indent = Cm(0.5)
p.paragraph_format.line_spacing = 1.4
p.paragraph_format.space_after = Pt(12)
for run in p.runs:
    run.font.size = Pt(10)

# Фото объекта
p = doc.add_paragraph()
run = p.add_run('Фото объекта (5 изображений)')
run.font.bold = True
run.font.size = Pt(11)
run.font.color.rgb = COLOR_PRIMARY

# Отобранные фото (пропускаем неподходящие)
photos_to_show = [
    ('/home/z/my-project/download/titan_objects_photos/obj6/photo_02.jpg', 'Силуэт дельфина — символ колонии'),
    ('/home/z/my-project/download/titan_objects_photos/obj6/photo_05.jpg', 'Главный корпус колонии'),
    ('/home/z/my-project/download/titan_objects_photos/obj6/photo_06.jpg', 'Комплекс зданий с решётками на окнах'),
    ('/home/z/my-project/download/titan_objects_photos/obj6/photo_01.jpg', 'Ограждение периметра с колючей проволокой'),
    ('/home/z/my-project/download/titan_objects_photos/obj6/photo_07.jpg', 'Интерьер — помещение изолятора'),
]

# Таблица 2x3 (или 2x2 + 1)
rows_count = (len(photos_to_show) + 1) // 2
photo_table = doc.add_table(rows=rows_count, cols=2)
photo_table.alignment = WD_TABLE_ALIGNMENT.CENTER

for i, (img_path, caption) in enumerate(photos_to_show):
    row_idx = i // 2
    col_idx = i % 2
    cell = photo_table.rows[row_idx].cells[col_idx]
    cell.text = ''
    cell.width = Cm(8)
    # Изображение
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    try:
        run = p.add_run()
        run.add_picture(img_path, width=Cm(7.5))
    except Exception as e:
        run = p.add_run(f'[Фото не загружено]')
        run.font.size = Pt(9)
        run.font.italic = True
        run.font.color.rgb = COLOR_SECONDARY
    # Подпись под фото
    p2 = cell.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p2.add_run(caption)
    run.font.size = Pt(8)
    run.font.italic = True
    run.font.color.rgb = COLOR_SECONDARY

# Если нечётное — последняя ячейка пустая
if len(photos_to_show) % 2 == 1:
    last_cell = photo_table.rows[-1].cells[1]
    last_cell.text = ''

# Источник
doc.add_paragraph()
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
run = p.add_run('Источники: ')
run.font.size = Pt(9)
run.font.italic = True
run.font.color.rgb = COLOR_SECONDARY
run = p.add_run('lada.rest, entermedia.io, lapsha.media, lemur59.ru')
run.font.size = Pt(9)
run.font.italic = True
run.font.color.rgb = COLOR_ACCENT
run.font.underline = True

# Сохраняем
doc.save(OUT)
print(f'✅ Сохранено: {OUT}')
size_kb = os.path.getsize(OUT) // 1024
print(f'Размер файла: {size_kb} KB')
