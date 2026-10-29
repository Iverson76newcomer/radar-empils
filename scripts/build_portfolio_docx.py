"""Создание DOCX-портфолио объектов TITAN с фото."""
from docx import Document
from docx.shared import Pt, Cm, RGBColor, Inches, Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import os
import json

with open('/home/z/my-project/scripts/objects_final.json', 'r', encoding='utf-8') as f:
    objects = json.load(f)

# Палитра TITAN
COLOR_PRIMARY = RGBColor(0x0B, 0x1C, 0x2C)
COLOR_ACCENT = RGBColor(0xF4, 0x76, 0x29)
COLOR_BODY = RGBColor(0x1A, 0x2B, 0x3D)
COLOR_SECONDARY = RGBColor(0x5B, 0x6B, 0x7D)
COLOR_WHITE = RGBColor(0xFF, 0xFF, 0xFF)
COLOR_GREEN = RGBColor(0x16, 0xA3, 0x4A)

doc = Document()
section = doc.sections[0]
section.top_margin = Cm(2.0)
section.bottom_margin = Cm(2.0)
section.left_margin = Cm(2.0)
section.right_margin = Cm(2.0)

# Стиль Normal
style = doc.styles['Normal']
style.font.name = 'Arial'
style.font.size = Pt(10)
style.font.color.rgb = COLOR_BODY

def shade_cell(cell, hex_color):
    """Закрасить ячейку цветом."""
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    tcPr.append(shd)

def set_cell_borders(cell, color='D0D0D0', size='4'):
    """Установить тонкие границы ячейки."""
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for border_name in ['top', 'left', 'bottom', 'right']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'single')
        border.set(qn('w:sz'), size)
        border.set(qn('w:color'), color)
        tcBorders.append(border)
    tcPr.append(tcBorders)

# === ОБЛОЖКА / ЗАГОЛОВОК ===
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run('TITAN')
run.font.size = Pt(32)
run.font.bold = True
run.font.color.rgb = COLOR_ACCENT

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run('Портфолио объектов')
run.font.size = Pt(20)
run.font.bold = True
run.font.color.rgb = COLOR_PRIMARY

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run('Объекты, где применялась грунт-эмаль TITAN по металлу')
run.font.size = Pt(12)
run.font.italic = True
run.font.color.rgb = COLOR_SECONDARY

# Разделитель
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run('━' * 30)
run.font.color.rgb = COLOR_ACCENT
run.font.size = Pt(11)

# Краткое описание
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run(f'5 объектов · 33 фото · {sum(len(o.get("local_images", [])) for o in objects)} изображений')
run.font.size = Pt(10)
run.font.color.rgb = COLOR_SECONDARY

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run('Дата составления: сентябрь 2026')
run.font.size = Pt(10)
run.font.italic = True
run.font.color.rgb = COLOR_SECONDARY

doc.add_paragraph()

# === СВОДНАЯ ТАБЛИЦА ОБЪЕКТОВ ===
p = doc.add_paragraph()
run = p.add_run('Сводный список объектов')
run.font.size = Pt(14)
run.font.bold = True
run.font.color.rgb = COLOR_PRIMARY

table = doc.add_table(rows=1, cols=4)
table.alignment = WD_TABLE_ALIGNMENT.CENTER
table.style = 'Light Grid Accent 1'
hdr = table.rows[0].cells
headers = ['№', 'Название объекта', 'Локация', 'Применённые материалы']
for i, h in enumerate(headers):
    hdr[i].text = ''
    p = hdr[i].paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(h)
    run.font.bold = True
    run.font.size = Pt(10)
    run.font.color.rgb = COLOR_WHITE
    shade_cell(hdr[i], '0B1C2C')

for obj in objects:
    row = table.add_row()
    cells = row.cells
    # №
    cells[0].text = ''
    p = cells[0].paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(str(obj['id']))
    run.font.bold = True
    run.font.size = Pt(11)
    run.font.color.rgb = COLOR_ACCENT

    # Название
    cells[1].text = ''
    p = cells[1].paragraphs[0]
    run = p.add_run(obj['name'])
    run.font.bold = True
    run.font.size = Pt(10)

    # Локация
    cells[2].text = ''
    p = cells[2].paragraphs[0]
    run = p.add_run(obj['location'])
    run.font.size = Pt(9)

    # Материалы
    cells[3].text = ''
    p = cells[3].paragraphs[0]
    run = p.add_run(obj['materials'])
    run.font.size = Pt(9)
    run.font.color.rgb = COLOR_GREEN

# Ширины
widths = [Cm(0.8), Cm(6.5), Cm(4.5), Cm(6.5)]
for row in table.rows:
    for i, w in enumerate(widths):
        if i < len(row.cells):
            row.cells[i].width = w

doc.add_paragraph()

# === РАЗДЕЛЫ ПО ОБЪЕКТАМ ===
for obj_idx, obj in enumerate(objects):
    if obj_idx > 0:
        doc.add_page_break()

    # Заголовок объекта
    p = doc.add_paragraph()
    run = p.add_run(f'Объект {obj["id"]}. {obj["name"]}')
    run.font.size = Pt(16)
    run.font.bold = True
    run.font.color.rgb = COLOR_PRIMARY

    # Локация
    p = doc.add_paragraph()
    run = p.add_run('📍 ')
    run.font.size = Pt(11)
    run = p.add_run(obj['location'])
    run.font.size = Pt(11)
    run.font.bold = True
    run.font.color.rgb = COLOR_ACCENT

    # Источник
    p = doc.add_paragraph()
    run = p.add_run('🔗 ')
    run.font.size = Pt(10)
    run = p.add_run(obj['url'])
    run.font.size = Pt(9)
    run.font.color.rgb = COLOR_SECONDARY
    run.font.italic = True

    # Применённые материалы
    p = doc.add_paragraph()
    run = p.add_run('🎨 Применённые материалы: ')
    run.font.bold = True
    run.font.size = Pt(11)
    run.font.color.rgb = COLOR_PRIMARY
    run = p.add_run(obj['materials'])
    run.font.size = Pt(11)
    run.font.color.rgb = COLOR_GREEN
    run.font.bold = True

    # Описание
    p = doc.add_paragraph()
    run = p.add_run('Описание объекта')
    run.font.bold = True
    run.font.size = Pt(11)
    run.font.color.rgb = COLOR_PRIMARY

    p = doc.add_paragraph(obj['description'])
    p.paragraph_format.first_line_indent = Cm(0.5)
    p.paragraph_format.line_spacing = 1.4
    p.paragraph_format.space_after = Pt(12)
    for run in p.runs:
        run.font.size = Pt(10)

    # Фотографии
    p = doc.add_paragraph()
    run = p.add_run(f'Фото объекта ({len(obj.get("local_images", []))} изображений)')
    run.font.bold = True
    run.font.size = Pt(11)
    run.font.color.rgb = COLOR_PRIMARY

    # Размещаем фото в таблице 2x2 или 2x3
    images = obj.get('local_images', [])
    # Берём максимум 6 фото для документа
    images_to_show = images[:6]

    if images_to_show:
        # Таблица 2 колонки × 3 ряда
        rows_count = (len(images_to_show) + 1) // 2
        photo_table = doc.add_table(rows=rows_count, cols=2)
        photo_table.alignment = WD_TABLE_ALIGNMENT.CENTER

        for i, img_path in enumerate(images_to_show):
            row_idx = i // 2
            col_idx = i % 2
            cell = photo_table.rows[row_idx].cells[col_idx]
            cell.text = ''
            # Ширина колонки
            cell.width = Cm(8)
            # Добавляем параграф с изображением
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            try:
                run = p.add_run()
                run.add_picture(img_path, width=Cm(7.5))
            except Exception as e:
                run = p.add_run(f'[Фото {i+1}: не удалось вставить]')
                run.font.size = Pt(9)
                run.font.italic = True
                run.font.color.rgb = COLOR_SECONDARY
            # Границы ячейки
            set_cell_borders(cell, 'FFFFFF', '4')

        # Если нечётное количество — добавляем пустую ячейку
        if len(images_to_show) % 2 == 1:
            last_row = photo_table.rows[-1]
            empty_cell = last_row.cells[1]
            empty_cell.text = ''

        # Заполняем недостающие ячейки в последнем ряду
        for ri in range(rows_count):
            for ci in range(2):
                cell = photo_table.rows[ri].cells[ci]
                if not cell.text and not any(p.runs for p in cell.paragraphs if any(r.element.xpath('.//pic:pic', namespaces={'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture'}) for r in p.runs)):
                    pass  # уже пустая

    doc.add_paragraph()

    # Контакты / ссылка на оригинал
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = p.add_run('Источник: ')
    run.font.size = Pt(9)
    run.font.italic = True
    run.font.color.rgb = COLOR_SECONDARY
    run = p.add_run(obj['url'])
    run.font.size = Pt(9)
    run.font.italic = True
    run.font.color.rgb = COLOR_ACCENT
    run.font.underline = True

# === ЗАКЛЮЧЕНИЕ ===
doc.add_page_break()
p = doc.add_paragraph()
run = p.add_run('Заключение')
run.font.size = Pt(16)
run.font.bold = True
run.font.color.rgb = COLOR_PRIMARY

conclusion_text = (
    'Грунт-эмаль TITAN 3 в 1 по ржавчине применена на 5 объектах разного типа: '
    'ретрит-центр, эко-парк, баня-юрта, база отдыха, агротуристический комплекс. '
    'Объекты расположены в разных климатических зонах — от Центральной России '
    '(Валдай) до Дальнего Востока (Хабаровский край) и Урала (Башкортостан). '
    'Это подтверждает универсальность продукта для защиты металлоконструкций в '
    'условиях атмосферных воздействий и переменного климата.'
)
p = doc.add_paragraph(conclusion_text)
p.paragraph_format.first_line_indent = Cm(0.5)
p.paragraph_format.line_spacing = 1.4
for run in p.runs:
    run.font.size = Pt(11)

# Преимущества
p = doc.add_paragraph()
run = p.add_run('Подтверждённые преимущества TITAN по результатам эксплуатации на объектах:')
run.font.size = Pt(11)
run.font.bold = True
run.font.color.rgb = COLOR_PRIMARY

advantages = [
    'Антикоррозийная защита металлоконструкций в условиях переменного климата (от -40°C до +80°C)',
    'Универсальность применения: каркасы юрт, навесы, ограждения, беседки, металлические элементы эко-домов',
    'Возможность нанесения по ржавчине без сложной подготовки поверхности (3 в 1: грунт + эмаль + защита)',
    'Долговечность покрытия: объекты эксплуатируются в течение нескольких сезонов без повторной покраски',
    'Сочетаемость с другими продуктами бренда: AVANTGARDE (дерево), PROFILUX (яхтный лак)',
    'География применения: от Центральной России до Дальнего Востока — подходит для всех климатических зон',
]
for adv in advantages:
    p = doc.add_paragraph(style='List Bullet')
    run = p.add_run(adv)
    run.font.size = Pt(10)

# Финальный блок
doc.add_paragraph()
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run('━' * 30)
run.font.color.rgb = COLOR_ACCENT
run.font.size = Pt(11)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run('ЗАО «Эмпилс» · Бренд TITAN · titanpaints.ru')
run.font.size = Pt(10)
run.font.bold = True
run.font.color.rgb = COLOR_PRIMARY

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run('Сентябрь 2026')
run.font.size = Pt(9)
run.font.italic = True
run.font.color.rgb = COLOR_SECONDARY

# Сохраняем
out_path = '/home/z/my-project/download/TITAN_Портфолио_объектов.docx'
doc.save(out_path)
print(f'✅ Сохранено: {out_path}')

# Проверка размера
size_kb = os.path.getsize(out_path) // 1024
print(f'Размер файла: {size_kb} KB')
print(f'\nСтруктура:')
print(f'  Объектов: {len(objects)}')
print(f'  Фото всего: {sum(len(o.get("local_images", [])) for o in objects)}')
for obj in objects:
    print(f'  {obj["id"]}. {obj["name"]}: {len(obj.get("local_images", []))} фото')
