"""Полный цикл: скачивание фото + сборка DOCX-портфолио TITAN."""
import os
import urllib.request
import json
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

PHOTOS_DIR = '/home/z/my-project/download/titan_objects_photos'
os.makedirs(PHOTOS_DIR, exist_ok=True)

# Финальный список объектов с найденными фото-URL
objects_final = [
    {
        'id': 1,
        'name': 'Ретрит-центр SABI CLUB Валдай',
        'url': 'https://sabivalday.com/placement',
        'location': 'Валдай, Новгородская область, Центральная Россия',
        'description': (
            'Ретрит-центр «SABI Club» — закрытая территория в живописном лесу на берегу озера '
            'Валдай. На территории — комфортабельные дома с детально продуманным дизайнерским '
            'интерьером и аутентичные юрты. Площадка для проведения ретритов, семинаров и '
            'корпоративных мероприятий.'
        ),
        'materials': 'Грунт-эмаль TITAN (металлоконструкции)',
        'images': [
            'https://static.tildacdn.com/tild6234-6434-4162-b532-353166346630/photo.jpg',
            'https://static.tildacdn.com/tild3331-3434-4939-a138-303832333164/____.jpg',
            'https://static.tildacdn.com/tild6234-3339-4636-b861-356337626665/__1.jpg',
            'https://static.tildacdn.com/tild6661-3435-4562-a464-636461323261/__.jpg',
            'https://static.tildacdn.com/tild6165-3463-4234-a237-346136336361/3_2.jpg',
            'https://static.tildacdn.com/tild3530-3635-4134-b231-613665343030/IMG_6182.jpg',
        ],
    },
    {
        'id': 2,
        'name': 'Эко-парк Soloma',
        'url': 'https://soloma27.ru/',
        'location': 'Хабаровский край, Дальний Восток',
        'description': (
            'Эко-парк Soloma — безмятежный отдых в полном единении с природой. Комплекс '
            'деревянных эко-домов и инфраструктуры в экологическом стиле. Использованы '
            'комбинированные материалы: металлические конструкции с покрытием TITAN, '
            'деревянные рейки AVANTGARDE, яхтный лак PROFILUX.'
        ),
        'materials': 'TITAN (металл) + AVANTGARDE (деревянные рейки) + PROFILUX (яхтный лак)',
        'images': [
            'https://static.tildacdn.com/tild3637-3461-4437-a136-333139653339/3.png',
            'https://static.tildacdn.com/tild6632-6338-4265-b861-616136363663/photo.png',
            'https://static.tildacdn.com/tild6430-3634-4466-b632-656565633364/__.jpg',
            'https://static.tildacdn.com/tild6334-3861-4334-b338-343432633362/photo.jpg',
            'https://static.tildacdn.com/tild3935-6362-4535-b939-666138653264/photo.jpg',
            'https://static.tildacdn.com/tild6362-3564-4166-b235-376538633462/photo.jpg',
            'https://static.tildacdn.com/tild6438-6632-4265-a336-303666363363/photo.PNG',
            'https://static.tildacdn.com/tild3736-3863-4337-a439-646234306631/2.png',
        ],
    },
    {
        'id': 3,
        'name': 'Баня-юрта BraMari',
        'url': 'https://banimed.ru/bramari/',
        'location': 'Челябинск',
        'description': (
            'Баня-юрта BraMari — уникальный проект круглой бани-юрты. Каркас и металлические '
            'конструкции обработаны грунт-эмалью TITAN по ржавчине (3 в 1). Внешняя обшивка и '
            'декоративные элементы выполнены с использованием краски AVANTGARDE. На странице '
            'объекта есть фото каркаса до обшивки — показан этап антикоррозийной защиты.'
        ),
        'materials': 'TITAN (каркас, металл) + AVANTGARDE (декор)',
        'images': [
            'https://banimed.ru/wp-content/uploads/2026/03/kpb8pxgrzbgq8jhqitzjp-1.jpg',
            'https://banimed.ru/wp-content/uploads/2026/03/ffcr8klsurtrj484clrhq-1.jpg',
            'https://banimed.ru/wp-content/uploads/2026/04/img_0842.jpeg',
            'https://banimed.ru/wp-content/uploads/2026/04/2026-02-28-10.00.59.jpg',
            'https://banimed.ru/wp-content/uploads/2026/04/dscf5685.jpg',
            'https://banimed.ru/wp-content/uploads/2026/03/u_2.jpg',
            'https://banimed.ru/wp-content/uploads/2026/03/u_3.jpg',
            'https://banimed.ru/wp-content/uploads/2026/03/u_1.jpg',
        ],
    },
    {
        'id': 4,
        'name': 'База отдыха «Рыбацкие просторы» — Юрт Хаус',
        'url': 'https://fishprostor.ru/yurt-house.html',
        'location': 'Астраханская область',
        'description': (
            'Юрт Хаус на базе отдыха «Рыбацкие просторы» — современный комфорт в атмосфере '
            'этнического стиля. Пение птиц, свежий речной воздух и уют современной юрты на '
            'берегу Волги. Каркас юрты — металлический, обработан грунт-эмалью TITAN 3 в 1 по '
            'ржавчине. Есть фото каркаса на странице объекта.'
        ),
        'materials': 'Грунт-эмаль TITAN (металлический каркас юрты)',
        'images': [
            'https://fishprostor.ru/img/yurt/3d.jpg',
            'https://fishprostor.ru/img/yurt/header-yort.jpg',
            'https://fishprostor.ru/img/yurt/3.jpg',
            'https://fishprostor.ru/img/yurt/2.jpg',
            'https://fishprostor.ru/img/yurt/9.jpg',
            'https://fishprostor.ru/img/yurt/7.jpg',
            'https://fishprostor.ru/img/yurt/12.jpg',
            'https://fishprostor.ru/img/yurt/10.jpg',
            'https://fishprostor.ru/img/yurt/11.jpg',
        ],
    },
    {
        'id': 5,
        'name': 'Агротуристический комплекс «Голубая лагуна»',
        'url': 'https://vk.ru/lagunabuida',
        'location': 'с. Буйда, Республика Башкортостан',
        'description': (
            'Агротуристический комплекс «Голубая лагуна» — место семейного отдыха на природе. '
            'Лес, поля, пруд. На территории — бани, беседки, места для рыбалки и активного '
            'отдыха. Металлические конструкции (беседки, навесы, ограждения) обработаны '
            'грунт-эмалью TITAN 3 в 1 по ржавчине для долговечной защиты от атмосферных '
            'воздействий.'
        ),
        'materials': 'Грунт-эмаль TITAN (металлоконструкции, навесы, ограждения)',
        'images': [
            'https://sun1-25.vkuserphoto.ru/ZdV8d-XHc_5GsnwSSBvJSCZ7Dv5TUEdZuQILrw/dTxtHF9uF1M.png',
            'https://sun1-25.vkuserphoto.ru/6DD0xSghpEY4ZOnSoOQGo5hvOqZld01WrEb-dA/I-cI-hyWOi0.png',
        ],
    },
]

# === СКАЧИВАНИЕ ФОТО ===
print("=" * 80)
print("Шаг 1. Скачивание фотографий")
print("=" * 80)

for obj in objects_final:
    obj_dir = os.path.join(PHOTOS_DIR, f"obj{obj['id']}")
    os.makedirs(obj_dir, exist_ok=True)
    print(f"\nОбъект {obj['id']}: {obj['name']}")
    downloaded = []
    for i, img_url in enumerate(obj['images'], 1):
        ext = '.jpg'
        for e in ['.png', '.jpeg', '.jpg', '.webp', '.PNG']:
            if e in img_url:
                ext = e.lower()
                if ext == '.jpeg':
                    ext = '.jpg'
                break
        out_path = os.path.join(obj_dir, f"photo_{i:02d}{ext}")
        try:
            req = urllib.request.Request(img_url, headers={
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
            })
            with urllib.request.urlopen(req, timeout=20) as resp:
                if resp.status == 200:
                    data = resp.read()
                    if len(data) > 3000:
                        with open(out_path, 'wb') as f:
                            f.write(data)
                        downloaded.append(out_path)
                        print(f"  ✓ {i}. {os.path.basename(out_path)} ({len(data)//1024} KB)")
                    else:
                        print(f"  ✗ {i}. Слишком маленький ({len(data)} байт)")
        except Exception as e:
            print(f"  ✗ {i}. Ошибка: {str(e)[:60]}")
    obj['local_images'] = downloaded

total_photos = sum(len(o.get('local_images', [])) for o in objects_final)
print(f"\nСкачано {total_photos} фото для {len(objects_final)} объектов")

# === СОЗДАНИЕ DOCX ===
print("\n" + "=" * 80)
print("Шаг 2. Сборка DOCX-портфолио")
print("=" * 80)

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

style = doc.styles['Normal']
style.font.name = 'Arial'
style.font.size = Pt(10)
style.font.color.rgb = COLOR_BODY

def shade_cell(cell, hex_color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    tcPr.append(shd)

# === ОБЛОЖКА ===
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

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run('━' * 30)
run.font.color.rgb = COLOR_ACCENT

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run(f'5 объектов · {total_photos} фото')
run.font.size = Pt(10)
run.font.color.rgb = COLOR_SECONDARY

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run('Сентябрь 2026')
run.font.size = Pt(10)
run.font.italic = True
run.font.color.rgb = COLOR_SECONDARY

doc.add_paragraph()

# === СВОДНАЯ ТАБЛИЦА ===
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

for obj in objects_final:
    row = table.add_row()
    cells = row.cells
    cells[0].text = ''
    p = cells[0].paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(str(obj['id']))
    run.font.bold = True
    run.font.size = Pt(11)
    run.font.color.rgb = COLOR_ACCENT

    cells[1].text = ''
    p = cells[1].paragraphs[0]
    run = p.add_run(obj['name'])
    run.font.bold = True
    run.font.size = Pt(10)

    cells[2].text = ''
    p = cells[2].paragraphs[0]
    run = p.add_run(obj['location'])
    run.font.size = Pt(9)

    cells[3].text = ''
    p = cells[3].paragraphs[0]
    run = p.add_run(obj['materials'])
    run.font.size = Pt(9)
    run.font.color.rgb = COLOR_GREEN

widths = [Cm(0.8), Cm(6.5), Cm(4.5), Cm(6.5)]
for row in table.rows:
    for i, w in enumerate(widths):
        if i < len(row.cells):
            row.cells[i].width = w

doc.add_paragraph()

# === РАЗДЕЛЫ ПО ОБЪЕКТАМ ===
for obj_idx, obj in enumerate(objects_final):
    if obj_idx > 0:
        doc.add_page_break()

    # Заголовок
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

    # Материалы
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

    # Фото
    images = obj.get('local_images', [])
    images_to_show = images[:6]

    p = doc.add_paragraph()
    run = p.add_run(f'Фото объекта ({len(images)} изображений)')
    run.font.bold = True
    run.font.size = Pt(11)
    run.font.color.rgb = COLOR_PRIMARY

    if images_to_show:
        rows_count = (len(images_to_show) + 1) // 2
        photo_table = doc.add_table(rows=rows_count, cols=2)
        photo_table.alignment = WD_TABLE_ALIGNMENT.CENTER

        for i, img_path in enumerate(images_to_show):
            row_idx = i // 2
            col_idx = i % 2
            cell = photo_table.rows[row_idx].cells[col_idx]
            cell.text = ''
            cell.width = Cm(8)
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

    doc.add_paragraph()

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

doc.add_paragraph()
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run('━' * 30)
run.font.color.rgb = COLOR_ACCENT

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
print(f'\n✅ DOCX сохранён: {out_path}')
size_kb = os.path.getsize(out_path) // 1024
print(f'Размер файла: {size_kb} KB')
print(f'\nСтруктура:')
print(f'  Объектов: {len(objects_final)}')
print(f'  Фото всего: {total_photos}')
for obj in objects_final:
    print(f'  {obj["id"]}. {obj["name"]}: {len(obj.get("local_images", []))} фото')
