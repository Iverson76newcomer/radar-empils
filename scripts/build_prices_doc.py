"""Создание DOCX-документа с мониторингом цен TITAN и конкурентов.
   Правила: цены без скидки (если перечёркнута — берём её).
   Добавлено: Лакра 5в1 Грунт-эмаль (кат. 4), TOPGRADE (кат. 5).
   Все продукты — со ссылкой на сайт-источник для быстрой проверки.
"""
import json
from docx import Document
from docx.shared import Pt, Cm, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn, nsmap
from docx.oxml import OxmlElement

with open('/home/z/my-project/scripts/vsi_prices_final.json', 'r', encoding='utf-8') as f:
    data = json.load(f)

# Группировка по категориям
categories = {}
for r in data:
    cat = r['cat']
    if cat not in categories:
        categories[cat] = {'title': '', 'products': []}
    categories[cat]['products'].append(r)

# Заголовки категорий
cat_titles = {
    1: 'Категория 1. Грунт-эмаль 3в1 высокоглянцевая, чёрный RAL 9005',
    2: 'Категория 2. Быстросохнущая матовая, серая RAL 7040',
    3: 'Категория 3. Эмаль с молотковым эффектом, золотистая',
    4: 'Категория 4. Грунт-эмаль 5в1 (Лакра) — НОВАЯ',
    5: 'Категория 5. Грунт-эмаль TOPGRADE by Hammerite — НОВАЯ',
}

# Создаём документ
doc = Document()

# Настройка стилей
section = doc.sections[0]
section.top_margin = Cm(2.54)
section.bottom_margin = Cm(2.54)
section.left_margin = Cm(2.5)
section.right_margin = Cm(2.0)

# Стиль Normal
style = doc.styles['Normal']
style.font.name = 'Arial'
style.font.size = Pt(10)
style.font.color.rgb = RGBColor(0x1A, 0x2B, 0x3D)

# === Заголовок ===
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run('TITAN — Актуальные цены конкурентов')
run.font.size = Pt(18)
run.font.bold = True
run.font.color.rgb = RGBColor(0x0B, 0x1C, 0x2C)
run.font.name = 'Arial'

# Подзаголовок
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run('Источник: vseinstrumenti.ru + альтернативные магазины (для отсутствующих на ВИ)')
run.font.size = Pt(11)
run.font.italic = True
run.font.color.rgb = RGBColor(0x60, 0x60, 0x60)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run('Дата проверки: 14 сентября 2026 г.')
run.font.size = Pt(11)
run.font.italic = True
run.font.color.rgb = RGBColor(0x60, 0x60, 0x60)

# Правила мониторинга
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run('⚠️ ВАЖНО: все цены указаны БЕЗ учёта скидок и распродаж. Если цена перечёркнута — берём её (постоянная цена). Скидки — временная опция, для мониторинга нужна постоянная цена.')
run.font.size = Pt(10)
run.font.bold = True
run.font.color.rgb = RGBColor(0xC0, 0x39, 0x2B)

doc.add_paragraph()

# Функция для добавления гиперссылки в ячейку
def add_hyperlink(paragraph, url, text):
    if not url:
        run = paragraph.add_run(text)
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor(0x80, 0x80, 0x80)
        return
    part = paragraph.part
    r_id = part.relate_to(url,
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
        is_external=True)
    hyperlink = OxmlElement('w:hyperlink')
    hyperlink.set(qn('r:id'), r_id)
    new_run = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')
    # Цвет ссылки
    color = OxmlElement('w:color')
    color.set(qn('w:val'), '1E3A5F')
    rPr.append(color)
    # Подчёркивание
    u = OxmlElement('w:u')
    u.set(qn('w:val'), 'single')
    rPr.append(u)
    # Размер шрифта
    sz = OxmlElement('w:sz')
    sz.set(qn('w:val'), '18')  # 9pt = 18 half-points
    rPr.append(sz)
    # Шрифт
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:ascii'), 'Arial')
    rFonts.set(qn('w:hAnsi'), 'Arial')
    rPr.append(rFonts)
    new_run.append(rPr)
    t = OxmlElement('w:t')
    t.text = text
    new_run.append(t)
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)

# Создание таблицы для категории
def add_category_table(cat_num, products):
    # Заголовок категории
    p = doc.add_paragraph()
    run = p.add_run(cat_titles.get(cat_num, f'Категория {cat_num}'))
    run.font.size = Pt(13)
    run.font.bold = True
    run.font.color.rgb = RGBColor(0x0B, 0x1C, 0x2C)

    # Сортировка по цене (возрастание)
    products_sorted = sorted(products, key=lambda x: x.get('final_price') or 0)

    # Таблица
    table = doc.add_table(rows=1, cols=6)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = 'Light Grid Accent 1'

    # Заголовки
    headers = ['№', 'Бренд / Продукт', 'Фасовка', 'Цена без скидки', 'Ссылка на ВИ', 'Источник']
    hdr_cells = table.rows[0].cells
    for i, h in enumerate(headers):
        hdr_cells[i].text = ''
        p = hdr_cells[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(h)
        run.font.bold = True
        run.font.size = Pt(10)
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        # Заливка
        tcPr = hdr_cells[i]._tc.get_or_add_tcPr()
        shd = OxmlElement('w:shd')
        shd.set(qn('w:fill'), '313131')
        tcPr.append(shd)

    # Данные
    for i, p_data in enumerate(products_sorted, 1):
        row = table.add_row()
        cells = row.cells

        # №
        cells[0].text = ''
        p = cells[0].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(str(i))
        run.font.size = Pt(10)

        # Бренд / Продукт
        cells[1].text = ''
        p = cells[1].paragraphs[0]
        run = p.add_run(p_data['brand'])
        run.font.bold = True
        run.font.size = Pt(10)
        if p_data['brand'] == 'ОРЕОЛ TITAN':
            run.font.color.rgb = RGBColor(0xF4, 0x76, 0x29)
        run = p.add_run('\n' + p_data['product'])
        run.font.size = Pt(9)
        run.font.italic = True
        run.font.color.rgb = RGBColor(0x60, 0x60, 0x60)

        # Фасовка
        cells[2].text = ''
        p = cells[2].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(p_data.get('facovka', '—'))
        run.font.size = Pt(10)

        # Цена без скидки
        cells[3].text = ''
        p = cells[3].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        price = p_data.get('final_price')
        if price:
            run = p.add_run(f'{price:,} ₽'.replace(',', ' '))
            run.font.bold = True
            run.font.size = Pt(11)
            # Цвет: TITAN — оранжевый, дешёвые — зелёные, дорогие — красные
            if p_data['brand'] == 'ОРЕОЛ TITAN':
                run.font.color.rgb = RGBColor(0xF4, 0x76, 0x29)
            elif price < 1000:
                run.font.color.rgb = RGBColor(0x16, 0xA3, 0x4A)
            elif price > 1300:
                run.font.color.rgb = RGBColor(0xC0, 0x39, 0x2B)
        else:
            run = p.add_run('—')
            run.font.size = Pt(10)

        # Ссылка на ВИ
        cells[4].text = ''
        p = cells[4].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        url = p_data.get('url', '')
        if url and 'vseinstrumenti.ru' in url:
            add_hyperlink(p, url, 'Открыть ↗')
        elif url:
            add_hyperlink(p, url, 'Ссылка ↗')
        else:
            run = p.add_run('—')
            run.font.size = Pt(9)

        # Источник
        cells[5].text = ''
        p = cells[5].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        src = p_data.get('source', '—')
        if 'vseinstrumenti' in (url or '') and 'vseinstrumenti' in src.lower():
            src_text = 'vseinstrumenti.ru'
        elif 'snippet' in src:
            src_text = url.split('//')[1].split('/')[0] if '://' in (url or '') else 'snippet'
        else:
            src_text = src[:25] if src else '—'
        run = p.add_run(src_text)
        run.font.size = Pt(8)
        run.font.italic = True
        run.font.color.rgb = RGBColor(0x80, 0x80, 0x80)

    # Ширины колонок
    widths = [Cm(0.8), Cm(5.0), Cm(2.5), Cm(2.5), Cm(2.0), Cm(3.5)]
    for row in table.rows:
        for i, w in enumerate(widths):
            if i < len(row.cells):
                row.cells[i].width = w

    doc.add_paragraph()

# Добавляем таблицы для каждой категории
for cat_num in sorted(categories.keys()):
    add_category_table(cat_num, categories[cat_num]['products'])

# === Итоговая сводка ===
doc.add_page_break()
p = doc.add_paragraph()
run = p.add_run('Сводная аналитика')
run.font.size = Pt(16)
run.font.bold = True
run.font.color.rgb = RGBColor(0x0B, 0x1C, 0x2C)

# Сводная таблица по брендам
p = doc.add_paragraph()
run = p.add_run('Сравнение цен по брендам (минимальная цена в категории)')
run.font.size = Pt(12)
run.font.bold = True
run.font.color.rgb = RGBColor(0x0B, 0x1C, 0x2C)

# Группировка по брендам
brand_prices = {}
for r in data:
    b = r['brand']
    if b not in brand_prices:
        brand_prices[b] = []
    if r.get('final_price'):
        brand_prices[b].append({'cat': r['cat'], 'product': r['product'], 'price': r['final_price']})

# Сводная таблица
table = doc.add_table(rows=1, cols=4)
table.style = 'Light Grid Accent 1'
hdr = table.rows[0].cells
headers = ['Бренд', 'Кол-во позиций', 'Средняя цена', 'Диапазон цен']
for i, h in enumerate(headers):
    hdr[i].text = ''
    p = hdr[i].paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(h)
    run.font.bold = True
    run.font.size = Pt(10)
    run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    tcPr = hdr[i]._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:fill'), '313131')
    tcPr.append(shd)

# Сортируем бренды по средней цене (возрастание)
brand_avg = []
for b, items in brand_prices.items():
    if not items:
        continue
    avg = sum(i['price'] for i in items) / len(items)
    brand_avg.append((b, items, avg))
brand_avg.sort(key=lambda x: x[2])

for b, items, avg in brand_avg:
    row = table.add_row()
    cells = row.cells
    prices_list = [i['price'] for i in items]
    min_p = min(prices_list)
    max_p = max(prices_list)

    cells[0].text = ''
    p = cells[0].paragraphs[0]
    run = p.add_run(b)
    run.font.bold = True
    run.font.size = Pt(10)
    if b == 'ОРЕОЛ TITAN':
        run.font.color.rgb = RGBColor(0xF4, 0x76, 0x29)

    cells[1].text = ''
    p = cells[1].paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(str(len(items)))
    run.font.size = Pt(10)

    cells[2].text = ''
    p = cells[2].paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(f'{int(avg):,} ₽'.replace(',', ' '))
    run.font.size = Pt(10)
    run.font.bold = True

    cells[3].text = ''
    p = cells[3].paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if min_p == max_p:
        run = p.add_run(f'{min_p:,} ₽'.replace(',', ' '))
    else:
        run = p.add_run(f'{min_p:,} – {max_p:,} ₽'.replace(',', ' '))
    run.font.size = Pt(10)

# Ширины
widths = [Cm(5), Cm(3), Cm(3), Cm(5)]
for row in table.rows:
    for i, w in enumerate(widths):
        if i < len(row.cells):
            row.cells[i].width = w

doc.add_paragraph()

# Ключевые выводы
p = doc.add_paragraph()
run = p.add_run('Ключевые выводы мониторинга от 14.09.2026:')
run.font.size = Pt(12)
run.font.bold = True
run.font.color.rgb = RGBColor(0x0B, 0x1C, 0x2C)

# Считаем выводы
titan_cat1 = next((r for r in data if r['cat'] == 1 and r['brand'] == 'ОРЕОЛ TITAN'), None)
titan_cat2 = next((r for r in data if r['cat'] == 2 and r['brand'] == 'ОРЕОЛ TITAN'), None)
titan_cat3 = next((r for r in data if r['cat'] == 3 and r['brand'] == 'ОРЕОЛ TITAN'), None)

cat1_prices = [r['final_price'] for r in data if r['cat'] == 1 and r.get('final_price')]
cat2_prices = [r['final_price'] for r in data if r['cat'] == 2 and r.get('final_price')]
cat3_prices = [r['final_price'] for r in data if r['cat'] == 3 and r.get('final_price')]

findings = []

if titan_cat1 and titan_cat1.get('final_price'):
    price = titan_cat1['final_price']
    avg = sum(cat1_prices) / len(cat1_prices)
    findings.append(f'В категории 1 (3в1 RAL 9005) TITAN стоит {price:,} ₽ — это {"дешевле" if price < avg else "дороже"} средней ({int(avg):,} ₽) на {int(abs(price-avg)):,} ₽.'.replace(',', ' '))

if titan_cat2 and titan_cat2.get('final_price'):
    price = titan_cat2['final_price']
    avg = sum(cat2_prices) / len(cat2_prices)
    findings.append(f'В категории 2 (быстросохнущая RAL 7040) TITAN стоит {price:,} ₽ — это {"дешевле" if price < avg else "дороже"} средней ({int(avg):,} ₽) на {int(abs(price-avg)):,} ₽.'.replace(',', ' '))

if titan_cat3 and titan_cat3.get('final_price'):
    price = titan_cat3['final_price']
    avg = sum(cat3_prices) / len(cat3_prices)
    findings.append(f'В категории 3 (молотковая) TITAN стоит {price:,} ₽ — это {"дешевле" if price < avg else "дороже"} средней ({int(avg):,} ₽) на {int(abs(price-avg)):,} ₽.'.replace(',', ' '))

# Самый дешёвый/дорогой в каждой категории
for cat in [1, 2, 3]:
    cat_items = [r for r in data if r['cat'] == cat and r.get('final_price')]
    if cat_items:
        cheapest = min(cat_items, key=lambda x: x['final_price'])
        expensive = max(cat_items, key=lambda x: x['final_price'])
        findings.append(f'В категории {cat}: самый дешёвый — {cheapest["brand"]} ({cheapest["final_price"]:,} ₽), самый дорогой — {expensive["brand"]} ({expensive["final_price"]:,} ₽).'.replace(',', ' '))

# Новые позиции
findings.append('Добавлены новые позиции: Лакра 5в1 грунт-эмаль (1 103 ₽, категория 4) — прямой конкурент TITAN в категории грунт-эмалей по ржавчине.')
findings.append('Добавлен TOPGRADE by Hammerite (1 700 ₽, категория 5) — премиальный сегмент, ближайший конкурент Tikkurila METALLISTA.')

for finding in findings:
    p = doc.add_paragraph(style='List Bullet')
    run = p.add_run(finding)
    run.font.size = Pt(10)

# Методология
doc.add_paragraph()
p = doc.add_paragraph()
run = p.add_run('Методология мониторинга')
run.font.size = Pt(12)
run.font.bold = True
run.font.color.rgb = RGBColor(0x0B, 0x1C, 0x2C)

methodology = [
    'Источник цен — сайт vseinstrumenti.ru (каталог товаров).',
    'Для товаров, недоступных на ВИ, — альтернативные источники (СтройЛандия, market.stworka.com, tetrisvn.ru).',
    'Все цены указаны БЕЗ учёта скидок и распродаж (по правилу: если цена перечёркнута — берём перечёркнутую как постоянную).',
    'Актуальные цены на 14 сентября 2026 г.',
    'Каждая позиция содержит прямую ссылку на товар для быстрой проверки на сайте-источнике.',
    'Цены могут незначительно меняться — рекомендуется перепроверка перед ключевыми решениями по ценообразованию.',
]
for m in methodology:
    p = doc.add_paragraph(style='List Bullet')
    run = p.add_run(m)
    run.font.size = Pt(10)

# Сохраняем
out_path = '/home/z/my-project/download/TITAN_Актуальные_цены_конкурентов_14.09.2026.docx'
doc.save(out_path)
print(f'✅ Сохранено: {out_path}')
print(f'\nОбработано: {len(data)} продуктов в {len(categories)} категориях')
