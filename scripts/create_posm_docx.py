#!/usr/bin/env python3
"""
Создание DOCX: Бенчмарк POSm-материалов конкурентов для программы Эмпилс
С акцентом на TCO-калькулятор как ключевой инструмент
"""

from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml
import os

IMG_DIR = "/home/z/my-project/download/posm_images"
OUT = "/home/z/my-project/download/Бенчмарк_POSm_конкурентов_Эмпилс.docx"

# Colors
NAVY = RGBColor(0x0A, 0x25, 0x40)
CYAN = RGBColor(0x00, 0xB4, 0xD8)
DARK_BLUE = RGBColor(0x1B, 0x2A, 0x4A)
BODY_COLOR = RGBColor(0x33, 0x33, 0x33)
ACCENT_ORANGE = RGBColor(0xE8, 0x6C, 0x00)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)

doc = Document()

# ===== PAGE SETUP =====
for section in doc.sections:
    section.page_width = Cm(21)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(2.5)
    section.bottom_margin = Cm(2.0)
    section.left_margin = Cm(2.5)
    section.right_margin = Cm(2.0)

# ===== STYLES =====
style = doc.styles['Normal']
style.font.name = 'Carlito'
style.font.size = Pt(11)
style.font.color.rgb = BODY_COLOR
style.paragraph_format.line_spacing = 1.15
style.paragraph_format.space_after = Pt(6)
# Set CJK font
rpr = style.element.get_or_add_rPr()
rFonts = parse_xml(f'<w:rFonts {nsdecls("w")} w:ascii="Carlito" w:hAnsi="Carlito" w:eastAsia="Noto Sans SC" w:cs="Carlito"/>')
rpr.append(rFonts)

def add_heading_styled(text, level=1):
    h = doc.add_heading(text, level=level)
    for run in h.runs:
        run.font.color.rgb = NAVY
        run.font.name = 'Carlito'
    return h

def add_body(text, bold=False, italic=False, color=None):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.bold = bold
    run.italic = italic
    run.font.size = Pt(11)
    run.font.name = 'Carlito'
    if color:
        run.font.color.rgb = color
    p.paragraph_format.space_after = Pt(6)
    return p

def add_rich_para(parts):
    """parts = [(text, bold, italic, color, size), ...]"""
    p = doc.add_paragraph()
    for text, bold, italic, color, size in parts:
        run = p.add_run(text)
        run.bold = bold
        run.italic = italic
        run.font.name = 'Carlito'
        run.font.size = Pt(size if size else 11)
        if color:
            run.font.color.rgb = color
    p.paragraph_format.space_after = Pt(6)
    return p

def add_image(filename, width_cm=14, caption=None):
    path = os.path.join(IMG_DIR, filename)
    if os.path.exists(path) and os.path.getsize(path) > 5000:
        try:
            from PIL import Image as PILImage
            im = PILImage.open(path)
            im.verify()
            # Re-open after verify (verify can close the file)
            im = PILImage.open(path)
            # Convert to RGB if necessary (for PNG with transparency etc.)
            if im.mode not in ('RGB', 'L'):
                im = im.convert('RGB')
            # Save as clean JPEG to temp
            import tempfile
            tmp = tempfile.NamedTemporaryFile(suffix='.jpg', delete=False)
            im.save(tmp.name, 'JPEG', quality=90)
            tmp.close()
            img_path = tmp.name
        except Exception as e:
            img_path = path
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run()
        try:
            run.add_picture(img_path, width=Cm(width_cm))
        except Exception:
            p.text = f"[Изображение: {filename}]"
        p.paragraph_format.space_after = Pt(4)
        if caption:
            cap = doc.add_paragraph()
            cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = cap.add_run(caption)
            r.italic = True
            r.font.size = Pt(9)
            r.font.color.rgb = RGBColor(0x66, 0x66, 0x66)
            cap.paragraph_format.space_after = Pt(12)
    else:
        p = doc.add_paragraph(f"[Изображение: {filename}]")
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(12)

def set_cell_shading(cell, color_hex):
    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}" w:val="clear"/>')
    cell._tc.get_or_add_tcPr().append(shading)

def add_table_row(table, cells_data, header=False):
    row = table.add_row()
    for i, (text, bold, align) in enumerate(cells_data):
        cell = row.cells[i]
        cell.text = ""
        p = cell.paragraphs[0]
        run = p.add_run(text)
        run.bold = bold
        run.font.size = Pt(10) if not header else Pt(10)
        run.font.name = 'Carlito'
        if header:
            run.font.color.rgb = WHITE
            set_cell_shading(cell, "0A2540")
        p.alignment = align
    return row

# ====================================================================
# TITLE PAGE
# ====================================================================
for _ in range(6):
    doc.add_paragraph()

title_p = doc.add_paragraph()
title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = title_p.add_run("БЕНЧМАРК POSm-МАТЕРИАЛОВ КОНКУРЕНТОВ")
r.bold = True
r.font.size = Pt(26)
r.font.color.rgb = NAVY
r.font.name = 'Carlito'

sub_p = doc.add_paragraph()
sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = sub_p.add_run("Визуальные примеры для программы развития Эмпилс 2026-2028")
r.font.size = Pt(14)
r.font.color.rgb = CYAN
r.font.name = 'Carlito'

doc.add_paragraph()

accent_p = doc.add_paragraph()
accent_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = accent_p.add_run("Ключевой фокус: TCO-калькулятор как центральный инструмент продаж")
r.bold = True
r.font.size = Pt(13)
r.font.color.rgb = ACCENT_ORANGE
r.font.name = 'Carlito'

doc.add_paragraph()
doc.add_paragraph()

meta_p = doc.add_paragraph()
meta_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = meta_p.add_run("ЗАО «Эмпилс» · Июнь 2026 · Для внутреннего использования")
r.font.size = Pt(11)
r.font.color.rgb = RGBColor(0x88, 0x88, 0x88)
r.font.name = 'Carlito'

doc.add_page_break()

# ====================================================================
# TABLE OF CONTENTS (manual)
# ====================================================================
add_heading_styled("Содержание", level=1)
toc_items = [
    "1. TCO-калькулятор: главный инструмент трансформации продаж",
    "2. Dulux Weathershield: 3-шаговая система с гарантией",
    "3. Sikkens (AkzoNobel): система древесных покрытий",
    "4. Rawlplug: модульная навигация по типу поверхности",
    "5. Sherwin-Williams ColorSnap Studio: интерактивный подбор",
    "6. Farrow & Ball: премиум-бутик и консультация",
    "7. Benjamin Moore: программа трансформации магазина",
    "8. Caparol: системная навигация и экологичность",
    "9. Сводная таблица: что позаимствовать для Эмпилс",
    "10. Рекомендации по внедрению TCO-калькулятора",
]
for item in toc_items:
    p = doc.add_paragraph(item)
    p.paragraph_format.space_after = Pt(4)
    for run in p.runs:
        run.font.size = Pt(11)

doc.add_page_break()

# ====================================================================
# 1. TCO-КАЛЬКУЛЯТОР
# ====================================================================
add_heading_styled("1. TCO-калькулятор: главный инструмент трансформации продаж", level=1)

add_body(
    "TCO-калькулятор (Total Cost of Ownership — совокупная стоимость владения) — это инструмент, "
    "который переводит разговор с клиентом от «цены за банку» к «стоимости года эксплуатации». "
    "Это самое мощное оружие против возражения «дорого», потому что оно не опирается на эмоции, "
    "а показывает конкретные цифры: премиум-решение Эмпилс дешевле в 2,6 раза в пересчёте на год, "
    "чем эконом-вариант. Ни один из обследованных конкурентов в мире не использует TCO-калькулятор "
    "как структурированный POSm-инструмент — это уникальное конкурентное преимущество для Эмпилс."
)

add_heading_styled("Почему TCO-калькулятор — ключевой элемент", level=2)

add_body(
    "Анализ мировых практик показывает: конкуренты подходят к обоснованию цены премиума через гарантии "
    "(Dulux — 10 лет гарантии при использовании системы), через визуализацию результата (Farrow & Ball — "
    "консультация по цвету), через интерактивный подбор (Sherwin-Williams — ColorSnap). Но ни один из них "
    "не даёт клиенту калькулятор, который считает стоимость года эксплуатации. Это белое пятно на рынке — "
    "и именно его должен занять Эмпилс."
)

add_rich_para([
    ("Формула TCO:", True, False, NAVY, 12),
    ("  TCO = Материалы + Подготовка + Нанесение + Ремонты + Простой", False, False, BODY_COLOR, 11),
])

add_heading_styled("Пример расчёта: комната 20 м²", level=2)

# TCO comparison table
table = doc.add_table(rows=1, cols=3)
table.alignment = WD_TABLE_ALIGNMENT.CENTER
table.style = 'Table Grid'

# Header
for i, text in enumerate(["Показатель", "Эконом + Средний (Эмпилс)", "Премиум (Эмпилс)"]):
    cell = table.rows[0].cells[i]
    cell.text = ""
    p = cell.paragraphs[0]
    r = p.add_run(text)
    r.bold = True
    r.font.size = Pt(10)
    r.font.name = 'Carlito'
    r.font.color.rgb = WHITE
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_cell_shading(cell, "0A2540")

rows_data = [
    ("Материалы", "3 600 ₽", "8 720 ₽"),
    ("Работа (грунт + шпатлёвка + покраска)", "15 750 ₽", "15 750 ₽"),
    ("Общая стоимость ремонта", "19 350 ₽", "24 470 ₽"),
    ("Срок службы покрытия", "3 года", "10 лет"),
    ("Стоимость 1 года эксплуатации", "6 450 ₽", "2 447 ₽"),
    ("Экономия в год", "—", "4 003 ₽"),
    ("Экономия за 10 лет", "—", "33 580 ₽"),
]

for label, econ, prem in rows_data:
    row = table.add_row()
    for i, text in enumerate([label, econ, prem]):
        cell = row.cells[i]
        cell.text = ""
        p = cell.paragraphs[0]
        r = p.add_run(text)
        r.font.size = Pt(10)
        r.font.name = 'Carlito'
        if i == 0:
            r.bold = True
        if label == "Стоимость 1 года эксплуатации":
            r.bold = True
            if i == 2:
                r.font.color.rgb = RGBColor(0x1B, 0x7D, 0x46)
            if i == 1:
                r.font.color.rgb = RGBColor(0xC0, 0x39, 0x2B)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER if i > 0 else WD_ALIGN_PARAGRAPH.LEFT

doc.add_paragraph()

add_rich_para([
    ("Результат:", True, False, ACCENT_ORANGE, 12),
    (" Премиум дешевле в 2,6 раза в год эксплуатации. Переплата за премиум — всего +7% к смете, но -60% к стоимости года.", False, False, BODY_COLOR, 11),
])

add_heading_styled("Как конкуренты решают проблему «дорого» vs TCO-подход Эмпилс", level=2)

add_body(
    "Все конкуренты пытаются обосновать цену премиума, но делают это косвенно — через гарантии, "
    "экологичность, визуализацию. Эмпилс может стать первым, кто даст клиенту прозрачный расчёт "
    "стоимости владения — и в магазине (Excel-калькулятор у продавца), и на сайте (веб-версия), "
    "и прямо на полке (shelf talker с TCO-сравнением)."
)

# Comparison approach table
table2 = doc.add_table(rows=1, cols=3)
table2.alignment = WD_TABLE_ALIGNMENT.CENTER
table2.style = 'Table Grid'

for i, text in enumerate(["Конкурент", "Метод обоснования цены", "Ограничение"]):
    cell = table2.rows[0].cells[i]
    cell.text = ""
    p = cell.paragraphs[0]
    r = p.add_run(text)
    r.bold = True
    r.font.size = Pt(10)
    r.font.name = 'Carlito'
    r.font.color.rgb = WHITE
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_cell_shading(cell, "0A2540")

approach_rows = [
    ("Dulux", "Гарантия 10 лет при использовании системы", "Нет расчёта стоимости года; гарантия — не аргумент для эконом-сегмента"),
    ("Sikkens", "Системный подход: 3 продукта = максимальная защита", "Нет калькулятора; клиент не видит экономику системы"),
    ("Sherwin-Williams", "ColorSnap — интерактивный подбор цвета", "Подбор цвета, а не расчёт стоимости; не работает для обоснования цены"),
    ("Farrow & Ball", "Платная консультация по цвету (£100-500)", "Только для премиума; недоступно массового клиента"),
    ("Caparol", "Экологичность + системная маркировка", "Аргумент «экологично», но не «дешевле в год»"),
    ("Эмпилс (план)", "TCO-калькулятор: стоимость года эксплуатации", "Прямой расчёт; клиент видит: премиум = 2 447 ₽/год vs эконом = 6 450 ₽/год"),
]

for brand, method, limit in approach_rows:
    row = table2.add_row()
    for i, text in enumerate([brand, method, limit]):
        cell = row.cells[i]
        cell.text = ""
        p = cell.paragraphs[0]
        r = p.add_run(text)
        r.font.size = Pt(10)
        r.font.name = 'Carlito'
        if i == 0:
            r.bold = True
        if brand == "Эмпилс (план)":
            set_cell_shading(cell, "E8F5E9")
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT

doc.add_paragraph()

doc.add_page_break()

# ====================================================================
# 2. DULUX WEATHERSHIELD
# ====================================================================
add_heading_styled("2. Dulux Weathershield: 3-шаговая система с гарантией", level=1)

add_body(
    "Dulux Weathershield — самый прямой аналог концепции «сгруппированных технологий» Эмпилс. "
    "Это 3-шаговая система для фасадных работ, где каждый шаг — обязательный элемент, а гарантия "
    "10 лет действует только при использовании всех трёх продуктов вместе. Этот подход уже доказал "
    "свою эффективность в B&Q, Homebase и других DIY-сетях Великобритании и Австралии."
)

add_heading_styled("3-шаговая система", level=2)

steps_table = doc.add_table(rows=1, cols=3)
for i, text in enumerate(["Шаг", "Продукт", "Роль в системе"]):
    cell = steps_table.rows[0].cells[i]
    cell.text = ""
    p = cell.paragraphs[0]
    r = p.add_run(text)
    r.bold = True; r.font.size = Pt(10); r.font.name = 'Carlito'; r.font.color.rgb = WHITE
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_cell_shading(cell, "0A2540")

for step, prod, role in [
    ("Шаг 1", "Weathershield Exterior Primer", "Грунт — подготавливает поверхность, обеспечивает адгезию"),
    ("Шаг 2", "Weathershield Exterior Undercoat", "Подкладка — выравнивает, создаёт барьер для влаги"),
    ("Шаг 3", "Weathershield Exterior Gloss/Satin", "Финишное покрытие — защита от UV и погоды на 10 лет"),
]:
    row = steps_table.add_row()
    for i, text in enumerate([step, prod, role]):
        cell = row.cells[i]
        cell.text = ""
        p = cell.paragraphs[0]
        r = p.add_run(text)
        r.font.size = Pt(10); r.font.name = 'Carlito'
        if i == 0: r.bold = True
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT

doc.add_paragraph()

add_heading_styled("POSm-элементы Dulux", level=2)

add_body(
    "На банках крупно напечатано «10 Year Guarantee» — это главный selling point. Гарантия действует "
    "ТОЛЬКО при использовании всех 3-х продуктов системы, что автоматически стимулирует покупку комплекта. "
    "В магазинах B&Q и Homebase используются shelf talkers с визуальными шагами 1→2→3, а цветовые "
    "карточки привязаны к системе. На банке Step 1 написано: «Use with Weathershield Undercoat + Gloss "
    "for 10 year guarantee» — это кросс-продажа встроена в упаковку."
)

add_rich_para([
    ("Что позаимствовать для Эмпилс:", True, False, CYAN, 11),
    (" Аналогичный подход «3 шага к гарантии результата» — Грунт → Шпатлёвка → Краска. На каждой банке AVANTGARDE указать: «Шаг 3 из 4. Используйте с ОРЕОЛ грунт + NEWTONE шпатлёвка для гарантии результата». Гарантия 10 лет при соблюдении технологии — мощнейший аргумент, которого нет у конкурентов в РФ.", False, False, BODY_COLOR, 11),
])

# ====================================================================
# 3. SIKKENS
# ====================================================================
add_heading_styled("3. Sikkens (AkzoNobel): система древесных покрытий", level=1)

add_body(
    "Sikkens — премиум-бренд древесных покрытий, входящий в группу AkzoNobel (как и Dulux). "
    "Sikkens продаёт не отдельные продукты, а полную систему защиты дерева из 3-х шагов: "
    "пропитка/грунт (Cetol HLS Plus) → промежуточный слой (Cetol Filter 7) → финишное покрытие "
    "(Cetol BL Topcoat). Этот подход аналогичен концепции «длинной цепочки» Эмпилс для деревянного домика."
)

add_image("sikkens_system.jpg", 10, "Sikkens Cetol HLS Plus — первый шаг системы древесных покрытий")

add_heading_styled("POSm-решения Sikkens", level=2)

add_body(
    "Цветовые веера Sikkens привязаны к системе, а не к отдельному продукту. На каждом веере указан "
    "рекомендуемый порядок нанесения всех трёх продуктов. В магазинах ЕС продукты продаются наборами "
    "(bundle) на одной полке, а shelf wobbler гласит: «Complete Sikkens System = Maximum Protection». "
    "Ключевой приём — визуальная связка: баночка с образцом цвета стоит рядом с рекомендованным "
    "грунтом и финишем, формируя «решение» в голове покупателя."
)

add_rich_para([
    ("Что позаимствовать для Эмпилс:", True, False, CYAN, 11),
    (" Bundle-выкладка: полка «Ремонт ванной» = грунт + шпатлёвка + краска + лак рядом, а не на разных стеллажах. "
     "Веер цветов с указанием полного состава цепочки. Shelf wobbler: «Полная система Эмпилс = Гарантия результата».", False, False, BODY_COLOR, 11),
])

# ====================================================================
# 4. RAWLPLUG
# ====================================================================
add_heading_styled("4. Rawlplug: модульная навигация по типу поверхности", level=1)

add_body(
    "Rawlplug — польский производитель крепёжных систем, создавший одну из лучших в мире модульных "
    "POS-систем для DIY-ритейла. Несмотря на то, что категория крепёж — не ЛКМ, подход Rawlplug "
    "к навигации по типу поверхности является идеальным аналогом для концепции «решения по задачам» Эмпилс. "
    "Вместо навигации по продукту (дюбель, саморез, анкер) Rawlplug организует полки по типу основания: "
    "дерево, бетон, гипсокартон, универсальные. Покупатель ищет «дюбель для гипсокартона», а не конкретную модель."
)

add_image("rawlplug_display.jpg", 12, "Модульный стенд Rawlplug с навигацией по типу поверхности")

add_image("rawlplug_modular.jpg", 12, "Полки Rawlplug: Universal / Plasterboard / Solid Wall")

add_heading_styled("Ключевые элементы POS-системы Rawlplug", level=2)

add_body(
    "Модульные стенды с заголовками по типу поверхности; POS Essentials Guide — каталог для ритейлеров "
    "с инструкциями по выкладке; гибкая модульная система, адаптируемая под размер магазина; "
    "на каждой полке — иконка типа поверхности (кирпич, дерево, гипсокартон) + список продуктов для этого основания. "
    "Простота и наглядность — покупатель находит нужное за 10 секунд."
)

add_rich_para([
    ("Что позаимствовать для Эмпилс:", True, False, CYAN, 11),
    (" Стенд «Ванная / Фасад / Дерево / Детская» вместо навигации по брендам. "
     "Иконка задачи + полный состав цепочки на каждой полке. "
     "POS Essentials Guide для партнёров Эмпилс — инструкция по выкладке решений.", False, False, BODY_COLOR, 11),
])

doc.add_page_break()

# ====================================================================
# 5. SHERWIN-WILLIAMS
# ====================================================================
add_heading_styled("5. Sherwin-Williams ColorSnap Studio: интерактивный подбор", level=1)

add_body(
    "Sherwin-Williams — крупнейшая сеть лакокрасочных магазинов в США (4 400+ точек). В каждом магазине "
    "установлена ColorSnap Studio — интерактивная станция подбора цвета с сенсорным экраном. "
    "Клиент выбирает цвет, визуализирует его в интерьере и получает полный список продуктов системы: "
    "primer + paint + topcoat. Подход «from color to complete solution» — от цвета к полному решению."
)

add_heading_styled("POSm-элементы Sherwin-Williams", level=2)

add_body(
    "ColorSnap Studio — стационарная интерактивная станция; ColorSnap Visualizer App — мобильное приложение "
    "с тем же функционалом; shelf talker с QR-кодом на приложение; стенд «Solution Center» — зона консультаций "
    "в магазине; персонал обучен продавать систему, а не отдельный продукт. Ключевая идея: "
    "выбрал цвет → получил список всех продуктов для его реализации."
)

add_rich_para([
    ("Что позаимствовать для Эмпилс:", True, False, CYAN, 11),
    (" TCO-калькулятор на планшете в магазине — аналог ColorSnap Studio, но с фокусом на стоимость, а не цвет. "
     "QR-код на shelf talker ведёт на веб-калькулятор. Зона «Решения Эмпилс» — аналог Solution Center.", False, False, BODY_COLOR, 11),
])

# ====================================================================
# 6. FARROW & BALL
# ====================================================================
add_heading_styled("6. Farrow & Ball: премиум-бутик и консультация", level=1)

add_body(
    "Farrow & Ball — эталон премиум-подхода в ЛКМ. Бренд не продаёт «банку краски» — продаёт консультацию "
    "и систему. Только 132 оттенка (vs 2000+ у масс-маркета), но каждый с полной системой: "
    "Primer + Undercoat + Topcoat + Dead Flat. Магазины оформлены как бутики с брендированными "
    "консультантами в фартуках «F&B Studio»."
)

add_image("farrow_ball_studio.jpg", 10, "Farrow & Ball Studio — брендированные консультанты")

add_image("farrow_ball_storefront.jpg", 10, "Витрина Farrow & Ball с Colour Consultancy Service")

add_image("farrow_ball_display.jpg", 10, "Стенд с банками и образцами Farrow & Ball")

add_heading_styled("POSm-элементы Farrow & Ball", level=2)

add_body(
    "F&B Studio — брендированные консультанты в фартуках; Colour Consultancy — платная услуга подбора "
    "цвета (£100-500); полка организована по finish/эффекту (Modern Emulsion, Estate Eggshell, Full Gloss), "
    "а не по цене; карточка цвета = «Complete the look» с рекомендованными праймерами и покрытиями; "
    "в магазине можно купить только полный комплект — отдельно праймер не продаётся без указания "
    "соответствующего топ-коута."
)

add_rich_para([
    ("Что позаимствовать для Эмпилс:", True, False, CYAN, 11),
    (" Сертифицированный «Эксперт Эмпилс» с бейджем — аналог F&B Studio. "
     "Карточка каждого продукта = «С чем использовать» (рекомендованная цепочка). "
     "Организация полки по задаче/эффекту, а не по цене.", False, False, BODY_COLOR, 11),
])

doc.add_page_break()

# ====================================================================
# 7. BENJAMIN MOORE
# ====================================================================
add_heading_styled("7. Benjamin Moore: программа трансформации магазина", level=1)

add_body(
    "Benjamin Moore предлагает ритейлерам полную трансформацию магазина под бренд — «Convert Your Store». "
    "7 500+ независимых магазинов по всей США работают по единому стандарту: брендированные стеллажи, "
    "цветовые стены, Gennex Color System (проприетарная система колеровки), Solution Station "
    "(зона консультаций), обучение персонала и сертификация. Это самая масштабная программа "
    "партнёрства в ЛКМ-индустрии."
)

add_image("bm_store.jpg", 12, "Интерьер магазина Benjamin Moore с организованной выкладкой")

add_rich_para([
    ("Что позаимствовать для Эмпилс:", True, False, CYAN, 11),
    (" «Стань партнёром Эмпилс» — брендированная зона решений в магазине партнёра. "
     "Посадочная страница для привлечения новых партнёров. "
     "Обучение и сертификация персонала партнёра.", False, False, BODY_COLOR, 11),
])

# ====================================================================
# 8. CAPAROL
# ====================================================================
add_heading_styled("8. Caparol: системная навигация и экологичность", level=1)

add_body(
    "Caparol (Германия, DAW-группа) — лидер немецкого рынка ЛКМ с подходом системной навигации. "
    "Caparol 3D-System Plus предлагает 1 383 цвета, но навигация организована по области применения: "
    "внутренние, фасадные, лаки. Премиум-линейка Caparol Icons делает акцент на экологичность — "
    "«Kinderspielzeuggeeignet» (пригодна для детских игрушек), что является мощным аргументом "
    "для осознанного потребителя."
)

add_image("caparol_3d_fan.jpg", 10, "Цветовой веер Caparol 3D-System Plus с системной навигацией")

add_heading_styled("POSm-элементы Caparol", level=2)

add_body(
    "На банках крупная маркировка «Grundiersystem» (грунтовочная система) — указание на обязательное "
    "использование с грунтом. Shelf talker: «Systemgerecht» = «в составе системы». Цветовой веер "
    "с указанием области применения (интерьер/фасад/лак). На продуктах Icons — значок экологичности "
    "и безопасности для детей."
)

add_rich_para([
    ("Что позаимствовать для Эмпилс:", True, False, CYAN, 11),
    (" Системная маркировка на банке AVANTGARDE: «Шаг 3 из 4. Используйте с ОРЕОЛ грунт + NEWTONE шпатлёвка». "
     "Shelf talker «Systemgerecht» → «В составе системы Эмпилс». "
     "Экологичность как дополнительный аргумент для премиума.", False, False, BODY_COLOR, 11),
])

doc.add_page_break()

# ====================================================================
# 9. СВОДНАЯ ТАБЛИЦА
# ====================================================================
add_heading_styled("9. Сводная таблица: что позаимствовать для Эмпилс", level=1)

summary_table = doc.add_table(rows=1, cols=3)
summary_table.alignment = WD_TABLE_ALIGNMENT.CENTER
summary_table.style = 'Table Grid'

for i, text in enumerate(["Приём", "Кто использует", "Как адаптировать для Эмпилс"]):
    cell = summary_table.rows[0].cells[i]
    cell.text = ""
    p = cell.paragraphs[0]
    r = p.add_run(text)
    r.bold = True; r.font.size = Pt(9); r.font.name = 'Carlito'; r.font.color.rgb = WHITE
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_cell_shading(cell, "0A2540")

summary_rows = [
    ("Система из N шагов с гарантией", "Dulux Weathershield",
     "«3 шага к гарантии результата» — Грунт → Шпатлёвка → Краска. Гарантия 10 лет при соблюдении технологии"),
    ("Навигация по типу поверхности", "Rawlplug",
     "Стенд «Ванная / Фасад / Дерево / Детская» вместо навигации по брендам"),
    ("Shelf talker «стоимость года»", "Аналогов нет — уникальное для Эмпилс",
     "Shelf wobbler: «Эконом 6 450 ₽/год → Премиум 2 447 ₽/год» прямо на полке"),
    ("Интерактивная станция подбора", "Sherwin-Williams",
     "TCO-калькулятор + 3D-тур на планшете в магазине"),
    ("Премиум-консультант", "Farrow & Ball",
     "Сертифицированный «Эксперт Эмпилс» с бейджем в точке продаж"),
    ("Bundle-выкладка (комплектом)", "Sikkens",
     "Полка «Ремонт ванной» = грунт + шпатлёвка + краска + лак рядом"),
    ("Системная маркировка на банке", "Caparol",
     "На банке AVANTGARDE: «Шаг 3 из 4. Используйте с ОРЕОЛ грунт + NEWTONE шпатлёвка»"),
    ("Программа трансформации магазина", "Benjamin Moore",
     "«Стань партнёром Эмпилс» — брендированная зона решений в магазине партнёра"),
    ("TCO-калькулятор (Excel + веб)", "Аналогов нет — уникальное для Эмпилс",
     "Продавец за 2 мин показывает: премиум дешевле в 2,6× в год. На сайте — онлайн-версия с лидогенерацией"),
]

for priem, who, how in summary_rows:
    row = summary_table.add_row()
    for i, text in enumerate([priem, who, how]):
        cell = row.cells[i]
        cell.text = ""
        p = cell.paragraphs[0]
        r = p.add_run(text)
        r.font.size = Pt(9); r.font.name = 'Carlito'
        if i == 0: r.bold = True
        if "уникальное" in who.lower():
            r.font.color.rgb = ACCENT_ORANGE
            r.bold = True
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT

doc.add_paragraph()

# ====================================================================
# 10. РЕКОМЕНДАЦИИ ПО TCO
# ====================================================================
add_heading_styled("10. Рекомендации по внедрению TCO-калькулятора", level=1)

add_body(
    "TCO-калькулятор — это не просто Excel-файл. Это экосистема из трёх уровней, каждый из которых "
    "работает на разных этапах взаимодействия с клиентом. Ниже приведены рекомендации по поэтапному "
    "внедрению, основанные на анализе мировых практик."
)

add_heading_styled("Уровень 1: Excel-калькулятор для менеджера (август 2026)", level=2)

add_body(
    "Создать Excel-калькулятор с предзаполненными сценариями (комната 20 м², фасад дома, ванная). "
    "Менеджер вводит площадь и тип помещения → получает автоматический расчёт стоимости материалов, "
    "работ, срока службы и стоимости 1 года эксплуатации. Сравнение «эконом vs премиум» с визуализацией "
    "экономии. Защита формул, инструкция для менеджера. Аналог в мире: ни у кого нет — это будет "
    "первый подобный инструмент в ЛКМ-индустрии."
)

add_heading_styled("Уровень 2: Веб-калькулятор на сайте (сентябрь-октябрь 2026)", level=2)

add_body(
    "Адаптация Excel-калькулятора в веб-версию на сайте empils.ru. Интеграция с каталогом продуктов. "
    "Сохранение результатов расчёта в PDF для клиента. Мобильная адаптация. Каждый расчёт — "
    "потенциальная заявка (лидогенерация). Аналог: Sherwin-Williams ColorSnap работает как "
    "интерактивный инструмент, но для подбора цвета, а не расчёта стоимости."
)

add_heading_styled("Уровень 3: POSm-интеграция (октябрь 2026)", level=2)

add_body(
    "Добавление TCO-сравнения «Эконом: 6 450 ₽/год → Премиум: 2 447 ₽/год» на ценники, выкрасы, "
    "полочные бирки. Shelf talker с QR-кодом на веб-калькулятор. POSm «под задачу» в DIY-ритейле. "
    "Ценник на полке сам продаёт премиум — клиент видит выгоду в год эксплуатации без помощи продавца. "
    "Аналог: Dulux Weathershield использует «10 Year Guarantee» на банке как встроенный продавец. "
    "Эмпилс пойдёт дальше — не просто гарантия, а конкретный расчёт стоимости."
)

add_heading_styled("Ключевое преимущество TCO-подхода", level=2)

add_rich_para([
    ("Ни один конкурент в мире не использует TCO-калькулятор как POSm-инструмент. ", True, False, ACCENT_ORANGE, 11),
    ("Dulux обосновывает цену гарантией. Farrow & Ball — консультацией. Sherwin-Williams — интерактивным подбором цвета. "
     "Но никто не даёт клиенту прозрачный расчёт: «Премиум дешевле в 2,6 раза в год». Это уникальное "
     "конкурентное преимущество, которое Эмпилс должен занять первым. Калькулятор работает 24/7 — "
     "в магазине, на сайте, на полке. Он не устаёт, не забывает аргументы, не нуждается в обучении. "
     "Это самый масштабируемый инструмент продаж в программе развития Эмпилс.", False, False, BODY_COLOR, 11),
])

# ===== SAVE =====
doc.save(OUT)
print(f"Saved: {OUT}")
