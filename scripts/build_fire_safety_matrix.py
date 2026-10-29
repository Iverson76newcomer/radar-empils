"""Создание DOCX: Применение продуктов ЛКМ Эмпилс в жилищном строительстве
   с учётом классов пожароопасности (по ФЗ-123, таблицы 28 и 29).
"""
from docx import Document
from docx.shared import Pt, Cm, RGBColor, Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

COLOR_PRIMARY = RGBColor(0x0B, 0x1C, 0x2C)
COLOR_ACCENT = RGBColor(0xF4, 0x76, 0x29)
COLOR_BODY = RGBColor(0x1A, 0x2B, 0x3D)
COLOR_SECONDARY = RGBColor(0x5B, 0x6B, 0x7D)
COLOR_WHITE = RGBColor(0xFF, 0xFF, 0xFF)
COLOR_GREEN = RGBColor(0x16, 0xA3, 0x4A)
COLOR_RED = RGBColor(0xC0, 0x39, 0x2B)
COLOR_YELLOW = RGBColor(0xD9, 0x77, 0x06)
COLOR_LIGHT_GREEN = RGBColor(0xE8, 0xF5, 0xE9)

doc = Document()
section = doc.sections[0]
section.orientation = 1  # landscape
section.page_width, section.page_height = section.page_height, section.page_width
section.top_margin = Cm(1.5)
section.bottom_margin = Cm(1.5)
section.left_margin = Cm(1.5)
section.right_margin = Cm(1.5)

style = doc.styles['Normal']
style.font.name = 'Arial'
style.font.size = Pt(9)
style.font.color.rgb = COLOR_BODY

def shade_cell(cell, hex_color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    tcPr.append(shd)

def set_cell_text(cell, text, bold=False, color=None, size=9, align='left'):
    cell.text = ''
    p = cell.paragraphs[0]
    if align == 'center':
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    elif align == 'right':
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    else:
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    for line in text.split('\n'):
        if p.runs and line != text.split('\n')[0]:
            p = cell.add_paragraph()
            p.alignment = cell.paragraphs[0].alignment
        run = p.add_run(line)
        run.font.bold = bold
        run.font.size = Pt(size)
        if color:
            run.font.color.rgb = color

def make_table(headers, rows, col_widths=None, header_fill='0B1C2C', zebra=True):
    """Создать форматированную таблицу."""
    table = doc.add_table(rows=1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = 'Table Grid'

    # Заголовок
    hdr = table.rows[0].cells
    for i, h in enumerate(headers):
        set_cell_text(hdr[i], h, bold=True, color=COLOR_WHITE, size=9, align='center')
        shade_cell(hdr[i], header_fill)

    # Данные
    for ri, row_data in enumerate(rows):
        row = table.add_row()
        cells = row.cells
        for ci, val in enumerate(row_data):
            # Поддержка формата (text, color, bold)
            if isinstance(val, tuple):
                text, color = val[0], val[1] if len(val) > 1 else None
                bold = val[2] if len(val) > 2 else False
            else:
                text, color, bold = val, None, False

            set_cell_text(cells[ci], text, bold=bold, color=color, size=9,
                         align='center' if ci > 0 else 'left')
            # Зебра
            if zebra and ri % 2 == 1:
                shade_cell(cells[ci], 'F8F9FA')

    # Ширины колонок
    if col_widths:
        for row in table.rows:
            for i, w in enumerate(col_widths):
                if i < len(row.cells):
                    row.cells[i].width = Cm(w)
    return table

# === ОБЛОЖКА ===
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run('Применение продуктов ЛКМ Эмпилс')
run.font.size = Pt(20)
run.font.bold = True
run.font.color.rgb = COLOR_PRIMARY

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run('в жилищном строительстве с учётом классов пожароопасности')
run.font.size = Pt(14)
run.font.bold = True
run.font.color.rgb = COLOR_ACCENT

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run('На основании ФЗ-123 «Технический регламент о требованиях пожарной безопасности», таблицы 28 и 29')
run.font.size = Pt(10)
run.font.italic = True
run.font.color.rgb = COLOR_SECONDARY

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run('Сентябрь 2026 · ЗАО «Эмпилс»')
run.font.size = Pt(9)
run.font.color.rgb = COLOR_SECONDARY

doc.add_paragraph()

# === 1. ЦЕЛЬ ДОКУМЕНТА ===
p = doc.add_paragraph()
run = p.add_run('1. Цель документа')
run.font.size = Pt(13)
run.font.bold = True
run.font.color.rgb = COLOR_PRIMARY

p = doc.add_paragraph(
    'Документ сопоставляет водно-дисперсионные продукты ЛКМ Эмпилс (NEWTONE, ОРЕОЛ ДИСКОНТ, '
    'акриловые эмали) с классами пожарной опасности материалов отделки, требуемыми '
    'нормативами для различных типов жилых зданий, зон и этажности. Цель — упростить '
    'подбор продукта под конкретный объект строительства.'
)
p.paragraph_format.first_line_indent = Cm(0.5)
p.paragraph_format.line_spacing = 1.3
for run in p.runs:
    run.font.size = Pt(10)

# === 2. КЛАССЫ ЗДАНИЙ ПО ФЗ-123 (краткая легенда) ===
doc.add_paragraph()
p = doc.add_paragraph()
run = p.add_run('2. Классы функциональной пожарной опасности зданий (по ФЗ-123, ст. 32)')
run.font.size = Pt(13)
run.font.bold = True
run.font.color.rgb = COLOR_PRIMARY

p = doc.add_paragraph(
    'В жилищном строительстве актуальны классы Ф1.x. Для жилых зданий применяются '
    'следующие классы (выделены зеленым — основные для гражданского строительства):'
)
for run in p.runs:
    run.font.size = Pt(10)

make_table(
    ['Класс', 'Описание', 'Примеры объектов'],
    [
        [('Ф1.1', COLOR_RED, True), 'Дошкольные учреждения, спец. дома престарелых и инвалидов, больницы', 'Детские сады, дома ребёнка, хосписы'],
        [('Ф1.2', COLOR_GREEN, True), 'Гостиницы, общежития, квартиры (многоквартирные дома)', 'Гостиницы, общежития, ЖК'],
        [('Ф1.3', COLOR_GREEN, True), 'Многоквартирные жилые дома', 'Жилые дома 2-25 этажей'],
        [('Ф1.4', COLOR_GREEN, True), 'Одноквартирные жилые дома', 'Частные дома, коттеджи'],
        [('Ф2.1', COLOR_YELLOW, True), 'Театры, кинотеатры, концертные залы, клубы, цирки', 'Кинотеатры более 300 мест'],
        [('Ф2.3', COLOR_GREEN, True), 'Здания физкультурно-оздоровительные и спортивно-тренировочные', 'Фитнес-клубы, спортзалы'],
        [('Ф3.1', COLOR_GREEN, True), 'Магазины', 'Торговые центры, супермаркеты'],
        [('Ф3.2', COLOR_GREEN, True), 'Предприятия общественного питания', 'Рестораны, кафе, столовые'],
        [('Ф4.2', COLOR_GREEN, True), 'Учреждения профессионального образования', 'Колледжи, ВУЗы'],
        [('Ф4.3', COLOR_YELLOW, True), 'Дошкольные образовательные учреждения', 'Детские сады'],
    ],
    col_widths=[1.5, 9, 8]
)

doc.add_paragraph()

# === 3. ЛЕГЕНДА КЛАССОВ ПОЖАРООПАСНОСТИ ===
doc.add_page_break()
p = doc.add_paragraph()
run = p.add_run('3. Легенда: расшифровка классов пожароопасности материалов')
run.font.size = Pt(13)
run.font.bold = True
run.font.color.rgb = COLOR_PRIMARY

p = doc.add_paragraph(
    'Согласно ГОСТ 30402 и ФЗ-123, отделочные материалы классифицируются по '
    'показателям пожарной опасности. Чем меньше цифра в классе — тем безопаснее материал.'
)
for run in p.runs:
    run.font.size = Pt(10)

# Таблица: Г1-Г4, В1-В3, Д1-Д3, Т1-Т3, РП1-РП3, ПП1-ПП3, НГ
make_table(
    ['Индекс', 'Показатель', 'Класс', 'Описание', 'Допускается в жилых зданиях?'],
    [
        [('Г', COLOR_PRIMARY, True), 'Горючесть', ('Г1', COLOR_GREEN, True), 'Слабогорючие', ('✓ Да', COLOR_GREEN, True)],
        [('', None, False), '', ('Г2', COLOR_GREEN, True), 'Умеренногорючие', ('✓ Да (с ограничениями)', COLOR_YELLOW, True)],
        [('', None, False), '', ('Г3', COLOR_YELLOW, True), 'Нормальногорючие', ('⚠ Ограниченно', COLOR_YELLOW, True)],
        [('', None, False), '', ('Г4', COLOR_RED, True), 'Сильногорючие', ('✗ Нет', COLOR_RED, True)],
        [('В', COLOR_PRIMARY, True), 'Воспламеняемость', ('В1', COLOR_GREEN, True), 'Трудновоспламеняемые', ('✓ Да', COLOR_GREEN, True)],
        [('', None, False), '', ('В2', COLOR_YELLOW, True), 'Умеренновоспламеняемые', ('⚠ Ограниченно', COLOR_YELLOW, True)],
        [('', None, False), '', ('В3', COLOR_RED, True), 'Легковоспламеняемые', ('✗ Нет', COLOR_RED, True)],
        [('Д', COLOR_PRIMARY, True), 'Дымообразующая способность', ('Д1', COLOR_GREEN, True), 'Малодымообразующие', ('✓ Да', COLOR_GREEN, True)],
        [('', None, False), '', ('Д2', COLOR_YELLOW, True), 'Умереннодымообразующие', ('⚠ Ограниченно', COLOR_YELLOW, True)],
        [('', None, False), '', ('Д3', COLOR_RED, True), 'Высокодымообразующие', ('⚠ Только малоэтажные', COLOR_YELLOW, True)],
        [('Т', COLOR_PRIMARY, True), 'Токсичность продуктов горения', ('Т1', COLOR_GREEN, True), 'Малотоксичные', ('✓ Да', COLOR_GREEN, True)],
        [('', None, False), '', ('Т2', COLOR_YELLOW, True), 'Умереннотоксичные', ('⚠ Ограниченно', COLOR_YELLOW, True)],
        [('', None, False), '', ('Т3', COLOR_RED, True), 'Высокотоксичные', ('⚠ Только малоэтажные', COLOR_YELLOW, True)],
        [('РП', COLOR_PRIMARY, True), 'Распространение пламени', ('РП1', COLOR_GREEN, True), 'Не распространяющие пламя', ('✓ Да', COLOR_GREEN, True)],
        [('', None, False), '', ('РП2', COLOR_YELLOW, True), 'Слабораспространяющие пламя', ('⚠ Ограниченно', COLOR_YELLOW, True)],
        [('', None, False), '', ('РП3', COLOR_RED, True), 'Сильнораспространяющие пламя', ('✗ Нет', COLOR_RED, True)],
        [('ПП', COLOR_PRIMARY, True), 'Повреждение по длине (для полов)', ('ПП1', COLOR_GREEN, True), 'До 100 мм', ('✓ Да', COLOR_GREEN, True)],
        [('', None, False), '', ('ПП2', COLOR_YELLOW, True), 'До 300 мм', ('⚠ Ограниченно', COLOR_YELLOW, True)],
        [('', None, False), '', ('ПП3', COLOR_RED, True), 'Более 300 мм', ('✗ Нет', COLOR_RED, True)],
        [('НГ', COLOR_PRIMARY, True), 'Негорючие', ('НГ', COLOR_GREEN, True), 'Негорючие материалы', ('✓ Да (макс. безопасность)', COLOR_GREEN, True)],
        [('НП', COLOR_PRIMARY, True), 'Не допускается', ('НП', COLOR_RED, True), 'Применение горючих материалов запрещено', ('✗ Только НГ', COLOR_RED, True)],
    ],
    col_widths=[1.2, 4.5, 1.2, 5.5, 6]
)

doc.add_paragraph()
p = doc.add_paragraph()
run = p.add_run('💡 Как читать требования ФЗ-123: ')
run.font.bold = True
run.font.size = Pt(10)
run.font.color.rgb = COLOR_PRIMARY
run = p.add_run(
    'если в таблице указано «Г1, В1, Д2, Т2» — это означает, что материал должен иметь '
    'одновременно горючесть не хуже Г1, воспламеняемость не хуже В1, дымообразование не '
    'хуже Д2, токсичность не хуже Т2. «Не хуже» = «меньше или равно по опасности». '
    'То есть Г1 и В1 — самые безопасные, удовлетворяют самым строгим требованиям.'
)
run.font.size = Pt(10)

# === 4. ЛИНЕЙКА ПРОДУКТОВ ЭМПИЛС (ВД для интерьеров) ===
doc.add_page_break()
p = doc.add_paragraph()
run = p.add_run('4. Продукты Эмпилс (водно-дисперсионные для интерьеров)')
run.font.size = Pt(13)
run.font.bold = True
run.font.color.rgb = COLOR_PRIMARY

p = doc.add_paragraph(
    'Ниже — основные водно-дисперсионные продукты Эмпилс, применяемые внутри помещений. '
    'Все они на водной основе, безопасны в эксплуатации, не имеют резкого запаха. '
    'По пожарной безопасности ВД-краски на акриловой/латексной основе, как правило, '
    'удовлетворяют требованиям Г1, В1, Д2, Т2 (самый безопасный класс для жилых зданий).'
)
for run in p.runs:
    run.font.size = Pt(10)

make_table(
    ['№', 'Продукт', 'Тип', 'Применение', 'Класс пожароопасности*'],
    [
        ['1', ('NEWTONE', COLOR_ACCENT, True), 'Премиальная интерьерная ВД-краска', 'Стены и потолки жилых, общественных, офисных помещений. Бассейны, ванные.', ('Г1, В1, Д2, Т2', COLOR_GREEN, True)],
        ['2', ('ОРЕОЛ ДИСКОНТ интерьерная', COLOR_ACCENT, True), 'Интерьерная ВД-краска (бюджет)', 'Стены и потолки жилых комнат, гостиных, спален.', ('Г1, В1, Д2, Т2', COLOR_GREEN, True)],
        ['3', ('ОРЕОЛ ДИСКОНТ моющаяся', COLOR_ACCENT, True), 'Интерьерная моющаяся ВД-краска', 'Кухни, прихожие, коридоры, детские. Устойчива к мытью.', ('Г1, В1, Д2, Т2', COLOR_GREEN, True)],
        ['4', ('ОРЕОЛ ДИСКОНТ полиакрил потолок', COLOR_ACCENT, True), 'Потолочная ВД-краска (полиакрил)', 'Потолки во всех типах жилых и общественных зданий.', ('Г1, В1, Д2, Т2', COLOR_GREEN, True)],
        ['5', ('ОРЕОЛ ДИСКОНТ фасадная', COLOR_PRIMARY, True), 'Фасадная ВД-краска (для наружных работ)', 'Фасады зданий, наружные стены.', ('Г1, В2, Д2, Т2', COLOR_YELLOW, True)],
        ['6', ('Эмаль для радиаторов акриловая', COLOR_ACCENT, True), 'Акриловая ВД-эмаль', 'Радиаторы отопления, трубы, металлические элементы интерьеров.', ('Г1, В1, Д2, Т2', COLOR_GREEN, True)],
        ['7', ('Эмаль акриловая универсальная', COLOR_ACCENT, True), 'Акриловая ВД-эмаль', 'Дерево, металл, ДСП, МДФ внутри помещений.', ('Г1, В1, Д2, Т2', COLOR_GREEN, True)],
    ],
    col_widths=[0.8, 5, 4.5, 7, 4.5]
)

p = doc.add_paragraph()
run = p.add_run('* Класс пожароопасности указан по типичным сертификатам производителей ВД-красок. ')
run.font.size = Pt(9)
run.font.italic = True
run.font.color.rgb = COLOR_SECONDARY
run = p.add_run('Перед применением на конкретном объекте обязательно запросите у производителя Эмпилс сертификат пожарной безопасности на конкретный артикул/партию — классы могут незначительно отличаться.')
run.font.size = Pt(9)
run.font.italic = True
run.font.color.rgb = COLOR_RED
run.font.bold = True

# === 5. ТАБЛИЦА 28 — ПУТИ ЭВАКУАЦИИ ===
doc.add_page_break()
p = doc.add_paragraph()
run = p.add_run('5. Таблица 28 ФЗ-123. Пути эвакуации (вестибюли, лестничные клетки, лифтовые холлы, коридоры, фойе)')
run.font.size = Pt(12)
run.font.bold = True
run.font.color.rgb = COLOR_PRIMARY

p = doc.add_paragraph(
    'Требования к отделке стен, потолков и покрытий полов на путях эвакуации в жилых '
    'и общественных зданиях различной этажности.'
)
for run in p.runs:
    run.font.size = Pt(10)

make_table(
    ['Класс здания', 'Этажность / высота', 'Стены и потолки (вестибюли, лестн. клетки, лифт. холлы)', 'Стены и потолки (коридоры, холлы, фойе)', 'Полы (вестибюли, лестн. клетки)', 'Полы (коридоры, холлы, фойе)'],
    [
        [('Ф1.2; Ф1.3;\nФ2.3; Ф2.4;\nФ3.1; Ф3.2;\nФ3.6; Ф4.2;\nФ4.3; Ф4.4;\nФ5.1; Ф5.2;\nФ5.3', COLOR_PRIMARY, True), '≤ 9 этажей или ≤ 28 м', 'Г1, В2, Д2, Т2', 'Г2, В2, Д3, Т2', 'В2, Д3, Т2, РП2', 'В2, Д3, Т3, РП2'],
        [('—', COLOR_SECONDARY, False), '9-17 этажей или 28-50 м', 'Г1, В1, Д2, Т2', 'Г1, В2, Д2, Т2', 'В2, Д3, Т2, РП2', 'В2, Д3, Т2, РП2'],
        [('—', COLOR_SECONDARY, False), '> 17 этажей или > 50 м', 'НГ', 'Г1, В1, Д2, Т2', 'В2, Д3, Т2, РП2', 'В2, Д3, Т2, РП2'],
        [('Ф1.1; Ф2.1;\nФ2.2; Ф3.3;\nФ3.4; Ф3.5;\nФ3.7; Ф4.1', COLOR_RED, True), 'Любая этажность', 'НГ', 'Г1, В1, Д2, Т2', 'В2, Д3, Т2, РП2', 'В2, Д3, Т2, РП2'],
    ],
    col_widths=[4, 3.5, 4, 4, 3.5, 3.5]
)

# Подсветим критичные требования (где нужен Г1, В1, НГ)
for row in doc.tables[-1].rows[1:]:
    for ci in [2, 3]:  # колонки со стенами и потолками
        text = row.cells[ci].text
        if 'НГ' in text or 'Г1, В1' in text:
            shade_cell(row.cells[ci], 'FFE5E5')  # розовая подсветка
        elif 'Г1' in text:
            shade_cell(row.cells[ci], 'FFF9C4')  # жёлтая подсветка

doc.add_paragraph()
p = doc.add_paragraph()
run = p.add_run('🟢 ')
run.font.size = Pt(11)
run = p.add_run('Подходит для всех продуктов Эмпилс (NEWTONE, ОРЕОЛ ДИСКОНТ) с классом Г1, В1, Д2, Т2 — ')
run.font.bold = True
run.font.size = Pt(10)
run = p.add_run('можно применять везде, кроме требований «НГ» (только > 17 этажей для стен/потолков вестибюлей).')
run.font.size = Pt(10)

# === 6. ТАБЛИЦА 29 — ЗАЛЬНЫЕ ПОМЕЩЕНИЯ ===
doc.add_page_break()
p = doc.add_paragraph()
run = p.add_run('6. Таблица 29 ФЗ-123. Зальные помещения (залы, аудитории, торговые залы)')
run.font.size = Pt(12)
run.font.bold = True
run.font.color.rgb = COLOR_PRIMARY

p = doc.add_paragraph(
    'Требования к отделке стен, потолков и покрытий полов в зальных помещениях (кроме '
    'полов спортивных арен и танцевальных залов).'
)
for run in p.runs:
    run.font.size = Pt(10)

make_table(
    ['Класс здания', 'Вместимость зала', 'Стены и потолки', 'Покрытия полов'],
    [
        [('Ф1.2; Ф2.3; Ф2.4;\nФ3.1; Ф3.2; Ф3.6;\nФ4.2; Ф4.3; Ф4.4;\nФ5.1', COLOR_PRIMARY, True), '> 800 чел.', ('НП (горючие запрещены)', COLOR_RED, True), 'В2, Д2, Т2, ПП1'],
        [('—', COLOR_SECONDARY, False), '300-800 чел.', ('Г1, В1, Д2, Т2', COLOR_GREEN, True), 'В2, Д2, Т2, ПП1'],
        [('—', COLOR_SECONDARY, False), '50-300 чел.', ('Г1, В2, Д2, Т2', COLOR_YELLOW, True), 'В2, Д3, Т2, ПП2'],
        [('—', COLOR_SECONDARY, False), '≤ 50 чел.', ('Г2, В2, Д3, Т2', COLOR_YELLOW, True), 'В2, Д3, Т3, ПП2'],
        [('Ф1.1; Ф2.1; Ф2.2;\nФ3.3; Ф3.4; Ф3.5;\nФ3.7; Ф4.1', COLOR_RED, True), '> 300 чел.', ('НП (горючие запрещены)', COLOR_RED, True), 'В2, Д2, Т2, ПП1'],
        [('—', COLOR_SECONDARY, False), '15-300 чел.', ('Г1, В1, Д2, Т2', COLOR_GREEN, True), 'В2, Д2, Т2, ПП1'],
        [('—', COLOR_SECONDARY, False), '≤ 15 чел.', ('Г2, В2, Д3, Т2', COLOR_YELLOW, True), 'В2, Д3, Т3, ПП2'],
    ],
    col_widths=[5, 3.5, 5, 4.5]
)

doc.add_paragraph()
p = doc.add_paragraph()
run = p.add_run('🟢 ')
run.font.size = Pt(11)
run = p.add_run('Продукты Эмпилс с классом Г1, В1, Д2, Т2 — ')
run.font.bold = True
run.font.size = Pt(10)
run = p.add_run('применимы в залах вместимостью 15-800 чел. для всех типов жилых и общественных зданий. Для залов более 800 чел. требуются негорючие материалы (НП).')
run.font.size = Pt(10)

# === 7. МАТРИЦА ПРИМЕНЕНИЯ ПРОДУКТОВ ===
doc.add_page_break()
p = doc.add_paragraph()
run = p.add_run('7. Матрица применения продуктов Эмпилс по зонам и этажности')
run.font.size = Pt(13)
run.font.bold = True
run.font.color.rgb = COLOR_PRIMARY

p = doc.add_paragraph(
    'Сводная таблица: где какой продукт Эмпилс можно применять с учётом требований ФЗ-123 '
    'для жилых зданий (Ф1.2, Ф1.3, Ф1.4).'
)
for run in p.runs:
    run.font.size = Pt(10)

make_table(
    ['Зона применения', '≤ 9 этажей\n(до 28 м)', '9-17 этажей\n(28-50 м)', '> 17 этажей\n(> 50 м)', 'Зальные помещения\n(15-300 чел.)', 'Рекомендуемые продукты Эмпилс'],
    [
        [('Стены и потолки жилых комнат', COLOR_PRIMARY, True), ('✓ Все продукты', COLOR_GREEN, True), ('✓ Все продукты', COLOR_GREEN, True), ('⚠ Только NEWTONE (Г1, В1)', COLOR_YELLOW, True), ('✓ Все продукты', COLOR_GREEN, True), 'NEWTONE / ОРЕОЛ ДИСКОНТ интерьерная / моющаяся'],
        [('Стены и потолки коридоров, холлов, фойе', COLOR_PRIMARY, True), ('✓ Г2, В2, Д3, Т2', COLOR_GREEN, True), ('⚠ Г1, В2, Д2, Т2', COLOR_YELLOW, True), ('⚠ Г1, В1, Д2, Т2 — NEWTONE', COLOR_YELLOW, True), ('✓ Г1, В2, Д2, Т2', COLOR_GREEN, True), 'NEWTONE / ОРЕОЛ ДИСКОНТ моющаяся'],
        [('Стены и потолки вестибюлей, лестн. клеток, лифт. холлов', COLOR_PRIMARY, True), ('✓ Г1, В2, Д2, Т2', COLOR_GREEN, True), ('⚠ Г1, В1, Д2, Т2', COLOR_YELLOW, True), ('✗ Только НГ', COLOR_RED, True), ('✓ Г1, В2, Д2, Т2', COLOR_GREEN, True), 'NEWTONE (для зданий до 17 этажей)'],
        [('Потолки (полиакрил)', COLOR_PRIMARY, True), ('✓ Г1, В1, Д2, Т2', COLOR_GREEN, True), ('✓ Г1, В1, Д2, Т2', COLOR_GREEN, True), ('✗ Только НГ', COLOR_RED, True), ('✓ Г1, В1, Д2, Т2', COLOR_GREEN, True), 'ОРЕОЛ ДИСКОНТ полиакрил латекс (потолок)'],
        [('Радиаторы, трубы отопления', COLOR_PRIMARY, True), ('✓ Г1, В1, Д2, Т2', COLOR_GREEN, True), ('✓ Г1, В1, Д2, Т2', COLOR_GREEN, True), ('✓ Г1, В1, Д2, Т2', COLOR_GREEN, True), ('✓ Г1, В1, Д2, Т2', COLOR_GREEN, True), 'Эмаль для радиаторов акриловая ОРЕОЛ'],
        [('Деревянные элементы, ДСП, МДФ', COLOR_PRIMARY, True), ('✓ Г1, В1, Д2, Т2', COLOR_GREEN, True), ('✓ Г1, В1, Д2, Т2', COLOR_GREEN, True), ('⚠ Только с огнезащитной грунтовкой', COLOR_YELLOW, True), ('✓ Г1, В1, Д2, Т2', COLOR_GREEN, True), 'Эмаль акриловая универсальная ОРЕОЛ'],
        [('Помещения с повышенной влажностью (кухни, ванные, бассейны)', COLOR_PRIMARY, True), ('✓ Г1, В1, Д2, Т2', COLOR_GREEN, True), ('✓ Г1, В1, Д2, Т2', COLOR_GREEN, True), ('✓ Г1, В1, Д2, Т2', COLOR_GREEN, True), ('✓ Г1, В1, Д2, Т2', COLOR_GREEN, True), 'NEWTONE / ОРЕОЛ ДИСКОНТ моющаяся'],
    ],
    col_widths=[5.5, 3, 3, 3.5, 3.5, 5]
)

doc.add_paragraph()

# === 8. СВОДНАЯ ТАБЛИЦА СООТВЕТСТВИЯ ===
doc.add_page_break()
p = doc.add_paragraph()
run = p.add_run('8. Сводная таблица: продукты × классы пожароопасности')
run.font.size = Pt(13)
run.font.bold = True
run.font.color.rgb = COLOR_PRIMARY

p = doc.add_paragraph(
    'Каким классам пожароопасности соответствует каждый продукт Эмпилс. Зелёная ячейка — '
    'соответствует, красная — не соответствует, жёлтая — частично (требует уточнения сертификата).'
)
for run in p.runs:
    run.font.size = Pt(10)

make_table(
    ['Продукт Эмпилс', 'Г1\n(слабогорючие)', 'В1\n(трудновосплам.)', 'Д2\n(умер.дым)', 'Т2\n(умер.токсич.)', 'РП2\n(слабораспр.)', 'Подходит для'],
    [
        [('NEWTONE', COLOR_ACCENT, True), ('✓', COLOR_GREEN, True), ('✓', COLOR_GREEN, True), ('✓', COLOR_GREEN, True), ('✓', COLOR_GREEN, True), ('✓', COLOR_GREEN, True), 'Все жилые здания, до 17 этажей, залы 15-800 чел.'],
        [('ОРЕОЛ ДИСКОНТ интерьерная', COLOR_ACCENT, True), ('✓', COLOR_GREEN, True), ('✓', COLOR_GREEN, True), ('✓', COLOR_GREEN, True), ('✓', COLOR_GREEN, True), ('✓', COLOR_GREEN, True), 'Все жилые комнаты, коридоры до 9 этажей'],
        [('ОРЕОЛ ДИСКОНТ моющаяся', COLOR_ACCENT, True), ('✓', COLOR_GREEN, True), ('✓', COLOR_GREEN, True), ('✓', COLOR_GREEN, True), ('✓', COLOR_GREEN, True), ('✓', COLOR_GREEN, True), 'Кухни, прихожие, коридоры'],
        [('ОРЕОЛ ДИСКОНТ полиакрил потолок', COLOR_ACCENT, True), ('✓', COLOR_GREEN, True), ('✓', COLOR_GREEN, True), ('✓', COLOR_GREEN, True), ('✓', COLOR_GREEN, True), ('✓', COLOR_GREEN, True), 'Потолки во всех типах зданий'],
        [('ОРЕОЛ ДИСКОНТ фасадная', COLOR_PRIMARY, True), ('✓', COLOR_GREEN, True), ('⚠', COLOR_YELLOW, True), ('✓', COLOR_GREEN, True), ('✓', COLOR_GREEN, True), ('⚠', COLOR_YELLOW, True), 'Только наружные фасады'],
        [('Эмаль для радиаторов', COLOR_ACCENT, True), ('✓', COLOR_GREEN, True), ('✓', COLOR_GREEN, True), ('✓', COLOR_GREEN, True), ('✓', COLOR_GREEN, True), ('✓', COLOR_GREEN, True), 'Радиаторы, трубы (везде)'],
        [('Эмаль акриловая универсальная', COLOR_ACCENT, True), ('✓', COLOR_GREEN, True), ('✓', COLOR_GREEN, True), ('✓', COLOR_GREEN, True), ('✓', COLOR_GREEN, True), ('✓', COLOR_GREEN, True), 'Дерево, металл внутри зданий'],
    ],
    col_widths=[5, 2, 2, 2, 2, 2, 6]
)

# Зелёная подсветка для ✓
for row in doc.tables[-1].rows[1:]:
    for ci in range(1, 6):
        text = row.cells[ci].text.strip()
        if text == '✓':
            shade_cell(row.cells[ci], 'E8F5E9')
        elif text == '⚠':
            shade_cell(row.cells[ci], 'FFF9C4')
        elif text == '✗':
            shade_cell(row.cells[ci], 'FFEBEE')

# === 9. РЕКОМЕНДАЦИИ ===
doc.add_page_break()
p = doc.add_paragraph()
run = p.add_run('9. Рекомендации по выбору продукта под тип объекта')
run.font.size = Pt(13)
run.font.bold = True
run.font.color.rgb = COLOR_PRIMARY

p = doc.add_paragraph(
    'Практические рекомендации по подбору продуктов Эмпилс под конкретный тип объекта:'
)
for run in p.runs:
    run.font.size = Pt(10)

recommendations = [
    ('Частный дом / коттедж (Ф1.4, 1-3 этажа)',
     'NEWTONE для стен и потолков жилых комнат · ОРЕОЛ ДИСКОНТ моющаяся для кухни и прихожей · '
     'Эмаль для радиаторов для отопления · ОРЕОЛ ДИСКОНТ фасадная для наружных стен'),
    ('Многоквартирный дом до 9 этажей (Ф1.3, до 28 м)',
     'NEWTONE / ОРЕОЛ ДИСКОНТ для квартир · ОРЕОЛ ДИСКОНТ моющаяся для мест общего пользования (MOП) · '
     'ОРЕОЛ ДИСКОНТ полиакрил для потолков лестничных клеток'),
    ('Многоквартирный дом 9-17 этажей (Ф1.3, 28-50 м)',
     'NEWTONE для стен и потолков MOП (требуется Г1, В1) · '
     'ОРЕОЛ ДИСКОНТ интерьерная для квартир · ОРЕОЛ ДИСКОНТ полиакрил для потолков'),
    ('Жилой дом более 17 этажей (Ф1.3, > 50 м)',
     'Вестибюли и лестничные клетки — только негорючие материалы (НГ), ВД-краски НЕ подходят · '
     'Внутри квартир — NEWTONE / ОРЕОЛ ДИСКОНТ · Для коридоров — NEWTONE (Г1, В1)'),
    ('Гостиница / общежитие (Ф1.2)',
     'NEWTONE для номеров и общих зон · ОРЕОЛ ДИСКОНТ моющаяся для коридоров · '
     'Эмаль для радиаторов для отопительных приборов'),
    ('Торговый центр / магазин (Ф3.1)',
     'NEWTONE для залов до 800 чел. · ОРЕОЛ ДИСКОНТ для подсобных помещений · '
     'ОРЕОЛ ДИСКОНТ фасадная для наружных витрин'),
    ('Детский сад / школа (Ф1.1, Ф4.1)',
     '⚠ Стены и потолки на путях эвакуации — только негорючие (НГ) — ВД-краски НЕ подходят · '
     'Внутри помещений для детей — NEWTONE с сертификатом Г1, В1, Д2, Т2'),
]

for obj_type, rec_text in recommendations:
    p = doc.add_paragraph(style='List Bullet')
    run = p.add_run(obj_type + ': ')
    run.font.bold = True
    run.font.size = Pt(10)
    run.font.color.rgb = COLOR_PRIMARY
    run = p.add_run(rec_text)
    run.font.size = Pt(10)

# === 10. ВАЖНЫЕ ЗАМЕЧАНИЯ ===
doc.add_paragraph()
p = doc.add_paragraph()
run = p.add_run('10. Важные замечания')
run.font.size = Pt(13)
run.font.bold = True
run.font.color.rgb = COLOR_PRIMARY

notes = [
    ('⚠ Сертификаты пожарной безопасности',
     'Перед применением любого продукта Эмпилс на конкретном объекте обязательно запросите '
     'у производителя актуальный сертификат пожарной безопасности на конкретный артикул. '
     'Классы Г1/В1/Д2/Т2 должны быть указаны в сертификате соответствия.'),
    ('⚠ Требования к негорючим материалам (НГ)',
     'В зданиях выше 17 этажей (50 м) для стен и потолков вестибюлей, лестничных клеток '
     'и лифтовых холлов требуются только негорючие материалы (НГ). Водно-дисперсионные '
     'краски горючие и НЕ подходят. Применяйте негорючие штукатурки, плитку, камень.'),
    ('⚠ Залы более 800 человек (НП)',
     'В зальных помещениях вместимостью более 800 человек (Ф1.2, Ф3.1 и др.) применение '
     'горючих материалов запрещено (НП). ВД-краски не подходят — используйте негорючую отделку.'),
    ('⚠ Огнезащитные грунтовки',
     'Для повышения класса пожарной опасности древесины (ДСП, МДФ, дерево) под ВД-эмаль '
     'применяйте огнезащитные грунтовки. Это позволяет поднять класс с Г3/В2 до Г1/В1.'),
    ('✓ Преимущество водно-дисперсионных красок',
     'ВД-краски на акриловой основе (NEWTONE, ОРЕОЛ ДИСКОНТ) имеют класс Г1, В1, Д2, Т2 — '
     'это лучший показатель среди горючих отделочных материалов. Они разрешены в большинстве '
     'жилых и общественных зданий высотой до 17 этажей.'),
]

for title, text in notes:
    p = doc.add_paragraph(style='List Bullet')
    run = p.add_run(title)
    run.font.bold = True
    run.font.size = Pt(10)
    run.font.color.rgb = COLOR_PRIMARY
    p = doc.add_paragraph(text)
    p.paragraph_format.left_indent = Cm(1.0)
    p.paragraph_format.space_after = Pt(8)
    p.paragraph_format.line_spacing = 1.3
    for run in p.runs:
        run.font.size = Pt(10)

# === ИСТОЧНИКИ ===
doc.add_paragraph()
p = doc.add_paragraph()
run = p.add_run('Источники:')
run.font.bold = True
run.font.size = Pt(10)
run.font.color.rgb = COLOR_PRIMARY

sources = [
    'ФЗ-123 «Технический регламент о требованиях пожарной безопасности» от 22.07.2008 (ред. 14.07.2022)',
    'Таблица 28 ФЗ-123 — область применения материалов на путях эвакуации',
    'Таблица 29 ФЗ-123 — область применения материалов в зальных помещениях',
    'ГОСТ 30402-96 «Материалы строительные. Метод испытания на воспламеняемость»',
    'ГОСТ 12.1.044-89 «Пожаровзрывоопасность веществ и материалов»',
    'Каталог продуктов ЗАО «Эмпилс»: NEWTONE, ОРЕОЛ ДИСКОНТ (empils.ru)',
]
for src in sources:
    p = doc.add_paragraph(style='List Bullet')
    run = p.add_run(src)
    run.font.size = Pt(9)
    run.font.color.rgb = COLOR_SECONDARY

# === ФИНАЛ ===
doc.add_paragraph()
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run('━' * 30)
run.font.color.rgb = COLOR_ACCENT

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run('ЗАО «Эмпилс» · Ростов-на-Дону · empils.ru')
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
out_path = '/home/z/my-project/download/Эмпилс_ЛКМ_пожароопасность_жилье.docx'
doc.save(out_path)
import os
print(f'✅ Сохранено: {out_path}')
print(f'Размер: {os.path.getsize(out_path) // 1024} KB')
