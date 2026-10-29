"""
Wordstat — план ручного исследования поисковых запросов для TITAN.
Список фраз, пошаговые действия, какие скриншоты сохранять.
"""
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

ORANGE = RGBColor(0xF4, 0x76, 0x29)
DARK   = RGBColor(0x31, 0x31, 0x31)
GREY   = RGBColor(0x6B, 0x72, 0x80)
WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
RED    = RGBColor(0xC0, 0x39, 0x2B)
GREEN  = RGBColor(0x16, 0xA3, 0x4A)
AMBER  = RGBColor(0xB8, 0x86, 0x0B)
BLUE   = RGBColor(0x1E, 0x3A, 0x5F)

doc = Document()
style = doc.styles['Normal']
style.font.name = 'Arial'
style.font.size = Pt(11)
style.element.rPr.rFonts.set(qn('w:eastAsia'), 'Arial')

for section in doc.sections:
    section.top_margin = Cm(2)
    section.bottom_margin = Cm(2)
    section.left_margin = Cm(2.0)
    section.right_margin = Cm(2.0)
    section.page_height = Cm(29.7)
    section.page_width = Cm(21.0)

def set_spacing(p, before=0, after=6, line=1.3):
    pf = p.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.line_spacing = line

def add_run(p, text, size=11, bold=False, color=None, italic=False):
    r = p.add_run(text)
    r.font.size = Pt(size)
    r.bold = bold
    r.italic = italic
    if color:
        r.font.color.rgb = color
    r.font.name = 'Arial'
    return r

def add_para(text, size=11, bold=False, color=None, italic=False, align='left', after=6, before=0):
    p = doc.add_paragraph()
    set_spacing(p, before=before, after=after, line=1.3)
    p.alignment = {'center': WD_ALIGN_PARAGRAPH.CENTER, 'right': WD_ALIGN_PARAGRAPH.RIGHT,
                   'left': WD_ALIGN_PARAGRAPH.LEFT, 'justify': WD_ALIGN_PARAGRAPH.JUSTIFY}.get(align, WD_ALIGN_PARAGRAPH.LEFT)
    add_run(p, text, size=size, bold=bold, color=color, italic=italic)
    return p

def add_h1(text, color=DARK, size=20, before=0, after=8, orange_bar=True):
    p = doc.add_paragraph()
    set_spacing(p, before=before, after=after, line=1.1)
    add_run(p, text, size=size, bold=True, color=color)
    if orange_bar:
        pPr = p._p.get_or_add_pPr()
        pBdr = OxmlElement('w:pBdr')
        left = OxmlElement('w:left')
        left.set(qn('w:val'), 'single')
        left.set(qn('w:sz'), '24')
        left.set(qn('w:space'), '8')
        left.set(qn('w:color'), 'F47629')
        pBdr.append(left)
        pPr.append(pBdr)
        p.paragraph_format.left_indent = Cm(0.3)
    return p

def add_h2(text, color=ORANGE, size=14, before=10, after=4):
    p = doc.add_paragraph()
    set_spacing(p, before=before, after=after, line=1.2)
    add_run(p, text, size=size, bold=True, color=color)
    return p

def add_h3(text, color=DARK, size=12, before=8, after=3):
    p = doc.add_paragraph()
    set_spacing(p, before=before, after=after, line=1.2)
    add_run(p, text, size=size, bold=True, color=color)
    return p

def add_bullet(text, size=11, bold=False):
    p = doc.add_paragraph(style='List Bullet')
    set_spacing(p, before=0, after=3, line=1.25)
    p.paragraph_format.left_indent = Cm(0.6)
    r = p.runs[0] if p.runs else p.add_run()
    r.text = text
    r.font.size = Pt(size)
    r.font.name = 'Arial'
    if bold:
        r.bold = True
    return p

def set_cell_bg(cell, color_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), color_hex)
    tcPr.append(shd)

def set_cell_borders(cell, color='D0D0D0', sz='4'):
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for edge in ('top', 'left', 'bottom', 'right'):
        el = OxmlElement(f'w:{edge}')
        el.set(qn('w:val'), 'single')
        el.set(qn('w:sz'), sz)
        el.set(qn('w:color'), color)
        tcBorders.append(el)
    tcPr.append(tcBorders)

def add_table(headers, rows, col_widths, font_size=10, header_bg='313131'):
    tbl = doc.add_table(rows=1 + len(rows), cols=len(headers))
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    for i, w in enumerate(col_widths):
        for cell in tbl.columns[i].cells:
            cell.width = Cm(w)
    for i, h in enumerate(headers):
        c = tbl.rows[0].cells[i]
        set_cell_bg(c, header_bg)
        set_cell_borders(c, color='313131', sz='4')
        c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_spacing(p, before=2, after=2, line=1.2)
        add_run(p, h, size=font_size, bold=True, color=WHITE)
    for ri, row in enumerate(rows):
        for ci, val in enumerate(row):
            c = tbl.rows[ri + 1].cells[ci]
            if ri % 2 == 1:
                set_cell_bg(c, 'F8F4F0')
            set_cell_borders(c, color='D0D0D0', sz='4')
            c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = c.paragraphs[0]
            set_spacing(p, before=2, after=2, line=1.25)
            add_run(p, str(val), size=font_size)
            if ci == 0:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                p.runs[0].bold = True
                p.runs[0].font.color.rgb = ORANGE
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p = doc.add_paragraph()
    set_spacing(p, before=0, after=4, line=1.0)
    return tbl

def add_alert_box(title, body, color='green'):
    color_map = {
        'red':    ('FEE2E2', 'C0392B'),
        'amber':  ('FEF3C7', 'B8860B'),
        'green':  ('D1FAE5', '166534'),
        'blue':   ('DBEAFE', '1E3A5F'),
    }
    bg, bar = color_map.get(color, color_map['green'])
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    cell = tbl.cell(0, 0)
    cell.width = Cm(17.0)
    set_cell_bg(cell, bg)
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for edge, c in (('top', bg), ('right', bg), ('bottom', bg), ('left', bar)):
        el = OxmlElement(f'w:{edge}')
        el.set(qn('w:val'), 'single')
        el.set(qn('w:sz'), '36' if edge == 'left' else '4')
        el.set(qn('w:color'), c)
        tcBorders.append(el)
    tcPr.append(tcBorders)
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    if title:
        p = cell.paragraphs[0]
        set_spacing(p, before=2, after=2, line=1.2)
        add_run(p, title, size=12, bold=True, color=RGBColor.from_string(bar))
    if body:
        p2 = cell.add_paragraph()
        set_spacing(p2, before=0, after=2, line=1.3)
        add_run(p2, body, size=11, color=DARK)
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for edge, val in (('top', '120'), ('left', '200'), ('bottom', '120'), ('right', '200')):
        el = OxmlElement(f'w:{edge}')
        el.set(qn('w:w'), val)
        el.set(qn('w:type'), 'dxa')
        tcMar.append(el)
    tcPr.append(tcMar)
    p = doc.add_paragraph()
    set_spacing(p, before=0, after=4, line=1.0)
    return tbl

def page_break():
    p = doc.add_paragraph()
    r = p.add_run()
    r.add_break(WD_BREAK.PAGE)

def hr():
    p = doc.add_paragraph()
    set_spacing(p, before=2, after=6, line=1.0)
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '6')
    bottom.set(qn('w:color'), 'F47629')
    pBdr.append(bottom)
    pPr.append(pBdr)

print("Utils ready")

# ============================================================
# ЗАГОЛОВОК
# ============================================================
add_h1('Wordstat — план исследования поисковых запросов', size=22, before=0, after=2)
add_para('Для запуска контекстной и таргетированной рекламы TITAN в Яндекс Директ',
         size=14, bold=True, color=DARK, after=2)
add_para('Ручной способ через wordstat.yandex.ru — бесплатно, ~2–3 часа работы',
         size=11, italic=True, color=GREY, after=10)

hr()

add_para('О документе', bold=True, size=14, color=DARK, before=4, after=4)
add_para(
    'Документ содержит: (1) список из 25 ключевых фраз для исследования; '
    '(2) пошаговый план действий с инструкциями, какие вкладки открывать и '
    'что записывать; (3) список скриншотов, которые нужно сохранять на каждом '
    'шаге. После сбора данных пришлите скриншоты — я сведу их в Excel-таблицу '
    'и подготовлю рекомендации по Яндекс Директ.',
    after=8
)

add_alert_box('Что вы получите в итоге',
    'Excel-таблицу со 100+ ключевыми фразами, распределением по регионам РФ, '
    'объёмами показов и рекомендациями: какие фразы, в каких регионах и в какие '
    'месяцы запускать рекламу в Яндекс Директ для TITAN.',
    color='green')

page_break()

# ============================================================
# ШАГ 0. ПОДГОТОВКА
# ============================================================
add_h1('Шаг 0. Подготовка к работе', size=18, before=0)

add_para('Перед началом — 5 минут подготовки:')

add_h3('1. Войдите в Яндекс', color=ORANGE)
add_para('Откройте https://wordstat.yandex.ru и авторизуйтесь через Яндекс ID. '
         'Без авторизации Wordstat не показывает полные данные.')

add_h3('2. Создайте папку для скриншотов', color=ORANGE)
add_para('На компьютере создайте папку с названием «Wordstat TITAN» — туда будете '
         'сохранять все скриншоты. Всего нужно будет сделать ~75 скриншотов '
         '(25 фраз × 3 вкладки).')

add_h3('3. Подготовьте Excel-таблицу для записей', color=ORANGE)
add_para('Создайте пустую Excel-таблицу с колонками:', after=4)
add_table(
    headers=['Колонка', 'Что записывать'],
    rows=[
        ['№',                'Порядковый номер фразы'],
        ['Ключевая фраза',    'Фраза, которую вводили в Wordstat'],
        ['Показов в месяц',   'Число показов (из левой колонки Wordstat)'],
        ['Регион',            'Название региона из вкладки «По регионам»'],
        ['Показов в регионе',  'Число показов в этом регионе'],
        ['Вкладка',           '«Левая» / «Правая» / «По регионам» / «История»'],
        ['Скриншот',          'Имя файла скриншота'],
    ],
    col_widths=[4.0, 13.0],
    font_size=10
)

add_h3('4. Откройте Wordstat', color=ORANGE)
add_para('URL: https://wordstat.yandex.ru — должна открыться строка поиска с '
         'кнопкой «Подобрать».')

page_break()

# ============================================================
# ШАГ 1. СПИСОК КЛЮЧЕВЫХ ФРАЗ
# ============================================================
add_h1('Шаг 1. Список ключевых фраз для исследования', size=18, before=0)

add_para(
    'Ниже — 25 фраз, разделённых на 4 группы. Вводите каждую фразу в Wordstat '
    'по очереди. Фразы сгруппированы по типу: брендовые, продуктовые, '
    'проблемные и конкурентные.',
    after=8
)

add_alert_box('Как вводить фразы',
    'Каждую фразу вводите БЕЗ кавычек и БЕЗ плюсиков — просто текст в строку. '
    'Нажимайте «Подобрать». Wordstat покажет 2 колонки: левая — что искали '
    'ВМЕСТЕ с фразой, правая — похожие запросы. Обе колонки нужны.',
    color='blue')

# --- Группа 1: Брендовые ---
add_h2('Группа 1. Брендовые запросы (5 фраз)', before=8, after=4)
add_para('Цель: понять, как ищут именно бренд TITAN и Эмпилс.', italic=True, color=GREY, after=4)

add_table(
    headers=['№', 'Фраза', 'Что ищем'],
    rows=[
        ['1',  'TITAN краска',              'Как ищут бренд TITAN'],
        ['2',  'Ореол TITAN',                'Полное название продукта'],
        ['3',  'Эмпилс',                     'Как ищут компанию-производителя'],
        ['4',  'краска TITAN по металлу',   'Бренд + назначение'],
        ['5',  'грунт-эмаль TITAN',          'Бренд + категория продукта'],
    ],
    col_widths=[0.8, 5.5, 10.7],
    font_size=10
)

# --- Группа 2: Продуктовые ---
add_h2('Группа 2. Продуктовые запросы (10 фраз)', before=8, after=4)
add_para('Цель: понять, как ищут категорию продуктов (грунт-эмали по металлу) '
         'без привязки к бренду. Это самые важные фразы для Яндекс Директ.',
         italic=True, color=GREY, after=4)

add_table(
    headers=['№', 'Фраза', 'Что ищем'],
    rows=[
        ['6',  'грунт-эмаль',                 'Категория в целом'],
        ['7',  'грунт-эмаль по ржавчине',      'Категория + проблема'],
        ['8',  'грунт-эмаль 3 в 1',            'Категория + технология'],
        ['9',  'краска для металла',            'Широкий продуктовый запрос'],
        ['10', 'краска по металлу',             'Альтернативная формулировка'],
        ['11', 'эмаль по металлу',              'Узкий продуктовый запрос'],
        ['12', 'быстросохнущая краска по металлу', 'Категория + свойство'],
        ['13', 'молотковая краска',             'Категория молотковых эмалей'],
        ['14', 'краска по ржавчине',           'Проблемный запрос'],
        ['15', 'грунт по металлу',              'Грунтовочный продукт'],
    ],
    col_widths=[0.8, 5.5, 10.7],
    font_size=10
)

# --- Группа 3: Применение ---
add_h2('Группа 3. Запросы по применению (6 фраз)', before=8, after=4)
add_para('Цель: понять, как ищут краску под конкретные объекты — ворота, '
         'гаражи, заборы. Это запросы с высокой конверсией в покупку.',
         italic=True, color=GREY, after=4)

add_table(
    headers=['№', 'Фраза', 'Что ищем'],
    rows=[
        ['16', 'краска для ворот',          'Краска для ворот'],
        ['17', 'краска для забора',          'Краска для забора'],
        ['18', 'краска для гаража',          'Краска для гаража'],
        ['19', 'краска для металлической двери', 'Краска для дверей'],
        ['20', 'покрасить металлический забор', 'Действие + объект'],
        ['21', 'чем покрасить металл',       'Вопросительный запрос'],
    ],
    col_widths=[0.8, 5.5, 10.7],
    font_size=10
)

# --- Группа 4: Конкурентные ---
add_h2('Группа 4. Конкурентные запросы (4 фразы)', before=8, after=4)
add_para('Цель: понять, как ищут конкурентов TITAN — для перехвата трафика '
         'в Яндекс Директ через рекламу по фразам конкурентов.',
         italic=True, color=GREY, after=4)

add_table(
    headers=['№', 'Фраза', 'Что ищем'],
    rows=[
        ['22', 'Dali краска по металлу',     'Конкурент Dali'],
        ['23', 'Салван краска',               'Конкурент Салван'],
        ['24', 'Husky краска',                'Конкурент Husky Klondike'],
        ['25', 'Dufa краска по металлу',      'Конкурент Dufa CORREX'],
    ],
    col_widths=[0.8, 5.5, 10.7],
    font_size=10
)

page_break()

# ============================================================
# ШАГ 2. ПОШАГОВЫЙ ПЛАН ДЕЙСТВИЙ
# ============================================================
add_h1('Шаг 2. Пошаговый план для каждой фразы', size=18, before=0)

add_para(
    'Для КАЖДОЙ из 25 фраз нужно сделать 3 действия. На каждое — 1 скриншот. '
    'Итого: 25 фраз × 3 скриншота = 75 скриншотов. Время на 1 фразу — ~5 минут. '
    'Общее время — ~2–3 часа.',
    after=8
)

# --- Действие 1 ---
add_h2('Действие 1. Вкладка «По словам» — левая и правая колонки', before=4, after=4)

add_para('Что делать:', bold=True, after=2)
add_bullet('Введите фразу в строку поиска Wordstat.')
add_bullet('Нажмите «Подобрать».')
add_bullet('Появятся 2 колонки: ЛЕВАЯ («Что искали вместе с фразой») и '
           'ПРАВАЯ («Похожие запросы»).')
add_bullet('Сделайте скриншот всей страницы — обе колонки должны быть видны.')

add_para('Что записать в Excel:', bold=True, after=2)
add_bullet('Топ-10 фраз из ЛЕВОЙ колонки с количеством показов.')
add_bullet('Топ-5 фраз из ПРАВОЙ колонки с количеством показов.')

add_para('Скриншот:', bold=True, after=2)
add_bullet('Имя файла: «01_TITAN_краска_слова.png» (номер фразы + фраза + «слова»).', bold=True)
add_bullet('На скриншоте должны быть видны: введённая фраза, обе колонки с числами.')

# --- Действие 2 ---
add_h2('Действие 2. Вкладка «По регионам» — распределение по регионам РФ', before=8, after=4)

add_para('Что делать:', bold=True, after=2)
add_bullet('НЕ меняя фразу в строке поиска — нажмите вкладку «По регионам» '
           '(вверху, рядом с «По словам»).')
add_bullet('Появится карта России с регионами, окрашенными по интенсивности.')
add_bullet('Под картой — таблица регионов с количеством показов.')

add_para('Что записать в Excel:', bold=True, after=2)
add_bullet('Топ-10 регионов по количеству показов — с числами.')
add_bullet('Общее число показов по России (вверху таблицы).')
add_bullet('Отметьте регионы, где показов > 1000 — это приоритет для рекламы.')

add_para('Скриншот:', bold=True, after=2)
add_bullet('Имя файла: «01_TITAN_краска_регионы.png» (номер + фраза + «регионы»).', bold=True)
add_bullet('На скриншоте должны быть видны: карта + таблица с топ-10 регионами.')

# --- Действие 3 ---
add_h2('Действие 3. Вкладка «История запросов» — динамика по месяцам', before=8, after=4)

add_para('Что делать:', bold=True, after=2)
add_bullet('НЕ меняя фразу — нажмите вкладку «История запросов» '
           '(третья вкладка вверху).')
add_bullet('Появится график с динамикой по месяцам за последние 2 года.')
add_bullet('Под графиком — таблица с числами по месяцам.')

add_para('Что записать в Excel:', bold=True, after=2)
add_bullet('Месяцы с пиком показов (максимальные числа).')
add_bullet('Месяцы с минимумом показов.')
add_bullet('Общий тренд — растёт, падает или стабильно.')

add_para('Скриншот:', bold=True, after=2)
add_bullet('Имя файла: «01_TITAN_краска_история.png» (номер + фраза + «история»).', bold=True)
add_bullet('На скриншоте должен быть виден график с пиками и спадами.')

add_alert_box('Важное замечание',
    'Если по фразе показов мало (меньше 50 в месяц) — всё равно сделайте '
    'скриншоты. Это тоже важная информация: фраза не работает, на неё '
    'не нужно запускать рекламу в Директе. Отметьте в Excel: «менее 50 показов — не использовать».',
    color='amber')

page_break()

# ============================================================
# ШАГ 3. ЧЕК-ЛИСТ СКРИНШОТОВ
# ============================================================
add_h1('Шаг 3. Чек-лист скриншотов (75 шт.)', size=18, before=0)

add_para(
    'Полный список всех скриншотов, которые нужно сделать. Отмечайте галочкой '
    'после каждого скриншота. Имя файла = номер фразы + фраза (коротко) + тип.',
    after=8
)

add_table(
    headers=['№', 'Фраза', 'Слова', 'Регионы', 'История'],
    rows=[
        ['1',  'TITAN краска',                    '☐', '☐', '☐'],
        ['2',  'Ореол TITAN',                      '☐', '☐', '☐'],
        ['3',  'Эмпилс',                           '☐', '☐', '☐'],
        ['4',  'краска TITAN по металлу',           '☐', '☐', '☐'],
        ['5',  'грунт-эмаль TITAN',                '☐', '☐', '☐'],
        ['6',  'грунт-эмаль',                      '☐', '☐', '☐'],
        ['7',  'грунт-эмаль по ржавчине',           '☐', '☐', '☐'],
        ['8',  'грунт-эмаль 3 в 1',                 '☐', '☐', '☐'],
        ['9',  'краска для металла',                '☐', '☐', '☐'],
        ['10', 'краска по металлу',                 '☐', '☐', '☐'],
        ['11', 'эмаль по металлу',                  '☐', '☐', '☐'],
        ['12', 'быстросохнущая краска по металлу',  '☐', '☐', '☐'],
        ['13', 'молотковая краска',                 '☐', '☐', '☐'],
        ['14', 'краска по ржавчине',               '☐', '☐', '☐'],
        ['15', 'грунт по металлу',                  '☐', '☐', '☐'],
        ['16', 'краска для ворот',                  '☐', '☐', '☐'],
        ['17', 'краска для забора',                 '☐', '☐', '☐'],
        ['18', 'краска для гаража',                  '☐', '☐', '☐'],
        ['19', 'краска для металлической двери',     '☐', '☐', '☐'],
        ['20', 'покрасить металлический забор',      '☐', '☐', '☐'],
        ['21', 'чем покрасить металл',               '☐', '☐', '☐'],
        ['22', 'Dali краска по металлу',             '☐', '☐', '☐'],
        ['23', 'Салван краска',                      '☐', '☐', '☐'],
        ['24', 'Husky краска',                       '☐', '☐', '☐'],
        ['25', 'Dufa краска по металлу',             '☐', '☐', '☐'],
    ],
    col_widths=[0.8, 7.0, 3.0, 3.0, 3.2],
    font_size=9
)

page_break()

# ============================================================
# ШАГ 4. ЧТО ДЕЛАТЬ ПОСЛЕ СБОРА ДАННЫХ
# ============================================================
add_h1('Шаг 4. Что делать после сбора данных', size=18, before=0)

add_para(
    'После того как все 75 скриншотов сделаны и данные записаны в Excel — '
    'пришлите мне:', after=6
)

add_bullet('Папку со всеми скриншотами (или загрузите в облако и дайте ссылку).', bold=True)
add_bullet('Excel-таблицу с записями (если заполняли).')
add_bullet('Или просто скриншоты — я сам извлеку данные и составлю таблицу.')

add_para('Что я сделаю с этими данными:', bold=True, before=8, after=4)

add_bullet('Сведу все 25 фраз в единую Excel-таблицу с колонками: фраза, '
           'показов в месяц, топ-10 регионов, сезонность.')
add_bullet('Рассчитаю приоритетность каждой фразы для Яндекс Директ (высокий, '
           'средний, низкий).')
add_bullet('Определю топ-10 регионов для запуска рекламы.')
add_bullet('Определю оптимальные месяцы для запуска (сезонность).')
add_bullet('Подготовлю рекомендации по бюджету и ставкам для Яндекс Директ.')
add_bullet('Создам тепловую карту регионов РФ по спросу.')
add_bullet('Сделаю визуализацию: топ-20 фраз, топ-10 регионов, график сезонности.')

page_break()

# ============================================================
# ПРИМЕР — КАК ВЫГЛЯДИТ ОДНА ФРАЗА
# ============================================================
add_h1('Пример — как выглядит заполненная одна фраза', size=18, before=0)

add_para('На примере фразы «грунт-эмаль 3 в 1» — вот что должно получиться:')

add_h2('Действие 1. По словам', before=4, after=4)
add_table(
    headers=['Левая колонка (что искали)', 'Показов', 'Правая колонка (похожие)', 'Показов'],
    rows=[
        ['грунт-эмаль 3 в 1',                 '12 345', 'грунт эмаль 3 в 1 по ржавчине', '8 901'],
        ['грунт-эмаль 3 в 1 по металлу',      '5 678',  'краска 3 в 1 по металлу',       '4 567'],
        ['грунт-эмаль 3 в 1 цена',             '3 456',  'эмаль 3 в 1 для металла',       '2 345'],
        ['купить грунт-эмаль 3 в 1',           '2 345',  'грунт эмаль по ржавчине 3 в 1', '1 890'],
        ['грунт-эмаль 3 в 1 TITAN',            '1 234',  'краска по металлу 3 в 1',       '1 567'],
    ],
    col_widths=[6.5, 2.0, 6.5, 2.0],
    font_size=9
)
add_para('Скриншот: «08_грунт-эмаль_3в1_слова.png»', italic=True, color=GREY)

add_h2('Действие 2. По регионам', before=8, after=4)
add_table(
    headers=['Регион', 'Показов', 'Доля, %'],
    rows=[
        ['Москва и Московская область',  '4 567', '37%'],
        ['Санкт-Петербург',              '1 890', '15%'],
        ['Краснодарский край',            '1 234', '10%'],
        ['Ростовская область',            '987',   '8%'],
        ['Свердловская область',          '654',   '5%'],
        ['Республика Татарстан',          '543',   '4%'],
        ['Новосибирская область',         '432',   '4%'],
        ['Нижегородская область',          '321',   '3%'],
        ['Самарская область',              '287',   '2%'],
        ['Челябинская область',            '234',   '2%'],
    ],
    col_widths=[7.0, 4.0, 4.0],
    font_size=10
)
add_para('Скриншот: «08_грунт-эмаль_3в1_регионы.png»', italic=True, color=GREY)

add_h2('Действие 3. История запросов', before=8, after=4)
add_table(
    headers=['Месяц', 'Показов', 'Месяц', 'Показов'],
    rows=[
        ['Январь 2026',  '3 456', 'Июль 2026',    '15 678'],
        ['Февраль 2026', '4 567', 'Август 2026',  '12 345'],
        ['Март 2026',    '7 890', 'Сентябрь 2026', '8 901'],
        ['Апрель 2026',  '12 345', 'Октябрь 2026', '5 678'],
        ['Май 2026',     '14 567', 'Ноябрь 2026',  '3 456'],
        ['Июнь 2026',    '15 678', 'Декабрь 2026', '2 345'],
    ],
    col_widths=[4.0, 3.5, 4.0, 3.5],
    font_size=10
)
add_para('Скриншот: «08_грунт-эмаль_3в1_история.png»', italic=True, color=GREY)

add_alert_box('Вывод по примеру',
    'Фраза «грунт-эмаль 3 в 1» — 12 345 показов в месяц. Пик — май–июль '
    '(сезон покраски). Топ-3 региона: Москва (37%), СПб (15%), Краснодар (10%). '
    'Рекомендация для Директа: запуск с апреля по август, гео — Москва + СПб + '
    'Краснодар + Ростов. Это тот вывод, который я сделаю по каждой фразе после '
    'получения ваших скриншотов.',
    color='green')

page_break()

# ============================================================
# ПОДСКАЗКИ И СОВЕТЫ
# ============================================================
add_h1('Подсказки и советы', size=18, before=0)

add_h2('1. Как сделать скриншот', before=4, after=4)
add_bullet('Windows: Win + Shift + S — выделите область экранка, вставьте в Paint и сохраните.')
add_bullet('Или: нажмите Print Screen (PrtScn) — весь экран, вставьте в Paint, обрежьте лишнее.')
add_bullet('Mac: Cmd + Shift + 4 — выделите область, файл сохранится на рабочий стол.')

add_h2('2. Что должно быть видно на скриншоте', before=8, after=4)
add_bullet('Строка поиска с введённой фразой (чтобы я видел, что именно искали).')
add_bullet('Числа показов (важны цифры, а не только фразы).')
add_bullet('Для «По регионам» — карта + таблица с топ-10 регионами.')
add_bullet('Для «Истории» — график с пиками и спадами.')

add_h2('3. Если фраза не даёт результатов', before=8, after=4)
add_bullet('Если показов меньше 50 — всё равно сделайте скриншот. Это важная информация: фраза не работает.')
add_bullet('Если Wordstat пишет «по вашему запросу ничего не найдено» — сделайте скриншот этого сообщения.')
add_bullet('Если фраза слишком широкая (например, «краска» — 500 000+ показов) — сделайте скриншот, но в Excel отметьте «слишком широкая — сузить».')

add_h2('4. Дополнительные фразы — по желанию', before=8, after=4)
add_para('Если по ходу работы в правой колонке Wordstat увидите интересные фразы, '
         'которых нет в списке — добавьте их в Excel и сделайте по ним те же 3 действия. '
         'Особенно ценные фразы с упоминанием брендов конкурентов или конкретных объектов '
         '(например, «покрасить ворота на даче»).')

add_h2('5. Время выполнения', before=8, after=4)
add_table(
    headers=['Действие', 'Время на 1 фразу', 'Время на 25 фраз'],
    rows=[
        ['Ввод фразы + «Подобрать»',           '30 сек',  '12 мин'],
        ['Просмотр + запись левой колонки',    '1 мин',    '25 мин'],
        ['Просмотр + запись правой колонки',   '30 сек',  '12 мин'],
        ['Переключение на «По регионам» + запись', '1,5 мин', '37 мин'],
        ['Переключение на «История» + запись', '1 мин',    '25 мин'],
        ['Скриншоты (3 шт. на фразу)',          '1 мин',    '25 мин'],
        ['ИТОГО',                                '~5 мин',  '~2 часа 20 мин'],
    ],
    col_widths=[7.0, 4.0, 4.0],
    font_size=10
)

add_para('Можно разбить на 2 дня — по 12–13 фраз (около 1 часа в день).', italic=True, color=GREY)

# ---------- Финал ----------
hr()
add_para('Wordstat — план исследования  ·  TITAN  ·  Сентябрь 2026',
         size=10, italic=True, color=GREY, align='center', after=2)
add_para('После сбора данных пришлите скриншоты — я сведу в Excel и подготовлю '
         'рекомендации по Яндекс Директ.',
         size=10, italic=True, color=GREY, align='center')

# ---------- Сохранение ----------
for table in doc.tables:
    if len(table.rows) > 0:
        tr = table.rows[0]._tr
        trPr = tr.get_or_add_trPr()
        tblHeader = OxmlElement('w:tblHeader')
        tblHeader.set(qn('w:val'), 'true')
        trPr.append(tblHeader)
        for row in table.rows:
            trPr = row._tr.get_or_add_trPr()
            cantSplit = OxmlElement('w:cantSplit')
            cantSplit.set(qn('w:val'), 'true')
            trPr.append(cantSplit)

out = '/home/z/my-project/download/Wordstat_план_исследования_TITAN.docx'
doc.save(out)
print(f'OK saved: {out}')
import os
print(f'Size: {os.path.getsize(out)} bytes')
