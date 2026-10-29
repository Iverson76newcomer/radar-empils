"""Обновление Excel и PPTX — добавление бренда "Расцвет" в рейтинг.
   Excel: вкладка 2 (Рейтинг брендов), вкладка 5 (Детализация), вкладка 6 (Матрица)
   PPTX: слайд 3 (Рейтинг узнаваемости брендов грунт-эмалей)
"""
import json
import shutil
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from copy import copy

# Загружаем данные Wordstat
with open('/home/z/my-project/scripts/rascvet_wordstat.json', 'r', encoding='utf-8') as f:
    rascvet_data = json.load(f)

queries_data = rascvet_data['queries']
total_rascvet = rascvet_data['total_7_requests']

# Данные всех брендов для пересчёта
brands_data = [
    {'brand': 'Dali',       'sum': 52968, 'color': 'краска',    'counts': {'краска': 19626, 'грунт-эмаль': 9977, '3 в 1': 7812, 'по металлу': 1766, 'по ржавчине': 10034, 'молотковая': 3486, 'быстросохнущая': 267}},
    {'brand': 'Престиж',    'sum': 31361, 'color': 'краска',    'counts': {'краска': 17998, 'грунт-эмаль': 2762, '3 в 1': 3102, 'по металлу': 2122, 'по ржавчине': 1708, 'молотковая': 1754, 'быстросохнущая': 1915}},
    {'brand': 'Elcon',      'sum': 19861, 'color': 'краска',   'counts': {'краска': 6584, 'грунт-эмаль': 4090, '3 в 1': 2876, 'по металлу': 2127, 'по ржавчине': 2675, 'молотковая': 266, 'быстросохнущая': 1243}},
    {'brand': 'Лакра',      'sum': 18061, 'color': 'краска',   'counts': {'краска': 12027, 'грунт-эмаль': 2138, '3 в 1': 954, 'по металлу': 658, 'по ржавчине': 1533, 'молотковая': 320, 'быстросохнущая': 431}},
    {'brand': 'Husky',      'sum': 5656,  'color': 'краска',   'counts': {'краска': 4372, 'грунт-эмаль': 54, '3 в 1': 52, 'по металлу': 705, 'по ржавчине': 149, 'молотковая': 297, 'быстросохнущая': 27}},
    {'brand': 'Расцвет',    'sum': total_rascvet, 'color': 'краска',
     'counts': {k: v['count'] for k, v in queries_data.items()}},  # ← НОВЫЙ БРЕНД
    {'brand': 'Новбытхим',  'sum': 4491,  'color': 'краска',   'counts': {'краска': 1596, 'грунт-эмаль': 1007, '3 в 1': 811, 'по металлу': 120, 'по ржавчине': 919, 'молотковая': 12, 'быстросохнущая': 26}},
    {'brand': 'Metallista', 'sum': 2386,  'color': 'краска',   'counts': {'краска': 1009, 'грунт-эмаль': 66, '3 в 1': 328, 'по металлу': 271, 'по ржавчине': 636, 'молотковая': 76, 'быстросохнущая': 0}},
    {'brand': 'TITAN',      'sum': 1430,  'color': 'краска',   'counts': {'краска': 424, 'грунт-эмаль': 310, '3 в 1': 329, 'по металлу': 199, 'по ржавчине': 93, 'молотковая': 56, 'быстросохнущая': 19}},
    {'brand': 'Correx',     'sum': 1185,  'color': 'краска',   'counts': {'краска': 145, 'грунт-эмаль': 337, '3 в 1': 193, 'по металлу': 152, 'по ржавчине': 320, 'молотковая': 29, 'быстросохнущая': 9}},
]

# Сортировка по убыванию суммы
brands_data.sort(key=lambda x: -x['sum'])

# TITAN - базовый品牌 для расчёта "во сколько раз >"
titan_sum = next(b['sum'] for b in brands_data if b['brand'] == 'TITAN')

# Копируем файлы для правки
shutil.copy('/home/z/my-project/upload/Wordstat_узнаваемость_брендов.xlsx',
            '/home/z/my-project/download/Wordstat_узнаваемость_брендов.xlsx')

wb = load_workbook('/home/z/my-project/download/Wordstat_узнаваемость_брендов.xlsx')

# Стили
THIN_BORDER = Border(left=Side(style='thin', color='D0D0D0'),
                     right=Side(style='thin', color='D0D0D0'),
                     top=Side(style='thin', color='D0D0D0'),
                     bottom=Side(style='thin', color='D0D0D0'))
NORMAL_FONT = Font(name='Arial', size=10)
BOLD_FONT = Font(name='Arial', size=10, bold=True)
RASCВET_FILL = PatternFill('solid', fgColor='FFFDE7')  # светло-жёлтый — помечаем новую строку

# ============================================================
# ЛИСТ 2: Рейтинг брендов — обновляем с Расцвет
# ============================================================
ws2 = wb['2. Рейтинг брендов']

# Стираем старые данные (строки 4..12)
for ri in range(4, 14):
    for ci in range(1, 6):
        c = ws2.cell(row=ri, column=ci)
        c.value = None
        c.fill = PatternFill(fill_type=None)
        c.font = NORMAL_FONT
        c.border = THIN_BORDER
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

# Заголовок оставляем — пишем данные
for ri, b in enumerate(brands_data, 1):
    ratio = b['sum'] / titan_sum if titan_sum else 0
    if ratio >= 10:
        ratio_str = f"{ratio:.0f}x"
    else:
        ratio_str = f"{ratio:.1f}x"
    is_new = b['brand'] == 'Расцвет'
    # № | Бренд | Сумма 7 запросов | Во сколько раз > TITAN | Статус
    row_data = [ri, b['brand'], b['sum'], ratio_str, '✅ Полные (7/7)']
    for ci, v in enumerate(row_data, 1):
        c = ws2.cell(row=3 + ri, column=ci, value=v)
        c.font = BOLD_FONT if is_new else NORMAL_FONT
        c.border = THIN_BORDER
        c.alignment = Alignment(horizontal='center', vertical='center')
        if is_new:
            c.fill = RASCВET_FILL

# Обновим заголовок листа — добавим примечание про новый бренд
title_cell = ws2.cell(row=1, column=1)
title_cell.value = 'РЕЙТИНГ УЗНАВАЕМОСТИ БРЕНДОВ ГРУНТ-ЭМАЛЕЙ — ОБЪЕКТИВНОЕ СРАВНЕНИЕ'

# Обновим подзаголовок (строка 2) с упоминанием Расцвет
subtitle_cell = ws2.cell(row=2, column=1)
subtitle_cell.value = f'Метод: одинаковый пул из 7 запросов для каждого бренда | Yandex Wordstat API, сентябрь 2026 | Обновлено: добавлен "Расцвет" ({total_rascvet} показов)'

# Ширины
widths = {'A': 5, 'B': 16, 'C': 18, 'D': 22, 'E': 22}
for col, w in widths.items():
    ws2.column_dimensions[col].width = w

# ============================================================
# ЛИСТ 5: Детализация по запросам — добавляем 7 строк для Расцвет
# ============================================================
ws5 = wb['5. Детализация по запросам']
# Найдём последнюю заполненную строку
last_row = ws5.max_row
# Добавляем после последней
new_row_start = last_row + 2  # пропускаем 1 пустую

# Заголовок для Расцвет
c = ws5.cell(row=new_row_start, column=1, value='Расцвет')
c.font = Font(name='Arial', size=10, bold=True, color='1E3A5F')
c.fill = RASCВET_FILL
ws5.merge_cells(start_row=new_row_start, start_column=1, end_row=new_row_start, end_column=4)
c.alignment = Alignment(horizontal='left', vertical='center')

# 7 строк с запросами Расцвет
queries_order = [
    ('Расцвет', 'краска',         'Расцвет краска'),
    ('Расцвет', 'грунт-эмаль',    'Расцвет грунт-эмаль'),
    ('Расцвет', '3 в 1',          'Расцвет 3 в 1'),
    ('Расцвет', 'по металлу',     'Расцвет по металлу'),
    ('Расцвет', 'по ржавчине',    'Расцвет по ржавчине'),
    ('Расцвет', 'молотковая',     'Расцвет молотковая'),
    ('Расцвет', 'быстросохнущая', 'Расцвет быстросохнущая'),
]
for i, (brand, key, query) in enumerate(queries_order, 1):
    cnt = queries_data[key]['count']
    ri = new_row_start + i
    row_data = [brand, key, query, cnt]
    for ci, v in enumerate(row_data, 1):
        c = ws5.cell(row=ri, column=ci, value=v)
        c.font = NORMAL_FONT
        c.border = THIN_BORDER
        c.alignment = Alignment(horizontal='center', vertical='center')
        c.fill = RASCВET_FILL

# ============================================================
# ЛИСТ 6: Матрица бренд × запрос — добавляем строку "Расцвет"
# ============================================================
ws6 = wb['6. Матрица бренд × запрос']
# Найдём строку с TITAN (для вставки после нее перед Correx)
# Лучшая стратегия: вставить "Расцвет" в правильную позицию в порядке убывания суммы

# Найдём последнюю строку с данными
last_data_row = ws6.max_row

# Текущий порядок брендов в матрице — отсортированный по сумме
# Читаем текущие строки
current_rows = []
for ri in range(3, last_data_row + 1):
    brand = ws6.cell(row=ri, column=1).value
    if brand and brand != 'ИТОГО':
        current_rows.append((ri, brand))

print(f"Текущие бренды в матрице: {[b for _, b in current_rows]}")

# Определим позицию для Расцвет (между Husky и Новбытхим — позиция 6)
# Чтобы не менять порядок строк, найдём позицию и вставим новую строку
# Простой подход: добавим в конец, если порядок не критичен

# Найдём позицию: Расцвет (3884) должна быть между Husky (5656) и Новбытхим (4491)
# Найдём строку Новбытхим
target_row = None
for ri, brand in current_rows:
    if brand == 'Новбытхим':
        target_row = ri
        break

if target_row:
    ws6.insert_rows(target_row)
    # Заполняем новую строку
    headers = ['Бренд', 'краска', 'грунт-эмаль', '3 в 1', 'по металлу', 'по ржавчине', 'молотковая', 'быстросохнущая', 'ИТОГО']
    rascvet_row_data = ['Расцвет']
    for h in headers[1:-1]:  # кроме Бренд и ИТОГО
        rascvet_row_data.append(queries_data[h]['count'])
    rascvet_row_data.append(total_rascvet)

    for ci, v in enumerate(rascvet_row_data, 1):
        c = ws6.cell(row=target_row, column=ci, value=v)
        c.font = BOLD_FONT
        c.border = THIN_BORDER
        c.alignment = Alignment(horizontal='center', vertical='center')
        c.fill = RASCВET_FILL

# Сохраняем Excel
wb.save('/home/z/my-project/download/Wordstat_узнаваемость_брендов.xlsx')
print("✅ Excel сохранён: /home/z/my-project/download/Wordstat_узнаваемость_брендов.xlsx")

# ============================================================
# PPTX: Слайд 3 — Рейтинг узнаваемости брендов грунт-эмалей
# ============================================================
print("\nОбновляю PPTX...")
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

shutil.copy('/home/z/my-project/upload/Wordstat_полный_анализ_TITAN_ с компанией.pptx',
            '/home/z/my-project/download/Wordstat_полный_анализ_TITAN_ с компанией.pptx')

prs = Presentation('/home/z/my-project/download/Wordstat_полный_анализ_TITAN_ с компанией.pptx')

# Слайд 3 — Рейтинг брендов (по 0-индексу)
slide = prs.slides[2]

# Найдём таблицу в этом слайде
target_table = None
for shape in slide.shapes:
    if shape.has_table:
        target_table = shape.table
        target_shape = shape
        break

if target_table:
    print(f"Найдена таблица: {len(target_table.rows)} строк × {len(target_table.columns)} колонок")
    # Структура таблицы:
    # Строка 0: заголовок (№ | Бренд | Сумма 7 запросов | Во сколько раз > TITAN)
    # Строки 1-9: данные (Dali, Престиж, Elcon, Лакра, Husky, Новбытхим, Metallista, TITAN, Correx)

    # Нужно вставить строку "Расцвет" после Husky (позиция 5, т.е. после строки 5)
    # Тогда порядок: Dali, Престиж, Elcon, Лакра, Husky, Расцвет, Новбытхим, Metallista, TITAN, Correx

    # python-pptx: добавление строк — через XML
    from pptx.oxml.ns import qn
    from copy import deepcopy

    tbl = target_table._tbl
    # Найдём строку Новбытхим (после которой нужно вставить Расцвет)
    # Точнее — найдём Husky (перед Новбытхим), и вставим после неё

    rows = list(tbl.iterchildren(qn('a:tr')))
    print(f"XML rows count: {len(rows)}")

    # Найдём индекс строки Husky
    husky_idx = None
    for i, row in enumerate(rows):
        cells = list(row.iterchildren(qn('a:tc')))
        if cells and len(cells) > 1:
            cell_text = cells[1].text if cells[1].text else ''
            # Извлекаем текст
            text = ''
            for t in cells[1].iter(qn('a:t')):
                text += t.text or ''
            if 'Husky' in text:
                husky_idx = i
                print(f"Husky найден в строке {i}")
                break

    if husky_idx is not None:
        # Клонируем строку Husky как шаблон
        new_row = deepcopy(rows[husky_idx])

        # Очищаем и заполняем новыми данными
        cells = list(new_row.iterchildren(qn('a:tc')))
        new_data = ['6', 'Расцвет', str(total_rascvet), f'{total_rascvet/1430:.1f}x']
        for ci, cell in enumerate(cells[:4]):
            # Удаляем существующий текст
            for p in cell.iter(qn('a:p')):
                for r in list(p.iter(qn('a:r'))):
                    p.remove(r)
            # Добавляем новый параграф с текстом
            # Создаём новый run
            from lxml import etree
            NSMAP = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}
            # Получаем первый параграф
            paras = list(cell.iter(qn('a:p')))
            if paras:
                p = paras[0]
                # Создаём run с правильным форматированием
                rPr_xml = f'''<a:r xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                    <a:rPr lang="ru-RU" sz="1400" b="1"/>
                    <a:t>{new_data[ci]}</a:t>
                </a:r>'''
                r = etree.fromstring(rPr_xml)
                p.append(r)
                # Удаляем лишние параграфы
                for extra_p in paras[1:]:
                    cell.remove(extra_p)

        # Вставляем новую строку после Husky
        husky_row_xml = rows[husky_idx]
        husky_row_xml.addnext(new_row)

        # Обновляем нумерацию последующих строк
        all_rows = list(tbl.iterchildren(qn('a:tr')))
        for i in range(husky_idx + 2, len(all_rows)):  # +2 = новая строка + следующая
            cells = list(all_rows[i].iterchildren(qn('a:tc')))
            if cells and len(cells) > 0:
                # Обновляем номер в первой ячейке
                new_num = str(i)  # i = позиция в массиве (1-индекс: header=0, первая строка=1)
                paras = list(cells[0].iter(qn('a:p')))
                if paras:
                    # Удаляем существующие runs
                    for r in list(paras[0].iter(qn('a:r'))):
                        paras[0].remove(r)
                    # Добавляем новый
                    from lxml import etree
                    rPr_xml = f'''<a:r xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                        <a:rPr lang="ru-RU" sz="1400"/>
                        <a:t>{new_num}</a:t>
                    </a:r>'''
                    r = etree.fromstring(rPr_xml)
                    paras[0].append(r)

    # Также обновим заголовок слайда — добавим "(обновлено)"
    # Найдём TextBox с заголовком
    for shape in slide.shapes:
        if shape.has_text_frame:
            for para in shape.text_frame.paragraphs:
                if 'Рейтинг узнаваемости брендов грунт-эмалей' in para.text:
                    # Изменим текст параграфа
                    for run in para.runs:
                        if 'Сумма 7' in run.text or 'сентябрь 2026' in run.text:
                            run.text = run.text.replace('сентябрь 2026', f'сентябрь 2026 (обновлено: добавлен «Расцвет»)')

# Сохраняем PPTX
prs.save('/home/z/my-project/download/Wordstat_полный_анализ_TITAN_ с компанией.pptx')
print("✅ PPTX сохранён: /home/z/my-project/download/Wordstat_полный_анализ_TITAN_ с компанией.pptx")

# Проверка
print("\n=== ПРОВЕРКА ===")
print(f"Расцвет — сумма 7 запросов: {total_rascvet}")
print(f"Сравнение с TITAN: {total_rascvet/1430:.2f}x (TITAN = 1430)")
print(f"Сравнение с Correx: {total_rascvet/1185:.2f}x (Correx = 1185)")
print(f"Позиция в рейтинге: 6 из 10 (после Husky, перед Новбытхим)")
