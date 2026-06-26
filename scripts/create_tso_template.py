#!/usr/bin/env python3
"""
Создание Excel-шаблона для ввода ответов бумажных анкет ТСО.
Каждая строка = один респондент. Выпадающие списки для закрытых вопросов.
"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, numbers
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from copy import copy

wb = openpyxl.Workbook()

# ─── Цветовая палитра ───
DARK_BLUE = "1B3A5C"
MED_BLUE = "2E75B6"
LIGHT_BLUE = "D6E4F0"
ACCENT_ORANGE = "ED7D31"
WHITE = "FFFFFF"
LIGHT_GRAY = "F2F2F2"
HEADER_FONT_COLOR = "FFFFFF"

# ─── Стили ───
header_font = Font(name="Calibri", bold=True, color=HEADER_FONT_COLOR, size=11)
header_fill = PatternFill(start_color=DARK_BLUE, end_color=DARK_BLUE, fill_type="solid")
subheader_fill = PatternFill(start_color=MED_BLUE, end_color=MED_BLUE, fill_type="solid")
subheader_font = Font(name="Calibri", bold=True, color=HEADER_FONT_COLOR, size=10)
data_font = Font(name="Calibri", size=10)
block_fill = PatternFill(start_color=LIGHT_BLUE, end_color=LIGHT_BLUE, fill_type="solid")
alt_fill = PatternFill(start_color=LIGHT_GRAY, end_color=LIGHT_GRAY, fill_type="solid")
orange_fill = PatternFill(start_color=ACCENT_ORANGE, end_color=ACCENT_ORANGE, fill_type="solid")
thin_border = Border(
    left=Side(style="thin", color="BFBFBF"),
    right=Side(style="thin", color="BFBFBF"),
    top=Side(style="thin", color="BFBFBF"),
    bottom=Side(style="thin", color="BFBFBF"),
)
center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)

# ═══════════════════════════════════════════════════════
# ЛИСТ 1: Ввод данных
# ═══════════════════════════════════════════════════════
ws = wb.active
ws.title = "Ввод данных"

# ─── Определение колонок ───
columns = [
    # (id, заголовок, тип, ширина, варианты ответа)
    ("id", "№ анкеты", "auto", 8, []),
    ("date", "Дата заполнения", "date", 14, []),
    
    # БЛОК 1
    ("q1", "В1. Должность", "single", 30, [
        "Руководитель / владелец компании",
        "Руководитель отдела закупок / категорийный менеджер",
        "Менеджер по продажам / торговый представитель",
        "Продавец / консультант в зале",
        "Другое",
    ]),
    ("q1_other", "В1. Другое (уточните)", "text", 25, []),
    ("q2", "В1. Опыт работы в ЛКМ", "single", 20, [
        "Менее 1 года",
        "1–3 года",
        "3–5 лет",
        "Более 5 лет",
    ]),
    ("q3", "В1. Формат продаж", "single", 30, [
        "Преимущественно розничные продажи (сеть магазинов / DIY)",
        "Преимущественно оптовые продажи (дистрибуция)",
        "Смешанный формат (опт + розница)",
        "Продажи проектным / корпоративным клиентам",
    ]),
    
    # БЛОК 2
    ("q4", "В2. Знакомство с ТСО", "single", 35, [
        "Да, хорошо понимаю и использую при продажах",
        "Слышал(а) термин, но не применял(а) на практике",
        "Не знаком(а) с термином, но интуитивно считаю полную стоимость",
        "Впервые слышу — всегда ориентируюсь только на цену банки",
    ]),
    ("q5_1", "В2.5 Цена банки", "multi", 10, []),
    ("q5_2", "В2.5 Грунтовка/подготовка", "multi", 10, []),
    ("q5_3", "В2.5 Количество слоёв", "multi", 10, []),
    ("q5_4", "В2.5 Время высыхания", "multi", 10, []),
    ("q5_5", "В2.5 Стоимость рабочей силы", "multi", 10, []),
    ("q5_6", "В2.5 Срок службы покрытия", "multi", 10, []),
    ("q5_7", "В2.5 Расход на инструмент", "multi", 10, []),
    ("q5_8", "В2.5 Затраты на утилизацию", "multi", 10, []),
    ("q6", "В2. Пример ТСО из практики", "text", 40, []),
    
    # БЛОК 3
    ("q7", "В3. Как рекомендуете ЛКМ", "single", 35, [
        "Называю цену и основные характеристики",
        "Сравниваю несколько вариантов по цене и качеству",
        "Считаю вместе с покупателем полную стоимость (грунт + краска + работа)",
        "Предлагаю несколько вариантов и оставляю выбор покупателю",
        "Другое",
    ]),
    ("q7_other", "В3. Другое (уточните)", "text", 25, []),
    ("q8", "В3. Лучший аргумент для премиума", "single", 35, [
        "«Это качественнее / долговечнее»",
        "«Экономит время — меньше слоёв, быстрее работа»",
        "«Полная стоимость покраски в итоге дешевле»",
        "«Этим пользуются профессионалы»",
        "«Не придётся перекрашивать через год»",
        "«Укрывистость лучше — меньше расход»",
        "Другое",
    ]),
    ("q8_other", "В3. Другое (уточните)", "text", 25, []),
    ("q9_1", "В3.9 «Чем дороже обычной?»", "multi", 10, []),
    ("q9_2", "В3.9 «Нет подешевле?»", "multi", 10, []),
    ("q9_3", "В3.9 «Нужно покрасить недорого»", "multi", 10, []),
    ("q9_4", "В3.9 «Покрашу сам — не за дорогую»", "multi", 10, []),
    ("q9_5", "В3.9 «Давайте подешевле — одинаково»", "multi", 10, []),
    ("q9_6", "В3.9 Другое (фраза)", "multi", 10, []),
    
    # БЛОК 4
    ("q10", "В4. ТСО повлияет на решение (0-10)", "scale", 12, []),
    ("q11", "В4. Кто восприимчив к ТСО", "single", 35, [
        "Дачник / хозяин дома — красит сам, ценит своё время",
        "Профессиональный маляр / строитель — считает рабочую силу",
        "Покупатель эконом-сегмента — всегда выбирает по цене",
        "Корпоративный заказчик — считает бюджет проекта",
        "Все одинаково — неважно, кто покупает",
    ]),
    ("q12_1", "В4.12 Грунт-эмали 3в1", "multi", 10, []),
    ("q12_2", "В4.12 Премиум-эмали (замена ПФ-115)", "multi", 10, []),
    ("q12_3", "В4.12 ВД краски повыш. стойкости", "multi", 10, []),
    ("q12_4", "В4.12 Фасадные с повыш. укрывист.", "multi", 10, []),
    ("q12_5", "В4.12 Лаки/пропитки премиум", "multi", 10, []),
    ("q12_6", "В4.12 Другое", "multi", 10, []),
    
    # БЛОК 5
    ("q13_1", "В5.13 Нет наглядного расчёта", "multi", 10, []),
    ("q13_2", "В5.13 Не хватает знаний", "multi", 10, []),
    ("q13_3", "В5.13 Покупателю не интересны расчёты", "multi", 10, []),
    ("q13_4", "В5.13 Нет времени на объяснения", "multi", 10, []),
    ("q13_5", "В5.13 Не верю, что заплатит больше", "multi", 10, []),
    ("q13_6", "В5.13 Конкуренты дешевле", "multi", 10, []),
    ("q13_7", "В5.13 Другое", "multi", 10, []),
    ("q14_1", "В5.14 Калькулятор ТСО", "multi", 10, []),
    ("q14_2", "В5.14 Печатная шпаргалка", "multi", 10, []),
    ("q14_3", "В5.14 Образцы покрытий", "multi", 10, []),
    ("q14_4", "В5.14 Обучение: ТСО за 2 мин", "multi", 10, []),
    ("q14_5", "В5.14 Видео/презентация", "multi", 10, []),
    ("q14_6", "В5.14 Сравнит. таблица vs конкуренты", "multi", 10, []),
    ("q14_7", "В5.14 Мотивационная программа", "multi", 10, []),
    ("q14_8", "В5.14 Другое", "multi", 10, []),
    ("q15", "В5. Срок внедрения инструмента", "single", 30, [
        "Сразу же — это то, чего нам не хватало",
        "В течение 1–2 месяцев после обучения",
        "Через полгода — нужно время на привыкание",
        "Скорее не внедрил бы — не уверен в эффективности",
    ]),
    
    # БЛОК 6
    ("q16_1", "В6.16 Готовые материалы", "multi", 10, []),
    ("q16_2", "В6.16 Обучение персонала ТСО", "multi", 10, []),
    ("q16_3", "В6.16 Политика ценообразования", "multi", 10, []),
    ("q16_4", "В6.16 Маркетинговая поддержка", "multi", 10, []),
    ("q16_5", "В6.16 Сопровождение на месте", "multi", 10, []),
    ("q16_6", "В6.16 Другое", "multi", 10, []),
    ("q17", "В6. Перспективность ТСО (0-10)", "scale", 12, []),
    
    ("notes", "Примечания", "text", 30, []),
]

# ─── Запись заголовков ───
block_names = {
    2: "БЛОК 1. Ваша роль и опыт",
    6: "БЛОК 2. Понимание методики ТСО",
    15: "БЛОК 3. Текущий подход к продаже ЛКМ",
    24: "БЛОК 4. ТСО как инструмент продажи",
    33: "БЛОК 5. Барьеры и потребности",
    48: "БЛОК 6. Ожидания от производителя",
}

for col_idx, (col_id, header, col_type, width, options) in enumerate(columns, 1):
    cell = ws.cell(row=1, column=col_idx, value=header)
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = center_align
    cell.border = thin_border
    ws.column_dimensions[get_column_letter(col_idx)].width = width

# ─── Блоковые подзаголовки (строка 2) — группировка по блокам ───
block_ranges = [
    (1, 5, "БЛОК 1: Роль и опыт"),
    (6, 14, "БЛОК 2: Понимание ТСО"),
    (15, 23, "БЛОК 3: Текущий подход"),
    (24, 31, "БЛОК 4: ТСО как инструмент"),
    (32, 49, "БЛОК 5: Барьеры и потребности"),
    (50, 57, "БЛОК 6: Ожидания"),
]
for start, end, label in block_ranges:
    ws.merge_cells(start_row=2, start_column=start, end_row=2, end_column=end)
    cell = ws.cell(row=2, column=start, value=label)
    cell.font = subheader_font
    cell.fill = subheader_fill
    cell.alignment = center_align
    cell.border = thin_border
    for c in range(start, end + 1):
        ws.cell(row=2, column=c).fill = subheader_fill
        ws.cell(row=2, column=c).border = thin_border

# ─── Подсказки (строка 3) ───
hints = {
    "id": "Авто",
    "date": "ДД.ММ.ГГГГ",
    "q1_other": "Если Q1=Другое",
    "q7_other": "Если Q7=Другое",
    "q8_other": "Если Q8=Другое",
    "q10": "0–10",
    "q17": "0–10",
}
hint_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
hint_font = Font(name="Calibri", italic=True, size=9, color="806000")

for col_idx, (col_id, header, col_type, width, options) in enumerate(columns, 1):
    hint_text = hints.get(col_id, "")
    if col_type == "multi":
        hint_text = "1=Да, 0=Нет"
    elif col_type == "single" and options:
        hint_text = "Выберите из списка"
    cell = ws.cell(row=3, column=col_idx, value=hint_text)
    cell.font = hint_font
    cell.fill = hint_fill
    cell.alignment = center_align
    cell.border = thin_border

# ─── Формат строк данных (50 строк для ввода) ───
for row in range(4, 54):
    for col_idx, (col_id, header, col_type, width, options) in enumerate(columns, 1):
        cell = ws.cell(row=row, column=col_idx)
        cell.font = data_font
        cell.border = thin_border
        cell.alignment = left_align if col_type == "text" else center_align
        # Чередование цветов
        if row % 2 == 0:
            cell.fill = alt_fill
        # Автонумерация
        if col_id == "id":
            cell.value = row - 3
        # Дата
        if col_id == "date":
            cell.number_format = "DD.MM.YYYY"

# ─── Выпадающие списки (Data Validation) ───
for col_idx, (col_id, header, col_type, width, options) in enumerate(columns, 1):
    if col_type == "single" and options:
        col_letter = get_column_letter(col_idx)
        # Формируем список через формулу (лист Справочники)
        dv = DataValidation(
            type="list",
            formula1=f"Справочники!$A${col_idx}",
            allow_blank=True,
        )
        dv.error = "Выберите значение из списка"
        dv.errorTitle = "Ошибка ввода"
        dv.prompt = "Нажмите для выбора"
        dv.promptTitle = header
        dv.showErrorMessage = True
        dv.showInputMessage = True
        dv.add(f"{col_letter}4:{col_letter}53")
        ws.add_data_validation(dv)
    elif col_type == "scale":
        col_letter = get_column_letter(col_idx)
        dv = DataValidation(
            type="whole",
            operator="between",
            formula1=0,
            formula2=10,
            allow_blank=True,
        )
        dv.error = "Введите число от 0 до 10"
        dv.errorTitle = "Ошибка"
        dv.showErrorMessage = True
        dv.add(f"{col_letter}4:{col_letter}53")
        ws.add_data_validation(dv)
    elif col_type == "multi":
        col_letter = get_column_letter(col_idx)
        dv = DataValidation(
            type="whole",
            operator="between",
            formula1=0,
            formula2=1,
            allow_blank=True,
        )
        dv.error = "Введите 1 (отмечено) или 0 (нет)"
        dv.errorTitle = "Ошибка"
        dv.showErrorMessage = True
        dv.add(f"{col_letter}4:{col_letter}53")
        ws.add_data_validation(dv)

# ─── Закрепление областей ───
ws.freeze_panes = "C4"

# ─── Автофильтр ───
ws.auto_filter.ref = f"A1:{get_column_letter(len(columns))}53"


# ═══════════════════════════════════════════════════════
# ЛИСТ 2: Справочники (варианты ответов для выпадающих списков)
# ═══════════════════════════════════════════════════════
ws_ref = wb.create_sheet("Справочники")

for col_idx, (col_id, header, col_type, width, options) in enumerate(columns, 1):
    if col_type == "single" and options:
        for opt_row, option in enumerate(options, 1):
            ws_ref.cell(row=opt_row, column=col_idx, value=option)

# Скрыть справочный лист
ws_ref.sheet_state = "hidden"


# ═══════════════════════════════════════════════════════
# ЛИСТ 3: Инструкция
# ═══════════════════════════════════════════════════════
ws_instr = wb.create_sheet("Инструкция")

instructions = [
    ("ИНСТРУКЦИЯ ПО ВВОДУ ДАННЫХ АНКЕТЫ ТСО", None),
    ("", None),
    ("1. КАК ПОЛЬЗОВАТЬСЯ ШАБЛОНОМ", "header"),
    ("Каждая заполненная бумажная анкета = одна строка в таблице (начиная со строки 4).", None),
    ("Столбцы соответствуют вопросам анкеты. Для закрытых вопросов — выпадающие списки.", None),
    ("Для вопросов с множественным выбором (☐) — каждый вариант в отдельном столбце: 1 = отмечено, 0 = нет.", None),
    ("Для шкальных вопросов (0–10) — введите число от 0 до 10.", None),
    ("Для открытых вопросов — вводите текст ответа как есть.", None),
    ("", None),
    ("2. ПОСЛЕДОВАТЕЛЬНОСТЬ ДЕЙСТВИЙ", "header"),
    ("Шаг 1: Возьмите стопку заполненных бумажных анкет.", None),
    ("Шаг 2: Присвойте каждой анкете номер (1, 2, 3...).", None),
    ("Шаг 3: Внесите ответы в соответствующую строку таблицы.", None),
    ("Шаг 4: Для вопросов «Другое» — вводите текст в столбец «Другое (уточните)».", None),
    ("Шаг 5: После ввода всех анкет — сохраните файл и запустите скрипт анализа.", None),
    ("", None),
    ("3. СПОСОБЫ ОЦИФРОВКИ БУМАЖНЫХ АНКЕТ", "header"),
    ("СПОСОБ А — Ручной ввод (рекомендуется для ≤30 анкет):", None),
    ("  • Откройте этот файл и вносите ответы вручную с каждой бумажной анкеты.", None),
    ("  • Выпадающие списки ускоряют ввод и исключают опечатки.", None),
    ("", None),
    ("СПОСОБ Б — OCR-распознавание (для 30+ анкет):", None),
    ("  • Сфотографируйте анкеты и распознайте через ABBYY FineReader или Яндекс.Формы.", None),
    ("  • Затем перенесите распознанные данные в этот шаблон.", None),
    ("", None),
    ("СПОСОБ В — Google/Яндекс Формы (для будущих опросов):", None),
    ("  • Создайте онлайн-версию анкеты — данные автоматически в Excel/CSV.", None),
    ("  • Этот шаблон можно использовать для сверки и доп. анализа.", None),
    ("", None),
    ("4. АНАЛИЗ РЕЗУЛЬТАТОВ", "header"),
    ("После заполнения запустите скрипт: python3 analyze_tso.py", None),
    ("Скрипт создаст: сводную таблицу, графики, интерпретацию результатов.", None),
    ("", None),
    ("5. ОБОЗНАЧЕНИЯ СТОЛБЦОВ", "header"),
    ("В1–В6 — номера блоков анкеты", None),
    ("Q1–Q17 — номера вопросов", None),
    ("Множественный выбор: 1 = отмечено галочкой, 0 = не отмечено, пусто = нет данных", None),
    ("Шкала: число от 0 до 10", None),
]

for row_idx, (text, style) in enumerate(instructions, 1):
    cell = ws_instr.cell(row=row_idx, column=1, value=text)
    if style == "header":
        cell.font = Font(name="Calibri", bold=True, size=12, color=DARK_BLUE)
    elif row_idx == 1:
        cell.font = Font(name="Calibri", bold=True, size=14, color=DARK_BLUE)
    else:
        cell.font = Font(name="Calibri", size=11)
    cell.alignment = Alignment(wrap_text=True)

ws_instr.column_dimensions["A"].width = 100


# ═══════════════════════════════════════════════════════
# Сохранение
# ═══════════════════════════════════════════════════════
output_path = "/home/z/my-project/download/Шаблон_ввода_анкет_ТСО.xlsx"
wb.save(output_path)
print(f"✅ Шаблон создан: {output_path}")
print(f"   Листы: {wb.sheetnames}")
print(f"   Столбцов: {len(columns)}")
print(f"   Строк для ввода: 50")
