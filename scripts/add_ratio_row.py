#!/usr/bin/env python3
"""
Добавить новую строку "Соотношение населения к производству ЛКМ"
между строкой "Производство" (R6) и "Экспорт" (R7) в файл
Экспорт_потенциал_полный.xlsx

Показатель: кг ЛКМ на 1 жителя в год = Производство (тыс. т) × 1000 / Население (млн чел)
"""

import os
import shutil
from copy import copy
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

SRC = "/home/z/my-project/upload/Экспорт_потенциал_полный.xlsx"
DST = "/home/z/my-project/download/Экспорт_потенциал_полный_с_соотношением.xlsx"
shutil.copy(SRC, DST)

# === Данные расчёта: (страна, колонка, население_млн, произв_тыс_т_среднее, диапазон) ===
# population и production берём из исходного файла

CALC = {
    "Турция":              {"col": 3, "pop": 88,    "pop_text": "88 млн.чел.",          "prod": 875,   "prod_range": "850-900"},
    "Иран":                {"col": 4, "pop": 93,    "pop_text": "93 млн.чел",           "prod": 450,   "prod_range": "оценочно ~450"},  # нет точных данных
    "Молдова":             {"col": 5, "pop": 2.4,   "pop_text": "2,4 млн. чел.",        "prod": 35,    "prod_range": "30-40"},
    "Саудовская Аравия":   {"col": 6, "pop": 36.4,  "pop_text": "36,4 млн чел.",        "prod": 425,   "prod_range": "400-450"},
    "Сербия":              {"col": 7, "pop": 6.57,  "pop_text": "6,57 млн чел.",        "prod": 115,   "prod_range": "100-130"},
    "Азербайджан":         {"col": 8, "pop": 10.4,  "pop_text": "10,4 млн чел.",        "prod": 27.5,  "prod_range": "25-30"},
    "Россия":              {"col": 9, "pop": 146,   "pop_text": "146 млн чел.",         "prod": 1115,  "prod_range": "1 115"},
}

# === Открываем файл и вставляем новую строку после R6 (Производство) ===

wb = load_workbook(DST)
ws = wb["Лист1"]

# Вставляем пустую строку 7 (сдвинет текущую R7 и далее вниз)
ws.insert_rows(7)

# Копируем формат со строки 6 (родительской), чтобы новая строка была похожа по стилю
# и копируем стиль ячейки колонки B (критерий) из R6, и ячеек данных из R6

# Создаём стиль для названия критерия
font_bold = Font(name="Calibri", size=11, bold=True, color="1F2937")
fill_header = PatternFill("solid", fgColor="F3F1ED")  # светлый фон как у строки "Производство"
font_val = Font(name="Calibri", size=11, color="1F2937", bold=False)

# Цвета по странам — как в прошлой версии
COUNTRY_COLORS = {
    "Турция":              None,                                  # исходные данные — без подсветки
    "Иран":                None,
    "Молдова":            "FFF9C4",  # светло-жёлтый
    "Саудовская Аравия":  "FFE0B2",  # светло-оранжевый
    "Сербия":             "E3F2FD",  # светло-голубой
    "Азербайджан":        "F8BBD0",  # светло-розовый
    "Россия":             "C8E6C9",  # светло-зелёный
}

thin_border = Border(
    top=Side(style="thin", color="DDD0C8"),
    bottom=Side(style="thin", color="DDD0C8"),
    left=Side(style="thin", color="DDD0C8"),
    right=Side(style="thin", color="DDD0C8"),
)

# Записываем название критерия в колонку B (R7C2)
crit_cell = ws.cell(row=7, column=2, value="Соотношение населения к производству ЛКМ")
crit_cell.font = font_bold
crit_cell.fill = fill_header
crit_cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
crit_cell.border = thin_border

# Заполняем ячейки данных для каждой страны
for country, info in CALC.items():
    pop = info["pop"]
    prod = info["prod"]
    prod_range = info["prod_range"]
    col = info["col"]

    # Рассчитываем кг на 1 жителя
    # prod (тыс. т) = prod * 1000 т = prod * 1 000 000 кг
    # pop (млн) = pop * 1 000 000 чел
    # kg_per_capita = (prod * 1 000 000) / (pop * 1 000 000) = prod / pop
    kg_per_capita = prod / pop
    # Также считаем тонн на 1 млн жителей (= кг/чел * 1000)
    t_per_million = kg_per_capita * 1000

    # Формируем текст ячейки
    # Для разбиения разрядов т/млн используем неразрывный пробел (функция format)
    t_per_million_str = f"{int(t_per_million):,}".replace(",", "\u202f")  # узкий неразрывный пробел

    if country == "Иран":
        text = (f"≈ {kg_per_capita:.2f} кг/чел. в год  ·  "
                f"{t_per_million_str} т на 1 млн жителей  ·  "
                f"оценочно (точное производство не опубликовано)")
    else:
        text = (f"{kg_per_capita:.2f} кг/чел. в год  ·  "
                f"{t_per_million_str} т на 1 млн жителей  ·  "
                f"произв. {prod_range} тыс. т / нас. {info['pop_text']}")

    cell = ws.cell(row=7, column=col, value=text)
    cell.font = font_val
    cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)

    # Подсветка
    color = COUNTRY_COLORS.get(country)
    if color:
        cell.fill = PatternFill("solid", fgColor=color)
    else:
        # Для Турции и Ирана — слегка осветлим, чтобы выделить новую строку
        cell.fill = PatternFill("solid", fgColor="F9F9F9")
    cell.border = thin_border

# Установим высоту строки 7 — побольше, чтобы текст поместился
ws.row_dimensions[7].height = 70

# Сохраним
wb.save(DST)

# Вывод результатов
print(f"✅ Файл сохранён: {DST}")
print(f"   Размер: {os.path.getsize(DST):,} байт")
print(f"\n=== РАСЧЁТ: кг ЛКМ на 1 жителя в год ===")
print(f"{'Страна':22} {'Население':>10} {'Произв.':>12} {'кг/чел':>10} {'т/млн чел':>12}")
print("-" * 70)
for country, info in CALC.items():
    pop = info["pop"]
    prod = info["prod"]
    kg = prod / pop
    t = kg * 1000
    print(f"{country:22} {pop:>8} млн {prod:>9} тыс.т {kg:>8.2f} {t:>10,.0f}".replace(",", " "))

# Сводка по рейтингу
print(f"\n=== РЕЙТИНГ ПО УДЕЛЬНОМУ ПРОИЗВОДСТВУ ===")
ranked = sorted(CALC.items(), key=lambda x: x[1]["prod"] / x[1]["pop"], reverse=True)
for i, (country, info) in enumerate(ranked, 1):
    kg = info["prod"] / info["pop"]
    print(f"  {i}. {country:22} — {kg:.2f} кг/чел")
