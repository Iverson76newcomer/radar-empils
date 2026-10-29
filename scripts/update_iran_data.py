#!/usr/bin/env python3
"""
Обновить данные по Ирану в файле Экспорт_потенциал_полный_с_потреблением_с_выводом.xlsx
на основе найденных через web-search данных:
- Производство: 500 тыс. т (фактическое, среднее за последние годы)
- Установленная мощность: 900 тыс. т (загрузка 55%)
- Источник: Paint and Resin Manufacturers Association of Iran (Coatings World)
"""

import os
import shutil
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

SRC = "/home/z/my-project/download/Экспорт_потенциал_полный_с_потреблением_с_выводом.xlsx"
DST = "/home/z/my-project/download/Экспорт_потенциал_полный_Иран_обновлён.xlsx"
shutil.copy(SRC, DST)

# === Обновлённые данные по Ирану ===
# Иран — колонка 8

# R6 «Производство» — обновляем
IRAN_PRODUCTION = (
    "Производство ЛКМ – ~500 тыс. т/год (среднее за последние годы). "
    "Установленная мощность 900 тыс. т/год (загрузка ~55%). "
    "350+ авторизованных заводов + 600+ малых предприятий. "
    "Источники: Paint and Resin Manufacturers Association of Iran, Mordor Intelligence (2024). "
    "Крупнейшие: Pars Diba Chemical (Barad), Pars Pamchal, Rangsazi Iran (RICO), Alvan Paint, Peka Chemie"
)

# R7 «Потребление» — обновляем
IRAN_CONSUMPTION = (
    "Потребление ЛКМ – ~600 тыс. т/год (оценочно). "
    "Дефицит ~100 тыс. т/год (16,7% от потребл.). "
    "Рынок: $454 млн (2025) → $537 млн (2031), CAGR 2,85% (Mordor Intelligence). "
    "На 1 жителя: ~6,45 кг/год. Импорт сырья 25% (для сложных систем ЛКМ). "
    "2-й в мире экспортёр ЛКМ после Турции"
)

# R8 «Соотношение населения к производству ЛКМ» — обновляем
def make_ratio_text(pop, prod):
    kg_per_capita = prod / pop
    t_per_million = kg_per_capita * 1000
    t_str = f"{int(t_per_million):,}".replace(",", "\u202f")  # узкий неразрывный пробел
    return (f"{kg_per_capita:.2f} кг/чел. в год  ·  {t_str} т на 1 млн жителей  ·  "
            f"произв. ~500 тыс. т / нас. 93 млн чел. (загрузка мощностей 55%)")

IRAN_RATIO = make_ratio_text(93, 500)

# === Открываем файл и обновляем ячейки ===

wb = load_workbook(DST)
ws = wb["Лист1"]

# Стили
thin_border = Border(
    top=Side(style="thin", color="DDD0C8"),
    bottom=Side(style="thin", color="DDD0C8"),
    left=Side(style="thin", color="DDD0C8"),
    right=Side(style="thin", color="DDD0C8"),
)

# Цвет для обновлённых ячеек — слегка выделим (светло-жёлтый, чтобы было видно что обновлено)
UPDATE_HIGHLIGHT = PatternFill("solid", fgColor="FFF59D")  # светло-жёлтый
font_val = Font(name="Calibri", size=11, color="1F2937")

# Обновляем ячейки Ирана (колонка 8)
updates = [
    (6, IRAN_PRODUCTION),
    (7, IRAN_CONSUMPTION),
    (8, IRAN_RATIO),
]

for row, text in updates:
    cell = ws.cell(row=row, column=8, value=text)
    cell.font = font_val
    cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
    cell.fill = UPDATE_HIGHLIGHT
    cell.border = thin_border

# Сохраняем
wb.save(DST)

print(f"✅ Файл сохранён: {DST}")
print(f"   Размер: {os.path.getsize(DST):,} байт")
print(f"\n=== ОБНОВЛЁННЫЕ ДАННЫЕ ПО ИРАНУ ===\n")
print(f"R6 (Производство):")
print(f"  {IRAN_PRODUCTION}\n")
print(f"R7 (Потребление):")
print(f"  {IRAN_CONSUMPTION}\n")
print(f"R8 (Соотношение):")
print(f"  {IRAN_RATIO}")
