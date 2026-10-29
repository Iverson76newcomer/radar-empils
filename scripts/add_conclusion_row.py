#!/usr/bin/env python3
"""
Заполнить строку R15 "Вывод" в файле Экспорт_потенциал_полный_с_потреблением.xlsx
выводами по каждой стране (приоритет для выхода Эмпилс).
"""

import os
import shutil
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

SRC = "/home/z/my-project/upload/Экспорт_потенциал_полный_с_потреблением.xlsx"
DST = "/home/z/my-project/download/Экспорт_потенциал_полный_с_потреблением_с_выводом.xlsx"
shutil.copy(SRC, DST)

# === Выводы по каждой стране ===
# Порядок колонок (по R4): C=Сербия, D=Молдова, E=Сауд.Аравия, F=Турция, G=Россия, H=Иран, I=Азербайджан

CONCLUSIONS = {
    # Сербия (col 3)
    "Сербия": (
        "НЕ ПРИОРИТЕТНО. Сербия — чистый экспортёр (профицит 45 тыс. т, 64% от произв.). "
        "Локальное производство выше потребления в 1,6 раза, рынок перенасыщен. "
        "Сильны локальные гиганты JUB Serbia, Irkom. Вход возможен только через нишевые продукты "
        "(PROFIWOOD, TITAN для промклиентов), массовый экспорт нецелесообразен."
    ),
    # Молдова (col 4)
    "Молдова": (
        "ПРИОРИТЕТ 2. Дефицит 30% (15 тыс. т/год) — импорт-доминирующий рынок. "
        "Малый рынок (50 тыс. т), но низкая конкуренция локальных производителей, "
        "простая логистика через Украину/РФ, договор СНГ. "
        "Подходит как «быстрая победа» с малыми инвестициями. Фокус: AvantGarde, NewTone в DIY."
    ),
    # Саудовская Аравия (col 5)
    "Саудовская Аравия": (
        "ПРИОРИТЕТ 3 (отложенный). Дефицит 125 тыс. т (23% потребл), но рынок зрелый "
        "(15 кг/чел), премиум-сегмент, жёсткая конкуренция (Jotun, Hempel, Sigma, National). "
        "Требует сертификатов, локального партнёра, выхода через хаб ОАЭ. "
        "Заходить после 2027 г. с защитными покрытиями TITAN для нефтегаза."
    ),
    # Турция (col 6)
    "Турция": (
        "НЕ ПРИОРИТЕТНО. Турция — крупный экспортёр (профицит 175 тыс. т, 5-е место в мире). "
        "Прямой конкурент Эмпилс на рынках СНГ и Центральной Азии. "
        "Сильная локальная промышленность (Marshall, DYO, Polisan, Filli). "
        "Вход возможен только через точечные партнёрства (Duayen — промпокрытия)."
    ),
    # Россия (col 7)
    "Россия": (
        "ВНУТРЕННИЙ РЫНОК. Россия — дом для Эмпилс (5-6 место в мире по произв.). "
        "Дефицит всего 6%, рынок зрелый (8,1 кг/чел). Стратегия — не экспорт, а удержание "
        "и расширение доли через импортозамещение и сдвиг в премиум (AvantGarde, EMPILS ULTRA). "
        "Рост водных ЛКМ +3,9%/год — основной тренд для развития."
    ),
    # Иран (col 8)
    "Иран": (
        "НЕ ПРИОРИТЕТНО (пока). Дефицит 150 тыс. т (25% потребл) — большой потенциал, "
        "НО санкции РФ + Ирана делают вход невозможным в текущих условиях. "
        "600+ локальных производителей закрывают большую часть спроса. "
        "Возможен только обмен сырьём или технологиями при изменении геополитики."
    ),
    # Азербайджан (col 9)
    "Азербайджан": (
        "ПРИОРИТЕТ 1. Самый перспективный рынок для Эмпилс: дефицит 42% (самый высокий в списке), "
        "удельное потребление 4,6 кг/чел (в 3 раза ниже Сербии — есть куда расти), "
        "рост отрасли 8-12%/год (самый высокий). Эмпилс уже работает с FAB Boya — готовая точка входа. "
        "Расширение: TITAN для нефтегаза + водные ЛКМ (AvantGarde, NewTone) в DIY."
    ),
}

# === Открываем файл и заполняем R15 ===

wb = load_workbook(DST)
ws = wb["Лист1"]

# Стили
thin_border = Border(
    top=Side(style="thin", color="DDD0C8"),
    bottom=Side(style="thin", color="DDD0C8"),
    left=Side(style="thin", color="DDD0C8"),
    right=Side(style="thin", color="DDD0C8"),
)

# Цвета по приоритету вывода
PRIORITY_COLORS = {
    "Азербайджан":         "C8E6C9",  # зелёный — приоритет 1
    "Молдова":             "DCEDC8",  # светло-зелёный — приоритет 2
    "Саудовская Аравия":  "FFE0B2",  # оранжевый — приоритет 3 (отложенный)
    "Сербия":              "FFCDD2",  # красный — не приоритет
    "Турция":              "FFCDD2",  # красный — не приоритет
    "Иран":                "E0E0E0",  # серый — невозможен (санкции)
    "Россия":              "BBDEFB",  # голубой — внутренний рынок
}

# Жирность шрифта по приоритету
def get_font(country):
    if country == "Азербайджан":
        return Font(name="Calibri", size=11, color="1B5E20", bold=True)  # тёмно-зелёный жирный
    elif country == "Молдова":
        return Font(name="Calibri", size=11, color="33691E", bold=True)  # зелёный жирный
    elif country == "Саудовская Аравия":
        return Font(name="Calibri", size=11, color="E65100", bold=True)  # оранжевый жирный
    elif country in ("Сербия", "Турция"):
        return Font(name="Calibri", size=11, color="B71C1C", bold=True)  # тёмно-красный жирный
    elif country == "Иран":
        return Font(name="Calibri", size=11, color="616161", italic=True)  # серый курсив
    elif country == "Россия":
        return Font(name="Calibri", size=11, color="0D47A1", bold=True)  # тёмно-синий жирный
    return Font(name="Calibri", size=11, color="1F2937")

# Маппинг колонок
COL_MAP = {
    "Сербия":             3,
    "Молдова":            4,
    "Саудовская Аравия": 5,
    "Турция":             6,
    "Россия":             7,
    "Иран":               8,
    "Азербайджан":       9,
}

# Заполняем
for country, text in CONCLUSIONS.items():
    col = COL_MAP[country]
    cell = ws.cell(row=15, column=col, value=text)
    cell.font = get_font(country)
    cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
    cell.fill = PatternFill("solid", fgColor=PRIORITY_COLORS[country])
    cell.border = thin_border

# Стиль для заголовка "Вывод" в R15C2
crit_cell = ws.cell(row=15, column=2)
crit_cell.font = Font(name="Calibri", size=11, bold=True, color="1F2937")
crit_cell.fill = PatternFill("solid", fgColor="F3F1ED")
crit_cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
crit_cell.border = thin_border

# Высота строки 15 — побольше, чтобы текст поместился
ws.row_dimensions[15].height = 110

# Сохраним
wb.save(DST)

print(f"✅ Файл сохранён: {DST}")
print(f"   Размер: {os.path.getsize(DST):,} байт")
print(f"\n=== СОДЕРЖИМОЕ R15 «Вывод» ===\n")
for country, text in CONCLUSIONS.items():
    print(f"🌍 {country}:")
    print(f"   {text}")
    print()
