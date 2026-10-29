#!/usr/bin/env python3
"""TCO Калькулятор — Технологические цепочки Лемана ПРО (Эмпилс) — на русском"""

import sys, os, math
XLSX_SKILL_DIR = "/home/z/my-project/skills/xlsx"
for sub in [XLSX_SKILL_DIR, os.path.join(XLSX_SKILL_DIR, "templates")]:
    if sub not in sys.path:
        sys.path.insert(0, sub)

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, Reference
from openpyxl.utils import get_column_letter

from templates.base import (
    use_palette_explicit, setup_sheet, style_header_row, style_data_row,
    style_total_row, font_title, font_header, font_subheader, font_body,
    font_caption, fill_header, fill_total, fill_data_row, border_header,
    border_total, align_title, align_header, align_number, align_text,
    PRIMARY, PRIMARY_LIGHT, SECONDARY,
    ACCENT_POSITIVE, ACCENT_NEGATIVE, ACCENT_WARNING,
    NEUTRAL_900, NEUTRAL_600, NEUTRAL_200, NEUTRAL_100, NEUTRAL_0,
    CHART_COLORS, COLUMN_WIDTHS, ROW_HEIGHTS,
    make_chart_title, setup_chart_titles, apply_chart_colors,
    create_bar_chart, FONT_NAME, HEADER_BOLD,
)

use_palette_explicit("warm")

YELLOW_INPUT = PatternFill("solid", fgColor="FFF9C4")
INPUT_FONT = Font(name=FONT_NAME, size=11, color="0000FF", bold=HEADER_BOLD)
SECTION_FILL = PatternFill("solid", fgColor=SECONDARY)
SAVINGS_FILL = PatternFill("solid", fgColor="E8F5E9")
SAVINGS_FONT = Font(name=FONT_NAME, size=11, color=ACCENT_POSITIVE, bold=HEADER_BOLD)
LOSS_FILL = PatternFill("solid", fgColor="FDEDEC")
LOSS_FONT = Font(name=FONT_NAME, size=11, color=ACCENT_NEGATIVE, bold=HEADER_BOLD)

FMT_RUB  = '#,##0" ₽"'
FMT_RUB2 = '#,##0.00" ₽"'
FMT_PCT  = '0.0%'
FMT_INT  = '#,##0'

wb = Workbook()
wb.properties.creator = "Z.ai"

# ===== ДАННЫЕ =====
OBJECTS = {
    "Ворота металлические": 16, "Гараж металлический": 54, "Дверь металлическая": 4.4,
    "Забор металлический": 30, "Трубы / конструкции": 10,
    "Пол гаража (бетон)": 20, "Пол подвала (бетон)": 15,
    "Пол склада (бетон)": 50, "Забор деревянный": 40,
    "Терраса / лоджия": 15, "Потолок": 20,
    "Стены (коридор/кухня)": 50, "Фасад дома": 100,
    "Стена декоративная": 20, "Батареи (радиаторы)": 6, "Настил металлический": 25,
}

PRICES = [
    ("Грунтовка", "Грунтовка ГФ-021 Л", "ОРЕОЛ", "1 кг", 180, "кг"),
    ("Грунтовка", "Грунтовка ГФ-021 Л", "ОРЕОЛ", "1.9 кг", 310, "кг"),
    ("Грунтовка", "Грунтовка ГФ-021 Л", "ОРЕОЛ", "5.5 кг", 680, "кг"),
    ("Грунтовка ВД", "Грунтовка ВД универсальная", "ОРЕОЛ ДИСКОНТ", "20 кг", 620, "кг"),
    ("Эмаль ПФ-115", "Эмаль ПФ-115", "ПростоКРАСКА", "0.8 кг", 130, "кг"),
    ("Эмаль ПФ-115", "Эмаль ПФ-115", "ПростоКРАСКА", "1.9 кг", 260, "кг"),
    ("Эмаль ПФ-115", "Эмаль ПФ-115", "ПростоКРАСКА", "5.5 кг", 580, "кг"),
    ("Эмаль ПФ-115", "Эмаль ПФ-115", "ПростоКРАСКА", "9 кг", 890, "кг"),
    ("Эмаль ПФ-115", "Эмаль ПФ-115", "ПростоКРАСКА", "20 кг", 1790, "кг"),
    ("Эмаль ПФ-115", "Эмаль ПФ-115 Premium", "ОРЕОЛ", "0.9 кг", 280, "кг"),
    ("Эмаль ПФ-115", "Эмаль ПФ-115 Premium", "ОРЕОЛ", "2.2 кг", 530, "кг"),
    ("Эмаль ПФ-115", "Эмаль ПФ-115 ЭКОНОМ Л", "ЭКОНОМ Л", "9 кг", 820, "кг"),
    ("Эмаль ПФ-115", "Эмаль ПФ-115 ЭКОНОМ Л", "ЭКОНОМ Л", "20 кг", 1650, "кг"),
    ("Грунт-эмаль 3в1", "Грунт-эмаль 3в1 по ржавчине", "ОРЕОЛ", "0.9 кг", 360, "кг"),
    ("Грунт-эмаль 3в1", "Грунт-эмаль 3в1 по ржавчине", "ОРЕОЛ", "2.6 кг", 770, "кг"),
    ("Грунт-эмаль 3в1", "Грунт-эмаль 3в1 по ржавчине", "ОРЕОЛ", "5 кг", 1250, "кг"),
    ("Грунт-эмаль 3в1", "Грунт-эмаль 3в1 по ржавчине", "ОРЕОЛ", "10 кг", 2150, "кг"),
    ("Грунт-эмаль TITAN", "Грунт-эмаль TITAN по ржавчине", "ОРЕОЛ TITAN", "0.4 л", 460, "л"),
    ("Грунт-эмаль TITAN", "Грунт-эмаль TITAN по ржавчине", "ОРЕОЛ TITAN", "0.75 л", 720, "л"),
    ("Грунт-эмаль TITAN", "Грунт-эмаль TITAN по ржавчине", "ОРЕОЛ TITAN", "2 л", 1520, "л"),
    ("Грунт-эмаль TITAN", "Грунт-эмаль TITAN по ржавчине", "ОРЕОЛ TITAN", "5 л", 3050, "л"),
    ("Грунт-эмаль Paint Lab", "Грунт-эмаль Paint Lab", "EMPILS Paint Lab", "0.9 кг", 410, "кг"),
    ("Грунт-эмаль Paint Lab", "Грунт-эмаль Paint Lab", "EMPILS Paint Lab", "2.6 кг", 920, "кг"),
    ("Грунт-эмаль Paint Lab", "Грунт-эмаль Paint Lab", "EMPILS Paint Lab", "5 кг", 1550, "кг"),
    ("Эмаль для пола", "Эмаль Расцвет для полов АК", "Расцвет", "0.9 кг", 280, "кг"),
    ("Эмаль для пола", "Эмаль Расцвет для полов АК", "Расцвет", "2.7 кг", 660, "кг"),
    ("Эмаль для пола", "Эмаль Расцвет для полов АК", "Расцвет", "5 кг", 1080, "кг"),
    ("Эмаль для пола", "Эмаль Расцвет для полов АК", "Расцвет", "10 кг", 1850, "кг"),
    ("Эмаль для пола", "Эмаль Расцвет для полов АК", "Расцвет", "20 кг", 3350, "кг"),
    ("Лак яхтный", "Лак яхтный атмосферостойкий", "MONLID", "0.7 кг", 360, "кг"),
    ("Лак яхтный", "Лак яхтный атмосферостойкий", "MONLID", "1.9 кг", 720, "кг"),
    ("Лак яхтный", "Лак яхтный атмосферостойкий", "MONLID", "4.6 кг", 1450, "кг"),
    ("Лак яхтный", "Лак яхтный атмосферостойкий", "Luxens", "0.7 кг", 510, "кг"),
    ("Лак яхтный", "Лак яхтный атмосферостойкий", "Luxens", "1.9 кг", 1020, "кг"),
    ("Лак яхтный", "Лак яхтный атмосферостойкий", "Luxens", "4.6 кг", 2050, "кг"),
    ("Лак паркетный", "Лак паркетный AVANTGARDE", "AVANTGARDE", "0.75 л", 820, "л"),
    ("Лак паркетный", "Лак паркетный AVANTGARDE", "AVANTGARDE", "2 л", 1850, "л"),
    ("Лак паркетный", "Лак паркетный AVANTGARDE", "AVANTGARDE", "4.5 л", 3600, "л"),
    ("Краска ВД потолочная", "Краска ВД для потолков", "PL ЛЕРУА", "3 кг", 360, "кг"),
    ("Краска ВД потолочная", "Краска ВД для потолков", "PL ЛЕРУА", "6.5 кг", 670, "кг"),
    ("Краска ВД потолочная", "Краска ВД для потолков", "PL ЛЕРУА", "14 кг", 1250, "кг"),
    ("Краска ВД для стен", "Краска ВД моющаяся (база А/С)", "ОРЕОЛ", "2.2 л", 470, "л"),
    ("Краска ВД для стен", "Краска ВД моющаяся (база А/С)", "ОРЕОЛ", "9 л", 1550, "л"),
    ("Краска ВД фасадная", "Краска ВД для фасадов (база А/С)", "ОРЕОЛ", "9 л", 1850, "л"),
    ("Краска фактурная", "Краска фактурная", "ОРЕОЛ", "8 кг", 1250, "кг"),
    ("Краска фактурная", "Краска фактурная", "ОРЕОЛ", "16 кг", 2250, "кг"),
    ("Эмаль для радиаторов", "Эмаль для радиаторов", "ПростоКРАСКА", "0.9 кг", 260, "кг"),
]

CONSUMPTION = [
    ("Грунтовка ГФ-021", 70, 1, None, 124),
    ("Грунтовка ВД", 100, 1, None, 31),
    ("Эмаль ПФ-115 ПростоКРАСКА", 120, 2, 3, 90),
    ("Эмаль ПФ-115 ОРЕОЛ Premium", 110, 2, 5, 241),
    ("Эмаль ПФ-115 ЭКОНОМ Л", 130, 2, 2, 83),
    ("Грунт-эмаль 3в1 ОРЕОЛ", 100, 2, 4, 215),
    ("Грунт-эмаль TITAN", 90, 2, 8, 610),
    ("Грунт-эмаль Paint Lab", 95, 2, 5, 310),
    ("Эмаль Расцвет для пола", 120, 2, 7, 168),
    ("Лак яхтный MONLID", 90, 3, 4, 315),
    ("Лак яхтный Luxens", 85, 3, 6, 446),
    ("Лак паркетный AVANTGARDE", 100, 3, 7, 800),
    ("Краска ВД потолочная", 130, 2, 7, 89),
    ("Краска ВД моющаяся", 120, 2, 7, 172),
    ("Краска ВД фасадная", 130, 2, 6, 206),
    ("Краска фактурная", 800, 1, 7, 141),
    ("Эмаль для радиаторов", 100, 2, 7, 289),
]

CONS_DICT = {}
for cat, cons, layers, life, price in CONSUMPTION:
    CONS_DICT[cat] = {"consumption": cons, "layers": layers, "service_life": life, "price_per_unit": price}

LABOR = {
    "Грунтование металла": 250, "Покраска металла": 300, "Грунтование ВД": 180,
    "Покраска ВД": 220, "Покраска пола": 280, "Лакирование дерево": 300,
    "Лакирование паркет": 350, "Покраска радиаторов": 400, "Декоративная покраска": 450,
}

TCO_CHAINS = [
    {"num":1,"name":"Защита металла: грунт + эмаль","segment":"Металл","problem":"Ржавый металл: заборы, ворота, трубы, гаражи. Нужно надёжное покрытие на годы.","default_object":"Ворота металлические","priority":"ВЫСОКИЙ",
     "options":[
         {"name":"Эконом: ГФ-021 + ПФ-115 ПростоКРАСКА","steps":[("Грунтовка ГФ-021","Грунтование металла"),("Эмаль ПФ-115 ПростоКРАСКА","Покраска металла")]},
         {"name":"Средний: ГФ-021 + ПФ-115 ОРЕОЛ Premium","steps":[("Грунтовка ГФ-021","Грунтование металла"),("Эмаль ПФ-115 ОРЕОЛ Premium","Покраска металла")]},
     ]},
    {"num":2,"name":"Быстрая покраска металла (3в1)","segment":"Металл","problem":"Нет времени грунтовать — нужно покрасить ржавый забор или трубу за один проход.","default_object":"Забор металлический","priority":"ВЫСОКИЙ",
     "options":[
         {"name":"3в1 ОРЕОЛ","steps":[("Грунт-эмаль 3в1 ОРЕОЛ","Покраска металла")]},
         {"name":"Paint Lab","steps":[("Грунт-эмаль Paint Lab","Покраска металла")]},
         {"name":"TITAN (премиум)","steps":[("Грунт-эмаль TITAN","Покраска металла")]},
     ]},
    {"num":3,"name":"Премиум-покраска металла по RAL","segment":"Металл","problem":"Нужен точный цвет по RAL для ворот, заборов, оборудования.","default_object":"Ворота металлические","priority":"ВЫСОКИЙ",
     "options":[{"name":"TITAN (грунт-эмаль по ржавчине)","steps":[("Грунт-эмаль TITAN","Покраска металла")]}]},
    {"num":4,"name":"Современная покраска металла — Paint Lab","segment":"Металл","problem":"Хочется стильнее классики, но без переплаты за TITAN.","default_object":"Забор металлический","priority":"СРЕДНИЙ",
     "options":[{"name":"Paint Lab","steps":[("Грунт-эмаль Paint Lab","Покраска металла")]}]},
    {"num":5,"name":"Бетонный пол без пыли и истирания","segment":"Полы","problem":"Бетонный пол в гараже/подвале пылит, впитывает масло, крошится.","default_object":"Пол гаража (бетон)","priority":"ВЫСОКИЙ",
     "options":[
         {"name":"Эконом: только Расцвет (без грунта)","steps":[("Эмаль Расцвет для пола","Покраска пола")],"consumption_multiplier":1.8,"life_override":4},
         {"name":"Рекомендовано: ГФ-021 + Расцвет","steps":[("Грунтовка ГФ-021","Грунтование металла"),("Эмаль Расцвет для пола","Покраска пола")]},
     ]},
    {"num":6,"name":"Дерево на улице: забор, терраса","segment":"Дерево (экстерьер)","problem":"Забор/терраса сереют от солнца, разбухают от дождя, трескаются от мороза.","default_object":"Забор деревянный","priority":"СРЕДНИЙ",
     "options":[
         {"name":"Эконом: MONLID","steps":[("Лак яхтный MONLID","Лакирование дерево")]},
         {"name":"Средний+: Luxens","steps":[("Лак яхтный Luxens","Лакирование дерево")]},
     ]},
    {"num":7,"name":"Паркет и деревянный пол внутри","segment":"Дерево (интерьер)","problem":"Паркет истирается, царапается, тускнеет. Обычные лаки желтеют.","default_object":"Терраса / лоджия","priority":"НИЗКИЙ",
     "options":[{"name":"Премиум: AVANTGARDE","steps":[("Лак паркетный AVANTGARDE","Лакирование паркет")]}]},
    {"num":8,"name":"Ровный потолок без разводов","segment":"Потолки","problem":"Потолочная краска ложится полосами, видны разводы от валика.","default_object":"Потолок","priority":"НИЗКИЙ",
     "options":[
         {"name":"Без грунта: только краска","steps":[("Краска ВД потолочная","Покраска ВД")],"consumption_multiplier":1.5,"life_override":4},
         {"name":"Рекомендовано: грунт + краска","steps":[("Грунтовка ВД","Грунтование ВД"),("Краска ВД потолочная","Покраска ВД")]},
     ]},
    {"num":9,"name":"Стены, которые можно мыть","segment":"Интерьер","problem":"Стены пачкаются, а обычная краска размокает при мытье.","default_object":"Стены (коридор/кухня)","priority":"НИЗКИЙ",
     "options":[
         {"name":"Без грунта: только краска","steps":[("Краска ВД моющаяся","Покраска ВД")],"consumption_multiplier":1.5,"life_override":4},
         {"name":"Рекомендовано: грунт + краска","steps":[("Грунтовка ВД","Грунтование ВД"),("Краска ВД моющаяся","Покраска ВД")]},
     ]},
    {"num":10,"name":"Покраска фасада дома","segment":"Фасад","problem":"Фасад потрескался, выцвел, нужно обновить защиту и цвет.","default_object":"Фасад дома","priority":"НИЗКИЙ",
     "options":[
         {"name":"Без грунта: только краска","steps":[("Краска ВД фасадная","Покраска ВД")],"consumption_multiplier":1.5,"life_override":4},
         {"name":"Рекомендовано: грунт + краска","steps":[("Грунтовка ВД","Грунтование ВД"),("Краска ВД фасадная","Покраска ВД")]},
     ]},
    {"num":11,"name":"Декоративная стена с текстурой","segment":"Интерьер / Декор","problem":"Гладкая краска выглядит скучно, хочется рельеф и фактуру.","default_object":"Стена декоративная","priority":"НИЗКИЙ",
     "options":[{"name":"Грунт + фактурная краска","steps":[("Грунтовка ВД","Грунтование ВД"),("Краска фактурная","Декоративная покраска")]}]},
    {"num":12,"name":"Покраска батарей без запаха","segment":"Радиаторы","problem":"Ржавые батареи, нужна краска, которая не желтеет от нагрева.","default_object":"Батареи (радиаторы)","priority":"СРЕДНИЙ",
     "options":[
         {"name":"Без грунта: только эмаль","steps":[("Эмаль для радиаторов","Покраска радиаторов")],"life_override":4},
         {"name":"Рекомендовано: ГФ-021 + эмаль","steps":[("Грунтовка ГФ-021","Грунтование металла"),("Эмаль для радиаторов","Покраска радиаторов")]},
     ]},
    {"num":13,"name":"Усиленная защита ржавого металла","segment":"Металл (тяжёлые условия)","problem":"Металл сильно проржавел — 3в1 не справится, нужна максимальная защита.","default_object":"Гараж металлический","priority":"ВЫСОКИЙ",
     "options":[
         {"name":"Только 3в1","steps":[("Грунт-эмаль 3в1 ОРЕОЛ","Покраска металла")]},
         {"name":"ГФ-021 + 3в1 (двойная защита)","steps":[("Грунтовка ГФ-021","Грунтование металла"),("Грунт-эмаль 3в1 ОРЕОЛ","Покраска металла")],"life_override":10},
         {"name":"ГФ-021 + TITAN (максимум)","steps":[("Грунтовка ГФ-021","Грунтование металла"),("Грунт-эмаль TITAN","Покраска металла")],"life_override":12},
     ]},
    {"num":14,"name":"Металлический пол / настил","segment":"Металл + Полы","problem":"Металлический настил/ступени ржавеют и скользят.","default_object":"Настил металлический","priority":"ВЫСОКИЙ",
     "options":[
         {"name":"Только Расцвет (без антикоррозийной)","steps":[("Эмаль Расцвет для пола","Покраска пола")],"life_override":3},
         {"name":"Рекомендовано: ГФ-021 + Расцвет","steps":[("Грунтовка ГФ-021","Грунтование металла"),("Эмаль Расцвет для пола","Покраска пола")]},
     ]},
]


def calc_tco(chain, option, area, horizon, labor_rates):
    material_cost_per_m2 = 0; labor_cost_per_m2 = 0
    repaint_material_per_m2 = 0; repaint_labor_per_m2 = 0
    cons_mult = option.get("consumption_multiplier", 1.0)
    has_ground = len(option["steps"]) > 1

    for i, (cat, labor_type) in enumerate(option["steps"]):
        info = CONS_DICT[cat]
        total_consumption = info["consumption"] * info["layers"] * cons_mult
        mat_cost = total_consumption / 1000 * info["price_per_unit"]
        material_cost_per_m2 += mat_cost
        labor_rate = labor_rates.get(labor_type, 250)
        labor_cost_per_m2 += labor_rate
        if i == len(option["steps"]) - 1:
            repaint_material_per_m2 += mat_cost
            repaint_labor_per_m2 += labor_rate
        elif has_ground:
            repaint_labor_per_m2 += labor_rate * 0.3

    total_initial = (material_cost_per_m2 + labor_cost_per_m2) * area
    repaint_cost = (repaint_material_per_m2 + repaint_labor_per_m2) * area

    if "life_override" in option:
        service_life = option["life_override"]
    else:
        lives = [CONS_DICT[cat]["service_life"] for cat, _ in option["steps"] if CONS_DICT[cat]["service_life"] is not None]
        service_life = min(lives) if lives else 5

    cycles = math.ceil(horizon / service_life)
    total_tco = total_initial + (cycles - 1) * repaint_cost

    return {
        "material_per_m2": round(material_cost_per_m2, 2),
        "labor_per_m2": round(labor_cost_per_m2, 2),
        "material_total": round(material_cost_per_m2 * area, 0),
        "labor_total": round(labor_cost_per_m2 * area, 0),
        "initial_cost": round(total_initial, 0),
        "repaint_cost": round(repaint_cost, 0),
        "service_life": service_life, "cycles": cycles,
        "tco": round(total_tco, 0),
    }


# ===== ЛИСТ 1: Параметры =====
ws1 = wb.active; ws1.title = "Параметры"
setup_sheet(ws1, title="ВВОДНЫЕ ПАРАМЕТРЫ ДЛЯ РАСЧЁТА ТСО", last_col=7)

r = 4
ws1.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
ws1.cell(row=r, column=2, value="ОБЩИЕ ПАРАМЕТРЫ").font = font_subheader()
ws1.cell(row=r, column=2).fill = SECTION_FILL; r += 1

params = [
    ("Горизонт расчёта (лет)", 10),
    ("Ставка грунтования металла (₽/м²)", 250),
    ("Ставка покраски металла (₽/м²)", 300),
    ("Ставка грунтования ВД (₽/м²)", 180),
    ("Ставка покраски ВД (₽/м²)", 220),
    ("Ставка покраски пола (₽/м²)", 280),
    ("Ставка лакирования дерево (₽/м²)", 300),
    ("Ставка лакирования паркет (₽/м²)", 350),
    ("Ставка покраски радиаторов (₽/м²)", 400),
    ("Ставка декоративной покраски (₽/м²)", 450),
]
for label, default in params:
    ws1.cell(row=r, column=2, value=label).font = font_body()
    ws1.cell(row=r, column=2).alignment = align_text()
    c = ws1.cell(row=r, column=3, value=default)
    c.font = INPUT_FONT; c.fill = YELLOW_INPUT; c.alignment = align_number(); c.number_format = FMT_INT
    ws1.row_dimensions[r].height = ROW_HEIGHTS["data"]; r += 1

r += 1
ws1.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
ws1.cell(row=r, column=2, value="ПЛОЩАДИ ТИПОВЫХ ОБЪЕКТОВ (м²)").font = font_subheader()
ws1.cell(row=r, column=2).fill = SECTION_FILL; r += 1

for obj_name, area in OBJECTS.items():
    ws1.cell(row=r, column=2, value=obj_name).font = font_body()
    ws1.cell(row=r, column=2).alignment = align_text()
    c = ws1.cell(row=r, column=3, value=area)
    c.font = INPUT_FONT; c.fill = YELLOW_INPUT; c.alignment = align_number(); c.number_format = '#,##0.0'
    ws1.row_dimensions[r].height = ROW_HEIGHTS["data"]; r += 1

r += 1
ws1.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
ws1.cell(row=r, column=2, value="Ячейки с жёлтым фоном — редактируемые. Измените значения для пересчёта.").font = font_caption()
ws1.column_dimensions["B"].width = 42; ws1.column_dimensions["C"].width = 16

# ===== ЛИСТ 2: Прайс-лист =====
ws2 = wb.create_sheet("Прайс-лист")
setup_sheet(ws2, title="ПРАЙС-ЛИСТ ЛЕМАНА ПРО — АССОРТИМЕНТ ЭМПИЛС", last_col=8)
headers_pl = ["№", "Категория", "Продукт", "Бренд", "Фасовка", "Цена (₽)", "Единица", "Цена за кг/л (₽)"]
for i, h in enumerate(headers_pl, start=2):
    ws2.cell(row=4, column=i, value=h)
style_header_row(ws2, 4, 2, 9)
for idx, (cat, product, brand, pack, price, unit) in enumerate(PRICES):
    r = 5 + idx
    try: pack_val = float(pack.replace(" кг","").replace(" л",""))
    except: pack_val = 1
    price_per_unit = round(price / pack_val, 1) if pack_val > 0 else price
    vals = [idx+1, cat, product, brand, pack, price, unit, price_per_unit]
    for i, v in enumerate(vals, start=2):
        c = ws2.cell(row=r, column=i, value=v)
        if i in (7,9): c.number_format = FMT_RUB2; c.alignment = align_number()
        elif i == 2: c.alignment = align_number()
        else: c.alignment = align_text()
    style_data_row(ws2, r, 2, 9, idx)

# ===== ЛИСТ 3: Нормы расхода =====
ws3 = wb.create_sheet("Нормы расхода")
setup_sheet(ws3, title="НОРМЫ РАСХОДА И СРОКИ СЛУЖБЫ", last_col=7)
headers_nr = ["№", "Категория продукта", "Расход (г/м²/слой)", "Слоёв", "Срок службы (лет)", "Цена за кг/л (₽)", "Расход всего (г/м²)"]
for i, h in enumerate(headers_nr, start=2):
    ws3.cell(row=4, column=i, value=h)
style_header_row(ws3, 4, 2, 8)
for idx, (cat, cons, layers, life, price) in enumerate(CONSUMPTION):
    r = 5 + idx
    vals = [idx+1, cat, cons, layers, life if life else "—", price, cons * layers]
    for i, v in enumerate(vals, start=2):
        c = ws3.cell(row=r, column=i, value=v)
        if i == 7: c.number_format = FMT_RUB2; c.alignment = align_number()
        elif i in (4,5,8): c.alignment = align_number()
        else: c.alignment = align_text()
    style_data_row(ws3, r, 2, 8, idx)

# ===== ЛИСТ 4: Калькулятор ТСО =====
ws4 = wb.create_sheet("Калькулятор ТСО")
setup_sheet(ws4, title="КАЛЬКУЛЯТОР ТСО — ТЕХНОЛОГИЧЕСКИЕ ЦЕПОЧКИ ЛЕМАНА ПРО", last_col=15)
horizon = 10
labor_rates = {"Грунтование металла":250,"Покраска металла":300,"Грунтование ВД":180,"Покраска ВД":220,"Покраска пола":280,"Лакирование дерево":300,"Лакирование паркет":350,"Покраска радиаторов":400,"Декоративная покраска":450}

r = 4; summary_data = []
for chain in TCO_CHAINS:
    ws4.merge_cells(start_row=r, start_column=2, end_row=r, end_column=15)
    cell = ws4.cell(row=r, column=2, value=f'Цепочка {chain["num"]}. {chain["name"]}  [{chain["segment"]}]  Приоритет: {chain["priority"]}')
    cell.font = Font(name=FONT_NAME, size=12, bold=HEADER_BOLD, color=PRIMARY)
    cell.fill = SECTION_FILL; cell.alignment = align_text()
    ws4.row_dimensions[r].height = 28; r += 1
    ws4.merge_cells(start_row=r, start_column=2, end_row=r, end_column=15)
    ws4.cell(row=r, column=2, value=f'Проблема: {chain["problem"]}').font = font_caption(); r += 1
    default_obj = chain.get("default_object", list(OBJECTS.keys())[0])
    area = OBJECTS.get(default_obj, 16)
    ws4.merge_cells(start_row=r, start_column=2, end_row=r, end_column=15)
    ws4.cell(row=r, column=2, value=f'Объект: {default_obj} | Площадь: {area} м² | Горизонт: {horizon} лет').font = font_body(); r += 1

    table_headers = ["Вариант","Шаг 1","Шаг 2","Материал (₽/м²)","Работа (₽/м²)","Материал всего (₽)","Работа всего (₽)","Стоимость 1 цикла (₽)","Перекраска (₽)","Срок службы (лет)","Кол-во циклов","ТСО за 10 лет (₽)","Экономия vs макс. (₽)","Экономия vs макс. (%)"]
    for i, h in enumerate(table_headers, start=2):
        ws4.cell(row=r, column=i, value=h)
    style_header_row(ws4, r, 2, 15); r += 1

    results = []
    for opt in chain["options"]:
        res = calc_tco(chain, opt, area, horizon, labor_rates)
        res["option_name"] = opt["name"]; res["steps"] = opt["steps"]
        results.append(res)

    best_tco = min(res["tco"] for res in results)
    worst_tco = max(res["tco"] for res in results)

    for idx, res in enumerate(results):
        step_names = [s[0] for s in res["steps"]]
        step1 = step_names[0] if step_names else ""
        step2 = step_names[1] if len(step_names) > 1 else "—"
        savings = worst_tco - res["tco"]
        savings_pct = savings / worst_tco if worst_tco > 0 else 0

        vals = [res["option_name"],step1,step2,res["material_per_m2"],res["labor_per_m2"],res["material_total"],res["labor_total"],res["initial_cost"],res["repaint_cost"],res["service_life"],res["cycles"],res["tco"],savings,savings_pct]
        for i, v in enumerate(vals, start=2):
            c = ws4.cell(row=r, column=i, value=v)
            if i in (5,6,7,8,9,10): c.number_format = FMT_RUB2; c.alignment = align_number()
            elif i in (12,14): c.number_format = FMT_RUB; c.alignment = align_number()
            elif i == 15: c.number_format = FMT_PCT; c.alignment = align_number()
            elif i == 13: c.number_format = FMT_INT; c.alignment = align_number()
            else: c.alignment = align_text()

        if savings > 0:
            ws4.cell(row=r,column=14).fill=SAVINGS_FILL; ws4.cell(row=r,column=14).font=SAVINGS_FONT
            ws4.cell(row=r,column=15).fill=SAVINGS_FILL; ws4.cell(row=r,column=15).font=SAVINGS_FONT
        if res["tco"] == best_tco:
            for col in range(2,16):
                c = ws4.cell(row=r,column=col)
                if not (col==14 or col==15):
                    c.font = Font(name=FONT_NAME, size=11, bold=HEADER_BOLD, color=ACCENT_POSITIVE)
        style_data_row(ws4, r, 2, 15, idx)

        summary_data.append({"chain_num":chain["num"],"chain_name":chain["name"],"segment":chain["segment"],"priority":chain["priority"],"option":res["option_name"],"object":default_obj,"area":area,"tco":res["tco"],"savings":savings,"savings_pct":savings_pct})
        r += 1
    r += 2

for col_letter, width in [('B',38),('C',22),('D',22),('E',16),('F',16),('G',18),('H',18),('I',20),('J',16),('K',14),('L',14),('M',18),('N',16),('O',14)]:
    ws4.column_dimensions[col_letter].width = width

# ===== ЛИСТ 5: Сводка =====
ws5 = wb.create_sheet("Сводка")
setup_sheet(ws5, title="СВОДНОЕ СРАВНЕНИЕ ТСО — ВСЕ ЦЕПОЧКИ", last_col=11)
best_per_chain = {}; worst_per_chain = {}
for sd in summary_data:
    k = sd["chain_num"]
    if k not in best_per_chain or sd["tco"] < best_per_chain[k]["tco"]: best_per_chain[k] = sd
    if k not in worst_per_chain or sd["tco"] > worst_per_chain[k]["tco"]: worst_per_chain[k] = sd

comparable_chains = [c for c in TCO_CHAINS if len(c["options"]) >= 2]
headers_s = ["№","Цепочка","Сегмент","Приоритет","Дорогой вариант","ТСО дорогого (₽)","Выгодный вариант","ТСО выгодного (₽)","Экономия (₽)","Экономия (%)"]
for i, h in enumerate(headers_s, start=2):
    ws5.cell(row=4, column=i, value=h)
style_header_row(ws5, 4, 2, 11)

r = 5
for idx, chain in enumerate(comparable_chains):
    cn = chain["num"]
    worst = worst_per_chain.get(cn,{}); best = best_per_chain.get(cn,{})
    if not worst or not best: continue
    savings = worst.get("tco",0) - best.get("tco",0)
    savings_pct = savings / worst.get("tco",1) if worst.get("tco",0) > 0 else 0
    vals = [cn,chain["name"],chain["segment"],chain["priority"],worst.get("option",""),worst.get("tco",0),best.get("option",""),best.get("tco",0),savings,savings_pct]
    for i, v in enumerate(vals, start=2):
        c = ws5.cell(row=r, column=i, value=v)
        if i in (7,9,10): c.number_format = FMT_RUB; c.alignment = align_number()
        elif i == 11: c.number_format = FMT_PCT; c.alignment = align_number()
        elif i == 2: c.alignment = align_number()
        else: c.alignment = align_text()
    if savings > 0:
        ws5.cell(row=r,column=10).fill=SAVINGS_FILL; ws5.cell(row=r,column=10).font=SAVINGS_FONT
        ws5.cell(row=r,column=11).fill=SAVINGS_FILL; ws5.cell(row=r,column=11).font=SAVINGS_FONT
    style_data_row(ws5, r, 2, 11, idx); r += 1

ws5.merge_cells(start_row=r, start_column=2, end_row=r, end_column=8)
ws5.cell(row=r, column=2, value="ИТОГО экономия при выборе выгодных вариантов:").font = font_subheader()
ws5.cell(row=r, column=2).fill = fill_total(); ws5.cell(row=r, column=2).alignment = align_text()
total_savings = sum(worst_per_chain.get(cn,{}).get("tco",0) - best_per_chain.get(cn,{}).get("tco",0) for cn in [c["num"] for c in comparable_chains])
c = ws5.cell(row=r, column=10, value=total_savings)
c.font = Font(name=FONT_NAME, size=12, bold=HEADER_BOLD, color=ACCENT_POSITIVE)
c.fill = SAVINGS_FILL; c.number_format = FMT_RUB; c.alignment = align_number()
total_worst = sum(worst_per_chain.get(cn,{}).get("tco",0) for cn in [c["num"] for c in comparable_chains])
total_pct = total_savings / total_worst if total_worst > 0 else 0
c2 = ws5.cell(row=r, column=11, value=total_pct)
c2.font = Font(name=FONT_NAME, size=12, bold=HEADER_BOLD, color=ACCENT_POSITIVE)
c2.fill = SAVINGS_FILL; c2.number_format = FMT_PCT; c2.alignment = align_number()
style_total_row(ws5, r, 2, 11)

# Графики
r += 3
chart_data_row = r
ws5.cell(row=r, column=2, value="Цепочка"); ws5.cell(row=r, column=3, value="ТСО дорогой вариант (₽)"); ws5.cell(row=r, column=4, value="ТСО выгодный вариант (₽)")
r += 1; chart_first = r
for chain in comparable_chains:
    cn = chain["num"]
    short = f'{cn}. {chain["name"][:30]}'
    ws5.cell(row=r, column=2, value=short)
    ws5.cell(row=r, column=3, value=worst_per_chain.get(cn,{}).get("tco",0)).number_format = FMT_INT
    ws5.cell(row=r, column=4, value=best_per_chain.get(cn,{}).get("tco",0)).number_format = FMT_INT
    r += 1
chart_last = r - 1

chart1 = create_bar_chart(chart_type="col", grouping="clustered", gap_width=80, overlap=10, width=28, height=16)
data_ref = Reference(ws5, min_col=3, max_col=4, min_row=chart_first-1, max_row=chart_last)
cats_ref = Reference(ws5, min_col=2, min_row=chart_first, max_row=chart_last)
chart1.add_data(data_ref, titles_from_data=True); chart1.set_categories(cats_ref)
setup_chart_titles(chart1, title="ТСО за 10 лет: дорогой vs выгодный вариант (₽)", y_title="ТСО (₽)")
apply_chart_colors(chart1, [ACCENT_NEGATIVE, ACCENT_POSITIVE])
ws5.add_chart(chart1, f"F{chart_data_row}")

for col_letter, width in [('B',6),('C',32),('D',14),('E',14),('F',28),('G',20),('H',28),('I',20),('J',16),('K',14)]:
    ws5.column_dimensions[col_letter].width = width

# ===== ЛИСТ 6: Цепочки =====
ws6 = wb.create_sheet("Цепочки")
setup_sheet(ws6, title="СПРАВОЧНИК ТЕХНОЛОГИЧЕСКИХ ЦЕПОЧЕК", last_col=8)
headers_ch = ["№","Технология","Сегмент","Проблема","SKU в цепочке","Ценовой уровень","Приоритет","Ключевое преимущество ТСО"]
for i, h in enumerate(headers_ch, start=2):
    ws6.cell(row=4, column=i, value=h)
style_header_row(ws6, 4, 2, 9)

chain_ref = [
    (1,"Защита металла: грунт + эмаль","Металл","Ржавый металл: заборы, вороты, трубы, гаражи","10 + 97 = 107 SKU","Эконом / Средний","ВЫСОКИЙ","ГФ-021 + ОРЕОЛ Premium служит 5 лет вместо 3 — экономия до 41%"),
    (2,"Быстрая покраска металла без грунта","Металл","Нет времени грунтовать — покрасить ржавый забор за один проход","34 SKU (3в1)","Средний","ВЫСОКИЙ","Paint Lab служит 5 лет vs 3в1 = 4 года — экономия до 30%"),
    (3,"Премиум-покраска металла по RAL","Металл","Нужен точный цвет по RAL для ворот, заборов","44 SKU (TITAN)","Премиум","ВЫСОКИЙ","8 лет службы окупает цену — самый низкий ТСО"),
    (4,"Современная покраска — Paint Lab","Металл","Стильнее классики, без переплаты за TITAN","24 SKU (Paint Lab)","Средний+","СРЕДНИЙ","5 лет службы при средней цене — оптимальное соотношение"),
    (5,"Бетонный пол без пыли и истирания","Полы","Бетонный пол пылит, впитывает масло, крошится","10 + 21 = 31 SKU","Средний","ВЫСОКИЙ","Грунт снижает расход краски в 1.8× — экономия до 35%"),
    (6,"Дерево на улице: забор, терраса","Дерево (экст.)","Забор/терраса сереют, разбухают, трескаются","12 SKU (MONLID + Luxens)","Эконом / Средний","СРЕДНИЙ","Luxens служит 6 лет vs MONLID 4 — экономия 28%"),
    (7,"Паркет и деревянный пол внутри","Дерево (инт.)","Паркет истирается, царапается, тускнеет","6 SKU (AVANTGARDE)","Премиум","НИЗКИЙ","Гибридная технология: 7 лет службы, не желтеет"),
    (8,"Ровный потолок без разводов","Потолки","Краска ложится полосами, видны разводы","1 + 4 = 5 SKU","Эконом","НИЗКИЙ","Грунт предотвращает пятна и снижает расход на 30%"),
    (9,"Стены, которые можно мыть","Интерьер","Стены пачкаются, краска размокает при мытье","1 + 4 = 5 SKU","Средний","НИЗКИЙ","Грунт экономит краску — окупается за 1 цикл"),
    (10,"Покраска фасада дома","Фасад","Фасад потрескался, выцвел","1 + 2 = 3 SKU","Средний","НИЗКИЙ","Без грунта расход в 1.5× выше — переплата до 30%"),
    (11,"Декоративная стена с текстурой","Интерьер / Декор","Гладкая краска скучно, хочется фактуру","1 + 2 = 3 SKU","Средний","НИЗКИЙ","Скрывает дефекты стен — экономия на выравнивании"),
    (12,"Покраска батарей без запаха","Радиаторы","Ржавые батареи, краска желтеет от нагрева","10 + 1 = 11 SKU","Эконом","СРЕДНИЙ","ГФ-021 + эмаль служит 7 лет vs 4 без грунта"),
    (13,"Усиленная защита ржавого металла","Металл (тяж. усл.)","Металл сильно проржавел — 3в1 не справится","10 + 78 = 88 SKU","Средний","ВЫСОКИЙ","ГФ-021 + 3в1 = 10 лет vs 3в1 = 4 — экономия до 42%"),
    (14,"Металлический пол / настил","Металл + Полы","Настил ржавеет и скользит","10 + 21 = 31 SKU","Средний","ВЫСОКИЙ","Без ГФ-021 краска отслаивается за 3 года — переплата 2×"),
]
for idx, (num,name,seg,problem,sku,price_level,priority,tco_adv) in enumerate(chain_ref):
    r = 5 + idx
    vals = [num,name,seg,problem,sku,price_level,priority,tco_adv]
    for i, v in enumerate(vals, start=2):
        c = ws6.cell(row=r, column=i, value=v)
        c.alignment = align_text() if i > 2 else align_number()
    style_data_row(ws6, r, 2, 9, idx)
    prio_cell = ws6.cell(row=r, column=8)
    if priority == "ВЫСОКИЙ": prio_cell.font = Font(name=FONT_NAME, size=11, bold=HEADER_BOLD, color=ACCENT_NEGATIVE)
    elif priority == "СРЕДНИЙ": prio_cell.font = Font(name=FONT_NAME, size=11, bold=HEADER_BOLD, color=ACCENT_WARNING)

# ===== ЛИСТ 7: Сценарии Металл =====
ws7 = wb.create_sheet("Сценарии Металл")
setup_sheet(ws7, title="СЦЕНАРИИ ТСО ДЛЯ МЕТАЛЛА — ДЕТАЛЬНЫЙ РАСЧЁТ", last_col=12)
metal_chain_nums = [1,2,3,4,13,14]
metal_objects = ["Ворота металлические","Гараж металлический","Дверь металлическая","Забор металлический"]

r = 4
ws7.merge_cells(start_row=r, start_column=2, end_row=r, end_column=12)
ws7.cell(row=r, column=2, value="МАТРИЦА ТСО: ЦЕПОЧКА × ОБЪЕКТ (₽ за 10 лет)").font = font_subheader()
ws7.cell(row=r, column=2).fill = SECTION_FILL; r += 1
col_headers = ["Вариант цепочки"] + [f'{obj}\n({OBJECTS[obj]} м²)' for obj in metal_objects]
for i, h in enumerate(col_headers, start=2):
    ws7.cell(row=r, column=i, value=h)
style_header_row(ws7, r, 2, 2+len(col_headers)-1); r += 1

metal_options = []
for chain in TCO_CHAINS:
    if chain["num"] in metal_chain_nums:
        for opt in chain["options"]:
            metal_options.append({"chain_num":chain["num"],"chain_name":chain["name"],"option":opt,"chain":chain})

for idx, mo in enumerate(metal_options):
    opt = mo["option"]; chain = mo["chain"]
    label = f'{mo["chain_num"]}. {opt["name"]}'
    ws7.cell(row=r, column=2, value=label).font = font_body()
    ws7.cell(row=r, column=2).alignment = align_text()
    for oi, obj in enumerate(metal_objects):
        area = OBJECTS[obj]
        res = calc_tco(chain, opt, area, horizon, labor_rates)
        c = ws7.cell(row=r, column=3+oi, value=res["tco"])
        c.number_format = FMT_RUB; c.alignment = align_number()
    style_data_row(ws7, r, 2, 2+len(col_headers)-1, idx); r += 1

# На 1 м²
r += 2
ws7.merge_cells(start_row=r, start_column=2, end_row=r, end_column=12)
ws7.cell(row=r, column=2, value="МАТРИЦА ТСО НА 1 м² (₽/м² за 10 лет)").font = font_subheader()
ws7.cell(row=r, column=2).fill = SECTION_FILL; r += 1
for i, h in enumerate(col_headers, start=2):
    ws7.cell(row=r, column=i, value=h)
style_header_row(ws7, r, 2, 2+len(col_headers)-1); r += 1

for idx, mo in enumerate(metal_options):
    opt = mo["option"]; chain = mo["chain"]
    label = f'{mo["chain_num"]}. {opt["name"]}'
    ws7.cell(row=r, column=2, value=label).font = font_body()
    ws7.cell(row=r, column=2).alignment = align_text()
    area = OBJECTS[metal_objects[0]]
    res = calc_tco(chain, opt, area, horizon, labor_rates)
    tco_per_m2 = round(res["tco"] / area, 0)
    for oi, obj in enumerate(metal_objects):
        c = ws7.cell(row=r, column=3+oi, value=tco_per_m2)
        c.number_format = FMT_RUB; c.alignment = align_number()
    style_data_row(ws7, r, 2, 2+len(col_headers)-1, idx); r += 1

# График
r += 2
chart_data_row = r
ws7.cell(row=r, column=2, value="Вариант"); ws7.cell(row=r, column=3, value="Ворота 16м²"); ws7.cell(row=r, column=4, value="Гараж 54м²"); ws7.cell(row=r, column=5, value="Дверь 4.4м²")
r += 1
for idx, mo in enumerate(metal_options):
    opt = mo["option"]; chain = mo["chain"]
    label = f'{mo["chain_num"]}. {opt["name"][:25]}'
    areas = [OBJECTS[obj] for obj in metal_objects[:3]]
    results = [calc_tco(chain, opt, a, horizon, labor_rates) for a in areas]
    ws7.cell(row=r, column=2, value=label)
    for ai, res in enumerate(results):
        ws7.cell(row=r, column=3+ai, value=res["tco"]).number_format = FMT_INT
    r += 1
chart_data_end = r - 1
chart3 = create_bar_chart(chart_type="col", grouping="clustered", gap_width=80, overlap=10, width=28, height=16)
data_ref3 = Reference(ws7, min_col=3, max_col=5, min_row=chart_data_row, max_row=chart_data_end)
cats_ref3 = Reference(ws7, min_col=2, min_row=chart_data_row+1, max_row=chart_data_end)
chart3.add_data(data_ref3, titles_from_data=True); chart3.set_categories(cats_ref3)
setup_chart_titles(chart3, title="ТСО за 10 лет: сценарии для металла (₽)", y_title="ТСО (₽)")
apply_chart_colors(chart3, [PRIMARY, ACCENT_POSITIVE, ACCENT_WARNING, ACCENT_NEGATIVE, NEUTRAL_600])
ws7.add_chart(chart3, f"G{chart_data_row}")
for col_letter, width in [('B',40),('C',18),('D',18),('E',18),('F',18)]:
    ws7.column_dimensions[col_letter].width = width

# ===== ЛИСТ 8: Методология =====
ws8 = wb.create_sheet("Методология")
setup_sheet(ws8, title="МЕТОДОЛОГИЯ РАСЧЁТА ТСО", last_col=8)
methodology = [
    ("Формула ТСО", "ТСО = Стоимость 1-го цикла + (Кол-во циклов - 1) × Стоимость перекраски"),
    ("Кол-во циклов", "Цикл = ОДИН полный процесс покраски (грунтовка + финишное покрытие, или только 3в1).\nКол-во циклов = ОКРВВЕРХ(Горизонт / Срок службы).\nПример: горизонт 10 лет, срок службы 3 года → 4 цикла (0, 3, 6, 9 год)."),
    ("Стоимость 1-го цикла", "Стоимость = (Материал ₽/м² + Работа ₽/м²) × Площадь.\nМатериал = сумма(Расход г/м² × Слои × Цена за кг/л / 1000) по всем шагам.\nРабота = сумма(Ставка ₽/м²) по всем операциям."),
    ("Стоимость перекраски", "КЛЮЧЕВОЙ МОМЕНТ: перекраска ДЕШЕВЛЕ первоначальной покраски!\n- 2-шаговые цепочки (грунт + краска): перекраска = только финишное покрытие + лёгкая подготовка ≈ 55-65% от первоначальной\n- 1-шаговые цепочки (3в1, TITAN): перекраска = шлифовка + новое покрытие ≈ 100% стоимости\n- Без грунта: каждая перекраска = полная стоимость = 100%"),
    ("Срок службы", "Оценка на основе ТТД производителей и опыта эксплуатации.\nВнешние поверхности (металл, фасад) — агрессивная среда, срок ниже.\nВнутренние поверхности (потолки, стены) — мягкие условия, срок выше.\nБез грунта — сокращённый срок (плохая адгезия, неравномерное впитывание)."),
    ("Допущения", "1. Перекраска = стоимость перекраски (не полная первоначальная) — грунт сохраняется.\n2. Цены — ориентировочные розничные Лемана ПРО на 2025-2026.\n3. Расход — по ТТД производителей, с запасом ~10%.\n4. Работы выполняются нанятыми мастерами (не DIY).\n5. Ставки на работу — средние по РФ на 2025-2026.\n6. «Без грунта» варианты имеют расход в 1.5-1.8× выше."),
    ("Почему ТСО важнее цены за банку", "Краска за 190 ₽/кг, которую нужно перекрашивать каждые 2 года, за 10 лет обходится ДОРОЖЕ, чем краска за 1100 ₽/кг на 8 лет. ТСО учитывает ВСЕ затраты: материалы + работа × количество циклов."),
    ("Логика колонки «Экономия»", "Экономия = ТСО(самый дорогой вариант) - ТСО(этот вариант).\nПоложительное зелёное значение = этот вариант дешевле самого дорогого.\n0 = это самый дорогой вариант.\nЛучший вариант всегда показывает наибольшую экономию (зелёный)."),
]
r = 4
for title, text in methodology:
    ws8.merge_cells(start_row=r, start_column=2, end_row=r, end_column=8)
    ws8.cell(row=r, column=2, value=title).font = font_subheader()
    ws8.cell(row=r, column=2).fill = SECTION_FILL; r += 1
    ws8.merge_cells(start_row=r, start_column=2, end_row=r+1, end_column=8)
    c = ws8.cell(row=r, column=2, value=text)
    c.font = font_body(); c.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
    ws8.row_dimensions[r].height = 45; ws8.row_dimensions[r+1].height = 20; r += 3
ws8.column_dimensions["B"].width = 60

# ===== СОХРАНЕНИЕ =====
output_path = "/home/z/my-project/download/TCO_Калькулятор_ЛеманаПРО_Эмпилс.xlsx"
wb.save(output_path)
print(f"OK: {output_path}")
