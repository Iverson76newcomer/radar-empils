#!/usr/bin/env python3
"""
Анализ результатов анкетирования ТСО.
Читает заполненный шаблон Excel, создаёт графики и текстовый отчёт.

Использование:
    python3 analyze_tso.py [путь_к_файлу.xlsx]

Если путь не указан — берёт последний файл шаблона в /download/
"""
import sys
import os
import glob
from datetime import datetime

import matplotlib
matplotlib.use("Agg")
import matplotlib.font_manager as fm
fm.fontManager.addfont('/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc')
fm.fontManager.addfont('/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf')
import matplotlib.pyplot as plt
plt.rcParams['font.sans-serif'] = ['WenQuanYi Zen Hei', 'DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False

import numpy as np
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.utils import get_column_letter

# ═══════════════════════════════════════════════════════
# КОНФИГУРАЦИЯ
# ═══════════════════════════════════════════════════════
OUTPUT_DIR = "/home/z/my-project/download"
CHARTS_DIR = os.path.join(OUTPUT_DIR, "tso_charts")
os.makedirs(CHARTS_DIR, exist_ok=True)

# Цвета
DARK_BLUE = "1B3A5C"
MED_BLUE = "2E75B6"
LIGHT_BLUE = "5B9BD5"
ACCENT_ORANGE = "ED7D31"
ACCENT_GREEN = "70AD47"
ACCENT_RED = "C00000"
PALETTE = ["#1B3A5C", "#2E75B6", "#5B9BD5", "#ED7D31", "#70AD47", "#7030A0", "#C00000", "#FFC000"]

# ═══════════════════════════════════════════════════════
# ОПРЕДЕЛЕНИЯ ВОПРОСОВ
# ═══════════════════════════════════════════════════════
SINGLE_QUESTIONS = {
    "q1": {"title": "В1. Должность", "short": "Должность"},
    "q2": {"title": "В1. Опыт работы в ЛКМ", "short": "Опыт работы"},
    "q3": {"title": "В1. Формат продаж", "short": "Формат продаж"},
    "q4": {"title": "В2. Знакомство с ТСО", "short": "Знание ТСО"},
    "q7": {"title": "В3. Как рекомендуете ЛКМ", "short": "Метод рекомендации"},
    "q8": {"title": "В3. Лучший аргумент для премиума", "short": "Аргумент для премиума"},
    "q11": {"title": "В4. Кто восприимчив к ТСО", "short": "Целевая аудитория ТСО"},
    "q15": {"title": "В5. Срок внедрения", "short": "Готовность внедрить"},
}

SCALE_QUESTIONS = {
    "q10": {"title": "В4. ТСО повлияет на решение (0–10)", "short": "Влияние ТСО на решение"},
    "q17": {"title": "В6. Перспективность ТСО (0–10)", "short": "Перспективность ТСО"},
}

MULTI_QUESTIONS = {
    "q5": {
        "title": "В2. Составляющие совокупной стоимости",
        "short": "Составляющие ТСО",
        "sub_items": [
            "Цена банки", "Грунтовка/подготовка", "Количество слоёв",
            "Время высыхания", "Стоимость рабочей силы", "Срок службы покрытия",
            "Расход на инструмент", "Затраты на утилизацию",
        ],
    },
    "q9": {
        "title": "В3. Фразы покупателя «по цене»",
        "short": "Фразы «по цене»",
        "sub_items": [
            "«Чем дороже обычной?»", "«Нет подешевле?»",
            "«Нужно покрасить недорого»", "«Покрашу сам — не за дорогую»",
            "«Давайте подешевле — одинаково»", "Другое",
        ],
    },
    "q12": {
        "title": "В4. Продукты для продажи через ТСО",
        "short": "Продукты для ТСО",
        "sub_items": [
            "Грунт-эмали 3в1", "Премиум-эмали (замена ПФ-115)",
            "ВД краски повыш. стойкости", "Фасадные с повыш. укрывист.",
            "Лаки/пропитки премиум", "Другое",
        ],
    },
    "q13": {
        "title": "В5. Барьеры для продажи через ТСО",
        "short": "Барьеры ТСО",
        "sub_items": [
            "Нет наглядного расчёта", "Не хватает знаний",
            "Покупателю не интересны расчёты", "Нет времени на объяснения",
            "Не верю, что заплатит больше", "Конкуренты дешевле", "Другое",
        ],
    },
    "q14": {
        "title": "В5. Инструменты для продажи через ТСО",
        "short": "Инструменты ТСО",
        "sub_items": [
            "Калькулятор ТСО", "Печатная шпаргалка",
            "Образцы покрытий", "Обучение: ТСО за 2 мин",
            "Видео/презентация", "Сравнит. таблица vs конкуренты",
            "Мотивационная программа", "Другое",
        ],
    },
    "q16": {
        "title": "В6. Ожидания от производителя",
        "short": "Ожидания от производителя",
        "sub_items": [
            "Готовые материалы", "Обучение персонала ТСО",
            "Политика ценообразования", "Маркетинговая поддержка",
            "Сопровождение на месте", "Другое",
        ],
    },
}


# ═══════════════════════════════════════════════════════
# ЧТЕНИЕ ДАННЫХ
# ═══════════════════════════════════════════════════════
def read_data(filepath):
    """Читает заполненный шаблон и возвращает заголовки + строки данных."""
    wb = openpyxl.load_workbook(filepath, data_only=True)
    ws = wb["Ввод данных"]
    
    headers = []
    for col in range(1, ws.max_column + 1):
        headers.append(ws.cell(row=1, column=col).value)
    
    rows = []
    for row in range(4, ws.max_row + 1):
        row_data = {}
        has_data = False
        for col in range(1, ws.max_column + 1):
            val = ws.cell(row=row, column=col).value
            key = headers[col - 1] if col - 1 < len(headers) else f"col_{col}"
            row_data[key] = val
            if val is not None and val != "" and key != "№ анкеты":
                has_data = True
        if has_data:
            rows.append(row_data)
    
    return headers, rows


def get_col_value(row_data, col_name):
    """Получает значение из строки по названию столбца (частичное совпадение)."""
    for key, val in row_data.items():
        if key and col_name in str(key):
            return val
    return None


# ═══════════════════════════════════════════════════════
# АНАЛИЗ
# ═══════════════════════════════════════════════════════
def analyze_single_choice(rows, q_prefix):
    """Подсчёт ответов для вопроса с одним выбором."""
    counts = {}
    for row in rows:
        val = get_col_value(row, q_prefix)
        if val is not None and val != "":
            val_str = str(val).strip()
            counts[val_str] = counts.get(val_str, 0) + 1
    return counts


def analyze_multi_choice(rows, q_prefix, sub_count):
    """Подсчёт ответов для вопроса с множественным выбором."""
    counts = {}
    for i in range(1, sub_count + 1):
        col_name = f"{q_prefix}_{i}"
        count = 0
        for row in rows:
            val = get_col_value(row, col_name)
            if val is not None and val != "":
                try:
                    if float(val) == 1:
                        count += 1
                except (ValueError, TypeError):
                    if str(val).strip() in ("1", "да", "Да", "+"):
                        count += 1
        counts[f"Вариант {i}"] = count
    return counts


def analyze_scale(rows, q_prefix):
    """Статистика для шкального вопроса."""
    values = []
    for row in rows:
        val = get_col_value(row, q_prefix)
        if val is not None and val != "":
            try:
                values.append(float(val))
            except (ValueError, TypeError):
                pass
    if not values:
        return {"count": 0, "mean": 0, "median": 0, "min": 0, "max": 0, "std": 0, "dist": {}}
    
    # Распределение
    dist = {}
    for v in values:
        bucket = int(v)
        dist[bucket] = dist.get(bucket, 0) + 1
    
    return {
        "count": len(values),
        "mean": np.mean(values),
        "median": np.median(values),
        "min": int(min(values)),
        "max": int(max(values)),
        "std": np.std(values),
        "dist": dist,
    }


# ═══════════════════════════════════════════════════════
# ВИЗУАЛИЗАЦИЯ
# ═══════════════════════════════════════════════════════
def chart_single_choice(data, title, filename, max_label_len=35):
    """Горизонтальная столбчатая диаграмма для одного выбора."""
    if not data:
        return
    
    labels = list(data.keys())
    values = list(data.values())
    total = sum(values)
    
    # Сокращение длинных подписей
    short_labels = []
    for l in labels:
        if len(l) > max_label_len:
            short_labels.append(l[:max_label_len - 3] + "...")
        else:
            short_labels.append(l)
    
    fig, ax = plt.subplots(figsize=(10, max(4, len(labels) * 0.7 + 1.5)))
    y_pos = np.arange(len(labels))
    bars = ax.barh(y_pos, values, color=PALETTE[:len(labels)], height=0.6, edgecolor="white")
    
    ax.set_yticks(y_pos)
    ax.set_yticklabels(short_labels, fontsize=9)
    ax.invert_yaxis()
    ax.set_xlabel("Количество ответов", fontsize=10)
    ax.set_title(title, fontsize=12, fontweight="bold", pad=12)
    
    # Подписи значений
    for bar, val in zip(bars, values):
        pct = val / total * 100
        ax.text(bar.get_width() + 0.3, bar.get_y() + bar.get_height() / 2,
                f"{val} ({pct:.0f}%)", va="center", fontsize=9)
    
    ax.set_xlim(0, max(values) * 1.35)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    
    plt.tight_layout()
    path = os.path.join(CHARTS_DIR, filename)
    plt.savefig(path, dpi=150, bbox_inches="tight")
    plt.close()
    print(f"  📊 {filename}")


def chart_multi_choice(data, title, filename, sub_items):
    """Горизонтальная столбчатая для множественного выбора."""
    if not sub_items:
        return
    
    labels = sub_items[:len(data)]
    values = list(data.values())
    
    fig, ax = plt.subplots(figsize=(10, max(4, len(labels) * 0.65 + 1.5)))
    y_pos = np.arange(len(labels))
    bars = ax.barh(y_pos, values, color=PALETTE[:len(labels)], height=0.6, edgecolor="white")
    
    ax.set_yticks(y_pos)
    ax.set_yticklabels(labels, fontsize=9)
    ax.invert_yaxis()
    ax.set_xlabel("Количество ответов", fontsize=10)
    ax.set_title(title, fontsize=12, fontweight="bold", pad=12)
    
    for bar, val in zip(bars, values):
        if val > 0:
            ax.text(bar.get_width() + 0.2, bar.get_y() + bar.get_height() / 2,
                    str(val), va="center", fontsize=9)
    
    ax.set_xlim(0, max(values) * 1.3 if max(values) > 0 else 1)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    
    plt.tight_layout()
    path = os.path.join(CHARTS_DIR, filename)
    plt.savefig(path, dpi=150, bbox_inches="tight")
    plt.close()
    print(f"  📊 {filename}")


def chart_scale(data, title, filename):
    """Гистограмма + статистика для шкального вопроса."""
    if data["count"] == 0:
        return
    
    dist = data["dist"]
    all_vals = list(range(11))
    counts = [dist.get(v, 0) for v in all_vals]
    
    fig, ax = plt.subplots(figsize=(10, 5))
    bars = ax.bar(all_vals, counts, color="#2E75B6", edgecolor="white", width=0.7)
    
    # Среднее
    ax.axvline(data["mean"], color="#ED7D31", linewidth=2, linestyle="--",
               label=f"Среднее: {data['mean']:.1f}")
    ax.axvline(data["median"], color="#70AD47", linewidth=2, linestyle=":",
               label=f"Медиана: {data['median']:.1f}")
    
    # Подписи
    for x, cnt in zip(all_vals, counts):
        if cnt > 0:
            ax.text(x, cnt + 0.1, str(cnt), ha="center", fontsize=9, fontweight="bold")
    
    ax.set_xticks(all_vals)
    ax.set_xlabel("Оценка", fontsize=10)
    ax.set_ylabel("Количество ответов", fontsize=10)
    ax.set_title(title, fontsize=12, fontweight="bold", pad=12)
    ax.legend(fontsize=10)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    
    plt.tight_layout()
    path = os.path.join(CHARTS_DIR, filename)
    plt.savefig(path, dpi=150, bbox_inches="tight")
    plt.close()
    print(f"  📊 {filename}")


def chart_tso_awareness_combo(rows):
    """Комбинированный график: знание ТСО + влияние + перспективность."""
    q4_data = analyze_single_choice(rows, "Знакомство") or analyze_single_choice(rows, "В2")
    q10_data = analyze_scale(rows, "ТСО повлияет") or {'count': 0}
    if q10_data['count'] == 0:
        q10_data = analyze_scale(rows, "В4")
    q17_data = analyze_scale(rows, "Перспективность") or {'count': 0}
    if q17_data['count'] == 0:
        q17_data = analyze_scale(rows, "В6")
    
    fig, axes = plt.subplots(1, 3, figsize=(16, 5))
    
    # Pie: Знание ТСО
    if q4_data:
        labels = list(q4_data.keys())
        # Сокращение
        short = []
        for l in labels:
            if "хорошо понимаю" in l:
                short.append("Знаю и применяю")
            elif "Слышал" in l or "Слышал(а)" in l:
                short.append("Слышал, не применял")
            elif "интуитивно" in l:
                short.append("Интуитивно считаю")
            elif "Впервые" in l or "впервые" in l:
                short.append("Впервые слышу")
            else:
                short.append(l[:25])
        
        vals = list(q4_data.values())
        colors = ["#70AD47", "#5B9BD5", "#ED7D31", "#C00000"][:len(vals)]
        wedges, texts, autotexts = axes[0].pie(vals, labels=short, colors=colors, autopct="%1.0f%%",
                                                  textprops={"fontsize": 8}, startangle=90)
        axes[0].set_title("Знакомство с ТСО", fontsize=11, fontweight="bold")
    
    # Bar: Влияние ТСО
    if q10_data["count"] > 0:
        dist = q10_data["dist"]
        x_vals = list(range(11))
        y_vals = [dist.get(v, 0) for v in x_vals]
        axes[1].bar(x_vals, y_vals, color="#2E75B6", edgecolor="white", width=0.7)
        axes[1].axvline(q10_data["mean"], color="#ED7D31", linewidth=2, linestyle="--")
        axes[1].set_xticks(x_vals)
        axes[1].set_title(f"Влияние ТСО на решение\n(ср. {q10_data['mean']:.1f})", fontsize=11, fontweight="bold")
        axes[1].set_xlabel("Оценка 0–10")
        axes[1].set_ylabel("Кол-во")
    
    # Bar: Перспективность
    if q17_data["count"] > 0:
        dist = q17_data["dist"]
        x_vals = list(range(11))
        y_vals = [dist.get(v, 0) for v in x_vals]
        axes[2].bar(x_vals, y_vals, color="#70AD47", edgecolor="white", width=0.7)
        axes[2].axvline(q17_data["mean"], color="#ED7D31", linewidth=2, linestyle="--")
        axes[2].set_xticks(x_vals)
        axes[2].set_title(f"Перспективность ТСО\n(ср. {q17_data['mean']:.1f})", fontsize=11, fontweight="bold")
        axes[2].set_xlabel("Оценка 0–10")
        axes[2].set_ylabel("Кол-во")
    
    plt.suptitle("СВОДНЫЙ ДАШБОРД: Готовность к ТСО", fontsize=14, fontweight="bold", y=1.02)
    plt.tight_layout()
    path = os.path.join(CHARTS_DIR, "00_сводный_дашборд_ТСО.png")
    plt.savefig(path, dpi=150, bbox_inches="tight")
    plt.close()
    print(f"  📊 00_сводный_дашборд_ТСО.png")


# ═══════════════════════════════════════════════════════
# ГЕНЕРАЦИЯ ОТЧЁТА В EXCEL
# ═══════════════════════════════════════════════════════
def generate_report_excel(rows, headers, output_path):
    """Создаёт Excel-файл с аналитикой и встроенными графиками."""
    wb = openpyxl.Workbook()
    
    # Стили
    title_font = Font(name="Calibri", bold=True, size=14, color=DARK_BLUE)
    section_font = Font(name="Calibri", bold=True, size=12, color=MED_BLUE)
    header_font = Font(name="Calibri", bold=True, size=10, color="FFFFFF")
    header_fill = PatternFill(start_color=DARK_BLUE, end_color=DARK_BLUE, fill_type="solid")
    data_font = Font(name="Calibri", size=10)
    pct_font = Font(name="Calibri", size=10, bold=True, color=ACCENT_ORANGE)
    border = Border(
        left=Side(style="thin", color="BFBFBF"),
        right=Side(style="thin", color="BFBFBF"),
        top=Side(style="thin", color="BFBFBF"),
        bottom=Side(style="thin", color="BFBFBF"),
    )
    
    # ─── Лист: Сводка ───
    ws = wb.active
    ws.title = "Сводка"
    
    r = 1
    ws.cell(row=r, column=1, value="АНАЛИЗ АНКЕТ ТСО").font = title_font
    r += 1
    ws.cell(row=r, column=1, value=f"Дата анализа: {datetime.now().strftime('%d.%m.%Y %H:%M')}").font = data_font
    r += 1
    ws.cell(row=r, column=1, value=f"Количество анкет: {len(rows)}").font = Font(name="Calibri", bold=True, size=11)
    r += 2
    
    # ─── Вопросы с одним выбором ───
    ws.cell(row=r, column=1, value="ВОПРОСЫ С ОДНИМ ВЫБОРОМ").font = section_font
    r += 1
    
    for q_key, q_info in SINGLE_QUESTIONS.items():
        # Поиск по ключевой части заголовка колонки
        search_str = q_info["title"].split(". ", 1)[-1] if ". " in q_info["title"] else q_info["title"]
        block_prefix = q_info["title"].split(".")[0]
        data = analyze_single_choice(rows, search_str)
        if not data:
            data = analyze_single_choice(rows, block_prefix)
        
        ws.cell(row=r, column=1, value=q_info["title"]).font = Font(name="Calibri", bold=True, size=10)
        r += 1
        
        # Заголовки таблицы
        for col, h in enumerate(["Вариант ответа", "Кол-во", "%"], 1):
            cell = ws.cell(row=r, column=col, value=h)
            cell.font = header_font
            cell.fill = header_fill
            cell.border = border
        r += 1
        
        total = sum(data.values()) if data else 0
        sorted_items = sorted(data.items(), key=lambda x: x[1], reverse=True) if data else []
        for label, count in sorted_items:
            ws.cell(row=r, column=1, value=label).font = data_font
            ws.cell(row=r, column=1).border = border
            ws.cell(row=r, column=2, value=count).font = data_font
            ws.cell(row=r, column=2).border = border
            ws.cell(row=r, column=2).alignment = Alignment(horizontal="center")
            pct = count / total * 100 if total > 0 else 0
            ws.cell(row=r, column=3, value=f"{pct:.1f}%").font = pct_font
            ws.cell(row=r, column=3).border = border
            ws.cell(row=r, column=3).alignment = Alignment(horizontal="center")
            r += 1
        r += 1
    
    # ─── Шкальные вопросы ───
    ws.cell(row=r, column=1, value="ШКАЛЬНЫЕ ВОПРОСЫ (0–10)").font = section_font
    r += 1
    
    for q_key, q_info in SCALE_QUESTIONS.items():
        search_str = q_info["title"].split(". ", 1)[-1] if ". " in q_info["title"] else q_info["title"]
        block_prefix = q_info["title"].split(".")[0]
        data = analyze_scale(rows, search_str)
        if data["count"] == 0:
            data = analyze_scale(rows, block_prefix)
        ws.cell(row=r, column=1, value=q_info["title"]).font = Font(name="Calibri", bold=True, size=10)
        r += 1
        
        stats = [
            ("Количество ответов", data["count"]),
            ("Среднее", f"{data['mean']:.1f}"),
            ("Медиана", f"{data['median']:.1f}"),
            ("Минимум", data["min"]),
            ("Максимум", data["max"]),
            ("Стд. отклонение", f"{data['std']:.1f}"),
        ]
        for label, val in stats:
            ws.cell(row=r, column=1, value=label).font = data_font
            ws.cell(row=r, column=1).border = border
            ws.cell(row=r, column=2, value=val).font = Font(name="Calibri", bold=True, size=10)
            ws.cell(row=r, column=2).border = border
            ws.cell(row=r, column=2).alignment = Alignment(horizontal="center")
            r += 1
        r += 1
    
    # ─── Вопросы с множественным выбором ───
    ws.cell(row=r, column=1, value="ВОПРОСЫ С МНОЖЕСТВЕННЫМ ВЫБОРОМ").font = section_font
    r += 1
    
    for q_key, q_info in MULTI_QUESTIONS.items():
        sub_count = len(q_info["sub_items"])
        data = analyze_multi_choice(rows, q_key, sub_count)
        
        ws.cell(row=r, column=1, value=q_info["title"]).font = Font(name="Calibri", bold=True, size=10)
        r += 1
        
        for col, h in enumerate(["Вариант ответа", "Кол-во", "% от респондентов"], 1):
            cell = ws.cell(row=r, column=col, value=h)
            cell.font = header_font
            cell.fill = header_fill
            cell.border = border
        r += 1
        
        total_resp = len(rows)
        for i, (label_key, count) in enumerate(data.items()):
            label = q_info["sub_items"][i] if i < len(q_info["sub_items"]) else label_key
            ws.cell(row=r, column=1, value=label).font = data_font
            ws.cell(row=r, column=1).border = border
            ws.cell(row=r, column=2, value=count).font = data_font
            ws.cell(row=r, column=2).border = border
            ws.cell(row=r, column=2).alignment = Alignment(horizontal="center")
            pct = count / total_resp * 100 if total_resp > 0 else 0
            ws.cell(row=r, column=3, value=f"{pct:.1f}%").font = pct_font
            ws.cell(row=r, column=3).border = border
            ws.cell(row=r, column=3).alignment = Alignment(horizontal="center")
            r += 1
        r += 1
    
    ws.column_dimensions["A"].width = 55
    ws.column_dimensions["B"].width = 15
    ws.column_dimensions["C"].width = 20
    
    # ─── Лист: Открытые ответы ───
    ws_open = wb.create_sheet("Открытые ответы")
    r = 1
    ws_open.cell(row=r, column=1, value="ОТКРЫТЫЕ ВОПРОСЫ").font = title_font
    r += 2
    
    open_questions = [
        ("q6", "В2. Пример ТСО из практики"),
        ("q1_other", "В1. Должность — другое"),
        ("q7_other", "В3. Метод рекомендации — другое"),
        ("q8_other", "В3. Аргумент для премиума — другое"),
    ]
    
    for q_key, title in open_questions:
        ws_open.cell(row=r, column=1, value=title).font = section_font
        r += 1
        for row in rows:
            val = get_col_value(row, q_key)
            if val is not None and str(val).strip() not in ("", "нет", "Нет"):
                ws_open.cell(row=r, column=1, value=str(val)).font = data_font
                r += 1
        r += 2
    
    ws_open.column_dimensions["A"].width = 80
    
    # ─── Лист: Интерпретация ───
    ws_interp = wb.create_sheet("Интерпретация")
    r = 1
    ws_interp.cell(row=r, column=1, value="ИНТЕРПРЕТАЦИЯ РЕЗУЛЬТАТОВ").font = title_font
    r += 2
    
    # Авто-интерпретация
    interpretations = []
    
    # Знание ТСО
    q4_data = analyze_single_choice(rows, "Знакомство") or analyze_single_choice(rows, "В2")
    if q4_data:
        total = sum(q4_data.values())
        know_and_use = sum(v for k, v in q4_data.items() if "хорошо понимаю" in k or "использую" in k)
        first_time = sum(v for k, v in q4_data.items() if "Впервые" in k or "впервые" in k)
        know_pct = know_and_use / total * 100 if total else 0
        first_pct = first_time / total * 100 if total else 0
        interpretations.append(
            f"ЗНАНИЕ ТСО: {know_pct:.0f}% хорошо понимают и применяют ТСО, "
            f"{first_pct:.0f}% слышат впервые. "
            f"{'Высокий уровень осведомлённости — можно переходить к практическим инструментам.' if know_pct > 40 else 'Низкая осведомлённость — необходимо начинать с базового обучения.'}"
        )
    
    # Влияние ТСО
    q10_data = analyze_scale(rows, "ТСО повлияет") or {'count': 0}
    if q10_data['count'] == 0:
        q10_data = analyze_scale(rows, "В4")
    if q10_data["count"] > 0:
        avg = q10_data["mean"]
        interpretations.append(
            f"ВЛИЯНИЕ ТСО НА РЕШЕНИЕ ПОКУПАТЕЛЯ: средняя оценка {avg:.1f}/10. "
            f"{'Высокая вера в эффективность аргумента — сотрудники готовы применять.' if avg >= 7 else 'Умеренная вера — нужны кейсы и доказательства.' if avg >= 4 else 'Низкая вера — требуется серьёзная работа по убеждению самих продавцов.'}"
        )
    
    # Перспективность
    q17_data = analyze_scale(rows, "Перспективность") or {'count': 0}
    if q17_data['count'] == 0:
        q17_data = analyze_scale(rows, "В6")
    if q17_data["count"] > 0:
        avg = q17_data["mean"]
        interpretations.append(
            f"ПЕРСПЕКТИВНОСТЬ ТСО: средняя оценка {avg:.1f}/10. "
            f"{'Респонденты видят высокий потенциал — можно инвестировать в инструменты.' if avg >= 7 else 'Умеренные ожидания — нужен пилотный проект для доказательства эффективности.' if avg >= 4 else 'Скептицизм — начните с малого пилота и соберите успешные кейсы.'}"
        )
    
    # Барьеры
    q13_data = analyze_multi_choice(rows, "q13", 7)
    if q13_data:
        top_barrier = max(q13_data, key=q13_data.get)
        top_val = q13_data[top_barrier]
        interpretations.append(
            f"ГЛАВНЫЙ БАРЬЕР: «{top_barrier}» ({top_val} ответов). "
            f"Фокус усилий на устранение этого барьера даст наибольший эффект."
        )
    
    # Инструменты
    q14_data = analyze_multi_choice(rows, "q14", 8)
    if q14_data:
        sorted_tools = sorted(q14_data.items(), key=lambda x: x[1], reverse=True)
        top3 = [f"«{k}» ({v})" for k, v in sorted_tools[:3] if v > 0]
        if top3:
            interpretations.append(
                f"ТОП-3 ЗАПРАШИВАЕМЫХ ИНСТРУМЕНТА: {', '.join(top3)}. "
                f"Приоритизируйте разработку именно этих инструментов."
            )
    
    # Готовность внедрить
    q15_data = analyze_single_choice(rows, "Срок внедрения") or analyze_single_choice(rows, "В5")
    if q15_data:
        total = sum(q15_data.values())
        immediately = sum(v for k, v in q15_data.items() if "Сразу" in k or "сразу" in k)
        imp_pct = immediately / total * 100 if total else 0
        interpretations.append(
            f"ГОТОВНОСТЬ ВНЕДРИТЬ: {imp_pct:.0f}% готовы внедрить сразу. "
            f"{'Отличная готовность к изменениям!' if imp_pct > 50 else 'Требуется подготовительная работа перед внедрением.' if imp_pct > 25 else 'Значительное сопротивление — начните с добровольных амбассадоров.'}"
        )
    
    for interp in interpretations:
        ws_interp.cell(row=r, column=1, value=interp).font = data_font
        ws_interp.cell(row=r, column=1).alignment = Alignment(wrap_text=True)
        r += 2
    
    # Рекомендации
    r += 1
    ws_interp.cell(row=r, column=1, value="РЕКОМЕНДАЦИИ").font = section_font
    r += 1
    recommendations = [
        "1. Начните с разработки калькулятора ТСО — это самый запрашиваемый инструмент.",
        "2. Проведите обучение «ТСО за 2 минуты» для продавцов — второй по спросу инструмент.",
        "3. Подготовьте сравнительную таблицу ТСО: продукты Эмпилс vs конкуренты.",
        "4. Запустите пилот с добровольными участниками — они станут амбассадорами методики.",
        "5. Собирайте успешные кейсы продажи через ТСО для формирования доказательной базы.",
        "6. Используйте мотивационную программу для стимулирования продаж через ТСО.",
    ]
    for rec in recommendations:
        ws_interp.cell(row=r, column=1, value=rec).font = data_font
        ws_interp.cell(row=r, column=1).alignment = Alignment(wrap_text=True)
        r += 1
    
    ws_interp.column_dimensions["A"].width = 100
    
    wb.save(output_path)
    print(f"  📋 Отчёт: {output_path}")


# ═══════════════════════════════════════════════════════
# ГЛАВНЫЙ СКРИПТ
# ═══════════════════════════════════════════════════════
def main():
    # Определяем файл для анализа
    if len(sys.argv) > 1:
        filepath = sys.argv[1]
    else:
        # Ищем последний шаблон
        files = glob.glob(os.path.join(OUTPUT_DIR, "Шаблон_ввода_анкет_ТСО*.xlsx"))
        if not files:
            print("❌ Файл с данными не найден. Укажите путь: python3 analyze_tso.py <файл.xlsx>")
            sys.exit(1)
        filepath = max(files, key=os.path.getmtime)
    
    print(f"📂 Файл: {filepath}")
    
    # Чтение данных
    headers, rows = read_data(filepath)
    n = len(rows)
    print(f"✅ Прочитано анкет: {n}")
    
    if n == 0:
        print("⚠️  Нет заполненных анкет. Заполните шаблон и запустите скрипт снова.")
        sys.exit(0)
    
    print(f"\n📊 Генерация графиков...")
    
    # Сводный дашборд
    chart_tso_awareness_combo(rows)
    
    # Вопросы с одним выбором — ищем по заголовку (title), а не q_key
    for q_key, q_info in SINGLE_QUESTIONS.items():
        # Извлекаем уникальную часть заголовка для поиска в столбцах
        search_str = q_info["title"].split(". ", 1)[-1] if ". " in q_info["title"] else q_info["title"]
        # Также пробуем короткий префикс блока (В1, В2 и т.д.)
        block_prefix = q_info["title"].split(".")[0]
        data = analyze_single_choice(rows, search_str)
        if not data:
            data = analyze_single_choice(rows, block_prefix)
        if data:
            chart_single_choice(data, q_info["title"], f"{q_key}_{q_info['short'].replace(' ', '_')}.png")
    
    # Шкальные вопросы
    for q_key, q_info in SCALE_QUESTIONS.items():
        search_str = q_info["title"].split(". ", 1)[-1] if ". " in q_info["title"] else q_info["title"]
        block_prefix = q_info["title"].split(".")[0]
        data = analyze_scale(rows, search_str)
        if data["count"] == 0:
            data = analyze_scale(rows, block_prefix)
        if data["count"] > 0:
            chart_scale(data, q_info["title"], f"{q_key}_{q_info['short'].replace(' ', '_')}.png")
    
    # Вопросы с множественным выбором
    for q_key, q_info in MULTI_QUESTIONS.items():
        sub_count = len(q_info["sub_items"])
        data = analyze_multi_choice(rows, q_key, sub_count)
        chart_multi_choice(data, q_info["title"], f"{q_key}_{q_info['short'].replace(' ', '_')}.png", q_info["sub_items"])
    
    # Генерация Excel-отчёта
    print(f"\n📋 Генерация отчёта...")
    report_path = os.path.join(OUTPUT_DIR, f"Отчёт_анализ_анкет_ТСО_{datetime.now().strftime('%Y%m%d')}.xlsx")
    generate_report_excel(rows, headers, report_path)
    
    print(f"\n✅ АНАЛИЗ ЗАВЕРШЁН!")
    print(f"   Анкет обработано: {n}")
    print(f"   Графики: {CHARTS_DIR}/")
    print(f"   Отчёт: {report_path}")


if __name__ == "__main__":
    main()
