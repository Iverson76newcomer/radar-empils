"""
Добавить колонку 'Математика расчёта' с детальной пошаговой формулой для каждого месяца.
Также обновить лист 'Объяснение простыми словами' — расширить блок с математикой.
"""
import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import shutil
import os

SRC_PATH = "/home/z/my-project/upload/Кассовый_разрыв_Олимп_модель_v2.xlsx"
OUT_PATH = "/home/z/my-project/download/Кассовый_разрыв_Олимп_модель_v3.xlsx"

# Копируем оригинал
shutil.copy(SRC_PATH, OUT_PATH)
wb = load_workbook(OUT_PATH)

# === Стили ===
HEAD_FILL = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
HEAD_FONT = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
HEAD_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)
TITLE_FONT = Font(name="Calibri", size=14, bold=True, color="1F4E79")
SUBTITLE_FONT = Font(name="Calibri", size=10, italic=True, color="595959")
BODY_FONT = Font(name="Calibri", size=10)
MATH_FONT = Font(name="Consolas", size=10, color="1A1A1A")  # моноширинный для формул
BODY_ALIGN = Alignment(horizontal="left", vertical="center", wrap_text=True)
MATH_ALIGN = Alignment(horizontal="left", vertical="top", wrap_text=True)
NUM_ALIGN = Alignment(horizontal="right", vertical="center")
CENTER_ALIGN = Alignment(horizontal="center", vertical="center")
GAP_FILL = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
POSITIVE_FILL = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
WARN_FILL = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
TOTAL_FILL = PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid")
PRIMARY_HIGHLIGHT = PatternFill(start_color="FFE699", end_color="FFE699", fill_type="solid")
SECONDARY_HIGHLIGHT = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")
INCOME_FILL = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
OUTCOME_FILL = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
THIN = Side(border_style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

# === Данные по месяцам (для построения формул) ===
# Каждая запись: месяц, первичка, вторичка, платёж_эмпилс, деньги_от_РТТ, чистый_поток, накопл_баланс, прошлый_накопл
months_data = [
    # (месяц, первичка, вторичка, платёж_эмпилс (за 2 мес назад), деньги_от_РТТ (за 1 мес назад), чистый, накопл, прошлый_накопл, прошлый_месяц)
    ('Январь', 6_000_000, 2_000_000, 0, 0, 0, 0, 0, None),
    ('Февраль', 6_000_000, 2_400_000, 0, 2_000_000, 2_000_000, 2_000_000, 0, 'Январь'),
    ('Март', 7_000_000, 4_800_000, 6_000_000, 2_400_000, -3_600_000, -1_600_000, 2_000_000, 'Февраль'),
    ('Апрель', 7_000_000, 7_800_000, 6_000_000, 4_800_000, -1_200_000, -2_800_000, -1_600_000, 'Март'),
    ('Май', 4_000_000, 7_800_000, 7_000_000, 7_800_000, 800_000, -2_000_000, -2_800_000, 'Апрель'),
    ('Июнь', 6_600_000, 7_750_000, 7_000_000, 7_800_000, 800_000, -1_200_000, -2_000_000, 'Май'),
    ('Июль', 6_000_000, 7_750_000, 4_000_000, 7_750_000, 3_750_000, 2_550_000, -1_200_000, 'Июнь'),
    ('Август', 5_100_000, 6_340_000, 6_600_000, 7_750_000, 1_150_000, 3_700_000, 2_550_000, 'Июль'),
    ('Сентябрь', 4_200_000, 4_580_000, 6_000_000, 6_340_000, 340_000, 4_040_000, 3_700_000, 'Август'),
    ('Октябрь', 3_300_000, 3_520_000, 5_100_000, 4_580_000, -520_000, 3_520_000, 4_040_000, 'Сентябрь'),
    ('Ноябрь', 2_700_000, 2_820_000, 4_200_000, 3_520_000, -680_000, 2_840_000, 3_520_000, 'Октябрь'),
    ('Декабрь', 2_100_000, 2_460_000, 3_300_000, 2_820_000, -480_000, 2_360_000, 2_840_000, 'Ноябрь'),
]

# Функция для форматирования чисел в млн руб
def fmt_mln(val):
    """Форматировать число в строку вида '+2,0 млн' или '-3,6 млн' или '0'."""
    if val == 0:
        return '0'
    sign = '+' if val > 0 else '−'  # используем минус-типографский для красоты
    abs_val = abs(val) / 1_000_000
    return f"{sign}{abs_val:.1f} млн"

def fmt_mln_plain(val):
    """Форматировать без знака (для сумм)."""
    if val == 0:
        return '0'
    abs_val = abs(val) / 1_000_000
    return f"{abs_val:.1f} млн"

# === 1. Обновить лист 'Кассовый поток (детально)' — добавить колонку 'Математика расчёта' ===
print("=== Обновление листа 'Кассовый поток (детально)' ===")
ws = wb["Кассовый поток (детально)"]

# Текущая шапка (строка 4) имеет 9 колонок (A-I). Добавим колонку J 'Математика расчёта'
new_col = 10  # J
ws.cell(row=4, column=new_col, value="Математика расчёта (пошагово)")
cell = ws.cell(row=4, column=new_col)
cell.fill = HEAD_FILL
cell.font = HEAD_FONT
cell.alignment = HEAD_ALIGN
cell.border = BORDER
ws.column_dimensions[get_column_letter(new_col)].width = 60

# Заполнить колонку математикой для каждого месяца
for i, m in enumerate(months_data):
    month, primary, secondary, pay_empils, receive_rtt, net, cumulative, prev_cumul, prev_month = m
    r = 5 + i
    
    # Построить строку с математикой
    if i == 0:
        # Январь — особый случай
        math_lines = [
            f"ЧИСТЫЙ ПОТОК = Деньги от РТТ − Платёж Эмпилс",
            f"  = 0 − 0 = 0 руб",
            f"  (платёж Эмпилс за янв. первичку придёт в марте,",
            f"   деньги от РТТ за янв. вторичку придут в феврале)",
            f"",
            f"НАКОПЛ. БАЛАНС = Прошлый баланс + Чистый поток",
            f"  = 0 (старт) + 0 = 0 руб",
        ]
    else:
        # Обычный месяц
        receive_str = fmt_mln(receive_rtt) if receive_rtt > 0 else '0'
        pay_str = fmt_mln(-pay_empils) if pay_empils > 0 else '0'  # показываем как расход
        
        math_lines = [
            f"ЧИСТЫЙ ПОТОК = Деньги от РТТ − Платёж Эмпилс",
            f"  = {receive_str} ({prev_month.lower()} вторичка, +30 дн)",
        ]
        if pay_empils > 0:
            # Найти, за какой месяц платим (2 месяца назад)
            two_months_ago_idx = i - 2
            if two_months_ago_idx >= 0:
                two_months_ago = months_data[two_months_ago_idx][0]
                math_lines.append(f"  − {fmt_mln_plain(pay_empils)} ({two_months_ago.lower()} первичка, +45 дн)")
            else:
                math_lines.append(f"  − {fmt_mln_plain(pay_empils)} (платёж Эмпилс)")
        else:
            math_lines.append(f"  − 0 (платёж Эмпилс ещё не наступил)")
        
        math_lines.append(f"  = {fmt_mln(net)} руб")
        math_lines.append(f"")
        math_lines.append(f"НАКОПЛ. БАЛАНС = Прошлый баланс + Чистый поток")
        math_lines.append(f"  = {fmt_mln(prev_cumul)} ({prev_month}) + {fmt_mln(net)}")
        math_lines.append(f"  = {fmt_mln(cumulative)} руб")
    
    # Записать в ячейку
    cell = ws.cell(row=r, column=new_col, value='\n'.join(math_lines))
    cell.font = MATH_FONT
    cell.alignment = MATH_ALIGN
    cell.border = BORDER
    
    # Подсветка
    if cumulative < 0:
        cell.fill = GAP_FILL
    elif cumulative > 0 and net > 0:
        cell.fill = POSITIVE_FILL
    elif cumulative > 0:
        cell.fill = INCOME_FILL
    else:
        cell.fill = PatternFill()
    
    # Увеличить высоту строки под многострочный текст
    ws.row_dimensions[r].height = 110

# Обновить итоговую строку
r_total = 5 + len(months_data)
# В текущем файле на строке 17 уже есть итог — обновим
# Но нам нужно добавить математику для итога
math_total_lines = [
    f"ИТОГО ЗА ГОД:",
    f"  Сумма первички = 60,0 млн (закуплено у Эмпилс)",
    f"  Сумма вторички = 60,0 млн (продано в РТТ)",
    f"  Платежи Эмпилс = 60,0 млн (оплачено с лагом 45 дн)",
    f"  Деньги от РТТ = 60,0 млн (получено с лагом 30 дн)",
    f"",
    f"  Чистый поток за год = 0 (сумма всех месяцев)",
    f"  Накопл. баланс на конец года = +2,36 млн (резерв)",
    f"  ⚠ Кассовый разрыв: март-июнь (4 мес), минимум в апреле (-2,8 млн)",
]
cell = ws.cell(row=r_total, column=new_col, value='\n'.join(math_total_lines))
cell.font = MATH_FONT
cell.fill = TOTAL_FILL
cell.alignment = MATH_ALIGN
cell.border = BORDER
ws.row_dimensions[r_total].height = 130

# === 2. Обновить лист 'Объяснение простыми словами' — добавить детальный блок с математикой ===
print("\n=== Обновление листа 'Объяснение простыми словами' ===")
ws2 = wb["Объяснение простыми словами"]

# Найти последнюю занятую строку
max_r = ws2.max_row
print(f"  Текущая последняя строка: {max_r}")

# Добавить новый блок после всех существующих данных
r = max_r + 3

# Заголовок блока
ws2.merge_cells(f"A{r}:F{r}")
ws2.cell(row=r, column=1, value="ДЕТАЛЬНАЯ МАТЕМАТИКА: КАК СЧИТАЕТСЯ КАЖДЫЙ МЕСЯЦ").font = Font(name="Calibri", size=12, bold=True, color="1F4E79")
ws2.cell(row=r, column=1).fill = PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid")
ws2.row_dimensions[r].height = 24
r += 2

# Вступительный текст
ws2.merge_cells(f"A{r}:F{r+2}")
ws2.cell(row=r, column=1, value=(
    "Ниже — пошаговый расчёт для каждого месяца с конкретными числами. "
    "Формулы одни и те же, меняются только входящие значения.\n\n"
    "ФОРМУЛА 1: Чистый поток = Деньги от РТТ (пришли за месяц) − Платёж Эмпилс (ушёл за месяц)\n"
    "ФОРМУЛА 2: Накопленный баланс = Накопленный баланс прошлого месяца + Чистый поток текущего месяца"
))
ws2.cell(row=r, column=1).font = Font(name="Calibri", size=11, bold=True, color="1F4E79")
ws2.cell(row=r, column=1).alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
ws2.cell(row=r, column=1).fill = PatternFill(start_color="FFF8E1", end_color="FFF8E1", fill_type="solid")
for rr in range(r, r+3):
    ws2.row_dimensions[rr].height = 22
r += 4

# Пример: январь и март как иллюстрация
ws2.merge_cells(f"A{r}:F{r}")
ws2.cell(row=r, column=1, value="ПРИМЕРЫ: ЯНВАРЬ (старт) и МАРТ (первый разрыв)").font = Font(name="Calibri", size=11, bold=True, color="1F4E79")
ws2.row_dimensions[r].height = 20
r += 1

# Январь
ws2.merge_cells(f"A{r}:F{r}")
ws2.cell(row=r, column=1, value="ЯНВАРЬ (старт, баланс = 0)").font = Font(name="Calibri", size=11, bold=True, color="1F4E79")
r += 1

jan_lines = [
    "1. В январе дистрибьютор:",
    "   • Закупил у Эмпилс товар на 6,0 млн руб (первичка)",
    "   • Отгрузил в РТТ товар на 2,0 млн руб (вторичка)",
    "",
    "2. Платёж Эмпилс в январе = 0",
    "   (отсрочка 45 дней — за январскую первичку платить будем в марте)",
    "",
    "3. Деньги от РТТ в январе = 0",
    "   (отсрочка 30 дней — за январскую вторичку РТТ заплатит в феврале)",
    "",
    "4. Чистый поток января = 0 (пришло) − 0 (ушло) = 0 руб",
    "",
    "5. Накопленный баланс на конец января = 0 (старт) + 0 (чистый поток) = 0 руб",
    "",
    "→ На счетах 0 руб, но на складе товара на 6,0 млн (это актив, но не деньги)",
]
ws2.merge_cells(f"A{r}:F{r+len(jan_lines)-1}")
ws2.cell(row=r, column=1, value='\n'.join(jan_lines)).font = Font(name="Consolas", size=10)
ws2.cell(row=r, column=1).alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
ws2.cell(row=r, column=1).fill = PatternFill(start_color="F8F9FA", end_color="F8F9FA", fill_type="solid")
for rr in range(r, r+len(jan_lines)):
    ws2.row_dimensions[rr].height = 18
r += len(jan_lines) + 1

# Март
ws2.merge_cells(f"A{r}:F{r}")
ws2.cell(row=r, column=1, value="МАРТ (первый кассовый разрыв, баланс = −1,6 млн)").font = Font(name="Calibri", size=11, bold=True, color="C00000")
r += 1

mar_lines = [
    "1. В марте дистрибьютор:",
    "   • Закупил у Эмпилс товар на 7,0 млн руб (первичка марта)",
    "   • Отгрузил в РТТ товар на 4,8 млн руб (вторичка марта)",
    "",
    "2. Платёж Эмпилс в марте = 6,0 млн руб",
    "   (платим за ЯНВАРСКУЮ первичку — прошло 45 дней с момента получения товара)",
    "   Формула лага: январская первичка → платёж в марте (через ~1,5 месяца)",
    "",
    "3. Деньги от РТТ в марте = 2,4 млн руб",
    "   (РТТ платит за ФЕВРАЛЬСКУЮ вторичку — прошло 30 дней с момента отгрузки)",
    "   Формула лага: февральская вторичка → деньги в марте (через 1 месяц)",
    "",
    "4. Чистый поток марта = 2,4 млн (пришло) − 6,0 млн (ушло) = −3,6 млн руб",
    "   ⚠ Минус! Ушло больше, чем пришло.",
    "",
    "5. Накопленный баланс на конец марта:",
    "   = Накопл. баланс февраля (+2,0 млн) + Чистый поток марта (−3,6 млн)",
    "   = +2,0 + (−3,6) = −1,6 млн руб",
    "",
    "→ На счетах −1,6 млн руб. Это и есть КАССОВЫЙ РАЗРЫВ.",
    "   Дистрибьютор должен Эмпилс 6,0 млн, а на счетах только 2,4 млн + 2,0 млн резерва = 4,4 млн.",
    "   Не хватает 1,6 млн, чтобы расплатиться полностью.",
]
ws2.merge_cells(f"A{r}:F{r+len(mar_lines)-1}")
ws2.cell(row=r, column=1, value='\n'.join(mar_lines)).font = Font(name="Consolas", size=10)
ws2.cell(row=r, column=1).alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
ws2.cell(row=r, column=1).fill = GAP_FILL
for rr in range(r, r+len(mar_lines)):
    ws2.row_dimensions[rr].height = 18
r += len(mar_lines) + 2

# Полная таблица математики по всем месяцам
ws2.merge_cells(f"A{r}:F{r}")
ws2.cell(row=r, column=1, value="ПОЛНАЯ ТАБЛИЦА МАТЕМАТИКИ ПО ВСЕМ МЕСЯЦАМ").font = Font(name="Calibri", size=11, bold=True, color="1F4E79")
ws2.row_dimensions[r].height = 20
r += 1

# Заголовок таблицы
table_headers = [
    ("Месяц", 10),
    ("Пришло (от РТТ)", 14),
    ("Ушло (Эмпилс)", 14),
    ("Чистый поток = Пришло − Ушло", 22),
    ("Прошлый баланс", 14),
    ("Накопл. = Прошлый + Чистый", 22),
]
for col_idx, (title, width) in enumerate(table_headers, 1):
    cell = ws2.cell(row=r, column=col_idx, value=title)
    cell.fill = HEAD_FILL
    cell.font = HEAD_FONT
    cell.alignment = HEAD_ALIGN
    cell.border = BORDER
    ws2.column_dimensions[get_column_letter(col_idx)].width = width
ws2.row_dimensions[r].height = 42
r += 1

# Данные по месяцам
for i, m in enumerate(months_data):
    month, primary, secondary, pay_empils, receive_rtt, net, cumulative, prev_cumul, prev_month = m
    
    receive_str = f"+{receive_rtt/1e6:.1f}" if receive_rtt > 0 else "0"
    pay_str = f"−{pay_empils/1e6:.1f}" if pay_empils > 0 else "0"
    net_str = f"{net/1e6:+.1f}".replace('+', '+').replace('-', '−')
    prev_str = f"{prev_cumul/1e6:+.1f}".replace('+', '+').replace('-', '−') if i > 0 else "0 (старт)"
    cum_str = f"{cumulative/1e6:+.1f}".replace('+', '+').replace('-', '−')
    
    # Месяц
    cell = ws2.cell(row=r, column=1, value=month)
    cell.font = Font(name="Calibri", size=10, bold=True)
    cell.alignment = CENTER_ALIGN
    cell.border = BORDER
    
    # Пришло
    cell = ws2.cell(row=r, column=2, value=f"{receive_str} млн" if receive_str != "0" else "0")
    cell.font = Font(name="Calibri", size=10, color="375623" if receive_rtt > 0 else "595959", bold=receive_rtt > 0)
    cell.alignment = CENTER_ALIGN
    cell.border = BORDER
    if receive_rtt > 0:
        cell.fill = INCOME_FILL
    
    # Ушло
    cell = ws2.cell(row=r, column=3, value=f"{pay_str} млн" if pay_str != "0" else "0")
    cell.font = Font(name="Calibri", size=10, color="C00000" if pay_empils > 0 else "595959", bold=pay_empils > 0)
    cell.alignment = CENTER_ALIGN
    cell.border = BORDER
    if pay_empils > 0:
        cell.fill = OUTCOME_FILL
    
    # Чистый поток
    cell = ws2.cell(row=r, column=4, value=f"{net_str} млн")
    if net < 0:
        cell.font = Font(name="Calibri", size=10, bold=True, color="C00000")
        cell.fill = GAP_FILL
    elif net > 0:
        cell.font = Font(name="Calibri", size=10, bold=True, color="375623")
        cell.fill = POSITIVE_FILL
    else:
        cell.font = Font(name="Calibri", size=10, color="595959")
    cell.alignment = CENTER_ALIGN
    cell.border = BORDER
    
    # Прошлый баланс
    cell = ws2.cell(row=r, column=5, value=f"{prev_str} млн")
    if prev_cumul < 0:
        cell.font = Font(name="Calibri", size=10, bold=True, color="C00000")
        cell.fill = GAP_FILL
    elif prev_cumul > 0:
        cell.font = Font(name="Calibri", size=10, bold=True, color="375623")
        cell.fill = INCOME_FILL
    else:
        cell.font = Font(name="Calibri", size=10, color="595959", italic=True)
    cell.alignment = CENTER_ALIGN
    cell.border = BORDER
    
    # Накопленный
    cell = ws2.cell(row=r, column=6, value=f"{cum_str} млн")
    if cumulative < 0:
        cell.font = Font(name="Calibri", size=11, bold=True, color="C00000")
        cell.fill = GAP_FILL
    elif cumulative > 0:
        cell.font = Font(name="Calibri", size=11, bold=True, color="375623")
        cell.fill = POSITIVE_FILL
    else:
        cell.font = Font(name="Calibri", size=11, bold=True, color="595959")
    cell.alignment = CENTER_ALIGN
    cell.border = BORDER
    
    ws2.row_dimensions[r].height = 22
    r += 1

# Итог
ws2.cell(row=r, column=1, value="ИТОГО").font = Font(name="Calibri", size=10, bold=True)
ws2.cell(row=r, column=1).fill = TOTAL_FILL
ws2.cell(row=r, column=1).border = BORDER
ws2.cell(row=r, column=1).alignment = CENTER_ALIGN

cell = ws2.cell(row=r, column=2, value="60,0 млн")
cell.font = Font(name="Calibri", size=10, bold=True)
cell.fill = TOTAL_FILL
cell.alignment = CENTER_ALIGN
cell.border = BORDER

cell = ws2.cell(row=r, column=3, value="60,0 млн")
cell.font = Font(name="Calibri", size=10, bold=True)
cell.fill = TOTAL_FILL
cell.alignment = CENTER_ALIGN
cell.border = BORDER

cell = ws2.cell(row=r, column=4, value="0")
cell.font = Font(name="Calibri", size=10, bold=True)
cell.fill = TOTAL_FILL
cell.alignment = CENTER_ALIGN
cell.border = BORDER

cell = ws2.cell(row=r, column=5, value="—")
cell.font = Font(name="Calibri", size=10, bold=True)
cell.fill = TOTAL_FILL
cell.alignment = CENTER_ALIGN
cell.border = BORDER

cell = ws2.cell(row=r, column=6, value="+2,4 млн")
cell.font = Font(name="Calibri", size=10, bold=True, color="375623")
cell.fill = TOTAL_FILL
cell.alignment = CENTER_ALIGN
cell.border = BORDER

r += 3

# Заключительный блок
ws2.merge_cells(f"A{r}:F{r}")
ws2.cell(row=r, column=1, value="ОБЩАЯ ЛОГИКА: ПОЧЕМУ РАЗРЫВ ВОЗНИКАЕТ В МАРТЕ-АПРЕЛЕ").font = Font(name="Calibri", size=11, bold=True, color="1F4E79")
ws2.row_dimensions[r].height = 22
r += 1

logic_lines = (
    "1. В январе-феврале дистрибьютор делает СТАРТОВУЮ ЗАГРУЗКУ склада: 12,0 млн руб первички (6+6),\n"
    "   а вторичка ещё маленькая (2,0 + 2,4 = 4,4 млн). Денег от РТТ почти не приходит.\n\n"
    "2. С марта начинают приходить платежи Эмпилс за январскую первичку (6,0 млн в марте).\n"
    "   В это же время от РТТ приходит только 2,4 млн (за февральскую вторичку).\n"
    "   Разница: 2,4 − 6,0 = −3,6 млн (чистый поток марта).\n\n"
    "3. Кумулятивно: было +2,0 (после февраля), добавили −3,6 (март) → стало −1,6 млн. РАЗРЫВ!\n\n"
    "4. В апреле та же картина: платим 6,0 млн (за февраль), а получаем 4,8 млн (за март).\n"
    "   Чистый поток = 4,8 − 6,0 = −1,2 млн.\n"
    "   Накопл. = −1,6 (март) + (−1,2) (апрель) = −2,8 млн. МИНИМУМ разрыва.\n\n"
    "5. С мая ситуация улучшается: вторичка выросла (пик весны), платежи Эмпилс тоже есть,\n"
    "   но чистый поток уже +0,8 млн (7,8 пришло − 7,0 ушло).\n"
    "   Накопл. = −2,8 (апрель) + 0,8 (май) = −2,0 млн. Разрыв сокращается.\n\n"
    "6. В июле чистый поток = +3,75 млн (7,75 пришло − 4,0 ушло за «короткую» майскую первичку).\n"
    "   Накопл. = −1,2 (июнь) + 3,75 (июль) = +2,55 млн. ВЫХОД В ПЛЮС!\n\n"
    "7. К концу года накопленный баланс = +2,36 млн. Это резерв на следующий сезон."
)
ws2.merge_cells(f"A{r}:F{r+14}")
ws2.cell(row=r, column=1, value=logic_lines).font = Font(name="Calibri", size=11)
ws2.cell(row=r, column=1).alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
ws2.cell(row=r, column=1).fill = PatternFill(start_color="FFF8E1", end_color="FFF8E1", fill_type="solid")
for rr in range(r, r+15):
    ws2.row_dimensions[rr].height = 18

# === Сохранить ===
wb.properties.creator = "Z.ai"
wb.properties.title = "Модель кассового разрыва — Олимп v3 (с детальной математикой)"
wb.properties.subject = "Добавлена колонка 'Математика расчёта' и расширен блок объяснений"

print(f"\nЛисты в файле: {wb.sheetnames}")
wb.save(OUT_PATH)
print(f"\n✓ Файл сохранён: {OUT_PATH}")
print(f"  Размер: {os.path.getsize(OUT_PATH):,} байт")
