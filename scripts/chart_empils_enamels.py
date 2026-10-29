"""
Create a chart showing dynamics of ALL Empils ground-enamel sales 2023-2026.
Data from user's screenshot (single "Общий итог" row):
  2023: 3 367 827
  2024: 3 400 445
  2025: 3 012 790
  2026: 2 945 402
"""
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import matplotlib.ticker as ticker
import numpy as np

fm.fontManager.addfont('/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf')
plt.rcParams['font.sans-serif'] = ['DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False

BG     = '#F5F2EC'
INK    = '#2F3A48'
MUTE   = '#8A95A3'
RED    = '#C00000'
SAND   = '#C9A77C'
BLUE   = '#6B8FA8'
GREEN  = '#8BAA8A'
DIV    = '#EFE9DD'
BORDER = '#D9D2C2'

# Data
years = ['2023', '2024', '2025', '2026']
values = [3367827, 3400445, 3012790, 2945402]

# Calculate growth rates
growth = []
for i in range(1, len(values)):
    g = (values[i] - values[i-1]) / values[i-1] * 100
    growth.append(g)

total_change = (values[-1] - values[0]) / values[0] * 100

# ── Create figure ─────────────────────────────────────────────────────
fig, ax = plt.subplots(figsize=(12, 6.5), dpi=200, facecolor=BG, constrained_layout=True)
ax.set_facecolor(BG)

# Bar chart
bar_colors = [MUTE, SAND, BLUE, RED]
bars = ax.bar(years, values, width=0.55, color=bar_colors,
              edgecolor='white', linewidth=1.5, zorder=3)

# Format values in millions
def fmt_mln(v):
    return f'{v/1e6:.2f} млн'

# Add value labels on top of bars
for i, (bar, val) in enumerate(zip(bars, values)):
    val_str = f'{val:,}'.replace(',', ' ')
    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + max(values)*0.015,
            val_str, ha='center', va='bottom', fontsize=12, fontweight='bold',
            color=INK, zorder=5)
    # Also show in millions below
    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + max(values)*0.06,
            f'({fmt_mln(val)})', ha='center', va='bottom', fontsize=9,
            color=MUTE, zorder=5)

# Add growth rate badges between bars
for i in range(1, len(values)):
    mid_x = i - 0.5
    mid_y = (values[i] + values[i-1]) / 2
    g = growth[i-1]
    color = GREEN if g > 0 else RED
    arrow = '▲' if g > 0 else '▼'
    ax.annotate(f'{arrow} {g:+.1f}%', xy=(mid_x, mid_y),
                ha='center', va='center', fontsize=10, fontweight='bold',
                color=color,
                bbox=dict(boxstyle='round,pad=0.3', facecolor='white',
                          edgecolor=color, linewidth=0.8, alpha=0.95),
                zorder=6)

# Connecting line through bar tops
ax.plot(range(len(years)), values, color=RED, linewidth=2, marker='o',
        markersize=8, markerfacecolor='white', markeredgewidth=2.5,
        markeredgecolor=RED, zorder=4, alpha=0.7)

# Title
ax.set_title('Динамика продаж всех грунт-эмалей Эмпилс (кг)',
             fontsize=16, fontweight='bold', color=INK, pad=20)

# Subtitle
subtitle = f'Период 2023—2026 · Изменение за период: {total_change:+.1f}% ({values[-1]-values[0]:+,} кг)'.replace(',', ' ')
ax.text(0.5, 1.02, subtitle, transform=ax.transAxes, ha='center', va='bottom',
        fontsize=10, color=MUTE, style='italic')

# Y-axis
ax.set_ylabel('Объём, кг', fontsize=11, color=INK, fontweight='bold')
ax.yaxis.set_major_formatter(ticker.FuncFormatter(lambda x, _: f'{int(x):,}'.replace(',', ' ')))
ax.tick_params(axis='y', labelcolor=MUTE, labelsize=10)
ax.set_ylim(0, max(values) * 1.25)

# X-axis
ax.tick_params(axis='x', labelcolor=INK, labelsize=12)

# Grid
ax.grid(axis='y', alpha=0.3, color=DIV, linestyle='--', linewidth=0.5, zorder=0)
ax.set_axisbelow(True)

# Spines
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)
ax.spines['left'].set_color(BORDER)
ax.spines['bottom'].set_color(BORDER)

# Summary box at bottom right
cagr = ((values[-1]/values[0])**(1/3)-1)*100
summary_text = (
    f'2023 → 2026: {values[0]:,} → {values[-1]:,} кг\n'.replace(',', ' ') +
    f'Изменение: {values[-1]-values[0]:+,} кг ({total_change:+.1f}%)\n'.replace(',', ' ') +
    f'CAGR (среднегодовой): {cagr:+.1f}%\n\n'
    f'Пик продаж: 2024 г. — {values[1]:,} кг\n'.replace(',', ' ') +
    f'Снижение 2024→2026: {values[1]-values[-1]:,} кг ({(values[-1]/values[1]-1)*100:+.1f}%)'.replace(',', ' ')
)
ax.text(0.98, 0.02, summary_text, transform=ax.transAxes,
        ha='right', va='bottom', fontsize=9, color=INK,
        bbox=dict(boxstyle='round,pad=0.5', facecolor='white',
                  edgecolor=BORDER, linewidth=0.8, alpha=0.95))

# Save
OUT = '/home/z/my-project/download/Эмпилс_грунт-эмали_динамика_2023_2026.png'
fig.savefig(OUT, dpi=200, facecolor=BG, bbox_inches='tight', pad_inches=0.3)
print(f"Saved: {OUT}")
plt.close(fig)
