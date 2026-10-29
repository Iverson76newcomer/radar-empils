"""
Create a clean, professional chart for ОРЕОЛ TITAN primary turnover 2023-2026.
Data extracted from user's screenshot:
  2023: 72 911
  2024: 122 899
  2025: 127 474
  2026: 178 751

Output: /home/z/my-project/download/TITAN_первичный_товарооборот_2023_2026.png
"""
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import matplotlib.ticker as ticker
import numpy as np

# Font
fm.fontManager.addfont('/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf')
plt.rcParams['font.sans-serif'] = ['DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False

# Palette (matches presentation — pastel)
BG     = '#F5F2EC'
INK    = '#2F3A48'
MUTE   = '#8A95A3'
RED    = '#C00000'
SAND   = '#C9A77C'
BLUE   = '#6B8FA8'
GREEN  = '#8BAA8A'
DIV    = '#EFE9DD'
BORDER = '#D9D2C2'

# Data — order: 2023, 2024, 2025, 2026 (без Лемана), 2026 (с Леманой)
years = ['2023', '2024', '2025', '2026\n(без Лемана Про)', '2026\n(с Леманой Про)']
values = [72911, 122899, 127474, 150751, 178751]

# Calculate growth rates — only for 2023→2024→2025→2026(без Лемана)
growth = []
for i in range(1, 4):  # indices 1,2,3
    g = (values[i] - values[i-1]) / values[i-1] * 100
    growth.append(g)

total_growth = (values[4] - values[0]) / values[0] * 100  # 2023→2026(с Леманой)
lemana_share = (values[4] - values[3]) / values[4] * 100   # Lemana's share of 2026 total

# ── Create figure ─────────────────────────────────────────────────────
fig, ax = plt.subplots(figsize=(12, 6.5), dpi=200, facecolor=BG, constrained_layout=True)
ax.set_facecolor(BG)

# Bar chart — 4th bar (без Лемана) blue, 5th bar (с Леманой) red
bar_colors = [MUTE, SAND, SAND, BLUE, RED]
bars = ax.bar(years, values, width=0.55, color=bar_colors,
              edgecolor='white', linewidth=1.5, zorder=3)

# Add value labels on top of bars
for i, (bar, val) in enumerate(zip(bars, values)):
    val_str = f'{val:,}'.replace(',', ' ')
    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + max(values)*0.02,
            val_str, ha='center', va='bottom', fontsize=13, fontweight='bold',
            color=INK, zorder=5)
    
    # Add growth rate between year bars (only 2023→2024, 2024→2025, 2025→2026-без-Лемана)
    if i > 0 and i < 4:
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

# Connecting line through bar tops (2023→2026 без Лемана — indices 0-3)
ax.plot(range(4), values[:4], color=RED, linewidth=2, marker='o',
        markersize=8, markerfacecolor='white', markeredgewidth=2.5,
        markeredgecolor=RED, zorder=4, alpha=0.7)

# Dashed line from "без Лемана" (bar 4, index 3) to "с Леманой" (bar 5, index 4)
ax.plot([3, 4], [values[3], values[4]], color=RED, linewidth=1.5,
        linestyle='--', marker='s', markersize=7, markerfacecolor='white',
        markeredgewidth=2, markeredgecolor=RED, zorder=4, alpha=0.7)

# Add delta label between "без Лемана" and "с Леманой"
delta = values[4] - values[3]
ax.annotate(f'+{delta:,} кг\n(Лемана Про)'.replace(',', ' '),
            xy=(3.5, (values[3] + values[4]) / 2),
            ha='center', va='center', fontsize=9, fontweight='bold',
            color=RED,
            bbox=dict(boxstyle='round,pad=0.3', facecolor='white',
                      edgecolor=RED, linewidth=0.8, alpha=0.95),
            zorder=6)

# Title
ax.set_title('Первичный товарооборот ОРЕОЛ TITAN (кг)',
             fontsize=16, fontweight='bold', color=INK, pad=20)

# Subtitle
ax.text(0.5, 1.02, f'Динамика 2023—2026 · Рост за период: +{total_growth:.0f}% · Доля Лемана Про в 2026: {lemana_share:.1f}%',
        transform=ax.transAxes, ha='center', va='bottom',
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

# Add summary box at bottom right
summary_text = (
    f'2023 → 2026 (с Леманой): {values[0]:,} → {values[4]:,} кг\n'.replace(',', ' ') +
    f'Прирост: +{values[4]-values[0]:,} кг (+{total_growth:.0f}%)\n'.replace(',', ' ') +
    f'CAGR (среднегодовой): +{((values[4]/values[0])**(1/3)-1)*100:.1f}%\n\n'
    f'2026 без Лемана Про: {values[3]:,} кг\n'.replace(',', ' ') +
    f'Доля Лемана Про: {lemana_share:.1f}% ({values[4]-values[3]:,} кг)'.replace(',', ' ')
)
ax.text(0.98, 0.02, summary_text, transform=ax.transAxes,
        ha='right', va='bottom', fontsize=9, color=INK,
        bbox=dict(boxstyle='round,pad=0.5', facecolor='white',
                  edgecolor=BORDER, linewidth=0.8, alpha=0.95))

# Save
OUT = '/home/z/my-project/download/TITAN_первичный_товарооборот_2023_2026.png'
fig.savefig(OUT, dpi=200, facecolor=BG, bbox_inches='tight', pad_inches=0.3)
print(f"Saved: {OUT}")
plt.close(fig)
