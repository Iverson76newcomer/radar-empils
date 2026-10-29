"""
Create a comparison chart: ALL Empils ground-enamels (declining) vs OREOL TITAN (growing)
2023-2026, showing that TITAN is growing while the overall portfolio is declining.
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

# ALL Empils ground-enamels (declining)
all_enamels = [3367827, 3400445, 3012790, 2945402]

# OREOL TITAN only (growing)
titan = [72911, 122899, 127474, 178751]

# Calculate TITAN's share of total
titan_share = [t/a*100 for t, a in zip(titan, all_enamels)]

# ── Create figure with 2 subplots ─────────────────────────────────────
fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(13, 9), dpi=200, facecolor=BG,
                                constrained_layout=True, height_ratios=[3, 2])

# ════════════════════════════════════════════════════════════════════════
# TOP CHART: Dual-axis — All enamels (left, declining) vs TITAN (right, growing)
# ════════════════════════════════════════════════════════════════════════
ax1.set_facecolor(BG)

# All enamels — large bars (left axis)
bars_all = ax1.bar(years, all_enamels, width=0.5, color=MUTE, alpha=0.6,
                   edgecolor='white', linewidth=1.5, zorder=2, label='Все грунт-эмали Эмпилс')

# Value labels for all enamels
for bar, val in zip(bars_all, all_enamels):
    ax1.text(bar.get_x() + bar.get_width()/2, bar.get_height() + max(all_enamels)*0.01,
             f'{val/1e6:.2f} млн', ha='center', va='bottom', fontsize=10, fontweight='bold',
             color=MUTE, zorder=5)

ax1.set_ylabel('Все грунт-эмали (кг)', fontsize=11, color=MUTE, fontweight='bold')
ax1.yaxis.set_major_formatter(ticker.FuncFormatter(lambda x, _: f'{x/1e6:.1f} млн'))
ax1.tick_params(axis='y', labelcolor=MUTE, labelsize=10)
ax1.set_ylim(0, max(all_enamels) * 1.25)
ax1.tick_params(axis='x', labelsize=11, colors=INK)
ax1.grid(axis='y', alpha=0.3, color=DIV, linestyle='--', linewidth=0.5, zorder=0)
ax1.set_axisbelow(True)
ax1.spines['top'].set_visible(False)

# TITAN — line + markers (right axis, growing)
ax1b = ax1.twinx()
line_titan = ax1b.plot(years, titan, color=RED, linewidth=3, marker='o',
                        markersize=12, markerfacecolor='white', markeredgewidth=3,
                        markeredgecolor=RED, zorder=5, label='ОРЕОЛ TITAN')

# Value labels for TITAN
for i, val in enumerate(titan):
    ax1b.text(i, val + max(titan)*0.06, f'{val:,}'.replace(',', ' '),
              ha='center', va='bottom', fontsize=11, fontweight='bold', color=RED, zorder=6)

ax1b.set_ylabel('ОРЕОЛ TITAN (кг)', fontsize=11, color=RED, fontweight='bold')
ax1b.yaxis.set_major_formatter(ticker.FuncFormatter(lambda x, _: f'{int(x):,}'.replace(',', ' ')))
ax1b.tick_params(axis='y', labelcolor=RED, labelsize=10)
ax1b.set_ylim(0, max(titan) * 1.5)
ax1b.spines['top'].set_visible(False)
ax1b.spines['right'].set_color(RED)
ax1b.spines['left'].set_visible(False)

# Growth badges on TITAN line
titan_growth = [(titan[1]-titan[0])/titan[0]*100, (titan[2]-titan[1])/titan[1]*100, (titan[3]-titan[2])/titan[2]*100]
for i, g in enumerate(titan_growth):
    mid_x = i + 0.5
    mid_y = (titan[i] + titan[i+1]) / 2
    color = GREEN if g > 0 else RED
    arrow = '▲' if g > 0 else '▼'
    ax1b.annotate(f'{arrow}{g:+.1f}%', xy=(mid_x, mid_y),
                  ha='center', va='center', fontsize=9, fontweight='bold', color=color,
                  bbox=dict(boxstyle='round,pad=0.2', facecolor='white',
                            edgecolor=color, linewidth=0.8, alpha=0.95), zorder=7)

# Decline badges on all enamels
all_growth = [(all_enamels[1]-all_enamels[0])/all_enamels[0]*100,
              (all_enamels[2]-all_enamels[1])/all_enamels[1]*100,
              (all_enamels[3]-all_enamels[2])/all_enamels[2]*100]
for i, g in enumerate(all_growth):
    mid_x = i + 0.5
    mid_y = (all_enamels[i] + all_enamels[i+1]) / 2
    color = GREEN if g > 0 else RED
    arrow = '▲' if g > 0 else '▼'
    ax1.annotate(f'{arrow}{g:+.1f}%', xy=(mid_x, mid_y),
                 ha='center', va='center', fontsize=8, fontweight='bold', color=color,
                 bbox=dict(boxstyle='round,pad=0.2', facecolor='white',
                           edgecolor=color, linewidth=0.6, alpha=0.9), zorder=4)

# Title for top chart
ax1.set_title('Все грунт-эмали Эмпилс (снижение) vs ОРЕОЛ TITAN (рост), кг',
              fontsize=14, fontweight='bold', color=INK, pad=15)

# Combined legend
lines1, labels1 = ax1.get_legend_handles_labels()
lines2, labels2 = ax1b.get_legend_handles_labels()
ax1.legend(lines1 + lines2, labels1 + labels2, loc='upper right',
           framealpha=0.95, facecolor='white', edgecolor=BORDER, fontsize=10)

# ════════════════════════════════════════════════════════════════════════
# BOTTOM CHART: TITAN's share of total ground-enamels (%)
# ════════════════════════════════════════════════════════════════════════
ax2.set_facecolor(BG)

# Area chart for TITAN share
ax2.fill_between(years, 0, titan_share, color=RED, alpha=0.2, zorder=2)
ax2.plot(years, titan_share, color=RED, linewidth=2.5, marker='o',
         markersize=10, markerfacecolor='white', markeredgewidth=2.5,
         markeredgecolor=RED, zorder=3)

# Value labels
for i, (y, s) in enumerate(zip(years, titan_share)):
    ax2.text(i, s + 0.15, f'{s:.2f}%', ha='center', va='bottom',
             fontsize=13, fontweight='bold', color=RED, zorder=5)

# Growth badges for share
share_growth = [(titan_share[1]-titan_share[0])/titan_share[0]*100,
                (titan_share[2]-titan_share[1])/titan_share[1]*100,
                (titan_share[3]-titan_share[2])/titan_share[2]*100]
for i, g in enumerate(share_growth):
    mid_x = i + 0.5
    mid_y = (titan_share[i] + titan_share[i+1]) / 2
    ax2.annotate(f'▲{g:+.0f}%', xy=(mid_x, mid_y),
                 ha='center', va='center', fontsize=9, fontweight='bold', color=GREEN,
                 bbox=dict(boxstyle='round,pad=0.2', facecolor='white',
                           edgecolor=GREEN, linewidth=0.8, alpha=0.95), zorder=6)

ax2.set_title('Доля ОРЕОЛ TITAN в общем объёме грунт-эмалей Эмпилс (%)',
              fontsize=13, fontweight='bold', color=INK, pad=12)
ax2.set_ylabel('Доля TITAN, %', fontsize=11, color=RED, fontweight='bold')
ax2.tick_params(axis='y', labelcolor=RED, labelsize=10)
ax2.tick_params(axis='x', labelsize=11, colors=INK)
ax2.set_ylim(0, max(titan_share) * 1.5)
ax2.yaxis.set_major_formatter(ticker.FuncFormatter(lambda x, _: f'{x:.1f}%'))
ax2.grid(axis='y', alpha=0.3, color=DIV, linestyle='--', linewidth=0.5, zorder=0)
ax2.set_axisbelow(True)
ax2.spines['top'].set_visible(False)
ax2.spines['right'].set_visible(False)
ax2.spines['left'].set_color(BORDER)
ax2.spines['bottom'].set_color(BORDER)

# Summary box at bottom right of bottom chart
all_change = (all_enamels[-1] - all_enamels[0]) / all_enamels[0] * 100
titan_change = (titan[-1] - titan[0]) / titan[0] * 100
share_change_abs = titan_share[-1] - titan_share[0]

summary = (
    f'Все грунт-эмали: {all_enamels[0]:,} → {all_enamels[-1]:,} кг ({all_change:+.1f}%)\n'.replace(',', ' ') +
    f'ОРЕОЛ TITAN: {titan[0]:,} → {titan[-1]:,} кг ({titan_change:+.1f}%)\n'.replace(',', ' ') +
    f'Доля TITAN: {titan_share[0]:.2f}% → {titan_share[-1]:.2f}% (+{share_change_abs:.2f} п.п.)'
)
ax2.text(0.98, 0.95, summary, transform=ax2.transAxes,
         ha='right', va='top', fontsize=9, color=INK,
         bbox=dict(boxstyle='round,pad=0.5', facecolor='white',
                   edgecolor=BORDER, linewidth=0.8, alpha=0.95))

# Suptitle
fig.suptitle('Контраст: рынок грунт-эмалей снижается, TITAN растёт',
             fontsize=16, fontweight='bold', color=INK, y=1.01)

# Save
OUT = '/home/z/my-project/download/Эмпилс_грунт-эмали_vs_TITAN_выводы.png'
fig.savefig(OUT, dpi=200, facecolor=BG, bbox_inches='tight', pad_inches=0.3)
print(f"Saved: {OUT}")
plt.close(fig)
