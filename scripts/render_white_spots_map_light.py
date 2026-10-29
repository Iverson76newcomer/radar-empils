"""
Re-render the white-spots map in a SOFT LIGHT palette.
Output: /home/z/my-project/download/slides_titan/assets/white_spots_map_light.png
"""
import geopandas as gpd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import matplotlib.patches as mpatches
from matplotlib.lines import Line2D
import numpy as np

fm.fontManager.addfont('/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf')
plt.rcParams['font.sans-serif'] = ['DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False

# Soft light palette
BG          = "#F5F2EC"   # warm cream
INK         = "#2F3A48"
INK_SOFT    = "#5A6776"
INK_MUTE    = "#8A95A3"
HAIRLINE    = "#E2DDD2"

COUNTRY_FILL = "#EEEAE0"
COUNTRY_EDGE = "#D6CFC0"

RU_FILL = "#F0EBE0"
RU_EDGE = "#D6CFC0"

TARGET_FILL = "#E8DDC9"   # warm sand tint
TARGET_EDGE = "#B5614A"   # soft terracotta border

MARKER_COLOR = "#B5614A"  # terracotta
MARKER_RING  = "#FFFFFF"

URL = "https://naciscdn.org/naturalearth/10m/cultural/ne_10m_admin_1_states_provinces.zip"
gdf = gpd.read_file(URL)
ru = gdf[gdf['admin'] == 'Russia'].copy()
target_names = ['Altay', 'Krasnoyarsk', 'Buryat', 'Omsk', 'Tatarstan']
target_regions = ru[ru['name'].isin(target_names)].copy()
other_ru = ru[~ru['name'].isin(target_names)].copy()

CITIES = [
    ("Барнаул",     "Altay",        53.34, 83.78,  "Олимп",                       "13.10",        0),
    ("Красноярск",  "Krasnoyarsk",  56.01, 92.85,  "ООО «ОЧАГ»",                  "05–11.11",     1),
    ("Улан-Удэ",    "Buryat",       51.83, 107.58, "Вегос-М",                     "19–25.11",     2),
    ("Омск",        "Omsk",         54.99, 73.37,  "ИП Шишков",                   "ноябрь",       3),
    ("Казань",      "Tatarstan",    55.79, 49.12,  "ТС Агава, Карат, Колорит",    "ноябрь",       4),
]
EMPILS_HQ = ("Ростов-на-Дону", 47.23, 39.70)

# colors per index (soft, on light bg)
ACCENT_BY_IDX = ["#B5614A", "#C9A77C", "#D4A574", "#8BAA8A", "#C9A77C"]

fig = plt.figure(figsize=(14.5, 7.6), dpi=200, facecolor=BG)
ax = fig.add_axes([0.02, 0.02, 0.96, 0.96])
ax.set_facecolor(BG)

other_ru.plot(ax=ax, color=RU_FILL, edgecolor=RU_EDGE, linewidth=0.35, zorder=2)
other_countries = gdf[(gdf['admin'] != 'Russia')].copy()
other_countries.plot(ax=ax, color=COUNTRY_FILL, edgecolor=COUNTRY_EDGE, linewidth=0.25, zorder=1)
target_regions.plot(ax=ax, color=TARGET_FILL, edgecolor=TARGET_EDGE, linewidth=1.4, zorder=4)

ax.set_xlim(38, 115)
ax.set_ylim(48, 60)

LABEL_OFFSETS = {
    "Барнаул":     (+2.0, -1.5, "left", "top"),
    "Красноярск":  (+2.0, +1.2, "left", "bottom"),
    "Улан-Удэ":    (-2.5, -1.5, "right","top"),
    "Омск":        (-2.5, +1.5, "right","bottom"),
    "Казань":      (+2.0, +1.0, "left", "bottom"),
}

for city, region_key, lat, lon, distr, date, idx in CITIES:
    accent = ACCENT_BY_IDX[idx]
    ax.scatter([lon], [lat], s=130, color=accent, edgecolors=MARKER_RING,
               linewidths=2.0, zorder=10, marker='o')
    ax.scatter([lon], [lat], s=30, color=MARKER_RING, zorder=11, marker='o')
    lo, la, ha, va = LABEL_OFFSETS[city]
    lx, ly = lon + lo, lat + la
    ax.plot([lon, lx], [lat, ly], color=INK_MUTE, linewidth=0.8,
            linestyle='-', zorder=9, alpha=0.55)
    ax.text(lon, lat, str(idx + 1), ha='center', va='center',
            fontsize=8, fontweight='bold', color=MARKER_RING, zorder=12)
    label_text = f"{city}\n{distr}\n{date}"
    ax.text(lx, ly, label_text, ha=ha, va=va,
            fontsize=8.5, color=INK, fontweight='normal',
            linespacing=1.45,
            bbox=dict(boxstyle="round,pad=0.4", facecolor='white',
                      edgecolor=accent, linewidth=0.9, alpha=0.96),
            zorder=13)

hq_name, hq_lat, hq_lon = EMPILS_HQ
ax.scatter([hq_lon], [hq_lat], s=110, color=INK, edgecolors=MARKER_RING,
           linewidths=1.5, zorder=10, marker='D')
ax.text(hq_lon + 1.5, hq_lat + 0.5, f"Эмпилс\n{hq_name}",
        ha='left', va='bottom', fontsize=8, color=INK_SOFT, fontweight='600',
        bbox=dict(boxstyle="round,pad=0.3", facecolor=BG,
                  edgecolor=INK_MUTE, linewidth=0.6, alpha=0.95),
        zorder=13)

# Legend
legend_x, legend_y = 105.5, 58.0
ax.text(legend_x, legend_y + 0.6, "УСЛОВНЫЕ ОБОЗНАЧЕНИЯ",
        fontsize=8.5, fontweight='bold', color=INK, ha='left', va='top', zorder=15)
ax.add_patch(mpatches.Rectangle((legend_x, legend_y - 0.2), 1.2, 0.6,
                                 facecolor=TARGET_FILL, edgecolor=TARGET_EDGE,
                                 linewidth=0.8, zorder=15))
ax.text(legend_x + 1.5, legend_y + 0.1, "Целевой регион",
        fontsize=8, color=INK, ha='left', va='center', zorder=15)
ax.scatter([legend_x + 0.6], [legend_y - 1.4], s=100, color=MARKER_COLOR,
           edgecolors=MARKER_RING, linewidths=1.5, zorder=15, marker='o')
ax.text(legend_x + 1.5, legend_y - 1.4, "Дистрибьютор (встреча)",
        fontsize=8, color=INK, ha='left', va='center', zorder=15)
ax.scatter([legend_x + 0.6], [legend_y - 2.4], s=80, color=INK,
           edgecolors=MARKER_RING, linewidths=1.2, zorder=15, marker='D')
ax.text(legend_x + 1.5, legend_y - 2.4, "Эмпилс (Ростов-на-Дону)",
        fontsize=8, color=INK, ha='left', va='center', zorder=15)
ax.add_patch(mpatches.Rectangle((legend_x, legend_y - 3.4), 1.2, 0.6,
                                 facecolor=RU_FILL, edgecolor=RU_EDGE,
                                 linewidth=0.5, zorder=15))
ax.text(legend_x + 1.5, legend_y - 3.1, "Прочие регионы РФ",
        fontsize=8, color=INK, ha='left', va='center', zorder=15)
ax.text(legend_x, legend_y - 4.2, "ЦЕЛЕВЫЕ РЕГИОНЫ:",
        fontsize=7.5, fontweight='bold', color=INK_MUTE, ha='left', va='top', zorder=15)
ax.text(legend_x, legend_y - 4.8,
        "1. Алтайский край\n2. Красноярский край\n3. Республика Бурятия\n4. Омская область\n5. Республика Татарстан",
        fontsize=7.5, color=INK, ha='left', va='top', linespacing=1.5, zorder=15)

ax.set_xticks([]); ax.set_yticks([])
for spine in ax.spines.values():
    spine.set_visible(False)
rect = mpatches.Rectangle((38, 48), 77, 12, fill=False, edgecolor=HAIRLINE,
                          linewidth=1.0, zorder=1)
ax.add_patch(rect)

OUT = "/home/z/my-project/download/slides_titan/assets/white_spots_map_light.png"
fig.savefig(OUT, dpi=200, facecolor=BG, bbox_inches='tight', pad_inches=0.15)
print(f"Saved: {OUT}")
plt.close(fig)
