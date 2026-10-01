"""
Render a pastel geographic map of Russia (Asian part + Volga) showing 5 target regions
where Empils competitors are located, with markers and distributor info.
Output: /home/z/my-project/download/slides_titan/assets/white_spots_map.png
"""
import geopandas as gpd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import matplotlib.patches as mpatches
from matplotlib.patches import FancyBboxPatch, Circle
from matplotlib.lines import Line2D
import numpy as np

# Font registration — DejaVu Sans supports Cyrillic
fm.fontManager.addfont('/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf')
plt.rcParams['font.sans-serif'] = ['DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False

# Dark contrast palette (matches user's existing slide 1)
BG          = "#0A0A0A"   # near-black background
INK         = "#FFFFFF"   # white text
INK_SOFT    = "#CCCCCC"
INK_MUTE    = "#888888"
HAIRLINE    = "#333333"

# Country fill — very dark
COUNTRY_FILL = "#161616"
COUNTRY_EDGE = "#2A2A2A"

# Non-target Russian regions
RU_FILL = "#161616"
RU_EDGE = "#2A2A2A"

# Target regions — accent yellow (like TITAN brand accent)
TARGET_FILL = "#1E1E1E"
TARGET_EDGE = "#FFD60A"  # yellow accent

# Highlight border
TARGET_BORDER = "#E63946"  # red primary

# Marker (city point)
MARKER_COLOR = "#E63946"
MARKER_RING  = "#0A0A0A"

# Load Natural Earth admin_1
URL = "https://naciscdn.org/naturalearth/10m/cultural/ne_10m_admin_1_states_provinces.zip"
print("Loading admin_1…")
gdf = gpd.read_file(URL)

ru = gdf[gdf['admin'] == 'Russia'].copy()
print(f"Russia subjects: {len(ru)}")

# Target region names in dataset
target_names = ['Altay', 'Krasnoyarsk', 'Buryat', 'Omsk', 'Tatarstan']
target_regions = ru[ru['name'].isin(target_names)].copy()
other_ru = ru[~ru['name'].isin(target_names)].copy()

# Cities with distributor info
# (city, region_name_for_lookup, lat, lon, distributor, date, color_group)
CITIES = [
    # Барнаул, Алтайский край — Олимп, 13.10
    ("Барнаул",       "Altay",       53.34, 83.78,  "Олимп",          "13.10.2026", 0),
    # Красноярск, Красноярский край — ООО ОЧАГ, 05.11–11.11
    ("Красноярск",    "Krasnoyarsk", 56.01, 92.85,  "ООО «ОЧАГ»",     "05–11.11.2026", 1),
    # Улан-Удэ, Бурятия — Вегос-М, 19.11–25.11
    ("Улан-Удэ",      "Buryat",      51.83, 107.58, "Вегос-М",        "19–25.11.2026", 2),
    # Омск, Омская область — ИП Шишков, ноябрь
    ("Омск",          "Omsk",        54.99, 73.37,  "ИП Шишков",      "ноябрь 2026", 3),
    # Казань, Татарстан — ТС Агава, Карат, Колорит, ноябрь
    ("Казань",        "Tatarstan",   55.79, 49.12,  "ТС Агава, Карат, Колорит", "ноябрь 2026", 4),
]

# Empils HQ — Ростов-на-Дону (для ориентира)
EMPILS_HQ = ("Ростов-на-Дону", 47.23, 39.70)

# ────────────────────────────────────────────────────────────────────────
# FIGURE
# ────────────────────────────────────────────────────────────────────────
fig = plt.figure(figsize=(14.5, 7.6), dpi=200, facecolor=BG)
ax = fig.add_axes([0.02, 0.02, 0.96, 0.96])
ax.set_facecolor(BG)

# Plot non-target Russian regions (very light taupe)
other_ru.plot(ax=ax, color=RU_FILL, edgecolor=RU_EDGE, linewidth=0.35, zorder=2)

# Plot neighbouring countries for context
other_countries = gdf[(gdf['admin'] != 'Russia')].copy()
other_countries.plot(ax=ax, color=COUNTRY_FILL, edgecolor=COUNTRY_EDGE, linewidth=0.25, zorder=1)

# Plot TARGET regions — pastel sand fill with thicker blue border
target_regions.plot(ax=ax, color=TARGET_FILL, edgecolor=TARGET_BORDER, linewidth=1.4, zorder=4)

# ── Map extent — focus on European Russia + Siberia
# Russia spans roughly 20°E to 180°E. We focus 38°E–115°E (Volga to Baikal)
ax.set_xlim(38, 115)
ax.set_ylim(48, 60)

# ────────────────────────────────────────────────────────────────────────
# CITIES + MARKERS
# ────────────────────────────────────────────────────────────────────────
# Marker positions for labels (offsets in degrees to avoid overlap)
# (lon_offset, lat_offset, ha, va, leader_line_to)
LABEL_OFFSETS = {
    "Барнаул":     (+2.0,  -1.5, "left", "top",     "right-down"),
    "Красноярск":  (+2.0,  +1.2, "left", "bottom",   "right-up"),
    "Улан-Удэ":    (-2.5,  -1.5, "right","top",      "left-down"),
    "Омск":        (-2.5,  +1.5, "right","bottom",   "left-up"),
    "Казань":      (+2.0,  +1.0, "left", "bottom",   "right-up"),
}

# Draw markers + leader lines + labels
for city, region_key, lat, lon, distr, date, idx in CITIES:
    # Marker (city point)
    ax.scatter([lon], [lat], s=130, color=MARKER_COLOR, edgecolors=MARKER_RING,
               linewidths=2.0, zorder=10, marker='o')

    # Inner dot
    ax.scatter([lon], [lat], s=30, color=MARKER_RING, zorder=11, marker='o')

    # Label position
    lo, la, ha, va, side = LABEL_OFFSETS[city]
    lx, ly = lon + lo, lat + la

    # Leader line from marker to label start
    ax.plot([lon, lx], [lat, ly], color=INK_MUTE, linewidth=0.8,
            linestyle='-', zorder=9, alpha=0.6)

    # Number badge on marker
    ax.text(lon, lat, str(idx + 1), ha='center', va='center',
            fontsize=8, fontweight='bold', color=MARKER_RING, zorder=12)

    # Label box
    label_text = f"{city}\n{distr}\n{date}"
    ax.text(lx, ly, label_text, ha=ha, va=va,
            fontsize=8.5, color=INK, fontweight='normal',
            linespacing=1.45,
            bbox=dict(
                boxstyle="round,pad=0.4",
                facecolor='#161616',
                edgecolor=MARKER_COLOR,
                linewidth=0.9,
                alpha=0.96
            ),
            zorder=13)

# ────────────────────────────────────────────────────────────────────────
# EMPILS HQ MARKER (anchor, different shape — diamond)
# ────────────────────────────────────────────────────────────────────────
hq_name, hq_lat, hq_lon = EMPILS_HQ
ax.scatter([hq_lon], [hq_lat], s=110, color=INK, edgecolors=MARKER_RING,
           linewidths=1.5, zorder=10, marker='D')
ax.text(hq_lon + 1.5, hq_lat + 0.5, f"Эмпилс\n{hq_name}",
        ha='left', va='bottom', fontsize=8, color=INK_SOFT, fontweight='600',
        bbox=dict(boxstyle="round,pad=0.3", facecolor='#0A0A0A',
                  edgecolor=INK_MUTE, linewidth=0.6, alpha=0.95),
        zorder=13)

# ────────────────────────────────────────────────────────────────────────
# LEGEND (top-right)
# ────────────────────────────────────────────────────────────────────────
legend_x = 105.5
legend_y = 58.0
ax.text(legend_x, legend_y + 0.6, "УСЛОВНЫЕ ОБОЗНАЧЕНИЯ",
        fontsize=8.5, fontweight='bold', color=INK, ha='left', va='top',
        zorder=15)

# Target region swatch
ax.add_patch(mpatches.Rectangle((legend_x, legend_y - 0.2), 1.2, 0.6,
                                 facecolor=TARGET_FILL, edgecolor=TARGET_BORDER,
                                 linewidth=0.8, zorder=15))
ax.text(legend_x + 1.5, legend_y + 0.1, "Целевой регион",
        fontsize=8, color=INK, ha='left', va='center', zorder=15)

# Distributor marker
ax.scatter([legend_x + 0.6], [legend_y - 1.4], s=100, color=MARKER_COLOR,
           edgecolors=MARKER_RING, linewidths=1.5, zorder=15, marker='o')
ax.text(legend_x + 1.5, legend_y - 1.4, "Дистрибьютор (встреча)",
        fontsize=8, color=INK, ha='left', va='center', zorder=15)

# HQ marker
ax.scatter([legend_x + 0.6], [legend_y - 2.4], s=80, color=INK,
           edgecolors=MARKER_RING, linewidths=1.2, zorder=15, marker='D')
ax.text(legend_x + 1.5, legend_y - 2.4, "Эмпилс (Ростов-на-Дону)",
        fontsize=8, color=INK, ha='left', va='center', zorder=15)

# Non-target region
ax.add_patch(mpatches.Rectangle((legend_x, legend_y - 3.4), 1.2, 0.6,
                                 facecolor=RU_FILL, edgecolor=RU_EDGE,
                                 linewidth=0.5, zorder=15))
ax.text(legend_x + 1.5, legend_y - 3.1, "Прочие регионы РФ",
        fontsize=8, color=INK, ha='left', va='center', zorder=15)

# Target regions list box
ax.text(legend_x, legend_y - 4.2, "ЦЕЛЕВЫЕ РЕГИОНЫ:",
        fontsize=7.5, fontweight='bold', color=INK_MUTE, ha='left', va='top',
        zorder=15)
ax.text(legend_x, legend_y - 4.8,
        "1. Алтайский край\n2. Красноярский край\n3. Республика Бурятия\n4. Омская область\n5. Республика Татарстан",
        fontsize=7.5, color=INK, ha='left', va='top',
        linespacing=1.5, zorder=15)

# ────────────────────────────────────────────────────────────────────────
# STYLING
# ────────────────────────────────────────────────────────────────────────
ax.set_xticks([])
ax.set_yticks([])
for spine in ax.spines.values():
    spine.set_visible(False)

# Subtle frame
rect = mpatches.Rectangle((38, 48), 77, 12, fill=False, edgecolor=HAIRLINE,
                          linewidth=1.0, zorder=1)
ax.add_patch(rect)

# Save
OUT = "/home/z/my-project/download/slides_titan/assets/white_spots_map.png"
fig.savefig(OUT, dpi=200, facecolor=BG, bbox_inches='tight', pad_inches=0.15)
print(f"Saved: {OUT}")
plt.close(fig)
