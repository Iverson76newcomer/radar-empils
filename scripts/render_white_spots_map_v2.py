"""
Render the white-spots map v2 — adds 3 new regions:
- Московская область (ООО Вертикаль, ООО Стройтерминал центр красок)
- Воронежская область (ООО Вертикаль)
- Липецкая область (ООО Стройбаза-спектр)

Changes from v1:
- Extended xlim from (38,115) to (35,115) to include Moscow
- Added 3 new target regions to the highlight list
- New regions use SAND markers (vs RED for existing 5 with scheduled meetings)
- Legend updated with "Новые регионы (потенциальные)" section
"""
import geopandas as gpd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import matplotlib.patches as mpatches
import os

fm.fontManager.addfont('/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf')
plt.rcParams['font.sans-serif'] = ['DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False

# Dark palette (matches existing slide)
BG          = "#0A0A0A"
INK         = "#FFFFFF"
INK_SOFT    = "#CCCCCC"
INK_MUTE    = "#888888"
HAIRLINE    = "#333333"
COUNTRY_FILL = "#161616"
COUNTRY_EDGE = "#2A2A2A"
RU_FILL = "#161616"
RU_EDGE = "#2A2A2A"

# Existing target regions — red border
TARGET_FILL = "#1E1E1E"
TARGET_EDGE_RED = "#E63946"  # red for existing 5

# New target regions — yellow/sand border
TARGET_EDGE_NEW = "#FFD60A"  # yellow for new 3

# Markers
MARKER_COLOR_EXISTING = "#E63946"  # red for existing
MARKER_COLOR_NEW = "#FFD60A"       # yellow for new
MARKER_RING  = "#0A0A0A"

URL = "https://naciscdn.org/naturalearth/10m/cultural/ne_10m_admin_1_states_provinces.zip"
print("Loading admin_1…")
gdf = gpd.read_file(URL)
ru = gdf[gdf['admin'] == 'Russia'].copy()
print(f"Russia subjects: {len(ru)}")

# Existing 5 target regions
existing_names = ['Altay', 'Krasnoyarsk', 'Buryat', 'Omsk', 'Tatarstan']
# New 3 target regions
new_names = ['Moskovskaya', 'Moskva', 'Voronezh', 'Lipetsk']
all_target_names = existing_names + new_names

target_existing = ru[ru['name'].isin(existing_names)].copy()
target_new = ru[ru['name'].isin(new_names)].copy()
other_ru = ru[~ru['name'].isin(all_target_names)].copy()

# Existing cities (5 with scheduled meetings)
CITIES_EXISTING = [
    ("Барнаул",       53.34, 83.78,  "Олимп",                       "13.10.2026",     0),
    ("Красноярск",    56.01, 92.85,  "ООО «ОЧАГ»",                  "05–11.11.2026",  1),
    ("Улан-Удэ",      51.83, 107.58, "Вегос-М",                     "19–25.11.2026",  2),
    ("Омск",          54.99, 73.37,  "ИП Шишков",                   "ноябрь 2026",    3),
    ("Казань",        55.79, 49.12,  "ТС Агава, Карат, Колорит",    "ноябрь 2026",    4),
]

# New cities (3 potential — no meetings scheduled yet)
CITIES_NEW = [
    ("Москва",        55.75, 37.62,  "Вертикаль, Стройтерминал",     "", 5),
    ("Воронеж",       51.67, 39.20,  "Вертикаль",                    "", 6),
    ("Липецк",        52.61, 39.58,  "Стройбаза-спектр",             "", 7),
]

EMPILS_HQ = ("Ростов-на-Дону", 47.23, 39.70)

# ─── Figure ───────────────────────────────────────────────────────────
fig = plt.figure(figsize=(14.5, 7.6), dpi=200, facecolor=BG)
ax = fig.add_axes([0.02, 0.02, 0.96, 0.96])
ax.set_facecolor(BG)

# Plot non-target regions
other_ru.plot(ax=ax, color=RU_FILL, edgecolor=RU_EDGE, linewidth=0.35, zorder=2)
other_countries = gdf[(gdf['admin'] != 'Russia')].copy()
other_countries.plot(ax=ax, color=COUNTRY_FILL, edgecolor=COUNTRY_EDGE, linewidth=0.25, zorder=1)

# Plot EXISTING target regions — red border
target_existing.plot(ax=ax, color=TARGET_FILL, edgecolor=TARGET_EDGE_RED, linewidth=1.4, zorder=4)

# Plot NEW target regions — yellow border
target_new.plot(ax=ax, color=TARGET_FILL, edgecolor=TARGET_EDGE_NEW, linewidth=1.4, zorder=4)

# Map extent — extended west to include Moscow (lon 37.62)
ax.set_xlim(35, 115)
ax.set_ylim(48, 60)

# ─── Existing city markers (RED) ──────────────────────────────────────
LABEL_OFFSETS = {
    "Барнаул":     (+2.0,  -1.5, "left", "top"),
    "Красноярск":  (+2.0,  +1.2, "left", "bottom"),
    "Улан-Удэ":    (-2.5,  -1.5, "right","top"),
    "Омск":        (-2.5,  +1.5, "right","bottom"),
    "Казань":      (+2.0,  +1.0, "left", "bottom"),
}

for city, lat, lon, distr, date, idx in CITIES_EXISTING:
    ax.scatter([lon], [lat], s=130, color=MARKER_COLOR_EXISTING, edgecolors=MARKER_RING,
               linewidths=2.0, zorder=10, marker='o')
    ax.scatter([lon], [lat], s=30, color=MARKER_RING, zorder=11, marker='o')
    lo, la, ha, va = LABEL_OFFSETS[city]
    lx, ly = lon + lo, lat + la
    ax.plot([lon, lx], [lat, ly], color=INK_MUTE, linewidth=0.8, linestyle='-', zorder=9, alpha=0.6)
    ax.text(lon, lat, str(idx + 1), ha='center', va='center', fontsize=8, fontweight='bold', color=MARKER_RING, zorder=12)
    label_text = f"{city}\n{distr}\n{date}"
    ax.text(lx, ly, label_text, ha=ha, va=va, fontsize=8.5, color=INK, fontweight='normal',
            linespacing=1.45,
            bbox=dict(boxstyle="round,pad=0.4", facecolor='#161616', edgecolor=MARKER_COLOR_EXISTING, linewidth=0.9, alpha=0.96),
            zorder=13)

# ─── New city markers (YELLOW/SAND) — potential regions ───────────────
NEW_LABEL_OFFSETS = {
    "Москва":   (+2.0,  +1.0, "left", "bottom"),   # label to upper-right
    "Воронеж":  (+2.0,  -1.5, "left", "top"),       # label to lower-right
    "Липецк":   (-2.5,  +1.0, "right","bottom"),    # label to upper-left
}

for city, lat, lon, distr, date, idx in CITIES_NEW:
    ax.scatter([lon], [lat], s=130, color=MARKER_COLOR_NEW, edgecolors=MARKER_RING,
               linewidths=2.0, zorder=10, marker='s')  # square marker for new
    ax.scatter([lon], [lat], s=30, color=MARKER_RING, zorder=11, marker='s')
    lo, la, ha, va = NEW_LABEL_OFFSETS[city]
    lx, ly = lon + lo, lat + la
    ax.plot([lon, lx], [lat, ly], color=INK_MUTE, linewidth=0.8, linestyle='--', zorder=9, alpha=0.6)
    ax.text(lon, lat, str(idx + 1), ha='center', va='center', fontsize=8, fontweight='bold', color=MARKER_RING, zorder=12)
    label_text = f"{city}\n{distr}"
    ax.text(lx, ly, label_text, ha=ha, va=va, fontsize=8.5, color=INK, fontweight='normal',
            linespacing=1.45,
            bbox=dict(boxstyle="round,pad=0.4", facecolor='#161616', edgecolor=MARKER_COLOR_NEW, linewidth=0.9, alpha=0.96),
            zorder=13)

# ─── Empils HQ marker ─────────────────────────────────────────────────
hq_name, hq_lat, hq_lon = EMPILS_HQ
ax.scatter([hq_lon], [hq_lat], s=110, color=INK, edgecolors=MARKER_RING,
           linewidths=1.5, zorder=10, marker='D')
ax.text(hq_lon + 1.5, hq_lat + 0.5, f"Эмпилс\n{hq_name}",
        ha='left', va='bottom', fontsize=8, color=INK_SOFT, fontweight='600',
        bbox=dict(boxstyle="round,pad=0.3", facecolor='#0A0A0A', edgecolor=INK_MUTE, linewidth=0.6, alpha=0.95),
        zorder=13)

# ─── Legend (top-right, compact 2-column) ─────────────────────────────
legend_x = 105.5
legend_y = 58.0
ax.text(legend_x, legend_y + 0.6, "УСЛОВНЫЕ ОБОЗНАЧЕНИЯ",
        fontsize=8.5, fontweight='bold', color=INK, ha='left', va='top', zorder=15)

# Existing target region (red border)
ax.add_patch(mpatches.Rectangle((legend_x, legend_y - 0.2), 1.2, 0.6,
                                 facecolor=TARGET_FILL, edgecolor=TARGET_EDGE_RED, linewidth=0.8, zorder=15))
ax.text(legend_x + 1.5, legend_y + 0.1, "Регион встречи (запланир.)",
        fontsize=7.5, color=INK, ha='left', va='center', zorder=15)

# New target region (yellow border)
ax.add_patch(mpatches.Rectangle((legend_x, legend_y - 1.0), 1.2, 0.6,
                                 facecolor=TARGET_FILL, edgecolor=TARGET_EDGE_NEW, linewidth=0.8, zorder=15))
ax.text(legend_x + 1.5, legend_y - 0.7, "Новый регион (потенциальн.)",
        fontsize=7.5, color=INK, ha='left', va='center', zorder=15)

# Existing distributor marker (red circle)
ax.scatter([legend_x + 0.6], [legend_y - 1.9], s=80, color=MARKER_COLOR_EXISTING,
           edgecolors=MARKER_RING, linewidths=1.3, zorder=15, marker='o')
ax.text(legend_x + 1.5, legend_y - 1.9, "Встреча запланирована",
        fontsize=7.5, color=INK, ha='left', va='center', zorder=15)

# New distributor marker (yellow square)
ax.scatter([legend_x + 0.6], [legend_y - 2.7], s=80, color=MARKER_COLOR_NEW,
           edgecolors=MARKER_RING, linewidths=1.3, zorder=15, marker='s')
ax.text(legend_x + 1.5, legend_y - 2.7, "Потенциальный партнёр",
        fontsize=7.5, color=INK, ha='left', va='center', zorder=15)

# HQ marker
ax.scatter([legend_x + 0.6], [legend_y - 3.5], s=65, color=INK,
           edgecolors=MARKER_RING, linewidths=1.0, zorder=15, marker='D')
ax.text(legend_x + 1.5, legend_y - 3.5, "Эмпилс (Ростов-на-Дону)",
        fontsize=7.5, color=INK, ha='left', va='center', zorder=15)

# Non-target region
ax.add_patch(mpatches.Rectangle((legend_x, legend_y - 4.3), 1.2, 0.6,
                                 facecolor=RU_FILL, edgecolor=RU_EDGE, linewidth=0.5, zorder=15))
ax.text(legend_x + 1.5, legend_y - 4.0, "Прочие регионы РФ",
        fontsize=7.5, color=INK, ha='left', va='center', zorder=15)

# ─── Styling ──────────────────────────────────────────────────────────
ax.set_xticks([])
ax.set_yticks([])
for spine in ax.spines.values():
    spine.set_visible(False)

rect = mpatches.Rectangle((35, 48), 80, 12, fill=False, edgecolor=HAIRLINE, linewidth=1.0, zorder=1)
ax.add_patch(rect)

OUT = "/home/z/my-project/download/slides_titan/assets/white_spots_map_v2.png"
fig.savefig(OUT, dpi=200, facecolor=BG, bbox_inches='tight', pad_inches=0.15)
print(f"Saved: {OUT} ({os.path.getsize(OUT)} bytes)")
plt.close(fig)
