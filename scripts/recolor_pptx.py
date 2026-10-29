"""
Recolor the user's Презентация_для_ГД.pptx from dark/high-contrast to soft/light.
Keeps ALL text content and ALL shape positions / sizes intact — only changes
fill / line / font colors by remapping every <a:srgbClr val="XXXXXX"/> in every
XML part of the .pptx archive.

Also replaces the embedded map PNG (image2.png) with a light-palette render.
"""
import zipfile, shutil, re, os
from pathlib import Path

SRC = "/home/z/my-project/upload/Презентация_для_ГД.pptx"
OUT = "/home/z/my-project/download/Презентация_для_ГД.pptx"
NEW_MAP = "/home/z/my-project/download/slides_titan/assets/white_spots_map_light.png"

# ── Color mapping (uppercase hex, no #) ─────────────────────────────────────
# Dark palette → Soft light palette
COLOR_MAP = {
    # ── Backgrounds ──
    '0A0A0A': 'F5F2EC',   # main bg → warm cream
    '0E0E0E': 'F5F2EC',
    '161616': 'FFFFFF',   # card surface → white
    '222222': 'EFE9DD',   # divider → soft cream
    '333333': 'D9D2C2',   # border → warm light gray

    # ── Text (inverted for light bg) ──
    'FFFFFF': '2F3A48',   # was white text on dark → dark slate
    'CCCCCC': '5A6776',   # was light gray → medium slate
    'DDDDDD': '8A95A3',   # was very light gray → soft slate
    'D1D5DB': '8A95A3',
    '888888': '8A95A3',   # was gray → soft slate (similar enough)
    '6B7280': '8A95A3',

    # ── Brand accents — keep identity but desaturate ──
    'E63946': 'B5614A',   # red → soft terracotta
    'FFD60A': 'C9A77C',   # yellow → warm sand
    '2ECC71': '8BAA8A',   # green → sage
    'FF8C00': 'D4A574',   # orange → warm amber

    # ── Pure black variants used for shapes ──
    '000000': '2F3A48',
}

# Build case-insensitive lookup
def lookup(hex6):
    return COLOR_MAP.get(hex6.upper())

# ── Work on a temp copy ─────────────────────────────────────────────────────
TMP = "/tmp/recolor_work"
if os.path.exists(TMP):
    shutil.rmtree(TMP)
os.makedirs(TMP)

# Unzip source
with zipfile.ZipFile(SRC, 'r') as zin:
    zin.extractall(TMP)

# ── Walk every XML file and remap colors ────────────────────────────────────
xml_files = []
for root, _, files in os.walk(TMP):
    for f in files:
        if f.endswith('.xml') or f.endswith('.rels'):
            xml_files.append(os.path.join(root, f))

total_replacements = 0
patterns_replaced = {}

# Pattern: srgbClr val="XXXXXX"  (case-insensitive hex)
pat = re.compile(r'(srgbClr\s+val=")([0-9A-Fa-f]{6})(")')

for path in xml_files:
    with open(path, 'r', encoding='utf-8', errors='ignore') as f:
        content = f.read()
    if 'srgbClr' not in content:
        continue

    def repl(m):
        global total_replacements
        old = m.group(2).upper()
        new = lookup(old)
        if new:
            total_replacements += 1
            patterns_replaced[old] = patterns_replaced.get(old, 0) + 1
            return f'{m.group(1)}{new}{m.group(3)}'
        return m.group(0)

    new_content = pat.sub(repl, content)
    if new_content != content:
        with open(path, 'w', encoding='utf-8') as f:
            f.write(new_content)

print(f"XML files scanned: {len(xml_files)}")
print(f"Total srgbClr replacements: {total_replacements}")
for old, n in sorted(patterns_replaced.items(), key=lambda x: -x[1]):
    new = lookup(old)
    print(f"  #{old} → #{new}   x{n}")

# ── Replace the map image (image2.png = 431385 bytes, 2868x1135) ────────────
map_path = os.path.join(TMP, 'ppt', 'media', 'image2.png')
if os.path.exists(map_path):
    orig_size = os.path.getsize(map_path)
    shutil.copy(NEW_MAP, map_path)
    new_size = os.path.getsize(map_path)
    print(f"\nMap image replaced: image2.png")
    print(f"  Original: {orig_size} bytes")
    print(f"  New:      {new_size} bytes")
else:
    print(f"WARNING: image2.png not found in pptx")

# ── Re-zip into OUT ─────────────────────────────────────────────────────────
if os.path.exists(OUT):
    os.remove(OUT)

with zipfile.ZipFile(OUT, 'w', zipfile.ZIP_DEFLATED) as zout:
    for root, _, files in os.walk(TMP):
        for f in files:
            full = os.path.join(root, f)
            arc = os.path.relpath(full, TMP)
            # Use forward slashes in archive
            arc = arc.replace(os.sep, '/')
            zout.write(full, arc)

print(f"\nSaved: {OUT} ({os.path.getsize(OUT)} bytes)")
