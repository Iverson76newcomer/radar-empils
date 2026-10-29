"""
Replace the embedded map image (ppt/media/image2.png) inside the user's
uploaded PPTX with the freshly rendered one (legend moved to bottom-left).
No XML color changes, no slide-layout changes — only swaps the PNG bytes.
"""
import zipfile, shutil, os

SRC = "/home/z/my-project/upload/Презентация_для_ГД_02.10.2026.pptx"
OUT = "/home/z/my-project/download/Презентация_для_ГД_02.10.2026.pptx"
NEW_MAP = "/home/z/my-project/download/slides_titan/assets/white_spots_map.png"

# Verify new map exists
assert os.path.exists(NEW_MAP), f"Missing: {NEW_MAP}"
print(f"New map: {os.path.getsize(NEW_MAP)} bytes")

# Read all entries from SRC, copy to OUT replacing image2.png
if os.path.exists(OUT):
    os.remove(OUT)

replaced = False
with zipfile.ZipFile(SRC, 'r') as zin, \
     zipfile.ZipFile(OUT, 'w', zipfile.ZIP_DEFLATED) as zout:
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename == 'ppt/media/image2.png':
            with open(NEW_MAP, 'rb') as f:
                data = f.read()
            replaced = True
            print(f"Replaced: {item.filename}  (orig {item.file_size} bytes → new {len(data)} bytes)")
        zout.writestr(item, data)

if not replaced:
    print("WARNING: image2.png was NOT found in source PPTX!")
else:
    print(f"\nSaved: {OUT} ({os.path.getsize(OUT)} bytes)")
