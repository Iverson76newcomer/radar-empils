import openpyxl
from openpyxl import load_workbook
from collections import defaultdict
import re

# === 1. Read assortment ===
wb1 = load_workbook("/home/z/my-project/upload/Ассортимент_Олимп_300_SKU.xlsx", data_only=True)
ws1 = wb1['300 SKU — детально']

# Read header (row 3) and data (rows 4+)
header = [c.value for c in ws1[3]]
print("Ассортимент header:", header)

assortment_items = []
for row in ws1.iter_rows(min_row=4, max_row=ws1.max_row, values_only=True):
    if row[0] is None:
        continue
    assortment_items.append({
        'idx': row[0],
        'sphere': row[1],
        'purpose': row[2],
        'base': row[3],
        'brand': row[4],
        'type': row[5],
        'name': row[6],
        'grp_code': row[7],
        'abc': row[8]
    })

print(f"\nВсего в ассортименте: {len(assortment_items)} SKU")

# Unique brands in assortment
brands_in_assort = sorted(set(item['brand'] for item in assortment_items if item['brand']))
print(f"\nБренды в ассортименте ({len(brands_in_assort)}):")
for b in brands_in_assort:
    cnt = sum(1 for it in assortment_items if it['brand'] == b)
    print(f"  '{b}': {cnt} SKU")

# === 2. Read price list ===
wb2 = load_workbook("/home/z/my-project/upload/21.08.2026 РОССИЯ с НДС, руб..xlsx", data_only=True)
print(f"\n\nЛисты прайс-листа ({len(wb2.sheetnames)}):")
for sn in wb2.sheetnames:
    print(f"  '{sn}'")

# Build price list dict
price_dict = defaultdict(list)  # brand_lower -> list of (name, фасовка_H, базовая_M, отгруз_N, sheet, row)

for sn in wb2.sheetnames:
    ws = wb2[sn]
    current_name = None
    for row_idx, row in enumerate(ws.iter_rows(min_row=14, max_row=ws.max_row, values_only=True), 14):
        # Check if this is a "name" row (col F has text, col M/N empty)
        f_val = row[5] if len(row) > 5 else None
        m_val = row[12] if len(row) > 12 else None
        n_val = row[13] if len(row) > 13 else None
        h_val = row[7] if len(row) > 7 else None  # фасовка
        
        # Update current_name if F is a string and looks like a product name (not СТО code)
        if f_val is not None and isinstance(f_val, str) and not f_val.startswith('СТО'):
            current_name = f_val.strip()
        
        # If has price, add to dict
        if (m_val is not None and isinstance(m_val, (int, float))) or (n_val is not None and isinstance(n_val, (int, float))):
            price_dict[sn].append({
                'name': current_name,
                'fасовка': h_val,
                'базовая': m_val,
                'отгрузочная': n_val,
                'sheet': sn,
                'row': row_idx
            })

print(f"\n\n=== Записей с ценами по листам ===")
for sn, items in price_dict.items():
    print(f"  '{sn}': {len(items)} цен")
