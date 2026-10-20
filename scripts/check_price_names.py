import openpyxl
from openpyxl import load_workbook
from collections import defaultdict

wb2 = load_workbook("/home/z/my-project/upload/21.08.2026 РОССИЯ с НДС, руб..xlsx", data_only=True)

price_dict = defaultdict(list)
for sn in wb2.sheetnames:
    ws = wb2[sn]
    current_name = None
    for row_idx, row in enumerate(ws.iter_rows(min_row=14, max_row=ws.max_row, values_only=True), 14):
        f_val = row[5] if len(row) > 5 else None
        m_val = row[12] if len(row) > 12 else None
        n_val = row[13] if len(row) > 13 else None
        h_val = row[7] if len(row) > 7 else None
        if f_val is not None and isinstance(f_val, str) and not f_val.startswith('СТО'):
            current_name = f_val.strip()
        if (m_val is not None and isinstance(m_val, (int, float))) or (n_val is not None and isinstance(n_val, (int, float))):
            price_dict[sn].append({
                'name': current_name,
                'fасовка': h_val,
                'базовая': m_val,
                'отгрузочная': n_val,
            })

# Search Расцвет for "3 в 1" or "3в1" or "3в 1"
print("=== Расцвет: search variants ===")
for p in price_dict['Расцвет']:
    if p['name'] and ('3 в 1' in p['name'] or '3в1' in p['name'] or '3 в1' in p['name']):
        if 'белая' in p['name'].lower() or 'белая' in p['name'].lower():
            print(f"  фас={p['fасовка']}, отгруз={p['отгрузочная']}, name='{p['name']}'")

print("\n=== Расцвет: первые 15 имён ===")
seen_names = []
for p in price_dict['Расцвет']:
    if p['name'] not in seen_names:
        seen_names.append(p['name'])
        if len(seen_names) <= 15:
            print(f"  '{p['name']}'")

print("\n=== NEWTONE: первые 20 имён ===")
seen_names = []
for p in price_dict['NEWTONE']:
    if p['name'] not in seen_names:
        seen_names.append(p['name'])
        if len(seen_names) <= 20:
            print(f"  '{p['name']}'")

print("\n=== PROFIWOOD: первые 15 имён ===")
seen_names = []
for p in price_dict['PROFIWOOD']:
    if p['name'] not in seen_names:
        seen_names.append(p['name'])
        if len(seen_names) <= 15:
            print(f"  '{p['name']}'")

# Check all names containing "3 в 1" or "3в1" across all sheets
print("\n=== All sheets: '3 в 1' or '3в1' ===")
for sn in wb2.sheetnames:
    for p in price_dict[sn]:
        if p['name'] and ('3 в 1' in p['name'].lower() or '3в1' in p['name'].lower()):
            print(f"  [{sn}] '{p['name']}' фас={p['fасовка']} отгруз={p['отгрузочная']}")
