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
        if f_val is not None and isinstance(f_val, str) and not f_val.startswith('СТО') and not f_val.startswith('ТУ') and not f_val.startswith('ГОСТ'):
            current_name = f_val.strip()
        if (m_val is not None and isinstance(m_val, (int, float))) or (n_val is not None and isinstance(n_val, (int, float))):
            price_dict[sn].append({
                'name': current_name,
                'fасовка': h_val,
                'базовая': m_val,
                'отгрузочная': n_val,
            })

# Look at all PROFIWOOD лак entries
print("=== PROFIWOOD: all лак entries ===")
for p in price_dict['PROFIWOOD']:
    if p['name'] and ('лак' in p['name'].lower() or 'паркет' in p['name'].lower() or 'яхт' in p['name'].lower()):
        print(f"  '{p['name']}' фас={p['fасовка']} отгруз={p['отгрузочная']}")

# Look at PROFIWOOD assortment items 246, 252
wb1 = load_workbook("/home/z/my-project/upload/Ассортимент_Олимп_300_SKU.xlsx", data_only=True)
ws1 = wb1['300 SKU — детально']
print("\n=== Assortment items 241-252 (PROFIWOOD) ===")
for row in ws1.iter_rows(min_row=4, max_row=ws1.max_row, values_only=True):
    if row[0] is None:
        continue
    if 240 <= row[0] <= 260:
        print(f"  [{row[0]}] {row[4]} | {row[6]}")
