import urllib.request
import os

url = "https://www.complexconstruction.ru/upload/iblock/793/Tikkurila_Catalog.pdf"
out_path = "tinting_research/tikkurila_catalog.pdf"

headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
    'Accept': 'application/pdf,*/*',
}

try:
    req = urllib.request.Request(url, headers=headers)
    with urllib.request.urlopen(req, timeout=30) as resp:
        with open(out_path, 'wb') as f:
            f.write(resp.read())
    print(f"✓ Tikkurila PDF сохранён: {out_path}")
    print(f"  Размер: {os.path.getsize(out_path):,} байт")
except Exception as e:
    print(f"Error: {e}")

# Извлечь текст
import subprocess
result = subprocess.run(['pdftotext', '-layout', out_path, out_path.replace('.pdf', '.txt')], 
                       capture_output=True, text=True)
print(f"Text extracted: {result.returncode == 0}")

# Прочитать
with open(out_path.replace('.pdf', '.txt'), 'r', encoding='utf-8') as f:
    text = f.read()
print(f"Длина текста: {len(text)}")

# Найти секции про колеровку
keywords = ['колеровочное оборудование', 'колеровочная машина', 'колеровочный', 'колеровка', 
            'дозирующая', 'смеситель', 'спектрофотометр', 'дозатор', 'колорант', 
            'маркетинговые материалы', 'магазин', 'розничн', 'точка продаж', 'дилер', 'POS']

print("\n=== Контекст по ключевым словам ===")
for kw in keywords:
    idx = 0
    count = 0
    while count < 2:
        idx = text.lower().find(kw.lower(), idx)
        if idx < 0:
            break
        start = max(0, idx - 100)
        end = min(len(text), idx + 500)
        context = text[start:end].strip()
        if len(context) > 50:
            print(f"\n[{kw}] ...{context}...")
            count += 1
        idx += len(kw)
