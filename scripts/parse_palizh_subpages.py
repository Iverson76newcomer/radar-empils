"""Extract clean text from Palizh sub-pages."""
import json, re

for fname, label in [
    ('palizh_company.json', 'COMPANY'),
    ('palizh_products.json', 'PRODUCTS'),
]:
    with open(f'/home/z/my-project/scripts/{fname}') as f:
        d = json.load(f)
    data = d.get('data', {})
    print(f"\n{'='*80}\n{label}: {data.get('url', 'N/A')}\n{'='*80}")
    html = data.get('html', '')
    clean = re.sub(r'<script[^>]*>.*?</script>', '', html, flags=re.DOTALL|re.IGNORECASE)
    clean = re.sub(r'<style[^>]*>.*?</style>', '', clean, flags=re.DOTALL|re.IGNORECASE)
    clean = re.sub(r'<!--.*?-->', '', clean, flags=re.DOTALL)
    clean = re.sub(r'</(div|p|section|article|header|footer|nav|ul|ol|li|h[1-6]|br)>', '\n', clean, flags=re.IGNORECASE)
    clean = re.sub(r'<br\s*/?>', '\n', clean, flags=re.IGNORECASE)
    text = re.sub(r'<[^>]+>', ' ', clean)
    text = text.replace('&nbsp;', ' ').replace('&amp;', '&').replace('&laquo;', '«').replace('&raquo;', '»').replace('&mdash;', '—').replace('&ndash;', '–')
    lines = [l.strip() for l in text.split('\n')]
    lines = [l for l in lines if l]
    # dedupe
    seen = set()
    final = []
    for l in lines:
        if l not in seen or len(l) > 60:
            final.append(l)
            seen.add(l)
    print('\n'.join(final)[:6000])
