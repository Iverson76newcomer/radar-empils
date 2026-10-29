"""Parse Palizh homepage HTML and extract main content."""
import json
import re
from html.parser import HTMLParser

with open('/home/z/my-project/scripts/palizh_home.json', 'r', encoding='utf-8') as f:
    data = json.load(f)

print("=== TITLE ===")
print(data.get('data', {}).get('title', 'N/A'))
print()
print("=== DESCRIPTION ===")
print(data.get('data', {}).get('description', 'N/A'))
print()
print("=== URL ===")
print(data.get('data', {}).get('url', 'N/A'))
print()

html = data.get('data', {}).get('html', '')
print(f"=== HTML length: {len(html)} ===\n")

# Strip scripts and styles
clean = re.sub(r'<script[^>]*>.*?</script>', '', html, flags=re.DOTALL|re.IGNORECASE)
clean = re.sub(r'<style[^>]*>.*?</style>', '', clean, flags=re.DOTALL|re.IGNORECASE)
clean = re.sub(r'<!--.*?-->', '', clean, flags=re.DOTALL)

# Replace block elements with newlines
clean = re.sub(r'</(div|p|section|article|header|footer|nav|ul|ol|li|h[1-6]|br)>', '\n', clean, flags=re.IGNORECASE)
clean = re.sub(r'<br\s*/?>', '\n', clean, flags=re.IGNORECASE)
# Strip all tags
text = re.sub(r'<[^>]+>', ' ', clean)
# Decode entities
text = text.replace('&nbsp;', ' ').replace('&amp;', '&').replace('&lt;', '<').replace('&gt;', '>').replace('&quot;', '"').replace('&laquo;', '«').replace('&raquo;', '»').replace('&mdash;', '—').replace('&ndash;', '–').replace('&hellip;', '…')
# Collapse whitespace per line
lines = [l.strip() for l in text.split('\n')]
lines = [l for l in lines if l]
text_clean = '\n'.join(lines)
# Remove duplicate consecutive lines
seen = set()
final = []
for l in lines:
    if l not in seen or len(l) > 60:
        final.append(l)
        seen.add(l)

print("=== TEXT CONTENT (cleaned) ===")
print('\n'.join(final)[:10000])
