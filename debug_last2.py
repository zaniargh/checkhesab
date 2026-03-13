import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
sys.path.append('.')
from app import parse_html, parse_excel

with open(r'd:\Checkhesab\nbnbnbnbnbnbnbnbn.Html', 'rb') as f:
    html_rows = parse_html(f.read())

credit_rows = [r for r in html_rows if r['credit'] > 0]
print(f'Total HTML credit rows: {len(credit_rows)}')
for r in credit_rows[-5:]:
    print(f'  Date={r["date"]} Amount={int(r["credit"]):,} desc_preview={r["desc"][:60]}')
