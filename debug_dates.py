import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
sys.path.append('.')
from app import parse_html, parse_excel

with open(r'd:\Checkhesab\nbnbnbnbnbnbnbnbn.Html', 'rb') as f:
    html_rows = parse_html(f.read())
with open(r'd:\Checkhesab\1-4 من.xlsx', 'rb') as f:
    excel_rows = parse_excel(f.read(), '1-4 من.xlsx')

credit_rows = [r for r in html_rows if r['credit'] > 0]
dates_html = sorted([r['date'] for r in credit_rows if r['date']])
dates_xls = sorted([t['date'] for t in excel_rows if t['date']])

print('=== HTML Credit dates range ===')
print(f'First: {dates_html[0]}  Last: {dates_html[-1]}')

print('\n=== Excel dates range ===')
print(f'First: {dates_xls[0]}  Last: {dates_xls[-1]}')

print('\n=== Last 2 HTML credit rows ===')
for r in credit_rows[-2:]:
    desc_parts = r['desc'].split(' | ')
    tx_type = desc_parts[2].strip() if len(desc_parts) > 2 else ''
    print(f'  Date={r["date"]}  Amount={int(r["credit"]):,}  Type={tx_type}')
