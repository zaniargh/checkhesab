import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
sys.path.append('.')
from app import parse_html, parse_excel

with open(r'd:\Checkhesab\nbnbnbnbnbnbnbnbn.Html', 'rb') as f:
    html_rows = parse_html(f.read())
with open(r'd:\Checkhesab\1-4 من.xlsx', 'rb') as f:
    excel_rows = parse_excel(f.read(), '1-4 من.xlsx')

target_amounts = [2000000000, 10000000000, 50000000000, 120000000000, 150000000000, 300000000000]
for amt in target_amounts:
    excel_matches = [t for t in excel_rows if t['amount'] and abs(t['amount'] - amt) < 1000]
    html_matches = [r for r in html_rows if r['credit'] and abs(r['credit'] - amt) < 1000]
    print(f'Amount {int(amt):,}:')
    print(f'  HTML rows: {len(html_matches)}')
    print(f'  Excel deposit rows: {len(excel_matches)}')
    for t in excel_matches[:3]:
        print(f'    -> Row {t["row_num"]}: {t["desc"][:60]} | codes={t["all_codes"][:2]}')
