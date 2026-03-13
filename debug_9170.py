import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
sys.path.append('.')
from app import parse_html, parse_excel

with open(r'd:\Checkhesab\nbnbnbnbnbnbnbnbn.Html', 'rb') as f:
    html_rows = parse_html(f.read())
with open(r'd:\Checkhesab\1-4 من.xlsx', 'rb') as f:
    excel_rows = parse_excel(f.read(), '1-4 من.xlsx')

# Check rows with codes 9170 and 8511
print('=== HTML rows containing 9170 or 8511 ===')
for r in html_rows:
    desc = r['desc']
    if '9170' in desc or '8511' in desc:
        crd = int(r['credit'])
        dbt = int(r['debit'])
        print(f'  credit={crd:,} debit={dbt:,} date={r["date"]} codes={r["codes"]}')
        print(f'  desc: {desc[:100]}')
        print()

print()
print('=== Excel rows with 9170 ===')
for t in excel_rows:
    combined = t.get('ref','') + t.get('desc','') + str(t.get('all_codes',''))
    if '9170' in combined:
        amt = int(t['amount'])
        print(f'  Row {t["row_num"]}: amt={amt:,} | {t["desc"][:80]}')

print()
print('=== Excel rows with 8511 ===')
for t in excel_rows:
    combined = t.get('ref','') + t.get('desc','') + str(t.get('all_codes',''))
    if '8511' in combined:
        amt = int(t['amount'])
        print(f'  Row {t["row_num"]}: amt={amt:,} | {t["desc"][:80]}')
