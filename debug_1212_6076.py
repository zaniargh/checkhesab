import sys
sys.path.append(r'd:\Checkhesab\receipt-checker')
from app import parse_excel

with open(r'd:\Checkhesab\1-4 من.xlsx', 'rb') as f:
    txns = parse_excel(f.read(), '1-4 من.xlsx')

found = False
for t in txns:
    if '6076' in t.get('all_codes', []) or t.get('last4') == '6076' or '6076' in t.get('ref', '') or '6076' in t.get('desc', ''):
        print(f"Found 6076 in Excel Row {t.get('row_num')}: Ref={t.get('ref')} Amount={t.get('amount')} Desc={t.get('desc')} Last4={t.get('last4')}")
        found = True

if not found:
    print("6076 NOT FOUND anywhere in Excel!")
