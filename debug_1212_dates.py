import sys
sys.path.append(r'd:\Checkhesab\receipt-checker')
from app import parse_excel

with open(r'd:\Checkhesab\1-4 من.xlsx', 'rb') as f:
    txns = parse_excel(f.read(), '1-4 من.xlsx')

for t in txns:
    if t.get('amount') == 1800000000:
        print(f"Row {t.get('row_num')}: Date={t.get('date')} Ref={t.get('ref')} Desc={t.get('desc')}")
