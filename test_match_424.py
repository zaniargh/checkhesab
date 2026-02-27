import sys
sys.path.append('d:/Checkhesab/receipt-checker')
from app import parse_excel

with open('d:/Checkhesab/ایران 1-6.xls', 'rb') as f:
    bank_txns = parse_excel(f.read(), 'ایران 1-6.xls')

print('Total bank txns:', len(bank_txns))
found = False
for r in bank_txns:
    if r['amount'] == 424680424:
        print('Excel Match:', r)
        found = True

if not found:
    print('Amount 424680424 not found in parsed Excel.')
