import sys
sys.path.append('d:/Checkhesab/receipt-checker')
from app import parse_html, parse_excel, match_receipts

with open('d:/Checkhesab/1212.Html', 'rb') as f:
    pdf_rows = parse_html(f.read())
print('Extracted', len(pdf_rows), 'rows from 1212.html')
for i, r in enumerate(pdf_rows):
    if r['amount'] == 9000000000 or i < 5:
        print(f"Row {i}: amount={r['amount']}, codes={r.get('codes', [])}, sender='{r.get('sender', '')}', desc='{r.get('desc', '')}'")

with open('d:/Checkhesab/ایران 1-6.xls', 'rb') as f:
    bank_txns = parse_excel(f.read(), 'ایران 1-6.xls')

matched = match_receipts(pdf_rows, bank_txns, credit_only=False)
for r in matched:
    if r['amount'] == 9000000000 or r['found']:
        print(f"Result {r['amount']}: Found={r['found']}, Method={r.get('match_method', '')}, BankRef={r.get('bank_ref', '')}")

