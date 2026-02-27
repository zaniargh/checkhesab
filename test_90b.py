import sys
sys.path.append('d:/Checkhesab/receipt-checker')
from app import parse_html, parse_excel, match_receipts

with open('d:/Checkhesab/1212.Html', 'rb') as f:
    pdf_rows = parse_html(f.read())
for r in pdf_rows:
    if r['amount'] == 90000000000:
        print(f"Row amount=90000000000, codes={r.get('codes', [])}, sender='{r.get('sender', '')}', desc='{r.get('desc', '')}'")

with open('d:/Checkhesab/ایران 1-6.xls', 'rb') as f:
    bank_txns = parse_excel(f.read(), 'ایران 1-6.xls')
for r in bank_txns:
    if r['amount'] == 90000000000:
        print(f"Bank Txn: amount=90000000000, ref={r.get('ref', '')}, sender='{r.get('sender', '')}', raw='{r.get('raw', '')}'")
