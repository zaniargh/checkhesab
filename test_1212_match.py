import sys
sys.path.append('d:/Checkhesab/receipt-checker')
from app import parse_html, parse_excel, match_receipts

with open('d:/Checkhesab/1212.Html', 'rb') as f:
    pdf_rows = parse_html(f.read())
with open('d:/Checkhesab/ایران 1-6.xls', 'rb') as f:
    bank_txns = parse_excel(f.read(), 'ایران 1-6.xls')

matched = match_receipts(pdf_rows, bank_txns, credit_only=False)
found_any = False
for r in matched:
    if r['found']:
        print(f"Match Found: amount={r['amount']}, method={r.get('match_method', '')}, sender={r['sender']}, codes={r.get('codes', '')}")
        found_any = True

if not found_any:
    print("NO matches found at all.")
    
    # Let's see all pdf rows to identify tracking codes that SHOULD match
    print("\n--- HTML Rows with Codes or Sender ---")
    for r in pdf_rows:
        if r.get('codes') or r.get('sender'):
            print(f"amount={r['amount']}, codes={r.get('codes', [])}, sender='{r.get('sender', '')}', desc='{r.get('desc', '')}'")

