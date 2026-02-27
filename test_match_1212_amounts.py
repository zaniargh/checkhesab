import sys
sys.path.append('d:/Checkhesab/receipt-checker')
from app import parse_html, parse_excel

with open('d:/Checkhesab/1212.Html', 'rb') as f:
    pdf_rows = parse_html(f.read())
html_amounts = {r['amount']: r for r in pdf_rows if r['amount'] > 0}

with open('d:/Checkhesab/ایران 1-6.xls', 'rb') as f:
    bank_txns = parse_excel(f.read(), 'ایران 1-6.xls')

print("--- Excel rows with matching amounts from 1212.html ---")
found_excel_amounts = set()
for bx in bank_txns:
    if bx['amount'] in html_amounts:
        amount = bx['amount']
        found_excel_amounts.add(amount)
        html_row = html_amounts[amount]
        print(f"\nAMOUNT: {amount}")
        print(f"  HTML: codes={html_row.get('codes', [])}, sender='{html_row.get('sender', '')}', desc='{html_row.get('desc', '')[:80]}'")
        print(f"  EXCEL: ref={bx.get('ref', '')}, last4={bx.get('last4', '')}, all_codes={bx.get('all_codes', [])}")
        print(f"  EXCEL RAW: {bx.get('raw', '')}")

print("\n--- HTML amounts NOT found in Excel ---")
for amt in sorted(html_amounts.keys()):
    if amt not in found_excel_amounts:
        print(amt)
