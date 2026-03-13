import sys
sys.path.append(r'd:\Checkhesab\receipt-checker')
from app import parse_html, parse_excel, match_receipts
import re

with open(r'd:\Checkhesab\1212.Html', 'rb') as f:
    pdf_rows = parse_html(f.read())
    
with open(r'd:\Checkhesab\1-4 من.xlsx', 'rb') as f:
    bank_txns = parse_excel(f.read(), '1-4 من.xlsx')
    
target_bank_tx = next((t for t in bank_txns if t.get('amount') == 450000000), None)
if target_bank_tx:
    print(f"Target Bank Tx Amount: {target_bank_tx['amount']}")
    print(f"Target Bank Tx Tx_type: {target_bank_tx['tx_type']}")
    print(f"Target Bank Tx All_codes: {target_bank_tx.get('all_codes')}")
    print(f"Target Bank Tx Last4: {target_bank_tx.get('last4')}")
else:
    print("NO BANK TX FOUND FOR 450M")

print("\nRunning custom indexing test...")

by_last4 = {}
tx_type_filter = 'all'

for tx in bank_txns:
    if tx_type_filter != "all" and tx.get("tx_type") != "unknown" and tx.get("tx_type") != tx_type_filter:
        continue
        
    all_codes = set(tx.get("all_codes", []))
    if tx.get("last4"):
        all_codes.add(tx["last4"])
        
    for code in all_codes:
        if code:
            by_last4.setdefault(code, []).append(tx)
            if len(code) > 4:
                by_last4.setdefault(code[-4:], []).append(tx)
                by_last4.setdefault(code[-5:], []).append(tx)

print(f"Is 2951 in by_last4? {bool(by_last4.get('2951'))}")
if by_last4.get('2951'):
    print(f"Found txns for 2951: {[t.get('amount') for t in by_last4['2951']]}")

print(f"Is 12951 in by_last4? {bool(by_last4.get('12951'))}")
if by_last4.get('12951'):
    print(f"Found txns for 12951: {[t.get('amount') for t in by_last4['12951']]}")

