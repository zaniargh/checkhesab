import sys
sys.path.append(r'd:\Checkhesab\receipt-checker')
from app import parse_excel

with open(r'd:\Checkhesab\1-4 من.xlsx', 'rb') as f:
    bank_txns = parse_excel(f.read(), '1-4 من.xlsx')

with open(r'd:\Checkhesab\receipt-checker\out_450.txt', 'w', encoding='utf-8') as f:
    for t in bank_txns:
        if t.get('amount') == 450000000:
            f.write("-" * 40 + "\n")
            f.write(f"Row {t['row_num']}:\n")
            f.write(f"  amount: {t['amount']}\n")
            f.write(f"  desc: {t.get('desc')}\n")
            f.write(f"  sender: {t.get('sender')}\n")
            f.write(f"  ref: {t.get('ref')}\n")
            f.write(f"  doc_num: {t.get('doc_num')}\n")
            f.write(f"  all_codes: {t.get('all_codes')}\n")
            f.write(f"  last4: {t.get('last4')}\n")
