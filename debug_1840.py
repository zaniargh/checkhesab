import sys
sys.path.append(r'd:\Checkhesab\receipt-checker')
from app import parse_html, parse_excel

with open(r'd:\Checkhesab\1212.Html', 'rb') as f:
    pdf_rows = parse_html(f.read())

with open(r'd:\Checkhesab\1-4 من.xlsx', 'rb') as f:
    bank_txns = parse_excel(f.read(), '1-4 من.xlsx')

TARGET_AMOUNT = 1_840_000_000

with open(r'd:\Checkhesab\receipt-checker\out_1840.txt', 'w', encoding='utf-8') as f:
    f.write("=== HTML rows with 1840000000 ===\n")
    for r in pdf_rows:
        if r.get('credit') == TARGET_AMOUNT or r.get('debit') == TARGET_AMOUNT:
            f.write(f"  codes: {r.get('codes')}\n")
            f.write(f"  amount: {r.get('credit') or r.get('debit')}\n")
            f.write(f"  desc: {r.get('desc','')[:100]}\n\n")

    f.write("\n=== BANK rows with 1840000000 ===\n")
    for t in bank_txns:
        if t.get('amount') == TARGET_AMOUNT:
            f.write(f"  row: {t['row_num']}\n")
            f.write(f"  all_codes: {t.get('all_codes')}\n")
            f.write(f"  last4: {t.get('last4')}\n")
            f.write(f"  amount: {t['amount']}\n")
            f.write(f"  tx_type: {t.get('tx_type')}\n")
            f.write(f"  desc: {t.get('desc','')[:100]}\n\n")

    # Now check if any of the html codes appear in bank codes
    html_codes_1840 = []
    for r in pdf_rows:
        if r.get('credit') == TARGET_AMOUNT or r.get('debit') == TARGET_AMOUNT:
            html_codes_1840.extend(r.get('codes', []))

    f.write(f"\n=== HTML codes for 1840: {html_codes_1840} ===\n")
    f.write("Checking each code in bank index...\n")

    by_last4 = {}
    for t in bank_txns:
        all_codes = set(t.get('all_codes', []))
        if t.get('last4'):
            all_codes.add(t['last4'])
        seen = set()
        for code in all_codes:
            if code:
                for key in [code] + ([code[-5:]] if len(code) > 5 else []) + ([code[-4:]] if len(code) > 4 else []):
                    if key and (key, t.get('row_num')) not in seen:
                        by_last4.setdefault(key, []).append(t)
                        seen.add((key, t.get('row_num')))

    for code in html_codes_1840:
        f.write(f"\nCode '{code}': found in bank? {bool(by_last4.get(code))}\n")
        if by_last4.get(code):
            for tx in by_last4[code]:
                f.write(f"  -> row {tx['row_num']}, amount={tx['amount']}, diff={abs(tx['amount']-TARGET_AMOUNT)}\n")
        if len(code) > 4:
            f.write(f"  last4 '{code[-4:]}': {bool(by_last4.get(code[-4:]))}\n")
            if by_last4.get(code[-4:]):
                for tx in by_last4[code[-4:]]:
                    f.write(f"    -> row {tx['row_num']}, amount={tx['amount']}\n")
        if len(code) > 5:
            f.write(f"  last5 '{code[-5:]}': {bool(by_last4.get(code[-5:]))}\n")
            if by_last4.get(code[-5:]):
                for tx in by_last4[code[-5:]]:
                    f.write(f"    -> row {tx['row_num']}, amount={tx['amount']}\n")
