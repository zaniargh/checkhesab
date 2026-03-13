import sys
sys.path.append(r'd:\Checkhesab\receipt-checker')
from app import parse_html, parse_excel, match_receipts

pdf_rows = [
    {
        "page": 1, "date": "1404/11/02", "doc_num": "", "desc": "Receipt 1",
        "doc_type": "بستانکار", "credit": 50000, "debit": 0, "amount": 50000,
        "codes": ["1234"], "sender": "Ali"
    },
    {
        "page": 1, "date": "1404/11/02", "doc_num": "", "desc": "Receipt 2 duplicate",
        "doc_type": "بستانکار", "credit": 50000, "debit": 0, "amount": 50000,
        "codes": ["1234"], "sender": "Ali"
    },
    {
        "page": 1, "date": "1404/11/02", "doc_num": "", "desc": "Receipt multiple",
        "doc_type": "بستانکار", "credit": 90000, "debit": 0, "amount": 90000,
        "codes": ["5678"], "sender": "Reza"
    }
]

bank_txns = [
    {
        "row_num": 10, "ref": "1234", "last4": "1234", "all_codes": ["1234"],
        "amount": 50000, "tx_type": "deposit", "date": "1404/11/02", "desc": "Bank 1", "sender": "Ali",
        "is_locked": False, "lock_text": ""
    },
    # Two identical bank rows for 5678 to trigger multiple
    {
        "row_num": 20, "ref": "5678", "last4": "5678", "all_codes": ["5678"],
        "amount": 90000, "tx_type": "deposit", "date": "1404/11/02", "desc": "Bank 2", "sender": "Reza",
        "is_locked": False, "lock_text": ""
    },
    {
        "row_num": 21, "ref": "56789", "last4": "5678", "all_codes": ["5678"],
        "amount": 90000, "tx_type": "deposit", "date": "1404/11/02", "desc": "Bank 3", "sender": "Reza",
        "is_locked": False, "lock_text": ""
    }
]

# Run first time
results1 = match_receipts(pdf_rows[:1], bank_txns, credit_only=True, use_tracking=True, use_name=True, use_amount=True, tx_type_filter="all", use_date=False)
print("--- FIRST RUN ---")
for r in results1:
    print(f"Row {r['idx']} - {r['desc']} => Status: {r['status']}, Method: {r['match_method']}")
    
# Lock the first result
bank_txns[0]["is_locked"] = True
bank_txns[0]["lock_text"] = "تطبیق شده - File1"

# Run second time for duplicate check
results2 = match_receipts(pdf_rows[1:2], bank_txns, credit_only=True, use_tracking=True, use_name=True, use_amount=True, tx_type_filter="all", use_date=False)
print("\n--- SECOND RUN (DUPLICATE) ---")
for r in results2:
    print(f"Row {r['idx']} - {r['desc']} => Status: {r['status']}, Method: {r['match_method']}, PrevLock: {r.get('duplicate_lock_text')}")

# Run third time for multiple
results3 = match_receipts(pdf_rows[2:], bank_txns, credit_only=True, use_tracking=True, use_name=True, use_amount=True, tx_type_filter="all", use_date=False)
print("\n--- THIRD RUN (MULTIPLE) ---")
for r in results3:
    print(f"Row {r['idx']} - {r['desc']} => Status: {r['status']}, Method: {r['match_method']}, BankRows: {r.get('bank_rows')}")
