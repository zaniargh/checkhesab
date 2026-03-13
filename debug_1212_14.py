import sys
from pathlib import Path

sys.path.append(r"d:\Checkhesab\receipt-checker")
from app import parse_html, parse_excel, match_receipts

html_path = r"d:\Checkhesab\1212.Html"
excel_path = r"d:\Checkhesab\1-4 من.xlsx"

print("Parsing HTML...")
with open(html_path, "rb") as f:
    pdf_rows = parse_html(f.read())
    
print("Parsing Excel...")
with open(excel_path, "rb") as f:
    bank_txns = parse_excel(f.read(), "1-4 من.xlsx")
    
# Find the specific HTML row
target_html = None
for r in pdf_rows:
    if r.get("amount") == 1800000000 and "6076" in r.get("codes", []):
        target_html = r
        break

if target_html:
    print("\n--- Found in HTML ---")
    for k, v in target_html.items():
        print(f"{k}: {v}")
else:
    print("\n--- NOT Found in HTML! ---")
    for r in pdf_rows:
        if r.get("amount") == 1800000000:
             print(f"HTML Amount Match: {r}")
             
print("\n--- Potential in Excel ---")
for t in bank_txns:
    if t.get("amount") == 1800000000:
        print(f"Row {t.get('row_num')}: Ref={t.get('ref')} Codes={t.get('all_codes')} Desc={t.get('desc')} Last4={t.get('last4')}")
        
    if "6076" in t.get("all_codes", []) or t.get("last4") == "6076":
         print(f"By Code 6076 - Row {t.get('row_num')}: Ref={t.get('ref')} Amount={t.get('amount')} Desc={t.get('desc')}")

results = match_receipts(
    pdf_rows, bank_txns,
    credit_only=True,
    use_tracking=True,
    use_name=True,
    use_amount=True,
    tx_type_filter="all",
    use_date=False
)

print("\n--- Match Result ---")
for res in results:
    if res["amount"] == 1800000000 and "6076" in res["codes"]:
        for k, v in res.items():
            print(f"{k}: {v}")
