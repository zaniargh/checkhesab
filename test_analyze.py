import sys
sys.path.append('d:/Checkhesab/receipt-checker')
from app import parse_html, parse_excel, match_receipts
import logging

logging.basicConfig(level=logging.INFO)

with open('d:/Checkhesab/333.Html', 'rb') as f:
    pdf_rows = parse_html(f.read())

with open('d:/Checkhesab/ایران 1-6.xls', 'rb') as f:
    bank_txns = parse_excel(f.read(), 'ایران 1-6.xls')

# Exactly as in the route:
results = match_receipts(
    pdf_rows, bank_txns,
    credit_only=False,
    use_tracking=True,
    use_name=True,
    use_amount=True,
)

print('Total results returned by match_receipts:', len(results))
if len(results) > 0:
    print('First result:', results[0])
    print('Last result:', results[-1])
