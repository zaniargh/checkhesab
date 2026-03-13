import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
sys.path.append('.')
from app import parse_html, parse_excel, match_receipts

with open(r'd:\Checkhesab\nbnbnbnbnbnbnbnbn.Html', 'rb') as f:
    html_rows = parse_html(f.read())
with open(r'd:\Checkhesab\1-4 من.xlsx', 'rb') as f:
    excel_rows = parse_excel(f.read(), '1-4 من.xlsx')

results = match_receipts(html_rows, excel_rows)

not_found = sum(1 for r in results if not r['found'])
exact = sum(1 for r in results if r['status'] == 'exact')
review = sum(1 for r in results if r['status'] == 'review')

print(f"Total HTML Rows processed: {len(results)}")
print(f"Exact Matches: {exact}")
print(f"Review Matches: {review}")
print(f"Not Found: {not_found}")

print("\n=== ALL MATCH RESULTS ===")
for r in results:
    print(f"  [{r['status'].upper()}] Amount: {int(r['amount']):,} | Method: {r['match_method']} | Bank Row: {r['bank_row']}")
