import sys
sys.path.append('d:/Checkhesab/receipt-checker')
from app import parse_html

try:
    with open('d:/Checkhesab/333.Html', 'rb') as f:
        html_bytes = f.read()

    rows = parse_html(html_bytes)
    print("Total HTML rows extracted from 333.Html:", len(rows))
    print("First row:", rows[0] if rows else "None")
    print("Last row:", rows[-1] if rows else "None")
except Exception as e:
    print("Error:", e)
