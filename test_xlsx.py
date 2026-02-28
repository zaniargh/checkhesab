import sys
sys.path.append('d:/Checkhesab/receipt-checker')
from app import parse_excel

try:
    with open('d:/Checkhesab/ایران 1-6.xlsx', 'rb') as f:
        data = f.read()
    txns = parse_excel(data, 'ایران 1-6.xlsx')
    
    locked_count = 0
    for t in txns:
        if t.get('is_locked'):
            locked_count += 1
            print(f"Locked Row {t['row_num']}: {t['raw'][:60]}...")
            
    print(f"\nTotal txns: {len(txns)}")
    print(f"Total locked: {locked_count}")
except Exception as e:
    print("Error:", e)
