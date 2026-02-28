import sys
sys.path.append('d:/Checkhesab/receipt-checker')
from app import parse_excel
import glob

for f in glob.glob('d:/Checkhesab/ایران*.xlsx'):
    try:
        with open(f, 'rb') as rd:
            data = rd.read()
            txns = parse_excel(data, f)
            locked = [t for t in txns if t.get('is_locked')]
            print(f'File {f} -> {len(locked)} locked.')
    except Exception as e:
        print(e)
