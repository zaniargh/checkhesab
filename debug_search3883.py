import sys
import pandas as pd

with open(r'd:\Checkhesab\receipt-checker\out_search3883.txt', 'w', encoding='utf-8') as out:
    df = pd.read_excel(r'd:\Checkhesab\1-4 من.xlsx', header=None, dtype=str)
    out.write("Searching for 3883, 8330, 978330 in any cell...\n\n")
    for code in ['3883', '8330', '978330', '1840000000']:
        out.write(f"\n=== Searching for '{code}' ===\n")
        for idx, row in df.iterrows():
            for col_idx, val in enumerate(row):
                if code in str(val):
                    out.write(f"  Row {idx}, Col {col_idx}: {val}\n")
