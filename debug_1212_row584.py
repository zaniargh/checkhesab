import sys
import pandas as pd

file_path = r'd:\Checkhesab\1-4 من.xlsx'
df = pd.read_excel(file_path, header=None, dtype=str)

with open(r'd:\Checkhesab\receipt-checker\out_row584.txt', 'w', encoding='utf-8') as f:
    for idx, row in df.iterrows():
        row_list = row.tolist()
        row_str = " ".join([str(x) for x in row_list])
        if "450000000" in row_str and "GPPC" in row_str:
            f.write(f"--- MATCH FOUND AT INDEX {idx} ---\n")
            for c_idx, val in enumerate(row_list):
                f.write(f"Col {c_idx}: {val}\n")
