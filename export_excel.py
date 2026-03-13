import pandas as pd
import json
import sys

# Read the Bank Saderat excel template
df = pd.read_excel(r'd:\Checkhesab\1-4 من.xlsx', nrows=10, header=None)

# Convert head to json
data = df.to_dict(orient='records')
with open('excel_head.json', 'w', encoding='utf-8') as f:
    json.dump(data, f, ensure_ascii=False, indent=2)

print("Exported to excel_head.json")
