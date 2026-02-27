import re

cases = [
    "| 0260/2502 |",
    "| 0260/2502",
    "محمد حسنی / 0260/2502",
    "0260-2502 / محمد",
    "| 0260 / 2502 |",
]

for desc_n in cases:
    codes = set()
    
    # Pattern 1a
    for m in re.finditer(r"[\u0600-\u06FF]+\s*/\s*([\d\s/-]{4,})", desc_n):
        for code in re.findall(r"\d{4,8}", m.group(1)):
            codes.add(code[-4:])
            
    # Pattern 1b
    for m in re.finditer(r"([\d\s/-]{4,})\s*/\s*[\u0600-\u06FF]+", desc_n):
        for code in re.findall(r"\d{4,8}", m.group(1)):
            codes.add(code[-4:])
            
    # Pattern 5
    for m in re.finditer(r"\|\s*([\d\s/-]{4,})\s*(?=\|)", desc_n):
        for code in re.findall(r"\d{4,8}", m.group(1)):
            codes.add(code[-4:])

    # Pattern 6
    for m in re.finditer(r"\|\s*([\d\s/-]{4,})\s*$", desc_n):
        for code in re.findall(r"\d{4,8}", m.group(1)):
            codes.add(code[-4:])

    print(f"Desc: {desc_n} -> Codes: {codes}")
