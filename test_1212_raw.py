import sys
sys.path.append('d:/Checkhesab/receipt-checker')
from bs4 import BeautifulSoup

def inspect_html(filepath):
    with open(filepath, 'rb') as f:
        html = f.read()
    soup = BeautifulSoup(html, "html.parser")
    tables = soup.find_all("table")
    for tbody in [t.find("tbody") or t for t in tables]:
        for tr in tbody.find_all("tr"):
            tds = tr.find_all(["td", "th"])
            cells = [t.get_text(strip=True) for t in tds]
            if len(cells) < 4: continue
            
            raw_text = " | ".join(cells)
            if "2,000,000,000" in raw_text or "35,000,000,000" in raw_text or "31,000,000,000" in raw_text:
                print("\nMATCH FOUND IN HTML ROW:")
                print(raw_text)

inspect_html('d:/Checkhesab/1212.Html')
