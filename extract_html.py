import json
import sys
sys.path.append('.')
from app import parse_html

try:
    with open(r'd:\Checkhesab\nbnbnbnbnbnbnbnbn.Html', 'rb') as f:
        html_bytes = f.read()
    
    results = parse_html(html_bytes)
    
    with open('html_parsed.json', 'w', encoding='utf-8') as f:
        json.dump(results, f, ensure_ascii=False, indent=2)
        
    print(f"Successfully exported {len(results)} rows to html_parsed.json")
except Exception as e:
    print(f"Error: {e}")
