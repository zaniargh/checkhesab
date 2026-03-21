import app
import codecs
from pprint import pprint

with open('d:/Checkhesab/8888.html', 'rb') as f:
    res = app.parse_html(f.read())

with codecs.open('debug_output.txt', 'w', encoding='utf-8') as out:
    for r in res:
        out.write(f"Desc: {r.get('desc')}\n")
        out.write(f"Customer: {r.get('customer_name')}\n")
        out.write(f"Sender: {r.get('sender')}\n")
        out.write(f"Amount: {r.get('amount')}\n")
        out.write(f"Codes: {r.get('codes')}\n")
        out.write('---\n')
