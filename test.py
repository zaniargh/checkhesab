import app

pdf = open(r'd:\Checkhesab\گزارش حساب زانيار حسن زاده(1817) 09120046814 (نمايش همه اسناد).pdf','rb').read()
xls = open(r'd:\Checkhesab\ایران 1-6.xls','rb').read()
rows = app.parse_pdf(pdf)
bank = app.parse_excel(xls, 'x.xls')
results = app.match_receipts(rows, bank)
matched = [r for r in results if r.get('found')]
pending = [r for r in results if not r.get('found')]
print(f'Total: {len(results)} | Matched: {len(matched)} | Pending: {len(pending)}')
print()
for r in matched[:20]:
    print(f"[{r['match_method']}] Amount:{r['amount']:,} | Ref:{r['bank_ref']} | Desc:{r['desc'][:50]}")
