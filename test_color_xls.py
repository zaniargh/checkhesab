import xlrd
from xlutils.copy import copy
import xlwt

try:
    print('Opening workbook...')
    rb = xlrd.open_workbook('d:/Checkhesab/ایران 1-6.xls', formatting_info=True)
    wb = copy(rb)
    sheet = wb.get_sheet(0)

    # Make yellow background style
    style = xlwt.easyxf('pattern: pattern solid, fore_colour yellow;')
    style.font.name = 'Tahoma'

    print('Coloring row 10...')
    for col_idx in range(10): # Color first 10 columns of row 10
        sheet.write(10, col_idx, "تست زرد", style)
        
    wb.save('d:/Checkhesab/output_test.xls')
    print('Saved to output_test.xls')
except Exception as e:
    print('Error:', e)
