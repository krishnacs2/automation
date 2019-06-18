import xlrd
from xlwt import Workbook

book = xlrd.open_workbook("Purchase_data.xlsx")
sheet = book.sheet_by_index(0)

temp_list = [1, 4, 5, 6, 7, 8, 9, 10, 15, 19, 26]
book1 = Workbook()
sheet1 = book1.add_sheet('test', cell_overwrite_ok=True)
total_cols = sheet.ncols
k = 0
for j in temp_list: #for rows
    for i in range(0, total_cols):
        value = sheet.cell_value(j-1,i)
        sheet1.write(k, i, value)
    k += 1
book1.save("copy-purchase-data.xls")