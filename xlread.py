import xlrd
import re
import sys
import xlwt
from datetime import datetime
import string
from xlwt import Workbook
import json
import xlsxwriter

try:
	commandKey = sys.argv[1]
	'''
	file_name = sys.argv[2]	
	style0 = xlwt.easyxf('font: name Times New Roman, color-index red, bold on',num_format_str='#,##0.00')	
	book = xlrd.open_workbook(file_name)
	sh = book.sheet_by_index(0)
    '''
	
	def reading_indices(argv2, argv3):
		#data = [[sh.cell_value(r, col) for col in range(sh.ncols)] for r in range(sh.nrows)]

		rowv =  argv2
		colv = argv3
		#print(D)

		# identifying the cell in excel
		def my_split(s):
			return filter(None, re.split(r'(\d+)', s))
			
		rowindexarray = my_split(rowv)	
		colindexarray = my_split(colv)

		#excel column headers reference index identifying 
		def col_to_num(col_str):
			""" Convert base26 column string to number. """
			expn = 0
			col_num = 0
			for char in reversed(col_str):
				col_num += (ord(char) - ord('A') + 1) * (26 ** expn)
				expn += 1
			return col_num

		#column and row index identifying
		rval = rowindexarray[0]
		rval1 = col_to_num(rval)

		print(rval1)
		#rval1 = I
		#print(type(rval))
		cval = colindexarray[0]
		cval1 = col_to_num(cval)
		print(cval1)

		#getting exact column and row index
		rowindexval = sh.cell_value(int(rowindexarray[1])-1,rval1-1)
		colindexval =sh.cell_value(int(colindexarray[1])-1,cval1-1)

		print(rowindexval)
		print(colindexval)
		return [rowindexval,colindexval]
	#wb = xlwt.Workbook()
	#addition action
	if commandKey == "add":
		argv2 =  sys.argv[2]
		argv3 =  sys.argv[3]
		def addition(n1,n2):
			return float(n1) + float(n2)
		reading_indice = reading_indices(argv2, argv3)	
		additionval = addition(reading_indice[0], reading_indice[1])

		print("Result........")
		print(additionval)
		wb = xlwt.Workbook()
		ws = wb.add_sheet('addition')	
		ws.write(0, 0, reading_indice[0], style0)
		ws.write(1, 0,  reading_indice[1], style0)
		ws.write(2, 2, xlwt.Formula("A3+B3"))
		ws.write(2, 0, additionval, style0)
		wb.save('example.xls')
	#subtraction action	
	elif commandKey == "subtract":
		argv2 =  sys.argv[2]
		argv3 =  sys.argv[3]
		def subtraction(n1,n2):
			return n1 - n2	
		reading_indice = reading_indices(argv2, argv3)
		subval = subtraction(reading_indice[0], reading_indice[1])
		print("Result........")
		print(subval)
		wb = xlwt.Workbook()
		ws = wb.add_sheet('subtraction')	
		ws.write(0, 0, reading_indice[0], style0)
		ws.write(1, 0, reading_indice[1], style0)
		ws.write(2, 0, subval, style0)
		wb.save('example.xls')
		
	elif commandKey == "copy":
		argv2 =  sys.argv[2]	
		temp_list1 = argv2
		temp_list2 = temp_list1.split(',')
		temp_list = list(map(int, temp_list2))
		print(temp_list)
		#temp_list = [1, 4, 5, 6, 7, 8, 9, 10, 15, 19, 26]
		book1 = Workbook()
		sheet1 = book1.add_sheet('test', cell_overwrite_ok=True)
		total_cols = sh.ncols
		k = 0
		for j in temp_list: #for rows
			for i in range(0, total_cols):
				value = sh.cell_value(int(j-1),i)
				sheet1.write(k, i, value)
			k += 1
		book1.save("copy-purchase-data.xls")
		
	elif commandKey == "OpenExcel":	
		try:
			file_name = sys.argv[2]				
			def open_file(file_name):
				book = xlrd.open_workbook(file_name)
				sheet = book.sheet_by_index(0)
				data_excel = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]
				'''
				excel_array = []
				for rx in range(sheet.nrows):
					excel_array.append(sheet.row(rx))
				'''	
				data = {}
				data['data'] = data_excel
				data['status_code'] = "200 OK"
				data['status'] = "File opened successfully"
				data = json.dumps(data)
				return data
				
			open_file_data = open_file(file_name)
			'''
			for rx in range(open_file.nrows):
				print(open_file.row(rx))
			'''	
			print(open_file_data)
		
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = e
			print(data)
			pass
			
	elif commandKey == "CreateExcel":	
		try:
			file_name =  sys.argv[2]
			sheet_name =  sys.argv[3]
			print(file_name)
			def create_file(file_name):
				excel_file = xlwt.Workbook()
				worksheet = excel_file.add_sheet(sheet_name)
				worksheet.write(0, 0, "")
				excel_file.save(file_name)
				data = {}
				data['status_code'] = "200 OK"
				data['status'] = "File created successfully"
				data = json.dumps(data)
				return data
				
			file_data = create_file(file_name)
			print(file_data)
		
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = e
			print(data)
			pass
			
	elif commandKey == "WriteSheet":	
		try:
			file_name =  sys.argv[2]
			sheet_name =  sys.argv[3]
			new_file_name =  sys.argv[4]
			new_sheet_name =  sys.argv[5]
			print(file_name)
			def write_file(file_name):
				book = xlrd.open_workbook(file_name)
				sheet = book.sheet_by_index(0)
				total_rows = sheet.nrows
				print(total_rows)
				temp_list = list(map(int, range(total_rows)))
				book1 = Workbook()
				sheet1 = book1.add_sheet(new_sheet_name, cell_overwrite_ok=True)
				total_cols = sheet.ncols

				k = 0
				for j in temp_list: #for rows
					for i in range(0, total_cols):
						value = sheet.cell_value(int(j),i)
						sheet1.write(k, i, value)
					k += 1
				book1.save(new_file_name)
				data = {}
				data['status_code'] = "200 OK"
				data['status'] = "File written successfully"
				data = json.dumps(data)
				return data
				
			file_data = write_file(file_name)
			print(file_data)
		
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = e
			print(data)
			pass
			
	elif commandKey == "CloseExcel":	
		try:
			file_name = sys.argv[2]				
			def open_file(file_name):
				book = xlrd.open_workbook(file_name)
				workbook = xlsxwriter.Workbook(file_name)
				workbook.close()
				data = {}
				data['status_code'] = "200 OK"
				data['status'] = "File Closed successfully"
				data = json.dumps(data)
				return data
				
			open_file_data = open_file(file_name)
			print(open_file_data)
		
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = e
			print(data)
			pass		
						

	else:
		print("Please enter proper Command Key")

	#saving excel file	

except Exception as e:
	data = {}
	data['status_code'] = "401"
	data['status'] = e 
	#data = json.dumps(data)
	print(data)
	pass


	
	

