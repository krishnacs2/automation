import xlrd
import re
import sys
import xlwt
import string
from xlwt import Workbook
import json
import io
import os.path
from xlutils.copy import copy as xl_copy
import Tkinter
import tkFileDialog
import shutil
import win32com.client
import operator


try:
	commandKey = sys.argv[1]
	'''
	file_name = sys.argv[2]	
	style0 = xlwt.easyxf('font: name Times New Roman, color-index red, bold on',num_format_str='#,##0.00')	
	book = xlrd.open_workbook(file_name)
	sh = book.sheet_by_index(0)
    '''
	
	def col_to_num(col_str):
		""" Convert base26 column string to number. """
		expn = 0
		col_num = 0
		for char in reversed(col_str):
			col_num += (ord(char) - ord('A') + 1) * (26 ** expn)
			expn += 1
		return col_num
	
	def reading_indices(file_name, sheet_name,temp_list1):
		#data = [[sh.cell_value(r, col) for col in range(sh.ncols)] for r in range(sh.nrows)]
		
		book = xlrd.open_workbook(file_name)
		sh = book.sheet_by_name(sheet_name)
		rowv =  temp_list1
		#colv = argv3

		# identifying the cell in excel
		def my_split(s):
			return filter(None, re.split(r'(\d+)', s))
			
		rowindexarray = my_split(rowv)	
		#colindexarray = my_split(colv)

		#excel column headers reference index identifying 

		#column and row index identifying
		rval = rowindexarray[0]
		rval1 = col_to_num(rval)

		#print(rval1)
		#rval1 = I
		#print(type(rval))
		#cval = colindexarray[0]
		#cval1 = col_to_num(cval)
		#print(cval1)

		#getting exact column and row index
		rowindexval = sh.cell_value(int(rowindexarray[1])-1,rval1-1)
		#colindexval =sh.cell_value(int(colindexarray[1])-1,cval1-1)

		#print(rowindexval)
		#print(colindexval)
		data = {}
		data['data'] = rowindexval
		data['status_code'] = "200 OK"
		data['status'] = "Cell value read successfully"
		data = json.dumps(data)
		return data	
	


	def preventexcelrename(file_string):
		sheet_name_out1s = list(file_string)
		sheetpreventlist = ['\\','/','*','[',']',':','?']
		sheetpreventresult = []
		for sheet_name_out1 in sheet_name_out1s:
			for sheetpreventlist1 in sheetpreventlist:
				if int(sheet_name_out1.find(sheetpreventlist1)) > -1:
					sheetpreventresult.append(str(sheet_name_out1.find(sheetpreventlist1)))	
		return sheetpreventresult
			
			
	def findexcel_column(excel_list2,wb_name,wk_sheet):
		book = xlrd.open_workbook(wb_name)
		sh = book.sheet_by_name(wk_sheet)
		for row in range(1):
			for column in range(0,sh.ncols):
				if excel_list2 == sh.cell(row, column).value:  
					return column		

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
			data['status'] = str(e)
			print(json.dumps(data))
			pass
			
			
	elif commandKey == "OpenSheet":	
		try:
			file_name = sys.argv[2]	
			sheet_name = sys.argv[3]				
			def open_file(file_name):
				book = xlrd.open_workbook(file_name)
				sheet = book.sheet_by_name(sheet_name)
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
			data['status'] = str(e)
			print(json.dumps(data))
			pass
			
			
	elif commandKey == "CreateExcel":	
		try:
			if os.path.exists(sys.argv[2]): 
				data = {}
				data['status_code'] = "401"
				data['status'] = "File name already exists"
				print(json.dumps(data))
				pass
					
			else: 
				filename, file_extension = os.path.splitext(sys.argv[2])
				if filename[-1] not in (' '):
					if file_extension in ('.xls', '.xlsx'):
						file_name =  sys.argv[2]
						sheet_name =  sys.argv[3]
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
						
					elif file_extension in (''):
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass
					else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file extension"
						print(json.dumps(data))
						pass
						
				else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass	
		
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e)
			print(json.dumps(data))
			pass
			
			
	elif commandKey == "CreateSheet":	
		try:
			if os.path.exists(sys.argv[2]): 
				filename, file_extension = os.path.splitext(sys.argv[2])			
				if filename[-1] not in (' '):
					if sys.argv[2].endswith('xls') or sys.argv[2].endswith('xlsx'):
						file_name =  sys.argv[2]
						sheet_name =  sys.argv[3]
						def add_sheet(file_name):
							readbook = xlrd.open_workbook(file_name, formatting_info=True)
							wb = xl_copy(readbook)
							Sheet1 = wb.add_sheet(sheet_name)
							wb.save(file_name)
							data = {}
							data['status_code'] = "200 OK"
							data['status'] = "Sheet added successfully"
							data = json.dumps(data)
							return data
							
						file_data = add_sheet(file_name)
						print(file_data)
						
					elif file_extension in (''):
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass
					else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file extension"
						print(json.dumps(data))
						pass
						
				else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass	
						
			else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide Existing file name"
						print(json.dumps(data))
						pass			
		
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e)
			print(json.dumps(data))
			pass
			
			
			
	elif commandKey == "CopySheet":	
		try:
			if os.path.exists(sys.argv[4]): 
				data = {}
				data['status_code'] = "401"
				data['status'] = "File name already exists"
				print(json.dumps(data))
				pass
					
			else: 
				filename_in, file_extension_in = os.path.splitext(sys.argv[2])
				filename_in1, file_extension_in1 =  sys.argv[2].split(".")
				filename_out, file_extension_out = sys.argv[4].split(".")
				filename_in, file_extension_out1 = os.path.splitext(sys.argv[4])
				if filename_out not in ("") and filename_in1 not in (""):
					if filename_out[-1] not in ('','/','\\') and filename_in1[-1] not in ('','/','\\'):
						if filename_out[-1] not in (' ') and filename_in1[-1] not in (' '):		
							if sys.argv[2].endswith(file_extension_in1) and sys.argv[4].endswith(file_extension_out):	
								if file_extension_in in ('.xls', '.xlsx') and file_extension_out1 in ('.xls', '.xlsx'):
									if file_extension_in1 == file_extension_out :
										file_name =  sys.argv[2]
										sheet_name =  sys.argv[3]
										new_file_name =  sys.argv[4]
										new_sheet_name =  sys.argv[5]
										def write_file(file_name):
											book = xlrd.open_workbook(file_name)
											sheet = book.sheet_by_name(sheet_name)
											total_rows = sheet.nrows
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
										
									elif filename_out[-1] in ('','/'):
										data = {}
										data['status_code'] = "401"
										data['status'] = "Please provide proper file name"
										print(json.dumps(data))
										pass	
										
									else:
										data = {}
										data['status_code'] = "401"
										data['status'] = "Please provide proper file extension"
										print(json.dumps(data))
										pass
										
								else:
									data = {}
									data['status_code'] = "401"
									data['status'] = "Please provide proper file extension"
									print(json.dumps(data))
									pass
											
														
							else:
								data = {}
								data['status_code'] = "401"
								data['status'] = "Please provide proper file extension"
								print(json.dumps(data))
								pass	
								
											
						else:
								data = {}
								data['status_code'] = "401"
								data['status'] = "Please provide proper file name"
								print(json.dumps(data))
								pass	
								
					else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass		
					
				else:
					data = {}
					data['status_code'] = "401"
					data['status'] = "Please provide proper file name"
					print(json.dumps(data))
					pass
				
							
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e)
			print(json.dumps(data))
			pass
			
	elif commandKey == "CloseExcel":	
		try:
			file_name = sys.argv[2]				
			def open_file(file_name):
				'''
				book = xlrd.open_workbook(file_name)
				workbook = xlsxwriter.Workbook(file_name)
				workbook.close()
				'''
				xlsx_filename=file_name
				with open(xlsx_filename, "rb") as f:
					in_mem_file = io.BytesIO(f.read())
				#wb = openpyxl.load_workbook(in_mem_file, read_only=True)
				in_mem_file.close()
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
			data['status'] = str(e)
			print(json.dumps(data))
			pass		
			
	elif commandKey == "RowCount":	
		try:
			file_name =  sys.argv[2]
			sheet_name =  sys.argv[3]
			def write_file(file_name):
				book = xlrd.open_workbook(file_name)
				sheet = book.sheet_by_name(sheet_name)
				total_rows = sheet.nrows
				#total_cols = sheet.ncols

				data = {}
				data['status_code'] = "200 OK"
				data['data'] = total_rows
				data = json.dumps(data)
				return data
				
			total_rows = write_file(file_name)
			print(total_rows)
		
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e)
			print(json.dumps(data))
			pass		
			
			
	elif commandKey == "ColumnCount":	
		try:
			file_name =  sys.argv[2]
			sheet_name =  sys.argv[3]
			def write_file(file_name):
				book = xlrd.open_workbook(file_name)
				sheet = book.sheet_by_name(sheet_name)
				#total_rows = sheet.nrows
				total_cols = sheet.ncols
				data = {}
				data['status_code'] = "200 OK"
				data['data'] = total_cols
				data = json.dumps(data)
				return data
				
			total_cols = write_file(file_name)
			print(total_cols)
		
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e)
			print(json.dumps(data))
			pass			
						
						
	elif commandKey == "RenameSheet":	
		try:
			if os.path.exists(sys.argv[2]): 
				filename, file_extension = os.path.splitext(sys.argv[2])			
				if filename[-1] not in (' '):
					if sys.argv[2].endswith('xls') or sys.argv[2].endswith('xlsx'):
						file_name =  sys.argv[2]
						sheet_name_in =  sys.argv[3]
						sheet_name_out =  sys.argv[4]
						sheet_name_out = sheet_name_out.strip()
						sheet_name_out1 = sheet_name_out.lower()
						if len(sheet_name_out) < 32:
							validatesheetname = preventexcelrename(sheet_name_out)
							if len(validatesheetname) < 1:
								if sheet_name_out not in (' '):
									def remane_sheet(file_name):
										readbook = xlrd.open_workbook(file_name)
										sheetlists = readbook.sheet_names()
										if sheet_name_in in sheetlists:
											if sheet_name_out1 in sheetlists:
												data = {}
												data['status_code'] = "401"
												data['status'] = "Duplicate sheet name not allowed"
												return data
												pass
											else:
												wb = xl_copy(readbook)
												# find the index of a sheet you wanna rename,
												# let's say you wanna rename Sheet1
												idx = readbook.sheet_names().index(sheet_name_in)							
												# now rename the sheet in the writable copy
												wb.get_sheet(idx).name = sheet_name_out
												wb.save(file_name)
												data = {}
												data['status_code'] = "200 OK"
												data['status'] = "Sheet renamed successfully"
												data = json.dumps(data)
												return data
												
										else:
											data = {}
											data['status_code'] = "401"
											data['status'] = "Sheet name not existed"
											return data
											pass		
										
									file_data = remane_sheet(file_name)
									print(file_data)
								
								else:
									data = {}
									data['status_code'] = "401"
									data['status'] = "Please provide proper Sheet name"
									print(json.dumps(data))
									pass
							
							else:
									data = {}
									data['status_code'] = "401"
									data['status'] = "Not able rename sheet name with '\\','/','*','[',']',':','?'"
									print(json.dumps(data))
									pass
									
						else:
							data = {}
							data['status_code'] = "401"
							data['status'] = "Sheet name allows only 31 characters length"
							print(json.dumps(data))
							pass				
					
					elif file_extension in (''):
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass
						
					else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file extension"
						print(json.dumps(data))
						pass
						
				else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass	
						
			else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide Existing file name"
						print(json.dumps(data))
						pass			
		
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e)
			print(json.dumps(data))
			pass
			
			
	elif commandKey == "DeleteCell":
		try :
			if os.path.exists(sys.argv[2]): 
				filename, file_extension = os.path.splitext(sys.argv[2])			
				if filename[-1] not in (' '):
					if sys.argv[2].endswith('xls') or sys.argv[2].endswith('xlsx'):	
						file_name =  sys.argv[2]	
						sheet_name = sys.argv[3]	
						temp_list1 = sys.argv[4]
						def my_split(s):
							return filter(None, re.split(r'(\d+)', s))
						
						def delete_cell(file_name):							
							rowindexarray = my_split(temp_list1)
							cval = rowindexarray[0]
							rval = col_to_num(cval)
							cval1 = rowindexarray[1]
							rb = xlrd.open_workbook(file_name)
							wb = xl_copy(rb)
							sheet = wb.get_sheet(sheet_name)
							rval = int(rval) -1
							cval1 = int(cval1) - 1
							sheet.write(cval1,rval,'')
							wb.save(file_name)							
							data = {}
							data['status_code'] = "200 OK"
							data['status'] = "Cell deleted successfully"
							data = json.dumps(data)
							return data
						
						file_data_delete = delete_cell(file_name)
						print(file_data_delete)	
															
					elif file_extension in (''):
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass
						
					else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file extension"
						print(json.dumps(data))
						pass	
						
				else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass			
						
			else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide Existing file name"
						print(json.dumps(data))
						pass				
				
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e)
			print(json.dumps(data))
			pass		

			
	elif commandKey == "FindReplace":
		try :
			if os.path.exists(sys.argv[2]): 
				filename, file_extension = os.path.splitext(sys.argv[2])			
				if filename[-1] not in (' '):
					if sys.argv[2].endswith('xls') or sys.argv[2].endswith('xlsx'):	
						file_name =  sys.argv[2]	
						sheet_name = sys.argv[3]	
						temp_list1 = sys.argv[4]
						replace_text = sys.argv[5]
												
						temp_list_in = temp_list1.split(',')
						cell_text_list_in = list(temp_list_in)
						temp_list_out = replace_text.split(',')
						cell_text_list_out = list(temp_list_out)
						
						if len(cell_text_list_in) ==1 and len(cell_text_list_out) ==1:
							def my_split(s):
								return filter(None, re.split(r'(\d+)', s))
							
							def find_replace(file_name):							
								rowindexarray = my_split(temp_list1)
								cval = rowindexarray[0]
								rval = col_to_num(cval)
								cval1 = rowindexarray[1]
								rb = xlrd.open_workbook(file_name)
								wb = xl_copy(rb)
								sheet = wb.get_sheet(sheet_name)
								rval = int(rval) -1
								cval1 = int(cval1) - 1
								sheet.write(cval1,rval,replace_text)
								wb.save(file_name)							
								data = {}
								data['status_code'] = "200 OK"
								data['status'] = "Cell Updated successfully"
								data = json.dumps(data)
								return data
							
							file_data_find = find_replace(file_name)
							print(file_data_find)	
																		
						else:
							data = {}
							data['status_code'] = "401"
							data['status'] = "Please provide proper format"
							print(json.dumps(data))
							pass	
															
					elif file_extension in (''):
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass
						
					else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file extension"
						print(json.dumps(data))
						pass	
						
				else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass			
						
			else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide Existing file name"
						print(json.dumps(data))
						pass				
				
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e)
			print(json.dumps(data))
			pass	
			

	elif commandKey == "SetCellValue":
		try :
			if os.path.exists(sys.argv[2]): 
				filename, file_extension = os.path.splitext(sys.argv[2])			
				if filename[-1] not in (' '):
					if sys.argv[2].endswith('xls') or sys.argv[2].endswith('xlsx'):	
						file_name =  sys.argv[2]	
						sheet_name = sys.argv[3]	
						temp_list1 = sys.argv[4]
						replace_text = sys.argv[5]
						def my_split(s):
							return filter(None, re.split(r'(\d+)', s))
						
						def set_cell_value(file_name,sheet_name,temp_list1,replace_text):							
							book = xlrd.open_workbook(file_name)
							sh = book.sheet_by_name(sheet_name)
							book1 = Workbook()
							rb = xlrd.open_workbook(file_name)
							wb = xl_copy(rb)
							sheet = wb.get_sheet(sheet_name)
							
							#original cell
							rowindexarray = my_split(temp_list1)
							cval = rowindexarray[0]
							rval = col_to_num(cval)
							cval1 = rowindexarray[1]
							rval = int(rval) -1
							cval1 = int(cval1) - 1
							
						    #replace cell index
							rowindexarray_replce = my_split(replace_text)
							cval_replace = rowindexarray_replce[0]
							rval_replace_cell = col_to_num(cval_replace)
							cval_replace_cell = rowindexarray_replce[1]
							rval_replace_cell = int(rval_replace_cell) -1
							cval_replace_cell = int(cval_replace_cell) - 1
							
							value = sh.cell_value(cval1, rval)
							sheet.write(cval_replace_cell,rval_replace_cell,value)
							
							#sheet.write(cval1,rval,replace_text)
							wb.save(file_name)							
							data = {}
							data['status_code'] = "200 OK"
							data['status'] = "Cell Updated successfully"
							data = json.dumps(data)
							return data
												
						temp_list_in = temp_list1.split(',')
						cell_text_list_in = list(temp_list_in)
						temp_list_out = replace_text.split(',')
						cell_text_list_out = list(temp_list_out)
						
						if len(cell_text_list_in) ==1 and len(cell_text_list_out) ==1:
							file_data_cellvalue = set_cell_value(file_name,sheet_name,temp_list1,replace_text)
							print(file_data_cellvalue)	
							
						else:
							data = {}
							data['status_code'] = "401"
							data['status'] = "Please provide proper format"
							print(json.dumps(data))
							pass	
							
							
					elif file_extension in (''):
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass
						
					else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file extension"
						print(json.dumps(data))
						pass	
						
				else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass			
						
			else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide Existing file name"
						print(json.dumps(data))
						pass				
				
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e)
			print(json.dumps(data))
			pass	
			

	elif commandKey == "GetCellValue":
		try :
			if os.path.exists(sys.argv[2]): 
				filename, file_extension = os.path.splitext(sys.argv[2])			
				if filename[-1] not in (' '):
					if sys.argv[2].endswith('xls') or sys.argv[2].endswith('xlsx'):	
						file_name =  sys.argv[2]	
						sheet_name = sys.argv[3]	
						temp_list1 = sys.argv[4]
						lowercase_letters = [c for c in temp_list1 if c.islower()]
						if len(lowercase_letters) > 0:
							data = {}
							data['status_code'] = "401"
							data['status'] = "Please provide proper cell format"
							print(json.dumps(data))
							pass
							
						else:
							temp_list_in = temp_list1.split(',')
							cell_text_list_in = list(temp_list_in)						
							if len(cell_text_list_in) == 1:				
								file_data_cellvalueget = reading_indices(file_name, sheet_name,temp_list1)
								print(file_data_cellvalueget)	
								
							else:
								data = {}
								data['status_code'] = "401"
								data['status'] = "Please provide proper format"
								print(json.dumps(data))
								pass											
										
					elif file_extension in (''):
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass
						
					else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file extension"
						print(json.dumps(data))
						pass	
						
				else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass			
						
			else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide Existing file name"
						print(json.dumps(data))
						pass				
				
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e)
			print(json.dumps(data))
			pass	
			
			
	elif commandKey == "SetCellRange":
		try :
			if os.path.exists(sys.argv[2]): 				
				filename, file_extension = os.path.splitext(sys.argv[2])			
				if filename[-1] not in (' '):
					if sys.argv[2].endswith('xls') or sys.argv[2].endswith('xlsx'):	
						file_name =  sys.argv[2]	
						sheet_name = sys.argv[3]	
						
						#cells 
						temp_list_in = sys.argv[4]
						temp_list2 = temp_list_in.split(',')
						cell_text_list = list(temp_list2)
						
						#values
						replace_text_in = sys.argv[5]
						replace_text_in1 = replace_text_in.split(',')
						replace_text_list = list(replace_text_in1)
						
						def my_split(s):
							return filter(None, re.split(r'(\d+)', s))
						
						def set_cell_range(file_name,cell,replace_col,sheet_name):	
													
							book = xlrd.open_workbook(file_name)
							sh = book.sheet_by_name(sheet_name)
							book1 = Workbook()
							rb = xlrd.open_workbook(file_name)
							wb = xl_copy(rb)
							sheet = wb.get_sheet(sheet_name)
							
							#original cells index
							rowindexarray = my_split(cell)
							cval = rowindexarray[0]
							rval = col_to_num(cval)
							cval1 = rowindexarray[1]
							rval = int(rval) -1
							cval1 = int(cval1) - 1
							
							#replace cells index
							rowindexarray_replce = my_split(replace_col)
							cval_replace = rowindexarray_replce[0]
							rval_replace_cell = col_to_num(cval_replace)
							cval_replace_cell = rowindexarray_replce[1]
							rval_replace_cell = int(rval_replace_cell) -1
							cval_replace_cell = int(cval_replace_cell) - 1
							
							value = sh.cell_value(cval1, rval)
							sheet.write(cval_replace_cell,rval_replace_cell,value)
							wb.save(file_name)							
							data = {}
							data['status_code'] = "200 OK"
							data['status'] = "Cell range updated successfully"
							data = json.dumps(data)
							return data
												
						if len(cell_text_list) == len(replace_text_list):						
							i = 0	
							for cell in cell_text_list:
								replace_col = replace_text_list[i]
								file_data_cellvalue = set_cell_range(file_name,cell,replace_col,sheet_name)
								i = i + 1
								
							print(file_data_cellvalue)	
							
						else:
							data = {}
							data['status_code'] = "401"
							data['status'] = "Please provide proper formate"
							print(json.dumps(data))
							pass		
							
															
					elif file_extension in (''):
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass
						
					else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file extension"
						print(json.dumps(data))
						pass	
						
				else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass			
						
			else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide Existing file name"
						print(json.dumps(data))
						pass				
				
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e)
			print(json.dumps(data))
			pass	


	elif commandKey == "GetCellRange":
		try :
			if os.path.exists(sys.argv[2]): 
				filename, file_extension = os.path.splitext(sys.argv[2])			
				if filename[-1] not in (' '):
					if sys.argv[2].endswith('xls') or sys.argv[2].endswith('xlsx'):	
						file_name =  sys.argv[2]	
						sheet_name = sys.argv[3]	
						#temp_list1 = sys.argv[4]
						#cells 
						temp_list_in1 = sys.argv[4]
						temp_val_list2 = temp_list_in1.split(',')
						cell_val_list = list(temp_val_list2)
						
						data1 = {}
						#data1['status_code'] = "200 OK"
						for cellval in cell_val_list:
							file_data_cellvalueget = reading_indices(file_name, sheet_name,cellval)
							file_data_cellvalueget = json.loads(file_data_cellvalueget)
							data1[cellval] = file_data_cellvalueget['data']
						
						data = {}
						data['status_code'] = "200 OK"
						data['data'] = data1						
						data = json.dumps(data)	
						print(json.dumps(data))	
															
					elif file_extension in (''):
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass
						
					else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file extension"
						print(json.dumps(data))
						pass	
						
				else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass			
						
			else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide Existing file name"
						print(json.dumps(data))
						pass				
				
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e)
			print(json.dumps(data))
			pass				
			
			
	elif commandKey == "SetRow":
		try :
			if os.path.exists(sys.argv[2]): 				
				filename, file_extension = os.path.splitext(sys.argv[2])			
				if filename[-1] not in (' '):
					if sys.argv[2].endswith('xls') or sys.argv[2].endswith('xlsx'):
						file_name =  sys.argv[2]	
						sheet_name =  sys.argv[3]
						column_in =  sys.argv[4]
						column_out =  sys.argv[5]
						#temp_list1 = argv2
						#temp_list2 = temp_list1.split(',')
						#temp_list = list(map(int, temp_list2))
						#print(temp_list)
						#temp_list = [1, 4, 5, 6, 7, 8, 9, 10, 15, 19, 26]
						def set_row(file_name):
							book = xlrd.open_workbook(file_name)
							sh = book.sheet_by_name(sheet_name)
							book1 = Workbook()
							#sheet1 = book1.add_sheet('test', cell_overwrite_ok=True)
							rb = xlrd.open_workbook(file_name)
							wb = xl_copy(rb)
							sheet = wb.get_sheet(sheet_name)
							total_cols = sh.ncols
							k = int(column_out) - 1
							#print(sys.argv[5])
							#print(column_in)
							column_in =  sys.argv[4]
							column_in = int(sys.argv[4]) -1 
							for i in range(0, total_cols):
								value = sh.cell_value(column_in,i)
								sheet.write(k, i, value)
								
							wb.save(file_name)
							data = {}
							data['status_code'] = "200 OK"
							data['status'] = "Row updated successfully"
							data = json.dumps(data)
							return data
						
						if int(column_in) > 0 and int(column_out) > 0:
							file_data_row = set_row(file_name)
							print(file_data_row)	
							
						else:
							data = {}
							data['status_code'] = "401"
							data['status'] = "Please provide proper index values"
							print(json.dumps(data))
							pass		
					
					elif file_extension in (''):
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass
						
					else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file extension"
						print(json.dumps(data))
						pass	
						
				else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass			
						
			else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide Existing file name"
						print(json.dumps(data))
						pass		
		
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e)
			print(json.dumps(data))
			pass
			
			
	elif commandKey == "GetRow":
		try :
			if os.path.exists(sys.argv[2]): 				
				filename, file_extension = os.path.splitext(sys.argv[2])			
				if filename[-1] not in (' '):
					if sys.argv[2].endswith('xls') or sys.argv[2].endswith('xlsx'):
						file_name =  sys.argv[2]	
						sheet_name =  sys.argv[3]
						column_in =  sys.argv[4]
						#temp_list1 = argv2
						#temp_list2 = temp_list1.split(',')
						#temp_list = list(map(int, temp_list2))
						#print(temp_list)
						#temp_list = [1, 4, 5, 6, 7, 8, 9, 10, 15, 19, 26]
						def get_row(file_name):
							book = xlrd.open_workbook(file_name)
							sh = book.sheet_by_name(sheet_name)
							book1 = Workbook()
							#sheet1 = book1.add_sheet('test', cell_overwrite_ok=True)
							rb = xlrd.open_workbook(file_name)
							wb = xl_copy(rb)
							sheet = wb.get_sheet(sheet_name)
							total_cols = sh.ncols
							#print(sys.argv[5])
							#print(column_in)
							column_in =  sys.argv[4]
							column_in = int(sys.argv[4]) -1 
							data = {}
							data_array = {}
							for i in range(0, total_cols):
								value = sh.cell_value(column_in,i)
								data_array[str( sys.argv[4]) +"*"+ str(i+1)] =  value
							data = {}
							data['status_code'] = "200 OK"
							data['data'] = data_array
							data = json.dumps(data)
							return data
						
						file_data_row = get_row(file_name)
						print(file_data_row)	
					
					elif file_extension in (''):
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass
						
					else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file extension"
						print(json.dumps(data))
						pass	
						
				else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass			
						
			else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide Existing file name"
						print(json.dumps(data))
						pass		
		
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e)
			print(json.dumps(data))
			pass	
			

	elif commandKey == "SetRows":
		try :
			if os.path.exists(sys.argv[2]): 				
				filename, file_extension = os.path.splitext(sys.argv[2])			
				if filename[-1] not in (' '):
					if sys.argv[2].endswith('xls') or sys.argv[2].endswith('xlsx'):
						file_name =  sys.argv[2]	
						sheet_name =  sys.argv[3]
						
						#row input 
						rows_in = sys.argv[4]
						rows_in2 = rows_in.split(',')
						rows_in_list1 = list(rows_in2)
						rows_in_list = list(map(int, rows_in_list1))
						
						#row output
						rows_out = sys.argv[5]
						rows_out2 = rows_out.split(',')
						rows_out_list1 = list(rows_out2)
						rows_out_list = list(map(int, rows_out_list1))
			
						def set_rows(file_name):
							book = xlrd.open_workbook(file_name)
							sh = book.sheet_by_name(sheet_name)
							book1 = Workbook()
							rb = xlrd.open_workbook(file_name)
							wb = xl_copy(rb)
							sheet = wb.get_sheet(sheet_name)
							total_cols = sh.ncols
							k = 0
							for j in rows_in_list: #for rows
								for i in range(0, total_cols):
									value = sh.cell_value(int(j-1),i)
									sheet.write(rows_out_list[k]-1, i, value)
								k = k+1
									
								
							wb.save(file_name)
							data = {}
							data['status_code'] = "200 OK"
							data['status'] = "Rows updated successfully"
							data = json.dumps(data)
							return data
						
						if len(rows_in_list) == len(rows_out_list):	
							file_data_rows = set_rows(file_name)
							print(file_data_rows)

						else:
							data = {}
							data['status_code'] = "401"
							data['status'] = "Please provide proper format"
							print(json.dumps(data))
							pass							
					
					elif file_extension in (''):
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass
						
					else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file extension"
						print(json.dumps(data))
						pass	
						
				else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass			
						
			else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide Existing file name"
						print(json.dumps(data))
						pass		
		
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e)
			print(json.dumps(data))
			pass
			
			
			
	elif commandKey == "GetRows":
		try :
			if os.path.exists(sys.argv[2]): 				
				filename, file_extension = os.path.splitext(sys.argv[2])			
				if filename[-1] not in (' '):
					if sys.argv[2].endswith('xls') or sys.argv[2].endswith('xlsx'):
						file_name =  sys.argv[2]	
						sheet_name =  sys.argv[3]
						
						#row input 
						rows_in = sys.argv[4]
						rows_in2 = rows_in.split(',')
						rows_in_list1 = list(rows_in2)
						rows_in_list = list(map(int, rows_in_list1))

						def get_rows(file_name):
							book = xlrd.open_workbook(file_name)
							sh = book.sheet_by_name(sheet_name)
							book1 = Workbook()
							rb = xlrd.open_workbook(file_name)
							wb = xl_copy(rb)
							sheet = wb.get_sheet(sheet_name)
							total_cols = sh.ncols
							data_array = {}
							for j in rows_in_list: #for rows
								for i in range(0, total_cols):
									value = sh.cell_value(int(j-1),i)
									data_array[str(j) +"*"+ str(i+1)] =  value
									
							data = {}
							data['status_code'] = "200 OK"
							data['data'] = data_array
							data = json.dumps(data)
							return data
						
						file_data_rowsget = get_rows(file_name)
						print(file_data_rowsget)	
					
					elif file_extension in (''):
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass
						
					else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file extension"
						print(json.dumps(data))
						pass	
						
				else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass			
						
			else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide Existing file name"
						print(json.dumps(data))
						pass		
		
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e)
			print(json.dumps(data))
			pass
			
			
	elif commandKey == "SetColumn":
		try :
			if os.path.exists(sys.argv[2]): 				
				filename, file_extension = os.path.splitext(sys.argv[2])			
				if filename[-1] not in (' '):
					if sys.argv[2].endswith('xls') or sys.argv[2].endswith('xlsx'):
						file_name =  sys.argv[2]	
						sheet_name =  sys.argv[3]
						column_in_pre =  sys.argv[4]
						column_out_pre =  sys.argv[5]
						
						column_in = col_to_num(column_in_pre) -1
						column_out = col_to_num(column_out_pre) -1
						def set_col(file_name):
							book = xlrd.open_workbook(file_name)
							sh = book.sheet_by_name(sheet_name)
							book1 = Workbook()
							rb = xlrd.open_workbook(file_name)
							wb = xl_copy(rb)
							sheet = wb.get_sheet(sheet_name)
							total_rows = sh.nrows
							k = int(column_out)
 
							for i in range(0, total_rows):
								value = sh.cell_value(i, column_in)
								sheet.write(i, k, value)
								
							wb.save(file_name)
							data = {}
							data['status_code'] = "200 OK"
							data['status'] = "Column updated successfully"
							data = json.dumps(data)
							return data
						
						file_data_col = set_col(file_name)
						print(file_data_col)	
					
					elif file_extension in (''):
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass
						
					else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file extension"
						print(json.dumps(data))
						pass	
						
				else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass			
						
			else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide Existing file name"
						print(json.dumps(data))
						pass		
		
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e)
			print(json.dumps(data))
			pass
			
					
	elif commandKey == "GetColumn":
		try :
			if os.path.exists(sys.argv[2]): 				
				filename, file_extension = os.path.splitext(sys.argv[2])			
				if filename[-1] not in (' '):
					if sys.argv[2].endswith('xls') or sys.argv[2].endswith('xlsx'):
						file_name =  sys.argv[2]	
						sheet_name =  sys.argv[3]
						column_in_pre =  sys.argv[4]
						
						column_in = col_to_num(column_in_pre) -1
						def set_col(file_name):
							book = xlrd.open_workbook(file_name)
							sh = book.sheet_by_name(sheet_name)
							book1 = Workbook()
							rb = xlrd.open_workbook(file_name)
							wb = xl_copy(rb)
							sheet = wb.get_sheet(sheet_name)
							total_rows = sh.nrows
 
							data_array = {}
							for i in range(0, total_rows):
								value = sh.cell_value(i, column_in)
								data_array[str(i + 1) +"*"+ str(column_in + 1)] =  value
									
							data = {}
							data['status_code'] = "200 OK"
							data['data'] = data_array
							data = json.dumps(data)
							return data
						
						file_data_col = set_col(file_name)
						print(file_data_col)	
					
					elif file_extension in (''):
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass
						
					else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file extension"
						print(json.dumps(data))
						pass	
						
				else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass			
						
			else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide Existing file name"
						print(json.dumps(data))
						pass		
		
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e)
			print(json.dumps(data))
			pass
			
			
			
	elif commandKey == "SetColumns":
		try :
			if os.path.exists(sys.argv[2]): 				
				filename, file_extension = os.path.splitext(sys.argv[2])			
				if filename[-1] not in (' '):
					if sys.argv[2].endswith('xls') or sys.argv[2].endswith('xlsx'):
						file_name =  sys.argv[2]	
						sheet_name =  sys.argv[3]
						#column_in_pre =  sys.argv[4]
						#column_out_pre =  sys.argv[5]
						
						#col input 
						cols_in = sys.argv[4]
						cols_in2 = cols_in.split(',')
						cols_in_list1 = list(cols_in2)
						cols_in_list = list(cols_in_list1)
						
						#row output
						cols_out = sys.argv[5]
						cols_out2 = cols_out.split(',')
						cols_out_list1 = list(cols_out2)
						cols_out_list = list(cols_out_list1)
						
						def set_col_range(file_name,col,replace_col,sheet_name):
							book = xlrd.open_workbook(file_name)
							sh = book.sheet_by_name(sheet_name)
							book1 = Workbook()
							rb = xlrd.open_workbook(file_name)
							wb = xl_copy(rb)
							sheet = wb.get_sheet(sheet_name)
							column_in = col_to_num(col) -1
							column_out = col_to_num(replace_col) -1
							total_rows = sh.nrows
							k = int(column_out)
 
							for i in range(0, total_rows):
								value = sh.cell_value(i, column_in)
								sheet.write(i, k, value)
								
							wb.save(file_name)
							data = {}
							data['status_code'] = "200 OK"
							data['status'] = "Columns updated successfully"
							data = json.dumps(data)
							return data
						
						if len(cols_in_list) == len(cols_out_list):	
							i = 0	
							for col in cols_in_list:
								replace_col = cols_out_list[i]
								file_data_colvalues = set_col_range(file_name,col,replace_col,sheet_name)
								i = i + 1
								
							print(file_data_colvalues)	
						
						
						else:
							data = {}
							data['status_code'] = "401"
							data['status'] = "Please provide proper format"
							print(json.dumps(data))
							pass	
							
					
					elif file_extension in (''):
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass
						
					else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file extension"
						print(json.dumps(data))
						pass	
						
				else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass			
						
			else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide Existing file name"
						print(json.dumps(data))
						pass		
		
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e)
			print(json.dumps(data))
			pass
					
			
			
	elif commandKey == "GetColumns":
		try :
			if os.path.exists(sys.argv[2]): 				
				filename, file_extension = os.path.splitext(sys.argv[2])			
				if filename[-1] not in (' '):
					if sys.argv[2].endswith('xls') or sys.argv[2].endswith('xlsx'):
						file_name =  sys.argv[2]	
						sheet_name =  sys.argv[3]
						#column_in_pre =  sys.argv[4]
						#column_out_pre =  sys.argv[5]
						
						#col input 
						cols_in = sys.argv[4]
						cols_in2 = cols_in.split(',')
						cols_in_list1 = list(cols_in2)
						cols_in_list = list(cols_in_list1)
												
						def get_col_range(file_name,col,sheet_name):
							book = xlrd.open_workbook(file_name)
							sh = book.sheet_by_name(sheet_name)
							book1 = Workbook()
							rb = xlrd.open_workbook(file_name)
							wb = xl_copy(rb)
							sheet = wb.get_sheet(sheet_name)
							column_in = col_to_num(col) -1
							total_rows = sh.nrows
							
							data_array = {}
							for i in range(0, total_rows):
								value = sh.cell_value(i, column_in)
								data_array[str(i + 1) +"*"+ str(column_in + 1)] =  value
								
							data = {}
							data['status_code'] = "200 OK"
							data['status'] = "Columns updated successfully"
							data['data'] = data_array
							data = json.dumps(data)
							return data
						
						
						data_col = {}
						for col in cols_in_list:
							file_data_colvalues = get_col_range(file_name,col,sheet_name)
							file_data_colvalueget = json.loads(file_data_colvalues)
							data_col[col] = file_data_colvalueget['data']
							
							
						data_col = json.dumps(data_col)	
						print(data_col)		
						
					
					elif file_extension in (''):
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass
						
					else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file extension"
						print(json.dumps(data))
						pass	
						
				else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass			
						
			else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide Existing file name"
						print(json.dumps(data))
						pass		
		
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e)
			print(json.dumps(data))
			pass
			
			
			
	elif commandKey == "WriteSheet":
		try :
			if os.path.exists(sys.argv[2]): 				
				filename, file_extension = os.path.splitext(sys.argv[2])			
				if filename[-1] not in (' '):
					if sys.argv[2].endswith('xls') or sys.argv[2].endswith('xlsx'):	
						file_name =  sys.argv[2]	
						sheet_name = sys.argv[3]	
						
						#cells 
						temp_list_in = sys.argv[4]
						temp_list2 = temp_list_in.split(',')
						cell_text_list = list(temp_list2)
						
						#values
						replace_text_in = sys.argv[5]
						replace_text_in1 = replace_text_in.split(',')
						replace_text_list = list(replace_text_in1)
						
						def my_split(s):
							return filter(None, re.split(r'(\d+)', s))
						
						def set_cell_range(file_name,cell,replace_value,sheet_name):							
							rowindexarray = my_split(cell)
							cval = rowindexarray[0]
							if cval.islower():
								data = {}
								data['status_code'] = "401"
								data['status'] = "Please provide proper column format"
								return data
							
							else:
								rval = col_to_num(cval)
								cval1 = rowindexarray[1]
								rb = xlrd.open_workbook(file_name)
								wb = xl_copy(rb)
								sheet = wb.get_sheet(sheet_name)
								rval = int(rval) -1
								cval1 = int(cval1) - 1
								sheet.write(cval1,rval,replace_value)
								wb.save(file_name)							
								data = {}
								data['status_code'] = "200 OK"
								data['status'] = "data written successfully"
								data = json.dumps(data)
								return data
						
						if len(cell_text_list) == len(replace_text_list):							
							i = 0	
							for cell in cell_text_list:
								replace_value = replace_text_list[i]
								replace_value = replace_value.strip()
								if not replace_value:
									data = {}
									data['status_code'] = "401"
									data['status'] = "Please provide the data"
									print(json.dumps(data))
									exit()	
									
								else:
									file_data_cellvalue = set_cell_range(file_name,cell,replace_value,sheet_name)
									i = i + 1
								
							print(file_data_cellvalue)	
							
						else:
							data = {}
							data['status_code'] = "401"
							data['status'] = "Please provide proper format"
							data = json.dumps(data)
							print(json.dumps(data))
							pass		
							
															
					elif file_extension in (''):
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						data = json.dumps(data)
						print(json.dumps(data))
						pass
						
					else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file extension"
						data = json.dumps(data)
						print(json.dumps(data))
						pass	
						
				else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						data = json.dumps(data)
						print(json.dumps(data))
						pass			
						
			else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide Existing file name"
						data = json.dumps(data)
						print(json.dumps(data))
						pass				
				
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e)
			print(json.dumps(data))
			pass	
		
			
	
	elif commandKey == "BrowseExcel":	
		try:
			def browse_file():	
				Tkinter.Tk().withdraw() # Close the root window
				in_path = tkFileDialog.askopenfilename()
				if in_path.endswith('xls') or in_path.endswith('xlsx'):
					#print in_path
					data = {}
					data['file'] = in_path
					#data['status'] = "File created successfully"
					data = json.dumps(data)
					return data
					
				else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please select excel file"
						return data
						pass	
						
				
			browse_file_path = browse_file()
			print(browse_file_path)
			
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e)
			print(json.dumps(data))
			pass		
			
	
	elif commandKey == "CopyExcel":	
		try:
			if os.path.exists(sys.argv[3]): 
				data = {}
				data['status_code'] = "401"
				data['status'] = "File name already exists"
				print(json.dumps(data))
				pass
					
			else: 
				filename_in, file_extension_in = os.path.splitext(sys.argv[2])
				filename_in1, file_extension_in1 =  sys.argv[2].split(".")
				filename_out, file_extension_out = sys.argv[3].split(".")
				filename_in, file_extension_out1 = os.path.splitext(sys.argv[3])
				if filename_out not in ("") and filename_in1 not in (""):
					if filename_out[-1] not in ('','/','\\') and filename_in1[-1] not in ('','/','\\'):
						if filename_out[-1] not in (' ') and filename_in1[-1] not in (' '):		
							if sys.argv[2].endswith(file_extension_in1) and sys.argv[3].endswith(file_extension_out):	
								if file_extension_in in ('.xls', '.xlsx') and file_extension_out1 in ('.xls', '.xlsx'):
									if file_extension_in1 == file_extension_out :
										argv2 =  sys.argv[2]
										argv3 =  sys.argv[3]
										def copy_paste_file(argv2, argv3):
											shutil.copy2(argv2, argv3)
											#another way to copy file
											#shutil.copyfile('/Users/pankaj/abc.txt', '/Users/pankaj/abc_copyfile.txt')
											data = {}
											data['status_code'] = "200 OK"
											data['status'] = "File Copied successfully"
											data = json.dumps(data)
											return data
											
										copy_file_content = copy_paste_file(argv2, argv3)	
										print(copy_file_content)
										
									
									elif filename_out[-1] in ('','/'):
										data = {}
										data['status_code'] = "401"
										data['status'] = "Please provide proper file name"
										print(json.dumps(data))
										pass	
										
									else:
										data = {}
										data['status_code'] = "401"
										data['status'] = "Please provide proper file extension"
										print(json.dumps(data))
										pass
										
								else:
									data = {}
									data['status_code'] = "401"
									data['status'] = "Please provide proper file extension"
									print(json.dumps(data))
									pass
											
														
							else:
								data = {}
								data['status_code'] = "401"
								data['status'] = "Please provide proper file extension"
								print(json.dumps(data))
								pass	
								
											
						else:
								data = {}
								data['status_code'] = "401"
								data['status'] = "Please provide proper file name"
								print(json.dumps(data))
								pass	
								
					else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass		
					
				else:
					data = {}
					data['status_code'] = "401"
					data['status'] = "Please provide proper file name"
					print(json.dumps(data))
					pass
				
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e)
			print(json.dumps(data))
			pass
			
			
			
	elif commandKey == "MoveExcel":	
		try:
			if os.path.exists(sys.argv[3]): 
				data = {}
				data['status_code'] = "401"
				data['status'] = "File name already exists"
				print(json.dumps(data))
				pass
					
			else: 
				filename_in, file_extension_in = os.path.splitext(sys.argv[2])
				filename_in1, file_extension_in1 =  sys.argv[2].split(".")
				filename_out, file_extension_out = sys.argv[3].split(".")
				filename_in, file_extension_out1 = os.path.splitext(sys.argv[3])
				if filename_out not in ("") and filename_in1 not in (""):
					if filename_out[-1] not in ('','/','\\') and filename_in1[-1] not in ('','/','\\'):
						if filename_out[-1] not in (' ') and filename_in1[-1] not in (' '):		
							if sys.argv[2].endswith(file_extension_in1) and sys.argv[3].endswith(file_extension_out):	
								if file_extension_in in ('.xls', '.xlsx') and file_extension_out1 in ('.xls', '.xlsx'):
									if file_extension_in1 == file_extension_out :
										argv2 =  sys.argv[2]
										argv3 =  sys.argv[3]
										def move_file(argv2, argv3):
											#shutil.move(argv2, argv3)
											#another way to copy file
											os.rename(argv2, argv3)
											#shutil.copyfile('/Users/pankaj/abc.txt', '/Users/pankaj/abc_copyfile.txt')
											data = {}
											data['status_code'] = "200 OK"
											data['status'] = "File Moved successfully"
											data = json.dumps(data)
											return data
											
										copy_file_content = move_file(argv2, argv3)	
										print(copy_file_content)
										
									
									elif filename_out[-1] in ('','/'):
										data = {}
										data['status_code'] = "401"
										data['status'] = "Please provide proper file name"
										print(json.dumps(data))
										pass	
										
									else:
										data = {}
										data['status_code'] = "401"
										data['status'] = "Please provide proper file extension"
										print(json.dumps(data))
										pass
										
								else:
									data = {}
									data['status_code'] = "401"
									data['status'] = "Please provide proper file extension"
									print(json.dumps(data))
									pass
											
														
							else:
								data = {}
								data['status_code'] = "401"
								data['status'] = "Please provide proper file extension"
								print(json.dumps(data))
								pass	
								
											
						else:
								data = {}
								data['status_code'] = "401"
								data['status'] = "Please provide proper file name"
								print(json.dumps(data))
								pass	
								
					else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass		
					
				else:
					data = {}
					data['status_code'] = "401"
					data['status'] = "Please provide proper file name"
					print(json.dumps(data))
					pass
				
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e)
			print(json.dumps(data))
			pass	
			
			
	elif commandKey == "DeleteExcel":	
		try:
			if os.path.exists(sys.argv[2]): 
				filename_in, file_extension_in = os.path.splitext(sys.argv[2])
				filename_in1, file_extension_in1 =  sys.argv[2].split(".")
				if filename_in1 not in (""):
					if filename_in1[-1] not in ('','/','\\'):
						if filename_in1[-1] not in (' '):			
								if sys.argv[2].endswith('xls') or sys.argv[2].endswith('xlsx'):
										argv2 =  sys.argv[2]
										def delete_file(argv2):
											shutil.os.remove(argv2)	
											data = {}
											data['status_code'] = "200 OK"
											data['status'] = "File Deleted successfully"
											data = json.dumps(data)
											return data
											
										copy_file_content = delete_file(argv2)	
										print(copy_file_content)
										
								else:
									data = {}
									data['status_code'] = "401"
									data['status'] = "Please provide proper file extension"
									print(json.dumps(data))
									pass											
											
						else:
								data = {}
								data['status_code'] = "401"
								data['status'] = "Please provide proper file name"
								print(json.dumps(data))
								pass	
								
					else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass		
					
				else:
					data = {}
					data['status_code'] = "401"
					data['status'] = "Please provide proper file name"
					print(json.dumps(data))
					pass
					
			else:
				data = {}
				data['status_code'] = "401"
				data['status'] = "Please provide existing file"
				print(json.dumps(data))
				pass
								
				
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e)
			print(json.dumps(data))
			pass	
	
	
	elif commandKey == "DeleteColumn":
		try :
			if os.path.exists(sys.argv[2]): 				
				filename, file_extension = os.path.splitext(sys.argv[2])			
				if filename[-1] not in (' '):
					if sys.argv[2].endswith('xls') or sys.argv[2].endswith('xlsx'):
						file_name =  sys.argv[2]	
						sheet_name =  sys.argv[3]
						column_in_pre =  sys.argv[4]
						#column_out_pre =  sys.argv[5]
						colNum = findexcel_column(column_in_pre,file_name,sheet_name)
						colNum123=colNum


						def set_col(file_name):
							book = xlrd.open_workbook(file_name)
							sh = book.sheet_by_name(sheet_name)
							book1 = Workbook()
							rb = xlrd.open_workbook(file_name)
							wb = xl_copy(rb)
							sheet = wb.get_sheet(sheet_name)
							total_rows = sh.nrows
							k = int(colNum)
 
							for i in range(0, total_rows):
								#value = sh.cell_value(i, column_in)
								sheet.write(i, k, "")
								
							wb.save(file_name)
							data = {}
							data['status_code'] = "200 OK"
							data['status'] = "Column removed successfully"
							data = json.dumps(data)
							return data
							
						if colNum123 != None:
							file_data_col = set_col(file_name)
							print(file_data_col)
							
						else:
							data = {}
							data['status_code'] = "401"
							data['status'] = "Please provide existing column name"
							print(json.dumps(data))
							pass
					
					elif file_extension in (''):
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass
						
					else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file extension"
						print(json.dumps(data))
						pass	
						
				else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass			
						
			else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide Existing file name"
						print(json.dumps(data))
						pass		
		
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e)
			print(json.dumps(data))
			pass
			
			
			
	elif commandKey == "DeleteColumns":
		try :
			if os.path.exists(sys.argv[2]): 				
				filename, file_extension = os.path.splitext(sys.argv[2])			
				if filename[-1] not in (' '):
					if sys.argv[2].endswith('xls') or sys.argv[2].endswith('xlsx'):
						file_name =  sys.argv[2]	
						sheet_name =  sys.argv[3]
						#column_in_pre =  sys.argv[4]
						#column_out_pre =  sys.argv[5]
						
						#col input 
						cols_in = sys.argv[4]
						cols_in2 = cols_in.split(',')
						cols_in_list1 = list(cols_in2)
						cols_in_list = list(cols_in_list1)

									
						def set_col_range(file_name,col,replace_col,sheet_name):
							book = xlrd.open_workbook(file_name)
							sh = book.sheet_by_name(sheet_name)
							book1 = Workbook()
							rb = xlrd.open_workbook(file_name)
							wb = xl_copy(rb)
							sheet = wb.get_sheet(sheet_name)
							#column_in = col_to_num(col) -1
							#column_out = col_to_num(replace_col) -1
							
							column_out = findexcel_column(col,file_name,sheet_name)
							column_out123=column_out
							total_rows = sh.nrows

							k = int(column_out)
 
							for i in range(0, total_rows):
								#value = sh.cell_value(i, column_in)
								sheet.write(i, k, "")
								
							wb.save(file_name)
							data = {}
							data['status_code'] = "200 OK"
							data['status'] = "Columns removed successfully"
							data = json.dumps(data)
							return data

								

						i = 0	
						x = 0 
						for col in cols_in_list:
							column_out = findexcel_column(col,file_name,sheet_name)
							column_out123=column_out
							if column_out123 != None:
								x = x + 1

							else:
								data = {}
								data['status_code'] = "401"
								data['status'] = "Please provide existing columns name"
								print(json.dumps(data))
								sys.exit()		
								
						for col in cols_in_list:	
							if x == len(cols_in_list):
								replace_col = cols_in_list[i]
								file_data_colvalues = set_col_range(file_name,col,replace_col,sheet_name)
								i = i + 1
							else:
								data = {}
								data['status_code'] = "401"
								data['status'] = "Please provide existing columns name"
								print(json.dumps(data))
								sys.exit()		
							
							
						print(file_data_colvalues)	
						
					elif file_extension in (''):
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass
						
					else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file extension"
						print(json.dumps(data))
						pass	
						
				else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass			
						
			else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide Existing file name"
						print(json.dumps(data))
						pass		
		
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e)
			print(json.dumps(data))
			pass
			
			
					
	elif commandKey == "GetCellColor":
		try :
			if os.path.exists(sys.argv[2]): 				
				filename, file_extension = os.path.splitext(sys.argv[2])			
				if filename[-1] not in (' '):
					if sys.argv[2].endswith('xls') or sys.argv[2].endswith('xlsx'):
						file_name =  sys.argv[2]	
						sheet_name =  sys.argv[3]
						temp_list1 = sys.argv[4]
						
						def getBGColor(book, sheet,temp_list1):
						
							def my_split(s):
								return filter(None, re.split(r'(\d+)', s))
			
							rowindexarray = my_split(temp_list1)	
							row1 = rowindexarray[0]
							col1 = col_to_num(row1)
							row = int(rowindexarray[1]) - 1
							col = int(col1) - 1
							xfx = sheet.cell_xf_index(row, col)
							xf = book.xf_list[xfx]
							bgx = xf.background.pattern_colour_index
							pattern_colour = book.colour_map[bgx]
							
							return pattern_colour
						
						
						lowercase_letters = [c for c in temp_list1 if c.islower()]
						if len(lowercase_letters) > 0:
							data = {}
							data['status_code'] = "401"
							data['status'] = "Please provide proper cell format"
							print(json.dumps(data))
							pass
							
						else:
							temp_list_in = temp_list1.split(',')
							cell_text_list_in = list(temp_list_in)						
							if len(cell_text_list_in) == 1:		
								book = xlrd.open_workbook(file_name, formatting_info=True)
								sheet = book.sheet_by_name(sheet_name)	
								color_name = getBGColor(book, sheet,temp_list1)
								#print(color_name)
								data = {}
								data['status_code'] = "200 OK"
								data['data'] = str(color_name)
								data = json.dumps(data)
								print(json.dumps(data))
								
							else:
								data = {}
								data['status_code'] = "401"
								data['status'] = "Please provide proper format"
								print(json.dumps(data))
								pass
						
									

					elif file_extension in (''):
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass
						
					else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file extension"
						print(json.dumps(data))
						pass	
						
				else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass			
						
			else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide Existing file name"
						print(json.dumps(data))
						pass		
						
						
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e)
			print(json.dumps(data))
			pass
					
								
	elif commandKey == "SetRangeColor":
		try :
			if os.path.exists(sys.argv[2]): 				
				filename, file_extension = os.path.splitext(sys.argv[2])			
				if filename[-1] not in (' '):
					if sys.argv[2].endswith('xls') or sys.argv[2].endswith('xlsx'):	
						file_name =  sys.argv[2]	
						sheet_name = sys.argv[3]	
						
						#cells 
						temp_list_in = sys.argv[4]
						temp_list2 = temp_list_in.split(',')
						cell_text_list = list(temp_list2)
						print(len(cell_text_list))
						
						#values
						replace_text_in = sys.argv[5]
						print(replace_text_in)
						replace_text_in1 = replace_text_in.split(' ,')
						replace_text_list = list(replace_text_in1)
						print(len(replace_text_list))
						def my_split(s):
							return filter(None, re.split(r'(\d+)', s))
							
							
						book = xlrd.open_workbook(file_name)
						sh = book.sheet_by_name(sheet_name)

						rb = xlrd.open_workbook(file_name)
						#sh = rb.sheet_by_name(sheet_name)
						wb = xl_copy(rb)
						sheet = wb.get_sheet(sheet_name)
						
						def set_range_color(cell,replace_col):	

							#original cells index
							rowindexarray = my_split(cell)
							cval = rowindexarray[0]
							rval = col_to_num(cval)
							cval1 = rowindexarray[1]
							rval = int(rval) -1
							cval1 = int(cval1) - 1

							#color adding to cells
							xlwt.add_palette_colour("custom_colour", 0x21)
							#book2 = xlwt.Workbook()
							color_params = replace_col[2:-2] 
							color_params = color_params.split(',')
							print(color_params)
							wb.set_colour_RGB(0x21,int(color_params[0]),int(color_params[1]),int(color_params[2]))

							# now you can use the colour in styles
							#sheet1 = book.add_sheet('Sheet 1')
							style = xlwt.easyxf('pattern: pattern solid, fore_colour custom_colour')
							#sheet1.write(0, 0, "",style)
							
							value = sh.cell_value(cval1,rval)
							print(cval1)
							print(rval)
							sheet.write(cval1,rval,value,style)
													
							print("saved")							
							data = {}
							data['status_code'] = "200 OK"
							data['status'] = "Cell range color updated successfully"
							data = json.dumps(data)
							return data
												
						if len(cell_text_list) == len(replace_text_list):						
							i = 0	
							for cell in cell_text_list:
								replace_col = replace_text_list[i]
								file_data_cellvalue = set_range_color(cell,replace_col)
								wb.save(file_name)	
								i = i + 1
							
							print(file_data_cellvalue)	
							
						else:
							data = {}
							data['status_code'] = "401"
							data['status'] = "Please provide proper formate"
							print(json.dumps(data))
							pass		
							
															
					elif file_extension in (''):
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass
						
					else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file extension"
						print(json.dumps(data))
						pass	
						
				else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass			
						
			else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide Existing file name"
						print(json.dumps(data))
						pass				
				
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e)
			print(json.dumps(data))
			pass	
					
	
	elif commandKey == "SetRangeColor1":
		try :
	
			'''			
			book = xlwt.Workbook()

			# add new colour to palette and set RGB colour value
			xlwt.add_palette_colour("custom_colour", 0x21)
			book.set_colour_RGB(0x21,250, 0, 0)

			# now you can use the colour in styles
			sheet1 = book.add_sheet('test123')
			style = xlwt.easyxf('pattern: pattern solid, fore_colour custom_colour')
			sheet1.write(1, 1, "",style)

			book.save('Docs/test2.xls')
			
			
			'''
			
			styles = dict(
						bold = 'font: bold 1',
						italic = 'font: italic 1',
						# Wrap text in the cell
						wrap_bold = 'font: bold 1; align: wrap 1;',
						# White text on a blue background
						reversed = 'pattern: pattern solid, fore_color blue; font: color white;',
						# Light orange checkered background
						light_orange_bg = 'pattern: pattern fine_dots, fore_color white, back_color orange;',
						# Heavy borders
						bordered = 'border: top thick, right thick, bottom thick, left thick;',
						# 16 pt red text
						big_red = 'font: height 320, color red;'
						)
			book = xlwt.Workbook()
			sheet = book.add_sheet('test1234')

			for idx, k in enumerate(sorted(styles)):
				style = xlwt.easyxf(styles[k])
				sheet.write(idx, 0, k)
				sheet.write(idx, 1, styles[k], style)

			book.save('Docs/test2.xls')
									
	
						
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e)
			print(json.dumps(data))
			pass


	elif commandKey == "SetCellValueByStartofDocument":
		try :
			if os.path.exists(sys.argv[2]): 
				filename, file_extension = os.path.splitext(sys.argv[2])			
				if filename[-1] not in (' '):
					if sys.argv[2].endswith('xls') or sys.argv[2].endswith('xlsx'):	
						file_name =  sys.argv[2]	
						sheet_name = sys.argv[3]	
						replace_text = sys.argv[4]

						def my_split(s):
							return filter(None, re.split(r'(\d+)', s))
						
						def set_cell_value(file_name,sheet_name,replace_text):							
							book = xlrd.open_workbook(file_name)
							sh = book.sheet_by_name(sheet_name)
							book1 = Workbook()
							rb = xlrd.open_workbook(file_name)
							wb = xl_copy(rb)
							sheet = wb.get_sheet(sheet_name)
							

							sheet.write(0,0,replace_text)
							
							#sheet.write(cval1,rval,replace_text)
							wb.save(file_name)							
							data = {}
							data['status_code'] = "200 OK"
							data['status'] = "Cell Updated successfully"
							data = json.dumps(data)
							return data
																
						file_data_cellvalue = set_cell_value(file_name,sheet_name,replace_text)
						print(file_data_cellvalue)	
							
							
					elif file_extension in (''):
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass
						
					else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file extension"
						print(json.dumps(data))
						pass	
						
				else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass			
						
			else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide Existing file name"
						print(json.dumps(data))
						pass				
				
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e)
			print(json.dumps(data))
			pass	
			
	
	elif commandKey == "GetCellValueByStartofDocument":
		try :
			if os.path.exists(sys.argv[2]): 
				filename, file_extension = os.path.splitext(sys.argv[2])			
				if filename[-1] not in (' '):
					if sys.argv[2].endswith('xls') or sys.argv[2].endswith('xlsx'):	
						file_name =  sys.argv[2]	
						sheet_name = sys.argv[3]	

						def my_split(s):
							return filter(None, re.split(r'(\d+)', s))
						
						def set_cell_value(file_name,sheet_name):							
							book = xlrd.open_workbook(file_name)
							sh = book.sheet_by_name(sheet_name)
							book1 = Workbook()
							rb = xlrd.open_workbook(file_name)
							wb = xl_copy(rb)
							sheet = wb.get_sheet(sheet_name)

							value = sh.cell_value(0,0)
							
							data = {}
							data['status_code'] = "200 OK"
							data['status'] = str(value)
							data = json.dumps(data)
							return data
																
						file_data_cellvalue = set_cell_value(file_name,sheet_name)
						print(file_data_cellvalue)	
							
							
					elif file_extension in (''):
						data = {}
						data['status_code'] = "401"
						data['data'] = "Please provide proper file name"
						print(json.dumps(data))
						pass
						
					else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file extension"
						print(json.dumps(data))
						pass	
						
				else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass			
						
			else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide Existing file name"
						print(json.dumps(data))
						pass				
				
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e)
			print(json.dumps(data))
			pass	


	elif commandKey == "SetCellValueByStartofColumn":
		try :
			if os.path.exists(sys.argv[2]): 				
				filename, file_extension = os.path.splitext(sys.argv[2])			
				if filename[-1] not in (' '):
					if sys.argv[2].endswith('xls') or sys.argv[2].endswith('xlsx'):
						file_name =  sys.argv[2]	
						sheet_name =  sys.argv[3]
						column_in_pre =  sys.argv[4]
						column_out_text =  sys.argv[5]
						
						column_in = col_to_num(column_in_pre) -1

						def set_col(file_name):
							book = xlrd.open_workbook(file_name)
							sh = book.sheet_by_name(sheet_name)
							book1 = Workbook()
							rb = xlrd.open_workbook(file_name)
							wb = xl_copy(rb)
							sheet = wb.get_sheet(sheet_name)
							total_rows = sh.nrows
							k = int(column_in)

							sheet.write(0, k, column_out_text)
								
							wb.save(file_name)
							data = {}
							data['status_code'] = "200 OK"
							data['status'] = "Cell updated successfully"
							data = json.dumps(data)
							return data
						

						lowercase_letters = [c for c in column_in_pre if c.islower()]
						if len(lowercase_letters) > 0:
							data = {}
							data['status_code'] = "401"
							data['status'] = "Please provide proper cell format"
							print(json.dumps(data))

						else:	
							def my_split(s):
								return filter(None, re.split(r'(\d+)', s))
								
							rowindexarray = my_split(column_in_pre)
							if len(rowindexarray) > 1:
								data = {}
								data['status_code'] = "401"
								data['status'] = "Please provide proper cell format"
								print(json.dumps(data))

							else:
								file_data_col = set_col(file_name)
								print(file_data_col)	
					
					elif file_extension in (''):
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass
						
					else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file extension"
						print(json.dumps(data))
						pass	
						
				else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass			
						
			else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide Existing file name"
						print(json.dumps(data))
						pass		
		
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e)
			print(json.dumps(data))
			pass
			

	elif commandKey == "GetCellValueByStartofColumn":
		try :
			if os.path.exists(sys.argv[2]): 				
				filename, file_extension = os.path.splitext(sys.argv[2])			
				if filename[-1] not in (' '):
					if sys.argv[2].endswith('xls') or sys.argv[2].endswith('xlsx'):
						file_name =  sys.argv[2]	
						sheet_name =  sys.argv[3]
						column_in_pre =  sys.argv[4]
						
						column_in = col_to_num(column_in_pre) -1

						def set_col(file_name):
							book = xlrd.open_workbook(file_name)
							sh = book.sheet_by_name(sheet_name)
							book1 = Workbook()
							rb = xlrd.open_workbook(file_name)
							wb = xl_copy(rb)
							sheet = wb.get_sheet(sheet_name)
							total_rows = sh.nrows
							k = int(column_in)

							value = sh.cell_value(0,k)
							
							data = {}
							data['status_code'] = "200 OK"
							data['data'] = str(value)
							data = json.dumps(data)
							return data
						

						lowercase_letters = [c for c in column_in_pre if c.islower()]
						if len(lowercase_letters) > 0:
							data = {}
							data['status_code'] = "401"
							data['status'] = "Please provide proper cell format"
							print(json.dumps(data))

						else:	
							def my_split(s):
								return filter(None, re.split(r'(\d+)', s))
								
							rowindexarray = my_split(column_in_pre)
							if len(rowindexarray) > 1:
								data = {}
								data['status_code'] = "401"
								data['status'] = "Please provide proper cell format"
								print(json.dumps(data))

							else:
								file_data_col = set_col(file_name)
								print(file_data_col)	
					
					elif file_extension in (''):
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass
						
					else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file extension"
						print(json.dumps(data))
						pass	
						
				else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass			
						
			else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide Existing file name"
						print(json.dumps(data))
						pass		
		
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e)
			print(json.dumps(data))
			pass


	elif commandKey == "SetCellValueByEndofColumn":
		try :
			if os.path.exists(sys.argv[2]): 				
				filename, file_extension = os.path.splitext(sys.argv[2])			
				if filename[-1] not in (' '):
					if sys.argv[2].endswith('xls') or sys.argv[2].endswith('xlsx'):
						file_name =  sys.argv[2]	
						sheet_name =  sys.argv[3]
						column_in_pre =  sys.argv[4]
						column_out_text =  sys.argv[5]
						
						column_in = col_to_num(column_in_pre) -1

						def set_col(file_name):
							book = xlrd.open_workbook(file_name)
							sh = book.sheet_by_name(sheet_name)
							book1 = Workbook()
							rb = xlrd.open_workbook(file_name)
							wb = xl_copy(rb)
							sheet = wb.get_sheet(sheet_name)
							total_rows = sh.nrows
							k = int(column_in)
							last_row = total_rows - 1
							sheet.write(last_row, k, column_out_text)
								
							wb.save(file_name)
							data = {}
							data['status_code'] = "200 OK"
							data['status'] = "Cell updated successfully"
							data = json.dumps(data)
							return data
						

						lowercase_letters = [c for c in column_in_pre if c.islower()]
						if len(lowercase_letters) > 0:
							data = {}
							data['status_code'] = "401"
							data['status'] = "Please provide proper cell format"
							print(json.dumps(data))

						else:	
							def my_split(s):
								return filter(None, re.split(r'(\d+)', s))
								
							rowindexarray = my_split(column_in_pre)
							if len(rowindexarray) > 1:
								data = {}
								data['status_code'] = "401"
								data['status'] = "Please provide proper cell format"
								print(json.dumps(data))

							else:
								file_data_col = set_col(file_name)
								print(file_data_col)	
					
					elif file_extension in (''):
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass
						
					else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file extension"
						print(json.dumps(data))
						pass	
						
				else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass			
						
			else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide Existing file name"
						print(json.dumps(data))
						pass		
		
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e)
			print(json.dumps(data))
			pass


	elif commandKey == "GetCellValueByEndofColumn":
		try :
			if os.path.exists(sys.argv[2]): 				
				filename, file_extension = os.path.splitext(sys.argv[2])			
				if filename[-1] not in (' '):
					if sys.argv[2].endswith('xls') or sys.argv[2].endswith('xlsx'):
						file_name =  sys.argv[2]	
						sheet_name =  sys.argv[3]
						column_in_pre =  sys.argv[4]
						
						column_in = col_to_num(column_in_pre) -1

						def set_col(file_name):
							book = xlrd.open_workbook(file_name)
							sh = book.sheet_by_name(sheet_name)
							book1 = Workbook()
							rb = xlrd.open_workbook(file_name)
							wb = xl_copy(rb)
							sheet = wb.get_sheet(sheet_name)
							total_rows = sh.nrows
							k = int(column_in)
							last_row = total_rows - 1

							value = sh.cell_value(last_row,k)
							
							data = {}
							data['status_code'] = "200 OK"
							data['data'] = str(value)
							data = json.dumps(data)
							return data
						

						lowercase_letters = [c for c in column_in_pre if c.islower()]
						if len(lowercase_letters) > 0:
							data = {}
							data['status_code'] = "401"
							data['status'] = "Please provide proper cell format"
							print(json.dumps(data))

						else:	
							def my_split(s):
								return filter(None, re.split(r'(\d+)', s))
								
							rowindexarray = my_split(column_in_pre)
							if len(rowindexarray) > 1:
								data = {}
								data['status_code'] = "401"
								data['status'] = "Please provide proper cell format"
								print(json.dumps(data))

							else:
								file_data_col = set_col(file_name)
								print(file_data_col)	
					
					elif file_extension in (''):
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass
						
					else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file extension"
						print(json.dumps(data))
						pass	
						
				else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass			
						
			else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide Existing file name"
						print(json.dumps(data))
						pass		
		
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e)
			print(json.dumps(data))
			pass


	elif commandKey == "SetCellValueByStartofRow":
		try :
			if os.path.exists(sys.argv[2]): 				
				filename, file_extension = os.path.splitext(sys.argv[2])			
				if filename[-1] not in (' '):
					if sys.argv[2].endswith('xls') or sys.argv[2].endswith('xlsx'):
						file_name =  sys.argv[2]	
						sheet_name =  sys.argv[3]
						column_in =  sys.argv[4]
						column_out =  sys.argv[5]

						def set_row(file_name):
							book = xlrd.open_workbook(file_name)
							sh = book.sheet_by_name(sheet_name)
							book1 = Workbook()
							#sheet1 = book1.add_sheet('test', cell_overwrite_ok=True)
							rb = xlrd.open_workbook(file_name)
							wb = xl_copy(rb)
							sheet = wb.get_sheet(sheet_name)
							total_cols = sh.ncols
							k = int(column_in) - 1
							#print(sys.argv[5])
							#print(column_in)

							sheet.write(k, 0, column_out)
								
							wb.save(file_name)
							data = {}
							data['status_code'] = "200 OK"
							data['status'] = "Cell updated successfully"
							data = json.dumps(data)
							return data
						

						file_data_row = set_row(file_name)
						print(file_data_row)	
								
					
					elif file_extension in (''):
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass
						
					else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file extension"
						print(json.dumps(data))
						pass	
						
				else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass			
						
			else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide Existing file name"
						print(json.dumps(data))
						pass		
		
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e)
			print(json.dumps(data))
			pass


	elif commandKey == "GetCellValueByStartofRow":
		try :
			if os.path.exists(sys.argv[2]): 				
				filename, file_extension = os.path.splitext(sys.argv[2])			
				if filename[-1] not in (' '):
					if sys.argv[2].endswith('xls') or sys.argv[2].endswith('xlsx'):
						file_name =  sys.argv[2]	
						sheet_name =  sys.argv[3]
						column_in =  sys.argv[4]

						def set_row(file_name):
							book = xlrd.open_workbook(file_name)
							sh = book.sheet_by_name(sheet_name)
							book1 = Workbook()
							#sheet1 = book1.add_sheet('test', cell_overwrite_ok=True)
							rb = xlrd.open_workbook(file_name)
							wb = xl_copy(rb)
							sheet = wb.get_sheet(sheet_name)
							total_cols = sh.ncols
							k = int(column_in) - 1

							value = sh.cell_value(k,0)
								
							wb.save(file_name)
							data = {}
							data['status_code'] = "200 OK"
							data['data'] = str(value)
							data = json.dumps(data)
							return data
						

						file_data_row = set_row(file_name)
						print(file_data_row)	
								
					
					elif file_extension in (''):
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass
						
					else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file extension"
						print(json.dumps(data))
						pass	
						
				else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass			
						
			else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide Existing file name"
						print(json.dumps(data))
						pass		
		
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e)
			print(json.dumps(data))
			pass


	elif commandKey == "SetCellValueByEndofRow":
		try :
			if os.path.exists(sys.argv[2]): 				
				filename, file_extension = os.path.splitext(sys.argv[2])			
				if filename[-1] not in (' '):
					if sys.argv[2].endswith('xls') or sys.argv[2].endswith('xlsx'):
						file_name =  sys.argv[2]	
						sheet_name =  sys.argv[3]
						column_in =  sys.argv[4]
						column_out =  sys.argv[5]

						def set_row(file_name):
							book = xlrd.open_workbook(file_name)
							sh = book.sheet_by_name(sheet_name)
							book1 = Workbook()
							#sheet1 = book1.add_sheet('test', cell_overwrite_ok=True)
							rb = xlrd.open_workbook(file_name)
							wb = xl_copy(rb)
							sheet = wb.get_sheet(sheet_name)
							total_cols = sh.ncols
							k = int(column_in) - 1
							#print(sys.argv[5])
							#print(column_in)
							last_column = total_cols - 1

							sheet.write(k, last_column, column_out)
								
							wb.save(file_name)
							data = {}
							data['status_code'] = "200 OK"
							data['status'] = "Cell updated successfully"
							data = json.dumps(data)
							return data
						

						file_data_row = set_row(file_name)
						print(file_data_row)	
								
					
					elif file_extension in (''):
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass
						
					else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file extension"
						print(json.dumps(data))
						pass	
						
				else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass			
						
			else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide Existing file name"
						print(json.dumps(data))
						pass		
		
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e)
			print(json.dumps(data))
			pass


	elif commandKey == "GetCellValueByEndofRow":
		try :
			if os.path.exists(sys.argv[2]): 				
				filename, file_extension = os.path.splitext(sys.argv[2])			
				if filename[-1] not in (' '):
					if sys.argv[2].endswith('xls') or sys.argv[2].endswith('xlsx'):
						file_name =  sys.argv[2]	
						sheet_name =  sys.argv[3]
						column_in =  sys.argv[4]

						def set_row(file_name):
							book = xlrd.open_workbook(file_name)
							sh = book.sheet_by_name(sheet_name)
							book1 = Workbook()
							#sheet1 = book1.add_sheet('test', cell_overwrite_ok=True)
							rb = xlrd.open_workbook(file_name)
							wb = xl_copy(rb)
							sheet = wb.get_sheet(sheet_name)
							total_cols = sh.ncols
							k = int(column_in) - 1
							last_column = total_cols - 1

							value = sh.cell_value(k,last_column)
								
							wb.save(file_name)
							data = {}
							data['status_code'] = "200 OK"
							data['data'] = str(value)
							data = json.dumps(data)
							return data
						

						file_data_row = set_row(file_name)
						print(file_data_row)	
								
					
					elif file_extension in (''):
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass
						
					else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file extension"
						print(json.dumps(data))
						pass	
						
				else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass			
						
			else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide Existing file name"
						print(json.dumps(data))
						pass		
		
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e)
			print(json.dumps(data))
			pass


	elif commandKey == "SetCurrentCellValue":
		try :
			if os.path.exists(sys.argv[2]): 
				filename, file_extension = os.path.splitext(sys.argv[2])			
				if filename[-1] not in (' '):
					if sys.argv[2].endswith('xls') or sys.argv[2].endswith('xlsx'):	
						file_name =  sys.argv[2]	
						sheet_name = sys.argv[3]	
						temp_list1 = sys.argv[4]
						replace_text = sys.argv[5]

						def my_split(s):
							return filter(None, re.split(r'(\d+)', s))
						
						def set_cell_value(file_name,sheet_name,temp_list1,replace_text):							
							book = xlrd.open_workbook(file_name)
							sh = book.sheet_by_name(sheet_name)
							book1 = Workbook()
							rb = xlrd.open_workbook(file_name)
							wb = xl_copy(rb)
							sheet = wb.get_sheet(sheet_name)
							
							#original cell
							rowindexarray = my_split(temp_list1)
							cval = rowindexarray[0]
							rval = col_to_num(cval)
							cval1 = rowindexarray[1]
							rval = int(rval) -1
							cval1 = int(cval1) - 1
														
							#value = sh.cell_value(cval1, rval)
							sheet.write(cval1,rval,replace_text)
							
							#sheet.write(cval1,rval,replace_text)
							wb.save(file_name)							
							data = {}
							data['status_code'] = "200 OK"
							data['status'] = "Cell Updated successfully"
							data = json.dumps(data)
							return data
												
						temp_list1_prevent = re.findall('[^A-Z0-9]',temp_list1)
						if len(temp_list1_prevent) < 1:
							if len(temp_list1) > 2 :
								if temp_list1[1].isdigit():
									if temp_list1[2:].isdigit():
										lowercase_letters = [c for c in temp_list1 if c.islower()]
										if len(lowercase_letters) > 0:
											data = {}
											data['status_code'] = "401"
											data['status'] = "Please provide proper cell format"
											print(json.dumps(data))

										else:	
											file_data_cellvalue = set_cell_value(file_name,sheet_name,temp_list1,replace_text)
											print(file_data_cellvalue)	
									else:
										data = {}
										data['status_code'] = "401"
										data['status'] = "Please provide proper cell index"
										print(json.dumps(data))
										pass
									
								else:
									lowercase_letters = [c for c in temp_list1 if c.islower()]
									if len(lowercase_letters) > 0:
										data = {}
										data['status_code'] = "401"
										data['status'] = "Please provide proper cell format"
										print(json.dumps(data))

									else:	
										file_data_cellvalue = set_cell_value(file_name,sheet_name,temp_list1,replace_text)
										print(file_data_cellvalue)	
						
							else:
								lowercase_letters = [c for c in temp_list1 if c.islower()]
								if len(lowercase_letters) > 0:
									data = {}
									data['status_code'] = "401"
									data['status'] = "Please provide proper cell format"
									print(json.dumps(data))

								else:	
									file_data_cellvalue = set_cell_value(file_name,sheet_name,temp_list1,replace_text)
									print(file_data_cellvalue)	
						
							
						else:
							data = {}
							data['status_code'] = "401"
							data['status'] = "Please provide proper cell index"
							print(json.dumps(data))
							pass
							
													
					elif file_extension in (''):
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass
						
					else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file extension"
						print(json.dumps(data))
						pass	
						
				else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass			
						
			else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide Existing file name"
						print(json.dumps(data))
						pass				
				
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e)
			print(json.dumps(data))
			pass	


	elif commandKey == "GetCurrentCellValue":
		try :
			if os.path.exists(sys.argv[2]): 
				filename, file_extension = os.path.splitext(sys.argv[2])			
				if filename[-1] not in (' '):
					if sys.argv[2].endswith('xls') or sys.argv[2].endswith('xlsx'):	
						file_name =  sys.argv[2]	
						sheet_name = sys.argv[3]	
						temp_list1 = sys.argv[4]
						
						def my_split(s):
							return filter(None, re.split(r'(\d+)', s))
						
						def set_cell_value(file_name,sheet_name,temp_list1):							
							book = xlrd.open_workbook(file_name)
							sh = book.sheet_by_name(sheet_name)
							book1 = Workbook()
							rb = xlrd.open_workbook(file_name)
							wb = xl_copy(rb)
							sheet = wb.get_sheet(sheet_name)
							
							#original cell
							rowindexarray = my_split(temp_list1)
							cval = rowindexarray[0]
							rval = col_to_num(cval)
							cval1 = rowindexarray[1]
							rval = int(rval) -1
							cval1 = int(cval1) - 1
														
							value = sh.cell_value(cval1, rval)
						
							data = {}
							data['status_code'] = "200 OK"
							data['data'] = str(value)
							data = json.dumps(data)
							return data
												
						temp_list1_prevent = re.findall('[^A-Z0-9]',temp_list1)
						if len(temp_list1_prevent) < 1:
							if len(temp_list1) > 2 :
								if temp_list1[1].isdigit():
									if temp_list1[2:].isdigit():
										lowercase_letters = [c for c in temp_list1 if c.islower()]
										if len(lowercase_letters) > 0:
											data = {}
											data['status_code'] = "401"
											data['status'] = "Please provide proper cell format"
											print(json.dumps(data))

										else:	
											file_data_cellvalue = set_cell_value(file_name,sheet_name,temp_list1)
											print(file_data_cellvalue)		
									else:
										data = {}
										data['status_code'] = "401"
										data['status'] = "Please provide proper cell index"
										print(json.dumps(data))
										pass
								
								else:
									lowercase_letters = [c for c in temp_list1 if c.islower()]
									if len(lowercase_letters) > 0:
										data = {}
										data['status_code'] = "401"
										data['status'] = "Please provide proper cell format"
										print(json.dumps(data))

									else:	
										file_data_cellvalue = set_cell_value(file_name,sheet_name,temp_list1)
										print(file_data_cellvalue)	
						
							else:
								lowercase_letters = [c for c in temp_list1 if c.islower()]
								if len(lowercase_letters) > 0:
									data = {}
									data['status_code'] = "401"
									data['status'] = "Please provide proper cell format"
									print(json.dumps(data))

								else:	
									file_data_cellvalue = set_cell_value(file_name,sheet_name,temp_list1)
									print(file_data_cellvalue)	
						
							
						else:
							data = {}
							data['status_code'] = "401"
							data['status'] = "Please provide proper cell index"
							print(json.dumps(data))
							pass
													
					elif file_extension in (''):
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass
						
					else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file extension"
						print(json.dumps(data))
						pass	
						
				else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass			
						
			else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide Existing file name"
						print(json.dumps(data))
						pass				
				
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e)
			print(json.dumps(data))
			pass	


	elif commandKey == "SetCellValueAbove":
		try :
			if os.path.exists(sys.argv[2]): 
				filename, file_extension = os.path.splitext(sys.argv[2])			
				if filename[-1] not in (' '):
					if sys.argv[2].endswith('xls') or sys.argv[2].endswith('xlsx'):	
						file_name =  sys.argv[2]	
						sheet_name = sys.argv[3]	
						temp_list1 = sys.argv[4]
						replace_text = sys.argv[5]
						def my_split(s):
							return filter(None, re.split(r'(\d+)', s))
						
						def set_cell_value(file_name,sheet_name,temp_list1,replace_text):							
							book = xlrd.open_workbook(file_name)
							sh = book.sheet_by_name(sheet_name)
							book1 = Workbook()
							rb = xlrd.open_workbook(file_name)
							wb = xl_copy(rb)
							sheet = wb.get_sheet(sheet_name)
							
							#original cell
							rowindexarray = my_split(temp_list1)
							cval = rowindexarray[0]
							rval = col_to_num(cval)
							cval1 = rowindexarray[1]
							rval = int(rval) -1
							cval1 = int(cval1) - 1
							cval_above = cval1 - 1 		
							
							if cval_above >= 0:
								#value = sh.cell_value(cval1, rval)
								sheet.write(cval_above,rval,replace_text)
								
								#sheet.write(cval1,rval,replace_text)
								wb.save(file_name)							
								data = {}
								data['status_code'] = "200 OK"
								data['status'] = "Cell Updated successfully"
								data = json.dumps(data)
								return data

							else:
								wb.save(file_name)							
								data = {}
								data['status_code'] = "401"
								data['status'] = "There is no cell above the current cell"
								data = json.dumps(data)
								return data
												
						temp_list1_prevent = re.findall('[^A-Z0-9]',temp_list1)
						if len(temp_list1_prevent) < 1:
							if len(temp_list1) > 2 :
								if temp_list1[1].isdigit():
									if temp_list1[2:].isdigit():
										lowercase_letters = [c for c in temp_list1 if c.islower()]
										if len(lowercase_letters) > 0:
											data = {}
											data['status_code'] = "401"
											data['status'] = "Please provide proper cell format"
											print(json.dumps(data))

										else:	
											file_data_cellvalue = set_cell_value(file_name,sheet_name,temp_list1,replace_text)
											print(file_data_cellvalue)
											
									else:
										data = {}
										data['status_code'] = "401"
										data['status'] = "Please provide proper cell index"
										print(json.dumps(data))
										pass
								
								else:
									lowercase_letters = [c for c in temp_list1 if c.islower()]
									if len(lowercase_letters) > 0:
										data = {}
										data['status_code'] = "401"
										data['status'] = "Please provide proper cell format"
										print(json.dumps(data))

									else:	
										file_data_cellvalue = set_cell_value(file_name,sheet_name,temp_list1,replace_text)
										print(file_data_cellvalue)	
						
							else:
								lowercase_letters = [c for c in temp_list1 if c.islower()]
								if len(lowercase_letters) > 0:
									data = {}
									data['status_code'] = "401"
									data['status'] = "Please provide proper cell format"
									print(json.dumps(data))

								else:	
									file_data_cellvalue = set_cell_value(file_name,sheet_name,temp_list1,replace_text)
									print(file_data_cellvalue)		
						
							
						else:
							data = {}
							data['status_code'] = "401"
							data['status'] = "Please provide proper cell index"
							print(json.dumps(data))
							pass
							
													
					elif file_extension in (''):
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass
						
					else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file extension"
						print(json.dumps(data))
						pass	
						
				else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass			
						
			else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide Existing file name"
						print(json.dumps(data))
						pass				
				
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e)
			print(json.dumps(data))
			pass


	elif commandKey == "GetCellValueAbove":
		try :
			if os.path.exists(sys.argv[2]): 
				filename, file_extension = os.path.splitext(sys.argv[2])			
				if filename[-1] not in (' '):
					if sys.argv[2].endswith('xls') or sys.argv[2].endswith('xlsx'):	
						file_name =  sys.argv[2]	
						sheet_name = sys.argv[3]	
						temp_list1 = sys.argv[4]

						def my_split(s):
							return filter(None, re.split(r'(\d+)', s))
						
						def set_cell_value(file_name,sheet_name,temp_list1):							
							book = xlrd.open_workbook(file_name)
							sh = book.sheet_by_name(sheet_name)
							book1 = Workbook()
							rb = xlrd.open_workbook(file_name)
							wb = xl_copy(rb)
							sheet = wb.get_sheet(sheet_name)
							
							#original cell
							rowindexarray = my_split(temp_list1)
							cval = rowindexarray[0]
							rval = col_to_num(cval)
							cval1 = rowindexarray[1]
							rval = int(rval) -1
							cval1 = int(cval1) - 1
							cval_above = cval1 - 1 		
							
							if cval_above >= 0:
								value = sh.cell_value(cval_above, rval)
							
								data = {}
								data['status_code'] = "200 OK"
								data['data'] = str(value)
								data = json.dumps(data)
								return data

							else:
								#wb.save(file_name)							
								data = {}
								data['status_code'] = "401"
								data['status'] = "There is no cell above the current cell"
								data = json.dumps(data)
								return data
												
						temp_list1_prevent = re.findall('[^A-Z0-9]',temp_list1)
						if len(temp_list1_prevent) < 1:
							if len(temp_list1) > 2 :
								if temp_list1[1].isdigit():
									if temp_list1[2:].isdigit():
										lowercase_letters = [c for c in temp_list1 if c.islower()]
										if len(lowercase_letters) > 0:
											data = {}
											data['status_code'] = "401"
											data['status'] = "Please provide proper cell format"
											print(json.dumps(data))

										else:	
											file_data_cellvalue = set_cell_value(file_name,sheet_name,temp_list1)
											print(file_data_cellvalue)	
											
									else:
										data = {}
										data['status_code'] = "401"
										data['status'] = "Please provide proper cell index"
										print(json.dumps(data))
										pass
								
								else:
									lowercase_letters = [c for c in temp_list1 if c.islower()]
									if len(lowercase_letters) > 0:
										data = {}
										data['status_code'] = "401"
										data['status'] = "Please provide proper cell format"
										print(json.dumps(data))

									else:	
										file_data_cellvalue = set_cell_value(file_name,sheet_name,temp_list1)
										print(file_data_cellvalue)	
						
						
							else:
								lowercase_letters = [c for c in temp_list1 if c.islower()]
								if len(lowercase_letters) > 0:
									data = {}
									data['status_code'] = "401"
									data['status'] = "Please provide proper cell format"
									print(json.dumps(data))

								else:	
									file_data_cellvalue = set_cell_value(file_name,sheet_name,temp_list1)
									print(file_data_cellvalue)	
						
							
						else:
							data = {}
							data['status_code'] = "401"
							data['status'] = "Please provide proper cell index"
							print(json.dumps(data))
							pass

													
					elif file_extension in (''):
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass
						
					else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file extension"
						print(json.dumps(data))
						pass	
						
				else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass			
						
			else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide Existing file name"
						print(json.dumps(data))
						pass				
				
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e)
			print(json.dumps(data))
			pass


	elif commandKey == "SetCellValueBelow":
		try :
			if os.path.exists(sys.argv[2]): 
				filename, file_extension = os.path.splitext(sys.argv[2])			
				if filename[-1] not in (' '):
					if sys.argv[2].endswith('xls') or sys.argv[2].endswith('xlsx'):	
						file_name =  sys.argv[2]	
						sheet_name = sys.argv[3]	
						temp_list1 = sys.argv[4]
						replace_text = sys.argv[5]
						def my_split(s):
							return filter(None, re.split(r'(\d+)', s))
						
						def set_cell_value(file_name,sheet_name,temp_list1,replace_text):							
							book = xlrd.open_workbook(file_name)
							sh = book.sheet_by_name(sheet_name)
							book1 = Workbook()
							rb = xlrd.open_workbook(file_name)
							wb = xl_copy(rb)
							sheet = wb.get_sheet(sheet_name)
							
							#original cell
							rowindexarray = my_split(temp_list1)
							cval = rowindexarray[0]
							rval = col_to_num(cval)
							cval1 = rowindexarray[1]
							rval = int(rval) -1
							cval1 = int(cval1) - 1
							cval_below = cval1 + 1 		
							if cval_below >= 1:
								#value = sh.cell_value(cval1, rval)
								sheet.write(cval_below,rval,replace_text)
								
								#sheet.write(cval1,rval,replace_text)
								wb.save(file_name)							
								data = {}
								data['status_code'] = "200 OK"
								data['status'] = "Cell Updated successfully"
								data = json.dumps(data)
								return data

							else:
								wb.save(file_name)							
								data = {}
								data['status_code'] = "401"
								data['status'] = "There is no cell below the current cell"
								data = json.dumps(data)
								return data
												

												
						temp_list1_prevent = re.findall('[^A-Z0-9]',temp_list1)
						if len(temp_list1_prevent) < 1:
							if len(temp_list1) > 2 :
								if temp_list1[1].isdigit():
									if temp_list1[2:].isdigit():
										lowercase_letters = [c for c in temp_list1 if c.islower()]
										if len(lowercase_letters) > 0:
											data = {}
											data['status_code'] = "401"
											data['status'] = "Please provide proper cell format"
											print(json.dumps(data))

										else:	
											file_data_cellvalue = set_cell_value(file_name,sheet_name,temp_list1,replace_text)
											print(file_data_cellvalue)
											
									else:
										data = {}
										data['status_code'] = "401"
										data['status'] = "Please provide proper cell index"
										print(json.dumps(data))
										pass
								
								else:						
									lowercase_letters = [c for c in temp_list1 if c.islower()]
									if len(lowercase_letters) > 0:
										data = {}
										data['status_code'] = "401"
										data['status'] = "Please provide proper cell format"
										print(json.dumps(data))

									else:	
										file_data_cellvalue = set_cell_value(file_name,sheet_name,temp_list1,replace_text)
										print(file_data_cellvalue)	
										
						
							else:
								lowercase_letters = [c for c in temp_list1 if c.islower()]
								if len(lowercase_letters) > 0:
									data = {}
									data['status_code'] = "401"
									data['status'] = "Please provide proper cell format"
									print(json.dumps(data))

								else:	
									file_data_cellvalue = set_cell_value(file_name,sheet_name,temp_list1,replace_text)
									print(file_data_cellvalue)	
						
							
						else:
							data = {}
							data['status_code'] = "401"
							data['status'] = "Please provide proper cell index"
							print(json.dumps(data))
							pass
						
						
													
					elif file_extension in (''):
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass
						
					else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file extension"
						print(json.dumps(data))
						pass	
						
				else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass			
						
			else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide Existing file name"
						print(json.dumps(data))
						pass				
				
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e)
			print(json.dumps(data))
			pass


	elif commandKey == "GetCellValueBelow":
		try :
			if os.path.exists(sys.argv[2]): 
				filename, file_extension = os.path.splitext(sys.argv[2])			
				if filename[-1] not in (' '):
					if sys.argv[2].endswith('xls') or sys.argv[2].endswith('xlsx'):	
						file_name =  sys.argv[2]	
						sheet_name = sys.argv[3]	
						temp_list1 = sys.argv[4]

						def my_split(s):
							return filter(None, re.split(r'(\d+)', s))
						
						def set_cell_value(file_name,sheet_name,temp_list1):							
							book = xlrd.open_workbook(file_name)
							sh = book.sheet_by_name(sheet_name)
							book1 = Workbook()
							rb = xlrd.open_workbook(file_name)
							wb = xl_copy(rb)
							sheet = wb.get_sheet(sheet_name)
							
							#original cell
							rowindexarray = my_split(temp_list1)
							cval = rowindexarray[0]
							rval = col_to_num(cval)
							cval1 = rowindexarray[1]
							rval = int(rval) -1
							cval1 = int(cval1) - 1
							cval_below = cval1 + 1 		
							if cval_below >= 1:
								value = sh.cell_value(cval_below, rval)
								#sheet.write(cval_below,rval,replace_text)
															
								data = {}
								data['status_code'] = "200 OK"
								data['data'] = str(value)
								data = json.dumps(data)
								return data

							else:
								wb.save(file_name)							
								data = {}
								data['status_code'] = "401"
								data['status'] = "There is no cell below the current cell"
								data = json.dumps(data)
								return data
												
						temp_list1_prevent = re.findall('[^A-Z0-9]',temp_list1)
						if len(temp_list1_prevent) < 1:
							if len(temp_list1) > 2 :
								if temp_list1[1].isdigit():
									if temp_list1[2:].isdigit():
										lowercase_letters = [c for c in temp_list1 if c.islower()]
										if len(lowercase_letters) > 0:
											data = {}
											data['status_code'] = "401"
											data['status'] = "Please provide proper cell format"
											print(json.dumps(data))

										else:	
											file_data_cellvalue = set_cell_value(file_name,sheet_name,temp_list1)
											print(file_data_cellvalue)
											
									else:
										data = {}
										data['status_code'] = "401"
										data['status'] = "Please provide proper cell index"
										print(json.dumps(data))
										pass
								
								else:	
									lowercase_letters = [c for c in temp_list1 if c.islower()]
									if len(lowercase_letters) > 0:
										data = {}
										data['status_code'] = "401"
										data['status'] = "Please provide proper cell format"
										print(json.dumps(data))

									else:	
										file_data_cellvalue = set_cell_value(file_name,sheet_name,temp_list1)
										print(file_data_cellvalue)	
						
							else:
								lowercase_letters = [c for c in temp_list1 if c.islower()]
								if len(lowercase_letters) > 0:
									data = {}
									data['status_code'] = "401"
									data['status'] = "Please provide proper cell format"
									print(json.dumps(data))

								else:	
									file_data_cellvalue = set_cell_value(file_name,sheet_name,temp_list1)
									print(file_data_cellvalue)	
						
							
						else:
							data = {}
							data['status_code'] = "401"
							data['status'] = "Please provide proper cell index"
							print(json.dumps(data))
							pass
							
													
					elif file_extension in (''):
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass
						
					else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file extension"
						print(json.dumps(data))
						pass	
						
				else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass			
						
			else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide Existing file name"
						print(json.dumps(data))
						pass				
				
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e)
			print(json.dumps(data))
			pass


	elif commandKey == "SetCellValueToLeft":
		try :
			if os.path.exists(sys.argv[2]): 
				filename, file_extension = os.path.splitext(sys.argv[2])			
				if filename[-1] not in (' '):
					if sys.argv[2].endswith('xls') or sys.argv[2].endswith('xlsx'):	
						file_name =  sys.argv[2]	
						sheet_name = sys.argv[3]	
						temp_list1 = sys.argv[4]
						replace_text = sys.argv[5]
						def my_split(s):
							return filter(None, re.split(r'(\d+)', s))
						
						def set_cell_value(file_name,sheet_name,temp_list1,replace_text):							
							book = xlrd.open_workbook(file_name)
							sh = book.sheet_by_name(sheet_name)
							book1 = Workbook()
							rb = xlrd.open_workbook(file_name)
							wb = xl_copy(rb)
							sheet = wb.get_sheet(sheet_name)
							
							#original cell
							rowindexarray = my_split(temp_list1)
							cval = rowindexarray[0]
							rval = col_to_num(cval)
							cval1 = rowindexarray[1]
							rval = int(rval) -1
							cval1 = int(cval1) - 1
							cell_left = rval - 1
							
							if cval1 >= 0 and cell_left >= 0:
								#value = sh.cell_value(cval1, rval)
								sheet.write(cval1,cell_left,replace_text)
								
								#sheet.write(cval1,rval,replace_text)
								wb.save(file_name)							
								data = {}
								data['status_code'] = "200 OK"
								data['status'] = "Cell Updated successfully"
								data = json.dumps(data)
								return data

							else:
								wb.save(file_name)							
								data = {}
								data['status_code'] = "401"
								data['status'] = "There is no cell left to the current cell"
								data = json.dumps(data)
								return data
												

						temp_list1_prevent = re.findall('[^A-Z0-9]',temp_list1)
						if len(temp_list1_prevent) < 1:
							if len(temp_list1) > 2 :
								if temp_list1[1].isdigit():
									if temp_list1[2:].isdigit():
										lowercase_letters = [c for c in temp_list1 if c.islower()]
										if len(lowercase_letters) > 0:
											data = {}
											data['status_code'] = "401"
											data['status'] = "Please provide proper cell format"
											print(json.dumps(data))

										else:	
											file_data_cellvalue = set_cell_value(file_name,sheet_name,temp_list1,replace_text)
											print(file_data_cellvalue)	
											
									else:
										data = {}
										data['status_code'] = "401"
										data['status'] = "Please provide proper cell index"
										print(json.dumps(data))
										pass
								
								else:	
									lowercase_letters = [c for c in temp_list1 if c.islower()]
									if len(lowercase_letters) > 0:
										data = {}
										data['status_code'] = "401"
										data['status'] = "Please provide proper cell format"
										print(json.dumps(data))

									else:	
										file_data_cellvalue = set_cell_value(file_name,sheet_name,temp_list1,replace_text)
										print(file_data_cellvalue)	
						

							else:
								lowercase_letters = [c for c in temp_list1 if c.islower()]
								if len(lowercase_letters) > 0:
									data = {}
									data['status_code'] = "401"
									data['status'] = "Please provide proper cell format"
									print(json.dumps(data))

								else:	
									file_data_cellvalue = set_cell_value(file_name,sheet_name,temp_list1,replace_text)
									print(file_data_cellvalue)
						
							
						else:
							data = {}
							data['status_code'] = "401"
							data['status'] = "Please provide proper cell index"
							print(json.dumps(data))
							pass
						
													
					elif file_extension in (''):
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass
						
					else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file extension"
						print(json.dumps(data))
						pass	
						
				else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass			
						
			else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide Existing file name"
						print(json.dumps(data))
						pass				
				
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e)
			print(json.dumps(data))
			pass


	elif commandKey == "GetCellValueToLeft":
		try :
			if os.path.exists(sys.argv[2]): 
				filename, file_extension = os.path.splitext(sys.argv[2])			
				if filename[-1] not in (' '):
					if sys.argv[2].endswith('xls') or sys.argv[2].endswith('xlsx'):	
						file_name =  sys.argv[2]	
						sheet_name = sys.argv[3]	
						temp_list1 = sys.argv[4]

						def my_split(s):
							return filter(None, re.split(r'(\d+)', s))
						
						def set_cell_value(file_name,sheet_name,temp_list1):							
							book = xlrd.open_workbook(file_name)
							sh = book.sheet_by_name(sheet_name)
							book1 = Workbook()
							rb = xlrd.open_workbook(file_name)
							wb = xl_copy(rb)
							sheet = wb.get_sheet(sheet_name)
							
							#original cell
							rowindexarray = my_split(temp_list1)
							cval = rowindexarray[0]
							rval = col_to_num(cval)
							cval1 = rowindexarray[1]
							rval = int(rval) -1
							cval1 = int(cval1) - 1
							cell_left = rval - 1
							
							if cval1 >= 0 and rval > 0:
								value = sh.cell_value(cval1, cell_left)
							
								data = {}
								data['status_code'] = "200 OK"
								data['data'] = str(value)
								data = json.dumps(data)
								return data

							else:
								wb.save(file_name)							
								data = {}
								data['status_code'] = "401"
								data['status'] = "There is no cell left to the current cell"
								data = json.dumps(data)
								return data
												
						temp_list1_prevent = re.findall('[^A-Z0-9]',temp_list1)
						if len(temp_list1_prevent) < 1:
							if len(temp_list1) > 2 :
								if temp_list1[1].isdigit():
									if temp_list1[2:].isdigit():
										lowercase_letters = [c for c in temp_list1 if c.islower()]
										if len(lowercase_letters) > 0:
											data = {}
											data['status_code'] = "401"
											data['status'] = "Please provide proper cell format"
											print(json.dumps(data))

										else:	
											file_data_cellvalue = set_cell_value(file_name,sheet_name,temp_list1)
											print(file_data_cellvalue)
											
									else:
										data = {}
										data['status_code'] = "401"
										data['status'] = "Please provide proper cell index"
										print(json.dumps(data))
										pass
								
								else:	
									lowercase_letters = [c for c in temp_list1 if c.islower()]
									if len(lowercase_letters) > 0:
										data = {}
										data['status_code'] = "401"
										data['status'] = "Please provide proper cell format"
										print(json.dumps(data))

									else:	
										file_data_cellvalue = set_cell_value(file_name,sheet_name,temp_list1)
										print(file_data_cellvalue)	
						

							else:
								lowercase_letters = [c for c in temp_list1 if c.islower()]
								if len(lowercase_letters) > 0:
									data = {}
									data['status_code'] = "401"
									data['status'] = "Please provide proper cell format"
									print(json.dumps(data))

								else:	
									file_data_cellvalue = set_cell_value(file_name,sheet_name,temp_list1)
									print(file_data_cellvalue)	
							
						else:
							data = {}
							data['status_code'] = "401"
							data['status'] = "Please provide proper cell index"
							print(json.dumps(data))
							pass
						
													
					elif file_extension in (''):
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass
						
					else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file extension"
						print(json.dumps(data))
						pass	
						
				else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass			
						
			else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide Existing file name"
						print(json.dumps(data))
						pass				
				
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e)
			print(json.dumps(data))
			pass


	elif commandKey == "SetCellValueToRight":
		try :
			if os.path.exists(sys.argv[2]): 
				filename, file_extension = os.path.splitext(sys.argv[2])			
				if filename[-1] not in (' '):
					if sys.argv[2].endswith('xls') or sys.argv[2].endswith('xlsx'):	
						file_name =  sys.argv[2]	
						sheet_name = sys.argv[3]	
						temp_list1 = sys.argv[4]
						replace_text = sys.argv[5]
						def my_split(s):
							return filter(None, re.split(r'(\d+)', s))
						
						def set_cell_value(file_name,sheet_name,temp_list1,replace_text):							
							book = xlrd.open_workbook(file_name)
							sh = book.sheet_by_name(sheet_name)
							book1 = Workbook()
							rb = xlrd.open_workbook(file_name)
							wb = xl_copy(rb)
							sheet = wb.get_sheet(sheet_name)
							
							#original cell
							rowindexarray = my_split(temp_list1)
							cval = rowindexarray[0]
							rval = col_to_num(cval)
							cval1 = rowindexarray[1]
							rval = int(rval) -1
							cval1 = int(cval1) - 1
							cell_left = rval + 1
							
							if cell_left >= 0:
								#value = sh.cell_value(cval1, rval)
								sheet.write(cval1,cell_left,replace_text)
								
								#sheet.write(cval1,rval,replace_text)
								wb.save(file_name)							
								data = {}
								data['status_code'] = "200 OK"
								data['status'] = "Cell Updated successfully"
								data = json.dumps(data)
								return data

							else:
								wb.save(file_name)							
								data = {}
								data['status_code'] = "401"
								data['status'] = "There is no cell right to the current cell"
								data = json.dumps(data)
								return data
												
						temp_list1_prevent = re.findall('[^A-Z0-9]',temp_list1)
						if len(temp_list1_prevent) < 1:
							if len(temp_list1) > 2 :
								if temp_list1[1].isdigit():
									if temp_list1[2:].isdigit():
										lowercase_letters = [c for c in temp_list1 if c.islower()]
										if len(lowercase_letters) > 0:
											data = {}
											data['status_code'] = "401"
											data['status'] = "Please provide proper cell format"
											print(json.dumps(data))

										else:	
											file_data_cellvalue = set_cell_value(file_name,sheet_name,temp_list1,replace_text)
											print(file_data_cellvalue)	
											
									else:
										data = {}
										data['status_code'] = "401"
										data['status'] = "Please provide proper cell index"
										print(json.dumps(data))
										pass
								
								else:	
									lowercase_letters = [c for c in temp_list1 if c.islower()]
									if len(lowercase_letters) > 0:
										data = {}
										data['status_code'] = "401"
										data['status'] = "Please provide proper cell format"
										print(json.dumps(data))

									else:	
										file_data_cellvalue = set_cell_value(file_name,sheet_name,temp_list1,replace_text)
										print(file_data_cellvalue)	
						
							else:
								lowercase_letters = [c for c in temp_list1 if c.islower()]
								if len(lowercase_letters) > 0:
									data = {}
									data['status_code'] = "401"
									data['status'] = "Please provide proper cell format"
									print(json.dumps(data))

								else:	
									file_data_cellvalue = set_cell_value(file_name,sheet_name,temp_list1,replace_text)
									print(file_data_cellvalue)	
							
						else:
							data = {}
							data['status_code'] = "401"
							data['status'] = "Please provide proper cell index"
							print(json.dumps(data))
							pass
							
													
					elif file_extension in (''):
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass
						
					else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file extension"
						print(json.dumps(data))
						pass	
						
				else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass			
						
			else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide Existing file name"
						print(json.dumps(data))
						pass				
				
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e)
			print(json.dumps(data))
			pass


	elif commandKey == "GetCellValueToRight":
		try :
			if os.path.exists(sys.argv[2]): 
				filename, file_extension = os.path.splitext(sys.argv[2])			
				if filename[-1] not in (' '):
					if sys.argv[2].endswith('xls') or sys.argv[2].endswith('xlsx'):	
						file_name =  sys.argv[2]	
						sheet_name = sys.argv[3]	
						temp_list1 = sys.argv[4]

						def my_split(s):
							return filter(None, re.split(r'(\d+)', s))
						
						def set_cell_value(file_name,sheet_name,temp_list1):							
							book = xlrd.open_workbook(file_name)
							sh = book.sheet_by_name(sheet_name)
							book1 = Workbook()
							rb = xlrd.open_workbook(file_name)
							wb = xl_copy(rb)
							sheet = wb.get_sheet(sheet_name)
							
							#original cell
							rowindexarray = my_split(temp_list1)
							cval = rowindexarray[0]
							rval = col_to_num(cval)
							cval1 = rowindexarray[1]
							rval = int(rval) -1
							cval1 = int(cval1) - 1
							cell_left = rval + 1

							if cval1 >= 0:
								value = sh.cell_value(cval1, cell_left)
							
								data = {}
								data['status_code'] = "200 OK"
								data['data'] = str(value)
								data = json.dumps(data)
								return data

							else:
								wb.save(file_name)							
								data = {}
								data['status_code'] = "401"
								data['status'] = "There is no cell right to the current cell"
								data = json.dumps(data)
								return data
												
						temp_list1_prevent = re.findall('[^A-Z0-9]',temp_list1)
						if len(temp_list1_prevent) < 1:
							if len(temp_list1) > 2 :
								if temp_list1[1].isdigit():
									if temp_list1[2:].isdigit():
										lowercase_letters = [c for c in temp_list1 if c.islower()]
										if len(lowercase_letters) > 0:
											data = {}
											data['status_code'] = "401"
											data['status'] = "Please provide proper cell format"
											print(json.dumps(data))

										else:	
											file_data_cellvalue = set_cell_value(file_name,sheet_name,temp_list1)
											print(file_data_cellvalue)	
											
									else:
										data = {}
										data['status_code'] = "401"
										data['status'] = "Please provide proper cell index"
										print(json.dumps(data))
										pass
								
								else:	
									lowercase_letters = [c for c in temp_list1 if c.islower()]
									if len(lowercase_letters) > 0:
										data = {}
										data['status_code'] = "401"
										data['status'] = "Please provide proper cell format"
										print(json.dumps(data))

									else:	
										file_data_cellvalue = set_cell_value(file_name,sheet_name,temp_list1)
										print(file_data_cellvalue)	
						
							else:
								lowercase_letters = [c for c in temp_list1 if c.islower()]
								if len(lowercase_letters) > 0:
									data = {}
									data['status_code'] = "401"
									data['status'] = "Please provide proper cell format"
									print(json.dumps(data))

								else:	
									file_data_cellvalue = set_cell_value(file_name,sheet_name,temp_list1)
									print(file_data_cellvalue)		
							
						else:
							data = {}
							data['status_code'] = "401"
							data['status'] = "Please provide proper cell index"
							print(json.dumps(data))
							pass
							
													
					elif file_extension in (''):
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass
						
					else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file extension"
						print(json.dumps(data))
						pass	
						
				else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass			
						
			else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide Existing file name"
						print(json.dumps(data))
						pass				
				
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e)
			print(json.dumps(data))
			pass


	elif commandKey == "DeleteCellValueByStartofDocument":
		try :
			if os.path.exists(sys.argv[2]): 
				filename, file_extension = os.path.splitext(sys.argv[2])			
				if filename[-1] not in (' '):
					if sys.argv[2].endswith('xls') or sys.argv[2].endswith('xlsx'):	
						file_name =  sys.argv[2]	
						sheet_name = sys.argv[3]	

						def my_split(s):
							return filter(None, re.split(r'(\d+)', s))
						
						def set_cell_value(file_name,sheet_name):							
							book = xlrd.open_workbook(file_name)
							sh = book.sheet_by_name(sheet_name)
							book1 = Workbook()
							rb = xlrd.open_workbook(file_name)
							wb = xl_copy(rb)
							sheet = wb.get_sheet(sheet_name)
							

							sheet.write(0,0,"")
							
							#sheet.write(cval1,rval,replace_text)
							wb.save(file_name)							
							data = {}
							data['status_code'] = "200 OK"
							data['status'] = "Cell Deleted successfully"
							data = json.dumps(data)
							return data
																
						file_data_cellvalue = set_cell_value(file_name,sheet_name)
						print(file_data_cellvalue)	
							
							
					elif file_extension in (''):
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass
						
					else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file extension"
						print(json.dumps(data))
						pass	
						
				else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass			
						
			else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide Existing file name"
						print(json.dumps(data))
						pass				
				
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e)
			print(json.dumps(data))
			pass


	elif commandKey == "DeleteCellValueByStartofColumn":
		try :
			if os.path.exists(sys.argv[2]): 				
				filename, file_extension = os.path.splitext(sys.argv[2])			
				if filename[-1] not in (' '):
					if sys.argv[2].endswith('xls') or sys.argv[2].endswith('xlsx'):
						file_name =  sys.argv[2]	
						sheet_name =  sys.argv[3]
						column_in_pre =  sys.argv[4]
						
						column_in = col_to_num(column_in_pre) -1

						def set_col(file_name):
							book = xlrd.open_workbook(file_name)
							sh = book.sheet_by_name(sheet_name)
							book1 = Workbook()
							rb = xlrd.open_workbook(file_name)
							wb = xl_copy(rb)
							sheet = wb.get_sheet(sheet_name)
							total_rows = sh.nrows
							k = int(column_in)

							sheet.write(0, k, "")
								
							wb.save(file_name)
							data = {}
							data['status_code'] = "200 OK"
							data['status'] = "Cell Deleted successfully"
							data = json.dumps(data)
							return data
						

						lowercase_letters = [c for c in column_in_pre if c.islower()]
						if len(lowercase_letters) > 0:
							data = {}
							data['status_code'] = "401"
							data['status'] = "Please provide proper cell format"
							print(json.dumps(data))

						else:	
							def my_split(s):
								return filter(None, re.split(r'(\d+)', s))
								
							rowindexarray = my_split(column_in_pre)
							if len(rowindexarray) > 1:
								data = {}
								data['status_code'] = "401"
								data['status'] = "Please provide proper cell format"
								print(json.dumps(data))

							else:
								file_data_col = set_col(file_name)
								print(file_data_col)	
					
					elif file_extension in (''):
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass
						
					else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file extension"
						print(json.dumps(data))
						pass	
						
				else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass			
						
			else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide Existing file name"
						print(json.dumps(data))
						pass		
		
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e)
			print(json.dumps(data))
			pass


	elif commandKey == "DeleteCellValueByEndofColumn":
		try :
			if os.path.exists(sys.argv[2]): 				
				filename, file_extension = os.path.splitext(sys.argv[2])			
				if filename[-1] not in (' '):
					if sys.argv[2].endswith('xls') or sys.argv[2].endswith('xlsx'):
						file_name =  sys.argv[2]	
						sheet_name =  sys.argv[3]
						column_in_pre =  sys.argv[4]
						
						column_in = col_to_num(column_in_pre) -1

						def set_col(file_name):
							book = xlrd.open_workbook(file_name)
							sh = book.sheet_by_name(sheet_name)
							book1 = Workbook()
							rb = xlrd.open_workbook(file_name)
							wb = xl_copy(rb)
							sheet = wb.get_sheet(sheet_name)
							total_rows = sh.nrows
							k = int(column_in)
							last_row = total_rows - 1
							sheet.write(last_row, k, "")
								
							wb.save(file_name)
							data = {}
							data['status_code'] = "200 OK"
							data['status'] = "Cell Deleted successfully"
							data = json.dumps(data)
							return data
						

						lowercase_letters = [c for c in column_in_pre if c.islower()]
						if len(lowercase_letters) > 0:
							data = {}
							data['status_code'] = "401"
							data['status'] = "Please provide proper cell format"
							print(json.dumps(data))

						else:	
							def my_split(s):
								return filter(None, re.split(r'(\d+)', s))
								
							rowindexarray = my_split(column_in_pre)
							if len(rowindexarray) > 1:
								data = {}
								data['status_code'] = "401"
								data['status'] = "Please provide proper cell format"
								print(json.dumps(data))

							else:
								file_data_col = set_col(file_name)
								print(file_data_col)	
					
					elif file_extension in (''):
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass
						
					else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file extension"
						print(json.dumps(data))
						pass	
						
				else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass			
						
			else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide Existing file name"
						print(json.dumps(data))
						pass		
		
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e)
			print(json.dumps(data))
			pass


	elif commandKey == "DeleteCellValueByStartofRow":
		try :
			if os.path.exists(sys.argv[2]): 				
				filename, file_extension = os.path.splitext(sys.argv[2])			
				if filename[-1] not in (' '):
					if sys.argv[2].endswith('xls') or sys.argv[2].endswith('xlsx'):
						file_name =  sys.argv[2]	
						sheet_name =  sys.argv[3]
						column_in =  sys.argv[4]

						def set_row(file_name):
							book = xlrd.open_workbook(file_name)
							sh = book.sheet_by_name(sheet_name)
							book1 = Workbook()
							#sheet1 = book1.add_sheet('test', cell_overwrite_ok=True)
							rb = xlrd.open_workbook(file_name)
							wb = xl_copy(rb)
							sheet = wb.get_sheet(sheet_name)
							total_cols = sh.ncols
							k = int(column_in) - 1
							#print(sys.argv[5])
							#print(column_in)

							sheet.write(k, 0, "")
								
							wb.save(file_name)
							data = {}
							data['status_code'] = "200 OK"
							data['status'] = "Cell Deleted successfully"
							data = json.dumps(data)
							return data
						

						file_data_row = set_row(file_name)
						print(file_data_row)	
								
					
					elif file_extension in (''):
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass
						
					else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file extension"
						print(json.dumps(data))
						pass	
						
				else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass			
						
			else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide Existing file name"
						print(json.dumps(data))
						pass		
		
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e)
			print(json.dumps(data))
			pass


	elif commandKey == "DeleteCellValueByEndofRow":
		try :
			if os.path.exists(sys.argv[2]): 				
				filename, file_extension = os.path.splitext(sys.argv[2])			
				if filename[-1] not in (' '):
					if sys.argv[2].endswith('xls') or sys.argv[2].endswith('xlsx'):
						file_name =  sys.argv[2]	
						sheet_name =  sys.argv[3]
						column_in =  sys.argv[4]

						def set_row(file_name):
							book = xlrd.open_workbook(file_name)
							sh = book.sheet_by_name(sheet_name)
							book1 = Workbook()
							#sheet1 = book1.add_sheet('test', cell_overwrite_ok=True)
							rb = xlrd.open_workbook(file_name)
							wb = xl_copy(rb)
							sheet = wb.get_sheet(sheet_name)
							total_cols = sh.ncols
							k = int(column_in) - 1
							#print(sys.argv[5])
							#print(column_in)
							last_column = total_cols - 1

							sheet.write(k, last_column, "")
								
							wb.save(file_name)
							data = {}
							data['status_code'] = "200 OK"
							data['status'] = "Cell Deleted successfully"
							data = json.dumps(data)
							return data
						

						file_data_row = set_row(file_name)
						print(file_data_row)	
								
					
					elif file_extension in (''):
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass
						
					else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file extension"
						print(json.dumps(data))
						pass	
						
				else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass			
						
			else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide Existing file name"
						print(json.dumps(data))
						pass		
		
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e)
			print(json.dumps(data))
			pass


	elif commandKey == "DeleteCurrentCellValue":
		try :
			if os.path.exists(sys.argv[2]): 
				filename, file_extension = os.path.splitext(sys.argv[2])			
				if filename[-1] not in (' '):
					if sys.argv[2].endswith('xls') or sys.argv[2].endswith('xlsx'):	
						file_name =  sys.argv[2]	
						sheet_name = sys.argv[3]	
						temp_list1 = sys.argv[4]

						def my_split(s):
							return filter(None, re.split(r'(\d+)', s))
						
						def set_cell_value(file_name,sheet_name,temp_list1):							
							book = xlrd.open_workbook(file_name)
							sh = book.sheet_by_name(sheet_name)
							book1 = Workbook()
							rb = xlrd.open_workbook(file_name)
							wb = xl_copy(rb)
							sheet = wb.get_sheet(sheet_name)
							
							#original cell
							rowindexarray = my_split(temp_list1)
							cval = rowindexarray[0]
							rval = col_to_num(cval)
							cval1 = rowindexarray[1]
							rval = int(rval) -1
							cval1 = int(cval1) - 1
														
							#value = sh.cell_value(cval1, rval)
							sheet.write(cval1,rval,"")
							
							#sheet.write(cval1,rval,replace_text)
							wb.save(file_name)							
							data = {}
							data['status_code'] = "200 OK"
							data['status'] = "Cell Deleted successfully"
							data = json.dumps(data)
							return data
												
						temp_list1_prevent = re.findall('[^A-Z0-9]',temp_list1)
						if len(temp_list1_prevent) < 1:
							if len(temp_list1) > 2 :
								if temp_list1[1].isdigit():
									if temp_list1[2:].isdigit():
										lowercase_letters = [c for c in temp_list1 if c.islower()]
										if len(lowercase_letters) > 0:
											data = {}
											data['status_code'] = "401"
											data['status'] = "Please provide proper cell format"
											print(json.dumps(data))

										else:	
											file_data_cellvalue = set_cell_value(file_name,sheet_name,temp_list1)
											print(file_data_cellvalue)	
											
									else:
										data = {}
										data['status_code'] = "401"
										data['status'] = "Please provide proper cell index"
										print(json.dumps(data))
										pass
								
								else:
									lowercase_letters = [c for c in temp_list1 if c.islower()]
									if len(lowercase_letters) > 0:
										data = {}
										data['status_code'] = "401"
										data['status'] = "Please provide proper cell format"
										print(json.dumps(data))

									else:	
										file_data_cellvalue = set_cell_value(file_name,sheet_name,temp_list1)
										print(file_data_cellvalue)	
									

							else:
								lowercase_letters = [c for c in temp_list1 if c.islower()]
								if len(lowercase_letters) > 0:
									data = {}
									data['status_code'] = "401"
									data['status'] = "Please provide proper cell format"
									print(json.dumps(data))

								else:	
									file_data_cellvalue = set_cell_value(file_name,sheet_name,temp_list1)
									print(file_data_cellvalue)		
							
						else:
							data = {}
							data['status_code'] = "401"
							data['status'] = "Please provide proper cell index"
							print(json.dumps(data))
							pass
									
													
					elif file_extension in (''):
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass
						
					else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file extension"
						print(json.dumps(data))
						pass	
						
				else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass			
						
			else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide Existing file name"
						print(json.dumps(data))
						pass				
				
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e)
			print(json.dumps(data))
			pass	


	elif commandKey == "DeleteCellValueAbove":
		try :
			if os.path.exists(sys.argv[2]): 
				filename, file_extension = os.path.splitext(sys.argv[2])			
				if filename[-1] not in (' '):
					if sys.argv[2].endswith('xls') or sys.argv[2].endswith('xlsx'):	
						file_name =  sys.argv[2]	
						sheet_name = sys.argv[3]	
						temp_list1 = sys.argv[4]

						def my_split(s):
							return filter(None, re.split(r'(\d+)', s))
						
						def set_cell_value(file_name,sheet_name,temp_list1):							
							book = xlrd.open_workbook(file_name)
							sh = book.sheet_by_name(sheet_name)
							book1 = Workbook()
							rb = xlrd.open_workbook(file_name)
							wb = xl_copy(rb)
							sheet = wb.get_sheet(sheet_name)
							
							#original cell
							rowindexarray = my_split(temp_list1)
							cval = rowindexarray[0]
							rval = col_to_num(cval)
							cval1 = rowindexarray[1]
							rval = int(rval) -1
							cval1 = int(cval1) - 1
							cval_above = cval1 - 1 		
							
							if cval_above >= 0:
								#value = sh.cell_value(cval1, rval)
								sheet.write(cval_above,rval,"")
								
								#sheet.write(cval1,rval,replace_text)
								wb.save(file_name)							
								data = {}
								data['status_code'] = "200 OK"
								data['status'] = "Cell Deleted successfully"
								data = json.dumps(data)
								return data

							else:
								wb.save(file_name)							
								data = {}
								data['status_code'] = "401"
								data['status'] = "There is no cell above the current cell"
								data = json.dumps(data)
								return data
												
						temp_list1_prevent = re.findall('[^A-Z0-9]',temp_list1)
						if len(temp_list1_prevent) < 1:
							if len(temp_list1) > 2 :
								if temp_list1[1].isdigit():
									if temp_list1[2:].isdigit():
										lowercase_letters = [c for c in temp_list1 if c.islower()]
										if len(lowercase_letters) > 0:
											data = {}
											data['status_code'] = "401"
											data['status'] = "Please provide proper cell format"
											print(json.dumps(data))

										else:	
											file_data_cellvalue = set_cell_value(file_name,sheet_name,temp_list1)
											print(file_data_cellvalue)	
											
									else:
										data = {}
										data['status_code'] = "401"
										data['status'] = "Please provide proper cell index"
										print(json.dumps(data))
										pass
								
								else:	
									lowercase_letters = [c for c in temp_list1 if c.islower()]
									if len(lowercase_letters) > 0:
										data = {}
										data['status_code'] = "401"
										data['status'] = "Please provide proper cell format"
										print(json.dumps(data))

									else:	
										file_data_cellvalue = set_cell_value(file_name,sheet_name,temp_list1)
										print(file_data_cellvalue)	
										
							else:
								lowercase_letters = [c for c in temp_list1 if c.islower()]
								if len(lowercase_letters) > 0:
									data = {}
									data['status_code'] = "401"
									data['status'] = "Please provide proper cell format"
									print(json.dumps(data))

								else:	
									file_data_cellvalue = set_cell_value(file_name,sheet_name,temp_list1)
									print(file_data_cellvalue)			
							
						else:
							data = {}
							data['status_code'] = "401"
							data['status'] = "Please provide proper cell index"
							print(json.dumps(data))
							pass	
						

													
					elif file_extension in (''):
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass
						
					else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file extension"
						print(json.dumps(data))
						pass	
						
				else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass			
						
			else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide Existing file name"
						print(json.dumps(data))
						pass				
				
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e)
			print(json.dumps(data))
			pass


	elif commandKey == "DeleteCellValueBelow":
		try :
			if os.path.exists(sys.argv[2]): 
				filename, file_extension = os.path.splitext(sys.argv[2])			
				if filename[-1] not in (' '):
					if sys.argv[2].endswith('xls') or sys.argv[2].endswith('xlsx'):	
						file_name =  sys.argv[2]	
						sheet_name = sys.argv[3]	
						temp_list1 = sys.argv[4]

						def my_split(s):
							return filter(None, re.split(r'(\d+)', s))
						
						def set_cell_value(file_name,sheet_name,temp_list1):							
							book = xlrd.open_workbook(file_name)
							sh = book.sheet_by_name(sheet_name)
							book1 = Workbook()
							rb = xlrd.open_workbook(file_name)
							wb = xl_copy(rb)
							sheet = wb.get_sheet(sheet_name)
							
							#original cell
							rowindexarray = my_split(temp_list1)
							cval = rowindexarray[0]
							rval = col_to_num(cval)
							cval1 = rowindexarray[1]
							rval = int(rval) -1
							cval1 = int(cval1) - 1
							cval_below = cval1 + 1 		
							if cval_below >= 1:
								#value = sh.cell_value(cval1, rval)
								sheet.write(cval_below,rval,"")
								
								#sheet.write(cval1,rval,replace_text)
								wb.save(file_name)							
								data = {}
								data['status_code'] = "200 OK"
								data['status'] = "Cell Deleted successfully"
								data = json.dumps(data)
								return data

							else:
								wb.save(file_name)							
								data = {}
								data['status_code'] = "401"
								data['status'] = "There is no cell below the current cell"
								data = json.dumps(data)
								return data
												

						temp_list1_prevent = re.findall('[^A-Z0-9]',temp_list1)
						if len(temp_list1_prevent) < 1:
							if len(temp_list1) > 2 :
								if temp_list1[1].isdigit():
									if temp_list1[2:].isdigit():
										lowercase_letters = [c for c in temp_list1 if c.islower()]
										if len(lowercase_letters) > 0:
											data = {}
											data['status_code'] = "401"
											data['status'] = "Please provide proper cell format"
											print(json.dumps(data))

										else:	
											file_data_cellvalue = set_cell_value(file_name,sheet_name,temp_list1)
											print(file_data_cellvalue)
											
									else:
										data = {}
										data['status_code'] = "401"
										data['status'] = "Please provide proper cell index"
										print(json.dumps(data))
										pass
								
								else:															
									lowercase_letters = [c for c in temp_list1 if c.islower()]
									if len(lowercase_letters) > 0:
										data = {}
										data['status_code'] = "401"
										data['status'] = "Please provide proper cell format"
										print(json.dumps(data))

									else:	
										file_data_cellvalue = set_cell_value(file_name,sheet_name,temp_list1)
										print(file_data_cellvalue)	
						
							else:
								lowercase_letters = [c for c in temp_list1 if c.islower()]
								if len(lowercase_letters) > 0:
									data = {}
									data['status_code'] = "401"
									data['status'] = "Please provide proper cell format"
									print(json.dumps(data))

								else:	
									file_data_cellvalue = set_cell_value(file_name,sheet_name,temp_list1)
									print(file_data_cellvalue)		
							
						else:
							data = {}
							data['status_code'] = "401"
							data['status'] = "Please provide proper cell index"
							print(json.dumps(data))
							pass
						
													
					elif file_extension in (''):
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass
						
					else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file extension"
						print(json.dumps(data))
						pass	
						
				else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass			
						
			else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide Existing file name"
						print(json.dumps(data))
						pass				
				
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e)
			print(json.dumps(data))
			pass


	elif commandKey == "DeleteCellValueToLeft":
		try :
			if os.path.exists(sys.argv[2]): 
				filename, file_extension = os.path.splitext(sys.argv[2])			
				if filename[-1] not in (' '):
					if sys.argv[2].endswith('xls') or sys.argv[2].endswith('xlsx'):	
						file_name =  sys.argv[2]	
						sheet_name = sys.argv[3]	
						temp_list1 = sys.argv[4]

						def my_split(s):
							return filter(None, re.split(r'(\d+)', s))
						
						def set_cell_value(file_name,sheet_name,temp_list1):							
							book = xlrd.open_workbook(file_name)
							sh = book.sheet_by_name(sheet_name)
							book1 = Workbook()
							rb = xlrd.open_workbook(file_name)
							wb = xl_copy(rb)
							sheet = wb.get_sheet(sheet_name)
							
							#original cell
							rowindexarray = my_split(temp_list1)
							cval = rowindexarray[0]
							rval = col_to_num(cval)
							cval1 = rowindexarray[1]
							rval = int(rval) -1
							cval1 = int(cval1) - 1
							cell_left = rval - 1
							
							if cval1 >= 0 and cell_left >= 0:
								#value = sh.cell_value(cval1, rval)
								sheet.write(cval1,cell_left,"")
								
								#sheet.write(cval1,rval,replace_text)
								wb.save(file_name)							
								data = {}
								data['status_code'] = "200 OK"
								data['status'] = "Cell Deleted successfully"
								data = json.dumps(data)
								return data

							else:
								wb.save(file_name)							
								data = {}
								data['status_code'] = "401"
								data['status'] = "There is no cell left to the current cell"
								data = json.dumps(data)
								return data
												
						temp_list1_prevent = re.findall('[^A-Z0-9]',temp_list1)
						if len(temp_list1_prevent) < 1:
							if len(temp_list1) > 2 :
								if temp_list1[1].isdigit():
									if temp_list1[2:].isdigit():
										lowercase_letters = [c for c in temp_list1 if c.islower()]
										if len(lowercase_letters) > 0:
											data = {}
											data['status_code'] = "401"
											data['status'] = "Please provide proper cell format"
											print(json.dumps(data))

										else:	
											file_data_cellvalue = set_cell_value(file_name,sheet_name,temp_list1)
											print(file_data_cellvalue)
											
									else:
										data = {}
										data['status_code'] = "401"
										data['status'] = "Please provide proper cell index"
										print(json.dumps(data))
										pass
								
								else:	
									lowercase_letters = [c for c in temp_list1 if c.islower()]
									if len(lowercase_letters) > 0:
										data = {}
										data['status_code'] = "401"
										data['status'] = "Please provide proper cell format"
										print(json.dumps(data))

									else:	
										file_data_cellvalue = set_cell_value(file_name,sheet_name,temp_list1)
										print(file_data_cellvalue)	
							
							else:
								lowercase_letters = [c for c in temp_list1 if c.islower()]
								if len(lowercase_letters) > 0:
									data = {}
									data['status_code'] = "401"
									data['status'] = "Please provide proper cell format"
									print(json.dumps(data))

								else:	
									file_data_cellvalue = set_cell_value(file_name,sheet_name,temp_list1)
									print(file_data_cellvalue)		
							
						else:
							data = {}
							data['status_code'] = "401"
							data['status'] = "Please provide proper cell index"
							print(json.dumps(data))
							pass

													
					elif file_extension in (''):
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass
						
					else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file extension"
						print(json.dumps(data))
						pass	
						
				else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass			
						
			else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide Existing file name"
						print(json.dumps(data))
						pass				
				
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e)
			print(json.dumps(data))
			pass


	elif commandKey == "DeleteCellValueToRight":
		try :
			if os.path.exists(sys.argv[2]): 
				filename, file_extension = os.path.splitext(sys.argv[2])			
				if filename[-1] not in (' '):
					if sys.argv[2].endswith('xls') or sys.argv[2].endswith('xlsx'):	
						file_name =  sys.argv[2]	
						sheet_name = sys.argv[3]	
						temp_list1 = sys.argv[4]

						def my_split(s):
							return filter(None, re.split(r'(\d+)', s))
						
						def set_cell_value(file_name,sheet_name,temp_list1):							
							book = xlrd.open_workbook(file_name)
							sh = book.sheet_by_name(sheet_name)
							book1 = Workbook()
							rb = xlrd.open_workbook(file_name)
							wb = xl_copy(rb)
							sheet = wb.get_sheet(sheet_name)
							
							#original cell
							rowindexarray = my_split(temp_list1)
							cval = rowindexarray[0]
							rval = col_to_num(cval)
							cval1 = rowindexarray[1]
							rval = int(rval) -1
							cval1 = int(cval1) - 1
							cell_left = rval + 1
							
							if cell_left >= 0:
								#value = sh.cell_value(cval1, rval)
								sheet.write(cval1,cell_left,"")
								
								#sheet.write(cval1,rval,replace_text)
								wb.save(file_name)							
								data = {}
								data['status_code'] = "200 OK"
								data['status'] = "Cell Deleted successfully"
								data = json.dumps(data)
								return data

							else:
								wb.save(file_name)							
								data = {}
								data['status_code'] = "401"
								data['status'] = "There is no cell right to the current cell"
								data = json.dumps(data)
								return data
												
						temp_list1_prevent = re.findall('[^A-Z0-9]',temp_list1)
						if len(temp_list1_prevent) < 1:
							if len(temp_list1) > 2 :
								if temp_list1[1].isdigit():
									if temp_list1[2:].isdigit():
										lowercase_letters = [c for c in temp_list1 if c.islower()]
										if len(lowercase_letters) > 0:
											data = {}
											data['status_code'] = "401"
											data['status'] = "Please provide proper cell format"
											print(json.dumps(data))

										else:	
											file_data_cellvalue = set_cell_value(file_name,sheet_name,temp_list1)
											print(file_data_cellvalue)
											
									else:
										data = {}
										data['status_code'] = "401"
										data['status'] = "Please provide proper cell index"
										print(json.dumps(data))
										pass
								
								else:
									lowercase_letters = [c for c in temp_list1 if c.islower()]
									if len(lowercase_letters) > 0:
										data = {}
										data['status_code'] = "401"
										data['status'] = "Please provide proper cell format"
										print(json.dumps(data))

									else:	
										file_data_cellvalue = set_cell_value(file_name,sheet_name,temp_list1)
										print(file_data_cellvalue)	
						
							else:
								lowercase_letters = [c for c in temp_list1 if c.islower()]
								if len(lowercase_letters) > 0:
									data = {}
									data['status_code'] = "401"
									data['status'] = "Please provide proper cell format"
									print(json.dumps(data))

								else:	
									file_data_cellvalue = set_cell_value(file_name,sheet_name,temp_list1)
									print(file_data_cellvalue)		
							
						else:
							data = {}
							data['status_code'] = "401"
							data['status'] = "Please provide proper cell index"
							print(json.dumps(data))
							pass
						
													
					elif file_extension in (''):
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass
						
					else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file extension"
						print(json.dumps(data))
						pass	
						
				else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						#print(json.dumps(data))
						pass			
						
			else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide Existing file name"
						print(json.dumps(data))
						#print(json.dumps(data))
						pass				
				
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e)
			print(json.dumps(data))
			#print(json.dumps(data))
			pass

			
	elif commandKey == "GetFirstColumn":
		try :
			if os.path.exists(sys.argv[2]): 				
				filename, file_extension = os.path.splitext(sys.argv[2])			
				if filename[-1] not in (' '):
					if sys.argv[2].endswith('xls') or sys.argv[2].endswith('xlsx'):
						file_name =  sys.argv[2]	
						sheet_name =  sys.argv[3]
						
						#column_in = col_to_num(column_in_pre) -1
						def set_col(file_name):
							book = xlrd.open_workbook(file_name)
							sh = book.sheet_by_name(sheet_name)
							book1 = Workbook()
							rb = xlrd.open_workbook(file_name)
							wb = xl_copy(rb)
							sheet = wb.get_sheet(sheet_name)
							total_rows = sh.nrows
							first_column = 0
							
							data_array = {}
							for i in range(0, total_rows):
								value = sh.cell_value(i, first_column)
								data_array[str(i + 1) +"*"+ str(first_column + 1)] =  value
									
							data = {}
							data['status_code'] = "200 OK"
							data['data'] = data_array
							data = json.dumps(data)
							return data
						
						file_data_col = set_col(file_name)
						print(file_data_col)	
					
					elif file_extension in (''):
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass
						
					else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file extension"
						print(json.dumps(data))
						pass	
						
				else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass			
						
			else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide Existing file name"
						print(json.dumps(data))
						pass		
		
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e)
			print(json.dumps(data))
			pass
			
			
	elif commandKey == "GetLastColumn":
		try :
			if os.path.exists(sys.argv[2]): 				
				filename, file_extension = os.path.splitext(sys.argv[2])			
				if filename[-1] not in (' '):
					if sys.argv[2].endswith('xls') or sys.argv[2].endswith('xlsx'):
						file_name =  sys.argv[2]	
						sheet_name =  sys.argv[3]
						
						#column_in = col_to_num(column_in_pre) -1
						def set_col(file_name):
							book = xlrd.open_workbook(file_name)
							sh = book.sheet_by_name(sheet_name)
							book1 = Workbook()
							rb = xlrd.open_workbook(file_name)
							wb = xl_copy(rb)
							sheet = wb.get_sheet(sheet_name)
							total_rows = sh.nrows
							last_column = sh.ncols -1
 
							data_array = {}
							for i in range(0, total_rows):
								value = sh.cell_value(i, last_column)
								data_array[str(i + 1) +"*"+ str(last_column + 1)] =  value
									
							data = {}
							data['status_code'] = "200 OK"
							data['data'] = data_array
							data = json.dumps(data)
							return data
						
						file_data_col = set_col(file_name)
						print(file_data_col)	
					
					elif file_extension in (''):
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass
						
					else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file extension"
						print(json.dumps(data))
						pass	
						
				else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass			
						
			else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide Existing file name"
						print(json.dumps(data))
						pass		
		
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e)
			print(json.dumps(data))
			pass
			
			
	elif commandKey == "GetNextColumn":
		try :
			if os.path.exists(sys.argv[2]): 				
				filename, file_extension = os.path.splitext(sys.argv[2])			
				if filename[-1] not in (' '):
					if sys.argv[2].endswith('xls') or sys.argv[2].endswith('xlsx'):
						file_name =  sys.argv[2]	
						sheet_name =  sys.argv[3]
						column_in_pre =  sys.argv[4]
						lowercase_letters = [c for c in column_in_pre if c.islower()]
						if len(lowercase_letters) > 0:
							data = {}
							data['status_code'] = "401"
							data['status'] = "Please provide proper column format"
							print(json.dumps(data))

						else:
							column_in = col_to_num(column_in_pre)
							def set_col(file_name):
								book = xlrd.open_workbook(file_name)
								sh = book.sheet_by_name(sheet_name)
								book1 = Workbook()
								rb = xlrd.open_workbook(file_name)
								wb = xl_copy(rb)
								sheet = wb.get_sheet(sheet_name)
								total_rows = sh.nrows
								
								if column_in > 0:
									data_array = {}
									for i in range(0, total_rows):
										value = sh.cell_value(i, column_in)
										data_array[str(i + 1) +"*"+ str(column_in + 1)] =  value
											
									data = {}
									data['status_code'] = "200 OK"
									data['data'] = data_array
									data = json.dumps(data)
									return data
								else:						
									data = {}
									data['status_code'] = "401"
									data['status'] = "There is no column next to the current column"
									data = json.dumps(data)
									return data	
							
							file_data_col = set_col(file_name)
							print(file_data_col)	
					
					elif file_extension in (''):
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass
						
					else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file extension"
						print(json.dumps(data))
						pass	
						
				else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass			
						
			else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide Existing file name"
						print(json.dumps(data))
						pass		
		
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e)
			print(json.dumps(data))
			pass
			
			
	elif commandKey == "GetPreviousColumn":
		try :
			if os.path.exists(sys.argv[2]): 				
				filename, file_extension = os.path.splitext(sys.argv[2])			
				if filename[-1] not in (' '):
					if sys.argv[2].endswith('xls') or sys.argv[2].endswith('xlsx'):
						file_name =  sys.argv[2]	
						sheet_name =  sys.argv[3]
						column_in_pre =  sys.argv[4]
						
						lowercase_letters = [c for c in column_in_pre if c.islower()]
						if len(lowercase_letters) > 0:
							data = {}
							data['status_code'] = "401"
							data['status'] = "Please provide proper column format"
							print(json.dumps(data))

						else:
							column_in = col_to_num(column_in_pre) - 2
							def set_col(file_name):
								book = xlrd.open_workbook(file_name)
								sh = book.sheet_by_name(sheet_name)
								book1 = Workbook()
								rb = xlrd.open_workbook(file_name)
								wb = xl_copy(rb)
								sheet = wb.get_sheet(sheet_name)
								total_rows = sh.nrows

								if column_in > -1:
									data_array = {}
									for i in range(0, total_rows):
										value = sh.cell_value(i, column_in)
										data_array[str(i + 1) +"*"+ str(column_in + 1)] =  value
											
									data = {}
									data['status_code'] = "200 OK"
									data['data'] = data_array
									data = json.dumps(data)
									return data
									
								else:						
									data = {}
									data['status_code'] = "401"
									data['status'] = "There is no column previous to the current column"
									data = json.dumps(data)
									return data	
							
							
							file_data_col = set_col(file_name)
							print(file_data_col)	
					
					elif file_extension in (''):
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass
						
					else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file extension"
						print(json.dumps(data))
						pass	
						
				else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass			
						
			else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide Existing file name"
						print(json.dumps(data))
						pass		
		
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e)
			print(json.dumps(data))
			pass
			
			
	elif commandKey == "GetFirstRow":
		try :
			if os.path.exists(sys.argv[2]): 				
				filename, file_extension = os.path.splitext(sys.argv[2])			
				if filename[-1] not in (' '):
					if sys.argv[2].endswith('xls') or sys.argv[2].endswith('xlsx'):
						file_name =  sys.argv[2]	
						sheet_name =  sys.argv[3]

						def get_row(file_name):
							book = xlrd.open_workbook(file_name)
							sh = book.sheet_by_name(sheet_name)
							book1 = Workbook()
							#sheet1 = book1.add_sheet('test', cell_overwrite_ok=True)
							rb = xlrd.open_workbook(file_name)
							wb = xl_copy(rb)
							sheet = wb.get_sheet(sheet_name)
							total_cols = sh.ncols
							#print(sys.argv[5])
							#print(column_in)
							column_in = 0 
							data = {}
							data_array = {}
							for i in range(0, total_cols):
								value = sh.cell_value(column_in,i)
								data_array[str(column_in + 1) +"*"+ str(i+1)] =  value
							data = {}
							data['status_code'] = "200 OK"
							data['data'] = data_array
							data = json.dumps(data)
							return data
						
						file_data_row = get_row(file_name)
						print(file_data_row)	
					
					elif file_extension in (''):
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass
						
					else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file extension"
						print(json.dumps(data))
						pass	
						
				else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass			
						
			else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide Existing file name"
						print(json.dumps(data))
						pass		
		
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e)
			print(json.dumps(data))
			pass
			

	elif commandKey == "GetLastRow":
		try :
			if os.path.exists(sys.argv[2]): 				
				filename, file_extension = os.path.splitext(sys.argv[2])			
				if filename[-1] not in (' '):
					if sys.argv[2].endswith('xls') or sys.argv[2].endswith('xlsx'):
						file_name =  sys.argv[2]	
						sheet_name =  sys.argv[3]

						def get_row(file_name):
							book = xlrd.open_workbook(file_name)
							sh = book.sheet_by_name(sheet_name)
							book1 = Workbook()
							#sheet1 = book1.add_sheet('test', cell_overwrite_ok=True)
							rb = xlrd.open_workbook(file_name)
							wb = xl_copy(rb)
							sheet = wb.get_sheet(sheet_name)
							total_cols = sh.ncols
							last_row = sh.nrows - 1  
							data_array = {}
							for i in range(0, total_cols):
								value = sh.cell_value(last_row,i)
								data_array[str(last_row + 1) +"*"+ str(i+1)] =  value
							data = {}
							data['status_code'] = "200 OK"
							data['data'] = data_array
							data = json.dumps(data)
							return data
						
						file_data_row = get_row(file_name)
						print(file_data_row)	
					
					elif file_extension in (''):
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass
						
					else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file extension"
						print(json.dumps(data))
						pass	
						
				else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass			
						
			else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide Existing file name"
						print(json.dumps(data))
						pass		
		
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e)
			print(json.dumps(data))
			pass
		

	elif commandKey == "GetNextRow":
		try :
			if os.path.exists(sys.argv[2]): 				
				filename, file_extension = os.path.splitext(sys.argv[2])			
				if filename[-1] not in (' '):
					if sys.argv[2].endswith('xls') or sys.argv[2].endswith('xlsx'):
						file_name =  sys.argv[2]	
						sheet_name =  sys.argv[3]
						column_in =  sys.argv[4]

						def get_row(file_name):
							book = xlrd.open_workbook(file_name)
							sh = book.sheet_by_name(sheet_name)
							book1 = Workbook()
							#sheet1 = book1.add_sheet('test', cell_overwrite_ok=True)
							rb = xlrd.open_workbook(file_name)
							wb = xl_copy(rb)
							sheet = wb.get_sheet(sheet_name)
							total_cols = sh.ncols
							#print(sys.argv[5])
					
							column_in = int(sys.argv[4])
							if column_in > 0:
								data = {}
								data_array = {}
								for i in range(0, total_cols):
									value = sh.cell_value(column_in,i)
									data_array[str( column_in + 1) +"*"+ str(i+1)] =  value
								data = {}
								data['status_code'] = "200 OK"
								data['data'] = data_array
								data = json.dumps(data)
								return data
							else:						
								data = {}
								data['status_code'] = "401"
								data['status'] = "There is no row next to the current row"
								data = json.dumps(data)
								return data
							
						file_data_row = get_row(file_name)
						print(file_data_row)	
					
					elif file_extension in (''):
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass
						
					else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file extension"
						print(json.dumps(data))
						pass	
						
				else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass			
						
			else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide Existing file name"
						print(json.dumps(data))
						pass		
		
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e)
			print(json.dumps(data))
			pass	
		
		
	elif commandKey == "GetPreviousRow":
		try :
			if os.path.exists(sys.argv[2]): 				
				filename, file_extension = os.path.splitext(sys.argv[2])			
				if filename[-1] not in (' '):
					if sys.argv[2].endswith('xls') or sys.argv[2].endswith('xlsx'):
						file_name =  sys.argv[2]	
						sheet_name =  sys.argv[3]
						column_in =  sys.argv[4]
						column_in1 = int(column_in)
						def get_row(file_name):
							book = xlrd.open_workbook(file_name)
							sh = book.sheet_by_name(sheet_name)
							book1 = Workbook()
							#sheet1 = book1.add_sheet('test', cell_overwrite_ok=True)
							rb = xlrd.open_workbook(file_name)
							wb = xl_copy(rb)
							sheet = wb.get_sheet(sheet_name)
							total_cols = sh.ncols
							#print(sys.argv[5])

							column_in = int(sys.argv[4]) - 2
							if column_in1 > 1:
								data_array = {}
								for i in range(0, total_cols):
									value = sh.cell_value(column_in,i)
									data_array[str( column_in + 1) +"*"+ str(i+1)] =  value
								data = {}
								data['status_code'] = "200 OK"
								data['data'] = data_array
								data = json.dumps(data)
								return data
								
							else:						
								data = {}
								data['status_code'] = "401"
								data['status'] = "There is no row previous to the current row"
								data = json.dumps(data)
								return data	
						
						file_data_row = get_row(file_name)
						print(file_data_row)	
					
					elif file_extension in (''):
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass
						
					else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file extension"
						print(json.dumps(data))
						pass	
						
				else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass			
						
			else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide Existing file name"
						print(json.dumps(data))
						pass		
		
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e)
			print(json.dumps(data))
			pass	
		
		
	elif commandKey == "ExcelMacro":
		try :
			if os.path.exists(sys.argv[2]): 				
				filename, file_extension = os.path.splitext(sys.argv[2])			
				if filename[-1] not in (' '):
					if sys.argv[2].endswith('xls') or sys.argv[2].endswith('xlsx'):
						file_name =  sys.argv[2]	
						macro_name =  sys.argv[3]
						
						def run_macro(file_name,macro_name):
								xl=win32com.client.Dispatch("Excel.Application")
								xl.Workbooks.Open(os.path.abspath(file_name), ReadOnly=1)
								xl.Application.Run(macro_name)
								#xl.Application.Save() # if you want to save then uncomment this line and change delete the ", ReadOnly=1" part from the open function.
								xl.Application.Quit() # Comment this out if your excel script closes
								del xl
								data = {}
								data['status_code'] = "200 OK"
								data['status'] = "Task completed successfully"
								data = json.dumps(data)
								return data
							
						
						file_data_row = run_macro(file_name,macro_name)
						print(file_data_row)	
					
					elif file_extension in (''):
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass
						
					else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file extension"
						print(json.dumps(data))
						pass	
						
				else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass			
						
			else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide Existing file name"
						print(json.dumps(data))
						pass		
		
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e)
			print(json.dumps(data))
			pass	
	
	
	elif commandKey == "SortTableByASC":
		try :		
			if os.path.exists(sys.argv[2]): 				
				filename, file_extension = os.path.splitext(sys.argv[2])			
				if filename[-1] not in (' '):
					if sys.argv[2].endswith('xls') or sys.argv[2].endswith('xlsx'):
						file_name =  sys.argv[2]	
						sheet_name =  sys.argv[3]
						column_name =  sys.argv[4]
						def findexcel_column_row(excel_list2):
							book = xlrd.open_workbook(file_name)
							sh = book.sheet_by_name(sheet_name)
							for row in range(sh.nrows):
								for column in range(sh.ncols):
									if excel_list2 == sh.cell(row, column).value:  
										return [row,column]
													

						col_row_Num = findexcel_column_row(column_name)	
						row_number = int(col_row_Num[0])
						col_number = int(col_row_Num[1])
						col_number1 = int(col_row_Num[1]) + 1

						#print(col_row_Num)			
						
						target_column = col_number

						book = xlrd.open_workbook(file_name)
						sheet = book.sheet_by_name(sheet_name)
						data = [sheet.col_values(i) for i in xrange(sheet.ncols)]
						labels = data[row_number]
						data = list(data[col_number][row_number +1:])
						final_data1 = sorted(data)
						final_data = filter(None, final_data1)
						
						rb = xlrd.open_workbook(file_name)
						wb = xl_copy(rb)
						sheet = wb.get_sheet(sheet_name)
						
						rowNum1 = row_number + 1 
						for i in final_data1:
							sheet.write(rowNum1, int(col_number), "")
							rowNum1 = rowNum1 + 1
							#k = k+1
						
						rowNum = row_number + 1 
						for i in final_data:
							sheet.write(rowNum, int(col_number), i)
							rowNum = rowNum + 1
							#k = k+1

						wb.save(file_name)
						data = {}
						data['status_code'] = "200 OK"
						data['status'] = "Column updated successfully"
						data = json.dumps(data)
						print(data)
						
					elif file_extension in (''):
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass
						
					else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file extension"
						print(json.dumps(data))
						pass	
						
				else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass			
						
			else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide Existing file name"
						print(json.dumps(data))
						pass
						

		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e)
			print(json.dumps(data))
			pass
			
			
			
	elif commandKey == "SortRowByASC":
		try :		
			if os.path.exists(sys.argv[2]): 				
				filename, file_extension = os.path.splitext(sys.argv[2])			
				if filename[-1] not in (' '):
					if sys.argv[2].endswith('xls') or sys.argv[2].endswith('xlsx'):
						file_name =  sys.argv[2]	
						sheet_name =  sys.argv[3]
						column_name =  sys.argv[4]
						def findexcel_column_row(excel_list2):
							book = xlrd.open_workbook(file_name)
							sh = book.sheet_by_name(sheet_name)
							for row in range(sh.nrows):
								for column in range(sh.ncols):
									if excel_list2 == sh.cell(row, column).value:  
										return [row,column]
													

						col_row_Num = findexcel_column_row(column_name)	
						row_number = int(col_row_Num[0])
						col_number = int(col_row_Num[1])
						col_number1 = int(col_row_Num[1]) + 1
								
						target_column = col_number

						book = xlrd.open_workbook(file_name)
						sheet = book.sheet_by_name(sheet_name)
						#data = [sheet.col_values(i) for i in xrange(sheet.ncols)]
						data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]
						#print(data)
						labels = data[row_number]
						final_data1 = sorted(data,key=operator.itemgetter(col_number))
						final_data = filter(None, final_data1)
						rb = xlrd.open_workbook(file_name)
						wb = xl_copy(rb)
						sheet = wb.get_sheet(sheet_name)
						
						rowNum1 = row_number + 1 
						for i in final_data1:
							sheet.write(rowNum1, int(col_number), "")
							rowNum1 = rowNum1 + 1
							#k = k+1
						
						rowNum = row_number + 1 
						for rownum, sublist in enumerate(final_data1):
							for colnum, value in enumerate(sublist):
								sheet.write(rownum+1, colnum, value)
							#k = k+1
							

						wb.save(file_name)
						data = {}
						data['status_code'] = "200 OK"
						data['status'] = "Column updated successfully"
						data = json.dumps(data)
						print(data)
						
					elif file_extension in (''):
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass
						
					else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file extension"
						print(json.dumps(data))
						pass	
						
				else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass			
						
			else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide Existing file name"
						print(json.dumps(data))
						pass
						

		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e)
			print(json.dumps(data))
			pass

			
	elif commandKey == "SortTableByDSC":
		try :		
			if os.path.exists(sys.argv[2]): 				
				filename, file_extension = os.path.splitext(sys.argv[2])			
				if filename[-1] not in (' '):
					if sys.argv[2].endswith('xls') or sys.argv[2].endswith('xlsx'):
						file_name =  sys.argv[2]	
						sheet_name =  sys.argv[3]
						column_name =  sys.argv[4]
						def findexcel_column_row(excel_list2):
							book = xlrd.open_workbook(file_name)
							sh = book.sheet_by_name(sheet_name)
							for row in range(sh.nrows):
								for column in range(sh.ncols):
									if excel_list2 == sh.cell(row, column).value:  
										return [row,column]
													

						col_row_Num = findexcel_column_row(column_name)	
						row_number = int(col_row_Num[0])
						col_number = int(col_row_Num[1])
						col_number1 = int(col_row_Num[1]) + 1

						#print(col_row_Num)			
						
						target_column = col_number

						book = xlrd.open_workbook(file_name)
						sheet = book.sheet_by_name(sheet_name)
						data = [sheet.col_values(i) for i in xrange(sheet.ncols)]
						labels = data[row_number]
						data = list(data[col_number][row_number +1:])
						final_data1 = sorted(data, reverse=True)
						final_data = filter(None, final_data1)
						
						rb = xlrd.open_workbook(file_name)
						wb = xl_copy(rb)
						sheet = wb.get_sheet(sheet_name)
						
						rowNum1 = row_number + 1 
						for i in final_data1:
							sheet.write(rowNum1, int(col_number), "")
							rowNum1 = rowNum1 + 1
							#k = k+1
						
						rowNum = row_number + 1 
						for i in final_data:
							sheet.write(rowNum, int(col_number), i)
							rowNum = rowNum + 1
							#k = k+1

						wb.save(file_name)
						data = {}
						data['status_code'] = "200 OK"
						data['status'] = "Column updated successfully"
						data = json.dumps(data)
						print(data)
						
					elif file_extension in (''):
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass
						
					else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file extension"
						print(json.dumps(data))
						pass	
						
				else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass			
						
			else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide Existing file name"
						print(json.dumps(data))
						pass
						

		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e)
			print(json.dumps(data))
			pass
			
			
	elif commandKey == "FilterTable":
		try :
			if os.path.exists(sys.argv[2]): 				
				filename, file_extension = os.path.splitext(sys.argv[2])			
				if filename[-1] not in (' '):
					if sys.argv[2].endswith('xls') or sys.argv[2].endswith('xlsx'):	
						file_name =  sys.argv[2]	
						sheet_name = sys.argv[3]	
						
						#cells 
						temp_list_in = sys.argv[4]
						temp_list2 = temp_list_in.split(',')
						cell_text_list = list(temp_list2)
						
						def my_split(s):
							return filter(None, re.split(r'(\d+)', s))
																									
						book = xlrd.open_workbook(file_name)
						sh = book.sheet_by_name(sheet_name)
						book1 = Workbook()
						rb = xlrd.open_workbook(file_name)
						wb = xl_copy(rb)
						sheet = wb.get_sheet(sheet_name)
						final_index = 0
						
						def set_cell_range(file_name,cell,sheet_name,data_index):	
							#original cells index
							rowindexarray = my_split(cell)
							cval = rowindexarray[0]
							rval = col_to_num(cval)
							cval1 = rowindexarray[1]
							rval = int(rval) -1
							cval1 = int(cval1) - 1	
							global final_index
							if final_index == 0:
								final_index = data_index + 1
							else:
								final_index = final_index
							
							#print(final_index)
							
							value = sh.cell_value(cval1, rval)
							value_final = value
							#print(value)
							value1 = sh.cell_value(data_index, rval)
							
							if value_final != value1:
								data2 = [sh.row_values(i) for i in xrange(sh.nrows)]
								for i in range(0, len(data2)):
									data = [sh.cell_value(i, col) for col in range(sh.ncols)]

									if data[rval] == value_final :
										for index, value in enumerate(data):
											#sheet.write(final_index, index, value)
											#print(value)
											sheet.write(final_index, index, value)
												
										final_index = final_index + 1
									
									
							wb.save(file_name)							
							data = {}
							data['status_code'] = "200 OK"
							data['status'] = "Data filtered successfully"
							data = json.dumps(data)
							return data
												
						total_cols = sh.ncols
						total_rows = sh.nrows
						
						data_find = []
						for j in range(0, total_rows): #for rows
								for i in range(0, total_cols):
									value = sh.cell_value(j,i)
									if value != "" and len(data_find) < 1:
										data_find.append(j)
										break
						
						if len(data_find) == 1:
							data2 = [sh.row_values(i) for i in xrange(sh.nrows)]
							data1 = data2[data_find[0]] 
						else:
							data1 = []
						
						for j in range(0, total_rows): #for rows
								for i in range(0, total_cols):
									sheet.write(j, i, "")
						
						for index, value in enumerate(data1):
							sheet.write(int(data_find[0]), index, value)
							
						i = 0	
						data_find_index = data_find[0]
						for cell in cell_text_list:
							file_data_cellvalue = set_cell_range(file_name,cell,sheet_name,data_find_index)
							data_find_index = data_find_index + 1
							i = i + 1
							
						print(file_data_cellvalue)	
						
															
					elif file_extension in (''):
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass
						
					else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file extension"
						print(json.dumps(data))
						pass	
						
				else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass			
						
			else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide Existing file name"
						print(json.dumps(data))
						pass				
				
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e)
			print(json.dumps(data))
			pass
			
		
	elif commandKey == "GetColumnRange":
		try :
			if os.path.exists(sys.argv[2]): 				
				filename, file_extension = os.path.splitext(sys.argv[2])			
				if filename[-1] not in (' '):
					if sys.argv[2].endswith('xls') or sys.argv[2].endswith('xlsx'):
						file_name =  sys.argv[2]	
						sheet_name =  sys.argv[3]
						
						#row input 
						cols_in = sys.argv[4]
						header_status = sys.argv[5]
						range_text1 = sys.argv[6]
						range_text1 = string.replace(range_text1, ':', ',')
						range_text2 = range_text1.split(',')

						
						if header_status == "FALSE":
							cols_in2 = cols_in.split(',')
							cols_in_list1 = list(cols_in2)
							cols_in_list = list(cols_in_list1)
													

							def get_rows(file_name):
								book = xlrd.open_workbook(file_name)
								sh = book.sheet_by_name(sheet_name)
								book1 = Workbook()
								rb = xlrd.open_workbook(file_name)
								wb = xl_copy(rb)
								sheet = wb.get_sheet(sheet_name)
								total_cols = sh.ncols
								total_rows = sh.nrows
								data_array = []
								
								for j in range(1, total_rows+1): #for rows
									data_array_elemnts = {}
									for i in cols_in_list:
										#print(i)
										column_in = col_to_num(i) -1
										#print(column_in)
										value = sh.cell_value(int(j-1),column_in)
										data_array_elemnts[str(i)] = value
										#data_array_elemnts['value'] = value
										data_array_elemnts['row_index'] = str(j)
									data_array.append(data_array_elemnts)	
									
									
									
								if len(range_text2) ==2:
									var1_range = int(range_text2[0])
									var2_range = int(range_text2[1]) + 1 
									range_text = range(var1_range,var2_range)
									range_text[:] = [x - 1 for x in range_text]
									data_array = [data_array[i] for i in range_text]
								
								elif len(range_text2) ==1:	
									var1_range = int(range_text2[0]) 
									if var1_range > 0:
										range_text = var1_range - 1
										data_array = data_array[range_text]
									else:
										data = {}
										data['status_code'] = "401"
										data['status'] = "Please provide proper range"
										print(json.dumps(data))
										pass
										sys.exit()
								#print(range	
								else:
									data = {}
									data['status_code'] = "401"
									data['status'] = "Please provide proper range"
									print(json.dumps(data))
									pass
									sys.exit()	
									
										
								data = {}
								data['status_code'] = "200 OK"
								data['data'] = data_array
								data = json.dumps(data)
								return data
							
							file_data_rowsget = get_rows(file_name)
							print(file_data_rowsget)	
							
						elif header_status == "TRUE":
							cols_in2 = cols_in.split(',')
							cols_in_list1 = list(cols_in2)
							cols_in_list = list(cols_in_list1)
												

							def findexcel_column1(excel_list_index):
								book = xlrd.open_workbook(file_name)
								sh = book.sheet_by_name(sheet_name)
								for row in range(sh.nrows):
									for column in range(sh.ncols):
										if excel_list_index == sh.cell(row, column).value:  
											return column												

							def get_rows11(file_name):
								book = xlrd.open_workbook(file_name)
								sh = book.sheet_by_name(sheet_name)
								book1 = Workbook()
								rb = xlrd.open_workbook(file_name)
								wb = xl_copy(rb)
								sheet = wb.get_sheet(sheet_name)
								total_cols = sh.ncols
								total_rows = sh.nrows
								data_array11 = []
								

								for j in range(1, total_rows): #for rows
									data_array_elemnts = {}
									k = 0	
									x = 0 
									for i in cols_in_list:
										column_out = findexcel_column(i,file_name,sheet_name)
										column_out123=column_out

										if column_out123 != None:
											x = x + 1

										else:
											data = {}
											data['status_code'] = "401"
											data['status'] = "Please provide existing columns name"
											print(json.dumps(data))
											sys.exit()	
											
									excel_lists2_arr = []		
									for i in cols_in_list:
										if x == len(cols_in_list):
											column_in = col_to_num(i) -1
											
											findexcel_column_num = findexcel_column1(i)
											excel_list21= sh.cell_value(j,int(findexcel_column_num))

											excel_list21 =  excel_list21
											data_array_elemnts[str(i)] = excel_list21

											data_array_elemnts['row_index'] = str(j)

										else:
											data = {}
											data['status_code'] = "401"
											data['status'] = "Please provide existing columns name"
											print(json.dumps(data))
											sys.exit()		
											
									data_array11.append(data_array_elemnts)	
								
								
								
								if len(range_text2) ==2:
									var1_range = int(range_text2[0])
									var2_range = int(range_text2[1]) + 1 
									range_text = range(var1_range,var2_range)
									range_text[:] = [x - 1 for x in range_text]
									data_array11 = [data_array11[i] for i in range_text]
								
								elif len(range_text2) ==1:	
									var1_range = int(range_text2[0]) 
									if var1_range > 0:
										range_text = var1_range - 1
										data_array11 = data_array11[range_text]
									else:
										data = {}
										data['status_code'] = "401"
										data['status'] = "Please provide proper range"
										print(json.dumps(data))
										pass
										sys.exit()
								#print(range	
								else:
									data = {}
									data['status_code'] = "401"
									data['status'] = "Please provide proper range"
									print(json.dumps(data))
									pass
									sys.exit()

									
							
								data = {}
								data['status_code'] = "200 OK"
								data['data'] = data_array11
								data = json.dumps(data)
								return data
							
							
							file_data_rowsget = get_rows11(file_name)
							print(file_data_rowsget)
					
						
						else:
							data = {}
							data['status_code'] = "401"
							data['status'] = "Please provide proper syntax"
							print(json.dumps(data))
							pass		
					
					
					elif file_extension in (''):
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass
						
					else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file extension"
						print(json.dumps(data))
						pass	
						
				else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass			
						
			else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide Existing file name"
						print(json.dumps(data))
						pass		
		
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e)
			print(json.dumps(data))
			pass
			
			
	elif commandKey == "GetColumnHeader":
		try :
			if os.path.exists(sys.argv[2]): 				
				filename, file_extension = os.path.splitext(sys.argv[2])			
				if filename[-1] not in (' '):
					if sys.argv[2].endswith('xls') or sys.argv[2].endswith('xlsx'):
						file_name =  sys.argv[2]	
						sheet_name =  sys.argv[3]
						
						#number to column generation header in excel sheet
						def colnum_string(n):
							string = ""
							while n > 0:
								n, remainder = divmod(n - 1, 26)
								string = chr(65 + remainder) + string
							return string
							
						def get_row(file_name):
							book = xlrd.open_workbook(file_name)
							sh = book.sheet_by_name(sheet_name)
							book1 = Workbook()
							#sheet1 = book1.add_sheet('test', cell_overwrite_ok=True)
							rb = xlrd.open_workbook(file_name)
							wb = xl_copy(rb)
							sheet = wb.get_sheet(sheet_name)
							total_cols = sh.ncols + 1
							#print(column_in)
							column_in = 0 
							data = {}
							data_array = []
							for i in range(1, total_cols):
								data_array.append(colnum_string(i))
							data = {}
							data['status_code'] = "200 OK"
							data['data'] = data_array
							data = json.dumps(data)
							return data
						
						file_data_row = get_row(file_name)
						print(file_data_row)	
					
					elif file_extension in (''):
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass
						
					else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file extension"
						print(json.dumps(data))
						pass	
						
				else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass			
						
			else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide Existing file name"
						print(json.dumps(data))
						pass		
		
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e)
			print(json.dumps(data))
			pass
			

	elif commandKey == "GetFirstHeaderRow":
		try :
			if os.path.exists(sys.argv[2]): 				
				filename, file_extension = os.path.splitext(sys.argv[2])			
				if filename[-1] not in (' '):
					if sys.argv[2].endswith('xls') or sys.argv[2].endswith('xlsx'):
						file_name =  sys.argv[2]	
						sheet_name =  sys.argv[3]

						def get_row(file_name):
							book = xlrd.open_workbook(file_name)
							sh = book.sheet_by_name(sheet_name)
							book1 = Workbook()
							#sheet1 = book1.add_sheet('test', cell_overwrite_ok=True)
							rb = xlrd.open_workbook(file_name)
							wb = xl_copy(rb)
							sheet = wb.get_sheet(sheet_name)
							total_cols = sh.ncols
							#print(sys.argv[5])
							#print(column_in)
							column_in = 0 
							data = {}
							data_array = []
							for i in range(0, total_cols):
								value = sh.cell_value(column_in,i)
								data_array.append(value)
								
							data = {}
							data['status_code'] = "200 OK"
							data['data'] = data_array
							data = json.dumps(data)
							return data
						
						file_data_row = get_row(file_name)
						print(file_data_row)	
					
					elif file_extension in (''):
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass
						
					else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file extension"
						print(json.dumps(data))
						pass	
						
				else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass			
						
			else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide Existing file name"
						print(json.dumps(data))
						pass		
		
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e)
			print(json.dumps(data))
			pass	
			
			
		
	else:
		print("Please enter proper Command Key")

	#saving excel file	

except Exception as e:
	data = {}
	data['status_code'] = "401"
	data['status'] = str(e) 
	print(json.dumps(data))
	#data = json.dumps(data)
	#print(json.dumps(data))
	pass

	

	
	

