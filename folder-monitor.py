#!/usr/bin/python
# -*- coding: latin-1 -*-
import os
import os
import sys
import filemapper as fm
import shutil

try:
	argv1 = sys.argv[1]
	#argv2 =  sys.argv[2]
	#argv3 =  sys.argv[3]
except Exception as e:
	print(e)
	print("please enter proper arguments")
	
folderpath = argv1	
	#filename = argv2

def readfile(filename):
	text_file = open(filename,'r')
	#open the file
	#text_file = open('/Users/pankaj/abc.txt','r')

	#get the list of line
	line_list = text_file.readlines();

	#for each line from the list, print the line
	for line in line_list:
		print(line)

	text_file.close() #don't forget to close the file
	return  text_file
		
filecontent = ""
for root, dirs, files in os.walk(folderpath, topdown=False):
    for name in files:
        print(os.path.join(root, name))
        if name.split('.')[1] == "txt":
			filename = folderpath+"/"+name
			filecontent = readfile(filename)
			print(filecontent)
	print("--------------")	
    for name in dirs:
        print(os.path.join(root, name))
        #stuff
		
	