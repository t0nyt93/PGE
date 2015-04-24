'''
Author: Anthony Tyrrell
Date: 4/8/2015
Written for Portland General Electric
Purpose: This script takes a directory of text files as input and parses them, writing the desired information into a .xls file that will then
		 be used and interpreted by md5test.py
'''
import re
import md5Object
import csv
import xlwt
import os
from xlrd import *
from xlutils.copy import *

########## Test Variables and Indexers
count = 0
foundFlag = 0
#######################################

########## Excel master sheet Name
myTable = 'MD5SCAN.xls'
previous_dir = os.getcwd()
#########

######### Change our working directory to the location of the e-mail bodies we will be parsing. 
os.chdir('emailBodies')
######### For all of the files in this directory (Should only be .txts in a certain format) iterate through them
for f in os.listdir('.'):
######### Where f is the name of our current file	
	index = 0
	sheetIndex = 0
	
	### Custom Object as defined in md5Object.py
	myObject = md5Object.emailParse()
	
	### In case we are still in the wrong directory after an iteration or two change back to correct one
	if os.getcwd() == previous_dir:
		os.chdir('emailBodies')

	### Open the current .txt file for reading
	fd = open(f,'r+')
	fd.read(498)
	
	#Get our Machines Name
	machineName = fd.read(10)
	#Skip some more junk (300 bits) so that we can find the pattern
	fd.read(300)

	#Read the rest of the file in as a string that will be parsed
	parseMe = fd.read()

	######### from regular expressions module, used for pattern matching 
	for m in re.findall("[|](.*)[|](.*)[|](.*)[|](.*)[|](.*)[|]",parseMe):
		myObject.foundFiles.append(m[0])
		myObject.foundMD5.append(m[1])
	#########
	fd.close()
	######## Strip blank spaces from the Machine Name in case it's shorter than what we read.
	machineName.strip()
	machineName = ''.join(machineName.split())
	##### Now that we have filenames and MD5's lets go back up a directory and write them to our central .xls file
	os.chdir(previous_dir)
	
	##### Open our workbook MD5SCAN.xls
	rb = open_workbook(myTable)
	
	##### make a copy for editing
	wb = copy(rb)
	##### List of the current sheets in so that we can append if we already have a record of it
	existingSheets = rb.sheet_names()
	####################################################################################################
	#Check to make sure that this machine isn't already in our file
	for x in existingSheets:
		#### If it is in our system, write to existing sheet instead of creating a new one
		if x == machineName:
			#### Gets the sheet based on how many times we have gone through this process
			currSheet = wb.get_sheet(sheetIndex)
			foundFlag = 1
			t_sheet = rb.sheet_by_index(sheetIndex)
			myRows = t_sheet.nrows

			#Writing File Names to Worksheet
			for n in myObject.foundFiles:
				currSheet.write((index+myRows),0,n)
				index+=1
			index = 0	
			
			#Writing MD5's to worksheet
			for k in myObject.foundMD5:
				currSheet.write((index+myRows),1,k)
				index+=1
			index = 0	

		sheetIndex+=1	

	#################################################################################################
	# If we don't have a record of the machine already, we have to append a new sheet to the file
	if foundFlag == 0:
		myNewSheet = machineName
		new_sheet = wb.add_sheet(myNewSheet,cell_overwrite_ok = True)
	#Writing File Names to Worksheet
		index = 0
		for n in myObject.foundFiles:
			new_sheet.write(index,0,n)
			index+=1
		index = 0	
			
	#Writing MD5's to worksheet
		for k in myObject.foundMD5:
			
			new_sheet.write(index,1,k)
			index+=1
		index = 0	
	#################################################################################################

	### Clearing our entries for the past machine to make sure that we don't rewrite information
	del(myObject.foundFiles[:])
	del(myObject.foundMD5[:])
	count+=1
	foundFlag = 0
	#Save the new workbook over the old one
	wb.save(myTable)