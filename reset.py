import os
from xlutils.copy import *
from xlrd import *
from xlwt import *
homeDir = os.getcwd()
from sys import *
import time
################ CLEAR E_MAIL BODIES DIRECTORY ##########################

os.chdir('emailBodies')
myPath = os.path.dirname(os.path.abspath(__file__))
for f in os.listdir('.'):
	os.remove(myPath + '/' +f)

os.chdir(homeDir)


os.chdir('leefLogs')
myPath = os.path.dirname(os.path.abspath(__file__))
for f in os.listdir('.'):
	print"removing %s"%f
	os.remove(myPath + '/' + f)

myTable = 'MD5SCAN.xls'
os.chdir(homeDir)

while True:
		try:
			workbook = open_workbook(myTable)
			break
		except Exception as e:
			print e

print (workbook.sheet_by_index(0))

new_book = Workbook()
new_sheet = 'PLACEHOLDER'
new_book.add_sheet(new_sheet)

while True:
	try:
		new_book.save(myTable)
		break
	except Exception as err:
		print err
		stdout.flush()
		print "\r Please close the Excel file so we can save"
		stdout.flush()
		time.sleep(1)


