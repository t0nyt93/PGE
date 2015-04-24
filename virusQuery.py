'''
Author: Anthony Tyrrell
Date: 4/8/2015
Written for Portland General Electric
Purpose: This Script is used to communicate with VirusTotal and detect/report Malicious md5 values
		For more information regarding syntax and operations, read up on the VirusTotal Public API at ->
	https://www.virustotal.com/en/documentation/public-api/
'''
import json, requests, urllib, urllib2
from xlrd import *
from xlutils.copy import *
import xlwt
import time
import re
import md5Object
import subprocess
import os
from sys import *

MY_KEY = "df25ff20c9a41fb8752ec252a911688c684a40c0f8b8b918ff8284a7eb9698a8"
MD5_URL = "https://www.virustotal.com/vtapi/v2/file/report"
LEEF_HEADER = 'LEEF:1.0|PGE|SECOPS_CUSTOM|0.2|1|\t'

def queryVT(currentmd5,myObject,sheet,fileName):
	try :
		parameters = {"resource":currentmd5,"apikey":MY_KEY}
		data = urllib.urlencode(parameters)
		req = urllib2.Request(MD5_URL,data)
		response = urllib2.urlopen(req)
		print "-- %s --"%x
		json_data = json.loads(response.read())
				
		#############################################################
		#VirusTotal Responds with a JSON Data object that contains all of our information
		#this is stored in another custom object as defined in md5Object.py
	except Exception as err:
		print "Exceeded query limit: Waiting..."
		for i in range(60,-1,-1):
			time.sleep(1)
		parameters = {"resource":currentmd5,"apikey":MY_KEY}
		data = urllib.urlencode(parameters)
		req = urllib2.Request(MD5_URL,data)
		response = urllib2.urlopen(req)
		json_data = json.loads(response.read())	
		
		if json_data['response_code'] == 1:
			myObject.machineName = sheet
			myObject.fileName = fileName	
			myObject.scanId = json_data['scan_id']
			myObject.sha1 = json_data['sha1']
			myObject.resource = json_data['resource']
			myObject.responseCode = json_data['response_code']
			myObject.scanDate = json_data['scan_date']
			myObject.permaLink = json_data['permalink']
			myObject.verboseMsg = json_data['verbose_msg']
			myObject.sha256 = json_data['sha256']
			myObject.positives = int(json_data['positives'])
			myObject.total = json_data['total']
			myObject.md5 = json_data['md5']
			holySnap = json_data['scans']																						 #				
			#Iterate through these systems and find out which of them have flagged our md5 as malicious                          #
			for z in holySnap:                                                                                                   #
				#If it has been flagged, add the host name to our data collection                                                #
				if json_data['scans'][z]['detected'] is True:                                                                    #
					myObject.detectedBy.append(z)

			if myObject.positives > 0:
				myIP = pingMachine(myObject.machineName)
				print myIP
				myObject.ip = myIP
				return myObject
	return None


def pingMachine(machine_name):
	proc = subprocess.Popen(["python","ping.py",machine_name], stdout=subprocess.PIPE, shell=True)
	(out, err) = proc.communicate()
	out = out.split()

	if out[0] ==  'request':
		compIP= 'Not Found'
		compDomain = 'Not Found'
# Information retrieved about machine
	else:
		compDomain = out[0]
		compIP = out[1]
		compIP=compIP.strip("[]")

	return compIP


def writeLEEF(myContent):
	previousDir = os.getcwd()
	os.chdir('leefLogs')
	keyIP = 'src=' +myContent.ip+'\t'
	keyFile = 'filename='+myContent.fileName+'\t'
	keyURL = 'url='+myContent.permaLink
	logName = 'LEEF'+myContent.machineName + '.log'
	fd = open(logName,'w+')
	fd.write(LEEF_HEADER)
	fd.write(keyURL)
	fd.write(keyFile)
	fd.write(keyIP)
	fd.close()
	os.chdir(previousDir)


def main():
	#############################################
	###         Constants and Index variables
	index = 0

	#############################################	
	myTable = 'MD5SCAN.xls'
	resultList= []
	previousDir = os.getcwd()
	#### Open our excel file for reading.
	rb = open_workbook(myTable)
	### Every sheet_name correlates to the files that were found on a single machine
	#### so for every machine that had suspicious files, analyze the md5's found on it
	for sheet_name in rb.sheet_names():
		### Skip the first sheet everytime so that our excel file won't disappear
		if sheet_name != 'PLACEHOLDER':
			curSheet = rb.sheet_by_name(sheet_name)	
			index = 0
			result = 0
			#### md5s are located in the second column (zero indexed) while file names are located in the first
			#### so for every md5 located on the current sheet
			for row in curSheet.col(1):
				### VirusTotal publically only allows 4 queries per minute, so requesting more than that will actually result in an error 
				# Grab our md5 from the excel sheet
				x = curSheet.row_values(index)[1]
				newObject = md5Object.myMD5()
				namefile = curSheet.row_values(index)[0]
				testVar = queryVT(x,newObject,sheet_name,namefile)
				
				if testVar == None :
					continue
				else:
					resultList.append(testVar)
					writeLEEF(testVar)
				
			
				index +=1	

if __name__ == "__main__":
	main()

	