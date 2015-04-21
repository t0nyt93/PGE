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
#############################################
###         Constants and Index variables
first = 1
counter = 0
index = 0
MD5_URL = "https://www.virustotal.com/vtapi/v2/file/report"
MY_KEY = "df25ff20c9a41fb8752ec252a911688c684a40c0f8b8b918ff8284a7eb9698a8"
LEEF_HEADER = 'LEEF:1.0|PGE|SECOPS_CUSTOM|0.2|1|\t'
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
			
			# each md5 gets its own object
			resultList.append(md5Object.myMD5())
			############################################################
			# these calls come directly from the VirusTotal API
			
			
			try :
				parameters = {"resource":x,"apikey":MY_KEY}
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
				parameters = {"resource":x,"apikey":MY_KEY}
				data = urllib.urlencode(parameters)
				req = urllib2.Request(MD5_URL,data)
				response = urllib2.urlopen(req)
				json_data = json.loads(response.read())	
				
			##############################################################
			# VirusTotal has a record of the MD5 so lets looks at what it has to say
			
			if json_data['response_code'] == 1:
				resultList[counter].machineName = sheet_name
				resultList[counter].fileName = curSheet.row_values(index)[0]	
				resultList[counter].scanId = json_data['scan_id']
				resultList[counter].sha1 = json_data['sha1']
				resultList[counter].resource = json_data['resource']
				resultList[counter].responseCode = json_data['response_code']
				resultList[counter].scanDate = json_data['scan_date']
				resultList[counter].permaLink = json_data['permalink']
				resultList[counter].verboseMsg = json_data['verbose_msg']
				resultList[counter].sha256 = json_data['sha256']
				resultList[counter].positives = int(json_data['positives'])
				resultList[counter].total = json_data['total']
				resultList[counter].md5 = json_data['md5']
				#HolySnap is the string containing all of the virus protection systems the md5 has been crossed against.  
				######################################################################################################################
				holySnap = json_data['scans']																						 #				
				#Iterate through these systems and find out which of them have flagged our md5 as malicious                          #
				for z in holySnap:                                                                                                   #
					#If it has been flagged, add the host name to our data collection                                                #
					if json_data['scans'][z]['detected'] is True:                                                                    #
						resultList[counter].detectedBy.append(z)                                                                     #
				#######################################################################################################################	
				# If the md5 has been flagged as malicious, we probably need to report it	
				if resultList[counter].positives > 0:
					# Open a subproccess that attempts to retrieve more information about the machine this md5 was found on
					proc = subprocess.Popen(["python","ping.py", resultList[counter].machineName], stdout=subprocess.PIPE, shell=True)
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
					# Now is the time to write out gathered information into something that QRADAR can understand and document
					os.chdir('leefLogs')
					# This sets everything to the correct LEEF format
					keyDomain = 'domain='+compDomain+'\t'
					keyIP = 'src=' + compIP+'\t'
					keyFile = 'filename='+resultList[counter].fileName+'\t'
					keyURL = 'url='+resultList[counter].permaLink
					logName = 'LEEF'+resultList[counter].machineName + '.log'
					#### Open our logfile and begin writing to it
					fd = open(logName,'w+')
					fd.write(LEEF_HEADER)
					fd.write(keyDomain)
					fd.write(keyIP)
					fd.write(keyFile)
					fd.write(keyURL)
					fd.close()
					# After we have written to the correct file traverse up a directory so that we can read from the excel file again
					os.chdir(previousDir)
			index +=1	
			counter += 1
				
		first = 0
				

	