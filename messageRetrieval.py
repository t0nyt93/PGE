'''
Author: Anthony Tyrrell
Date: 4/8/2015
Written for Portland General Electric
Purpose:
--- This script uses the win32com API provided by Microsoft to interact with Microsoft Outlook
--- Documentation for this API can be found at https://msdn.microsoft.com/en-us/library/microsoft.office.interop.outlook(v=office.14).aspx
--- Be warned...It is confusing and does not have direct code examples for the Python language.
'''
import win32com
import win32com.client
import string
import os


def findFolder(folderName,iterObject):
	try:
		lowerAccount = iterObject.Folders
		for x in lowerAccount:
			if x.Name == folderName:
				print 'found it %s'%x.Name
				objective = x
				return objective
		return None
	except Exception as error:
		print "Looks like we had an issue accessing ITSecOps"
		print (error)
		return None

def main():

	''' 
	the variable outlook is defind by win32com as an _Application object...available methods can be found by 
	printing dir(outlook). Properties of this object are available in the documentation provided by MSDN(Link below)
	https://msdn.microsoft.com/en-us/library/microsoft.office.interop.outlook._application_members(v=office.14).aspx
	'''
	outlook=win32com.client.Dispatch("Outlook.Application")
	'''
	the variable ons is a _Namespace object (ons = outlook namespace) that refers to the user account being used within Outlook. 
	The link to it's methods and properties is included below. 
	https://msdn.microsoft.com/en-us/library/microsoft.office.interop.outlook._namespace_members(v=office.14).aspx
	'''
	ons = outlook.GetNamespace("MAPI")
	
	one = 'IT.SECOPS@pgn.com'
	two = "IT Security Operations"
	
	Folder1 = findFolder(two,ons)
	Folder2 = findFolder('Inbox',Folder1)
	Folder3 = findFolder('McAfee',Folder2)
	Folder4 = findFolder('Virus Research- Avertlabs',Folder3)
		
	new_folder = '(processed)Virus Research-Avertlabs'
	messages = Folder4.Items

	#Iterate through the messages contained within our subfolder
	for xx in messages:
		try:
		#This makes the objects properties comparable for our language
			temp = xx.Subject.encode('ascii','ignore')
	#Compare each e-mails subject to see if it contains info that we are interested in
			if temp[:48] == 'Submission through GetSusp (Reference WorkItemID':
				eID = temp[49:]
		#If the e-mail pattern matches. We're going to write it's body to a text file. 
				if eID != '':
					eID = int(eID[:8])
				#Naming convention for our files, text+machine name
					filename = '/text' + str(eID) + '.txt'

					three = "C:\Users\e04675\Desktop\serverScripts\PGE\emailBodies"
					four = "C:\Users\itsecops\Desktop\serverScripts\PGE\emailBodies"
			#Path to where we're going to write it
					fileDirect =  three
					final = fileDirect+filename
					printStr = xx.Body
			#Open the file, write to it, and close it. Simple Enough
					fd = open(final,'w+')
					fd.write(printStr)
				fd.close()
		except Exception as err:
			print "Error saving file"
			print err		

if __name__ == "__main__":
	main()