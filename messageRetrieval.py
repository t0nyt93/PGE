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

'''
upperAccount refers to a MAPIFolder Object which  is a child of ons. This MAPIFolder object refers to the e-mail account being accessed within the namespace
In this case my Outlook consists of my PGE e-mail Anthony.Tyrrell@pgn.com, and the IT Security Operations E-mail. So we iterate through Folder until we
find the account we want to look inside.
Information on methods and properties of MAPIFolder objects can be found at the link below. 

https://msdn.microsoft.com/en-us/library/microsoft.office.interop.outlook.mapifolder_members(v=office.14).aspx

try:
	upperAccount = ons.Folders
	for x in upperAccount:
		upperAccount = ons.Folders.GetNext()
		if x.Name == "IT Security Operations":
			itAcc = x
			break
except Exception as error:
	print "Looks like we had an issue accessing ITSecOps"
	print (error)


'''
'''
After we have retrieved the correct account, we look for the mail folder we are interested in. In this case it's the Inbox, which is another MAPIFolder instance.
'''
try:
	secopsInbox = ons.Folders
	# Loop through available folders until we match
	for y in secopsInbox:
		if y.Name == "Inbox":
			secInbox = y
			break
except Exception as e:
	print "Couldn't find Inbox"
	print e
#Once We have inbox, we need to go through the subfolders there. 
''' More Mapi Folder Instances '''
try:
	for z in secInbox.Folders:
		if z.Name == 'McAfee':
			subFold = z
			break
	subFold2 = subFold.Folders
except Exception as er:
	print "Had trouble navigating folders"
	print er

'''
In this iteration, we once again (for the last time) iterate through the subfolders of a MAPIFolder object.
'''
try:
	for t in subFold2:
		if t.Name == 'Virus Research- Avertlabs':
			myFolder = t
			break
except Exception as error:
	print "More Trouble navigating folders"
	print error


'''
Done with MAPI_folder Instances, now we are dealing with _MailItem objects (e-mails). Documentation on the _MailItem object can be found below.
https://msdn.microsoft.com/en-us/library/microsoft.office.interop.outlook._mailitem_members(v=office.14).aspx
'''

messages = myFolder.Items

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
			#Path to where we're going to write it
				fileDirect = "C:\Users\e04675\Desktop\Scripts\emailBodies" 
				final = fileDirect+filename
				printStr = xx.Body
			#Open the file, write to it, and close it. Simple Enough
				fd = open(final,'w+')
				fd.write(printStr)
			fd.close()
	except Exception as err:
		print "Error saving file"
		print err		