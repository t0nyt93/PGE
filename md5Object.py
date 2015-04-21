class myMD5:
	"""This class is designed to hold output from virusTotal JSON objects. """
	scan_id = ''
	sha1 = ''
	resource =''
	responseCode =''
	scanDate =''
	permaLink=''
	verboseMsg=''
	scanFinished=''
	sha256=''
	positives= 0
	total=''
	md5 = ''
	detectedBy = []
	machineName = ''
	fileName = ''
class emailParse:
	'''this class is designed to pull information contained in emails for getsusp'''
	machineID= ''
	foundFiles = []
	foundMD5 = []