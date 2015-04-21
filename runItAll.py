import subprocess
import os
import sys
import time

print "Resetting formats and files to avoid repetitions"


proc0 = subprocess.Popen(["python","reset.py"],stdout = subprocess.PIPE, shell=True)
(out0,err) = proc0.communicate()
out0 = out0.split()
time.sleep(1)
#### First step is to run readStuff.py
print "Running messageRetrieval.py"
time.sleep(1)

print "Don't forget to give the script access to Outlook"
try:
	proc1 = subprocess.Popen(["python","messageRetrieval.py"], stdout=subprocess.PIPE, shell=True)
	(out1, err) = proc1.communicate()
	out1 = out1.split()
except Exception as e1:
	print "Error running process 1"
	print e1.args


print "If there was no error running process 1, messageRetrieval.py ran successfully"
time.sleep(1)
#### Second Step is to run md5grabber.py
print "Running messageParser.py"
try:
	proc2 = subprocess.Popen(["python","messageParser.py"], stdout=subprocess.PIPE, shell=True)
	(out2, err) = proc2.communicate()
	out2 = out2.split()
except Exception as e2:
	print "Error running process 2"
	print e2

print "If there was no error running process 2, messageParser.py ran successfully"

##### Third Step is to run md5test.py
print "Running virusQuery.py"
time.sleep(1)
try:
	proc3 = subprocess.Popen(["python","virusQuery.py"], stdout=subprocess.PIPE)
	lines_iter = iter(proc3.stdout.readline, b"")
	for line in lines_iter:
		print line

except Exception as e3:
	print "Error running process 3"
	print e3

print "If there was no error running process 3, virusQuery.py ran successfully"

##### Fourth Step is to run reset.py
print "Resetting everything to empty"
try:
	proc4 = subprocess.Popen(["python","reset.py"], stdout=subprocess.PIPE, shell=True)
	(out4, err) = proc4.communicate()
	out4 = out4.split()
except Exception as e4:
	print "Error running process 4"
	print e4
print "Processes completed."