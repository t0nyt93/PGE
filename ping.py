import os
import sys
import subprocess

proc = subprocess.Popen(["ping", sys.argv[1]], stdout=subprocess.PIPE, shell=True)
(out, err) = proc.communicate()

out =  out.split()

print out[1] + ' ' + out[2]