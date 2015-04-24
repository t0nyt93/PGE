import requests

headers = {'Version': '3.0','SEC': 'b1b563dc-aade-4512-9267-0e8269f7abe4','Accept': 'application/json','Allow-Experimental': 'true'}
r = requests.get('https://10.80.192.2/api/reference_data/tables/virustotal_correlated_ips',auth = ('e04675',"Bachelor2015"),headers= headers)
print r.status_code