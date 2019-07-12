import sys, os
sys.path.append(os.path.join(os.path.dirname(__file__), '..'))
import barcode_scanner as bs
from pprint import pprint 

dirname = os.path.dirname(__file__)
filename = os.path.join(dirname, 'data/all_barcodes.txt')

input_list = []
barcode_issues = []
with open(filename, 'r', encoding='utf-8') as f:
	contents = f.read()
	input_list = contents.splitlines()

# Remove Duplicates
input_list = list(dict.fromkeys(input_list))	
sys.stdout = open('file','w')
for code in input_list:
	barcode_dict = bs.gs1_decoder(code)
	if(not(('GTIN-14' in barcode_dict)&('Product Net Weight' in barcode_dict))):
		barcode_issues.append(code)

pprint(barcode_issues)
print('\n')
