#!/Users/wilsonsa/anaconda2/bin/python
'''
Main file to parse and display gs1-128 data
'''
#Libraries
import sys
import tabulate
import pandas as pd
from _settings import APP_IDS
#
# Ex:
# Input: 0100070919010926320200449413190122214107381404
# Output: {
#           GTIN-14:                 00070919010926
#           Product Net Weight (lb): 44.94 lbs
#           Packaging Date:          January 22, 2019
#           Serial Number:           4107381404
#          }
def gs1_decoder(string):
    '''Decodes one individual barcode.'''
    output = {}
    while string:
        for app_id in APP_IDS:
            regex = app_id.get_RE()
            result = regex.search(string)
            if result:
                match = result.group()
                output.update(app_id.get_description(match))
                string = string[len(match):]
                break # Restart the search in case order of data is different
        else:
            break # Here's your infinite safety :)
    return output

def format_barcodes(barcode):
    '''Format and prints barcodes
    Input: ['019932793300799431020010091318100921450056032605', 
            '019932793300799431020010091318100921450056032605', 
            '019932793300799431020010091318100921450056032605']
    Output: 
            +----------------+----------------------+
            |        GTIN-14 | Product Net Weight   |
            +================+======================+
            | 99327933007994 | 22.24 lbs            |
            +----------------+----------------------+
            | 99327933007994 | 22.24 lbs            |
            +----------------+----------------------+
            | 99327933007994 | 22.24 lbs            |
            +----------------+----------------------+
                Number of Items: 3'''
    ## Run settings
    default_len=46
    critical_headers = ['GTIN-14',
                        'Product Net Weight',]
    
    ## Parse barcodes using the default length
    output = [gs1_decoder(barcode) ]

    return output