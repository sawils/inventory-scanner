    # from gs1_parser import gs1_decoder
import os
import pandas as pd
import tabulate
import barcode_scanner 
from _settings import CUSTOMERS, ACTIONS
from pprint import pprint
from pathlib import Path
import win32com.client as win32
from datetime import datetime
import xlsxwriter

master_dict = {}
datastream = []
customers = []
barcodes = [] 

#Initialize
input_str = str(input('>> Customer: ')).lower()
datastream.append(input_str)
customers.append(input_str)
master_dict[input_str] = barcodes

def pretty_customers(customer):
    for c in CUSTOMERS: 
        if str(c.uuid) == customer:
            return c.name
    return customer

def is_barcode(barcode):
    return len(barcode) >= 46


while True:
    last_inp = datastream[-1:][0]
    curr_customer = pretty_customers(customers[-1:][0])

    if last_inp.lower() == ACTIONS['exit']: break
    elif last_inp.lower() == ACTIONS['next']:
        # uses pretty_name
        # old_customer = curr_customer
        # master_dict[old_customer] = barcodes
        barcodes = []
        new_customer = str(input('>> Please input your customer: '))
        master_dict[new_customer] = barcodes
        customers.append(new_customer)
        datastream.append(new_customer)
        continue
    else:
        inp = str(input('>> Insert barcode for %s: ' % curr_customer))
        if is_barcode(inp):
            barcodes.append(inp)
        datastream.append(inp)
    # os.system('clear')
    os.system('cls') #windows
    for customer in master_dict:
        print('===============================================')
        print('For Customer %s' % pretty_customers(customer))
        output = barcode_scanner.format_barcodes(master_dict[customer])
        ## Create dataframe and filter by critical_headers
        # print(output)
        critical_headers = ['GTIN-14',
                            'Product Net Weight',]
        out_df = pd.DataFrame(output)[critical_headers]
        out_df.set_index("GTIN-14", inplace=True)
        '''Splits into Tables by Unique GTIN-14'''
        for gtin in out_df.index.unique():
            row = out_df[out_df.index == gtin]
            print(tabulate.tabulate(row,critical_headers,tablefmt='grid'))
            print("Number of Items: %d" % len(row))
            print("Total Weight: %s lbs" % row['Product Net Weight'].sum())

    pprint("Customers: ")
    pprint(customers)
    pprint("Barcodes: ")
    pprint(barcodes)
    pprint("Datastream: ")
    pprint(datastream)
    pprint("Master: ")
    pprint(master_dict)
print("You have exited. Goodbye!")

#print(datastream)
# Create File
# date = datetime.today().strftime("%Y%m%d-%H%M%S")
# filename = date + "_mm_inv.xlsx"
# out_file = Path.cwd()/filename
# # out_df.to_excel(out_file)  # Brings dataframe to excel 
# writer = pd.ExcelWriter(out_file, engine='xlsxwriter')
# out_df.to_excel(writer, sheet_name='final')
# datastream_df = pd.DataFrame(datastream)
# datastream_df.to_excel(writer, sheet_name='datastream')
# writer.save()
# # Open up Excel and make it visible
# excel = win32.gencache.EnsureDispatch('Excel.Application')
# excel.Visible = True

# # Open up the file
# excel.Workbooks.Open(out_file)

# # Wait before closing it
# _ = input("Press enter to close Excel")
# excel.Application.Quit() 