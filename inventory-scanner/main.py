# Do I need some sort of #! thingy here?
import os
import pandas as pd
from _settings import CUSTOMERS, ACTIONS
import barcode_scanner as bs
import to_xl as xl
from pprint import pprint

master_dict = {'cust_id':[],
               'cust_name':[],
               'gtin':[],
               'gtin_name':[],
               'weight':[],}
customers = []
barcodes = []
datastream = []
error_barcodes = []
def pretty_customers(customer):
    for c in CUSTOMERS: 
        if str(c.uuid) == customer:
            return c.name
    return customer

'''
this function is a product of poor code design. 
1. You're mpn dict is calling excel and creating a new dict every time
2. You did this because you wanted to avoid creating a global variable 'MPN_DICT'
What's the solution here?
'''
def get_gtin_name(gtin):
    mpn_dict = get_mpn_dict()
    if gtin in mpn_dict:
        return mpn_dict[gtin]
    else:
        return 'No Name Found!'

def get_mpn_dict():
    curpath = os.path.dirname(__file__)
    fn = os.path.abspath('C:\\Users\\Julie\\Documents\\dev\\inventory-scanner-git\\inventory-scanner\\docs\\master_mpn_list.xlsx')
    # fn = os.path.relpath('docs\\master_mpn_list.xlsx', curpath)
    df = pd.read_excel(fn, converters={'MPN': lambda x: str(x)}, index=False) #Keeps leading zeros for MPN
    mpn_dict = df.set_index('MPN').to_dict()['Description']
    return mpn_dict

def is_barcode(input):
    return len(input) >= 46

def is_action(input):
    return input in ACTIONS.values()

def format_data(dictionary):
    df =  pd.DataFrame(dictionary)    
    df = df[['cust_id','cust_name','gtin','gtin_name','weight']]
    print(df)
    return df
def main():
    #Initialize
    input_str = str(input('>> Customer: ')).lower()
    datastream.append(input_str)
    customers.append(input_str)

    while True:
        last_inp = datastream[-1:][0]
        if last_inp.lower() == ACTIONS['exit']: 
            datastream.append('exit Logged')
            break
        elif last_inp.lower() == ACTIONS['next']:
            datastream.append('next Logged')
            new_customer = str(input('>> Please input your customer: '))
            customers.append(new_customer)
        else:
            last_customer = pretty_customers(customers[-1])
            inp = str(input('>> Insert barcode for %s: '%last_customer))
            if is_barcode(inp):
                barcodes.append(inp)
                output = bs.gs1_decoder(barcodes[-1])
                master_dict['cust_id'].append(customers[-1:][0])
                master_dict['cust_name'].append(pretty_customers(customers[-1:][0]))
                master_dict['gtin'].append(output['GTIN-14'])
                master_dict['gtin_name'].append(get_gtin_name(output['GTIN-14']))
                master_dict['weight'].append(output['Product Net Weight'])

            elif is_action(inp):
                print('Action --%s-- occured'%inp)
            else:
                error_barcodes.append(inp)
                print('An error occurred! Barcode invalid. Add error handling!')
                continue
            datastream.append(inp)
    print("You have exited. Goodbye!")
    print("Error Barcodes: \n")
    pprint(error_barcodes)
    
    # Convert to Dataframe & Group It
    df = format_data(master_dict)
    ds = pd.DataFrame(datastream)
    eb = pd.DataFrame(error_barcodes)
    xl.excelify(df, ds, eb)
    
if __name__ == "__main__":
    main()
