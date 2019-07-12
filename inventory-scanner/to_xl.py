from datetime import datetime
import pandas as pd
import win32com.client as win32
from pathlib import Path
import xlsxwriter
import time
# Purpose is to get a master dictionary and group by correctly. 

def get_date():
    dow = ['Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday']
    today = datetime.today()
    today_date = today.strftime("%m/%d/%Y")
    today_day = dow[today.weekday()]
    print(today_date)
    print(today_day)
    return str(today_day+", "+today_date)
# Add Excel specific formatting to the workbook
def format_excel_table(writer, table_coords, title_coords):
    # Get the workbook and the summary sheet so we can add the formatting
    workbook = writer.book
    worksheet = writer.sheets['final'] #pandas version

    # Formatting columns
    num_fmt = workbook.add_format({'num_format': 0, 'align': 'center'})
    gtin_fmt = workbook.add_format({'num_format': 0, 'align': 'right'})
    float_fmt = workbook.add_format({'num_format': '0.00', 'align': 'center'})
    title_fmt = workbook.add_format({'bold':True, 'underline': True, 'font_size': 18})   
    
    worksheet.write(0, 0, 'DATE: %s' %get_date(), title_fmt)
    # worksheet.set_column('A:A', 5)
    # worksheet.set_column('B:B', 8, num_fmt)
    # worksheet.set_column('C:C', 30, num_fmt)
    # worksheet.set_column('D:D', 16, gtin_fmt)
    # worksheet.set_column('E:E', 25, num_fmt)
    # worksheet.set_column('F:F', 10, float_fmt)
    worksheet.set_column('A:A', 8, num_fmt)
    worksheet.set_column('B:B', 25, num_fmt)
    worksheet.set_column('C:C', 16, gtin_fmt)
    worksheet.set_column('D:D', 25, num_fmt)
    worksheet.set_column('E:E', 10, float_fmt)
    for title in title_coords:
        worksheet.write(title_coords[title],0, 'CUSTOMER: %s'%title, title_fmt)
    for coordinate in table_coords:
        worksheet.add_table(coordinate, {'columns': [{'header': 'ID',
                                                      'total_string': ' '},
                                                     {'header': 'Customer Name',
                                                      'total_string': ' '},
                                                     {'header': 'GTIN-14',
                                                      'total_string': ' '},
                                                     {'header': 'MPN Name', #remove total_string here
                                                      'total_function': 'count'},
                                                     {'header': 'Weight',
                                                      'total_function': 'sum'}
                                                     ], 
                                       'autofilter': False,
                                       'total_row': True,
                                       'style': 'Table Style Medium 20'})

'''
Input: Dataframe
Output: Dictionary:
{
    '1000000': 
                '00070919010926':
                                    DATAFRAME
                '90061741193416':
}

'''
def format_df(df):
    nested_dict = {}
    for cid,cid_df in df.groupby('cust_id'):
        gtin_dict = {}
        for gtin,gtin_df in cid_df.groupby('gtin'):
            gtin_dict[gtin] = gtin_df
            nested_dict[cid] = gtin_dict
    return nested_dict

def format_excel(writer, nested_dict):
    start_row = 2
    padding = 1
    title_coords = {}
    table_coords = []
    for by_customer in nested_dict:
        title_coords[by_customer] =  start_row-1
        for by_gtin in nested_dict[by_customer]:
            df = nested_dict[by_customer][by_gtin]
            df_length = len(df.index)
            df.to_excel(writer, sheet_name='final', startrow=start_row, index=False)
            end_row = start_row + df_length + padding
            table_coords.append("A%s:E%s"%(str(start_row+1), str(end_row+1))) 
            start_row = end_row + padding
        start_row = start_row + 2*padding
    format_excel_table(writer, table_coords, title_coords)

def mk_file():
    date = datetime.today().strftime("%Y%m%d-%H%M%S")
    filename = date + "_mm_inv.xlsx" 
    return Path.home()/'Dropbox'/'mm_data'/filename

def excelify(df, datastream, eb):

    nested_dict = format_df(df)
    # Create File
    out_file = mk_file()
    writer = pd.ExcelWriter(out_file, engine='xlsxwriter')
    format_excel(writer, nested_dict)
    datastream.to_excel(writer, sheet_name='datastream')
    eb.to_excel(writer, sheet_name='barcode errors')
    writer.save()
    # Open up Excel and make it visible
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = True
    # Open up the file
    excel.Workbooks.Open(out_file)
    # Wait before closing it
    _ = input("Press enter to close Excel")
    excel.Workbooks.Close()
    excel.Application.DisplayAlerts = True
    excel.Application.Quit() 
    