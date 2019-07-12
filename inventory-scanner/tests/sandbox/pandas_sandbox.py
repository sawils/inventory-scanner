from datetime import datetime
import pandas as pd
import win32com.client as win32
from pathlib import Path
import xlsxwriter
from pprint import pprint
import time
# Purpose is to get a master dictionary and group by correctly. 

def format_excel_table(writer, start_row, tbl_size, table_coords, title_coords):
    """ Add Excel specifi c formatting to the workbook
    """
    # Get the workbook and the summary sheet so we can add the formatting
    workbook = writer.book
    worksheet = writer.sheets['final'] #pandas version
    # worksheet = workbook.add_worksheet()
    # Add currency formatting and apply it
    num_fmt = workbook.add_format({'num_format': 0, 'align': 'center'})
    gtin_fmt = workbook.add_format({'num_format': 0, 'align': 'right'})
    float_fmt = workbook.add_format({'num_format': '0.00', 'align': 'center'})
    title_fmt = workbook.add_format({'bold':True, 'underline': True, 'font_size': 18})   
    
    worksheet.write(0, 0, 'CUSTOMER: %s'%'c0_name_here', title_fmt)
    worksheet.set_column('A:A', 5)
    worksheet.set_column('B:C', 10, num_fmt)
    worksheet.set_column('D:D', 16, gtin_fmt)
    worksheet.set_column('E:E', 10, float_fmt)
    table_coords
    title_coords
    coordinates = ['A2:E4','A5:E8','A11:E16','A17:E23']
    title_coord = ['A1','A10']
    for x in title_coord:
        worksheet.write(x, 'CUSTOMER: %s'%'c0_name_here', title_fmt)
    # for coordinate in table_coords:
    for coordinate in coordinates:
        worksheet.add_table(coordinate, {'columns': [{'header': '#',
                                                    'total_string': 'Total'},
                                                   {'header': 'ID',
                                                    'total_string': 'sum'},
                                                   {'header': 'Name',
                                                    'total_string': ' '},
                                                    {'header': 'GTIN-14',
                                                    'total_function': 'count'},
                                                    {'header': 'Weight',
                                                    'total_function': 'sum'}],
                                       'autofilter': False,
                                       'total_row': True,
                                       'style': 'Table Style Medium 20'})

def excelify(df):
    # grouped = df_0.groupby(['cust_id','gtin'])
    # df.sort_values(by=['cust_id','gtin'])
    # df.set_index(keys=['cust_id','gtin'])
    # unique_cust = df['cust_id'].unique().tolist()
    # unique_gtin = df['gtin'].unique().tolist()

    # group by cust_id,gtin
    # Can access dataframe
    # 
    nested_dict = {}
    for cid,cid_df in df.groupby('cust_id'):
        gtin_dict = {}
        for gtin,gtin_df in cid_df.groupby('gtin'):
            gtin_dict[gtin] = gtin_df
            nested_dict[cid] = gtin_dict
    pprint(nested_dict)    
    # groupby_customer = dict(tuple(df.groupby('cust_id')))
    # pprint(groupby_customer)
    # print(grouped.sum())
    # print(grouped.count())

    # Create File
    date = datetime.today().strftime("%Y%m%d-%H%M%S")
    filename = date + "_mm_inv.xlsx" 
    out_file = Path.cwd()/filename

    writer = pd.ExcelWriter(out_file, engine='xlsxwriter')

    start_row = 1
    padding = 1
    title_coords = {}
    table_coords = []
    for by_customer in nested_dict:
        title_coords[by_customer] =  start_row-1
        for by_gtin in nested_dict[by_customer]:
            df = nested_dict[by_customer][by_gtin]
            df_length = len(df.index)
            df.to_excel(writer, sheet_name='final', startrow=start_row)
            end_row = start_row + df_length + padding
            table_coords.append("A%s:E%s"%(str(start_row), str(end_row))) 
            start_row = end_row + padding
        start_row = start_row + 2*padding
    print(title_coords)
    print(table_coords)
    format_excel_table(writer, start_row, df_length, table_coords, title_coords)
    writer.save()
    # Open up Excel and make it visible
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = True

    # Open up the file
    excel.Workbooks.Open(out_file)

    # Wait before closing it
    _ = input("Press enter to close Excel")
    excel.Application.Quit() 

master_000 = {'cust_id':[1000001,1000000,1000000,1000000,1000001,1000001,1000001,1000001,1000001,1000001,1000001,1000001],
             'name':['B','A','A','A','B','B','B','B','B','B','B','B'],
             'gtin':['0100070919011022320200968613190124214002553673',
                     '0100070919011022320200945613190124214002553690',
                     '0100070919011022320200945613190124214002553690',
                     '0100070919011022320200945613190124214002553690',
                     '0100070919011022320200945613190124214002553690',
                     '0100070919011022320200945613190124214002553690',
                     '0100070919011022320200945613190124214002553690',
                     '019009642376897232010005791119010321327900304767',
                     '019009642376897232010005791119010321327900304767',
                     '019009642376897232010005791119010321327900304767',
                     '0100070919011022320200968613190124214002553673',
                     '0100070919011022320200968613190124214002553673',],
             'weight':[57.9,94.56,94.56,94.56,
                       94.56,94.56,94.56,
                       96.86,96.86,96.86,
                       57.9,57.9]
            }

master_001 = {'cust_id':[1000001,1000000,1000000,1000000,1000001,1000001,1000001,1000001,1000001,1000001,1000001,1000001],
             'name':['B','A','A','A','B','B','B','B','B','B','B','B'],
             'gtin':['00070919011022',
                     '00070919011022',
                     '90096423768972',
                     '90096423768972',
                     '90096423768972',
                     '00070919011022',
                     '00070919011022',
                     '90096423768972',
                     '90096423768972',
                     '90096423768972',
                     '90096423768972',
                     '00070919011022',],
             'weight':[57.9,94.56,94.56,
                       94.56,94.56,94.56,
                       94.56,96.86,96.86,
                       96.86,57.9,57.9]
            }
df = pd.DataFrame(master_001)
excelify(df)