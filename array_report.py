from datetime import datetime, timedelta
import pandas as pd
from zipfile import ZipFile
from math import trunc
import os
from sharepoint import file_upload


def add_data_columns(array_name, used, charge):
    center = workbook.add_format({'align': 'center'})
    currency = workbook.add_format({'num_format': '$#,##0.00'})
    bold_head = workbook.add_format({'align': 'center', 'bold': True, 'font_size': 11})
    data_start = 19
    sna = 'SCL-0000185'
    dwf = 'SCL-0000184'
    # Iter through csv
    for column in arrayCsv.iterrows():
        # Put row into a dictionary
        x = {'Time': column[1]['Time'].strftime('%m/%d/%Y'),
             'Effective_Used (Byte)': trunc(column[1]['Effective_Used (Byte)'] / 1024 / 1024 / 1024), 'array': column[1]['Array_Name']}
        # Put array data into specific columns based on Array_Name
        if column[1]['Array_Name'] == array_name:
            worksheet2.write(f'A{data_start}', x['Time'])
            worksheet2.write('B18', sna, bold_head)
            worksheet2.write(f'{used}{data_start}', x['Effective_Used (Byte)'], center)
            worksheet2.write(f'{charge}18', 'Daily Charge', bold_head)
            worksheet2.write('F18', dwf, bold_head)
            worksheet2.write(f'{charge}{data_start}', round(x['Effective_Used (Byte)'] * daily_charge, 2), center)
            data_start += 1
    total_start = data_start + 1
    # Add totals and change to currency format
    worksheet2.write(f'C{total_start}', f'=SUM(C19:C{data_start})', currency)
    worksheet2.write(f'G{total_start}', f'=SUM(G19:G{data_start})', currency)


today = datetime.today().strftime('%Y-%m-%d')
fileName = 'array_capacity.csv'
technologent_charge = 0.075
daily_charge = technologent_charge * 12 / 365
# reserve = 51200

today_date = datetime.today()
first = today_date.replace(day=1)
lastMonth = first - timedelta(days=1)
lastMonthname = lastMonth.strftime('%B')
lastMonthyear = lastMonth.strftime('%Y')

# Unzip report into script folder
with ZipFile(f'/PathTo/array_capacity_{today}.zip', 'r') as zf:
    zf.extract(fileName)
os.remove(f'array_capacity_{today}.zip')

# Put raw report data into DataFrame
arrayCsv = pd.read_csv(fileName)
# Take time off date column and change to datetime
arrayCsv['Time'] = arrayCsv['Time'].str.split('T').str[0]
arrayCsv['Time'] = pd.to_datetime(arrayCsv['Time'])
arrayCsv = arrayCsv[arrayCsv['Time'].dt.strftime('%Y-%m') == lastMonth.strftime('%Y-%m')]

# Initialize writer, workbook and worksheet2
writer = pd.ExcelWriter(f'file_name-{lastMonthname}{lastMonthyear}_processed_array.xlsx', engine='xlsxwriter', options={'nan_inf_to_errors': True})
df = pd.DataFrame()
df.to_excel(writer, sheet_name='Array Report')
workbook = writer.book
worksheet2 = writer.sheets['Array Report']

# Create excel formats
bold = workbook.add_format({'align': 'bottom align', 'bold': True, 'font_size': 10})
title = workbook.add_format({'align': 'bottom align', 'bold': True, 'font_size': 12})
center_array = workbook.add_format({'align': 'center', 'bold': True})

# Add static items to sheet
worksheet2.insert_image('A1', 'pure.png')
worksheet2.merge_range('A4:C4', 'usage', title)
worksheet2.write('A5', 'Report Period', bold)
worksheet2.write('B5', lastMonth.replace(day=1).strftime('%m-%d-%Y') + ' - ' + lastMonth.strftime('%m-%d-%Y'))
worksheet2.write('A6', 'Customer', bold)
worksheet2.write('B6', 'customer')
worksheet2.write('A7', 'Partner', bold)
worksheet2.write('B7', 'Technologent')
worksheet2.write('A8', 'Partner Email', bold)
worksheet2.write('B8', 'email')
worksheet2.write('A11', 'Service Start Date', bold)
worksheet2.write('B11', '2019-10-03')
worksheet2.write('A12', 'Subscription Number', bold)
worksheet2.write('A13', 'Site ID', bold)
worksheet2.write('B12', 'siteid')
worksheet2.write('B13', 'Site Name', bold)
worksheet2.write('C13', 'Site Address', bold)
worksheet2.write('F11', 'charge', bold)
worksheet2.write('F12', technologent_charge)
worksheet2.write('G11', 'charge', bold)
worksheet2.write('G12', daily_charge)
worksheet2.write('A14', 'cluster')
worksheet2.write('A15', 'cluster')
worksheet2.write('B14', 'Fcomp')
worksheet2.write('C14', 'address')
worksheet2.write('B15', 'comp')
worksheet2.write('C15', 'address2')
worksheet2.write('A17', 'Date', bold)
worksheet2.merge_range('B17:C17', 'array', center_array)
worksheet2.merge_range('F17:G17', 'array', center_array)

# Run data column function to add computed data to sheet
add_data_columns('array', 'B', 'C')
add_data_columns('array', 'F', 'G')

# Increase column width of sheet
worksheet2.set_column('A:H', 20)

writer.save()
file_upload(writer.path, 'folder_here')
