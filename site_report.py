# Creates On-Demand / Site Array report

from pure_pull import *
from array_report import *
from datetime import datetime, timedelta
import pandas as pd
from math import trunc
from sharepoint import file_upload


def add_data_columns(array_name, used, charge, ondemand):
    center = workbook.add_format({'align': 'center'})
    currency = workbook.add_format({'num_format': '$#,##0.00'})
    bold_head = workbook.add_format({'align': 'center', 'bold': True, 'font_size': 11})

    data_start = 19
    sna = 'array1'
    dwf = 'array2'
    # Iter through csv
    for column in arrayCsv.iterrows():
        # Put row into a dictionary
        x = {'Time': column[1]['Time'].strftime('%m/%d/%Y'),
             'Effective_Used (Byte)': trunc(column[1]['Effective_Used (Byte)'] / 1024 / 1024 / 1024), 'array': column[1]['Array_Name']}
        # Put array data into specific columns based on Array_Name
        if column[1]['Array_Name'] == array_name:
            worksheet.write(f'A{data_start}', x['Time'])
            worksheet.write('B18', sna, bold_head)
            worksheet.write(f'{used}{data_start}', x['Effective_Used (Byte)'], center)
            worksheet.write(f'{charge}18', 'Daily Charge', bold_head)
            worksheet.write('F18', dwf, bold_head)
            worksheet.write(f'{charge}{data_start}', round(x['Effective_Used (Byte)'] * daily_charge, 2), center)
            worksheet.write(f'{ondemand}18', 'On Demand Charge', bold_head)
            worksheet.write(f'{ondemand}{data_start}', round((x['Effective_Used (Byte)'] - reserve) * ondemand_charge * 12 / 365, 2), center)
            data_start += 1
    total_start = data_start + 1
    # Add totals and change to currency format
    worksheet.write(f'C{total_start}', f'=SUM(C19:C{data_start})', currency)
    worksheet.write(f'D{total_start}', f'=SUM(D19:D{data_start})', currency)
    worksheet.write(f'G{total_start}', f'=SUM(G19:G{data_start})', currency)
    worksheet.write(f'H{total_start}', f'=SUM(H19:H{data_start})', currency)


today = datetime.today().strftime('%Y-%m-%d')
fileName = 'array_capacity.csv'
technologent_charge = 0.075
daily_charge = technologent_charge * 12 / 365
ondemand_charge = 0.06
reserve = 51200

today_date = datetime.today()
first = today_date.replace(day=1)
lastMonth = first - timedelta(days=1)
lastMonthname = lastMonth.strftime('%B')
lastMonthyear = lastMonth.strftime('%Y')


# Put raw report data into DataFrame
arrayCsv = pd.read_csv(fileName)
# Take time off date column and change to datetime
arrayCsv['Time'] = arrayCsv['Time'].str.split('T').str[0]
arrayCsv['Time'] = pd.to_datetime(arrayCsv['Time'])
arrayCsv = arrayCsv[arrayCsv['Time'].dt.strftime('%Y-%m') == lastMonth.strftime('%Y-%m')]

# Initialize writer, workbook and worksheet
writer = pd.ExcelWriter(f'report_name-{lastMonthname}{lastMonthyear}_processed.xlsx', engine='xlsxwriter', options={'nan_inf_to_errors': True})
df = pd.DataFrame()
df.to_excel(writer, sheet_name='Site Report')
workbook = writer.book
worksheet = writer.sheets['Site Report']

# Create excel formats
bold = workbook.add_format({'align': 'bottom align', 'bold': True, 'font_size': 10})
title = workbook.add_format({'align': 'bottom align', 'bold': True, 'font_size': 12})
center_array = workbook.add_format({'align': 'center', 'bold': True})


# Add static items to sheet
worksheet.insert_image('A1', 'pure.png')
worksheet.merge_range('A4:C4', 'Pure-as-a-Service Usage Report', title)
worksheet.write('A5', 'Report Period', bold)
worksheet.write('B5', lastMonth.replace(day=1).strftime('%m-%d-%Y') + ' - ' + lastMonth.strftime('%m-%d-%Y'))
worksheet.write('A6', 'Customer', bold)
worksheet.write('B6', 'customer')
worksheet.write('A7', 'Partner', bold)
worksheet.write('B7', 'Technologent')
worksheet.write('A8', 'Partner Email', bold)
worksheet.write('B8', 'email')
worksheet.write('A11', 'Service Start Date', bold)
worksheet.write('B11', '2019-10-03')
worksheet.write('A12', 'Subscription Number', bold)
worksheet.write('A13', 'Site ID', bold)
worksheet.write('B12', 'ID')
worksheet.write('B13', 'Site Name', bold)
worksheet.write('C13', 'Site Address', bold)
worksheet.write('E11', 'Reserve Capacity Commit (in GiB)', bold)
worksheet.write('E12', reserve)
worksheet.write('F11', 'demand', bold)
worksheet.write('F12', ondemand_charge)
worksheet.write('G11', 'charge', bold)
worksheet.write('G12', technologent_charge)
worksheet.write('H11', 'charge', bold)
worksheet.write('H12', daily_charge)
worksheet.write('A14', 'cluster')
worksheet.write('A15', 'cluster')
worksheet.write('B14', 'Fgrp')
worksheet.write('C14', 'address1')
worksheet.write('B15', 'grp')
worksheet.write('C15', 'address2')
worksheet.write('A17', 'Date', bold)
worksheet.merge_range('B17:D17', 'cluster', center_array)
worksheet.merge_range('F17:H17', 'cluster', center_array)

# Run data column function to add computed data to sheet
add_data_columns('array1', 'B', 'C', 'D')
add_data_columns('array2', 'F', 'G', 'H')

# Increase column width of sheet
worksheet.set_column('A:H', 20)

writer.save()
os.remove('array_capacity.csv')
file_upload(writer.path, 'folder_path')
