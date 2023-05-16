from datetime import timedelta

from openpyxl import Workbook
from openpyxl.utils import get_column_letter, column_index_from_string

from openpyxl.styles import PatternFill
from openpyxl.styles.borders import Border, Side
from openpyxl.styles import Font, Alignment

from excel_sf_functions.create_sales_sheet import create_excel_array
from excel_sf_functions.create_NSST_sheets import create_nsst_sheet
from excel_sf_functions.format_sf_sheets import format_sales_forecast
from excel_sf_functions.format_nsst_sheets import format_nsst
import time
import threading


def create_sales_forecast_file(data, developmentinputdata, pledges):
    start_time = time.time()
    category = developmentinputdata['Category']
    worksheet_names = []
    if len(category) > 1:
        filename = f'Sales Forecast{category[0].split(" ")[0]}'
        sheet_name = f"SF {category[0]} & {category[1]}"

    else:
        sheet_name = f"SF {category[0]}"
        filename = f'Sales Forecast{category[0]}'
        worksheet_names.append(sheet_name)

    wb = Workbook()
    worksheet_data, interest_on_funds_drawn, interest_on_funds_in_momentum = create_excel_array(data)

    ws = wb.active
    ws.title = sheet_name
    ws.sheet_properties.tabColor = "1072BA"

    row1_data = [f"{sheet_name} - {developmentinputdata['date']}"]
    # insert row1_data into the beginning of worksheet_data
    worksheet_data.insert(0, row1_data)
    # loop through worksheet_data and append each item to the worksheet
    for item in worksheet_data:
        ws.append(item)

    if len(category) > 1:
        for index, item in enumerate(category):
            sheet_name = f"SF {item}"
            row1_data = [f"{sheet_name} - {developmentinputdata['date']}"]
            # filter data based on the category into new list using list comprehension
            data_input = [x for x in data if x['Category'] == item]
            worksheet_data, interest_on_funds_drawn, interest_on_funds_in_momentum = create_excel_array(data_input)

            # insert row1_data into the beginning of worksheet_data
            worksheet_data.insert(0, row1_data)
            # create a new worksheet
            ws = wb.create_sheet(sheet_name)
            ws.sheet_properties.tabColor = "1072BA"
            # loop through worksheet_data and append each item to the worksheet
            for item1 in worksheet_data:
                ws.append(item1)

    # LOOP THROUGH EACH SHEET AND FORMAT ACCORDINGLY
    threads = []
    for sheet in wb.worksheets:
        t = threading.Thread(target=format_sales_forecast, args=(sheet,))
        t.start()
        threads.append(t)

    for t in threads:
        t.join()

    # CREATE NSST SHEET
    if len(category) > 1:
        sheet_name = f"NSST {category[0].split(' ')[0]}"

    else:
        sheet_name = f"NSST {category[0]}"

    worksheets = wb.sheetnames

    # if len(worksheets) == 1:
    index = 0

    ws = wb.create_sheet(sheet_name)
    # Make Tab Color of new Sheet 'Green'

    ws.sheet_properties.tabColor = "539165"

    # if len(worksheets) == 1:

    nsst_data = create_nsst_sheet(category, developmentinputdata, pledges, index, sheet_name, worksheets)

    for item in nsst_data:
        ws.append(item)

    if len(worksheets) == 3:
        # filter pledges to only include the category[0]
        new_pledges = [pledge for pledge in pledges if pledge['Category'] == category[0]
                       and pledge['opportunity_code'] != 'Unallocated']
        index = 1
        sheet_name = f"NSST {category[0]}"
        ws = wb.create_sheet(sheet_name)
        # Make Tab Color of new Sheet 'Green'
        ws.sheet_properties.tabColor = "539165"
        nsst_data = create_nsst_sheet(category, developmentinputdata, new_pledges, index, sheet_name, worksheets)

        for item in nsst_data:
            ws.append(item)

        if len(pledges):
            new_pledges = [pledge for pledge in pledges if pledge['Category'] == category[1]
                           and pledge['opportunity_code'] != 'Unallocated']
        else:
            new_pledges = []
        index = 2
        sheet_name = f"NSST {category[1]}"

        ws = wb.create_sheet(sheet_name)
        # Make Tab Color of new Sheet 'Green'
        ws.sheet_properties.tabColor = "539165"
        nsst_data = create_nsst_sheet(category, developmentinputdata, new_pledges, index, sheet_name, worksheets)

        for item in nsst_data:
            ws.append(item)

    num_sheets = len(wb.sheetnames)

    for index, sheet in enumerate(wb.worksheets):
        format_nsst(num_sheets, index, sheet)

    # SAVE TO FILE
    wb.save(f"excel_files/{filename}.xlsx")

    end_time = time.time()

    elapsed = end_time - start_time
    minutes, seconds = divmod(elapsed, 60)
    hours, minutes = divmod(minutes, 60)
    days, hours = divmod(hours, 24)

    print(f"Elapsed time: {int(days)} days, {int(hours)} hours, {int(minutes)} minutes and {round(seconds, 2)} seconds")


def create_investment_list(data, request):
    wb = Workbook()

    worksheet_data = []

    report_array = request['Category']
    if len(report_array) > 1:
        sheet_name = f"Investment List Heron"
    else:
        sheet_name = f"Investment List {report_array[0]}"

    ws = wb.active
    ws.title = sheet_name
    ws.sheet_properties.tabColor = "1072BA"

    row1_data = [f"{sheet_name} - {request['date']}"]

    # insert row1_data into the beginning of worksheet_data
    worksheet_data.append(row1_data)
    # loop through worksheet_data and append each item to the worksheet
    row2_data = ["Unit No", "Block", "Investor", "Capital Amount", "Fund Release Date", "Unit Sold Status",
                 "Occupation Date", "Estimated Transfer Date", "Actual Transfer Date", "Exit Deadline (712 Days)",
                 "Current Report Date", "Time remaining (per LA)", "180 day Threshhold Warning", "Days to Exit"]
    worksheet_data.append(row2_data)

    for item in data:
        row_data = [item['opportunity_code'], item['block'], item['investment_name'], float(item['investment_amount']), item['release_date'],
                    item['opportunity_sold'], item['occupation_date'], item['estimated_transfer_date'],
                    item['final_transfer_date'], "", item['report_date'],
                    "", "", ""]
        worksheet_data.append(row_data)

    for item in worksheet_data:
        ws.append(item)

    # Format column D to currency
    for cell in ws['D']:
        cell.number_format = '#,##0.00'


    cols = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N']
    # Make Column C 20 wide
    ws.column_dimensions['C'].width = 30
    for col in cols:
        ws.column_dimensions[col].width = 15


    # make columns D through N 15 wide

    for col in ws.iter_cols(min_col=1, max_col=14):
        for cell in col:
            cell.alignment = Alignment(horizontal='center')

            cell.font = Font(size=10)
            cell.border = Border(left=Side(border_style='thin', color='000000'),
                                 right=Side(border_style='thin', color='000000'),
                                 top=Side(border_style='thin', color='000000'),
                                 bottom=Side(border_style='thin', color='000000'))


        # make row 1 bold, center, fot size 14, text color white, background color blue
    for cell in ws[1]:
        cell.font = Font(bold=True, size=14, color='FFFFFF')
        cell.alignment = Alignment(horizontal='center')
        cell.fill = PatternFill(start_color='1072BA', end_color='1072BA', fill_type='solid')

    # make row 2 bold, center, font size 12, text color black, background color light grey
    for cell in ws[2]:
        cell.font = Font(bold=True, size=12, color='000000')
        cell.alignment = Alignment(horizontal='center')
        cell.fill = PatternFill(start_color='E9E9E9', end_color='E9E9E9', fill_type='solid')

    # for all rows after row 2, if the value in column F is True, make the row background color light green
    for row in ws.iter_rows(min_row=3):
        if row[5].value == True:
            for cell in row:
                cell.fill = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')

    # for all rows after 2, set a formula in column j to equal to column E + (365 * 2)
    for row in ws.iter_rows(min_row=3):
        row[9].value = f'=E{row[0].row}+730'
    # format column J as date YYYY-MM-DD
    for cell in ws['J']:
        cell.number_format = 'YYYY-MM-DD'

    # set a formula in colmn L to equal to J - K formatted as a date YYYY-MM-DD
    for row in ws.iter_rows(min_row=3):
        row[11].value = f'=J{row[0].row}-K{row[0].row}'
    # format column L as integer
    for cell in ws['L']:
        cell.number_format = '0'

    # in column N, if column I is not blank, set the value to I - K formatted as integer, else set the value to H - K
    # formatted as integer
    for row in ws.iter_rows(min_row=3):
        row[13].value = f'=IF(I{row[0].row}="",H{row[0].row}-K{row[0].row},I{row[0].row}-K{row[0].row})'
    # format column N as integer
    for cell in ws['N']:
        cell.number_format = '0'





    # in row2 set all cells to bold and wrap text
    for cell in ws[2]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(wrap_text=True)

    rows_for_full_merge = [1]
    for row in rows_for_full_merge:
        ws.merge_cells(f'A{row}:N{row}')








    wb.save(f"excel_files/{sheet_name}.xlsx")

