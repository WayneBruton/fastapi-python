from openpyxl import Workbook
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
