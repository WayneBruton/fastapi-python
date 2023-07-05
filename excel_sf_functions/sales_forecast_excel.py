# from datetime import timedelta
from datetime import datetime

from openpyxl import Workbook
# from openpyxl.utils import get_column_letter, column_index_from_string

from openpyxl.styles import PatternFill
from openpyxl.styles.borders import Border, Side
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter

# from openpyxl.utils import range_boundaries

from excel_sf_functions.create_sales_sheet import create_excel_array
from excel_sf_functions.create_NSST_sheets import create_nsst_sheet
from excel_sf_functions.format_sf_sheets import format_sales_forecast
from excel_sf_functions.format_nsst_sheets import format_nsst
import time
import threading


def create_sales_forecast_file(data, developmentinputdata, pledges, firstName, listData, request):
    start_time = time.time()
    # print(firstName)
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

    ##################################

    worksheet_data = []

    report_array = request['Category']
    if len(report_array) > 1:
        sheet_name = f"Investor Exit List Heron"
    else:
        sheet_name = f"Investor Exit List {report_array[0]}"

    ws = wb.create_sheet(sheet_name)

    # ws = wb.active
    ws.title = sheet_name

    # make tab color dark red
    ws.sheet_properties.tabColor = "B31312"

    row1_data = [f"{sheet_name} - {request['date']}"]

    # insert row1_data into the beginning of worksheet_data
    worksheet_data.append(row1_data)
    # loop through worksheet_data and append each item to the worksheet
    row2_data = ["Total Units", "Unit No", "Block", "Investor", "Capital Amount", "Fund Release Date",
                 "Unit Sold Status",
                 "Occupation Date", "Estimated Transfer Date", "Actual Transfer Date", "Exit Deadline (712 Days)",
                 "Current Report Date", "Days to Contact Exit", "180 day Threshhold Warning",
                 "Days to Estimated Exit",
                 "Investor Contract expiry exit", "Capital & Interest to be Exited", "Investor Exit Value On Sales",
                 "Exited by Developer",
                 "Date of Exit", "Early Release", "Investor pay Back On transfer"]
    # print(row2_data[16])
    worksheet_data.append(row2_data)
    row3_data = ["", "", "", "Investor Capital Deployed", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
                 "",
                 "", "", ""]
    worksheet_data.append(row3_data)

    for item in listData:
        row_data = [item['investor_acc_number'], item['opportunity_code'], item['block'], item['investment_name'],
                    float(item['investment_amount']),
                    item['release_date'],
                    item['opportunity_sold'], item['occupation_date'], item['estimated_transfer_date'],
                    item['final_transfer_date'], "", item['report_date'],
                    item['days_to_exit_deadline'], "", "", item['investment_interest'], item['investment_interest'],
                    0, item['exited_by_developer'], item['date_of_exit'], item['early_release'], ""]
        # print(row_data[16], row_data[17], row_data[18])
        worksheet_data.append(row_data)

    for item in worksheet_data:
        ws.append(item)

    # Format column D to currency
    for cell in ws['E']:
        cell.number_format = '#,##0.00'
    for cell in ws['P']:
        cell.number_format = '#,##0.00'
    for cell in ws['Q']:
        cell.number_format = '#,##0.00'
    for cell in ws['R']:
        cell.number_format = '#,##0.00'
    for cell in ws['S']:
        cell.number_format = '#,##0.00'
    for cell in ws['V']:
        cell.number_format = '#,##0.00'

    cols = ['E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'P', 'Q', 'R', 'S', 'T', 'U', 'V']
    # Make Column C 20 wide
    ws.column_dimensions['C'].width = 10
    ws.column_dimensions['D'].width = 30
    for col in cols:
        ws.column_dimensions[col].width = 15

    # make columns D through N 15 wide

    for col in ws.iter_cols(min_col=1, max_col=22):
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
    for row in ws.iter_rows(min_row=4):
        if row[6].value:
            for cell in row:
                cell.fill = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')

    # for all rows after row 2, if the value in column J has a value, make the row background color light red and
    # hide the row
    for row in ws.iter_rows(min_row=4):

        if row[9].value != "":

            for cell in row:
                cell.fill = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
            ws.row_dimensions[row[0].row].hidden = True

    # loop through cells in column M from row 4 to the end of the worksheet, if the value is <= 90, make the cell
    # background color light red

    # for all rows after 2, set a formula in column j to equal to column E + (365 * 2)
    for row in ws.iter_rows(min_row=4):
        row[10].value = f'=F{row[0].row}+730'
        if row[6].value == True and row[20].value == True and row[9].value == "":
            row[21].value = row[18].value
        else:
            row[21].value = 0
    # format column J as date YYYY-MM-DD
    for cell in ws['K']:
        cell.number_format = 'YYYY-MM-DD'

    for cell in ws['M']:
        cell.number_format = '0'

    # if cells after row 4 in column 'M' are <= 90, make the cell in colmn 'M' have a background color light red
    for row in ws.iter_rows(min_row=4):
        if row[12].value <= 90:
            row[12].fill = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')

            if row[18].value != 0:
                row[4].value = 0

        else:
            row[15].value = 0
        if row[18].value != 0:
            row[4].value = 0

    sheet1 = wb.sheetnames[0]
    # print("sheets", sheet1)
    ws2 = wb[sheet1]
    # print("ws", ws2)
    # get the column name, for example 'A' or 'B' etc of the last column of ws2
    last_col = get_column_letter(ws2.max_column)
    # print("last_col", last_col)

    # set a formula in colmn L to equal to J - K formatted as a date YYYY-MM-DD
    for row in ws.iter_rows(min_row=4):
        # get row number and print it
        # print("row", row[0].row)

        # =IF(S20 <> 0, 0, IF(J20="", K20 - L20, 0))
        row[12].value = f'=IF(S{row[0].row}<>0,0,IF(J{row[0].row}="",K{row[0].row}-L{row[0].row},0))'
        row[15].value = f'=IF(M{row[0].row}<90,+Q{row[0].row},0)'
        # row[17].value = row[6].value
        if row[6].value:
            row[17].value = row[16].value
        else:
            row[17].value = 0

        # DO SOMETHING HERE

    # format column L as integer

    # in column N, if column I is not blank, set the value to I - K formatted as integer, else set the value to H - K
    # formatted as integer
    for row in ws.iter_rows(min_row=4):
        row[14].value = f'=IF(J{row[0].row}="",I{row[0].row}-L{row[0].row},J{row[0].row}-L{row[0].row})'
    for row in ws.iter_rows(min_row=3, max_row=3):
        row[1].value = f'=COUNTIFS(B4:B{ws.max_row},"<>",J4:J{ws.max_row},"=")'
        row[0].value = listData[0]['count_of_units']
        row[4].value = f'=SUMIFS(E4:E{ws.max_row}, J4:J{ws.max_row}, "")'
        row[15].value = f'=SUM(P4:P{ws.max_row})'
        row[16].value = f'=SUM(Q4:Q{ws.max_row})'
        row[17].value = f'=SUM(R4:R{ws.max_row})'
        row[18].value = f'=SUM(S4:S{ws.max_row})'
        row[21].value = f'=SUM(V4:V{ws.max_row})'

    # put the contents of column B, column E and column J into a new list of dictionaries
    list_to_filter = []
    for row in ws.iter_rows(min_row=4):
        list_to_filter.append(
            {'unit': row[1].value, 'amount': row[4].value, 'date': row[9].value, 'with_interest': row[16].value,
             'sold': row[6].value})
        # filter out of list_to_filter where the value of 'date' is equal to ""
        list_to_filter = list(filter(lambda x: x['date'] == "", list_to_filter))

    # total = sum([x['with_interest'] for x in list_to_filter]) - sum([x['amount'] for x in list_to_filter])
    # print("total", total)

    # make row 3 bold, center, font size 12, text color white, background color red
    for cell in ws[3]:
        cell.font = Font(bold=True, size=12, color='FFFFFF')
        cell.alignment = Alignment(horizontal='center')
        cell.fill = PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid')

    for cell in ws['O']:
        cell.number_format = '0'

    for row in ws.iter_rows(min_row=4):
        # print("row", row[0].row)
        row[
            16].value = f"=IF(A{row[0].row}&B{row[0].row}<>A{row[0].row - 1}&B{row[0].row - 1},SUMIFS('{sheet1}'!$G$78:${last_col}$78, '{sheet1}'!$G$10:${last_col}$10,A{row[0].row}, '{sheet1}'!$G$4:${last_col}$4, B{row[0].row}),0)"
        # =IF(A20 <> A19, IF(Q20 <> 0, SUMIFS(
        #     'SF Endulini'!$G$19:$HE$19, 'SF Endulini'!$G$10:$HE$10, A20, 'SF Endulini'!$G$4:$HE$4, B20), 0), 0)
        row[
            4].value = f"=IF(A{row[0].row}&B{row[0].row}<>A{row[0].row - 1}&B{row[0].row - 1},IF(Q{row[0].row},SUMIFS('{sheet1}'!$G$19:${last_col}$19, '{sheet1}'!$G$10:${last_col}$10,A{row[0].row}, '{sheet1}'!$G$4:${last_col}$4, B{row[0].row}),0),0)"
        row[
            18].value = f"=IF(A{row[0].row}<>A{row[0].row - 1},SUMIFS('{sheet1}'!$G$60:${last_col}$60, '{sheet1}'!$G$10:${last_col}$10,A{row[0].row}, '{sheet1}'!$G$4:${last_col}$4, B{row[0].row}),0)"
        # if row[16].value != 0:
        #     row[15].value = row[16].value
        # print("row[18].value", row[18].value)
        # if row[18].value != 0:
        #     row[12].value = 0

    # Hide columns N & O
    cols_to_hide = ['F', 'G', 'H', 'I', 'J', 'K', 'L', 'N', 'T', 'U']
    for col in cols_to_hide:
        ws.column_dimensions[col].hidden = True
    # ws.column_dimensions['N'].hidden = True
    # ws.column_dimensions['O'].hidden = True

    # in row2 set all cells to bold and wrap text
    for cell in ws[2]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(wrap_text=True, vertical='center', horizontal='center')

    rows_for_full_merge = [1]
    for row in rows_for_full_merge:
        ws.merge_cells(f'A{row}:N{row}')
    ##################################

    # CREATE NSST SHEET

    # print(firstName)
    if firstName == 'Wayne' or firstName == 'Wynand' or firstName == 'Nick' or firstName == 'Deric' or \
            firstName == 'Debbie' or firstName == 'Leandri':
        if len(category) > 1:
            sheet_name = f"NSST {category[0].split(' ')[0]}"

        else:
            sheet_name = f"NSST {category[0]}"

        worksheets = wb.sheetnames
        # print(worksheets)

        # if len(worksheets) == 1:
        index = 0

        ws = wb.create_sheet(sheet_name)
        # Make Tab Color of new Sheet 'Green'

        ws.sheet_properties.tabColor = "539165"

        nsst_data = create_nsst_sheet(category, developmentinputdata, pledges, index, sheet_name, worksheets)

        for item in nsst_data:
            ws.append(item)

        if len(worksheets) == 4:
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
            # print("index", index, "sheet", sheet)
            format_nsst(num_sheets, index, sheet, list_to_filter)

    else:
        # loop through each sheet and delete all rows after row 40
        for sheet in wb.worksheets:
            sheet.delete_rows(40, 1000)

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
        sheet_name = f"Investor Exit List Heron"
    else:
        sheet_name = f"Investor Exit List {report_array[0]}"

    ws = wb.active
    ws.title = sheet_name
    ws.sheet_properties.tabColor = "1072BA"

    row1_data = [f"{sheet_name} - {request['date']}"]

    # insert row1_data into the beginning of worksheet_data
    worksheet_data.append(row1_data)
    # loop through worksheet_data and append each item to the worksheet
    row2_data = ["Total Units", "Unit No", "Block", "Investor", "Capital Amount", "Fund Release Date",
                 "Unit Sold Status",
                 "Occupation Date", "Estimated Transfer Date", "Actual Transfer Date", "Exit Deadline (712 Days)",
                 "Current Report Date", "Days to Contact Exit", "180 day Threshhold Warning",
                 "Days to Estimated Exit",
                 "Investor Contract expiry exit", "Capital & Interest to be Exited", "Investor Exit Value On Sales",
                 "Exited by Developer",
                 "Date of Exit"]
    worksheet_data.append(row2_data)
    row3_data = ["", "", "", "Investor Capital Deployed", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
                 ""]
    worksheet_data.append(row3_data)

    for item in data:
        row_data = [item['investor_acc_number'], item['opportunity_code'], item['block'], item['investment_name'],
                    float(item['investment_amount']),
                    item['release_date'],
                    item['opportunity_sold'], item['occupation_date'], item['estimated_transfer_date'],
                    item['final_transfer_date'], "", item['report_date'],
                    item['days_to_exit_deadline'], "", "", item['investment_interest'], item['investment_interest'],
                    0, item['exited_by_developer'], item['date_of_exit']]
        worksheet_data.append(row_data)

    for item in worksheet_data:
        ws.append(item)

    # Format column D to currency
    for cell in ws['E']:
        cell.number_format = '#,##0.00'
    for cell in ws['P']:
        cell.number_format = '#,##0.00'
    for cell in ws['Q']:
        cell.number_format = '#,##0.00'
    for cell in ws['R']:
        cell.number_format = '#,##0.00'
    for cell in ws['S']:
        cell.number_format = '#,##0.00'

    cols = ['E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'P', 'Q', 'R', 'S', 'T']
    # Make Column C 20 wide
    ws.column_dimensions['C'].width = 10
    ws.column_dimensions['D'].width = 30
    for col in cols:
        ws.column_dimensions[col].width = 15

    # make columns D through N 15 wide

    for col in ws.iter_cols(min_col=1, max_col=20):
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
    for row in ws.iter_rows(min_row=4):
        if row[6].value:
            for cell in row:
                cell.fill = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')

    # for all rows after row 2, if the value in column J has a value, make the row background color light red and
    # hide the row
    for row in ws.iter_rows(min_row=4):
        if row[9].value != "":
            for cell in row:
                cell.fill = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
            ws.row_dimensions[row[0].row].hidden = True

    # loop through cells in column M from row 4 to the end of the worksheet, if the value is <= 90, make the cell
    # background color light red

    # for all rows after 2, set a formula in column j to equal to column E + (365 * 2)
    for row in ws.iter_rows(min_row=4):
        row[10].value = f'=F{row[0].row}+730'
    # format column J as date YYYY-MM-DD
    for cell in ws['K']:
        cell.number_format = 'YYYY-MM-DD'

    # if cells after row 4 in column 'M' are <= 90, make the cell in colmn 'M' have a background color light red
    for row in ws.iter_rows(min_row=4):
        if row[12].value <= 90:
            row[12].fill = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
            if row[18].value != 0:
                row[4].value = 0

        else:
            row[15].value = 0

    # set a formula in colmn L to equal to J - K formatted as a date YYYY-MM-DD
    for row in ws.iter_rows(min_row=4):
        # =IF(S112 <> 0, 0, IF(J112="", K112 - L112, 0))
        row[12].value = f'=IF(S{row[0].row}<>0,0,IF(J{row[0].row}="",K{row[0].row}-L{row[0].row},0))'
        # row[17].value = row[6].value
        if row[6].value:
            row[17].value = row[16].value
        else:
            row[17].value = 0

    # format column L as integer
    for cell in ws['M']:
        cell.number_format = '0'

    # in column N, if column I is not blank, set the value to I - K formatted as integer, else set the value to H - K
    # formatted as integer
    for row in ws.iter_rows(min_row=4):
        row[14].value = f'=IF(J{row[0].row}="",I{row[0].row}-L{row[0].row},J{row[0].row}-L{row[0].row})'
    for row in ws.iter_rows(min_row=3, max_row=3):
        # =SUM(--(LEN(UNIQUE(FILTER(B4:B192, J4:J192="", ""))) > 0))
        row[0].value = f'=SUM(--(LEN(UNIQUE(FILTER(B4:B{ws.max_row}, J4:J{ws.max_row}, ""))) > 0))'
        # =COUNTIFS($B$4:$B$192, "<>",$J$4:$J$192, "=")
        row[1].value = f'=COUNTIF(B4:B{ws.max_row}, "<>", J4:J{ws.max_row}, "=")'

        row[4].value = f'=SUMIFS(E4:E{ws.max_row}, J4:J{ws.max_row}, "")'
        row[15].value = f'=SUM(P4:P{ws.max_row})'
        row[16].value = f'=SUM(Q4:Q{ws.max_row})'
        row[17].value = f'=SUM(R4:R{ws.max_row})'
        row[18].value = f'=SUM(S4:S{ws.max_row})'

    # make row 3 bold, center, font size 12, text color white, background color red
    for cell in ws[3]:
        cell.font = Font(bold=True, size=12, color='FFFFFF')
        cell.alignment = Alignment(horizontal='center')
        cell.fill = PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid')

    for cell in ws['O']:
        cell.number_format = '0'

    # Hide columns N & O
    ws.column_dimensions['N'].hidden = True
    ws.column_dimensions['O'].hidden = True

    # in row2 set all cells to bold and wrap text
    for cell in ws[2]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(wrap_text=True)

    rows_for_full_merge = [1]
    for row in rows_for_full_merge:
        ws.merge_cells(f'A{row}:N{row}')

    wb.save(f"excel_files/{sheet_name}.xlsx")


def create_cash_flow(data, request, other_data):
    # print("data", data)
    # print("request", request)

    worksheet_data = []

    if len(request['Category']) > 1:
        heading = 'Heron'
    else:
        heading = request['Category'][0]

    wb = Workbook()

    ws = wb.active
    ws.title = "Cash Flow"

    row1_data = [f"Weekly Cashflow ({heading}) - {request['date']}","",""]

    row2_data = ["","",""]

    row3_data = ["Date", "Amount", "Units"]

    worksheet_data.append(row1_data)
    worksheet_data.append(row2_data)
    worksheet_data.append(row3_data)
    worksheet_data.append(row2_data)

    for item in data:
        row_data = [item['date'], item['total_cashflow'], item['units']]
        worksheet_data.append(row_data)

    for item in worksheet_data:
        # print(item)
        ws.append(item)

    # format column B as currency
    for cell in ws['B']:
        cell.number_format = 'R#,##0.00'

    # format column A as date
    for cell in ws['A']:
        cell.number_format = 'YYYY-MM-DD'

    # set column B width to 20
    ws.column_dimensions['B'].width = 20
    ws.column_dimensions['A'].width = 15

    # add a worksheet
    ws2 = wb.create_sheet("Cash Flow Summary")

    worksheet_data = []
    row1_data = ["Date", "Amount", "Units"]
    worksheet_data.append(row1_data)

    for item in other_data:
        row_data = [item['opportunity_end_date'], item['nett_cashflow'], item['opportunity_code']]
        worksheet_data.append(row_data)

    for item in worksheet_data:
        # print(item)
        ws2.append(item)



    for cell in ws2['B']:
        cell.number_format = 'R#,##0.00'

    # format column A as date
    for cell in ws2['A']:
        cell.number_format = 'YYYY-MM-DD'

    # set column B width to 20
    ws2.column_dimensions['B'].width = 20
    ws2.column_dimensions['A'].width = 15



    wb.save(f"excel_files/Cashflow {heading}.xlsx")

    return f"Cashflow {heading}.xlsx"


# create_cash_flow([{
#     "date": "2023/06/25",
#     "total_cashflow": 1520710.5982950684,
#     "units": "HVC102,HVC106,HVC105,HVC103,HVC101,HVC104"
# },
#     {
#         "date": "2023/06/27",
#         "total_cashflow": 58374.981506849406,
#         "units": "HFB214"
#     },
#     {
#         "date": "2023/07/18",
#         "total_cashflow": 2381304.6562584257,
#         "units": "HVC305,HVC205,HVC302,HVC204,HVC301,HVC206,HVC304,HVC303,HVC306,HVC201,HVC202,HVC203"
#     },
#     {
#         "date": "2023/08/19",
#         "total_cashflow": 4305378.80258411,
#         "units": "HVP101,HVP102,HVP201,HVP103,HVP203,HVP301,HVP302,HVP303,HVP202"
#     },
#     {
#         "date": "2023/09/19",
#         "total_cashflow": 1311740.05,
#         "units": "HFB315"
#     },
#     {
#         "date": "2023/11/07",
#         "total_cashflow": 2547835.3119508224,
#         "units": "HVD101,HVD102,HVD103,HVD104,HVD202,HVD201,HVD203,HVD204,HVD301,HVD302,HVD303,HVD304"
#     },
#     {
#         "date": "2023/12/05",
#         "total_cashflow": 3449508.8032786306,
#         "units": "HVN101,HVN102,HVN103,HVN104,HVN201,HVN202,HVN203,HVN204,HVN301,HVN302,HVN303,HVN304"
#     }],
#     {
#         "Category": ["Heron Fields", "Heron View"],
#         "date": "2023/05/31",
#         "firstName": "Wayne"
#     }
# )
