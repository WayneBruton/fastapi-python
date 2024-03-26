import copy
import os
from calendar import calendar

from dateutil.relativedelta import relativedelta
from openpyxl import Workbook, formatting
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
import re
from openpyxl.formula import Tokenizer
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.formatting.rule import CellIsRule
from datetime import datetime


# from routes.cashflow_routes import calculate_vat_due


def calculate_vat_due(sale_date):
    vat_periods = {
        1: "03/31",
        2: "03/31",
        3: "05/31",
        4: "05/31",
        5: "07/31",
        6: "07/31",
        7: "09/30",
        8: "09/30",
        9: "11/30",
        10: "11/30",
        11: "01/31",
        12: "01/31",
    }

    sale_date = datetime.strptime(sale_date.replace("-", "/"), '%Y/%m/%d')

    sale_month = sale_date.month
    sale_year = sale_date.year
    if sale_month > 10:
        vat_year = sale_year + 1
    else:
        vat_year = sale_year

    vat_date = f"{vat_year}/{vat_periods[sale_month]}"
    # print("vat_date", vat_date)
    return vat_date


def format_header(ws, header_data):
    header = [key.replace("_", " ").title() for key in header_data]
    # append a blank row?
    ws.append([])
    ws.append([])

    ws.append(header)
    # append header in row 4

    for cell in ws[4]:
        cell.font = Font(bold=True)


def format_columns(ws, column_widths, number_format_ranges):
    for col_letter, width in column_widths.items():
        ws.column_dimensions[col_letter].width = 30
    for col_range, number_format in number_format_ranges.items():
        for row in ws.iter_rows(min_row=5, **col_range):
            for cell in row:
                cell.number_format = number_format


def format_cells(ws, last_row):
    for row in ws.iter_rows(min_row=4, max_row=last_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            cell.alignment = Alignment(horizontal='center', vertical='center')


def create_sheet(wb, title, data, column_widths, number_format_ranges):
    ws = wb.create_sheet(title)
    ws.sheet_properties.tabColor = "1072BA"
    ws['A1'] = title
    format_header(ws, data[0])
    for item in data:
        ws.append([item[key] for key in item])
    format_columns(ws, column_widths, number_format_ranges)
    last_row = ws.max_row
    format_cells(ws, last_row)
    ws.auto_filter.ref = f"A4:{get_column_letter(ws.max_column)}{last_row}"
    ws.freeze_panes = "A5"

    return ws


def cashflow_projections(invest, construction, sales, operational_costs, xero,opportunities, report_date):
    # cashflow_p&l_files/cashflow_projection.xlsx exists, then delete it
    if os.path.exists("cashflow_p&l_files/cashflow_projection.xlsx"):
        os.remove("cashflow_p&l_files/cashflow_projection.xlsx")

    for sale in sales:
        sale['refanced_value'] = float(sale['profit_loss'])
        # if sale['Category'] == "Endulini" and sale['refinanced']:
        #     print("sale", sale)
        #     print()
    # print("sales_data", sales[0])

    try:
        report_date = datetime.strptime(report_date, '%Y-%m-%d')

        # print(col_number)
        wb = Workbook()
        ws = create_sheet(wb, 'Investors', invest, {'B': 40}, {})

        # Add a column called "Interest to Date" to the investors sheet



        last_row = ws.max_row
        columns = ['M', 'Q', 'R', 'S','T']
        # in row 2 add a subtotal for each column
        for col in columns:
            ws[f"{col}2"] = f"=SUBTOTAL(9,{col}5:{col}{last_row})"
            ws[f"{col}2"].number_format = '#,##0.00'

        ws1 = wb.create_sheet('Opportunities')
        # make tab color red
        ws1.sheet_properties.tabColor = "FF204E"
        ws1['A1'] = "Opportunities"
        ws1.append([])
        ws1.append([])
        ws1.append([])
        headers = []
        for item in opportunities[0]:
            headers.append(item.replace("_", " ").title())
        ws1.append(headers)

        for item in opportunities:
            ws1.append([item[key] for key in item])
        ws1.auto_filter.ref = f"A5:{get_column_letter(ws1.max_column)}{ws1.max_row}"
        ws1.freeze_panes = "A6"

        # in row 3 add subtotals for columns E and H
        ws1['E3'] = f"=SUBTOTAL(9,E5:E{ws1.max_row})"
        ws1['E3'].number_format = '#,##0.00'
        ws1['H3'] = f"=SUBTOTAL(9,H5:H{ws1.max_row})"




        ws2 = wb.create_sheet('Construction')
        # make tab color red
        ws2.sheet_properties.tabColor = "1072BA"

        ws2['A1'] = "Construction"
        ws2['A2'] = ""
        ws2['A3'] = ""

        headers = []

        for item in construction:
            # if the first 6 characters in item["Blocks"] is equal to "Block " then item['Renamed_block'] is equal to
            # the 7th character only else item['Renamed_block'] is equal to item['Blocks']
            if item['Blocks'][:6] == "Block ":
                item['Renamed_block'] = item['Blocks'][6]
            else:
                item['Renamed_block'] = item['Blocks']

        for key in construction[0]:
            headers.append(key.replace("_", " ").title())
        ws2.append(headers)

        # print("headers", headers)
        for item in construction:

            row = []
            for key in item:
                row.append(item[key])

            ws2.append(row)

        ws2.auto_filter.ref = f"A4:{get_column_letter(ws2.max_column)}{ws2.max_row}"
        ws2.freeze_panes = "A5"

        subtotal_columns = ['I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q']
        # print(ws2.max_column)
        subtotal_columns = [get_column_letter(col) for col in range(9, ws2.max_column + 1)]
        # print(subtotal_columns)

        # format column in range from 9 to ws2.max_column as currency
        for col in range(9, ws2.max_column + 1):
            ws2[f"{get_column_letter(col)}5"].number_format = '#,##0.00'

        for col in subtotal_columns:
            ws2[f"{col}2"] = f"=SUBTOTAL(9,{col}5:{col}{ws2.max_row})"
            ws2[f"{col}2"].number_format = '#,##0.00'

        # make all columns a width of 35
        for col in ws2.columns:
            max_length = 0
            column = col[0].column_letter  # Get the column name
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)
                except Exception:
                    pass
            adjusted_width = (max_length + 2)
            ws2.column_dimensions[column].width = adjusted_width

        # loop through all cells and if the cell in column A is True, then make the row green
        for row in ws2.iter_rows(min_row=5, max_row=ws2.max_row, min_col=1, max_col=1):
            for cell in row:
                if cell.value == True:
                    for inner_row in ws2.iter_rows(min_row=row[0].row, max_row=row[0].row, min_col=1,
                                                   max_col=ws2.max_column):
                        for cells in inner_row:
                            cells.fill = PatternFill(start_color="4CCD99", end_color="4CCD99", fill_type="solid")

        # original_dict = {}  # loop through all cells and if the cell in column A is False, then make the row red
        original_list = []
        converted_dicts = []
        # make a deep copy of the construction list
        # new_construction_list = copy.deepcopy(construction)

        for index, item in enumerate(construction):
            # if index == 0:
            #     print(item)
            original_dict = item
            original_list.append(original_dict)

        # Regular expression pattern to extract amount
        amount_pattern = r"R\s(\d{1,3}(?:\s\d{3})*,\d{2})"

        for ind, original_dict in enumerate(original_list):
            # print(ind, original_dict)
            # print()
            # print(original_dict)
            for key, value in original_dict.items():
                # print("got this far", "key",key, "value",value, "index",ind, "Original_dict", original_dict)
                if key.endswith("Actual") or re.match(r"\d{2}-[A-Za-z]{3}-\d{2}", key):
                    date = key.split()[0] if key.endswith("Actual") else key
                    amount = value
                    if key.endswith("Actual"):
                        actual = True
                    else:
                        actual = False
                    # print()
                    actual_dict = {
                        "Whitebox-Able": original_dict["Whitebox-Able"],
                        "Blocks": original_dict["Blocks"],
                        "Complete Build": original_dict["Complete Build"],
                        "Actual": actual,
                        "Date": date,
                        "Amount": amount,
                        "Renamed_block": original_dict["Renamed_block"]
                    }

                    # print("Actual_dict", actual_dict)
                    # print()
                    converted_dicts.append(actual_dict)

            # print("Gat this far-2", ind )

            # Print the converted dictionaries
        months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

        # print("Gat this far-3", ind)

        for i, d in enumerate(converted_dicts):
            d['Date'] = d['Date'].replace("-", "/")
            # result = calculate_vat_due(d['Date'])
            # d['Vat_due'] = result

            # print(d)
            for index, month in enumerate(months, 1):
                d['Date'] = d['Date'].replace(month, f"{index:02d}")
                # result = calculate_vat_due(d['Date'])
                # d['Vat_due'] = result

            try:
                # Convert the date to a datetime object
                # if d['Date'] is not a datetime object

                if not isinstance(d['Date'], datetime):
                    d['Date'] = datetime.strptime(d['Date'], '%Y/%m/%d')

            except Exception as e:
                # print("XXX",e)
                # Convert the date to a datetime object
                if not isinstance(d['Date'], datetime):
                    d['Date'] = datetime.strptime(d['Date'], '%d/%m/%y')

        for item in converted_dicts:
            # print(item['Date'])
            # convert the date to a string
            date_calc = item['Date'].strftime('%Y/%m/%d')
            # print(date_calc)
            result = calculate_vat_due(date_calc)
            item['Vat_due'] = result
            if not isinstance(item['Vat_due'], datetime):
                item['Vat_due'] = datetime.strptime(item['Vat_due'], '%Y/%m/%d')

        # print(d)
        # print("converted", converted_dicts[0])
        ws2b = wb.create_sheet('Updated Construction')
        # make tab color red
        ws2b.sheet_properties.tabColor = "FF204E"

        ws2b['A1'] = "Updated Construction"

        headers = ["Whitebox-Able", "Blocks", "Complete Build", "Actual", "Date", "Amount", "Renamed_block",
                   "Vat Due", ]
        ws2b.append(headers)
        for item in converted_dicts:
            ws2b.append([item[key] for key in item])

        ws3 = create_sheet(wb, 'Sales', sales, {'B': 44, 'F': 44, 'G': 44, 'H': 44}, {})

        last_row = ws3.max_row
        columns = ['I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S']
        # in row 2 add a subtotal for each column
        for col in columns:
            ws3[f"{col}2"] = f"=SUBTOTAL(9,{col}5:{col}{last_row})"
            ws3[f"{col}2"].number_format = '#,##0.00'

        for row in ws3.iter_rows(min_row=5, max_row=ws3.max_row, min_col=18, max_col=18):
            for cell in row:
                # "=(SUMIFS(Investors!$M:$M,Investors!$E:$E,Sales!$C5,Investors!$O:$O,FALSE)+SUMIFS(Investors!$S:$S,Investors!$E:$E,Sales!$C5,Investors!$O:$O,FALSE))*$F5"
                cell.value = f"=(SUMIFS(Investors!$M:$M,Investors!$E:$E,Sales!$C{cell.row},Investors!$O:$O,FALSE)+SUMIFS(Investors!$S:$S,Investors!$E:$E,Sales!$C{cell.row},Investors!$O:$O,FALSE))*$F{cell.row}"
                cell.number_format = '#,##0.00'

        for row in ws3.iter_rows(min_row=5, max_row=ws3.max_row, min_col=19, max_col=19):
            for cell in row:
                # "=IF(T6=FALSE,Q6-R6,V6)"
                cell.value = f"=IF(T{cell.row}=FALSE,Q{cell.row}-R{cell.row},+V{cell.row})"
                cell.number_format = '#,##0.00'

        for row in ws3.iter_rows(min_row=5, max_row=ws3.max_row, min_col=22, max_col=22):
            for cell in row:
                # cell.value = f"=Q{cell.row}-R{cell.row}"
                cell.number_format = '#,##0.00'

        ws4 = create_sheet(wb, 'Operational Costs', operational_costs, {'A': 44, 'B': 44, 'F': 44, 'G': 44, 'H': 44},
                           {})

        last_row = ws4.max_row
        columns = ['N']

        # in row 2 add a subtotal for each column
        for col in columns:
            # get column number of col

            ws4[f"{col}2"] = f"=SUBTOTAL(9,{col}5:{col}{last_row})"
            ws4[f"{col}2"].number_format = '#,##0.00'

        ws5 = create_sheet(wb, 'Xero', xero, {'A': 44, 'B': 44, 'F': 44, 'G': 44, 'H': 44}, {})

        last_row = ws5.max_row
        columns = ['F', 'G']
        # in row 2 add a subtotal for each column
        for col in columns:
            ws5[f"{col}2"] = f"=SUBTOTAL(9,{col}5:{col}{last_row})"
            ws5[f"{col}2"].number_format = '#,##0.00'
            # format from row 5 to last row as currency

        ws5b = wb.create_sheet('Other Costs')
        # make tab color blue
        ws5b.sheet_properties.tabColor = "1072BA"
        ws5b.title = "Other Costs"
        ws5b['A1'] = "Other Costs"
        ws5b['A2'] = ""
        ws5b['A3'] = ""

        headers_other_costs = []
        for item in other_costs[0]:
            headers.append(item.replace("_", " ").title())
        ws5b.append(headers_other_costs)

        for item in other_costs:

            if not isinstance(item['ReportDate'], datetime):
                item['ReportDate'] = datetime.strptime(item['ReportDate'], '%Y-%m-%d')

            row = []
            for key in item:
                row.append(item[key])
            ws5b.append(row)

        # CASH PROJECTION

        ws6 = wb.create_sheet('Cashflow Projection')
        # make tab color red
        ws6.sheet_properties.tabColor = "FF204E"

        ws6['A1'] = "Cashflow Projection"
        ws6["A1"].font = Font(bold=True, color="31304D", size=14)



        ws6['A2'] = 'Date'
        ws6["A2"].font = Font(bold=True, color="31304D", size=14)
        ws6['B2'] = report_date.strftime('%d-%b-%Y')
        ws6["B2"].font = Font(bold=True, color="31304D", size=14)
        ws6['A3'] = 'NSST Cashflow Projection'
        ws6["A3"].font = Font(bold=True, color="FFFFFF", size=24)
        # fill with olive green
        ws6['A3'].fill = PatternFill(start_color="7F9F80", end_color="FFFFFF", fill_type="solid")
        ws6['A3'].border = Border(left=Side(style='medium'),
                                  right=Side(style='medium'),
                                  top=Side(style='medium'),
                                  bottom=Side(style='medium'))
        # merge cells from A3 to AB3

        # make text white

        ws6['A4'] = ''
        sale_info = []
        # sale_profit_info = []
        for sale in sales:
            if not sale['transferred']:
                insert = {
                    "Development": sale['Category'],
                    "Block": sale['block'],
                    "Complete Build": sale['complete_build'],
                    "Totals": "",
                }
                sale_info.append(insert)
        sale_info = [dict(t) for t in {tuple(d.items()) for d in sale_info}]
        sale_info = sorted(sale_info, key=lambda x: (x['Development'], x['Block']))
        row = []
        # filter through keys and append them to row
        for key in sale_info[0]:
            row.append(key)
        ws6.append(row)

        ws6['A6'] = 'SALES'
        ws6['A6'].font = Font(bold=True, color="31304D", size=14)
        ws6['A5'] = 'PROJECTION'
        ws6['B5'] = ''
        ws6['C5'] = ''
        ws6['D5'] = 'TOTAL'

        ws6['A5'].font = Font(bold=True, color="31304D", size=14)
        ws6['D5'].font = Font(bold=True, color="31304D", size=14)

        toggles_start = ws6.max_row + 1
        row = []
        for sale in sale_info:
            for key in sale:
                row.append(sale[key])
            ws6.append(row)
            # index > 0 then center the row

            row = []

        toggles_end = ws6.max_row
        print("toggles_start", toggles_start)
        print("toggles_end", toggles_end)

        # calculate how many months are between the report date and the 31 December of the report date year
        # print("report_date.month", report_date.month)
        months_to_use = (12 - report_date.month) * 2
        # print("report_date.day", report_date.day)
        # print("report_date.year", report_date.year)
        # print("report_date.month", report_date.month)


        # month_headings = ['F', 'H', 'J', 'L', 'N', 'P', 'R', 'T', 'V', 'X', 'Z', 'AB']
        month_headings = ['F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W',
                          'X', 'Y', 'Z', 'AA', 'AB', 'AC']

        # reduce month_headings to only include the number of months to use
        month_headings = month_headings[:months_to_use]

        ws6.merge_cells(f"A3:{month_headings[len(month_headings) - 1]}3")
        # center the text
        ws6['A3'].alignment = Alignment(horizontal='center', vertical='center')

        ws6[f"{month_headings[len(month_headings) - 1]}1"] = 'C.3.e'
        ws6[f"{month_headings[len(month_headings) - 1]}1"].font = Font(bold=True, size=22)
        # center the above
        ws6[f"{month_headings[len(month_headings) - 1]}1"].alignment = Alignment(horizontal='center', vertical='center')
        # put a border around the above cell
        ws6[f"{month_headings[len(month_headings) - 1]}1"].border = Border(left=Side(style='medium'),
                                   right=Side(style='medium'),
                                   top=Side(style='medium'),
                                   bottom=Side(style='medium'))

        # print("month_headings", len(month_headings))

        row = []
        ws6.append([])
        ws6.append([])
        ws6.append(['VAT ON SALES', "", 1])

        vat_row = ws6.max_row
        ws6[f"A{vat_row}"].font = Font(bold=True, color="31304D", size=14)
        ws6[f"C{vat_row}"].font = Font(bold=True, color="31304D", size=14)
        ws6[f"D{vat_row}"].number_format = '#,##0'
        ws6[f"D{vat_row}"].font = Font(bold=True, color="31304D", size=14)
        ws6[f"D{vat_row}"].fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
        ws6[f"D{vat_row}"].border = ws6[f"A{vat_row}"].border + Border(left=Side(style='medium'),
                                                                       right=Side(style='medium'),
                                                                       top=Side(style='medium'),
                                                                       bottom=Side(style='medium'))

        print("vat_row", vat_row)

        ws6.append([])
        ws6.append(["Income (Profit on Sale)"])
        profit_on_sale = ws6.max_row
        print("profit_on_sale", profit_on_sale)
        ws6.append([])
        ws6.append(["Costs To Complete"])
        ws6.append(["Construction"])
        ws6['A31'].font = Font(bold=True, color="31304D", size=14)
        ws6['A32'].font = Font(bold=True, color="31304D", size=14)
        block_costs_start = ws6.max_row + 1
        # create list called construction_info from sale_info filtered to only include Development = "Heron View"
        # print(sale_info)
        construction_blocks = []
        for item in construction:
            construction_blocks.append(item['Renamed_block'])
        construction_blocks = list(set(construction_blocks))
        construction_blocks = sorted(construction_blocks)
        # print(construction_blocks)
        for block in construction_blocks:
            filtered_construction = [item for item in construction if
                                     item['Renamed_block'] == block and item['Whitebox-Able'] == True]
            # print(filtered_construction)
            if len(filtered_construction) > 0:
                value = filtered_construction[0]['Complete Build']
                ws6.append(["", f"{block}", value])
            else:
                ws6.append(["", f"{block}", 1])
        block_costs_end = ws6.max_row
        print("block_costs_start", block_costs_start)
        print("block_costs_end", block_costs_end)

        ws6.append([])
        ws6.append(["VAT ON CONSTRUCTION", "", 1])
        vat_construction = ws6.max_row
        ws6[f"A{vat_construction}"].font = Font(bold=True, color="31304D", size=14)
        ws6[f"C{vat_construction}"].font = Font(bold=True, color="31304D", size=14)

        ws6.append([])
        ws6.append(["OPERATING EXPENSES", "", 1])
        operating_expenses = ws6.max_row
        ws6[f"A{operating_expenses}"].font = Font(bold=True, color="31304D", size=14)
        ws6[f"C{operating_expenses}"].font = Font(bold=True, color="31304D", size=14)
        # format the above cell as a percentage
        ws6[f"C{operating_expenses}"].number_format = '0%'
        ws6.append([])
        ws6.append(["MONTHLY"])
        monthly = ws6.max_row
        ws6[f"A{monthly}"].font = Font(bold=True, color="31304D", size=14)
        # ws6[f"C{monthly}"].font = Font(bold=True, color="31304D", size=14)

        ws6.append([])
        ws6.append(["RUNNING BALANCE"])
        running = ws6.max_row
        ws6[f"A{running}"].font = Font(bold=True, color="31304D", size=14)

        ws6[f"D{profit_on_sale}"].font = Font(bold=True, color="31304D", size=14)
        ws6[f"D{profit_on_sale}"].number_format = '#,##0'

        ws6[f"D{profit_on_sale}"].border = Border(left=Side(style='medium'),
                                                  right=Side(style='medium'),
                                                  top=Side(style='medium'),
                                                  bottom=Side(style='medium'))

        ws6[f"D{vat_construction}"].fill = PatternFill(start_color="BFEA7C", end_color="BFEA7C", fill_type="solid")
        ws6[f"D{vat_construction}"].number_format = '#,##0'
        ws6[f"D{vat_construction}"].font = Font(bold=True, color="31304D", size=14)
        ws6[f"D{vat_construction}"].border = ws6[f"A{vat_construction}"].border + Border(left=Side(style='medium'),
                                                                                         right=Side(style='medium'),
                                                                                         top=Side(style='medium'),
                                                                                         bottom=Side(style='medium'))

        ws6[f"D{operating_expenses}"].fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
        ws6[f"D{operating_expenses}"].number_format = '#,##0'
        ws6[f"D{operating_expenses}"].font = Font(bold=True, color="31304D", size=14)
        ws6[f"D{operating_expenses}"].border = ws6[f"A{operating_expenses}"].border + Border(left=Side(style='medium'),
                                                                                             right=Side(style='medium'),
                                                                                             top=Side(style='medium'),
                                                                                             bottom=Side(
                                                                                                 style='medium'))

        # ws6[f"D{monthly}"].fill = PatternFill(start_color="BFEA7C", end_color="BFEA7C", fill_type="solid")
        ws6[f"D{monthly}"].number_format = '#,##0'
        ws6[f"D{monthly}"].font = Font(bold=True, color="31304D", size=14)
        ws6[f"D{monthly}"].border = ws6[f"A{monthly}"].border + Border(left=Side(style='medium'),
                                                                       right=Side(style='medium'),
                                                                       top=Side(style='medium'),
                                                                       bottom=Side(style='medium'))

        ws6.conditional_formatting.add(f"D{monthly}",
                                       formatting.rule.CellIsRule(operator='lessThan', formula=['0'],
                                                                  fill=PatternFill(
                                                                      start_color="FFC7CE",
                                                                      end_color="FFC7CE",
                                                                      fill_type="solid")))
        ws6.conditional_formatting.add(f"D{monthly}",
                                       formatting.rule.CellIsRule(operator='greaterThan', formula=['0'],
                                                                  fill=PatternFill(
                                                                      start_color="BFEA7C",
                                                                      end_color="BFEA7C",
                                                                      fill_type="solid")))
        ws6.conditional_formatting.add(f"D{monthly}",
                                       formatting.rule.CellIsRule(operator='equal', formula=['0'],
                                                                  fill=PatternFill(
                                                                      start_color="FFF67E",
                                                                      end_color="FFF67E",
                                                                      fill_type="solid")))

        ws6[f"D{running}"].number_format = '#,##0'
        ws6[f"D{running}"].font = Font(bold=True, color="31304D", size=14)
        ws6[f"D{running}"].border = ws6[f"A{running}"].border + Border(left=Side(style='medium'),
                                                                       right=Side(style='medium'),
                                                                       top=Side(style='medium'),
                                                                       bottom=Side(style='medium'))

        ws6.conditional_formatting.add(f"D{running}",
                                       formatting.rule.CellIsRule(operator='lessThan', formula=['0'],
                                                                  fill=PatternFill(
                                                                      start_color="FFC7CE",
                                                                      end_color="FFC7CE",
                                                                      fill_type="solid")))
        ws6.conditional_formatting.add(f"D{running}",
                                       formatting.rule.CellIsRule(operator='greaterThan', formula=['0'],
                                                                  fill=PatternFill(
                                                                      start_color="BFEA7C",
                                                                      end_color="BFEA7C",
                                                                      fill_type="solid")))
        ws6.conditional_formatting.add(f"D{running}",
                                       formatting.rule.CellIsRule(operator='equal', formula=['0'],
                                                                  fill=PatternFill(
                                                                      start_color="FFF67E",
                                                                      end_color="FFF67E",
                                                                      fill_type="solid")))

        ws6.conditional_formatting.add(f"D{profit_on_sale}",
                                       formatting.rule.CellIsRule(operator='lessThan', formula=['0'],
                                                                  fill=PatternFill(
                                                                      start_color="FFC7CE",
                                                                      end_color="FFC7CE",
                                                                      fill_type="solid")))

        ws6.conditional_formatting.add(f"D{profit_on_sale}",
                                       formatting.rule.CellIsRule(operator='greaterThan', formula=['0'],
                                                                  fill=PatternFill(
                                                                      start_color="BFEA7C",
                                                                      end_color="BFEA7C",
                                                                      fill_type="solid")))

        ws6.conditional_formatting.add(f"D{profit_on_sale}",
                                       formatting.rule.CellIsRule(operator='equal', formula=['0'],
                                                                  fill=PatternFill(
                                                                      start_color="FFF67E",
                                                                      end_color="FFF67E",
                                                                      fill_type="solid")))

        ws6[f"A{profit_on_sale}"].font = Font(bold=True, color="31304D", size=14)
        # ws6[f"D{profit_on_sale}"].fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

        ws6[f"D{toggles_start - 1}"] = f"=SUM(D{toggles_start}:D{toggles_end})"
        ws6[f"D{toggles_start - 1}"].number_format = '#,##0'
        ws6[f"D{toggles_start - 1}"].font = Font(bold=True, color="31304D", size=14)
        ws6[f"E{toggles_start - 1}"] = f"=SUM(E{toggles_start}:E{toggles_end})"
        ws6[f"E{toggles_start - 1}"].number_format = '#,##0'
        ws6[f"E{toggles_start - 1}"].font = Font(bold=True, color="31304D", size=14)
        ws6[f"D{toggles_start - 1}"].fill = PatternFill(start_color="7F9F80", end_color="7F9F80", fill_type="solid")

        ws6[f"D{block_costs_start - 1}"] = f"=SUM(D{block_costs_start}:D{block_costs_end})"
        ws6[f"D{block_costs_start - 1}"].font = Font(bold=True, color="31304D", size=14)
        ws6[f"D{block_costs_start - 1}"].number_format = '#,##0'
        ws6[f"D{block_costs_start - 1}"].fill = PatternFill(start_color="7F9F80", end_color="7F9F80", fill_type="solid")
        # Apply borders to the cell
        ws6[f"D{block_costs_start - 1}"].border = Border(left=Side(style='medium'),
                                                         right=Side(style='medium'),
                                                         top=Side(style='medium'),
                                                         bottom=Side(style='medium'))

        # ws6[f"D{block_costs_start - 1}"].fill = PatternFill(start_color="7F9F80", end_color="7F9F80", fill_type="solid")

        for i in range(toggles_start, toggles_end + 1):
            formula_start = "="
            count_formula_start = "="
            for index, col in enumerate(month_headings):
                if index % 2 == 0:
                    formula_start += f"+{col}{i}"
                else:
                    count_formula_start += f"+{col}{i}"

            ws6[f"D{i}"] = f"{formula_start}"
            ws6[f"D{i}"].number_format = '#,##0'
            ws6[f"D{i}"].font = Font(bold=True, color="A9A9A9", size=14)

            ws6[f"E{i}"] = f"{count_formula_start}"
            ws6[f"E{i}"].number_format = '#,##0'
            ws6[f"E{i}"].font = Font(bold=True, color="A9A9A9", size=14)

        for i in range(vat_row, vat_row + 1):
            formula_start_vat = "="
            for index, col in enumerate(month_headings):
                formula_start = "="
                if index % 2 == 0:
                    formula_start_vat += f"+{col}{i}"

                ws6[f"D{i}"] = f"{formula_start_vat}"

        for i in range(profit_on_sale, profit_on_sale + 1):
            formula_start_vat = "="
            for index, col in enumerate(month_headings):
                formula_start = "="
                if index % 2 == 0:
                    formula_start_vat += f"+{col}{i}"

                ws6[f"D{i}"] = f"{formula_start_vat}"

        for i in range(block_costs_start, block_costs_end + 1):
            formula_start_block = "="
            for index, col in enumerate(month_headings):

                if index % 2 == 0:
                    formula_start_block += f"+{col}{i}"

                    ws6[f"D{i}"] = f"{formula_start_block}"
                    ws6[f"D{i}"].number_format = '#,##0'
                    ws6[f"D{i}"].font = Font(bold=True, color="A9A9A9", size=14)

        for i in range(vat_construction, vat_construction + 1):
            formula_start_block = "="
            for index, col in enumerate(month_headings):

                if index % 2 == 0:
                    formula_start_block += f"+{col}{i}"

                    ws6[f"D{i}"] = f"{formula_start_block}"
                    ws6[f"D{i}"].number_format = '#,##0'
                    ws6[f"D{i}"].font = Font(bold=True, color="31304D", size=14)

        for i in range(operating_expenses, operating_expenses + 1):
            formula_start_block = "="
            for index, col in enumerate(month_headings):

                if index % 2 == 0:
                    formula_start_block += f"+{col}{i}"

                    ws6[f"D{i}"] = f"{formula_start_block}"
                    ws6[f"D{i}"].number_format = '#,##0'
                    ws6[f"D{i}"].font = Font(bold=True, color="31304D", size=14)

        for i in range(monthly, monthly + 1):
            formula_start_block = "="
            for index, col in enumerate(month_headings):

                if index % 2 == 0:
                    formula_start_block += f"+{col}{i}"

                    ws6[f"D{i}"] = f"{formula_start_block}"
                    ws6[f"D{i}"].number_format = '#,##0'
                    ws6[f"D{i}"].font = Font(bold=True, color="31304D", size=14)

        for i in range(running, running + 1):
            # "=SUMIFS(Xero!$G:$G, Xero!$B:$B, 'Cashflow Projection'!$B$2, Xero!$D:$D, "84*")"
            ws6[f"D{i}"] = f"=SUMIFS(Xero!$G:$G, Xero!$B:$B, 'Cashflow Projection'!$B$2, Xero!$D:$D, \"84*\")+8823977"
            ws6[f"D{i}"].number_format = '#,##0'
            ws6[f"D{i}"].font = Font(bold=True, color="31304D", size=14)

        for index, col in enumerate(month_headings):

            ws6[f"{col}{toggles_start - 1}"] = f"=SUM({col}{toggles_start}:{col}{toggles_end})"
            ws6[f"{col}{toggles_start - 1}"].number_format = '#,##0'
            ws6[f"D{toggles_start - 1}"].fill = PatternFill(start_color="D4E7C5", end_color="D4E7C5", fill_type="solid")
            ws6[f"{col}{toggles_start - 1}"].font = Font(bold=True, color="31304D", size=14)
            # apply borders to the cell
            ws6[f"{col}{toggles_start - 1}"].border = ws6[f"A{toggles_start - 1}"].border + Border(
                left=Side(style='medium'),
                right=Side(style='medium'),
                top=Side(style='medium'),
                bottom=Side(style='medium'))

            if index % 2 == 0:
                ws6[f"{col}{toggles_start - 1}"].fill = PatternFill(start_color="D4E7C5", end_color="D4E7C5",
                                                                    fill_type="solid")

                ws6[f"{col}{block_costs_start - 1}"] = f"=SUM({col}{block_costs_start}:{col}{block_costs_end})"
                ws6[f"{col}{block_costs_start - 1}"].number_format = '#,##0'
                ws6[f"{col}{block_costs_start - 1}"].fill = PatternFill(start_color="BFEA7C", end_color="BFEA7C",
                                                                        fill_type="solid")

                ws6[f"{col}{block_costs_start - 1}"].font = Font(bold=True, color="31304D", size=14)

                # apply borders to the cell
                ws6[f"{col}{block_costs_start - 1}"].border = ws6[f"A{block_costs_start - 1}"].border + Border(
                    left=Side(style='medium'),
                    right=Side(style='medium'),
                    top=Side(style='medium'),
                    bottom=Side(style='medium'))

            if index == 0:

                ws6[f"{col}{vat_row}"] = f"=-SUMIFS(Sales!$J:$J,Sales!$E:$E,FALSE, Sales!$U:$U," \
                                         f"\"<=\"&'Cashflow Projection'!{month_headings[index]}$5,Sales!$F:$F,1)*$C{vat_row}"
                ws6[f"{col}{vat_row}"].font = Font(bold=True, color="31304D", size=14)

                ws6[f"{col}{vat_row}"].border = ws6[f"A{vat_row}"].border + Border(left=Side(style='medium'),
                                                                                   right=Side(style='medium'),
                                                                                   top=Side(style='medium'),
                                                                                   bottom=Side(style='medium'))
                # fill equals light red
                ws6[f"{col}{vat_row}"].fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
                ws6[f"{col}{vat_row}"].number_format = '#,##0'

                ws6[
                    f"{col}{profit_on_sale}"] = (f"=SUMIFS(Sales!$S:$S,Sales!$E:$E,FALSE,Sales!$F:$F,1,Sales!$H:$H,"
                                                 f"\"<=\"&'Cashflow Projection'!{month_headings[index]}$5,"
                                                 f"Sales!$H:$H,\">\"&'Cashflow Projection'!$B$2)")
                ws6[f"{col}{profit_on_sale}"].border = ws6[f"A{profit_on_sale}"].border + Border(
                    left=Side(style='medium'),
                    right=Side(style='medium'),
                    top=Side(style='medium'),
                    bottom=Side(style='medium'))
                ws6[f"{col}{profit_on_sale}"].font = Font(bold=True, color="31304D", size=14)

                ws6[f"{col}{profit_on_sale}"].number_format = '#,##0'

                ws6.conditional_formatting.add(f"{col}{profit_on_sale}",
                                               formatting.rule.CellIsRule(operator='lessThan', formula=['0'],
                                                                          fill=PatternFill(
                                                                              start_color="FFC7CE",
                                                                              end_color="FFC7CE",
                                                                              fill_type="solid")))

                ws6.conditional_formatting.add(f"{col}{profit_on_sale}",
                                               formatting.rule.CellIsRule(operator='greaterThan', formula=['0'],
                                                                          fill=PatternFill(
                                                                              start_color="BFEA7C",
                                                                              end_color="BFEA7C",
                                                                              fill_type="solid")))

                ws6.conditional_formatting.add(f"{col}{profit_on_sale}",
                                               formatting.rule.CellIsRule(operator='equal', formula=['0'],
                                                                          fill=PatternFill(
                                                                              start_color="FFF67E",
                                                                              end_color="FFF67E",
                                                                              fill_type="solid")))

                ws6[f"{col}{toggles_start - 2}"] = f"=EOMONTH(EDATE($B$2, 0), 1)"
                ws6[f"{col}{toggles_start - 2}"].number_format = 'dd-mmm-yy'
                ws6[f"{col}{toggles_start - 2}"].font = Font(bold=True, color="31304D", size=14)

                for i in range(toggles_start, toggles_end + 1):
                    ws6[
                        f"{col}{i}"] = (f"=SUMIFS(Sales!$S:$S,Sales!$H:$H,\"<=\"&'Cashflow Projection'!F$5,"
                                        f"Sales!$E:$E,FALSE,Sales!$F:$F,'Cashflow Projection'!$C{i},Sales!$A:$A,"
                                        f"'Cashflow Projection'!$A{i},Sales!$B:$B,'Cashflow Projection'!$B{i},"
                                        f"Sales!$H:$H,\">\"&'Cashflow Projection'!B$2)*$C{i}")
                    ws6[f"{col}{i}"].number_format = '#,##0'
                    ws6[f"{col}{i}"].font = Font(bold=True, color="A9A9A9", size=13)

                for i in range(block_costs_start, block_costs_end + 1):
                    ws6[
                        f"{col}{i}"] = (f"=-(SUMIFS('Updated Construction'!$F:$F,'Updated Construction'!$A:$A,FALSE,"
                                        f"'Updated Construction'!$G:$G,'Cashflow Projection'!$B{i},"
                                        f"'Updated Construction'!$E:$E,\"<=\"&'Cashflow Projection'!"
                                        f"{month_headings[index]}$5,'Updated Construction'!$E:$E,\">\"&'Cashflow "
                                        f"Projection'!$B$2)+(SUMIFS('Updated Construction'!$F:$F,"
                                        f"'Updated Construction'!$A:$A,TRUE,'Updated Construction'!$G:$G,"
                                        f"'Cashflow Projection'!$B{i},'Updated Construction'!$E:$E,\"<=\"&'Cashflow "
                                        f"Projection'!{month_headings[index]}$5,'Updated Construction'!$E:$E,"
                                        f"\">\"&'Cashflow Projection'!$B$2)*$C{i}))*1.15")
                    ws6[f"{col}{i}"].number_format = '#,##0'
                    ws6[f"{col}{i}"].font = Font(bold=True, color="A9A9A9", size=13)

                for i in range(vat_construction, vat_construction + 1):
                    ws6[
                        f"{col}{i}"] = (f"=SUMIFS('Updated Construction'!$F:$F,'Updated Construction'!$C:$C,1,'Updated "
                                        f"Construction'!$H:$H,\">\"&'Cashflow Projection'!$B$2,"
                                        f"'Updated Construction'!$H:$H,\"<=\"&'Cashflow Projection'!"
                                        f"{month_headings[index]}$5)*0.15*$C{i}")
                    ws6[f"{col}{i}"].number_format = '#,##0'
                    ws6[f"{col}{i}"].font = Font(bold=True, color="31304D", size=14)

                    ws6[f"{col}{i}"].fill = PatternFill(start_color="BFEA7C", end_color="BFEA7C",
                                                        fill_type="solid")

                    ws6[f"{col}{i}"].border = ws6[f"A{vat_construction}"].border + Border(
                        left=Side(style='medium'),
                        right=Side(style='medium'),
                        top=Side(style='medium'),
                        bottom=Side(style='medium'))

                for i in range(operating_expenses, operating_expenses + 1):
                    ws6[
                        f"{col}{i}"] = f"=-'Operational Costs'!$N$2*$C{i}"
                    ws6[f"{col}{i}"].number_format = '#,##0'
                    ws6[f"{col}{i}"].font = Font(bold=True, color="31304D", size=14)
                    ws6[f"{col}{i}"].fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE",
                                                        fill_type="solid")

                    ws6[f"{col}{i}"].border = ws6[f"A{vat_construction}"].border + Border(
                        left=Side(style='medium'),
                        right=Side(style='medium'),
                        top=Side(style='medium'),
                        bottom=Side(style='medium'))

                for i in range(monthly, monthly + 1):
                    ws6[
                        f"{col}{i}"] = f"={col}{vat_row}+{col}{profit_on_sale}+{col}{block_costs_start - 1}+{col}{vat_construction}+{col}{operating_expenses}"
                    ws6[f"{col}{i}"].number_format = '#,##0'
                    ws6[f"{col}{i}"].font = Font(bold=True, color="31304D", size=14)
                    ws6[f"{col}{i}"].border = ws6[f"A{vat_construction}"].border + Border(
                        left=Side(style='medium'),
                        right=Side(style='medium'),
                        top=Side(style='medium'),
                        bottom=Side(style='medium'))

                    ws6.conditional_formatting.add(f"{col}{i}",
                                                   formatting.rule.CellIsRule(operator='lessThan', formula=['0'],
                                                                              fill=PatternFill(
                                                                                  start_color="FFC7CE",
                                                                                  end_color="FFC7CE",
                                                                                  fill_type="solid")))
                    ws6.conditional_formatting.add(f"{col}{i}",
                                                   formatting.rule.CellIsRule(operator='greaterThan', formula=['0'],
                                                                              fill=PatternFill(
                                                                                  start_color="BFEA7C",
                                                                                  end_color="BFEA7C",
                                                                                  fill_type="solid")))
                    ws6.conditional_formatting.add(f"{col}{i}",
                                                   formatting.rule.CellIsRule(operator='equal', formula=['0'],
                                                                              fill=PatternFill(
                                                                                  start_color="FFF67E",
                                                                                  end_color="FFF67E",
                                                                                  fill_type="solid")))

                for i in range(running, running + 1):
                    "=D53+F51"
                    ws6[
                        f"{col}{i}"] = f"=D{running}+{col}{monthly}"
                    ws6[f"{col}{i}"].number_format = '#,##0'
                    ws6[f"{col}{i}"].font = Font(bold=True, color="31304D", size=14)

                    # apply borders around the cell
                    ws6[f"{col}{i}"].border = ws6[f"A{running}"].border + Border(
                        left=Side(style='medium'),
                        right=Side(style='medium'),
                        top=Side(style='medium'),
                        bottom=Side(style='medium'))

                    ws6.conditional_formatting.add(f"{col}{i}",
                                                   formatting.rule.CellIsRule(operator='lessThan', formula=['0'],
                                                                              fill=PatternFill(
                                                                                  start_color="FFC7CE",
                                                                                  end_color="FFC7CE",
                                                                                  fill_type="solid")))
                    ws6.conditional_formatting.add(f"{col}{i}",
                                                   formatting.rule.CellIsRule(operator='greaterThan', formula=['0'],
                                                                              fill=PatternFill(
                                                                                  start_color="BFEA7C",
                                                                                  end_color="BFEA7C",
                                                                                  fill_type="solid")))
                    ws6.conditional_formatting.add(f"{col}{i}",
                                                   formatting.rule.CellIsRule(operator='equal', formula=['0'],
                                                                              fill=PatternFill(
                                                                                  start_color="FFF67E",
                                                                                  end_color="FFF67E",
                                                                                  fill_type="solid")))
            elif index == 1:
                for i in range(toggles_start, toggles_end + 1):
                    ws6[
                        f"{col}{i}"] = (f"=COUNTIFS(Sales!$H:$H,\"<=\"&'Cashflow Projection'!F$5,Sales!$H:$H,"
                                        f"\">\"&'Cashflow Projection'!B$2,Sales!$A:$A,'Cashflow Projection'!$A{i},"
                                        f"Sales!$B:$B,'Cashflow Projection'!$B{i},Sales!$F:$F,'Cashflow "
                                        f"Projection'!$C{i},Sales!$E:$E,FALSE)*$C{i}")
                    ws6[f"{col}{i}"].number_format = '#,##0'
                    ws6[f"{col}{i}"].font = Font(bold=True, color="A9A9A9", size=13)

            elif index > 1 and index % 2 == 0:
                ws6[f"{col}{vat_row}"] = (f"=-SUMIFS(Sales!$J:$J,Sales!$E:$E,FALSE, Sales!$U:$U,\"<=\"&'Cashflow "
                                          f"Projection'!{month_headings[index]}$5,Sales!$U:$U,\""
                                          f">\"&'Cashflow Projection'!{month_headings[index - 2]}$5,Sales!$F:$F,"
                                          f"1)*$C{vat_row}")
                ws6[f"{col}{vat_row}"].font = Font(bold=True, color="31304D", size=14)
                ws6[f"{col}{vat_row}"].fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
                ws6[f"{col}{vat_row}"].number_format = '#,##0'
                ws6[
                    f"{col}{profit_on_sale}"] = (f"=SUMIFS(Sales!$S:$S,Sales!$E:$E,FALSE,Sales!$F:$F,1,Sales!$H:$H,"
                                                 f"\"<=\"&'Cashflow Projection'!{month_headings[index]}$5,"
                                                 f"Sales!$H:$H,\">\"&'Cashflow Projection'!"
                                                 f"{month_headings[index - 2]}$5)")
                ws6[f"{col}{profit_on_sale}"].font = Font(bold=True, color="31304D", size=14)

                ws6[f"{col}{profit_on_sale}"].number_format = '#,##0'
                ws6[f"{col}{profit_on_sale}"].border = ws6[f"A{profit_on_sale}"].border + Border(
                    left=Side(style='medium'),
                    right=Side(style='medium'),
                    top=Side(style='medium'),
                    bottom=Side(style='medium'))

                ws6.conditional_formatting.add(f"{col}{profit_on_sale}",
                                               formatting.rule.CellIsRule(operator='lessThan', formula=['0'],
                                                                          fill=PatternFill(
                                                                              start_color="FFC7CE",
                                                                              end_color="FFC7CE",
                                                                              fill_type="solid")))

                ws6.conditional_formatting.add(f"{col}{profit_on_sale}",
                                               formatting.rule.CellIsRule(operator='greaterThan', formula=['0'],
                                                                          fill=PatternFill(
                                                                              start_color="BFEA7C",
                                                                              end_color="BFEA7C",
                                                                              fill_type="solid")))

                ws6.conditional_formatting.add(f"{col}{profit_on_sale}",
                                               formatting.rule.CellIsRule(operator='equal', formula=['0'],
                                                                          fill=PatternFill(
                                                                              start_color="FFF67E",
                                                                              end_color="FFF67E",
                                                                              fill_type="solid")))

                ws6[
                    f"{col}{toggles_start - 2}"] = (f"=EOMONTH(EDATE(${month_headings[index - 2]}${toggles_start - 2}, "
                                                    f"0), 1)")
                ws6[f"{col}{toggles_start - 2}"].font = Font(bold=True, color="31304D", size=14)
                ws6[f"{col}{toggles_start - 2}"].number_format = 'dd-mmm-yy'

                for i in range(toggles_start, toggles_end + 1):
                    ws6[
                        f"{col}{i}"] = f"=SUMIFS(Sales!$S:$S,Sales!$H:$H,\"<=\"&'Cashflow Projection'!{month_headings[index]}$5,Sales!$E:$E,FALSE,Sales!$F:$F,'Cashflow Projection'!$C{i},Sales!$A:$A,'Cashflow Projection'!$A{i},Sales!$B:$B,'Cashflow Projection'!$B{i},Sales!$H:$H,\">\"&'Cashflow Projection'!{month_headings[index - 2]}$5)*$C{i}"
                    ws6[f"{col}{i}"].number_format = '#,##0'
                    ws6[f"{col}{i}"].font = Font(bold=True, color="A9A9A9", size=13)
                    ws6[f"{col}{vat_row}"].border = ws6[f"A{vat_row}"].border + Border(left=Side(style='medium'),
                                                                                       right=Side(style='medium'),
                                                                                       top=Side(style='medium'),
                                                                                       bottom=Side(style='medium'))

                for i in range(block_costs_start, block_costs_end + 1):
                    ws6[
                        f"{col}{i}"] = f"=-(SUMIFS('Updated Construction'!$F:$F,'Updated Construction'!$A:$A,FALSE,'Updated Construction'!$G:$G,'Cashflow Projection'!$B{i},'Updated Construction'!$E:$E," \
                                       f"\"<=\"&'Cashflow Projection'!{month_headings[index]}$5,'Updated Construction'!$E:$E,\">\"&'Cashflow Projection'!{month_headings[index - 2]}$5)+(SUMIFS('Updated Construction'!$F:$F,'Updated Construction'!$A:$A,TRUE,'Updated Construction'!$G:$G,'Cashflow Projection'!$B{i},'Updated Construction'!$E:$E," \
                                       f"\"<=\"&'Cashflow Projection'!{month_headings[index]}$5,'Updated Construction'!$E:$E,\">\"&'Cashflow Projection'!{month_headings[index - 2]}$5)*$C{i}))*1.15"
                    ws6[f"{col}{i}"].number_format = '#,##0'
                    ws6[f"{col}{i}"].font = Font(bold=True, color="A9A9A9", size=13)

                for i in range(vat_construction, vat_construction + 1):
                    ws6[
                        f"{col}{i}"] = f"=SUMIFS('Updated Construction'!$F:$F,'Updated Construction'!$C:$C,1,'Updated Construction'!$H:$H,\">\"&'Cashflow Projection'!{month_headings[index - 2]}$5,'Updated Construction'!$H:$H,\"<=\"&'Cashflow Projection'!{month_headings[index]}$5)*0.15*$C{i}"
                    ws6[f"{col}{i}"].number_format = '#,##0'
                    ws6[f"{col}{i}"].font = Font(bold=True, color="31304D", size=14)

                    ws6[f"{col}{i}"].fill = PatternFill(start_color="BFEA7C", end_color="BFEA7C",
                                                        fill_type="solid")

                    ws6[f"{col}{i}"].border = ws6[f"A{vat_construction}"].border + Border(
                        left=Side(style='medium'),
                        right=Side(style='medium'),
                        top=Side(style='medium'),
                        bottom=Side(style='medium'))

                for i in range(operating_expenses, operating_expenses + 1):
                    "=-'Operational Costs'!$N$2*$C$50"
                    ws6[
                        f"{col}{i}"] = f"=-'Operational Costs'!$N$2*$C{i}"
                    ws6[f"{col}{i}"].number_format = '#,##0'
                    ws6[f"{col}{i}"].font = Font(bold=True, color="31304D", size=14)
                    ws6[f"{col}{i}"].fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE",
                                                        fill_type="solid")

                    ws6[f"{col}{i}"].border = ws6[f"A{vat_construction}"].border + Border(
                        left=Side(style='medium'),
                        right=Side(style='medium'),
                        top=Side(style='medium'),
                        bottom=Side(style='medium'))

                for i in range(monthly, monthly + 1):
                    # "=F27+F29+F32+F48+F50"
                    ws6[
                        f"{col}{i}"] = f"={col}{vat_row}+{col}{profit_on_sale}+{col}{block_costs_start - 1}+{col}{vat_construction}+{col}{operating_expenses}"
                    ws6[f"{col}{i}"].number_format = '#,##0'
                    ws6[f"{col}{i}"].font = Font(bold=True, color="31304D", size=14)

                    ws6[f"{col}{i}"].border = ws6[f"A{vat_construction}"].border + Border(
                        left=Side(style='medium'),
                        right=Side(style='medium'),
                        top=Side(style='medium'),
                        bottom=Side(style='medium'))

                    ws6.conditional_formatting.add(f"{col}{i}",
                                                   formatting.rule.CellIsRule(operator='lessThan', formula=['0'],
                                                                              fill=PatternFill(
                                                                                  start_color="FFC7CE",
                                                                                  end_color="FFC7CE",
                                                                                  fill_type="solid")))
                    ws6.conditional_formatting.add(f"{col}{i}",
                                                   formatting.rule.CellIsRule(operator='greaterThan', formula=['0'],
                                                                              fill=PatternFill(
                                                                                  start_color="BFEA7C",
                                                                                  end_color="BFEA7C",
                                                                                  fill_type="solid")))
                    ws6.conditional_formatting.add(f"{col}{i}",
                                                   formatting.rule.CellIsRule(operator='equal', formula=['0'],
                                                                              fill=PatternFill(
                                                                                  start_color="FFF67E",
                                                                                  end_color="FFF67E",
                                                                                  fill_type="solid")))

                for i in range(running, running + 1):
                    "=F53+H51"
                    ws6[
                        f"{col}{i}"] = f"={month_headings[index - 2]}{running}+{col}{monthly}"
                    ws6[f"{col}{i}"].number_format = '#,##0'
                    ws6[f"{col}{i}"].font = Font(bold=True, color="31304D", size=14)

                    # apply borders around the cell
                    ws6[f"{col}{i}"].border = ws6[f"A{running}"].border + Border(
                        left=Side(style='medium'),
                        right=Side(style='medium'),
                        top=Side(style='medium'),
                        bottom=Side(style='medium'))

                    ws6.conditional_formatting.add(f"{col}{i}",
                                                   formatting.rule.CellIsRule(operator='lessThan', formula=['0'],
                                                                              fill=PatternFill(
                                                                                  start_color="FFC7CE",
                                                                                  end_color="FFC7CE",
                                                                                  fill_type="solid")))
                    ws6.conditional_formatting.add(f"{col}{i}",
                                                   formatting.rule.CellIsRule(operator='greaterThan', formula=['0'],
                                                                              fill=PatternFill(
                                                                                  start_color="BFEA7C",
                                                                                  end_color="BFEA7C",
                                                                                  fill_type="solid")))
                    ws6.conditional_formatting.add(f"{col}{i}",
                                                   formatting.rule.CellIsRule(operator='equal', formula=['0'],
                                                                              fill=PatternFill(
                                                                                  start_color="FFF67E",
                                                                                  end_color="FFF67E",
                                                                                  fill_type="solid")))

            elif index > 2 and index % 2 != 0:
                for i in range(toggles_start, toggles_end + 1):
                    ws6[
                        f"{col}{i}"] = (
                        f"=COUNTIFS(Sales!$H:$H,\"<=\"&'Cashflow Projection'!{month_headings[index - 1]}"
                        f"$5,Sales!$H:$H,\">\"&'Cashflow Projection'!{month_headings[index - 3]}$5,"
                        f"Sales!$A:$A,'Cashflow Projection'!$A{i},Sales!$B:$B,'Cashflow "
                        f"Projection'!$B{i},Sales!$F:$F,'Cashflow Projection'!$C{i},Sales!$E:$E,"
                        f"FALSE)*$C{i}")
                    ws6[f"{col}{i}"].number_format = '#,##0'
                    ws6[f"{col}{i}"].font = Font(bold=True, color="A9A9A9", size=13)



        for i in range(7, 25):
            ws6[f"B{i}"].alignment = Alignment(horizontal='center', vertical='center')
            ws6[f"C{i}"].alignment = Alignment(horizontal='center', vertical='center')

        for i in range(block_costs_start, block_costs_end + 1):

            ws6[f"B{i}"].border = ws6[f"B{i}"].border + Border(left=Side(style='medium'), right=Side(style='medium'),
                                                               top=Side(style='medium'), bottom=Side(style='medium'))
            ws6[f"C{i}"].border = ws6[f"C{i}"].border + Border(left=Side(style='medium'), right=Side(style='medium'),
                                                               top=Side(style='medium'), bottom=Side(style='medium'))

            ws6[f"B{i}"].fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
            ws6[f"C{i}"].fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")


        for row in ws2b.iter_rows(min_row=5, max_row=ws2b.max_row, min_col=3, max_col=3):
            for cell in row:
                cell.value = f"=IF(A{cell.row}=FALSE, 1,SUMIFS('Cashflow Projection'!$C${block_costs_start}:$C${block_costs_end}, 'Cashflow Projection'!$B$33:$B$45, 'Updated Construction'!G{cell.row}))"

        for row in ws6.iter_rows(min_row=toggles_start, max_row=toggles_end, min_col=3, max_col=3):
            for cell in row:
                cell.value = f"=IF(A{cell.row}<>\"Heron View\",1,IF(ISERROR(VLOOKUP($B{cell.row},$B$33:$C$45,2,FALSE)),1,SUMIFS($C$33:$C$45,$B$33:$B$45,B{cell.row})))"

        for row in ws3.iter_rows(min_row=5, max_row=ws3.max_row, min_col=6, max_col=6):
            for cell in row:
                cell.value = f"=SUMIFS('Cashflow Projection'!$C${toggles_start}:$C${toggles_end},'Cashflow Projection'!$B${toggles_start}:$B${toggles_end},Sales!$B{cell.row},'Cashflow Projection'!$A${toggles_start}:$A${toggles_end},Sales!$A{cell.row})"

        for i in range(1, ws6.max_column + 1):
            if i == 1:
                ws6.column_dimensions[get_column_letter(i)].width = 20
            elif i == 2:
                ws6.column_dimensions[get_column_letter(i)].width = 12
            elif i == 3:
                ws6.column_dimensions[get_column_letter(i)].width = 12
            elif i > 3 and i % 2 == 0:
                ws6.column_dimensions[get_column_letter(i)].width = 14
            elif i > 3 and i % 2 != 0:
                ws6.column_dimensions[get_column_letter(i)].width = 7

        # freeze panes at D7
        ws6.freeze_panes = ws6["D6"]

        ws7 = wb.create_sheet('Dashboard')
        # make tab color red
        ws7.sheet_properties.tabColor = "FF204E"

        ws7['C1'] = "C.3.f"
        ws7["C1"].font = Font(bold=True, color="31304D", size=22)
        # center the text
        ws7['C1'].alignment = Alignment(horizontal='center', vertical='center')
        # put a border around the cell
        ws7['C1'].border = Border(left=Side(style='medium'), right=Side(style='medium'), top=Side(style='medium'),
                                  bottom=Side(style='medium'))

        ws7['A2'] = "CASHFLOW DASHBOARD"
        ws7["A2"].font = Font(bold=True, color="31304D", size=14)

        # merge the cells in the first row
        ws7.merge_cells('A2:C2')
        # center the text
        ws7['A2'].alignment = Alignment(horizontal='center', vertical='center')
        ws7['A2'].border = Border(left=Side(style='medium'), top=Side(style='medium'),
                                    bottom=Side(style='medium'))
        ws7['B2'].border = Border( top=Side(style='medium'),
                                  bottom=Side(style='medium'))
        ws7['C2'].border = Border(right=Side(style='medium'), top=Side(style='medium'),
                                  bottom=Side(style='medium'))
        # fill with light grey
        ws7['A2'].fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
        ws7['B2'].fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
        ws7['C2'].fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")


        for i in range(1, ws7.max_column + 1):
            if i == 1:
                ws7.column_dimensions[get_column_letter(i)].width = 55
                # set the font to bold and size 13
                # ws7[f"{get_column_letter(i)}1"].font = Font(bold=False, size=13)
            else:
                ws7.column_dimensions[get_column_letter(i)].width = 25
                # set the font to bold and size 13
                # ws7[f"{get_column_letter(i)}1"].font = Font(bold=True, size=13)
                # set as currency
                # ws7[f"{get_column_letter(i)}1"].number_format = 'R #,##0'

        ws7.append(["", "Projected", "Actual"])
        start = ws7.max_row
        for i in range(start, start + 1):
            # center the text in the cells
            ws7[f"A{i}"].alignment = Alignment(horizontal='center', vertical='center')
            ws7[f"A{i}"].font = Font(bold=True, size=13)
            ws7[f"B{i}"].alignment = Alignment(horizontal='center', vertical='center')
            ws7[f"B{i}"].font = Font(bold=True, size=13)
            ws7[f"C{i}"].alignment = Alignment(horizontal='center', vertical='center')
            ws7[f"C{i}"].font = Font(bold=True, size=13)

        ws7.append([])
        ws7.append(["Actual Transfer Income to hit FNB - Endulini Sold",
                    '=SUMIFS(Sales!$S:$S,Sales!$A:$A,"="&"Endulini",Sales!$D:$D,TRUE,Sales!$E:$E,FALSE)',
                    '=SUMIFS(Sales!$S:$S,Sales!$A:$A,"="&"Endulini",Sales!$D:$D,TRUE,Sales!$E:$E,FALSE)'])
        transfer_endulini = ws7.max_row
        print("Transfer Endulini", transfer_endulini)
        # ws7[f'B{transfer_endulini}'].number_format = '#,##0'

        ws7.append(["Actual Transfer Income to hit FNB - Heron Sold",
                    '=SUMIFS(Sales!$S:$S,Sales!$A:$A,"<>"&"Endulini",Sales!$D:$D,TRUE,Sales!$E:$E,FALSE)',
                    '=SUMIFS(Sales!$S:$S,Sales!$A:$A,"<>"&"Endulini",Sales!$D:$D,TRUE,Sales!$E:$E,FALSE)'])
        transfer_heron = ws7.max_row
        ws7.append(["Projected Transfer Income  - Endulini  Not yet Sold (Profit)",
                    '=SUMIFS(Sales!$S:$S,Sales!$A:$A,"="&"Endulini",Sales!$D:$D,FALSE,Sales!$E:$E,FALSE)'])
        transfer_endulini_profit = ws7.max_row
        ws7.append(["Projected Transfer Income  - Heron Not yet Sold (Profit)", "=SUMIFS(Sales!$S:$S,Sales!$A:$A,\"<>\"&\"Endulini\",Sales!$D:$D,FALSE,Sales!$E:$E,FALSE,Sales!$F:$F,1)"])
        transfer_heron_profit = ws7.max_row
        ws7.append(["Projected Not Allocated asd yet",
                    f"='Cashflow Projection'!D{profit_on_sale}-Dashboard!B{transfer_endulini}-Dashboard!B{transfer_heron}-Dashboard!B{transfer_endulini_profit}-Dashboard!B{transfer_heron_profit}"])
        not_allocated = ws7.max_row
        ws7.append([])
        ws7.append(["Total Income", f"=sum(B{transfer_endulini}:B{not_allocated})",
                    f"=sum(C{transfer_endulini}:C{not_allocated})"])
        total_income = ws7.max_row
        ws7.append([])
        ws7.append(["Momentum funds avaialble to draw",
                    "=SUMIFS(Xero!$G:$G,Xero!$E:$E,\"Momentum Investors Account RU502229930\",Xero!$B:$B,'Cashflow Projection'!$B$2)",
                    "=SUMIFS(Xero!$G:$G,Xero!$E:$E,\"Momentum Investors Account RU502229930\",Xero!$B:$B,'Cashflow Projection'!$B$2)"])
        momentum_funds = ws7.max_row
        ws7.append(["FNB Bank",
                    f"=SUMIFS(Xero!$G:$G, Xero!$B:$B, 'Cashflow Projection'!$B$2, Xero!$D:$D, \"84*\")-B{momentum_funds}",
                    f"=SUMIFS(Xero!$G:$G, Xero!$B:$B, 'Cashflow Projection'!$B$2, Xero!$D:$D, \"84*\")-B{momentum_funds}"])
        fnb_bank = ws7.max_row
        ws7.append(["Deposit made", 5598623.87, 5598623.87])
        ws7.append(["Deposit made",  3143782 , 0])
        deposit_made = ws7.max_row
        ws7.append(["Momentum interest to be earned", 0, 0])
        momentum_interest = ws7.max_row
        ws7.append(["Investor funds to hit Momentum Account to draw", 0, 0])
        investor_funds = ws7.max_row
        ws7.append([])
        ws7.append(
            ["Total Cash", f"=sum(B{momentum_funds}:B{investor_funds})", f"=sum(B{momentum_funds}:B{investor_funds})"])
        total_funds = ws7.max_row
        ws7.append([])
        ws7.append(
            ["Total funds available to draw", f"=B{total_funds}+B{total_income}", f"=C{total_funds}+C{total_income}"])
        total_funds_draw = ws7.max_row
        ws7.append([])
        ws7.append(["Cost to complete Heron Projects (CPC)", f"='Cashflow Projection'!D{block_costs_start - 1}",
                    f"='Cashflow Projection'!D{block_costs_start - 1}"])
        cpc = ws7.max_row
        ws7.append(["Company Running Costs", f"='Cashflow Projection'!D{operating_expenses}",
                    f"='Cashflow Projection'!D{operating_expenses}"])
        company_running_costs = ws7.max_row
        ws7.append(["VAT Payable", f"='Cashflow Projection'!D27+'Cashflow Projection'!D47",
                     f"='Cashflow Projection'!D27+'Cashflow Projection'!D47"])
        vat_payable = ws7.max_row

        ws7.append([])
        ws7.append(["Project Costs", f"=B{cpc}+B{company_running_costs}+B{vat_payable}", f"=C{cpc}+C{company_running_costs}+C{vat_payable}"])
        project_costs = ws7.max_row
        ws7.append([])
        ws7.append(["NETTO EFFECT", f"=B{total_funds_draw}+B{project_costs}", f"=C{total_funds_draw}+C{project_costs}"])
        nett_effect = ws7.max_row

        style = '#,##0.00'
        rows_chosen = ['A', 'B', 'C']
        for i in range(transfer_endulini, ws7.max_row + 1):
            for index, row in enumerate(rows_chosen):
                if index == 0:
                    ws7[f"{row}{i}"].font = Font(bold=False, size=14)
                elif index == 1 or index == 2:
                    ws7[f"{row}{i}"].font = Font(bold=False, size=14)
                    ws7[f"{row}{i}"].number_format = style

        ws7[f"A{total_income}"].font = Font(bold=True, size=14)
        ws7[f"B{total_income}"].font = Font(bold=True, size=14)
        ws7[f"B{total_income}"].border = Border(top=Side(style='medium'), bottom=Side(style='double'))
        ws7[f"C{total_income}"].font = Font(bold=True, size=14)
        ws7[f"C{total_income}"].border = Border(top=Side(style='medium'), bottom=Side(style='double'))

        ws7[f"A{total_funds}"].font = Font(bold=True, size=14)
        ws7[f"B{total_funds}"].font = Font(bold=True, size=14)
        ws7[f"B{total_funds}"].border = Border(top=Side(style='medium'), bottom=Side(style='double'))
        ws7[f"C{total_funds}"].font = Font(bold=True, size=14)
        ws7[f"C{total_funds}"].border = Border(top=Side(style='medium'), bottom=Side(style='double'))

        ws7[f"A{total_funds_draw}"].font = Font(bold=True, size=14)
        ws7[f"B{total_funds_draw}"].font = Font(bold=True, size=14)
        ws7[f"B{total_funds_draw}"].border = Border(top=Side(style='medium'), bottom=Side(style='double'))
        ws7[f"C{total_funds_draw}"].font = Font(bold=True, size=14)
        ws7[f"C{total_funds_draw}"].border = Border(top=Side(style='medium'), bottom=Side(style='double'))

        ws7[f"A{project_costs}"].font = Font(bold=True, size=14)
        ws7[f"B{project_costs}"].font = Font(bold=True, size=14)
        ws7[f"B{project_costs}"].border = Border(top=Side(style='medium'), bottom=Side(style='double'))
        ws7[f"C{project_costs}"].font = Font(bold=True, size=14)
        ws7[f"C{project_costs}"].border = Border(top=Side(style='medium'), bottom=Side(style='double'))

        ws7[f"A{nett_effect}"].font = Font(bold=True, size=14)
        ws7[f"B{nett_effect}"].font = Font(bold=True, size=14)
        ws7[f"B{nett_effect}"].border = Border(top=Side(style='medium'), bottom=Side(style='double'))
        ws7[f"C{nett_effect}"].font = Font(bold=True, size=14)
        ws7[f"C{nett_effect}"].border = Border(top=Side(style='medium'), bottom=Side(style='double'))

        ws['T4'] = "Interest to Date"



        "=IF(J5<>"",M5*N5/365/100*('Cashflow Projection'!$B$2-Investors!J5),0)"
        inv_last = ws.max_row
        for i in range(5, inv_last + 1):
            ws[f"T{i}"] = f"=IF(J{i}<>\"\",M{i}*N{i}/365/100*('Cashflow Projection'!$B$2-Investors!J{i}),0)"
            ws[f"T{i}"].number_format = '#,##0'

        ws6[f"D{running}"] = f"==Dashboard!B20"

        ws8 = wb.create_sheet('Heron')
        # make tab color red
        ws8.sheet_properties.tabColor = "FF204E"
        ws8['A1'] = "HERON P&L - Based on Cash Projection and Dashboard"
        ws8["A1"].font = Font(bold=True, color="31304D", size=14)
        ws8.append([])
        ws8.append([])
        ws8.append(["Heron Fields (Pty) Ltd", "Heron Fields"])
        fields = ws8.max_row
        ws8.append(["Heron View (Pty) Ltd", "Heron View"])
        view = ws8.max_row
        ws8.append(["Cape Projects Construction (Pty) Ltd"])
        cpc = ws8.max_row
        ws8.append([])
        ws8.append([])

        reporting_months = []

        # filter Xero and only include where ReportTitle = "Heron Fields (Pty) Ltd" or ReportTitle = "Heron View (Pty) Ltd" using list comprehension
        xero_heron = [x for x in xero if
                      x['ReportTitle'] == "Heron Fields (Pty) Ltd" or x['ReportTitle'] == "Heron View (Pty) Ltd"]
        # xero_heron to only include where Category = "Income" or Category = "Expenses"
        xero_heron = [x for x in xero_heron if x['Category'] == "Revenue" or x['Category'] == "Expenses"]
        # print("Xero Heron", xero_heron)
        income_accounts = []
        expense_accounts = []
        # print("Xero Heron", xero_heron[0])
        # print()
        for item in xero_heron:
            reporting_months.append(item['ReportDate'])
            if item['Category'] == "Revenue":
                income_accounts.append(item['AccountName'])
            elif item['Category'] == "Expenses":
                expense_accounts.append(item['AccountName'])
        reporting_months = list(set(reporting_months))

        income_accounts = list(set(income_accounts))
        # sort the income accounts
        income_accounts.sort()
        expense_accounts = list(set(expense_accounts))
        # sort the expense accounts
        expense_accounts.sort()

        expense_accounts.append("Heron View Land")
        # print("Expense Accounts", expense_accounts)

        # sort the reporting months
        reporting_months.sort()
        # Insert "Month", "","" at the beginning of the list
        reporting_months.insert(0, "Month")
        reporting_months.insert(1, "")
        reporting_months.insert(2, "")
        # add 12 months to the list

        ws8.append(reporting_months)
        months_report = ws8.max_row
        # make months report row bold
        for i in range(1, ws8.max_column + 1):
            ws8[f"{get_column_letter(i)}{months_report}"].font = Font(bold=True, size=14)
            ws8[f"{get_column_letter(i)}{months_report}"].alignment = Alignment(horizontal='center', vertical='center')
            ws8[f"{get_column_letter(i)}{months_report}"].fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE",
                                                                             fill_type="solid")
            if i > 3:
                # format as "YYYY-MM-DD"
                ws8[f"{get_column_letter(i)}{months_report}"].number_format = 'yyyy-mm-dd'

        for i in range(ws8.max_column + 1, ws8.max_column + 19):
            # f"=EOMONTH(EDATE($B$2, 0), 1)"
            ws8[
                f"{get_column_letter(i)}{months_report}"] = f"=EOMONTH(EDATE({get_column_letter(i - 1)}{months_report}, 0),1)"
            ws8[f"{get_column_letter(i)}{months_report}"].number_format = 'yyyy-mm-dd'
            ws8[f"{get_column_letter(i)}{months_report}"].font = Font(bold=True, size=14)
            ws8[f"{get_column_letter(i)}{months_report}"].alignment = Alignment(horizontal='center', vertical='center')
            ws8[f"{get_column_letter(i)}{months_report}"].fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE",
                                                                             fill_type="solid")

        for index, i in enumerate(range(ws8.max_column + 1, ws8.max_column + 3)):
            if index == 0:
                ws8[f"{get_column_letter(i)}{months_report}"] = "Total"
            else:
                ws8[f"{get_column_letter(i)}{months_report}"] = "NSST"
            ws8[f"{get_column_letter(i)}{months_report}"].font = Font(bold=True, size=14)
            ws8[f"{get_column_letter(i)}{months_report}"].alignment = Alignment(horizontal='center', vertical='center')
            ws8[f"{get_column_letter(i)}{months_report}"].fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE",
                                                                             fill_type="solid")

        ws8.append([])
        ws8.append(['Income'])
        # make the income row bold and a bit bigger
        ws8[f"A{ws8.max_row}"].font = Font(bold=True, size=14)
        ws8.append([])
        for index, item in enumerate(income_accounts):
            ws8.append([item])
            ws8[f"A{ws8.max_row}"].font = Font(bold=False, size=12)
            if index == 0:
                income_start = ws8.max_row
        income_end = ws8.max_row
        print("Income Start", income_start)
        print("Income End", income_end)

        ws8.append([])
        ws8.append(['Total Income'])
        ws8[f"A{ws8.max_row}"].font = Font(bold=True, size=14)

        for i in range(4, ws8.max_column + 1):
            ws8[
                f"{get_column_letter(i)}{ws8.max_row}"] = f"=sum({get_column_letter(i)}{income_start}:{get_column_letter(i)}{income_end})"
            ws8[f"{get_column_letter(i)}{ws8.max_row}"].number_format = '#,##0.00'
            ws8[f"{get_column_letter(i)}{ws8.max_row}"].font = Font(bold=True, size=14)
            ws8[f"{get_column_letter(i)}{ws8.max_row}"].border = Border(top=Side(style='medium'),
                                                                        bottom=Side(style='double'))

        ws8.append([])
        ws8.append(['Expenses'])
        # make the income row bold and a bit bigger
        ws8[f"A{ws8.max_row}"].font = Font(bold=True, size=14)
        ws8.append([])

        for index, item in enumerate(expense_accounts):
            ws8.append([item])
            ws8[f"A{ws8.max_row}"].font = Font(bold=False, size=12)
            if index == 0:
                expense_start = ws8.max_row

        cpc_costs = ["COS - Heron Fields - Printing & Stationary", "COS - Heron Fields - Construction",
                     "COS - Heron Fields - P & G", "COS - Heron View - P&G", "COS - Heron View - Construction",
                     "COS - Heron View - Printing & Stationary"]
        # sort the cpc costs
        cpc_costs.sort()
        for index, item in enumerate(cpc_costs):
            ws8.append([item])
            ws8[f"A{ws8.max_row}"].font = Font(bold=False, size=12)
            if index == 0:
                cpc_start = ws8.max_row
        cpc_end = ws8.max_row

        other_costs_headers = []
        for item in other_costs:
            other_costs_headers.append(item['AccountName'])

        other_costs_headers = list(set(other_costs_headers))
        other_costs_headers.sort()

        for index, item in enumerate(other_costs_headers):
            print("Item", item)
            ws8.append([item])
            ws8[f"A{ws8.max_row}"].font = Font(bold=False, size=12)
            if index == 0:
                other_costs_start = ws8.max_row
        other_costs_end = ws8.max_row

        print("CPC Start", cpc_start)
        print("CPC End", cpc_end)

        print("Other Costs Start", other_costs_start)
        print("Other Costs End", other_costs_end)

        # other_costs_to_input = []
        # for index, item in enumerate(other_costs):

        expense_end = ws8.max_row
        print("Expense Start", expense_start)
        print("Expense End", expense_end)

        ws8.append([])
        ws8.append(['Total Expenses'])
        ws8[f"A{ws8.max_row}"].font = Font(bold=True, size=14)

        for i in range(4, ws8.max_column + 1):
            ws8[
                f"{get_column_letter(i)}{ws8.max_row}"] = f"=sum({get_column_letter(i)}{expense_start}:{get_column_letter(i)}{expense_end})"
            ws8[f"{get_column_letter(i)}{ws8.max_row}"].number_format = '#,##0.00'
            ws8[f"{get_column_letter(i)}{ws8.max_row}"].font = Font(bold=True, size=14)
            ws8[f"{get_column_letter(i)}{ws8.max_row}"].border = Border(top=Side(style='medium'),
                                                                        bottom=Side(style='double'))

        ws8.append([])
        ws8.append([])
        ws8.append(['Net Profit'])
        ws8[f"A{ws8.max_row}"].font = Font(bold=True, size=14)

        for i in range(4, ws8.max_column + 1):
            ws8[
                f"{get_column_letter(i)}{ws8.max_row}"] = f"={get_column_letter(i)}{income_end + 2}-{get_column_letter(i)}{expense_end + 2}"
            ws8[f"{get_column_letter(i)}{ws8.max_row}"].number_format = '#,##0.00'
            ws8[f"{get_column_letter(i)}{ws8.max_row}"].font = Font(bold=True, size=14)
            ws8[f"{get_column_letter(i)}{ws8.max_row}"].border = Border(top=Side(style='medium'),
                                                                        bottom=Side(style='double'))

        for i in range(1, ws8.max_column + 1):
            if i == 1:
                ws8.column_dimensions[get_column_letter(i)].width = 35
            elif i > 3:
                ws8.column_dimensions[get_column_letter(i)].width = 18

        for i in range(4, ws8.max_column - 1):
            for row in range(income_start, income_end + 1):
                # print(ws8[f"A{row}"].value)
                if ws8[f"A{row}"].value == "Sales - Heron View Sales":
                    ws8[
                        f"{get_column_letter(i)}{row}"].value = f"=-SUMIFS(Xero!$F:$F,Xero!$B:$B,Heron!{get_column_letter(i)}$9,Xero!$A:$A,Heron!$A$4,Xero!$E:$E,Heron!$A{row})-SUMIFS(Xero!$F:$F,Xero!$B:$B,Heron!{get_column_letter(i)}$9,Xero!$A:$A,Heron!$A$5,Xero!$E:$E,Heron!$A{row})+SUMIFS(Sales!$K:$K,Sales!$E:$E,FALSE,Sales!$A:$A,Heron!$B$5,Sales!$H:$H,\"<=\"&Heron!{get_column_letter(i)}$9,Sales!$H:$H,\">\"&Heron!{get_column_letter(i - 1)}$9,Sales!$F:$F,1)"
                    ws8[f"{get_column_letter(i)}{row}"].number_format = '#,##0.00'
                    ws8[f"{get_column_letter(i)}{row}"].font = Font(bold=False, size=12)
                elif ws8[f"A{row}"].value == "Sales - Heron Fields":
                    ws8[
                        f"{get_column_letter(i)}{row}"].value = f"=-SUMIFS(Xero!$F:$F,Xero!$B:$B,Heron!{get_column_letter(i)}$9,Xero!$A:$A,Heron!$A$4,Xero!$E:$E,Heron!$A{row})-SUMIFS(Xero!$F:$F,Xero!$B:$B,Heron!{get_column_letter(i)}$9,Xero!$A:$A,Heron!$A$5,Xero!$E:$E,Heron!$A{row})+SUMIFS(Sales!$K:$K,Sales!$E:$E,FALSE,Sales!$A:$A,Heron!$B$4,Sales!$H:$H,\"<=\"&Heron!{get_column_letter(i)}$9,Sales!$H:$H,\">\"&Heron!{get_column_letter(i - 1)}$9,Sales!$F:$F,1)"
                    ws8[f"{get_column_letter(i)}{row}"].number_format = '#,##0.00'
                    ws8[f"{get_column_letter(i)}{row}"].font = Font(bold=False, size=12)
                else:
                    ws8[
                        f"{get_column_letter(i)}{row}"].value = f"=-SUMIFS(Xero!$F:$F,Xero!$B:$B,Heron!{get_column_letter(i)}$9,Xero!$A:$A,Heron!$A$4,Xero!$E:$E,Heron!$A{row})-SUMIFS(Xero!$F:$F,Xero!$B:$B,Heron!{get_column_letter(i)}$9,Xero!$A:$A,Heron!$A$5,Xero!$E:$E,Heron!$A{row})"
                    ws8[f"{get_column_letter(i)}{row}"].number_format = '#,##0.00'
                    ws8[f"{get_column_letter(i)}{row}"].font = Font(bold=False, size=12)
            for row in range(expense_start, cpc_start):
                if ws8[f"A{row}"].value == "COS - Commission HV Units":
                    ws8[
                        f"{get_column_letter(i)}{row}"].value = f"=SUMIFS(Xero!$F:$F,Xero!$B:$B,Heron!{get_column_letter(i)}$9,Xero!$A:$A,Heron!$A$4,Xero!$E:$E,Heron!$A{row})+SUMIFS(Xero!$F:$F,Xero!$B:$B,Heron!{get_column_letter(i)}$9,Xero!$A:$A,Heron!$A$5,Xero!$E:$E,Heron!$A{row})+SUMIFS(Sales!$O:$O,Sales!$E:$E,FALSE,Sales!$A:$A,Heron!$B$5,Sales!$H:$H,\"<=\"&Heron!{get_column_letter(i)}$9,Sales!$H:$H,\">\"&Heron!{get_column_letter(i - 1)}$9,Sales!$F:$F,1)/1.15"
                    ws8[f"{get_column_letter(i)}{row}"].number_format = '#,##0.00'
                    ws8[f"{get_column_letter(i)}{row}"].font = Font(bold=False, size=12)

                elif ws8[f"A{row}"].value == "COS - Commission HF Units":
                    ws8[
                        f"{get_column_letter(i)}{row}"].value = f"=SUMIFS(Xero!$F:$F,Xero!$B:$B,Heron!{get_column_letter(i)}$9,Xero!$A:$A,Heron!$A$4,Xero!$E:$E,Heron!$A{row})+SUMIFS(Xero!$F:$F,Xero!$B:$B,Heron!{get_column_letter(i)}$9,Xero!$A:$A,Heron!$A$5,Xero!$E:$E,Heron!$A{row})+SUMIFS(Sales!$O:$O,Sales!$E:$E,FALSE,Sales!$A:$A,Heron!$B$4,Sales!$H:$H,\"<=\"&Heron!{get_column_letter(i)}$9,Sales!$H:$H,\">\"&Heron!{get_column_letter(i - 1)}$9,Sales!$F:$F,1)/1.15"
                    ws8[f"{get_column_letter(i)}{row}"].number_format = '#,##0.00'
                    ws8[f"{get_column_letter(i)}{row}"].font = Font(bold=False, size=12)
                elif ws8[f"A{row}"].value == "COS - Legal Fees":
                    ws8[
                        f"{get_column_letter(i)}{row}"].value = f"=SUMIFS(Xero!$F:$F,Xero!$B:$B,Heron!{get_column_letter(i)}$9,Xero!$A:$A,Heron!$A$4,Xero!$E:$E,Heron!$A{row})+SUMIFS(Xero!$F:$F,Xero!$B:$B,Heron!{get_column_letter(i)}$9,Xero!$A:$A,Heron!$A$5,Xero!$E:$E,Heron!$A{row})+SUMIFS(Sales!$L:$L,Sales!$E:$E,FALSE,Sales!$A:$A,\"<>\"&\"Endulini\",Sales!$H:$H,\"<=\"&Heron!{get_column_letter(i)}$9,Sales!$H:$H,\">\"&Heron!{get_column_letter(i - 1)}$9,Sales!$F:$F,1)+SUMIFS(Sales!$M:$M,Sales!$E:$E,FALSE,Sales!$A:$A,\"<>\"&\"Endulini\",Sales!$H:$H,\"<=\"&Heron!{get_column_letter(i)}$9,Sales!$H:$H,\">\"&Heron!{get_column_letter(i - 1)}$9,Sales!$F:$F,1)+SUMIFS(Sales!$N:$N,Sales!$E:$E,FALSE,Sales!$A:$A,\"<>\"&\"Endulini\",Sales!$H:$H,\"<=\"&Heron!{get_column_letter(i)}$9,Sales!$H:$H,\">\"&Heron!{get_column_letter(i - 1)}$9,Sales!$F:$F,1)+SUMIFS(Sales!$P:$P,Sales!$E:$E,FALSE,Sales!$A:$A,\"<>\"&\"Endulini\",Sales!$H:$H,\"<=\"&Heron!{get_column_letter(i)}$9,Sales!$H:$H,\">\"&Heron!{get_column_letter(i - 1)}$9,Sales!$F:$F,1)"
                    ws8[f"{get_column_letter(i)}{row}"].number_format = '#,##0.00'
                    ws8[f"{get_column_letter(i)}{row}"].font = Font(bold=False, size=12)
                elif ws8[f"A{row}"].value == "Interest Paid - Investors @ 18%":
                    # "+SUMIFS(Xero!$F:$F,Xero!$B:$B,Heron!D$9,Xero!$A:$A,Heron!$A$5,Xero!$E:$E,Heron!$A73)+SUMIFS(Investors!$S:$S,Investors!$G:$G,FALSE,Investors!$O:$O,FALSE,Investors!$P:$P,FALSE,Investors!$K:$K,"<="&Heron!D$9,Investors!$K:$K,">"&Heron!C$9)"
                    ws8[
                        f"{get_column_letter(i)}{row}"].value = f"=SUMIFS(Xero!$F:$F,Xero!$B:$B,Heron!{get_column_letter(i)}$9,Xero!$A:$A,Heron!$A$4,Xero!$E:$E,Heron!$A{row})+SUMIFS(Xero!$F:$F,Xero!$B:$B,Heron!{get_column_letter(i)}$9,Xero!$A:$A,Heron!$A$5,Xero!$E:$E,Heron!$A{row})+SUMIFS(Investors!$S:$S,Investors!$G:$G,FALSE,Investors!$O:$O,FALSE,Investors!$P:$P,FALSE,Investors!$K:$K,\"<=\"&Heron!{get_column_letter(i)}$9,Investors!$K:$K,\">\"&Heron!{get_column_letter(i - 1)}$9)"
                    ws8[f"{get_column_letter(i)}{row}"].number_format = '#,##0.00'
                    ws8[f"{get_column_letter(i)}{row}"].font = Font(bold=False, size=12)
                else:
                    ws8[
                        f"{get_column_letter(i)}{row}"].value = f"=SUMIFS(Xero!$F:$F,Xero!$B:$B,Heron!{get_column_letter(i)}$9,Xero!$A:$A,Heron!$A$4,Xero!$E:$E,Heron!$A{row})+SUMIFS(Xero!$F:$F,Xero!$B:$B,Heron!{get_column_letter(i)}$9,Xero!$A:$A,Heron!$A$5,Xero!$E:$E,Heron!$A{row})"
                    ws8[f"{get_column_letter(i)}{row}"].number_format = '#,##0.00'
                    ws8[f"{get_column_letter(i)}{row}"].font = Font(bold=False, size=12)

            for row in range(cpc_start, cpc_end + 1):
                if ws8[f"A{row}"].value == "COS - Heron View - Construction":
                    ws8[
                        f"{get_column_letter(i)}{row}"].value = f"=SUMIFS(Xero!$F:$F,Xero!$B:$B,Heron!{get_column_letter(i)}$9,Xero!$A:$A,Heron!$A$6,Xero!$E:$E,Heron!$A{row})+SUMIFS('Updated Construction'!$F:$F,'Updated Construction'!$C:$C,1,'Updated Construction'!$E:$E,\"<=\"&Heron!{get_column_letter(i)}$9,'Updated Construction'!$E:$E,\">\"&Heron!{get_column_letter(i - 1)}$9,'Updated Construction'!$D:$D,FALSE)"
                    ws8[f"{get_column_letter(i)}{row}"].number_format = '#,##0.00'
                    ws8[f"{get_column_letter(i)}{row}"].font = Font(bold=False, size=12)
                    # fill with light grey
                    ws8[f"{get_column_letter(i)}{row}"].fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3",
                                                                           fill_type="solid")
                else:
                    ws8[
                        f"{get_column_letter(i)}{row}"].value = f"=SUMIFS(Xero!$F:$F,Xero!$B:$B,Heron!{get_column_letter(i)}$9,Xero!$A:$A,Heron!$A$6,Xero!$E:$E,Heron!$A{row})"
                    ws8[f"{get_column_letter(i)}{row}"].number_format = '#,##0.00'
                    ws8[f"{get_column_letter(i)}{row}"].font = Font(bold=False, size=12)
                    # fill with light grey
                    ws8[f"{get_column_letter(i)}{row}"].fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3",
                                                                           fill_type="solid")

            for row in range(other_costs_start, other_costs_end + 1):
                "=SUMIFS('Other Costs'!$D:$D,'Other Costs'!$A:$A,Heron!$A106,'Other Costs'!$C:$C," <= "&Heron!D$9,'Other Costs'!$C:$C," > "&Heron!C$9)"
                ws8[
                    f"{get_column_letter(i)}{row}"].value = f"=SUMIFS('Other Costs'!$D:$D,'Other Costs'!$A:$A,Heron!$A{row},'Other Costs'!$C:$C,\"<=\"&Heron!{get_column_letter(i)}$9,'Other Costs'!$C:$C,\">\"&Heron!{get_column_letter(i - 1)}$9)"
                ws8[f"{get_column_letter(i)}{row}"].number_format = '#,##0.00'
                ws8[f"{get_column_letter(i)}{row}"].font = Font(bold=False, size=12)
                # fill with light orange
                ws8[f"{get_column_letter(i)}{row}"].fill = PatternFill(start_color="FFD966", end_color="FFD966",
                                                                       fill_type="solid")

        for i in range(ws8.max_column - 1, ws8.max_column):
            # print("I", i)
            for row in range(income_start, income_end + 1):
                ws8[
                    f"{get_column_letter(i)}{row}"].value = f"=sum(D{row}:{get_column_letter(i - 1)}{row})"
                ws8[f"{get_column_letter(i)}{row}"].number_format = '#,##0.00'
                ws8[f"{get_column_letter(i)}{row}"].font = Font(bold=False, size=12)

            for row in range(expense_start, expense_end + 1):
                ws8[
                    f"{get_column_letter(i)}{row}"].value = f"=sum(D{row}:{get_column_letter(i - 1)}{row})"
                ws8[f"{get_column_letter(i)}{row}"].number_format = '#,##0.00'
                ws8[f"{get_column_letter(i)}{row}"].font = Font(bold=False, size=12)

        # freeze panes at cell B10
        ws8.freeze_panes = ws8['B10']

        ws9 = wb.create_sheet('NSST Print')
        # make tab color red
        ws9.sheet_properties.tabColor = "FF204E"

        ws9.append(["", "", "", "", "C.3_d"])
        # Merger the first 4 cells in the first row
        ws9.merge_cells('A1:D1')
        # make D1 bold, make the font 20 and put a border around it
        ws9["E1"].font = Font(bold=True, size=22)
        # center the text
        ws9["E1"].alignment = Alignment(horizontal='center', vertical='center')
        ws9["E1"].border = Border(top=Side(style='medium'), bottom=Side(style='medium'), left=Side(style='medium'),
                                  right=Side(style='medium'))

        ws9.append(["NSST HERON PROJECT REPORT"])
        ws9.append(["Report Date", report_date])
        # format cell B3 as date
        ws9["B3"].number_format = 'dd MMMM yyyy'
        ws9.append(["Development", "Heron Fields and Heron View"])
        ws9.append(["CAPITAL"])
        ws9.append(["Total Investment capital to be raised (Estimated)", 236217976.39])
        ws9.append(["Total Investment capital received", "=SUMIFS(Investors!$M:$M,Investors!$P:$P,FALSE,Investors!$I:$I,\"<=\"&'NSST Print'!B3)"])
        ws9.append(["Total Funds Drawn Down into Development", "=SUMIFS(Investors!$M:$M,Investors!$J:$J,\"<=\"&'NSST Print'!B3)"])
        ws9.append(["Momentum Investment Account", "=Dashboard!B13"])
        ws9.append(["Capital not Raised", "=B6-B7-B11"])
        ws9.append(["Available to be raised (Estimated)", "=SUMIFS(Opportunities!$E:$E,Opportunities!$D:$D,FALSE,Opportunities!$C:$C,FALSE)-SUMIFS(Opportunities!$H:$H,Opportunities!$D:$D,FALSE,Opportunities!$C:$C,FALSE)"])
        ws9.append(["Capital repaid", "=SUMIFS(Investors!$M:$M,Investors!$G:$G,TRUE,Investors!$K:$K,\"<=\"&'NSST Print'!B3)-(SUMIFS(Investors!$M:$M,Investors!$G:$G,FALSE,Investors!$O:$O,TRUE,Investors!$K:$K,\"<=\"&'NSST Print'!B3)-(SUMIFS(Investors!$M:$M,Investors!$I:$I,\"<=\"&'NSST Print'!B3,Investors!$J:$J,\">\"&'NSST Print'!B3)+SUMIFS(Investors!$M:$M,Investors!$I:$I,\"<=\"&'NSST Print'!B3,Investors!$J:$J,\"\")))"])
        ws9.append(["Current Investor Capital deployed", "=SUMIFS(Investors!$M:$M,Investors!$O:$O,FALSE,Investors!$P:$P,FALSE,Investors!$K:$K,\">\"&'NSST Print'!$B$3,Investors!$I:$I,\"<=\"&'NSST Print'!$B$3)-(SUMIFS(Investors!$M:$M,Investors!$I:$I,\"<=\"&'NSST Print'!B3,Investors!$J:$J,\">\"&'NSST Print'!B3)+SUMIFS(Investors!$M:$M,Investors!$I:$I,\"<=\"&'NSST Print'!B3,Investors!$J:$J,\"\"))"])
        ws9.append(["INVESTMENTS"])
        ws9.append(["No. of Capital Investments received", "=COUNTIFS(Investors!$G:$G,FALSE)+COUNTIFS(Investors!$G:$G,TRUE)"])
        ws9.append(["No. Investments exited to date", "=COUNTIFS(Investors!$G:$G,TRUE)+COUNTIFS(Investors!$G:$G,FALSE,Investors!$O:$O,TRUE,Investors!$P:$P,FALSE)"])
        ws9.append(["No. Investments still in Development", "=B15-B16"])
        ws9.append(["SALES INCOME"])
        ws9.append(["", "Total", "Transferred", "Sold", "Remaining"])
        ws9.append(["Units", "=COUNTIFS(Sales!$E:$E,TRUE,Sales!$A:$A,\"<>\"&\"Endulini\")+COUNTIFS(Sales!$E:$E,FALSE,Sales!$A:$A,\"<>\"&\"Endulini\")", "=COUNTIFS(Sales!$E:$E,TRUE,Sales!$A:$A,\"<>\"&\"Endulini\")", "=COUNTIFS(Sales!$E:$E,FALSE,Sales!$D:$D,TRUE,Sales!$A:$A,\"<>\"&\"Endulini\")", "=COUNTIFS(Sales!$E:$E,FALSE,Sales!$D:$D,FALSE,Sales!$A:$A,\"<>\"&\"Endulini\")"])
        ws9.append(["Sales Income", "=SUMIFS(Sales!$K:$K,Sales!$A:$A,\"<>\"&\"Endulini\",Sales!$D:$D,TRUE)+SUMIFS(Sales!$K:$K,Sales!$A:$A,\"<>\"&\"Endulini\",Sales!$D:$D,FALSE)+C46", "=SUMIFS(Sales!$K:$K,Sales!$E:$E,TRUE,Sales!$A:$A,\"<>\"&\"Endulini\")", "=SUMIFS(Sales!$K:$K,Sales!$E:$E,FALSE,Sales!$D:$D,TRUE,Sales!$A:$A,\"<>\"&\"Endulini\")", "=SUMIFS(Sales!$K:$K,Sales!$E:$E,FALSE,Sales!$D:$D,FALSE,Sales!$A:$A,\"<>\"&\"Endulini\")+C46"])
        ws9.append(["Commission", "=(SUMIFS(Sales!$O:$O,Sales!$A:$A,\"<>\"&\"Endulini\",Sales!$D:$D,TRUE)+SUMIFS(Sales!$O:$O,Sales!$A:$A,\"<>\"&\"Endulini\",Sales!$D:$D,FALSE))/1.15-C52", "=SUMIFS(Sales!$O:$O,Sales!$E:$E,TRUE,Sales!$A:$A,\"<>\"&\"Endulini\")/1.15", "=SUMIFS(Sales!$O:$O,Sales!$E:$E,FALSE,Sales!$D:$D,TRUE,Sales!$A:$A,\"<>\"&\"Endulini\")/1.15", "=SUMIFS(Sales!$O:$O,Sales!$E:$E,FALSE,Sales!$D:$D,FALSE,Sales!$A:$A,\"<>\"&\"Endulini\")/1.15-C52"])
        ws9.append(["Transfer Fees", "=SUMIFS(Sales!$L:$L,Sales!$A:$A,\"<>\"&\"Endulini\",Sales!$D:$D,TRUE)+SUMIFS(Sales!$L:$L,Sales!$A:$A,\"<>\"&\"Endulini\",Sales!$D:$D,FALSE)", "=SUMIFS(Sales!$L:$L,Sales!$E:$E,TRUE,Sales!$A:$A,\"<>\"&\"Endulini\")", "=SUMIFS(Sales!$L:$L,Sales!$E:$E,FALSE,Sales!$D:$D,TRUE,Sales!$A:$A,\"<>\"&\"Endulini\")", "=SUMIFS(Sales!$L:$L,Sales!$E:$E,FALSE,Sales!$D:$D,FALSE,Sales!$A:$A,\"<>\"&\"Endulini\")"])
        ws9.append(["Bond Registration", "=SUMIFS(Sales!$P:$P,Sales!$A:$A,\"<>\"&\"Endulini\",Sales!$D:$D,TRUE)+SUMIFS(Sales!$P:$P,Sales!$A:$A,\"<>\"&\"Endulini\",Sales!$D:$D,FALSE)", "=SUMIFS(Sales!$P:$P,Sales!$E:$E,TRUE,Sales!$A:$A,\"<>\"&\"Endulini\")", "=SUMIFS(Sales!$P:$P,Sales!$E:$E,FALSE,Sales!$D:$D,TRUE,Sales!$A:$A,\"<>\"&\"Endulini\")", "=SUMIFS(Sales!$P:$P,Sales!$E:$E,FALSE,Sales!$D:$D,FALSE,Sales!$A:$A,\"<>\"&\"Endulini\")"])
        ws9.append(["Security Release Fee", "=SUMIFS(Sales!$M:$M,Sales!$A:$A,\"<>\"&\"Endulini\",Sales!$D:$D,TRUE)+SUMIFS(Sales!$M:$M,Sales!$A:$A,\"<>\"&\"Endulini\",Sales!$D:$D,FALSE)", "=SUMIFS(Sales!$M:$M,Sales!$E:$E,TRUE,Sales!$A:$A,\"<>\"&\"Endulini\")", "=SUMIFS(Sales!$M:$M,Sales!$E:$E,FALSE,Sales!$D:$D,TRUE,Sales!$A:$A,\"<>\"&\"Endulini\")", "=SUMIFS(Sales!$M:$M,Sales!$E:$E,FALSE,Sales!$D:$D,FALSE,Sales!$A:$A,\"<>\"&\"Endulini\")"])
        ws9.append(["Unforseen (0.05%)", "=SUMIFS(Sales!$N:$N,Sales!$A:$A,\"<>\"&\"Endulini\",Sales!$D:$D,TRUE)+SUMIFS(Sales!$N:$N,Sales!$A:$A,\"<>\"&\"Endulini\",Sales!$D:$D,FALSE)-C53", "=SUMIFS(Sales!$N:$N,Sales!$E:$E,TRUE,Sales!$A:$A,\"<>\"&\"Endulini\")", "=SUMIFS(Sales!$N:$N,Sales!$E:$E,FALSE,Sales!$D:$D,TRUE,Sales!$A:$A,\"<>\"&\"Endulini\")", "=SUMIFS(Sales!$N:$N,Sales!$E:$E,FALSE,Sales!$D:$D,FALSE,Sales!$A:$A,\"<>\"&\"Endulini\")-C53"])
        ws9.append(["Transfer Income", "=B21-SUM(B22:B26)", "=C21-SUM(C22:C26)", "=D21-SUM(D22:D26)", "=E21-SUM(E22:E26)"])
        ws9.append(["INTEREST"])
        ws9.append(["Total Estimated Interest", "=SUMIFS(Investors!$S:$S,Investors!$G:$G,FALSE)+SUMIFS(Investors!$S:$S,Investors!$G:$G,TRUE)"])
        ws9.append(["Interest Paid to Date", "=SUMIFS(Investors!$S:$S,Investors!$G:$G,TRUE)+SUMIFS(Investors!$S:$S,Investors!$G:$G,FALSE,Investors!$O:$O,TRUE)"])
        ws9.append(["Remaining Interest to Exit", "=B29-B30"])
        ws9.append(["Interest Due to date", "=SUMIFS(Investors!$T:$T,Investors!$G:$G,FALSE,Investors!$O:$O,FALSE)+SUMIFS(Investors!$Q:$Q,Investors!$G:$G,FALSE,Investors!$O:$O,FALSE)"])
        ws9.append(["Interest Due to be Earned to Exit", "=B31-B32"])
        ws9.append(["FUNDING AVAILABLE"])
        ws9.append(["Total Draw funds available", "=Dashboard!B20"])
        ws9.append(["Projected Heron Projects Income", "=Dashboard!B11+C46"])
        ws9.append(["Total Current Funds Available", "=B35+B36"])
        ws9.append(["Current Construction Cost", "=-Dashboard!B28+C47+C48+C49+C50+C51"])
        ws9.append(["Total funds (required)/Surplus", "=B37-B38"])
        ws9.append(["PROJECTED PROFIT"])
        ws9.append(["Projected Nett Revenue", "=+B21+SUM(B56:B63)"])
        ws9.append(["Other Income (interest received)", "=-SUMIFS(Xero!$F:$F,Xero!$E:$E,\"Interest Received - Momentum\")"])
        ws9.append(["Total Estimated Development and Sales Costs", "=+B78-C47-C48-C49-C50+SUM(D22:E26)"])
        ws9.append(["Interest Expense", "=B29-B30-C51"])
        ws9.append(["Profit", "=B41+B42-B43-B44"])
        ws9.append(["Sales Increase", 0, "=SUMIFS(Sales!$K:$K,Sales!$E:$E,FALSE,Sales!$D:$D,FALSE,Sales!$A:$A,\"<>\"&\"Endulini\")*B46", "", "=C46/$E$20"])
        ws9.append(["CPC Construction", 0, "=SUMIFS('Updated Construction'!$F:$F,'Updated Construction'!$D:$D,FALSE)*B47",  "=SUMIFS('Updated Construction'!$F:$F,'Updated Construction'!$D:$D,FALSE)-C47", "=C47/$E$20"])
        ws9.append(["Rent Salaries and wages", 0, "=400000*B48", "=SUMIFS('Other Costs'!$D:$D,'Other Costs'!$A:$A,'NSST Print'!$A48,'Other Costs'!$C:$C,\">\"&'NSST Print'!$B$3)-C48", "=C48/$E$20"])
        ws9.append(["CPSD", 0, "=SUMIFS('Other Costs'!$D:$D,'Other Costs'!$A:$A,'NSST Print'!$A49,'Other Costs'!$C:$C,\">\"&'NSST Print'!$B$3)*B49", "=SUMIFS('Other Costs'!$D:$D,'Other Costs'!$A:$A,'NSST Print'!$A49,'Other Costs'!$C:$C,\">\"&'NSST Print'!$B$3)-C49", "=C49/$E$20"])
        ws9.append(["OppInvest", 0, "=SUMIFS('Other Costs'!$D:$D,'Other Costs'!$A:$A,'NSST Print'!$A75,'Other Costs'!$C:$C,\">\"&'NSST Print'!$B$3)*B50", "=SUMIFS('Other Costs'!$D:$D,'Other Costs'!$A:$A,'NSST Print'!$A75,'Other Costs'!$C:$C,\">\"&'NSST Print'!$B$3)-C50", "=C50/$E$20"])
        ws9.append(["investor interest", 0, "=(B29-B30)*B51", "=(B29-B30)-C51", "=C51/$E$20"])
        ws9.append(["Commissions", 0, "=SUMIFS(Sales!$O:$O,Sales!$E:$E,FALSE,Sales!$D:$D,FALSE,Sales!$A:$A,\"<>\"&\"Endulini\")/1.15*B52", "=SUMIFS(Sales!$O:$O,Sales!$E:$E,FALSE,Sales!$D:$D,FALSE,Sales!$A:$A,\"<>\"&\"Endulini\")/1.15-C52", "=C52/$E$20"])
        ws9.append(["Unforseen", 0, "=SUMIFS(Sales!$N:$N,Sales!$E:$E,FALSE,Sales!$D:$D,FALSE,Sales!$A:$A,\"<>\"&\"Endulini\")*B53", "=SUMIFS(Sales!$N:$N,Sales!$E:$E,FALSE,Sales!$D:$D,FALSE,Sales!$A:$A,\"<>\"&\"Endulini\")-C53", "=C53/$E$20"])
        ws9.append([])
        ws9.append([])
        ws9.append(["Rental Income", "=-SUMIFS(Xero!$F:$F,Xero!$E:$E,'NSST Print'!$A56)"])
        ws9.append(["Sales - Heron Fields occupational rent", "=-SUMIFS(Xero!$F:$F,Xero!$E:$E,'NSST Print'!$A57)"])
        ws9.append(["Sales - Heron View Occupational Rent", "=-SUMIFS(Xero!$F:$F,Xero!$E:$E,'NSST Print'!$A58)"])
        ws9.append(["Bond Origination", "=-SUMIFS(Xero!$F:$F,Xero!$E:$E,'NSST Print'!$A59)"])
        ws9.append(["Interest Received - Deposit Grey Heron", "=-SUMIFS(Xero!$F:$F,Xero!$E:$E,'NSST Print'!$A60)"])
        ws9.append(["Interest Received - Deposits", "=-SUMIFS(Xero!$F:$F,Xero!$E:$E,'NSST Print'!$A61)"])
        ws9.append(["Interest Received - STBB Trust Account", "=-SUMIFS(Xero!$F:$F,Xero!$E:$E,'NSST Print'!$A62)"])
        ws9.append(["Revaluation of Land (Not Incl)", "=-SUMIFS(Xero!$F:$F,Xero!$E:$E,'NSST Print'!$A63)"])
        ws9.append([])
        ws9.append([])
        ws9.append(["Costs", "=SUMIFS(Xero!$F:$F,Xero!$C:$C,\"Expenses\",Xero!$A:$A,\"<>\"&Heron!$A$6)"])
        ws9.append(["Heron View Land", "=SUMIFS(Xero!$F:$F,Xero!$E:$E,'NSST Print'!$A67)"])
        ws9.append(["COS - Heron Fields - Construction", "=SUMIFS(Xero!$F:$F,Xero!$E:$E,'NSST Print'!$A68)"])
        ws9.append(["COS - Heron Fields - P & G", "=SUMIFS(Xero!$F:$F,Xero!$E:$E,'NSST Print'!$A69)"])
        ws9.append(["COS - Heron Fields - Printing & Stationary", "=SUMIFS(Xero!$F:$F,Xero!$E:$E,'NSST Print'!$A70)"])
        ws9.append(["COS - Heron View - Construction", "=SUMIFS(Xero!$F:$F,Xero!$E:$E,'NSST Print'!$A71)+SUMIFS('Updated Construction'!$F:$F,'Updated Construction'!$D:$D,FALSE)"])
        ws9.append(["COS - Heron View - P&G", "=SUMIFS(Xero!$F:$F,Xero!$E:$E,'NSST Print'!$A72)"])
        ws9.append(["COS - Heron View - Printing & Stationary", "=SUMIFS(Xero!$F:$F,Xero!$E:$E,'NSST Print'!$A73)"])
        ws9.append(["CPSD", "=SUMIFS('Other Costs'!$D:$D,'Other Costs'!$A:$A,'NSST Print'!$A74)"])
        ws9.append(["OppInvest", "=SUMIFS('Other Costs'!$D:$D,'Other Costs'!$A:$A,'NSST Print'!$A75)"])
        ws9.append(["Professional Fees", "=SUMIFS('Other Costs'!$D:$D,'Other Costs'!$A:$A,'NSST Print'!$A76)"])
        ws9.append(["Rent Salaries and Wages", "=SUMIFS('Other Costs'!$D:$D,'Other Costs'!$A:$A,'NSST Print'!$A77)"])
        ws9.append(["Total Costs", "=SUM(B66:B77)"])


        # make the first column 55 and the others 20
        for i in range(1, ws9.max_column + 1):
            if i == 1:
                ws9.column_dimensions[get_column_letter(i)].width = 50
            else:
                ws9.column_dimensions[get_column_letter(i)].width = 20

        # format all cells as font size 14
        for i in range(1, ws9.max_column + 1):
            for j in range(1, ws9.max_row + 1):
                ws9[f"{get_column_letter(i)}{j}"].font = Font(size=14)
                # After row format all cells as currency except the first column and except rows 15, 16 and 17
                if j < 15 and j > 5 and i > 1:
                    ws9[f"{get_column_letter(i)}{j}"].number_format = 'R#,##0.00'

                if j > 17 and j < 20 and i > 1:
                    ws9[f"{get_column_letter(i)}{j}"].number_format = 'R#,##0.00'

                if j > 20 and i > 1:
                    ws9[f"{get_column_letter(i)}{j}"].number_format = 'R#,##0.00'
                # After row 2, put a border around all cells until row 46
                if j > 1 and j < 46:
                    ws9[f"{get_column_letter(i)}{j}"].border = Border(top=Side(style='thin'), bottom=Side(style='thin'),
                                                                      left=Side(style='thin'), right=Side(style='thin'))

        rows_for_full_merge = [2, 5, 14, 18, 28, 34, 40]
        for i in rows_for_full_merge:
            ws9.merge_cells(f"A{i}:E{i}")
            # make the cell bold and color it light olive and make the text white
            ws9[f"A{i}"].font = Font(bold=True, size=20, color="FFFFFF")
            ws9[f"A{i}"].fill = PatternFill(start_color="7F9F80", end_color="7F9F80", fill_type="solid")
            ws9[f"A{i}"].alignment = Alignment(horizontal='center', vertical='center')

        rows_for_part_merger = [3, 4, 6, 7, 8, 9, 10, 11, 12, 13, 15, 16, 17, 29, 30, 31, 32, 33, 35, 36, 37, 38, 39,
                                41, 42, 43, 44, 45]
        for i in rows_for_part_merger:
            ws9.merge_cells(f"B{i}:E{i}")
            ws9[f"B{i}"].alignment = Alignment(horizontal='center', vertical='center')

        for i in range(46,54):
            if i == 48:
                # format the cell in row 48 as an integer
                ws9[f"B{i}"].number_format = '0'
            else:
                #format as a percentage
                ws9[f"B{i}"].number_format = '0%'

        # ws9['F43'].value = "<< SALES COSTS DO NOT APPEAR TO INCLUDED PREVIOUSLY"
        # Hide row 10, 31,32,33
        ws9.row_dimensions[10].hidden = True
        ws9.row_dimensions[31].hidden = True
        ws9.row_dimensions[32].hidden = True
        ws9.row_dimensions[33].hidden = True

        # in row 19 and 20, center the text from column B to E
        for i in range(2, 6):
            ws9[f"{get_column_letter(i)}19"].alignment = Alignment(horizontal='center', vertical='center')
            ws9[f"{get_column_letter(i)}20"].alignment = Alignment(horizontal='center', vertical='center')


       # get a list of sheets in the workbook
        sheet_names = wb.sheetnames
        # print(sheet_names)
        sheet_names_to_hide = ['Updated Construction','Operational Costs', 'Xero', 'Other Costs','Investors', 'Sales','Construction', 'Opportunities']
        # loop through the sheets and hide the ones in the list

        for sheet in sheet_names_to_hide:
            ws = wb[sheet]
            ws.sheet_state = 'hidden'

        # move sheet 'NSST Print' to the beginning of the workbook



        # sheets_to_move = ['Investors', 'Construction', 'Sales']
        # # move the sheets to the nd of the workbook
        # for sheet in sheets_to_move:
        #     ws = wb[sheet]
        #     sheet_names.move_sheet(ws, -1)


        # remove the default sheet
        wb.remove(wb['Sheet'])
        wb.save('cashflow_p&l_files/cashflow_projection.xlsx')
        return "cashflow_p&l_files/cashflow_projection.xlsx"
    except Exception as e:
        print("Error XX", e)
        return "Error"


other_costs = [
    {
        "AccountName": "CPSD",
        "Category": "Expenses",
        "ReportDate": "2023-10-31",
        "Amount": 5125504.85
    },
    {
        "AccountName": "CPSD",
        "Category": "Expenses",
        "ReportDate": "2023-11-30",
        "Amount": 314037.88
    },
    {
        "AccountName": "CPSD",
        "Category": "Expenses",
        "ReportDate": "2023-12-31",
        "Amount": 314037.88
    },
    {
        "AccountName": "CPSD",
        "Category": "Expenses",
        "ReportDate": "2024-01-31",
        "Amount": 314037.88
    },
    {
        "AccountName": "CPSD",
        "Category": "Expenses",
        "ReportDate": "2024-02-29",
        "Amount": 314037.88
    },
    {
        "AccountName": "CPSD",
        "Category": "Expenses",
        "ReportDate": "2024-03-31",
        "Amount": 314037.88
    },
    {
        "AccountName": "CPSD",
        "Category": "Expenses",
        "ReportDate": "2024-04-30",
        "Amount": 314037.88
    },
    {
        "AccountName": "CPSD",
        "Category": "Expenses",
        "ReportDate": "2024-05-31",
        "Amount": 314037.88
    },
    {
        "AccountName": "CPSD",
        "Category": "Expenses",
        "ReportDate": "2024-06-30",
        "Amount": 314037.88
    },
    {
        "AccountName": "CPSD",
        "Category": "Expenses",
        "ReportDate": "2024-07-31",
        "Amount": 314037.88
    },
    {
        "AccountName": "CPSD",
        "Category": "Expenses",
        "ReportDate": "2024-08-31",
        "Amount": 314037.88
    },
    # {
    #     "AccountName": "CPSD",
    #     "Category": "Expenses",
    #     "ReportDate": "2024-09-30",
    #     "Amount": 314037.88
    # },
    # {
    #     "AccountName": "CPSD",
    #     "Category": "Expenses",
    #     "ReportDate": "2024-10-31",
    #     "Amount": 314037.88
    # },
    # {
    #     "AccountName": "CPSD",
    #     "Category": "Expenses",
    #     "ReportDate": "2024-11-30",
    #     "Amount": 314037.88
    # },
    # {
    #     "AccountName": "CPSD",
    #     "Category": "Expenses",
    #     "ReportDate": "2024-12-31",
    #     "Amount": 314037.88
    # },
    # {
    #     "AccountName": "CPSD",
    #     "Category": "Expenses",
    #     "ReportDate": "2025-01-31",
    #     "Amount": 314037.88
    # },
    # {
    #     "AccountName": "CPSD",
    #     "Category": "Expenses",
    #     "ReportDate": "2025-02-28",
    #     "Amount": 314037.88
    # },
    {
        "AccountName": "OppInvest",
        "Category": "Expenses",
        "ReportDate": "2023-10-31",
        "Amount": 5676905
    },
    {
        "AccountName": "OppInvest",
        "Category": "Expenses",
        "ReportDate": "2023-11-30",
        "Amount": 392914.302
    },
    {
        "AccountName": "OppInvest",
        "Category": "Expenses",
        "ReportDate": "2023-12-31",
        "Amount": 392914.302
    },
    {
        "AccountName": "OppInvest",
        "Category": "Expenses",
        "ReportDate": "2024-01-31",
        "Amount": 392914.302
    },
    {
        "AccountName": "OppInvest",
        "Category": "Expenses",
        "ReportDate": "2024-02-29",
        "Amount": 392914.302
    },
    {
        "AccountName": "OppInvest",
        "Category": "Expenses",
        "ReportDate": "2024-03-31",
        "Amount": 392914.302
    },
    {
        "AccountName": "OppInvest",
        "Category": "Expenses",
        "ReportDate": "2024-04-30",
        "Amount": 392914.302
    },
    {
        "AccountName": "OppInvest",
        "Category": "Expenses",
        "ReportDate": "2024-05-31",
        "Amount": 392914.302
    },
    {
        "AccountName": "OppInvest",
        "Category": "Expenses",
        "ReportDate": "2024-06-30",
        "Amount": 392914.302
    },
    {
        "AccountName": "OppInvest",
        "Category": "Expenses",
        "ReportDate": "2024-07-31",
        "Amount": 392914.302
    },
    {
        "AccountName": "OppInvest",
        "Category": "Expenses",
        "ReportDate": "2024-08-31",
        "Amount": 392914.302
    },
    # {
    #     "AccountName": "OppInvest",
    #     "Category": "Expenses",
    #     "ReportDate": "2024-09-30",
    #     "Amount": 392914.302
    # },
    # {
    #     "AccountName": "OppInvest",
    #     "Category": "Expenses",
    #     "ReportDate": "2024-10-31",
    #     "Amount": 392914.302
    # },
    # {
    #     "AccountName": "OppInvest",
    #     "Category": "Expenses",
    #     "ReportDate": "2024-11-30",
    #     "Amount": 392914.302
    # },
    # {
    #     "AccountName": "OppInvest",
    #     "Category": "Expenses",
    #     "ReportDate": "2024-12-31",
    #     "Amount": 392914.302
    # },
    # {
    #     "AccountName": "OppInvest",
    #     "Category": "Expenses",
    #     "ReportDate": "2025-01-31",
    #     "Amount": 392914.302
    # },
    # {
    #     "AccountName": "OppInvest",
    #     "Category": "Expenses",
    #     "ReportDate": "2025-02-28",
    #     "Amount": 392914.302
    # },
    {
        "AccountName": "Rent Salaries and Wages",
        "Category": "Expenses",
        "ReportDate": "2023-10-31",
        "Amount": 17272544
    },
    {
        "AccountName": "Rent Salaries and Wages",
        "Category": "Expenses",
        "ReportDate": "2023-11-30",
        "Amount": 800000
    },
    {
        "AccountName": "Rent Salaries and Wages",
        "Category": "Expenses",
        "ReportDate": "2023-12-31",
        "Amount": 800000
    },
    {
        "AccountName": "Rent Salaries and Wages",
        "Category": "Expenses",
        "ReportDate": "2024-01-31",
        "Amount": 800000
    },
    {
        "AccountName": "Rent Salaries and Wages",
        "Category": "Expenses",
        "ReportDate": "2024-02-29",
        "Amount": 800000
    },
    {
        "AccountName": "Rent Salaries and Wages",
        "Category": "Expenses",
        "ReportDate": "2024-03-31",
        "Amount": 800000
    },
    {
        "AccountName": "Rent Salaries and Wages",
        "Category": "Expenses",
        "ReportDate": "2024-04-30",
        "Amount": 800000
    },
    {
        "AccountName": "Rent Salaries and Wages",
        "Category": "Expenses",
        "ReportDate": "2024-05-31",
        "Amount": 800000
    },
    {
        "AccountName": "Rent Salaries and Wages",
        "Category": "Expenses",
        "ReportDate": "2024-06-30",
        "Amount": 800000
    },
    {
        "AccountName": "Rent Salaries and Wages",
        "Category": "Expenses",
        "ReportDate": "2024-07-31",
        "Amount": 800000
    },
    {
        "AccountName": "Rent Salaries and Wages",
        "Category": "Expenses",
        "ReportDate": "2024-08-31",
        "Amount": 800000
    },
    # {
    #     "AccountName": "Rent Salaries and Wages",
    #     "Category": "Expenses",
    #     "ReportDate": "2024-09-30",
    #     "Amount": 800000
    # },
    # {
    #     "AccountName": "Rent Salaries and Wages",
    #     "Category": "Expenses",
    #     "ReportDate": "2024-10-31",
    #     "Amount": 800000
    # },
    # {
    #     "AccountName": "Rent Salaries and Wages",
    #     "Category": "Expenses",
    #     "ReportDate": "2024-11-30",
    #     "Amount": 800000
    # },
    # {
    #     "AccountName": "Rent Salaries and Wages",
    #     "Category": "Expenses",
    #     "ReportDate": "2024-12-31",
    #     "Amount": 800000
    # },
    # {
    #     "AccountName": "Rent Salaries and Wages",
    #     "Category": "Expenses",
    #     "ReportDate": "2025-01-31",
    #     "Amount": 800000
    # },
    # {
    #     "AccountName": "Rent Salaries and Wages",
    #     "Category": "Expenses",
    #     "ReportDate": "2025-02-28",
    #     "Amount": 800000
    # },
    {
        "AccountName": "Professional Fees",
        "Category": "Expenses",
        "ReportDate": "2023-10-31",
        "Amount": 5361342.92
    }
]
