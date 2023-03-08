import copy
from datetime import datetime, timedelta

from dateutil.relativedelta import relativedelta
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill, Color
from openpyxl.styles.borders import Border, Side
from openpyxl.styles import borders
from openpyxl.styles.alignment import Alignment
from openpyxl.formula.translate import Translator
from openpyxl.utils.cell import coordinate_from_string

from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment


# CREATE HELPER FUNCTIONS
def create_excel_array(data):
    worksheet_data = []
    for item in data:
        if item['release_date'] != "" and item['opportunity_code'] != "ZZUN01":
            item['funds_drawn'] = item['investment_amount']
            item['funds_in_momentum'] = 0
            item['funds_to_be_raised'] = 0
        elif item['release_date'] == "" and item['opportunity_code'] != "ZZUN01":
            item['funds_drawn'] = 0
            item['funds_in_momentum'] = item['investment_amount']
            item['funds_to_be_raised'] = 0
        elif item['opportunity_code'] == "ZZUN01":
            item['funds_to_be_raised'] = item['opportunity_amount_required']
            item['funds_drawn'] = 0
            item['funds_in_momentum'] = 0

    interest_on_funds_drawn, interest_on_funds_in_momentum = calculate_interest_figures(data)

    # row1_data = [f"{sheet_name} - {developmentinputdata['date']}"]
    row2_data = ["Opportunity", "Total", "Interest", "Transferred", "sold", "Remaining"]
    row4_data = ["Sold", "", "", "", "", ""]
    row5_data = ["Transferred", "", "", "", "", ""]
    row7_data = ["Investor", "", "", "", "", ""]
    row8_data = ["Account", "", "", "", "", ""]
    row9_data = ["Project Interest", "", "", "", "", ""]
    row13_data = ["Capital Required", "", "", "", "", ""]
    row14_data = ["Capital Raised", "", "", "", "", ""]
    row15_data = ["Capital Drawn", "", interest_on_funds_drawn, "", "", ""]
    row16_data = ["Capital in Momentum", "", interest_on_funds_in_momentum, "", "", ""]
    row17_data = ["Capital to be Raised", "", "", "", "", ""]
    row19_data = ["Capital Invested", "", "", "", "", ""]
    row20_data = ["Momentum Deposit Date", "", "", "", "", ""]
    row21_data = ["Released Date", "", "", "", "", ""]
    row22_data = ["End Date", "", "", "", "", ""]
    row23_data = ["Days in Transaction", "", "", "", "", ""]
    row24_data = ["Investment Interest to Date", "", "", "", "", ""]
    row25_data = ["Released Interest to Date", "", "", "", "", ""]
    row26_data = ["Total Interest to Date", "", "", "", "", ""]
    row27_data = ["Contract End Date", "", "", "", "", ""]
    row28_data = ["Days to exit", "", "", "", "", ""]
    row29_data = ["Investment Account Interest Earned.", "", "", "", "", ""]
    row30_data = ["Released Interest Earned.", "", "", "", "", ""]
    row31_data = ["Total Interest Earned.", "", "", "", "", ""]
    row31A_data = ["Due to Investors.", "", "", "", "", ""]
    row33_data = ["Raising Commission (4%)", "", "", "", "", ""]
    row34_data = ["Structuring Fees (3%)", "", "", "", "", ""]
    row35_data = ["Total Fees", "", "", "", "", ""]
    row36_data = ["Available for construction after fees", "", "", "", "", ""]
    row38_data = ["Sales Price", "", "", "", "", ""]
    row39_data = ["VAT", "", "", "", "", ""]
    row40_data = ["Nett Sales Price", "", "", "", "", ""]
    row41_data = ["Commission", "", "", "", "", ""]
    row42_data = ["Transfer Fees", "", "", "", "", ""]
    row43_data = ["Bond Registration", "", "", "", "", ""]
    row44_data = ["Trust Release Fee", "", "", "", "", ""]
    row45_data = ["Unforeseen", "", "", "", "", ""]
    row46_data = ["Discount", "", "", "", "", ""]
    row47_data = ["Transfer Income", "", "", "", "", ""]
    row48_data = ["Due to Investors", "", "", "", "", ""]
    row49_data = ["Profit / Loss", "", "", "", "", ""]

    for item in data:
        row2_data.append(item['opportunity_code'])
        row4_data.append(item['opportunity_sold'])
        row5_data.append(item['opportunity_transferred'])

        # append investor name in the following manner, first character of name, full surname, if no name
        # then just surname
        if item['investor_name'] != "":
            row7_data.append(f"{item['investor_name'][0]}. {item['investor_surname']}")
        else:
            row7_data.append(f"{item['investor_surname']}")
        row8_data.append(item['investor_acc_number'])
        row9_data.append(float(item['investment_interest_rate']) / 100)
        row13_data.append(item['opportunity_amount_required'])

        # Create a new list from data based on the item['opportunity_code'] and sum the investment_amount in
        # the list then append the sum to the row14_data list
        row14_data.append(sum([float(x['investment_amount']) for x in data
                               if x['opportunity_code'] == item['opportunity_code']]))
        row15_data.append(sum([float(x['funds_drawn']) for x in data
                               if x['opportunity_code'] == item['opportunity_code']]))
        row16_data.append(sum([float(x['funds_in_momentum']) for x in data
                               if x['opportunity_code'] == item['opportunity_code']]))
        row19_data.append(item['investment_amount'])
        row20_data.append(item['deposit_date'])
        row21_data.append(item['planned_release_date'])
        # row22_data.append(item['opportunity_final_transfer_date'])
        if item['opportunity_code'] != "ZZUN01":
            row22_data.append(item['opportunity_final_transfer_date'])
        else:
            row22_data.append("")

        row24_data.append(item['investment_interest_today'])
        row25_data.append(item['released_interest_today'])
        # for row27_data if the item['opportunity_code'] is not equal to 'ZZUN01', take the item['deposit_date'], replace '-' with '/' then convert it to datetime and add exactly two years to it and append the new date to the list formatted as YYYY/MM/DD else just make it ""
        if item['opportunity_code'] != "ZZUN01" and item['deposit_date'] != "":
            row27_data.append(
                (datetime.strptime(item['deposit_date'], '%Y/%m/%d') + timedelta(days=730)).strftime('%Y/%m/%d'))
        else:
            row27_data.append("")
        row29_data.append(item['trust_interest_total'])
        row30_data.append(item['released_interest_total'])
        row33_data.append(float(item['raising_commission']) * float(item['investment_amount']))
        row34_data.append(float(item['structuring_fee']) * float(item['investment_amount']))
        row38_data.append(item['opportunity_sale_price'])
        row39_data.append(float(item['opportunity_sale_price']) / 1.15 * 0.15)
        row40_data.append(float(item['opportunity_sale_price']) / 1.15)
        row41_data.append(float(item['opportunity_sale_price']) / 1.15 * float(item['commission']))
        row42_data.append(float(item['transfer_fees']))
        row43_data.append(float(item['bond_registration']))
        row44_data.append(float(item['trust_release_fee']))
        row45_data.append(float(item['unforseen']) * float(item['opportunity_sale_price']))
        row46_data.append(0)

    # worksheet_data.append(row1_data)
    worksheet_data.append(row4_data)
    worksheet_data.append(row5_data)
    worksheet_data.append(row2_data)
    worksheet_data.append(row2_data)
    worksheet_data.append(row4_data)
    worksheet_data.append(row5_data)
    worksheet_data.append([])
    worksheet_data.append(row7_data)
    worksheet_data.append(row8_data)
    worksheet_data.append(row9_data)
    worksheet_data.append([])
    worksheet_data.append(row13_data)
    worksheet_data.append(row14_data)
    worksheet_data.append(row15_data)
    worksheet_data.append(row16_data)
    worksheet_data.append(row17_data)
    worksheet_data.append([])
    worksheet_data.append(row19_data)
    worksheet_data.append(row20_data)
    worksheet_data.append(row21_data)
    worksheet_data.append(row22_data)
    worksheet_data.append(row23_data)
    worksheet_data.append(row24_data)
    worksheet_data.append(row25_data)
    worksheet_data.append(row26_data)
    worksheet_data.append(row27_data)
    worksheet_data.append(row28_data)
    worksheet_data.append(row29_data)
    worksheet_data.append(row30_data)
    worksheet_data.append(row31_data)
    worksheet_data.append(row31A_data)
    worksheet_data.append([])
    worksheet_data.append(row33_data)
    worksheet_data.append(row34_data)
    worksheet_data.append(row35_data)
    worksheet_data.append(row36_data)
    worksheet_data.append([])
    worksheet_data.append(row38_data)
    worksheet_data.append(row39_data)
    worksheet_data.append(row40_data)
    worksheet_data.append(row41_data)
    worksheet_data.append(row42_data)
    worksheet_data.append(row43_data)
    worksheet_data.append(row44_data)
    worksheet_data.append(row45_data)
    worksheet_data.append(row46_data)
    worksheet_data.append(row47_data)
    worksheet_data.append(row48_data)
    worksheet_data.append(row49_data)

    merge_start = []
    merge_end = []

    # from column 6 in excel to the end of the row, add the row number to the merge_start list if the value
    # for each cell in row 3 is different to the cell to its immediete left
    for index, item in enumerate(row2_data):
        if index > 5:
            if item != row2_data[index - 1]:
                merge_start.append(index)

    for index, item in enumerate(row2_data):
        if 5 < index < len(row2_data) - 1:
            if item != row2_data[index + 1]:
                merge_end.append(index)
        elif index == len(row2_data) - 1:
            merge_end.append(index)

    # loop through the merge_start list and add 1 to each item using list comprehension
    merge_start = [x + 1 for x in merge_start]
    merge_end = [x + 1 for x in merge_end]

    return worksheet_data, interest_on_funds_drawn, interest_on_funds_in_momentum


def calculate_interest_figures(data):
    interest_on_funds_drawn = float(0)
    interest_on_funds_in_momentum = float(0)

    for item in data:
        # check if item['funds_drawn'] exists and if it does and if its > 0, then add the value of item[
        # 'released_interest_today'] to interest_on_funds_drawn
        if 'funds_drawn' in item and item['funds_drawn'] > 0:
            interest_on_funds_drawn += float(item['released_interest_today'])
        # Do the same as above, although instead of funds_drawn, check if funds_in_momentum exists and if it does and
        # if its > 0, then add the value of item['investment_interest_today'] to interest_on_funds_in_momentum
        if 'funds_in_momentum' in item and float(item['funds_in_momentum']) > 0:
            interest_on_funds_in_momentum += float(item['investment_interest_today'])

    return interest_on_funds_drawn, interest_on_funds_in_momentum


def create_nsst_sheet(category, developmentinputdata, pledges, index, sheet_name, worksheets):
    # deep copy category list so that changes to the copy do not affect the original
    category_new = copy.deepcopy(category)
    # if index == 0 and category length equals 2 then insert f"{category[0]} and {category[1]}" into the begining of category list
    if index == 0 and len(category_new) == 2:
        category_new.insert(0, f"{category_new[0]} and {category_new[1]}")
    if index == 0:
        category_index = 0
    elif index == 1:
        category_index = 0
    elif index == 2:
        category_index = 1
    nsst_data = []
    row1_data = [f"{sheet_name} Investor Report - {developmentinputdata['date']}"]
    nsst_data.append(row1_data)
    nsst_data.append([])
    nsst_data.append([])
    nsst_data.append(["Report Date", developmentinputdata['date']])
    nsst_data.append(["Development", category_new[category_index]])
    nsst_data.append([])
    nsst_data.append(["CAPITAL"])
    nsst_data.append(["Total Investment capital to be raised (Estimated)", f'=+\'{worksheets[index]}\'!B13'])
    nsst_data.append(["Available to be raised (Estimated)", f'=+\'{worksheets[index]}\'!B17'])
    nsst_data.append(["Total Investment capital received", f'=+\'{worksheets[index]}\'!B14'])
    nsst_data.append(["Total Funds Drawn Down into Development", f'=+\'{worksheets[index]}\'!B15'])
    nsst_data.append(["Pledges Due", ""])
    nsst_data.append(["Momentum Investment Account", f'=+\'{worksheets[index]}\'!B16'])
    nsst_data.append([])
    nsst_data.append(["INVESTMENTS"])
    nsst_data.append(["No. of Capital Investments received", ""])
    nsst_data.append(["No. Investments exited to date", ""])
    nsst_data.append(["No. Investments still in Development", ""])
    nsst_data.append([])
    nsst_data.append(["REPAYMENT"])
    nsst_data.append(["Investments released prior to transfer", ""])
    nsst_data.append(["Investment funds paid out to date (Incl Interest)", ""])
    nsst_data.append(["Capital plus Interest due (Anticipated Exit Date)", ""])
    nsst_data.append(["Investor repayment on current sales (Anticipated Exit Date)", ""])
    nsst_data.append(["Investor Loan Balance after sales Transfers (Anticipated Exit Date)", ""])
    nsst_data.append(["Gross Remaining Sales Income", ""])
    nsst_data.append(["Surplus / Deficit", ""])
    nsst_data.append(["Cost to Complete & Creditors", ""])
    nsst_data.append(["Interest Budget", ""])
    nsst_data.append(["Gross Income", ""])
    nsst_data.append([])
    nsst_data.append(["GROSS INCOME"])
    nsst_data.append(["", "Total", "Transferred", "Sold", "Remaining"])
    nsst_data.append(
        ["Units", f'=+\'{worksheets[index]}\'!B6', f'=+\'{worksheets[index]}\'!D6', f'=+\'{worksheets[index]}\'!E6',
         f'=+\'{worksheets[index]}\'!F6'])
    nsst_data.append(
        ["Sales Price", f'=+\'{worksheets[index]}\'!B39', f'=+\'{worksheets[index]}\'!D39',
         f'=+\'{worksheets[index]}\'!E39',
         f'=+\'{worksheets[index]}\'!F39'])
    nsst_data.append(
        ["VAT", f'=+\'{worksheets[index]}\'!B40', f'=+\'{worksheets[index]}\'!D40', f'=+\'{worksheets[index]}\'!E40',
         f'=+\'{worksheets[index]}\'!F40'])
    nsst_data.append(
        ["Gross", f'=+\'{worksheets[index]}\'!B41', f'=+\'{worksheets[index]}\'!D41', f'=+\'{worksheets[index]}\'!E41',
         f'=+\'{worksheets[index]}\'!F41'])
    nsst_data.append(["Commission (5 %)", f'=+\'{worksheets[index]}\'!B42', f'=+\'{worksheets[index]}\'!D42',
                      f'=+\'{worksheets[index]}\'!E42', f'=+\'{worksheets[index]}\'!F42'])
    nsst_data.append(
        ["Transfer Fees", f'=+\'{worksheets[index]}\'!B43', f'=+\'{worksheets[index]}\'!D43',
         f'=+\'{worksheets[index]}\'!E43',
         f'=+\'{worksheets[index]}\'!F43'])
    nsst_data.append(["Bond Registration", f'=+\'{worksheets[index]}\'!B44', f'=+\'{worksheets[index]}\'!D44',
                      f'=+\'{worksheets[index]}\'!E44', f'=+\'{worksheets[index]}\'!F44'])
    nsst_data.append(["Security Release Fee", f'=+\'{worksheets[index]}\'!B45', f'=+\'{worksheets[index]}\'!D45',
                      f'=+\'{worksheets[index]}\'!E45', f'=+\'{worksheets[index]}\'!F45'])
    nsst_data.append(["Unforseen (0.05%)", f'=+\'{worksheets[index]}\'!B46', f'=+\'{worksheets[index]}\'!D46',
                      f'=+\'{worksheets[index]}\'!E46', f'=+\'{worksheets[index]}\'!F46'])
    nsst_data.append(["Transfer Income", ""])
    nsst_data.append(["Interest Due to Investors  (Expected Exit Dates)", f'=+\'{worksheets[index]}\'!B31',
                      f'=+\'{worksheets[index]}\'!D31', f'=+\'{worksheets[index]}\'!E31',
                      f'=+\'{worksheets[index]}\'!F31'])
    nsst_data.append(["Capital Due to Investors  (Expected Exit Dates)", f'=+\'{worksheets[index]}\'!B14',
                      f'=+\'{worksheets[index]}\'!D14', f'=+\'{worksheets[index]}\'!E14',
                      f'=+\'{worksheets[index]}\'!F14'])
    nsst_data.append(["Total Due to Investors (Expected Exit Dates)", ""])
    nsst_data.append(["Surplus / Deficit", ""])
    nsst_data.append(["Cost to Complete/Creditors", ""])
    nsst_data.append(["Interest Budget", ""])
    nsst_data.append(["Gross Income", ""])

    return nsst_data


def create_sales_forecast_file(data, developmentinputdata, pledges):
    print(len(pledges))
    filename = 'Sales Forecast'
    category = developmentinputdata['Category']
    worksheet_names = []
    if len(category) > 1:
        sheet_name = f"SF {category[0]} & {category[1]}"

    else:
        sheet_name = f"SF {category[0]}"
        worksheet_names.append(sheet_name)

    print(category)
    print(developmentinputdata)

    wb = Workbook()
    worksheet_data, interest_on_funds_drawn, interest_on_funds_in_momentum = create_excel_array(data)
    # print(interest_on_funds_drawn, interest_on_funds_in_momentum)

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
            for item in worksheet_data:
                ws.append(item)

    # LOOP THROUGH EACH SHEET AND FORMAT ACCORDINGLY
    for sheet in wb.worksheets:
        # print(sheet.title)

        # make all column_widths in ws 20 wide
        for col in sheet.columns:
            col_width = 15
            col_letter = col[0].column_letter
            sheet.column_dimensions[col_letter].width = col_width

        # for rows 6 and 7, if the value in the cell from column 6 is FALSE, set the value to 'No' else set the value
        # to 'Yes'
        for index, row in enumerate(sheet.iter_rows(min_row=6, min_col=7, max_row=7, max_col=sheet.max_column)):
            for cell in row:
                if not cell.value:
                    sheet[f'{get_column_letter(cell.column)}{cell.row}'] = 'No'
                else:
                    sheet[f'{get_column_letter(cell.column)}{cell.row}'] = 'Yes'

        # In cell B6, count the values of all records from F6 to the last column in row 6
        sheet['B6'] = f'=COUNTIF(G7:{get_column_letter(sheet.max_column)}7,"*")'
        sheet['D6'] = f'=COUNTIF(G7:{get_column_letter(sheet.max_column)}7,"Yes")'
        sheet[
            'E6'] = f'=COUNTIF(G6:{get_column_letter(sheet.max_column)}6,"Yes")-' \
                    f'COUNTIF(G7:{get_column_letter(sheet.max_column)}7,"Yes") '
        # In Cell F6, Add the value of B6 less D6 less E6
        sheet['F6'] = f'=B6-D6-E6'

        # From column 6 in row 17, subtract the value in the cell in row 14 from the value in the cell in row 13
        for index, row in enumerate(sheet.iter_rows(min_row=17, min_col=7, max_row=17, max_col=sheet.max_column)):
            for cell in row:
                sheet[
                    f'{get_column_letter(cell.column)}17'] = f'=SUM({get_column_letter(cell.column)}13-' \
                                                             f'{get_column_letter(cell.column)}14)'
        for index, row in enumerate(
                sheet.iter_rows(min_row=23, min_col=7, max_row=23, max_col=sheet.max_column)):
            for cell in row:
                # DAYS IN TRANSACTION
                sheet[
                    f'{get_column_letter(cell.column)}23'] = f'=SUM({get_column_letter(cell.column)}22-' \
                                                             f'{get_column_letter(cell.column)}21)'
        for index, row in enumerate(
                sheet.iter_rows(min_row=26, min_col=7, max_row=26, max_col=sheet.max_column)):
            for cell in row:
                sheet[
                    f'{get_column_letter(cell.column)}26'] = f'=SUM({get_column_letter(cell.column)}24+' \
                                                             f'{get_column_letter(cell.column)}25) '
        for index, row in enumerate(
                sheet.iter_rows(min_row=28, min_col=7, max_row=28, max_col=sheet.max_column)):
            for cell in row:
                sheet[
                    f'{get_column_letter(cell.column)}28'] = f'=720-SUM({get_column_letter(cell.column)}23)'
        for index, row in enumerate(
                sheet.iter_rows(min_row=31, min_col=7, max_row=31, max_col=sheet.max_column)):
            for cell in row:
                sheet[
                    f'{get_column_letter(cell.column)}31'] = f'=SUM({get_column_letter(cell.column)}29+' \
                                                             f'{get_column_letter(cell.column)}30)'
        for index, row in enumerate(
                sheet.iter_rows(min_row=36, min_col=7, max_row=36, max_col=sheet.max_column)):
            for cell in row:
                sheet[
                    f'{get_column_letter(cell.column)}36'] = f'=SUM({get_column_letter(cell.column)}34+' \
                                                             f'{get_column_letter(cell.column)}35)'
        for index, row in enumerate(
                sheet.iter_rows(min_row=37, min_col=7, max_row=37, max_col=sheet.max_column)):
            for cell in row:
                sheet[
                    f'{get_column_letter(cell.column)}37'] = f'=SUM({get_column_letter(cell.column)}19-' \
                                                             f'{get_column_letter(cell.column)}36)'
        for index, row in enumerate(
                sheet.iter_rows(min_row=32, min_col=7, max_row=32, max_col=sheet.max_column)):
            for cell in row:
                sheet[
                    f'{get_column_letter(cell.column)}32'] = f'=SUM({get_column_letter(cell.column)}19+' \
                                                             f'{get_column_letter(cell.column)}31)'
        for index, row in enumerate(
                sheet.iter_rows(min_row=48, min_col=7, max_row=48, max_col=sheet.max_column)):
            for cell in row:
                sheet[
                    f'{get_column_letter(cell.column)}48'] = f'=SUM({get_column_letter(cell.column)}39' \
                                                             f'-{get_column_letter(cell.column)}42-' \
                                                             f'{get_column_letter(cell.column)}43-' \
                                                             f'{get_column_letter(cell.column)}44-' \
                                                             f'{get_column_letter(cell.column)}45-' \
                                                             f'{get_column_letter(cell.column)}46-' \
                                                             f'{get_column_letter(cell.column)}47) '
        for index, row in enumerate(
                sheet.iter_rows(min_row=49, min_col=7, max_row=49, max_col=sheet.max_column)):
            for cell in row:
                sheet[
                    f'{get_column_letter(cell.column)}49'] = f'=SUM({get_column_letter(cell.column)}32)'
        for index, row in enumerate(
                sheet.iter_rows(min_row=50, min_col=7, max_row=50, max_col=sheet.max_column)):
            for cell in row:
                sheet[
                    f'{get_column_letter(cell.column)}50'] = f'=SUM({get_column_letter(cell.column)}48-' \
                                                             f'{get_column_letter(cell.column)}49)'
        for index, row in enumerate(
                sheet.iter_rows(min_row=49, min_col=7, max_row=49, max_col=sheet.max_column)):
            for cell in row:
                sheet[
                    f'{get_column_letter(cell.column)}49'] = f'=SUMIFS($G32:${get_column_letter(sheet.max_column)}32,' \
                                                             f'$G4:${get_column_letter(sheet.max_column)}4,' \
                                                             f'{get_column_letter(cell.column)}4)'

        # MERGE CELLS

        merge_master = []
        merge_start = []
        merge_end = []

        # CREATE MERGE LISTS Loop through all the cells in row 4 until the last column and insert the column number
        # and the cell value as a dict into the merge_start list
        for index, row in enumerate(sheet.iter_rows(min_row=4, min_col=6, max_row=4, max_col=sheet.max_column)):
            for cell in row:
                merge_master.append({'column': cell.column, 'value': cell.value})

        # Loop through merge_master, from index 1 to the end of the list, if the value of 'value' is different to the
        # value of the previous item in the list, add the column number to the merge_start list
        for index, item in enumerate(merge_master):
            if index > 0:
                if item['value'] != merge_master[index - 1]['value']:
                    merge_start.append(item['column'])

            # Loop through merge_master, from index 1 to the end of the list, if the value of 'value' is different to
            # the value of the next item in the list, add the column number to the merge_end list
            if index > 0 and index < len(merge_master) - 1:
                if item['value'] != merge_master[index + 1]['value']:
                    merge_end.append(item['column'])

        merge_end.append(merge_master[len(merge_master) - 1]['column'])

        rows_to_format_currency = [11, 13, 14, 15, 16, 17, 19, 24, 25, 26, 29, 30, 31, 34, 35, 36, 37, 39, 40, 41, 42,
                                   43,
                                   44, 45, 46, 47, 48, 49, 50]
        # Loop through rows_to_format_currency list and format the cells from column 6 to the last column as currency
        # with 2 decimal places and comma every 3 digits
        for row in rows_to_format_currency:
            for col in sheet.iter_cols(min_row=row, min_col=2, max_row=row, max_col=sheet.max_column):
                for cell in col:
                    if row == 11:
                        cell.font = Font(bold=True, color='FFFFFF')
                        cell.number_format = '0.00%'
                    else:
                        cell.font = Font(bold=True, color='FFFFFF')
                        cell.number_format = 'R#,##0.00'

        # format 'B5' as white text

        rows_to_center = [5, 6, 7, 9, 10, 11, 13, 14, 15, 16, 17, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31,
                          32, 34, 35, 36, 37, 39, 40, 41, 42, 43, 44, 45, 46, 47, 48, 49, 50]
        # Loop through the rows_to_format_currency list and align the cells from column 6 to the last column in the
        # centre
        for row in rows_to_center:
            # COLUMNS TO FILL ETC
            if row in [9, 10, 11, 20, 21, 22, 23, 27, 28]:
                col_number = 7
            else:
                col_number = 2
            # DARKER COLOR

            if row in [23, 28, 13, 19, 32, 37, 41]:
                color_is = '3E54AC'
            elif row in [50]:
                color_is = 'E14D2A'
            else:
                color_is = '537FE7'
            for col in sheet.iter_cols(min_row=row, min_col=col_number, max_row=row, max_col=sheet.max_column):
                for cell in col:
                    if row == 11:
                        if cell.value == 0.18:
                            color_is = 'EA047E'
                        elif cell.value == 0.16:
                            color_is = 'ED50F1'
                        elif cell.value == 0.15:
                            color_is = 'FF4A4A'
                        elif cell.value == 0.14:
                            color_is = 'FF9E9E'
                        # sheet[f'{get_column_letter(cell.column)}11'].number_format = '0.00%'
                        # sheet[f'{get_column_letter(cell.column)}11'].alignment = Alignment(horizontal='center')

                    cell.fill = PatternFill(start_color=color_is, end_color=color_is, fill_type='solid')
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                    cell.font = Font(bold=True, color='FFFFFF')
                    cell.border = Border(left=Side(border_style='medium', color='000000'),
                                         right=Side(border_style='medium', color='000000'),
                                         top=Side(border_style='medium', color='000000'),
                                         bottom=Side(border_style='medium', color='000000'))

        rows_to_hide = [2, 3, 4]
        # Loop through the rows_to_hide list and hide the rows
        for row in rows_to_hide:
            sheet.row_dimensions[row].hidden = True

        # Set column 1 to 25 units wide and make column 1 bold
        for col in sheet.iter_cols(min_row=1, min_col=1, max_row=ws.max_row, max_col=5):
            for cell in col:
                sheet.column_dimensions[cell.column_letter].width = 30
                cell.font = Font(bold=True)

        # ROWS TO SUM
        rows_to_sum = [13, 14, 15, 16, 17, 19, 24, 25, 26, 29, 30, 31, 32, 34, 35, 36, 37, 39, 40, 41, 42, 43, 44, 45,
                       46, 47, 48, 49, 50]
        # for each row in rows_to_sum, sum the cells from column 6 to the last column in the sheet in insert the
        # formula in column 'B', with white font and bold and format the cell as currency with 2 decimal places and
        # comma every 3 digits

        columns_to_format = ['B', 'C', 'D', 'E', 'F']

        for row in rows_to_sum:
            sheet[f'B{row}'] = f'=SUM({get_column_letter(7)}{row}:{get_column_letter(sheet.max_column)}{row})'
            for column in columns_to_format:
                sheet[f'{column}{row}'].font = Font(bold=True, color='FFFFFF')
                sheet[f'{column}{row}'].number_format = 'R#,##0.00'
                sheet[f'{column}{row}'].alignment = Alignment(horizontal='center')
                if column == 'D':
                    sheet[
                        f'{column}{row}'] = f'=SUMIFS($G{row}:${get_column_letter(sheet.max_column)}{row},' \
                                            f'$G7:${get_column_letter(sheet.max_column)}7,' \
                                            f'"Yes")'

                if column == 'E':
                    sheet[
                        f'{column}{row}'] = f'=SUMIFS($G{row}:${get_column_letter(sheet.max_column)}{row},' \
                                            f'$G6:${get_column_letter(sheet.max_column)}6,' \
                                            f'"Yes")-SUMIFS($G{row}:${get_column_letter(sheet.max_column)}{row},' \
                                            f'$G7:${get_column_letter(sheet.max_column)}7,' \
                                            f'"Yes")'

                if column == 'F':
                    sheet[
                        f'{column}{row}'] = f'=B{row}-D{row}-E{row}'

        rows_to_merge = [5, 6, 7, 13, 14, 15, 16, 17, 39, 40, 41, 42, 43, 44, 45, 46, 47, 48, 49, 50]
        # Loop through the rows_to_merge list and merge the cells in the merge_start and merge_end lists
        for row in rows_to_merge:
            for index, item in enumerate(merge_start):
                sheet.merge_cells(start_row=row, start_column=item, end_row=row, end_column=merge_end[index])

        # Merge cells in row 6 and 7 in column B
        cells_to_merge = [6]
        columns_affected = ['B', 'C', 'D', 'E', 'F']
        for columns_affected in columns_affected:
            for row in cells_to_merge:
                sheet.merge_cells(f'{columns_affected}{row}:{columns_affected}{row + 1}')

        sheet.merge_cells('B6:B7')

        # Format 'B5' to 'F7' with white font and bold
        for row in sheet.iter_rows(min_row=5, min_col=2, max_row=7, max_col=6):
            for cell in row:
                cell.font = Font(bold=True, color='FFFFFF')

    # CREATE NSST SHEET
    if len(category) > 1:
        sheet_name = f"NSST {category[0].split(' ')[0]}"

    else:
        sheet_name = f"NSST {category[0]}"

    print(sheet_name)
    for sheet in wb:
        print("XXXX", sheet.title)

    worksheets = wb.sheetnames
    print("worksheets", worksheets)
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

        print("category XXX",category[1])
        if len(pledges):
            new_pledges = [pledge for pledge in pledges if pledge['Category'] == category[1]
                           and pledge['opportunity_code'] != 'Unallocated']
        else:
            new_pledges = []
        index = 2
        sheet_name = f"NSST {category[1]}"
        print("sheet_name XXX", sheet_name)
        ws = wb.create_sheet(sheet_name)
        # Make Tab Color of new Sheet 'Green'
        ws.sheet_properties.tabColor = "539165"
        nsst_data = create_nsst_sheet(category, developmentinputdata, new_pledges, index, sheet_name, worksheets)

        for item in nsst_data:
            ws.append(item)

    num_sheets = len(wb.sheetnames)
    print("num_sheets", num_sheets)

    for index, sheet in enumerate(wb.worksheets):
        print("index", index)
        print("sheet", sheet)
        sheet_index = []
        if num_sheets == 2:
            sheet_index = [1]
        else:
            sheet_index = [3, 4, 5]


    # SAVE TO FILE
    wb.save(f"excel_files/{filename}.xlsx")
