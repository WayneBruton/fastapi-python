from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.styles import Font, NumberFormatDescriptor, PatternFill, Color
from openpyxl.styles.borders import Border

from openpyxl.styles import borders
from openpyxl.styles.alignment import Alignment


# from openpyxl.styles.borders import Border


def create_sales_forecast_file(data, developmentinputdata):
    # print(developmentinputdata)
    filename = 'Sales Forecast'
    sheet_name = "Sales Forecast"
    print(filename)

    wb = Workbook()
    ws = wb.active
    ws.title = filename

    opportunity_code, opportunity_amount_required, opportunity_end_date, opportunity_final_transfer_date, \
    opportunity_sale_price, opportunity_sold, opportunity_transferred, report_date, trust_release_fee, \
    raising_commission, structuring_fee, commission, transfer_fees, bond_registration, unforseen, \
    investment_amount, investor_acc_number, investor_name, deposit_date, release_date, investment_number, \
    released_investment_amount, investment_end_date, investment_release_date, released_investment_number, \
    investment_interest_rate, rollover_amount, rollover_date, investment_interest_total, \
    investment_interest_to_date, released_interest_total, released_interest_to_date = [], [], [], [], [], [], \
                                                                                      [], [], [], [], [], [], [], \
                                                                                      [], [], [], [], [], [], [], \
                                                                                      [], [], [], [], [], [], [], \
                                                                                      [], [], [], [], []
    final_insert_data = []

    cell_labels = ["Opportunity",
                   "opportunity_amount_required",
                   "opportunity_end_date",
                   "opportunity_final_transfer_date",
                   "opportunity_sale_price",
                   "opportunity_sold",
                   "opportunity_transferred",
                   "report_date",
                   "trust_release_fee",
                   "raising_commission",
                   "structuring_fee",
                   "commission",
                   "transfer_fees",
                   "bond_registration",
                   "unforseen",
                   "investment_amount",
                   "investor_acc_number",
                   "investor_name",
                   "deposit_date",
                   "release_date",
                   "investment_number",
                   "released_investment_amount",
                   "investment_end_date",
                   "investment_release_date",
                   "released_investment_number",
                   "investment_interest_rate",
                   "rollover_amount",
                   "rollover_date",
                   "investment_interest_total",
                   "investment_interest_to_date",
                   "released_interest_total",
                   "released_interest_to_date", ]

    # APPEND DATA
    for info in data:
        opportunity_code.append(info['opportunity_code'])
        opportunity_amount_required.append(info['opportunity_amount_required'])
        opportunity_end_date.append(info['opportunity_end_date'])
        opportunity_final_transfer_date.append(info['opportunity_final_transfer_date'])
        opportunity_sale_price.append(info['opportunity_sale_price'])
        opportunity_sold.append(info['opportunity_sold'])
        opportunity_transferred.append(info['opportunity_transferred'])
        report_date.append(info['report_date'])
        trust_release_fee.append(info['trust_release_fee'])
        raising_commission.append(info['raising_commission'])
        structuring_fee.append(info['structuring_fee'])
        commission.append(info['commission'])
        transfer_fees.append(info['transfer_fees'])
        bond_registration.append(info['bond_registration'])
        unforseen.append(info['unforseen'])
        investment_amount.append(info['investment_amount'])
        investor_acc_number.append(info['investor_acc_number'])
        investor_name.append(info['investor_name'])
        deposit_date.append(info['deposit_date'])
        release_date.append(info['release_date'])
        investment_number.append(info['investment_number'])
        released_investment_amount.append(info['released_investment_amount'])
        investment_end_date.append(info['investment_end_date'])
        investment_release_date.append(info['investment_release_date'])
        released_investment_number.append(info['released_investment_number'])
        investment_interest_rate.append(info['investment_interest_rate'] / 100)
        rollover_amount.append(info['rollover_amount'])
        rollover_date.append(info['rollover_date'])
        investment_interest_total.append(info['investment_interest_total'])
        investment_interest_to_date.append(info['investment_interest_to_date'])
        released_interest_total.append(info['released_interest_total'])
        released_interest_to_date.append(info['released_interest_to_date'])

    # INSERT DATA INTO A SINGLE LIST
    final_insert_data.append(opportunity_code)
    final_insert_data.append(opportunity_amount_required)
    final_insert_data.append(opportunity_end_date)
    final_insert_data.append(opportunity_final_transfer_date)
    final_insert_data.append(opportunity_sale_price)
    final_insert_data.append(opportunity_sold)
    final_insert_data.append(opportunity_transferred)
    final_insert_data.append(report_date)
    final_insert_data.append(trust_release_fee)
    final_insert_data.append(raising_commission)
    final_insert_data.append(structuring_fee)
    final_insert_data.append(commission)
    final_insert_data.append(transfer_fees)
    final_insert_data.append(bond_registration)
    final_insert_data.append(unforseen)
    final_insert_data.append(investment_amount)
    final_insert_data.append(investor_acc_number)
    final_insert_data.append(investor_name)
    final_insert_data.append(deposit_date)
    final_insert_data.append(release_date)
    final_insert_data.append(investment_number)
    final_insert_data.append(released_investment_amount)
    final_insert_data.append(investment_end_date)
    final_insert_data.append(investment_release_date)
    final_insert_data.append(released_investment_number)
    final_insert_data.append(investment_interest_rate)
    final_insert_data.append(rollover_amount)
    final_insert_data.append(rollover_date)
    final_insert_data.append(investment_interest_total)
    final_insert_data.append(investment_interest_to_date)
    final_insert_data.append(released_interest_total)
    final_insert_data.append(released_interest_to_date)

    # APPEND SINGLE LIST INTO EXCEL
    for row in final_insert_data:
        ws.append(row)

    # MOVE DATA FROM TOP OF FILE
    ws.move_range(f"A1:{get_column_letter(ws.max_column)}{str(ws.max_row)}", rows=5, cols=6)

    # PUT LABELS INTO COLUMN A
    for x in range(6, 38):
        ws[f"A{x}"].value = cell_labels[x - 6]

    ws["B6"].value = "Total"
    ws["C6"].value = "Transferred"
    ws["D6"].value = "Sold"
    ws["E6"].value = "Remaining"
    ws["F6"].value = "Check"

    # REORGANISE DATA INTO CORRECT ORDER

    last_column = get_column_letter(ws.max_column)
    last_column_number = ws.max_column

    border1 = borders.Side(style=None, color=Color(indexed=0), border_style='medium')
    # border0 = borders.Side(style=None, color=None, border_style=None)
    medium = Border(left=border1, right=border1, bottom=border1, top=border1)

    rows_to_format_from_g = [6, 7, 10, 31]

    for specific_row in rows_to_format_from_g:
        for row in ws.iter_rows(min_row=specific_row, min_col=7, max_row=specific_row, max_col=last_column_number):
            for i, cell in enumerate(row):
                # cell.fill = PatternFill(bgColor="000000FF", fill_type="solid")
                cell.fill = PatternFill("solid", start_color="000066CC")
                cell.border = medium
                cell.font = Font(color="00FFFFFF")
                cell.alignment = Alignment(horizontal="center")
                if specific_row == 7 or specific_row == 10:
                    cell.number_format = '"R"#,##0.00_);[Red]("R"#,##0.00)'
                if specific_row == 31:
                    cell.number_format = '0%'

    # PREPARE FOR CELL MERGING

    merge_cells_start = []
    merge_cells_end = []
    rows_for_merging = [6, 7]

    for row in ws.iter_rows(min_row=6, min_col=7, max_row=6, max_col=last_column_number):
        for i, cell in enumerate(row):
            if row[i].value != row[i - 1].value:
                end_cell = f"{get_column_letter(i + 6)}"
                start_cell = f"{get_column_letter(i + 7)}"
                merge_cells_start.append(start_cell)
                merge_cells_end.append(end_cell)

    merge_cells_end.append(f"{last_column}")
    merge_cells_end = merge_cells_end[1:]

    # MERGE CELLS
    for row_number in rows_for_merging:
        for i, start in enumerate(merge_cells_start):
            merge_range = f"{start}{row_number}:{merge_cells_end[i]}{row_number}"
            ws.merge_cells(merge_range)

    # SAVE TO FILE
    wb.save(f"excel_files/{filename}.xlsx")
