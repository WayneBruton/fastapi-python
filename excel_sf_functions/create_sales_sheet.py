# CREATE HELPER FUNCTIONS
from datetime import datetime, timedelta


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
    row4A_data = ["Sold", "", "", "", "", ""]
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
    row25A_data = ["Interest on funds to be raised", "", "", "", "", ""]
    row26_data = ["Total Interest to Date", "", "", "", "", ""]
    row27_data = ["Contract End Date", "", "", "", "", ""]
    row28_data = ["Days to exit", "", "", "", "", ""]
    row29_data = ["Investment Account Interest Earned.", "", "", "", "", ""]
    row30AA_data = ["Supplemented Interest (2.75%)", "", "", "", "", ""]
    row30_data = ["Released Interest Earned.", "", "", "", "", ""]

    row30A_data = ["Interest on funds to be raised", "", "", "", "", ""]
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
    row49A_data = []

    row50_data = ["Early Released", "", "", "", "", ""]
    row51_data = ["Early Exit Date", "", "", "", "", ""]
    row52_data = ["Early Exit Amount", "", "", "", "", ""]
    row53_data = ["Early Exit", "", "", "", "", ""]
    row54_data = ["Transfer Income", "", "", "", "", ""]
    row55_data = ["Due to Investors (Adjusted)", "", "", "", "", ""]
    row56_data = ["Profit / Loss (Adjusted)", "", "", "", "", ""]
    rowblank_tfr_data = ["Transfer Date", "", "", "", "", ""]
    rowblank_1_data = ["RENTALS", "", "", "", "", ""]
    rowblank_data = []
    rental1_data = ["Marked for Rent", "", "", "", "", ""]
    rental1A_data = ["Potential Monthly Income", "", "", "", "", ""]

    rental2_data = ["Rented Out", "", "", "", "", ""]
    rental3_data = ["Deposit Held", "", "", "", "", ""]
    rental4_data = ["Gross Rent", "", "", "", "", ""]
    rental5_data = ["Levy", "", "", "", "", ""]
    rental6_data = ["Commission", "", "", "", "", ""]
    rental7_data = ["Rates", "", "", "", "", ""]
    rental8_data = ["Other", "", "", "", "", ""]
    rental9_data = ["Nett Rental", "", "", "", "", ""]
    rental10_data = ["Rental Start Date", "", "", "", "", ""]
    rental11_data = ["Rental End Date", "", "", "", "", ""]
    rental12_data = ["Income to date", "", "", "", "", ""]
    rental13_data = ["Rental Contract Income", "", "", "", "", ""]

    rollover_data1 = []
    rollover_data2 = []
    rollover_data3 = []
    rollover_data4 = ["Rollover from portal", "", "", "", "", ""]
    rollover_data5 = ["Rollover amount from portal", "", "", "", "", ""]
    rollover_data6 = ["Forecast Rollover Date", "", "", "", "", ""]
    rollover_data7 = ["Forecast Rollover Amount", "", "", "", "", ""]



    for item in data:
        row2_data.append(item['opportunity_code'])
        row4A_data.append(item['funds_in_momentum'])
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

        if item['opportunity_code'] != "ZZUN01":
            if item['early_release'] == True:
                row22_data.append(item['investment_end_date'])
            else:
                row22_data.append(item['opportunity_final_transfer_date'])
        else:
            row22_data.append("")

        if item['opportunity_final_transfer_date'] != "":
            rowblank_tfr_data.append(item['opportunity_end_date'])
        else:
            rowblank_tfr_data.append(item['opportunity_final_transfer_date'])

        row24_data.append(item['investment_interest_today'])
        row25_data.append(item['released_interest_today'])
        row25A_data.append(float(item['interest_to_date_still_to_be_raised']))
        # for row27_data if the item['opportunity_code'] is not equal to 'ZZUN01', take the item['deposit_date'],
        # replace '-' with '/' then convert it to datetime and add exactly two years to it and append the new date to
        # the list formatted as YYYY/MM/DD else just make it ""
        if item['opportunity_code'] != "ZZUN01" and item['deposit_date'] != "":
            row27_data.append(
                (datetime.strptime(item['deposit_date'], '%Y/%m/%d') + timedelta(days=730)).strftime('%Y/%m/%d'))

        else:
            row27_data.append("")
        row29_data.append(float(item['trust_interest_total']))
        row30AA_data.append(float(item['trust_interest_total'] / 100 * 2.75))
        row30_data.append(item['released_interest_total'])
        row30A_data.append(float(item['interest_total_still_to_be_raised']))
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
        row50_data.append(item['early_release'])

        rental1_data.append(item['rental_marked_for_rent'])
        rental1A_data.append(item['potential_income'])
        rental2_data.append(item['rental_rented_out'])
        rental3_data.append(item['rental_deposit_amount'])
        rental4_data.append(item['rental_gross_amount'])
        rental5_data.append(item['rental_levy_amount'])
        rental6_data.append(item['rental_commission'])
        rental7_data.append(item['rental_rates'])
        rental8_data.append(item['rental_other_expenses'])
        rental9_data.append(item['rental_nett_amount'])
        rental10_data.append(item['rental_start_date'])
        rental11_data.append(item['rental_end_date'])
        rental12_data.append(item['rental_income_to_date'])
        rental13_data.append(item['rental_income_to_contract_end'])

        rollover_data4.append(item['from_portal'])
        rollover_data5.append(item['rollover_amount_chosen'])



    worksheet_data += [row4_data, row5_data, row2_data, row2_data, row4_data, row5_data, [], row7_data, row8_data,
                       row9_data, [], row13_data, row14_data, row15_data, row16_data, row17_data, [], row19_data,
                       row20_data, row21_data, row22_data, row23_data, row24_data, row25_data, row25A_data, row26_data,
                       row27_data, row28_data, row29_data, row30AA_data, row30_data, row30A_data, row31_data,
                       row31A_data, [],
                       row33_data, row34_data, row35_data, row36_data, [], row38_data, row39_data, row40_data,
                       row41_data, row42_data, row43_data, row44_data, row45_data, row46_data, row47_data, row48_data,
                       row49_data, row4A_data, row49A_data, row49A_data, row49A_data, row50_data, row51_data,
                       row52_data, row53_data, row54_data, row55_data, row56_data, rowblank_tfr_data, rowblank_data,
                       rowblank_data, rowblank_data, rowblank_data, rowblank_data, rowblank_data, rowblank_data,
                       rowblank_data, rowblank_data, rowblank_data, rowblank_data, rowblank_data, rowblank_data,
                       rowblank_data, rowblank_data, rowblank_data, rowblank_data, rowblank_data, rowblank_data,
                       rowblank_data, rowblank_data, rowblank_data, rowblank_data, rowblank_data, rowblank_1_data,
                       rowblank_data, rental1_data,rental1A_data, rental2_data, rental3_data, rental4_data, rental5_data,
                       rental6_data, rental7_data, rental8_data, rental9_data, rental10_data, rental11_data,
                       rental12_data, rental13_data, rollover_data1, rollover_data2, rollover_data3, rollover_data4,
                       rollover_data5, rollover_data6, rollover_data7]

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
