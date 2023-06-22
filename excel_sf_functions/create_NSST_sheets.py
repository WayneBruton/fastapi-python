import copy


def create_nsst_sheet(category, developmentinputdata, pledges, index, sheet_name, worksheets):
    # if pledges is not empty, then get the total of all the investment_amounts in the pledges list else set
    # total_pledges to 0

    index_for_investor_report = len(worksheets) - 1
    # ='Investor Exit List Endulini Sbe'!Q3 - 'Investor Exit List Endulini Sbe'!E3

    if index == 0:
        totalFigure = f'=+\'{worksheets[index_for_investor_report]}\'!Q3-\'{worksheets[index_for_investor_report]}\'!E3'
        soldFigure = f'=+sumifs(\'{worksheets[index_for_investor_report]}\'!Q4:Q1000,' \
                     f'\'{worksheets[index_for_investor_report]}\'!G4:G1000, True,' \
                     f'\'{worksheets[index_for_investor_report]}\'!J4:J1000, "=")' \
                     f'-sumifs(\'{worksheets[index_for_investor_report]}\'!E4:E1000,' \
                     f'\'{worksheets[index_for_investor_report]}\'!G4:G1000, True,' \
                     f'\'{worksheets[index_for_investor_report]}\'!J4:J1000, "=")'

    elif index == 1:
        # WORKING ON THIS

        totalFigure = f'=+SUMIFS(\'{worksheets[index_for_investor_report]}\'!Q4: Q1000, ' \
                      f'\'{worksheets[index_for_investor_report]}\'!J4: J1000, "=",' \
                      f'\'{worksheets[index_for_investor_report]}\'!C4: C1000, "A")' \
                      f'+SUMIFS(\'{worksheets[index_for_investor_report]}\'!Q4: Q1000, ' \
                      f'\'{worksheets[index_for_investor_report]}\'!J4: J1000, "=", ' \
                      f'\'{worksheets[index_for_investor_report]}\'!C4: C1000, "B")' \
                      f'-SUMIFS(\'{worksheets[index_for_investor_report]}\'!E4: E1000, ' \
                      f'\'{worksheets[index_for_investor_report]}\'!J4: J1000, "=", ' \
                      f'\'{worksheets[index_for_investor_report]}\'!C4: C1000, "A")' \
                      f'-SUMIFS(\'{worksheets[index_for_investor_report]}\'!E4: E1000, ' \
                      f'\'{worksheets[index_for_investor_report]}\'!J4: J1000, "=", ' \
                      f'\'{worksheets[index_for_investor_report]}\'!C4: C1000, "B")'

        soldFigure = f'=+SUMIFS(\'{worksheets[index_for_investor_report]}\'!Q4: Q1000,' \
                     f' \'{worksheets[index_for_investor_report]}\'!J4: J1000, "=", ' \
                     f'\'{worksheets[index_for_investor_report]}\'!C4: C1000, "A",' \
                     f'\'{worksheets[index_for_investor_report]}\'!G4:G1000, TRUE)' \
                     f'+SUMIFS(\'{worksheets[index_for_investor_report]}\'!Q4: Q1000, ' \
                     f'\'{worksheets[index_for_investor_report]}\'!J4: J1000, "=", ' \
                     f'\'{worksheets[index_for_investor_report]}\'!C4: C1000, "B",' \
                     f'\'{worksheets[index_for_investor_report]}\'!G4:G1000, TRUE)' \
                     f'-SUMIFS(\'{worksheets[index_for_investor_report]}\'!E4: E1000, ' \
                     f'\'{worksheets[index_for_investor_report]}\'!J4: J1000, "=", ' \
                     f'\'{worksheets[index_for_investor_report]}\'!C4: C1000, "A",' \
                     f'\'{worksheets[index_for_investor_report]}\'!G4:G1000, TRUE)' \
                     f'-SUMIFS(\'{worksheets[index_for_investor_report]}\'!E4: E1000, ' \
                     f'\'{worksheets[index_for_investor_report]}\'!J4: J1000, "=", ' \
                     f'\'{worksheets[index_for_investor_report]}\'!C4: C1000, "B",' \
                     f'\'{worksheets[index_for_investor_report]}\'!G4:G1000, TRUE)'

    elif index == 2:
        'NSST Heron'
        'NSST Heron Fields'
        totalFigure = f"='NSST Heron'!B49 - 'NSST Heron Fields'!B49"
        soldFigure = f"='NSST Heron'!D49 - 'NSST Heron Fields'!D49"

    total_pledges = sum([float(item['investment_amount']) for item in pledges]) if pledges else 0

    # deep copy category list so that changes to the copy do not affect the original
    category_new = copy.deepcopy(category)
    # if index == 0 and category length equals 2 then insert f"{category[0]} and {category[1]}" into the begining of
    # category list
    if index == 0 and len(category_new) == 2:
        category_new.insert(0, f"{category_new[0]} and {category_new[1]}")
    if index == 0:
        category_index = 0
    elif index == 1:
        category_index = 0
    elif index == 2:
        category_index = 1
    nsst_data = []
    row1_data = [f"{sheet_name.upper()} INVESTOR REPORT - {developmentinputdata['date']}"]
    nsst_data.append(row1_data)
    nsst_data.append([])
    nsst_data.append([])
    nsst_data.append(["Report Date", developmentinputdata['date']])
    nsst_data.append(["Development", category_new[category_index]])
    nsst_data.append([])
    nsst_data.append(["CAPITAL"])
    nsst_data.append(["Total Investment capital to be raised (Estimated)", f'=+\'{worksheets[index]}\'!B13'])
    nsst_data.append(["Available to be raised (Estimated)", f'=+B40'])
    nsst_data.append(["Total Investment capital received", f'=+\'{worksheets[index]}\'!B14'])
    nsst_data.append(["Total Funds Drawn Down into Development", f'=+\'{worksheets[index]}\'!B15'])
    nsst_data.append(["Pledges Due", total_pledges])
    nsst_data.append(["Momentum Investment Account", f'=+\'{worksheets[index]}\'!B16'])
    nsst_data.append([])
    nsst_data.append(["INVESTMENTS"])
    nsst_data.append(["No. of Capital Investments received", f'=+\'{worksheets[index]}\'!B3'])
    nsst_data.append(["No. Investments exited to date", f'=+\'{worksheets[index]}\'!D3'])
    nsst_data.append(
        ["No. Investments still in Development", f'=+\'{worksheets[index]}\'!B3-\'{worksheets[index]}\'!D3'])
    nsst_data.append([])
    nsst_data.append(["GROSS INCOME"])
    nsst_data.append(["", "Total", "Transferred", "Sold", "Remaining"])
    nsst_data.append(
        ["Units", f'=+\'{worksheets[index]}\'!B6', f'=+\'{worksheets[index]}\'!D6', f'=+\'{worksheets[index]}\'!E6',
         f'=+\'{worksheets[index]}\'!F6'])
    nsst_data.append(
        ["Sales Price", f'=+\'{worksheets[index]}\'!B42', f'=+\'{worksheets[index]}\'!D42',
         f'=+\'{worksheets[index]}\'!E42',
         f'=+\'{worksheets[index]}\'!F42'])
    nsst_data.append(
        ["VAT", f'=+\'{worksheets[index]}\'!B43', f'=+\'{worksheets[index]}\'!D43', f'=+\'{worksheets[index]}\'!E43',
         f'=+\'{worksheets[index]}\'!F43'])
    nsst_data.append(
        ["Gross", f'=+\'{worksheets[index]}\'!B44', f'=+\'{worksheets[index]}\'!D44', f'=+\'{worksheets[index]}\'!E44',
         f'=+\'{worksheets[index]}\'!F44'])
    nsst_data.append(["Commission (5.75 %)", f'=+\'{worksheets[index]}\'!B45', f'=+\'{worksheets[index]}\'!D45',
                      f'=+\'{worksheets[index]}\'!E45', f'=+\'{worksheets[index]}\'!F45'])
    nsst_data.append(
        ["Transfer Fees", f'=+\'{worksheets[index]}\'!B46', f'=+\'{worksheets[index]}\'!D46',
         f'=+\'{worksheets[index]}\'!E46',
         f'=+\'{worksheets[index]}\'!F46'])
    nsst_data.append(["Bond Registration", f'=+\'{worksheets[index]}\'!B47', f'=+\'{worksheets[index]}\'!D47',
                      f'=+\'{worksheets[index]}\'!E47', f'=+\'{worksheets[index]}\'!F47'])
    nsst_data.append(["Security Release Fee", f'=+\'{worksheets[index]}\'!B48', f'=+\'{worksheets[index]}\'!D48',
                      f'=+\'{worksheets[index]}\'!E48', f'=+\'{worksheets[index]}\'!F48'])
    nsst_data.append(["Unforseen (0.05%)", f'=+\'{worksheets[index]}\'!B49', f'=+\'{worksheets[index]}\'!D49',
                      f'=+\'{worksheets[index]}\'!E49', f'=+\'{worksheets[index]}\'!F49'])
    nsst_data.append(["Discount", f'=+\'{worksheets[index]}\'!B50', f'=+\'{worksheets[index]}\'!D50',
                      f'=+\'{worksheets[index]}\'!E50', f'=+\'{worksheets[index]}\'!F50'])
    nsst_data.append(["Transfer Income", ""])
    nsst_data.append([])
    nsst_data.append(["CAPITAL"])
    nsst_data.append(
        ["Total Capital Available to be raised", f'=+\'{worksheets[index]}\'!B13', f'=+\'{worksheets[index]}\'!D13',
         f'=+\'{worksheets[index]}\'!E13', f'=+\'{worksheets[index]}\'!F13'])

    nsst_data.append(["Capital Drawn down", f'=+\'{worksheets[index]}\'!B15', f'=+\'{worksheets[index]}\'!D15',
                      f'=+\'{worksheets[index]}\'!E15', f'=+\'{worksheets[index]}\'!F15'])

    nsst_data.append(["Developer Capital - Early Investor Exit", f'=+\'{worksheets[index]}\'!B67', 0,
                      f'=+\'{worksheets[index]}\'!B86', f'=+\'{worksheets[index]}\'!B83'])

    nsst_data.append(["Current Investor Capital deployed", 0, 0, 0, 0])
    nsst_data.append(
        ["Capital Available for deployment", f'=+\'{worksheets[index]}\'!B16', f'=+\'{worksheets[index]}\'!D16',
         f'=+\'{worksheets[index]}\'!E16', f'=+\'{worksheets[index]}\'!F16'])
    nsst_data.append(["Available security for Capital to be raised", f'=+\'{worksheets[index]}\'!B17',
                      f'=+\'{worksheets[index]}\'!D17',
                      f'=+\'{worksheets[index]}\'!E17', f'=+\'{worksheets[index]}\'!F17'])
    nsst_data.append([])
    nsst_data.append(["INTEREST"])
    nsst_data.append(["Total Interest", f'=+\'{worksheets[index]}\'!B34', f'=+\'{worksheets[index]}\'!D34',
                      f'=+\'{worksheets[index]}\'!E34', f'=+\'{worksheets[index]}\'!F34'])
    nsst_data.append(["Interest on Capital Repaid", 0, f'=+\'{worksheets[index]}\'!D34',
                      0, 0])
    nsst_data.append(["Total Interest on Total Capital Drawn to estimated Exit date", f'=+\'{worksheets[index]}\'!B32',
                      f'=+\'{worksheets[index]}\'!D32',
                      f'=+\'{worksheets[index]}\'!E32', f'=+\'{worksheets[index]}\'!F32'])

    nsst_data.append(
        ["Investor Early Exit interest", f'=+\'{worksheets[index]}\'!B68', 0, f'=+\'{worksheets[index]}\'!B87',
         f'=+\'{worksheets[index]}\'!B84'])

    nsst_data.append(
        ["Interest on Capital to Be Drawn from Momentum", f'=+\'{worksheets[index]}\'!B30+\'{worksheets[index]}\'!B31',
         f'=+\'{worksheets[index]}\'!D30+\'{worksheets[index]}\'!D31',
         f'=+\'{worksheets[index]}\'!E30+\'{worksheets[index]}\'!E31',
         f'=+\'{worksheets[index]}\'!F30+\'{worksheets[index]}\'!F31'])
    nsst_data.append(
        ["Interest on Capital to be raised", f'=+\'{worksheets[index]}\'!B33', f'=+\'{worksheets[index]}\'!D33',
         f'=+\'{worksheets[index]}\'!E33', f'=+\'{worksheets[index]}\'!F33'])
    nsst_data.append(
        ["Interest due on deployed capital exit dates", totalFigure, 0, soldFigure,
         f'=+\'{worksheets[index]}\'!F34-\'{worksheets[index]}\'!F47'])
    nsst_data.append(
        ["Developer Interest earned from Investment Account", f'=+\'{worksheets[index]}\'!B30',
         f'=+\'{worksheets[index]}\'!D30', f'=+\'{worksheets[index]}\'!E30',
         f'=+\'{worksheets[index]}\'!F30'])
    nsst_data.append([])
    nsst_data.append(["PROJECTED GROSS PROFIT"])
    nsst_data.append(["Income after repayment of Capital & Interest", 0, 0, 0, 0])
    nsst_data.append(["Other Income", 0, 0, 0, 0])
    nsst_data.append(["Cost to complete funding requirement", 0, 0, 0, 0])
    nsst_data.append(["Capital available for funding", 0, 0, 0, 0])
    nsst_data.append(["Balance after funding utilization_Gross Profit", 0, 0, 0, 0])

    return nsst_data
