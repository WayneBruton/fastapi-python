import copy


def create_nsst_sheet(category, developmentinputdata, pledges, index, sheet_name, worksheets):
    # if pledges is not empty, then get the total of all the investment_amounts in the pledges list else set
    # total_pledges to 0
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
    nsst_data.append(["Available to be raised (Estimated)", f'=+\'{worksheets[index]}\'!B17'])
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
        ["Sales Price", f'=+\'{worksheets[index]}\'!B41', f'=+\'{worksheets[index]}\'!D41',
         f'=+\'{worksheets[index]}\'!E41',
         f'=+\'{worksheets[index]}\'!F41'])
    nsst_data.append(
        ["VAT", f'=+\'{worksheets[index]}\'!B42', f'=+\'{worksheets[index]}\'!D42', f'=+\'{worksheets[index]}\'!E42',
         f'=+\'{worksheets[index]}\'!F42'])
    nsst_data.append(
        ["Gross", f'=+\'{worksheets[index]}\'!B43', f'=+\'{worksheets[index]}\'!D43', f'=+\'{worksheets[index]}\'!E43',
         f'=+\'{worksheets[index]}\'!F43'])
    nsst_data.append(["Commission (5 %)", f'=+\'{worksheets[index]}\'!B44', f'=+\'{worksheets[index]}\'!D44',
                      f'=+\'{worksheets[index]}\'!E44', f'=+\'{worksheets[index]}\'!F44'])
    nsst_data.append(
        ["Transfer Fees", f'=+\'{worksheets[index]}\'!B45', f'=+\'{worksheets[index]}\'!D45',
         f'=+\'{worksheets[index]}\'!E45',
         f'=+\'{worksheets[index]}\'!F45'])
    nsst_data.append(["Bond Registration", f'=+\'{worksheets[index]}\'!B46', f'=+\'{worksheets[index]}\'!D46',
                      f'=+\'{worksheets[index]}\'!E46', f'=+\'{worksheets[index]}\'!F46'])
    nsst_data.append(["Security Release Fee", f'=+\'{worksheets[index]}\'!B47', f'=+\'{worksheets[index]}\'!D47',
                      f'=+\'{worksheets[index]}\'!E47', f'=+\'{worksheets[index]}\'!F47'])
    nsst_data.append(["Unforseen (0.05%)", f'=+\'{worksheets[index]}\'!B48', f'=+\'{worksheets[index]}\'!D48',
                      f'=+\'{worksheets[index]}\'!E48', f'=+\'{worksheets[index]}\'!F48'])
    nsst_data.append(["Discount", f'=+\'{worksheets[index]}\'!B49', f'=+\'{worksheets[index]}\'!D49',
                      f'=+\'{worksheets[index]}\'!E49', f'=+\'{worksheets[index]}\'!F49'])
    nsst_data.append(["Transfer Income", ""])
    nsst_data.append([])
    nsst_data.append(["CAPITAL"])
    nsst_data.append(
        ["Total Capital Available to be raised", f'=+\'{worksheets[index]}\'!B13-B37', f'=+\'{worksheets[index]}\'!D13',
         f'=+\'{worksheets[index]}\'!E13', f'=+\'{worksheets[index]}\'!F13'])

    nsst_data.append(["Capital Drawn down", f'=+\'{worksheets[index]}\'!B15', f'=+\'{worksheets[index]}\'!D15',
                      f'=+\'{worksheets[index]}\'!E15', f'=+\'{worksheets[index]}\'!F15'])

    nsst_data.append(["Capital - Early repayment", f'=+\'{worksheets[index]}\'!B66', 0,0, 0])

    nsst_data.append(["Current Capital deployed", 0, 0, 0, 0])
    nsst_data.append(
        ["Capital Available for deployment", f'=+\'{worksheets[index]}\'!B16', f'=+\'{worksheets[index]}\'!D16',
         f'=+\'{worksheets[index]}\'!E16', f'=+\'{worksheets[index]}\'!F16'])
    nsst_data.append(["Available security for Capital to be raised", f'=+\'{worksheets[index]}\'!B17',
                      f'=+\'{worksheets[index]}\'!D17',
                      f'=+\'{worksheets[index]}\'!E17', f'=+\'{worksheets[index]}\'!F17'])
    nsst_data.append([])
    nsst_data.append(["INTEREST"])
    nsst_data.append(["Total Interest", f'=+\'{worksheets[index]}\'!B33+B46', f'=+\'{worksheets[index]}\'!D33',
                      f'=+\'{worksheets[index]}\'!E33', f'=+\'{worksheets[index]}\'!F33'])
    nsst_data.append(["Capital Repaid", 0, f'=+\'{worksheets[index]}\'!D33',
                      0, 0])
    nsst_data.append(["Interest on Capital Drawn to estimated Exit date", f'=+\'{worksheets[index]}\'!B31',
                      f'=+\'{worksheets[index]}\'!D31',
                      f'=+\'{worksheets[index]}\'!E31', f'=+\'{worksheets[index]}\'!F31'])

    nsst_data.append(["Interest - Early Repayments", f'=+\'{worksheets[index]}\'!B67', 0, 0, 0])

    nsst_data.append(["Interest on Capital to Be Drawn from Momentum", f'=+\'{worksheets[index]}\'!B30',
                      f'=+\'{worksheets[index]}\'!D30',
                      f'=+\'{worksheets[index]}\'!E30', f'=+\'{worksheets[index]}\'!F30'])
    nsst_data.append(
        ["Interest on Capital to be raised", f'=+\'{worksheets[index]}\'!B32', f'=+\'{worksheets[index]}\'!D32',
         f'=+\'{worksheets[index]}\'!E32', f'=+\'{worksheets[index]}\'!F32'])
    nsst_data.append(
        ["Interest due to estimated exit dates", 0, 0, f'=+\'{worksheets[index]}\'!E33',
         f'=+\'{worksheets[index]}\'!F33-\'{worksheets[index]}\'!F46'])
    nsst_data.append([])
    nsst_data.append(["PROJECTED GROSS PROFIT"])
    nsst_data.append(["Income after repayment of Capital & Interest", 0, 0, 0, 0])
    nsst_data.append(["Other Income", 0, 0, 0, 0])
    nsst_data.append(["Cost to complete funding requirement", 0, 0, 0, 0])
    nsst_data.append(["Capital available for funding", 0, 0, 0, 0])
    nsst_data.append(["Balance after funding utilization_Gross Profit", 0, 0, 0, 0])

    return nsst_data
