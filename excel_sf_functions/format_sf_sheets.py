from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill
from openpyxl.styles.borders import Border, Side
from openpyxl.styles import Font, Alignment


def format_sales_forecast(sheet):
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

    # create a list of column letters from column 2 to the last column in the sheet
    column_letters_2 = [get_column_letter(x) for x in range(2, sheet.max_column + 1)]

    # splice this list into a new list starting only at column 7
    column_letters_7 = column_letters_2[5:]

    rows_to_add_formulas = [17, 23, 27, 29, 33, 34, 38, 39, 50, 51, 52, 54, 58, 59, 60, 61, 62, 63, 66, 67]

    for letter in column_letters_7:
        for row in rows_to_add_formulas:
            if row == 17:
                sheet[f'{letter}17'] = f'=SUM({letter}13-{letter}14)'
            elif row == 23:
                sheet[f'{letter}23'] = f'=SUM({letter}22-{letter}21)'
            elif row == 27:
                sheet[f'{letter}27'] = f'=SUM({letter}24+{letter}25+{letter}26)'
            elif row == 29:
                sheet[f'{letter}29'] = f'=720-{letter}23'
            elif row == 33:
                sheet[f'{letter}33'] = f'=SUM({letter}30+{letter}31+{letter}32)'
            elif row == 34:
                sheet[f'{letter}34'] = f'=SUM({letter}19+{letter}33)'
            elif row == 38:
                sheet[f'{letter}38'] = f'=SUM({letter}36+{letter}37)'
            elif row == 39:
                sheet[f'{letter}39'] = f'=SUM({letter}19-{letter}38)'
            elif row == 50:
                sheet[f'{letter}50'] = f'={letter}41-SUM({letter}44:{letter}49)'
            elif row == 51:
                sheet[f'{letter}51'] = f'=SUMIFS({column_letters_7[0]}34:' \
                                       f'{column_letters_7[len(column_letters_7) - 1]}34,{column_letters_7[0]}4:' \
                                       f'{column_letters_7[len(column_letters_7) - 1]}4, {letter}4)'
            elif row == 52:
                sheet[f'{letter}52'] = f'={letter}50-{letter}51'
            elif row == 54:
                # =IFERROR(+D51/D50,0)
                sheet[f'{letter}54'] = f'=IFERROR({letter}51/{letter}50,0)'
            elif row == 58:
                sheet[f'{letter}58'] = f'=IF({letter}57=TRUE, {letter}22, "")'
            elif row == 59:
                sheet[f'{letter}59'] = f'=IF({letter}57=TRUE, {letter}34, 0)'
            elif row == 60:
                # =SUMIFS($G59:$KO59,$G4:$KO4, G4)
                sheet[f'{letter}60'] = f'=SUMIFS($G59:' \
                                       f'${column_letters_7[len(column_letters_7) - 1]}59,$G4:' \
                                       f'${column_letters_7[len(column_letters_7) - 1]}4, {letter}4)'
            elif row == 61:
                sheet[f'{letter}61'] = f'=SUM({letter}50)'
            elif row == 62:
                sheet[f'{letter}62'] = f'=SUM({letter}51-{letter}60)'
            elif row == 63:
                sheet[f'{letter}63'] = f'=SUM({letter}61-{letter}62)'
            elif row == 66:
                # =SUMIFS(G19: KO19, G57: KO57, TRUE)
                sheet[f'B66'] = f'=SUMIFS(G19:' \
                                f'{column_letters_7[len(column_letters_7) - 1]}19,G57:' \
                                f'{column_letters_7[len(column_letters_7) - 1]}57, TRUE)'
            elif row == 67:
                # =SUMIFS(G33: KO33, G57: KO57, TRUE)
                sheet[f'B67'] = f'=SUMIFS(G33:' \
                                f'{column_letters_7[len(column_letters_7) - 1]}33,G57:' \
                                f'{column_letters_7[len(column_letters_7) - 1]}57, TRUE)'

    # format all rows with data except rows 1 to 11, 20 to 23, 28 and 29 as currency with 2 decimal places and
    # comma every 3 digits, bold and white font, and for row 11 as a percentage with 2 decimal places and comma
    # every 3 digits
    for row in sheet.iter_rows(min_row=12, min_col=2, max_row=63, max_col=sheet.max_column):
        for cell in row:
            if cell.row == 11:
                cell.font = Font(bold=True, color='FFFFFF')
                cell.number_format = '0.00%'
            if cell.row == 54:
                cell.number_format = '0.00%'
            elif cell.row == 20 or cell.row == 21 or cell.row == 22 or cell.row == 23 or cell.row == 28 \
                    or cell.row == 29:
                continue
            else:
                cell.font = Font(bold=True, color='FFFFFF')
                cell.number_format = 'R#,##0.00'

    rows_to_center = [5, 6, 7, 9, 10, 11, 13, 14, 15, 16, 17, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31,
                      32, 33, 34, 36, 37, 38, 39, 41, 42, 43, 44, 45, 46, 47, 48, 49, 50, 51, 52, 54, 58, 59, 60, 61,
                      62, 63]
    # Loop through the rows_to_format_currency list and align the cells from column 6 to the last column in the
    # centre
    for row in rows_to_center:
        # COLUMNS TO FILL ETC
        if row in [9, 10, 11, 20, 21, 22, 23, 28, 29]:
            col_number = 7
        else:
            col_number = 2
        # DARKER COLOR

        if row in [9, 10, 23, 13, 19, 27, 29, 32, 33, 34, 38, 39, 43]:
            color_is = '3E54AC'
        elif row in [52, 63]:
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

                    sheet[f'{get_column_letter(cell.column)}11'].number_format = '0.00%'
                    sheet[f'{get_column_letter(cell.column)}11'].alignment = Alignment(horizontal='center')

                if row == 9:
                    if cell.value == 'UnAllocated':
                        color_is = '00B7C2'
                    else:
                        color_is = '3E54AC'

                if row == 10:
                    if cell.value == 'ZZUN01':
                        color_is = '00B7C2'
                    else:
                        color_is = '3E54AC'

                if row == 32:
                    if cell.value != 0:
                        color_is = '00B7C2'
                    else:
                        color_is = '3E54AC'

                # if row == 58:
                #     if cell.value:
                #
                #         color_is = '00B7C2'
                #     else:
                #         color_is = '3E54AC'

                cell.fill = PatternFill(start_color=color_is, end_color=color_is, fill_type='solid')
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.font = Font(bold=True, color='FFFFFF')
                cell.border = Border(left=Side(border_style='medium', color='000000'),
                                     right=Side(border_style='medium', color='000000'),
                                     top=Side(border_style='medium', color='000000'),
                                     bottom=Side(border_style='medium', color='000000'))

    rows_to_hide = [2, 3, 4, 53, 57]
    # Loop through the rows_to_hide list and hide the rows
    for row in rows_to_hide:
        sheet.row_dimensions[row].hidden = True

    true_false = [57]

    for row in true_false:
        for col in sheet.iter_cols(min_row=row, min_col=1, max_row=row, max_col=sheet.max_column):
            for cell in col:
                # If the cell value in row 57 is True then set the color of the cell in row 58 to green
                if cell.value:
                    sheet[f'{get_column_letter(cell.column)}58'].fill = PatternFill(start_color='A555EC',
                                                                                    end_color='A555EC',
                                                                                    fill_type='solid')
                    sheet[f'{get_column_letter(cell.column)}59'].fill = PatternFill(start_color='A555EC',
                                                                                    end_color='A555EC',
                                                                                    fill_type='solid')
                    sheet[f'{get_column_letter(cell.column)}22'].fill = PatternFill(start_color='A555EC',
                                                                                    end_color='A555EC',
                                                                                    fill_type='solid')

    # Set column 1 to 25 units wide and make column 1 bold
    for col in sheet.iter_cols(min_row=1, min_col=1, max_row=sheet.max_row, max_col=5):
        for cell in col:
            sheet.column_dimensions[cell.column_letter].width = 30
            cell.font = Font(bold=True)

    # ROWS TO SUM
    rows_to_sum = [13, 14, 15, 16, 17, 19, 24, 25, 26, 27, 30, 31, 32, 33, 34, 36, 37, 38, 39, 41, 42, 43, 44, 45,
                   46, 47, 48, 49, 50, 51, 52, 54, 59, 60, 61, 62, 63]
    # for each row in rows_to_sum, sum the cells from column 6 to the last column in the sheet in insert the
    # formula in column 'B', with white font and bold and format the cell as currency with 2 decimal places and
    # comma every 3 digits

    sheet['B3'] = f'=COUNTIF({get_column_letter(7)}10:{get_column_letter(sheet.max_column)}10,"<>ZZUN01")'
    sheet['D3'] = f'=COUNTIF({get_column_letter(7)}3:{get_column_letter(sheet.max_column)}3,TRUE)'
    sheet['A54'] = f'LTV'

    columns_to_format = ['B', 'C', 'D', 'E', 'F']

    for row in rows_to_sum:
        sheet[f'B{row}'] = f'=SUM({get_column_letter(7)}{row}:{get_column_letter(sheet.max_column)}{row})'
        for column in columns_to_format:
            sheet[f'{column}{row}'].font = Font(bold=True, color='FFFFFF')
            sheet[f'{column}{row}'].number_format = 'R#,##0.00'
            sheet[f'{column}{row}'].alignment = Alignment(horizontal='center')

            if column == 'D':
                if row == 51:
                    sheet[
                        f'{column}{row}'] = f'=+D34'
                elif row == 52:
                    sheet[
                        f'{column}{row}'] = f'=+D50-D51'
                elif row == 54:
                    sheet[
                        f'{column}{row}'] = f'=IFERROR(+D51/D50,0)'
                    sheet[f'{column}{row}'].number_format = '0.00%'
                    sheet[f'{column}{row}'].font = Font(bold=True, color='000000')

                else:
                    sheet[
                        f'{column}{row}'] = f'=SUMIFS($G{row}:${get_column_letter(sheet.max_column)}{row},' \
                                            f'$G3:${get_column_letter(sheet.max_column)}3,' \
                                            f'TRUE)'

            if column == 'E':
                if row == 51:
                    sheet[
                        f'{column}{row}'] = f'=+E34'
                elif row == 52:
                    sheet[
                        f'{column}{row}'] = f'=+E50-E51'
                elif row == 54:
                    # =IFERROR(+D51 / D50, 0)
                    sheet[
                        f'{column}{row}'] = f'=IFERROR(+E51/E50,0)'
                    sheet[f'{column}{row}'].number_format = '0.00%'
                    sheet[f'{column}{row}'].font = Font(bold=True, color='000000')
                else:
                    sheet[
                        f'{column}{row}'] = f'=SUMIFS($G{row}:${get_column_letter(sheet.max_column)}{row},' \
                                            f'$G2:${get_column_letter(sheet.max_column)}2,' \
                                            f'TRUE)-SUMIFS($G{row}:${get_column_letter(sheet.max_column)}{row},' \
                                            f'$G3:${get_column_letter(sheet.max_column)}3,' \
                                            f'TRUE)'

            if column == 'F':
                sheet[
                    f'{column}{row}'] = f'=B{row}-D{row}-E{row}'

    # Join rows_to_add_formulas,rows_to_format_currency,rows_to_center and rows_to_sum as one new list ordered
    # and only unique values all_rows_to_format = list(set(rows_to_add_formulas + rows_to_format_currency +
    # rows_to_center + rows_to_sum)) all_rows_to_format.sort()

    # If the values in row 16 from column 2 are greater than 0, then the cells in row 16 from column 2 must be
    # filled with red and the font must be white
    sheet['B16'].fill = PatternFill(start_color='DF7857', end_color='DF7857', fill_type='solid')
    for col in sheet.iter_cols(min_row=16, min_col=7, max_row=16, max_col=sheet.max_column):
        for cell in col:
            if type(cell.value) == float:
                if cell.value > 0:
                    for row in range(16, 17):
                        cell.fill = PatternFill(start_color='DF7857', end_color='DF7857', fill_type='solid')
                        cell.font = Font(color='FFFFFF')

    # from column 7 to the last column in the sheet, if the value in the cell in row 2 is not equal to  0,
    # then the corresponding cell in row 19 must have a red fill and the font white
    for col in sheet.iter_cols(min_row=53, min_col=7, max_row=53, max_col=sheet.max_column):
        for cell in col:
            if cell.value != 0:
                for row in range(19, 20):
                    sheet[f'{get_column_letter(cell.column)}{row}'].fill = PatternFill(start_color='DF7857',
                                                                                       end_color='DF7857',
                                                                                       fill_type='solid')
                    sheet[f'{get_column_letter(cell.column)}{row}'].font = Font(color='FFFFFF')

    sheet['D5'].fill = PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid')

    # Make cell E5 fill with dark-green and the font white
    sheet['E5'].fill = PatternFill(start_color='539165', end_color='539165', fill_type='solid')

    # loop through cells in row 5, row 6 and row 7 from column 7. If the value in the cells in row 7 equals 'Yes'
    # then fill cells in row 5, 6 and 7 with red and the font white
    for col in sheet.iter_cols(min_row=7, min_col=7, max_row=7, max_col=sheet.max_column):
        for cell in col:
            # if cell.value == 'Yes' then make cells in row 5, 6 and 7 fill with red and the font white
            if cell.value == 'Yes':
                for row in range(5, 8):
                    sheet[f'{cell.column_letter}{row}'].fill = PatternFill(start_color='FF0000',
                                                                           end_color='FF0000',
                                                                           fill_type='solid')
                    sheet[f'{cell.column_letter}{row}'].font = Font(color='FFFFFF')

    # loop through cells in row 6 and row 7 from column 7. If the value in the cells in row 7 equals 'No' and the
    # cells in row 6 equals 'Yes' then fill cells in row 6 and 7 with dark-green and the font white
    for col in sheet.iter_cols(min_row=6, min_col=7, max_row=6, max_col=sheet.max_column):
        for cell in col:
            # if cell.value == 'Yes' then make cells in row 5, 6 and 7 fill with red and the font white
            if cell.value == 'Yes':
                for row in range(7, 8):
                    if sheet[f'{cell.column_letter}{row}'].value == 'No':
                        for row1 in range(5, 8):
                            sheet[f'{cell.column_letter}{row1}'].fill = PatternFill(start_color='539165',
                                                                                    end_color='539165',
                                                                                    fill_type='solid')
                            sheet[f'{cell.column_letter}{row1}'].font = Font(color='FFFFFF')

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
        if 0 < index < len(merge_master) - 1:
            if item['value'] != merge_master[index + 1]['value']:
                merge_end.append(item['column'])

    merge_end.append(merge_master[len(merge_master) - 1]['column'])

    # Create a dictionary to store the start and end columns for each row
    rows_to_merge = [5, 6, 7, 13, 14, 15, 16, 17, 41, 42, 43, 44, 45, 46, 47, 48, 49, 50, 51, 52, 54, 60, 61, 62, 63]
    merge_dict = {}

    # Loop through the rows_to_merge list and populate merge_dict with start and end columns for each row
    for row in rows_to_merge:
        merge_dict[row] = []
        start_column = None
        end_column = None
        for cell in sheet[row]:
            if cell.column in merge_start:
                start_column = cell.column
            elif cell.column in merge_end:
                end_column = cell.column
            if start_column and end_column:
                merge_dict[row].append((start_column, end_column))
                start_column = None
                end_column = None

    # Merge cells for each row in bulk using the merge_dict
    for row, merge_cols in merge_dict.items():
        for start, end in merge_cols:
            sheet.merge_cells(start_row=row, start_column=start, end_row=row, end_column=end)

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

    # Freeze frames at C6
    sheet.freeze_panes = 'C6'
