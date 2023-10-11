import os

import openpyxl


def insert_data_from_xero_profit_loss(base_data):
    # get all excel file names from the folder cashflow_p&l_files and insert filenames into a list called files
    files = []
    for filename in os.listdir('cashflow_p&l_files'):
        files.append(filename)
    # print("files",files)
    # filter the list of files and exclude files that begin with ~$
    files = [file for file in files if not file.startswith('~$')]
    # filter out  any files that do not end with .xlsx
    files = [file for file in files if file.endswith('.xlsx')]
    # exclude "2. Heron Master report.xlsx" from the list of files
    # files = [file for file in files if not file.startswith('2. Heron Master report.xlsx')]
    print("files",files)
    try:
        for file in files:
            # Load the existing Excel workbook
            workbook = openpyxl.load_workbook(f'cashflow_p&l_files/{file}')

            # Select the desired sheet by name
            sheet = workbook['CPC 24']

            # Insert data into the sheet
            for entry in base_data:
                account_name = entry['Account']
                amounts = entry['Amount']

                # Find the row in the sheet where the account matches
                for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row, min_col=1, max_col=1):
                    for cell in row:
                        if cell.value == account_name:
                            row_idx = cell.row

                            # Insert amount values in columns B to M
                            for col_idx, amount in enumerate(amounts, start=2):
                                sheet.cell(row=row_idx, column=col_idx, value=amount)

            # Save the modified workbook
            workbook.save(f'cashflow_p&l_files/{file}')
        return {"Success": True}
    except Exception as e:
        print("Error", e)
        return {"Success": False, "Error": str(e)}
