import os
# from operator import add, sub, mul, truediv
from zipfile import ZipFile
from typing import Union
from io import BytesIO, FileIO, BufferedReader, BufferedWriter
from pathlib import Path
from PyPDF2 import PdfFileReader, PdfFileWriter
from loan_agreement_files.investment_cover_letter import print_investor_cover_letter
from loan_agreement_files.loan_agreement import print_investor_loan_agreement
from loan_agreement_files.annexure_c import print_annexure_c


# HELPER FUNCTION TO CREATE A ZIP FILE - THIS IS USED IN THE MAIN FUNCTION BELOW:
def create_zip_file(file_list, linked_unit, investor):
    investor = investor.split(" ")
    investor = [z for z in investor if z != '']
    investor = '_'.join(investor)
    with ZipFile(f'loan_agreements/{investor}-{linked_unit}.zip', 'w') as zipObj2:
        for item in file_list:
            zipObj2.write(item)
    return {"link": f'loan_agreements/{investor}-{linked_unit}.zip'}


# CREATE THE FINAL LOAN AGREEMENT FILES AS A ZIP
def create_final_loan_agreement(linked_unit, investor, investor2, nsst, project, investment_amount,
                                investment_interest_rate,
                                investor_id, investor_id2, registered_company_name, registration_number):
    # GENERATE COVER LETTER
    cover_letter = print_investor_cover_letter(investor, investor2)

    # GENERATE LOAN AGREEMENT
    loan_agreement = print_investor_loan_agreement(investor, investor2, nsst, project, linked_unit, investment_amount,
                                                   investment_interest_rate, investor_id, investor_id2,
                                                   registered_company_name, registration_number)

    # GENERATE ANNEXURE C
    annexure_C = print_annexure_c(investor, nsst)

    # SPLIT LOAN AGREEMENT INTO SEPARATE PAGES
    pdf_file_path = 'LoanAgreement.pdf'
    file_base_name = pdf_file_path.replace('.pdf', '')
    output_folder_path = os.path.join(os.getcwd(), 'split_pdf_files')
    pdf = PdfFileReader(pdf_file_path)

    for page_number in range(pdf.numPages):
        pdf_writer = PdfFileWriter()
        pdf_writer.add_page(pdf.getPage(page_number))

        with open(os.path.join(output_folder_path, '{0}_page{1}.pdf'.format(file_base_name, page_number + 1)),
                  'wb', buffering=0) as f:
            f: Union[Path, str, BytesIO, BufferedReader, BufferedWriter, FileIO] = f
            pdf_writer.write(f)
            f.close()

    # FILE LIST OF DOCUMENTS TO ZIP
    file_list = [cover_letter]

    split_files = os.listdir('split_pdf_files')
    for doc in split_files:
        insert_doc = f'split_pdf_files/{doc}'
        file_list.append(insert_doc)

    # APPEND STANDARD ANNEXURES TO BE ZIPPED LATER
    file_list.append("annexures/Heron View - Annexure A.pdf")
    file_list.append("annexures/Annexure B - Ground Floor.pdf")
    file_list.append("annexures/Annexure B - First Floor.pdf")
    file_list.append("annexures/Annexure B - Second Floor.pdf")
    file_list.append(annexure_C)

    # ZIP ALL FILES INTO ONE FILE - THIS USES THE HELPER FUNCTION ABOVE
    final_doc = create_zip_file(file_list, linked_unit, investor)

    # DELETE ALL SPLIT FILES AS NO LONGER NEEDED
    for doc in split_files:
        os.remove(f'split_pdf_files/{doc}')

    # DELETE OTHER WASTED FILES CREATED EARLIER TO KEEP WASTAGE TO A MINIMUM
    os.remove(cover_letter)
    os.remove(loan_agreement)
    os.remove(annexure_C)

    # RETURN THE FINAL ZIP FILE SO IT CAN BE ACCESSED BY USERS
    return final_doc


# MY SANDBOX OR PLAY AREA

# x = 5000000.00
#
# name = 'wayne'
# print(f"R {x:,.2f}")
#
# print(f"Congratulations {name.title()}, you have won ${x:,.2f}")

# def calculator(operator, a, b):
#     ops = {
#         "+": add,
#         "-": sub,
#         "*": mul,
#         "/": truediv,
#     }
#     if operator in ops:
#         return ops[operator](a, b)
#     else:
#         return "Invalid operator"

# print(calculator('-', 3, 15))
