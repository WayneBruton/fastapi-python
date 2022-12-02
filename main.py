import os
from zipfile import ZipFile
from PyPDF2 import PdfFileReader, PdfFileWriter
from loan_agreement_files.investment_cover_letter import print_investor_cover_letter
from loan_agreement_files.loan_agreement import print_investor_loan_agreement
from loan_agreement_files.annexure_c import print_annexure_c


def create_zip_file(file_list, linked_unit, investor):
    investor = investor.split(" ")
    investor = [x for x in investor if x != '']
    investor = '_'.join(investor)
    with ZipFile(f'loan_agreements/{investor}-{linked_unit}.zip', 'w') as zipObj2:
        for item in file_list:
            zipObj2.write(item)
    return {"link": f'loan_agreements/{investor}-{linked_unit}.zip'}


def create_final_loan_agreement(linked_unit, investor, nsst, project, investment_amount, investment_interest_rate,
                                investor_id):
    pdf_list = []
    linked_unit_split = []

    cover_letter = print_investor_cover_letter(investor)
    pdf_list.append(cover_letter)
    loan_agreement = print_investor_loan_agreement(investor, nsst, project, linked_unit, investment_amount,
                                                   investment_interest_rate, investor_id)
    pdf_list.append(loan_agreement)
    pdf_list.append("annexures/Heron View - Annexure A.pdf")

    for letter in linked_unit:
        linked_unit_split.append(letter)

    if linked_unit_split[-3] == 1:
        pdf_list.append("annexures/Annexure B - Ground Floor.pdf")
    elif linked_unit_split[-3] == 2:
        pdf_list.append("annexures/Annexure B - First Floor.pdf")
    else:
        pdf_list.append("annexures/Annexure B - Second Floor.pdf")

    annexure_C = print_annexure_c(investor, nsst)
    pdf_list.append(annexure_C)

    # SPLIT LOAN AGREEMENT INTO SEPERATE PAGES
    pdf_file_path = 'LoanAgreement.pdf'
    file_base_name = pdf_file_path.replace('.pdf', '')
    output_folder_path = os.path.join(os.getcwd(), 'split_pdf_files')
    pdf = PdfFileReader(pdf_file_path)

    for page_number in range(pdf.numPages):
        pdf_writer = PdfFileWriter()
        pdf_writer.add_page(pdf.getPage(page_number))

        with open(os.path.join(output_folder_path, '{0}_page{1}.pdf'.format(file_base_name, page_number + 1)),
                  'wb') as f:
            pdf_writer.write(f)
            f.close()

    # FILE LIST OF DOCUMENTS TO ZIP
    file_list = [cover_letter]

    split_files = os.listdir('split_pdf_files')
    for doc in split_files:
        insert_doc = f'split_pdf_files/{doc}'
        file_list.append(insert_doc)
    print("split_files", file_list)
    file_list.append("annexures/Heron View - Annexure A.pdf")
    file_list.append("annexures/Annexure B - Ground Floor.pdf")
    file_list.append("annexures/Annexure B - First Floor.pdf")
    file_list.append("annexures/Annexure B - Second Floor.pdf")
    file_list.append(annexure_C)

    # ZIP ALL FILES INTO ONE FILE
    final_doc = create_zip_file(file_list, linked_unit, investor)

    for doc in split_files:
        os.remove(f'split_pdf_files/{doc}')

    os.remove(cover_letter)
    os.remove(loan_agreement)
    os.remove(annexure_C)

    return final_doc
