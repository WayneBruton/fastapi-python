import os
from loan_agreement_files.investment_cover_letter import print_investor_cover_letter
from loan_agreement_files.loan_agreement import print_investor_loan_agreement
from loan_agreement_files.annexure_c import print_annexure_c
from loan_agreement_files.merge_documents import merge_files


def create_final_loan_agreement(linked_unit, investor, nsst, project, investment_amount, investment_interest_rate):
    pdf_list = []
    linked_unit_split = []

    cover_letter = print_investor_cover_letter(investor)
    pdf_list.append(cover_letter)
    loan_agreement = print_investor_loan_agreement(investor, nsst, project, linked_unit, investment_amount,
                                                   investment_interest_rate)
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

    final_doc = merge_files(pdf_list, investor, project, linked_unit)

    os.remove(cover_letter)
    os.remove(loan_agreement)
    os.remove(annexure_C)

    return final_doc
