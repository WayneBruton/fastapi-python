from PyPDF2 import PdfMerger


# from loan_agreement_files.loan_agreement import linked_unit, lender, project


def merge_files(pdf_list, lender, project, linked_unit):
    merger = PdfMerger()
    for pdf in pdf_list:
        merger.append(pdf)
    merger.write(f"loan_agreements/Loan_Agreement_{lender}_{project}_{linked_unit}.pdf")

    merger.close()
    return {"link": f"Loan_Agreement_{lender}_{project}_{linked_unit}.pdf"}
