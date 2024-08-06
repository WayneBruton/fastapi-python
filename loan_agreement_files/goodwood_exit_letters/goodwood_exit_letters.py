from docxtpl import DocxTemplate
import os


def create_goodwood_exit_letters(data):
    # if loan_agreements/exit_letter_nsst.docx exists, delete it
    if os.path.exists("loan_agreements/exit_letter_nsst.docx"):
        os.remove("loan_agreements/exit_letter_nsst.docx")
        print("File deleted")
    if os.path.exists("loan_agreements/exit_letter_omh.docx"):
        os.remove("loan_agreements/exit_letter_omh.docx")
        print("File deleted")
    print("Data in Letter File:::",data)
    months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October","November", "December"]

    investor_name = data.get('investor_name', "")
    investor_name_nsst = data.get('investor_name', "")
    investor_name_nsst = investor_name_nsst.split(' ')

    # take the foirst character only of the first record and combine with the second record
    investor_name_nsst = investor_name_nsst[0][0] + ' ' + investor_name_nsst[1]
    print("Investor Name:::", investor_name_nsst)


    investor_acc_number = data.get('investor_acc_number', "")
    opp_code = data.get('opportunity_code', "")
    exit_date = data.get('exit_date', "")
    exit_date = exit_date.split('/')
    exit_date = f"{exit_date[2]} {months[int(exit_date[1])-1]} {exit_date[0]}"
    exit_value = data.get('exit_value', "")
    exit_value = exit_value.replace('R', '')
    exit_value = exit_value.replace(' ', '')
    exit_value = float(exit_value)
    exit_value = f"R {exit_value:,.2f}"
    # exit_value = f"R {data.get('exit_value', '')}"
    rollover_amount = float(data.get('rollover_amount', 0))
    exit_amount = float(data.get('exit_amount',0))
    exit_amount = f"R {exit_amount:,.2f}"
    rollover_amount = (f"R {rollover_amount:,.2f}")
    doc = DocxTemplate("loan_agreement_files/goodwood_exit_letters/NSST_Exit - Goodwood.docx")
    doc2 = DocxTemplate("loan_agreement_files/goodwood_exit_letters/OMH_Exit - Goodwood.docx")
    context = {
        'investor_name': investor_name,
        'investor_acc_number': investor_acc_number,
        'opp_code': opp_code,
        'exit_date': exit_date,
        'exit_value': exit_value,
        'exit_amount': exit_amount,
        'rollover_amount': rollover_amount,
        "investor_name_nsst": investor_name_nsst
    }
    doc.render(context)
    doc2.render(context)
    # context = {
    #     'erf_number': erf_number,
    #     'borrower': borrower,
    #     'trading_name': trading_name,
    #     'company_registration_number': company_registration_number,
    #     'vat_number': vat_number,
    #     'members_directors': members_directors,
    #     'investor': investor,
    #     'bank_account_holder': bank_account_holder,
    #     'bank_name': bank_name,
    #     'bank_account_type': bank_account_type,
    #     'bank_branch': bank_branch,
    #     'bank_account_number': bank_account_number,
    #     'bank_branch_code': bank_branch_code,
    #     'investor_physical_street': investor_physical_street,
    #     'investor_physical_suburb': investor_physical_suburb,
    #     'investor_physical_province': investor_physical_province,
    #     'investor_physical_city': investor_physical_city,
    #     'investor_physical_postal_code': investor_physical_postal_code,
    #     'investor_physical_country': investor_physical_country,
    #     'investor_postal_street_box': investor_postal_street_box,
    #     'investor_postal_suburb': investor_postal_suburb,
    #     'investor_postal_province': investor_postal_province,
    #     'investor_postal_city': investor_postal_city,
    #     'investor_postal_postal_code': investor_postal_postal_code,
    #     'investor_postal_country': investor_postal_country,
    #     'income_tax_number': income_tax_number,
    #     'investor_landline': investor_landline,
    #     'investor_mobile': investor_mobile,
    #     'investor_email': investor_email,
    #     'interest_rate': interest_rate,
    #     'investment_amount': investment_amount,
    #     'investor_name': investor_name,
    #     'investor_id_number': investor_id_number,
    #
    # }
    # doc.render(context)
    file_name = f"loan_agreements/exit_letter_nsst.docx"
    file_name2 = f"loan_agreements/exit_letter_omh.docx"
    doc.save(file_name)
    doc2.save(file_name2)
    # return {'link': file_name}
    return {'link': 'loan_agreements/exit_letter_nsst.docx',
            'link2': 'loan_agreements/exit_letter_omh.docx'}