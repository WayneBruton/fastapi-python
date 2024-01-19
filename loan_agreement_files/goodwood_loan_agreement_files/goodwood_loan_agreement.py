from docxtpl import DocxTemplate


def create_goodwood_la(data):
    # for item in data:
    #     print(item, ":", data[item])
    erf_number = data['pledges']['opportunity_code']
    # make erf number = the right 4 characters preceded by 'Erf '
    erf_number = erf_number[-4:]
    if data.get('registered_company_name', "") != '':
        borrower = data.get('registered_company_name', "")
    else:
        if data.get('investor_name2', "") == '':
            borrower = data.get('investor_name', "") + ' ' + data.get('investor_surname', "")
        else:
            borrower = data.get('investor_name', "") + ' ' + data.get('investor_surname', "") + ' and ' + data.get(
                'investor_name2', "") + ' ' + \
                       data.get('investor_surname2', "")

    if data.get('trading_name', '') != data.get('registered_company_name', ""):
        trading_name = data.get('trading_name', "")
    else:
        trading_name = ''
    company_registration_number = data.get('registration_number', "")
    vat_number = data.get('vat_number', "")
    if data.get('members_directors', "") != '' and data.get('members_directors2', "") == '':
        members_directors = data.get('members_directors', "") + '(' + data.get('members_directors_id', "") + ')'
    elif data.get('members_directors', "") == '' and data.get('members_directors2', "") == '':
        members_directors = ''
    elif data.get('members_directors', "") != '' and data.get('members_directors2', "") != '':
        members_directors = data.get('members_directors', "") + ' (' + data.get('members_directors_id',
                                                                                "") + ')' + ' and ' + data.get(
            'members_directors2', "") + ' (' + data.get('members_directors_id2', "") + ')'
    if data.get('investor_name2', "") == '' and data.get('registered_company_name', "") == '':
        investor = data.get('investor_name', "") + ' ' + data.get('investor_surname', "") + '(' + data.get(
            'investor_id', "") + ')'
    elif data.get('investor_name2', "") != '' and data.get('registered_company_name', "") == '':
        investor = data.get('investor_name', "") + ' ' + data.get('investor_surname', "") + '(' + data.get(
            'investor_id_number', "") + ')' + ' and ' + \
                   data.get('investor_name2', "") + ' ' + data.get('investor_surname2', "") + '(' + data.get(
            'investor_id2', "") + ')'
    else:
        investor = data.get('investor_name', "") + ' ' + data.get('investor_surname', "") + '(' + data.get(
            'investor_id_number', "") + ')'
    bank_account_holder = data.get('bank_account_holder', "")
    investor_id_number = data.get('investor_id_number', "")
    bank_name = data.get('bank_name', "")
    bank_account_type = data.get('bank_account_type', "")
    bank_branch = data.get('bank_branch', "")
    bank_account_number = data.get('bank_account_number', "")
    bank_branch_code = data.get('bank_branch_code', "")
    investor_physical_street = data.get('investor_physical_street', "")
    investor_physical_suburb = data.get('investor_physical_suburb', "")
    investor_physical_province = data.get('investor_physical_province', "")
    investor_physical_city = data.get('investor_physical_city', "")
    investor_physical_postal_code = data.get('investor_physical_postal_code', "")
    investor_physical_country = data.get('investor_physical_country', "")
    if data.get('investor_postal_street_box', "") == '':
        investor_postal_street_box = investor_physical_street
    else:
        investor_postal_street_box = data.get('investor_postal_street_box', "")
    if data.get('investor_postal_suburb', "") == '':
        investor_postal_suburb = investor_physical_suburb
    else:
        investor_postal_suburb = data.get('investor_postal_suburb', "")
    if data.get('investor_postal_province', "") == '':
        investor_postal_province = investor_physical_province
    else:
        investor_postal_province = data.get('investor_postal_province', "")
    if data.get('investor_postal_city', "") == '':
        investor_postal_city = investor_physical_city
    else:
        investor_postal_city = data.get('investor_postal_city', "")
    if data.get('investor_postal_postal_code', "") == '':
        investor_postal_postal_code = investor_physical_postal_code
    else:
        investor_postal_postal_code = data.get('investor_postal_postal_code', "")
    if data.get('investor_postal_country', "") == '':
        investor_postal_country = investor_physical_country
    else:
        investor_postal_country = data.get('investor_postal_country', "")
    income_tax_number = data.get('income_tax_number', "")
    investor_landline = data.get('investor_landline', "")
    investor_mobile = data.get('investor_mobile', "")
    investor_email = data.get('investor_email', "")
    interest_rate = data['pledges']['investment_interest_rate'] + ' %'
    investment_amount = float(data['pledges']['investment_amount'])
    investment_amount = "R {:,.2f}".format(investment_amount)
    investor_name = data.get('investor_name', "") + ' ' + data.get('investor_surname', "")

    doc = DocxTemplate("loan_agreement_files/goodwood_loan_agreement_files/Investment_Loan_Agreement_GW.docx")
    context = {
        'erf_number': erf_number,
        'borrower': borrower,
        'trading_name': trading_name,
        'company_registration_number': company_registration_number,
        'vat_number': vat_number,
        'members_directors': members_directors,
        'investor': investor,
        'bank_account_holder': bank_account_holder,
        'bank_name': bank_name,
        'bank_account_type': bank_account_type,
        'bank_branch': bank_branch,
        'bank_account_number': bank_account_number,
        'bank_branch_code': bank_branch_code,
        'investor_physical_street': investor_physical_street,
        'investor_physical_suburb': investor_physical_suburb,
        'investor_physical_province': investor_physical_province,
        'investor_physical_city': investor_physical_city,
        'investor_physical_postal_code': investor_physical_postal_code,
        'investor_physical_country': investor_physical_country,
        'investor_postal_street_box': investor_postal_street_box,
        'investor_postal_suburb': investor_postal_suburb,
        'investor_postal_province': investor_postal_province,
        'investor_postal_city': investor_postal_city,
        'investor_postal_postal_code': investor_postal_postal_code,
        'investor_postal_country': investor_postal_country,
        'income_tax_number': income_tax_number,
        'investor_landline': investor_landline,
        'investor_mobile': investor_mobile,
        'investor_email': investor_email,
        'interest_rate': interest_rate,
        'investment_amount': investment_amount,
        'investor_name': investor_name,
        'investor_id_number': investor_id_number,

    }
    doc.render(context)
    file_name = f"loan_agreements/{investor_name}-GW{erf_number}.docx"
    doc.save(file_name)
    return {'link': file_name}
