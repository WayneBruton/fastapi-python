from docxtpl import DocxTemplate


def create_goodwood_la(data):
    # for item in data:
    #     print(item, ":", data[item])
    erf_number = data['pledges']['opportunity_code']
    # make erf number = the right 4 characters preceded by 'Erf '
    erf_number = erf_number[-4:]
    if data['registered_company_name'] != '':
        borrower = data['registered_company_name']
    else:
        if data['investor_name2'] == '':
            borrower = data['investor_name'] + ' ' + data['investor_surname']
        else:
            borrower = data['investor_name'] + ' ' + data['investor_surname'] + ' and ' + data['investor_name2'] + ' ' + \
                       data['investor_surname2']

    if data['trading_name'] != data['registered_company_name']:
        trading_name = data['trading_name']
    else:
        trading_name = ''
    company_registration_number = data['registration_number']
    vat_number = data['vat_number']
    if data['members_directors'] != '' and data['members_directors2'] == '':
        members_directors = data['members_directors'] + '(' + data['members_directors_id'] + ')'
    elif data['members_directors'] == '' and data['members_directors2'] == '':
        members_directors = ''
    elif data['members_directors'] != '' and data['members_directors2'] != '':
        members_directors = data['members_directors'] + ' (' + data['members_directors_id'] + ')' + ' and ' + data[
            'members_directors2'] + ' (' + data['members_directors_id2'] + ')'
    if data['investor_name2'] == '' and data['registered_company_name'] == '':
        investor = data['investor_name'] + ' ' + data['investor_surname'] + '(' + data['investor_id'] + ')'
    elif data['investor_name2'] != '' and data['registered_company_name'] == '':
        investor = data['investor_name'] + ' ' + data['investor_surname'] + '(' + data['investor_id'] + ')' + ' and ' + \
                   data['investor_name2'] + ' ' + data['investor_surname2'] + '(' + data['investor_id2'] + ')'
    else:
        investor = data['investor_name'] + ' ' + data['investor_surname'] + '(' + data['investor_id_number'] + ')'
    bank_account_holder = data['bank_account_holder']
    bank_name = data['bank_name']
    bank_account_type = data['bank_account_type']
    bank_branch = data['bank_branch']
    bank_account_number = data['bank_account_number']
    bank_branch_code = data['bank_branch_code']
    investor_physical_street = data['investor_physical_street']
    investor_physical_suburb = data['investor_physical_suburb']
    investor_physical_province = data['investor_physical_province']
    investor_physical_city = data['investor_physical_city']
    investor_physical_postal_code = data['investor_physical_postal_code']
    investor_physical_country = data['investor_physical_country']
    if data['investor_postal_street_box'] == '':
        investor_postal_street_box = investor_physical_street
    else:
        investor_postal_street_box = data['investor_postal_street_box']
    if data['investor_postal_suburb'] == '':
        investor_postal_suburb = investor_physical_suburb
    else:
        investor_postal_suburb = data['investor_postal_suburb']
    if data['investor_postal_province'] == '':
        investor_postal_province = investor_physical_province
    else:
        investor_postal_province = data['investor_postal_province']
    if data['investor_postal_city'] == '':
        investor_postal_city = investor_physical_city
    else:
        investor_postal_city = data['investor_postal_city']
    if data['investor_postal_postal_code'] == '':
        investor_postal_postal_code = investor_physical_postal_code
    else:
        investor_postal_postal_code = data['investor_postal_postal_code']
    if data['investor_postal_country'] == '':
        investor_postal_country = investor_physical_country
    else:
        investor_postal_country = data['investor_postal_country']
    income_tax_number = data['income_tax_number']
    investor_landline = data['investor_landline']
    investor_mobile = data['investor_mobile']
    investor_email = data['investor_email']
    interest_rate = data['pledges']['investment_interest_rate'] + ' %'
    investment_amount = float(data['pledges']['investment_amount'])
    investment_amount = "R {:,.2f}".format(investment_amount)
    investor_name = data['investor_name'] + ' ' + data['investor_surname']

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

    }
    doc.render(context)
    file_name = f"loan_agreements/{investor_name}-GW{erf_number}.docx"
    doc.save(file_name)
    return {'link': file_name}
