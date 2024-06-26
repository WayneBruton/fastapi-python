import json

from fastapi import APIRouter, Request
from fastapi.responses import FileResponse

import os
from pydantic import BaseModel
import loan_agreement_files.lender as l1
from main import create_final_loan_agreement
import smtplib
from email.message import EmailMessage
from bson.objectid import ObjectId

from early_releases_excel_generation.early_releases_excel import early_release_creation
from loan_agreement_files.goodwood_loan_agreement_files.goodwood_loan_agreement import create_goodwood_la
from loan_agreement_files.NGAH_loan_agreement_files.ngah_loan_agreement import create_ngah_la
from configuration.db import db


# from verify_token import verify_jwt_token


class InvestorAccNumber(BaseModel):
    investor_acc_number: str
    opportunity_code: str
    token_received: str


class InvestorFile(BaseModel):
    file_name: str


investor = APIRouter()

# MONGO COLLECTIONS
investors = db.investors
rates = db.rates
opportunities = db.opportunities
sales_parameters = db.salesParameters
rollovers = db.investorRollovers


@investor.get("/")
async def get_all_investors():
    result = "AWESOME IT WORKS!!"
    return result


@investor.post("/email_investor_re_release")
async def email_investor_re_release(data: Request):
    request = await data.json()

    # convert the request['investment_amount'] to currency format and assign it to a variable called amount
    amount = "{:,.2f}".format(float(request['investment_amount']))

    # send email to the investor
    smtp_server = "depro8.fcomet.com"
    port = 465  # For starttls
    sender_email = 'omh-app@opportunitymanagement.co.za'
    password = "12071994Wb!"

    message = f"""\
            <html>
              <body>
                <p>Dear {request['name']},<br><br>
                
                We are pleased to inform you that your investment of R {amount} in {request['Category']} 
                ({request['opportunity_code']}) has been released into the project on the following date:
                {request['release_date']}.<br><br>
                
                As a result, the interest on your investment will be calculated at the higher rate.
                   

                      <br><br>
                      <u>Please do not reply to this email as it is not monitored</u>.<br><br>
                      <br /><br />
                      Regards<br />
                      <strong>The OPC Team</strong><br />
                </p>
              </body>
            </html>
            """

    msg = EmailMessage()
    msg['Subject'] = f"Funds Released - {request['Category']} ({request['opportunity_code']})"
    msg['From'] = sender_email
    msg['To'] = f"{request['email']}"
    # msg['To'] = f"{request['email']}, debbie@opportunity.co.za, nick@opportunity.co.za"
    # copy in nick@opportunity.co.za, wynand@capeprojects.co.za, debbie@opportunity.co.za and
    # dirk@cpconstruction.co.za to the email
    # msg['Cc'] = "nick@opportunity.co.za, wynand@capeprojects.co.za, debbie@opportunity.co.za, " \
    #             "dirk@cpconstruction.co.za"

    msg.set_content(message, subtype='html')

    try:
        with smtplib.SMTP_SSL(smtp_server, port) as server:
            server.ehlo()
            server.login(sender_email, password=password)
            server.send_message(msg)
            server.quit()
            return {"message": "Email sent successfully"}
    except Exception as e:
        print(e)
        return {"message": "Email failed to send"}
    return request


# GET LOAN AGREEMENT INFO FROM MONGO & CREATE PHYSICAL PDF DOCUMENT ACCORDINGLY
@investor.post("/investorloanagreement")
async def get_investor_for_loan_agreement(investor_acc_number: InvestorAccNumber):
    # print("investor_acc_number",investor_acc_number)
    # print()
    # token_verification = verify_jwt_token(investor_acc_number.token_received)
    token_verification = "ABC"

    try:

        if token_verification == "Verification Failed":
            return {"error": "User Not Verified"}
        else:

            result_loan = list(investors.aggregate(
                [{"$match": {"investor_acc_number": investor_acc_number.investor_acc_number}}, {"$unwind": "$pledges"},
                 {"$match": {"pledges.opportunity_code": investor_acc_number.opportunity_code}}, {
                     "$project": {"investor_name": 1, "investor_surname": 1, "investor_id_number": 1,
                                  "investor_email": 1,
                                  "investor_name2": 1, "investor_surname2": 1, "investor_id_number2": 1,
                                  "investor_mobile": 1, "investor_landline": 1, "investor_physical_street": 1,
                                  "investor_physical_suburb": 1, "investor_physical_city": 1,
                                  "investor_physical_province": 1,
                                  "investor_physical_country": 1, "investor_physical_postal_code": 1,
                                  "investor_postal_street_box": 1, "investor_postal_suburb": 1,
                                  "investor_postal_city": 1,
                                  "investor_postal_province": 1, "investor_postal_country": 1,
                                  "investor_postal_postal_code": 1,
                                  "registered_company_name": 1, "bank_account_holder": 1, "bank_account_type": 1,
                                  "bank_branch": 1, "bank_account_number": 1, "bank_branch_code": 1, "bank_name": 1,
                                  "income_tax_number": 1,
                                  "members_directors": 1, "members_directors2": 1, "members_directors_id": 1,
                                  "members_directors_id2": 1, "registration_number": 1,
                                  "telefax_number": 1, "trading_name": 1, "vat_number": 1, "pledges": 1,
                                  "id": {'$toString': "$_id"}, "_id": 0, }}]))

            if len(result_loan) == 0:
                return {
                    "error": f"No data found for {investor_acc_number.investor_acc_number} and "
                             f"{investor_acc_number.opportunity_code}"}
            else:

                if result_loan[0]['pledges']['Category'] == 'Goodwood':
                    # print("result_loan", result_loan)
                    # print()
                    print("result_loan", result_loan[0]['pledges']['Category'])
                    # print()
                    # print("Do some shit here!!!")
                    # print()
                    # print("final_doc {'link': 'loan_agreements/Wayne_Bruton_A-GW3616.zip'}")
                    final_doc = create_goodwood_la(result_loan[0])
                    return final_doc

                elif result_loan[0]['pledges']['Category'] == 'NGAH':
                    print("result_loan", result_loan[0]['pledges']['Category'])
                    print("AWESOME SO FAR")
                    final_doc = create_ngah_la(result_loan[0])
                    return final_doc


                else:

                    for i in range(0, len(l1.lender_info)):
                        l1.lender_info[i]['text'] = ""

                    physical_address2 = ""
                    physical_address3 = ""
                    postal_address2 = ""
                    postal_address3 = ""
                    linked_unit = ""
                    investor_name = ""
                    investor_name2 = ""
                    nsst = "Martin Deon van Rooyen"
                    project = ""
                    investment_amount = ""
                    investment_interest_rate = ""
                    investor_id = ""
                    investor_id2 = ""
                    registered_company_name = ""
                    registration_number = ""
                    for i in result_loan:
                        for key in i:

                            if key == "investor_name":
                                investor_name += i[key] + " "
                            if key == "investor_surname":
                                investor_name += i[key]
                                l1.lender_info[1]['text'] = investor_name
                            if key == "investor_name2":
                                investor_name2 += i[key] + " "
                            if key == "investor_surname2":
                                investor_name2 += i[key]
                                l1.lender_info[1]['text'] = f"{investor_name} - {investor_name2}"
                            if key == "registered_company_name":
                                registered_company_name = i[key]
                                l1.lender_info[0]['text'] = i[key]

                            if key == "trading_name":
                                l1.lender_info[2]['text'] = i[key]

                            if key == "registration_number":
                                registration_number = i[key]
                                l1.lender_info[3]['text'] = i[key]

                            if key == "vat_number":
                                l1.lender_info[4]['text'] = i[key]

                            if key == "members_directors":
                                l1.lender_info[5]['text'] = i[key]

                            if key == "investor_id_number":
                                l1.lender_info[6]['text'] = i[key]
                                investor_id = i[key]

                            if key == "investor_id_number2":
                                # l1.lender_info[6]['text'] = i[key]
                                investor_id2 = i[key]

                            if key == "members_directors2":
                                l1.lender_info[7]['text'] = i[key]

                            if key == "members_directors_id2":
                                l1.lender_info[8]['text'] = i[key]

                            if key == "bank_account_holder":
                                l1.lender_info[9]['text'] = i[key]

                            if key == "bank_name":
                                l1.lender_info[10]['text'] = i[key]

                            if key == "bank_account_type":
                                l1.lender_info[11]['text'] = i[key]

                            if key == "bank_branch":
                                l1.lender_info[12]['text'] = i[key]

                            if key == "bank_account_number":
                                l1.lender_info[13]['text'] = i[key]

                            if key == "bank_branch_code":
                                l1.lender_info[14]['text'] = i[key]

                            if key == "investor_physical_street":
                                l1.lender_info[15]['text'] = i[key]
                            elif key == "investor_physical_suburb":
                                physical_address2 = physical_address2 + i[key] + " "
                                l1.lender_info[16]['text'] = physical_address2
                            elif key == "investor_physical_city":
                                physical_address2 = physical_address2 + i[key] + " "
                                l1.lender_info[16]['text'] = physical_address2

                            elif key == "investor_physical_province":
                                physical_address3 = physical_address3 + i[key] + " "
                                l1.lender_info[17]['text'] = physical_address3

                            elif key == "investor_physical_country":
                                physical_address3 = physical_address3 + i[key] + " "
                                l1.lender_info[17]['text'] = physical_address3
                            elif key == "investor_physical_postal_code":
                                physical_address3 = physical_address3 + i[key]
                                l1.lender_info[17]['text'] = physical_address3
                            elif key == "investor_postal_street_box":
                                l1.lender_info[18]['text'] = i[key]
                            elif key == "investor_postal_suburb":
                                postal_address2 = postal_address2 + i[key] + " "
                                l1.lender_info[19]['text'] = postal_address2
                            elif key == "investor_postal_city":
                                postal_address2 = postal_address2 + i[key] + " "
                                l1.lender_info[19]['text'] = postal_address2
                            elif key == "investor_postal_province":
                                postal_address3 = postal_address3 + i[key] + " "
                                l1.lender_info[19]['text'] = postal_address3
                            elif key == "investor_postal_country":
                                postal_address3 = i[key] + " "
                                l1.lender_info[20]['text'] = postal_address3
                            elif key == "investor_postal_postal_code":
                                postal_address3 = postal_address3 + i[key]
                                l1.lender_info[20]['text'] = postal_address3
                            elif key == "income_tax_number":
                                l1.lender_info[21]['text'] = i[key]
                            elif key == "investor_landline":
                                l1.lender_info[22]['text'] = i[key]
                            elif key == "investor_mobile":
                                l1.lender_info[23]['text'] = i[key]
                            elif key == "telefax_number":
                                l1.lender_info[24]['text'] = i[key]
                            elif key == "investor_email":
                                l1.lender_info[25]['text'] = i[key]

                            if key == "pledges":
                                for pledge in i[key]:

                                    if pledge == "opportunity_code":
                                        linked_unit = i[key][pledge]
                                    if pledge == "Category":
                                        project = i[key][pledge]
                                    if pledge == "investment_amount":
                                        investment_amount = i[key][pledge]
                                    if pledge == "investment_interest_rate":
                                        investment_interest_rate = i[key][pledge]

                    amt_list = []
                    investment_amount = float(investment_amount)
                    amt_list.append(investment_amount)
                    investment_amount = ['R {:0,.2f}'.format(x) for x in amt_list][0]

                    final_doc = create_final_loan_agreement(linked_unit=linked_unit, investor=investor_name,
                                                            investor2=investor_name2, nsst=nsst,
                                                            project=project, investment_amount=investment_amount,
                                                            investment_interest_rate=investment_interest_rate,
                                                            investor_id=investor_id, investor_id2=investor_id2,
                                                            registered_company_name=registered_company_name,
                                                            registration_number=registration_number)

                    print("final_doc", final_doc)
                    return final_doc

    except Exception as e:
        print(e)
        return {"ERROR": f"Something went wrong!!- {e}"}


# GET LOAN AGREEMENT AS A ZIP FILE
@investor.get("/get_loan_agreement")
async def loan_agreement(loan_agreement_name):
    loan_agreement_name = loan_agreement_name.replace('$', '&').split('/')[1]
    print("HellOOOOOOO")
    print("loan_agreement_name", loan_agreement_name)
    extension = loan_agreement_name.split('.')[1]
    print("extension", extension)
    dir_path = "loan_agreements"
    dir_list = os.listdir(dir_path)

    if loan_agreement_name in dir_list:
        if extension == "zip":
            return FileResponse(f"{dir_path}/{loan_agreement_name}", media_type="application/zip")
        else:
            return FileResponse(f"{dir_path}/{loan_agreement_name}", media_type="application/docx")
        # return FileResponse(f"{dir_path}/{loan_agreement_name}", media_type="application/zip")
    else:
        return {"ERROR": "File does not exist!!"}


# Get early releases
@investor.get("/get_early_releases")
async def get_early_releases():
    # from the database get investors and project only investor_acc_number, investor_name, investor_surname and the
    # investments array
    try:
        investors = list(db.investors.aggregate(
            [
                {
                    '$project': {
                        'investor_acc_number': 1,
                        'investor_name': 1,
                        'investor_surname': 1,
                        'investments': 1,
                        '_id': 0,
                    }
                }
            ]
        ))
        for inv in investors:
            # filter investor['invstments'] to only include investments where if early_release exists and is True

            for i in inv['investments']:
                i['early_release'] = i.get('early_release', False)
            inv['investments'] = list(filter(lambda x: x['early_release'] == True, inv['investments']))

        # filter investors to only include investors where investments is not empty
        investors = list(filter(lambda x: len(x['investments']) > 0, investors))

        early_releases = []

        # get opportunities from the database
        opportunities_list = list(db.opportunities.find({}))
        for opportunity in opportunities_list:
            opportunity['_id'] = str(opportunity['_id'])
            opportunity['id'] = opportunity['_id']
            del opportunity['_id']

        for investor in investors:
            for inv in investor['investments']:
                # filter opportunity_list to only include opportunities where opportunity_code == inv[
                # 'opportunity_code']
                opportunity_list_filtered = list(filter(lambda x: x['opportunity_code'] == inv['opportunity_code'],
                                                        opportunities_list))
                insert = {}
                insert['investor_acc_number'] = investor['investor_acc_number']
                insert['investor_name'] = investor['investor_name']
                insert['investor_surname'] = investor['investor_surname']
                insert['opportunity_code'] = inv['opportunity_code']
                insert['Category'] = opportunity_list_filtered[0]['Category']
                insert['early_release_date'] = inv['end_date']
                insert['early_release_amount'] = float(inv['exit_value'].replace(" ", ""))
                insert['investment_amount'] = float(inv['investment_amount'].replace(" ", ""))
                insert['sold'] = opportunity_list_filtered[0]['opportunity_sold']
                insert['interest_rate'] = float(inv['investment_interest_rate'])
                if opportunity_list_filtered[0]['opportunity_final_transfer_date'] == '':
                    insert['transferred'] = False
                else:
                    insert['transferred'] = True

                early_releases.append(insert)

        result = early_release_creation(early_releases)
        # return FileResponse(f"early_releases_excel_generation/early_releases.xlsx", media_type="application/xlsx")

        return result
    except Exception as e:
        print(e)
        return {"ERROR": f"Something went wrong!!- {e}"}


@investor.get("/deliver_early_releases")
async def deliver_early_releases(file_name):
    try:
        file_name = file_name + ".xlsx"

        dir_path = "early_releases_excel_generation"
        dir_list = os.listdir(dir_path)

        if file_name in dir_list:

            return FileResponse(f"{dir_path}/{file_name}", filename=file_name)
        else:
            return {"ERROR": "File does not exist!!"}
    except Exception as e:
        print(e)
        return {"ERROR": "Please Try again"}



# def importNGAH():
#     # loop through NGAH.json and print each item
#     with open('NGAH.json') as f:
#         data = json.load(f)
#         for item in data:
#             item['opportunity_final_transfer_date'] = ""
#             print(item)
#         result = db.opportunities.insert_many(data)
#         print(result.inserted_ids)

# importNGAH()

# def edit_ngah_opportunities():
#     opportunities = db.opportunities.find({"Category": "NGAH"})
#     for opportunity in opportunities:
#         opportunity['_id'] = str(opportunity['_id'])
#         id = opportunity['_id']
#
#
#         adjust = opportunity['opportunity_code'].split("-")
#         print(adjust)
#         opportunity['opportunity_code'] = f"Section {adjust[1]} - {adjust[0]}"
#         print()
#         print(opportunity['opportunity_code'])
#         try:
#             db.opportunities.update_one({"_id": ObjectId(id)}, {"$set": {"opportunity_code": opportunity['opportunity_code']}})
#             print("success")
#         except Exception as e:
#             print(e)


# edit_ngah_opportunities()