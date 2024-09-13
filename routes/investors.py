import json

from PyPDF2 import PdfFileMerger
from fastapi import APIRouter, Request, UploadFile, Form, File
from fastapi.responses import FileResponse
from decouple import config
import boto3

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
from loan_agreement_files.goodwood_exit_letters.goodwood_exit_letters import create_goodwood_exit_letters
from cashflow_excel_functions.rolled_from_and_consultant import create_workbook
from configuration.db import db
from PyPDF2 import PdfFileMerger

# AWS BUCKET INFO - ENSURE IN VARIABLES ON HEROKU
AWS_BUCKET_NAME = config("AWS_BUCKET_NAME")
AWS_BUCKET_REGION = config("AWS_BUCKET_REGION")
AWS_ACCESS_KEY_ID = config("AWS_ACCESS_KEY_ID")
AWS_SECRET_ACCESS_KEY = config("AWS_SECRET_ACCESS_KEY")

# CONNECT TO S3 SERVICE
s3 = boto3.client(
    "s3",
    aws_access_key_id=AWS_ACCESS_KEY_ID,
    aws_secret_access_key=AWS_SECRET_ACCESS_KEY
)



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
                                  "telefax_number": 1, "trading_name": 1, "vat_number": 1, "pledges": 1, "alternate_contact": 1,
                                  "alternate_contact_details": 1,
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



@investor.post("/uploadLoanAgreement")
async def uploadLoanAgreement( doc: UploadFile, name: str = Form(...), inv: str = Form(...), unit: str = Form(...)):
    # data = await request.json()
    # print(data)
    print("NAME",name)
    print("INVESTOR",inv)
    print("UNIT",unit)
    try:
        # for doc in docs:
        # print(doc)
        name = name + unit
        name = name.replace(" ", "_")
        doc.filename = doc.filename.replace(" ", "_")
        # doc.filename = name
        # print("CCCCC",doc.filename)
        data = await doc.read()
        with open(f"upload_loanAgreement/{doc.filename}", "wb") as f:
            f.write(data)

        # print("UPLOAD PARTLY SUCCESSFUL")

        input_folder = 'upload_loanAgreement'
        output_pdf = f'upload_loanAgreement/{name}.pdf'
        merge_pdfs(input_folder, output_pdf)

        # delete all files in upload_fica except for the merged pdf
        for file in os.listdir(input_folder):
            if file != f"{name}.pdf":
                os.remove(f"{input_folder}/{file}")

        # upload to s3
        try:
            s3.upload_file(
                f"upload_loanAgreement/{name}.pdf",
                AWS_BUCKET_NAME,
                f"{name}.pdf",
            )
            link = f"https://{AWS_BUCKET_NAME}.s3.{AWS_BUCKET_REGION}.amazonaws.com/{name}.pdf"
            # print("SUCCESS")
            # print(link)
            investor = db.investors.find_one({"investor_acc_number": inv})
            # print("INVESTOR",investor)
            # investor['id'] = str(investor['_id'])
            # del investor['_id']
            for i in investor['trust']:
                if i['opportunityPlusInvestment'] == unit:
                    i['loanAgreement'] = link
            db.investors.update_one({"investor_acc_number": inv}, {"$set": investor})

        except Exception as err:
            print("ERR",err)

        # delete all files in upload_fica
        for file in os.listdir(input_folder):
            os.remove(f"{input_folder}/{file}")
        unit = unit.split(" ")[0]

        return {"agreement": link, "unit": unit}
    except Exception as e:
        print("E",e)
        return {"ERROR": "Please Try again"}


@investor.post("/uploadExitLetters")
async def uploadExitLetters( doc: UploadFile, name: str = Form(...), inv: str = Form(...), unit: str = Form(...)):
    # data = await request.json()
    # print(data)
    print("NAME",name)
    print("INVESTOR",inv)
    print("UNIT",unit)
    try:
        # for doc in docs:
        # print(doc)
        name = name + unit
        name = name.replace(" ", "_")
        doc.filename = doc.filename.replace(" ", "_")
        # doc.filename = name
        # print("CCCCC",doc.filename)
        data = await doc.read()
        with open(f"upload_exit_letters/{doc.filename}", "wb") as f:
            f.write(data)

        # print("UPLOAD PARTLY SUCCESSFUL")

        input_folder = 'upload_exit_letters'
        output_pdf = f'upload_exit_letters/{name}.pdf'
        merge_pdfs(input_folder, output_pdf)

        # delete all files in upload_fica except for the merged pdf
        for file in os.listdir(input_folder):
            if file != f"{name}.pdf":
                os.remove(f"{input_folder}/{file}")

        # upload to s3
        try:
            s3.upload_file(
                f"upload_exit_letters/{name}.pdf",
                AWS_BUCKET_NAME,
                f"{name}.pdf",
            )
            link = f"https://{AWS_BUCKET_NAME}.s3.{AWS_BUCKET_REGION}.amazonaws.com/{name}.pdf"
            print("SUCCESS")
            print(link)
            investor = db.investors.find_one({"investor_acc_number": inv})
            # print("INVESTOR",investor)
            # investor['id'] = str(investor['_id'])
            # del investor['_id']
            for i in investor['trust']:
                if i['opportunityPlusInvestment'] == unit:
                    i['exit_rollover_documents'] = link
            db.investors.update_one({"investor_acc_number": inv}, {"$set": investor})

        except Exception as err:
            print("ERR",err)

        # delete all files in upload_fica
        for file in os.listdir(input_folder):
            os.remove(f"{input_folder}/{file}")
        unit = unit.split(" ")[0]

        return {"exitDoc": link, "unit": unit}
    except Exception as e:
        print("E",e)
        return {"ERROR": "Please Try again"}


@investor.post("/upload_fica_files/")
async def upload_fica_files( docs: list[UploadFile], name: str = Form(...), inv: str = Form(...)):
    # data = await request.json()
    # print(data)
    print("NAME",name)
    try:
        for doc in docs:
            # print(doc)
            doc.filename = doc.filename.replace(" ", "_")
            print(doc.filename)
            data = await doc.read()
            with open(f"upload_fica/{doc.filename}", "wb") as f:
                f.write(data)

        input_folder = 'upload_fica'
        output_pdf = f'upload_fica/{name}.pdf'
        merge_pdfs(input_folder, output_pdf)

        # delete all files in upload_fica except for the merged pdf
        for file in os.listdir(input_folder):
            if file != f"{name}.pdf":
                os.remove(f"{input_folder}/{file}")

        # upload to s3
        try:
            s3.upload_file(
                f"upload_fica/{name}.pdf",
                AWS_BUCKET_NAME,
                f"{name}.pdf",
            )
            link = f"https://{AWS_BUCKET_NAME}.s3.{AWS_BUCKET_REGION}.amazonaws.com/{name}.pdf"
            print("SUCCESS")
            print(link)
            investor = db.investors.find_one({"investor_acc_number": inv})
            # print("INVESTOR",investor)
            # investor['id'] = str(investor['_id'])
            # del investor['_id']
            # for i in investor['trust']:
            #     if i['opportunityPlusInvestment'] == unit:
            #         i['exit_rollover_documents'] = link
            investor['fileFica'] = link
            db.investors.update_one({"investor_acc_number": inv}, {"$set": investor})

        except Exception as err:
            print(err)

        # delete all files in upload_fica
        for file in os.listdir(input_folder):
            os.remove(f"{input_folder}/{file}")

        return {"fica": link}
    except Exception as e:
        print(e)
        return {"ERROR": "Please Try again"}


@investor.post("/deleteAgreement")
async def delete_agreement_file(data: Request):
    request = await data.json()
    print(request)
    # return {"message": "File Deleted"}
    investor = db.investors.find_one({"investor_acc_number": request['investor_acc_number']})
    # print("INVESTOR",investor['trust'])
    for i in investor['trust']:
        if i.get('opportunityPlusInvestment',"") == request['name']:
            i['loanAgreement'] = ""
    db.investors.update_one({"investor_acc_number": request['investor_acc_number']}, {"$set": investor})
    return {"message": "File Deleted"}


@investor.post("/deleteExit")
async def delete_exit_file(data: Request):
    request = await data.json()
    print(request)
    # return {"message": "File Deleted"}
    investor = db.investors.find_one({"investor_acc_number": request['investor_acc_number']})
    # print("INVESTOR",investor['trust'])
    for i in investor['trust']:
        if i.get('opportunityPlusInvestment',"") == request['name']:
            i['exit_rollover_documents'] = ""
    db.investors.update_one({"investor_acc_number": request['investor_acc_number']}, {"$set": investor})
    return {"message": "File Deleted"}

@investor.post("/upload_general_files/")
async def upload_general_files( docs: list[UploadFile], name: str = Form(...),inv: str = Form(...)):
    # data = await request.json()
    # print(data)
    print("NAME",name)
    try:
        for doc in docs:
            # print(doc)
            doc.filename = doc.filename.replace(" ", "_")
            print(doc.filename)
            data = await doc.read()
            with open(f"upload_general/{doc.filename}", "wb") as f:
                f.write(data)

        input_folder = 'upload_general'
        output_pdf = f'upload_general/{name}.pdf'
        merge_pdfs(input_folder, output_pdf)

        # delete all files in upload_fica except for the merged pdf
        for file in os.listdir(input_folder):
            if file != f"{name}.pdf":
                os.remove(f"{input_folder}/{file}")

        # upload to s3
        try:
            s3.upload_file(
                f"upload_general/{name}.pdf",
                AWS_BUCKET_NAME,
                f"{name}.pdf",
            )
            link = f"https://{AWS_BUCKET_NAME}.s3.{AWS_BUCKET_REGION}.amazonaws.com/{name}.pdf"
            print("SUCCESS")
            print(link)
            investor = db.investors.find_one({"investor_acc_number": inv})
            # print("INVESTOR",investor)
            # investor['id'] = str(investor['_id'])
            # del investor['_id']
            # for i in investor['trust']:
            #     if i['opportunityPlusInvestment'] == unit:
            #         i['exit_rollover_documents'] = link
            investor['fileGeneral'] = link
            db.investors.update_one({"investor_acc_number": inv}, {"$set": investor})

        except Exception as err:
            print(err)

        # delete all files in upload_fica
        for file in os.listdir(input_folder):
            os.remove(f"{input_folder}/{file}")

        return {"general": link}
    except Exception as e:
        print(e)
        return {"ERROR": "Please Try again"}


def merge_pdfs(input_folder, output_pdf):
    """
    Merge all PDF files in the input folder and save the merged PDF to the output file.

    Args:
        input_folder (str): Path to the folder containing the PDF files to merge.
        output_pdf (str): Path to the output merged PDF file.
    """
    merger = PdfFileMerger()

    # Get a list of all PDF files in the input folder
    pdf_files = [os.path.join(input_folder, f) for f in os.listdir(input_folder) if f.endswith('.pdf')]

    # Append each PDF file to the merger
    for pdf_file in pdf_files:
        with open(pdf_file, 'rb') as file:
            merger.append(file)

    # Write the merged PDF to the output file
    with open(output_pdf, 'wb') as output_file:
        merger.write(output_file)

    print(f'Merged PDFs successfully. Output saved to: {output_pdf}')


@investor.post("/exit_letters")
async def early_exits(data: Request):
    request = await data.json()
    # print(request)
    investor = db.investors.find_one({"investor_acc_number": request['investor_acc_number']},{"investor_name": 1, "investor_surname": 1, "investor_acc_number": 1, "investments": 1,"investor_organisation": 1, "_id": 0})
    investor['investments'] = list(filter(lambda x: x['opportunity_code'] == request['opportunity_code'], investor['investments']))
    data = {
        "investor_name": f"{investor['investor_name']} {investor['investor_surname']}",
        "investor_acc_number": investor['investor_acc_number'],
        "opportunity_code": request['opportunity_code'],
        "exit_date": investor['investments'][0]['end_date'],
        "exit_amount": investor['investments'][0]['exit_amount'],
        "rollover_amount": investor['investments'][0]['rollover_amount'],
        "exit_value": investor['investments'][0]['exit_value'],
        "investment_exit_rollover": investor['investments'][0]['investment_exit_rollover'],
        "investor_organisation": investor['investor_organisation']
    }
    # print("Data",data)

    result = create_goodwood_exit_letters(data)
    return result

@investor.get("/get_exit_letter")
async def get_exit_letter(file_name):
    try:
        file_name = file_name

        dir_path = "loan_agreements"
        dir_list = os.listdir(dir_path)

        if file_name in dir_list:

            return FileResponse(f"{dir_path}/{file_name}", filename=file_name)
        else:
            return {"ERROR": "File does not exist!!"}
    except Exception as e:
        print(e)
        return {"ERROR": "Please Try again"}


    # print(investor)
    # return request

@investor.post("/get_rollover_consultants")
async def get_rollover_consultants():

    consultants = list(db.investmentCommissionStatement.find({},{"_id": 0, "development": 1, "unit_number": 1, "investor": 1, "investment_amount": 1, "investor_acc_number": 1, "consultantName": 1}))
    # consultants = list(consultants)
    # print(consultants)
    investors = list(db.investors.find({},{"_id": 0, "investor_acc_number": 1, "investor_name": 1, "investor_surname": 1, "investments": 1}))
    # print()
    # print(investors)
    investments = []
    for inv in investors:
        # filter investments to only include investments where Category != Southwark and Category != TEST and Category != Endulini
        inv['investments'] = list(filter(lambda x: x['Category'] != "Southwark" and x['Category'] != "TEST" and x['Category'] != "Endulini", inv['investments']))
        for i in inv['investments']:

            insert = {
                "investor_acc_number": inv['investor_acc_number'],
                "investor_name": f"{inv['investor_name']} {inv['investor_surname']}",
                "Category": i['Category'],
                "opportunity_code": i['opportunity_code'],
                "investment_amount": float(i['investment_amount']),
                "opportunity_code_rolled_from": i.get('opportunity_code_rolled_from',""),
            }
            investments.append(insert)

        # print(investments)
        for item in investments:
            consultants_filtered = list(filter(lambda x: x['investor_acc_number'] == item['investor_acc_number'] and x['unit_number'] == item['opportunity_code'] and x["investment_amount"], consultants))
            if len(consultants_filtered) > 0:
                item['consultantName'] = consultants_filtered[0]['consultantName']
            else:
                item['consultantName'] = ""

    print(len(investments))

    result = create_workbook(investments)
    return {"href": result }

@investor.get("/get_rolled_from")
async def get_rolled_from(file_name):
    try:
        print("file_name",file_name)
        file_name = file_name
        print(file_name)


        dir_path = "excel_files"
        dir_list = os.listdir(dir_path)

        if file_name in dir_list:

            return FileResponse(f"{dir_path}/{file_name}", filename=file_name)
        else:
            return {"ERROR": "File does not exist!!"}
    except Exception as e:
        print(e)
        return {"ERROR": "Please Try again"}

    # return investments

# def get_opportunities():
#     try:
#         opportunities = list(db.opportunities.find({}))
#         for index, opportunity in enumerate(opportunities):
#             opportunity['_id'] = str(opportunity['_id'])
#             id = opportunity['_id']
#
#             del opportunity['_id']
#             if opportunity['Category'] == "Goodwood":
#                 opportunity['block'] = "R"
#             elif opportunity['Category'] == "Southwark":
#                 opportunity['block'] = "S"
#             elif opportunity['Category'] == "NGAH":
#                 block = opportunity['opportunity_code'].split(" ")
#                 block = block[len(block) - 1][0:-3]
#                 # print("block",block)
#                 opportunity['block'] = block
#             else:
#                 opportunity['block'] = opportunity['opportunity_code'][-4:-3]
#                 # print("block",opportunity['opportunity_code'][-4:-3])
#             print(opportunity)
#             print()
#             # update opportunities with the block
#             db.opportunities.update_one({"_id": ObjectId(id)}, {"$set": { "block": opportunity['block']}})
#         # print(opportunities[0])
#
#     except Exception as e:
#         print("Error getting opportunities", e)

# get_opportunities()