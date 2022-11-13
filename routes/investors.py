from fastapi import APIRouter
from fastapi.responses import FileResponse
from config.db import db
import os
from bson.objectid import ObjectId
from pydantic import BaseModel
import loan_agreement_files.lender as l1

from main import create_final_loan_agreement

from verify_token import verify_jwt_token


class InvestorAccNumber(BaseModel):
    investor_acc_number: str
    opportunity_code: str
    token_received: str


class InvestorFile(BaseModel):
    file_name: str


investor = APIRouter()

investors = db.investors


@investor.get("/")
async def get_all_investors():
    # result = list(investors.find({}, {'_id': 0}))
    result = "AWESOME IT WORKS!!"
    return result


@investor.post("/investorloanagreement")
async def get_investor_for_loan_agreement(investor_acc_number: InvestorAccNumber):
    # token_verification = verify_jwt_token(investor_acc_number.token_received)
    token_verification = ""
    print(token_verification)
    if token_verification == "Verification Failed":
        return {"error":"User Not Verified"}
    else:

        # result = investors.find_one({"investor_acc_number": investor_acc_number.investor_acc_number}, {'_id': 0})
        result_loan = list(investors.aggregate(
            [{"$match": {"investor_acc_number": investor_acc_number.investor_acc_number}}, {"$unwind": "$pledges"},
             {"$match": {"pledges.opportunity_code": investor_acc_number.opportunity_code}}, {
                 "$project": {"investor_name": 1, "investor_surname": 1, "investor_id_number": 1, "investor_email": 1,
                              "investor_mobile": 1, "investor_landline": 1, "investor_physical_street": 1,
                              "investor_physical_suburb": 1, "investor_physical_city": 1,
                              "investor_physical_province": 1,

                              "investor_physical_country": 1, "investor_physical_postal_code": 1,
                              "investor_postal_street_box": 1, "investor_postal_suburb": 1, "investor_postal_city": 1,
                              "investor_postal_province": 1, "investor_postal_country": 1,
                              "investor_postal_postal_code": 1,
                              "registered_company_name": 1, "bank_account_holder": 1, "bank_account_type": 1,
                              "bank_branch": 1, "bank_account_number": 1, "bank_branch_code": 1, "bank_name": 1,
                              "income_tax_number": 1,
                              "members_directors": 1, "members_directors2": 1, "members_directors_id": 1,
                              "members_directors_id2": 1, "registration_number": 1,
                              "telefax_number": 1, "trading_name": 1, "vat_number": 1, "pledges": 1, "_id": 0}}]))

        if len(result_loan) == 0:
            return {
                "error": f"No data found for {investor_acc_number.investor_acc_number} and {investor_acc_number.opportunity_code}"}
        else:

            for i in range(0, len(l1.lender_info)):
                # print(i)
                l1.lender_info[i]['text'] = ""

            physical_address2 = ""
            physical_address3 = ""
            postal_address2 = ""
            postal_address3 = ""
            linked_unit = ""
            investor_name = ""
            nsst = "Martin Deon van Rooyen"
            project = ""
            investment_amount = ""
            investment_interest_rate = ""
            for i in result_loan:
                for key in i:
                    if key == "investor_name":
                        investor_name += i[key] + ""
                    if key == "investor_surname":
                        investor_name += i[key]
                        l1.lender_info[1]['text'] = investor_name
                    if key == "registered_company_name":
                        l1.lender_info[0]['text'] = i[key]

                    if key == "trading_name":
                        l1.lender_info[2]['text'] = i[key]

                    if key == "registration_number":
                        l1.lender_info[3]['text'] = i[key]

                    if key == "vat_number":
                        l1.lender_info[4]['text'] = i[key]

                    if key == "members_directors":
                        l1.lender_info[5]['text'] = i[key]

                    if key == "members_directors_id":
                        l1.lender_info[6]['text'] = i[key]

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
                        physical_address2 = physical_address2 + i[key] + ", "
                        l1.lender_info[16]['text'] = physical_address2
                    elif key == "investor_physical_city":
                        physical_address2 = physical_address2 + i[key] + " "
                        l1.lender_info[16]['text'] = physical_address2

                    elif key == "investor_physical_province":
                        physical_address3 = physical_address3 + i[key] + ", "
                        l1.lender_info[17]['text'] = physical_address3

                    elif key == "investor_physical_country":
                        physical_address3 = physical_address3 + i[key] + ", "
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
                            # print(pledge)
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

            final_doc = create_final_loan_agreement(linked_unit=linked_unit, investor=investor_name, nsst=nsst,
                                                    project=project, investment_amount=investment_amount,
                                                    investment_interest_rate=investment_interest_rate)
            return final_doc


@investor.get("/get_loan_agreement")
async def loan_agreement(loan_agreement_name):
    dir_path = "loan_agreements"
    dir_list = os.listdir(dir_path)

    if loan_agreement_name in dir_list:
        return FileResponse(f"loan_agreements/{loan_agreement_name}", media_type="application/pdf")
    else:
        return {"ERROR": "File does not exist!!"}

# http://127.0.0.1:8000/get_loan_agreement?loan_agreement_name=Loan_Agreement_Wayne%20Bruton_Heron%20View_HVC102.pdf
