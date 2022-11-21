from fastapi import APIRouter, Request
from fastapi.responses import FileResponse
from config.db import db
import os
import time
from datetime import datetime
from datetime import timedelta
from pydantic import BaseModel
import loan_agreement_files.lender as l1
from main import create_final_loan_agreement


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


@investor.post("/salesforecast")
async def get_sales_info(data: Request):
    request = await data.json()
    initial_forecast_data = list(investors.aggregate([
        {
            '$project': {
                'id': {
                    '$toString': '$_id'
                },
                "_id": 0,
                'investor_acc_number': 1,
                'investor_name': 1,
                'investor_surname': 1,
                'investments': 1,
                'trust': 1
            }
        }
    ]))

    opportunities_listed = list(opportunities.aggregate([
        {
            '$match': {
                'Category': {
                    '$in': request["Category"]
                }
            }
        },
        {
            '$project': {
                'id': {
                    '$toString': '$_id'
                },
                '_id': 0,
                'opportunity_code': 1,
                'Category': 1,
                'opportunity_amount_required': 1,
                'opportunity_end_date': 1,
                'opportunity_final_transfer_date': 1,
                'opportunity_sale_price': 1,
                'opportunity_sold': 1
            }
        }
    ]))

    rates_retrieved = list(rates.aggregate([
        {
            '$project': {
                'id': {
                    '$toString': '$_id'
                },
                '_id': 0,
                'Efective_date': 1,
                'rate': 1
            }
        }
    ]))

    costs = list(sales_parameters.aggregate([
        {
            '$match': {
                'Development': {
                    '$in': request["Category"]
                }
            }
        }, {
            '$project': {
                'id': {
                    '$toString': '$_id'
                },
                '_id': 0,
                'Effective_date': 1,
                'Development': 1,
                'Description': 1,
                'rate': 1
            }
        }
    ]))

    project_rollovers = list(rollovers.aggregate([
        {
            '$match': {
                'Category': {
                    '$in': request["Category"]
                }
            }
        }, {
            '$project': {
                'id': {
                    '$toString': '$_id'
                },
                '_id': 0,
                'investor_acc_number': 1,
                'investment_number': 1,
                'opportunity_code': 1,
                'available_date': 1,
                'rollover_amount': 1
            }
        }
    ]))

    # FILTER OUT NON RELEVENT OPPORTUNITIES
    for inv in initial_forecast_data:
        if len(request["Category"]) == 1:
            inv["trust"] = [x for x in inv["trust"] if x["Category"] == request["Category"][0]]
            inv["investments"] = [x for x in inv["investments"] if x["Category"] == request["Category"][0]]
        elif len(request["Category"]) == 2:
            inv["trust"] = [x for x in inv["trust"] if
                            x["Category"] == request["Category"][0] or x["Category"] == request["Category"][1]]
            inv["investments"] = [x for x in inv["investments"] if
                                  x["Category"] == request["Category"][0] or x["Category"] == request["Category"][1]]
    initial_forecast_data = [x for x in initial_forecast_data if len(x["trust"]) > 0]

    # SEPERATE INVESTMENTS AND TRUST

    trust_items = []
    investment_items = []
    interim_data = []

    finalData = []
    for inv in initial_forecast_data:
        for t in inv['trust']:
            t['investor_acc_number'] = inv['investor_acc_number']
            t["investor_name"] = f"{inv['investor_surname']} {inv['investor_name'][0:1]}"
            trust_items.append(t)

        for i in inv['investments']:
            i['investor_acc_number'] = inv['investor_acc_number']
            i["investor_name"] = f"{inv['investor_surname']} {inv['investor_name'][0:1]}"
            investment_items.append(i)

    # SORT RATES BY DATE DESCENDING
    for rate_item in rates_retrieved:
        rate_item["Efective_date"] = datetime.strptime(rate_item["Efective_date"], '%Y-%m-%d')
    rates_retrieved = sorted(rates_retrieved, key=lambda d: d['Efective_date'], reverse=True)

    # FILTER COSTS
    costs = [x for x in costs if x["Development"] == request["Category"][0]]

    # CREATE FINAL DATA (INITIAL)
    for opportunity in opportunities_listed:
        opportunity_code = opportunity["opportunity_code"]

        if "opportunity_end_date" not in opportunity:
            opportunity["opportunity_end_date"] = ""
        # ADD BASIC UNIT INFO TO FINAL DATA
        insert = {}
        insert["opportunity_code"] = opportunity["opportunity_code"]
        insert["opportunity_amount_required"] = opportunity["opportunity_amount_required"]
        insert["opportunity_end_date"] = opportunity["opportunity_end_date"]
        insert["opportunity_final_transfer_date"] = opportunity["opportunity_final_transfer_date"]
        insert["opportunity_sale_price"] = opportunity["opportunity_sale_price"]
        insert["opportunity_sold"] = opportunity["opportunity_sold"]

        if insert["opportunity_end_date"] == "":
            insert["opportunity_end_date"] = request["date"]

        if opportunity["opportunity_final_transfer_date"] == "":
            insert["opportunity_transferred"] = False
            opportunity["opportunity_final_transfer_date"] = insert["opportunity_end_date"]
        else:
            insert["opportunity_transferred"] = True
        insert["report_date"] = request["date"]

        for cost in costs:
            key = cost["Description"]
            rate = cost["rate"]
            insert[key] = rate

            interim_data.append(insert)

    final_data_trust = []
    final_data_trust_and_investment = []

    # REMOVE DUPLICATES
    interim_data = [i for n, i in enumerate(interim_data)
                    if i not in interim_data[n + 1:]]

    # POPULATE FINAL TRUST ITEMS
    for item in trust_items:
        opportunity_code = item['opportunity_code']
        filtered_data = list(filter(lambda x: x['opportunity_code'] == opportunity_code, interim_data))
        # print(len(filtered_data))
        insert1 = {}
        for data in filtered_data:
            for key in data:
                insert1[key] = data[key]
        insert1['investment_amount'] = float(item['investment_amount'])
        insert1["investor_acc_number"] = item["investor_acc_number"]
        insert1["investor_name"] = item["investor_name"]
        insert1['deposit_date'] = item['deposit_date']
        insert1['release_date'] = item['release_date']
        insert1['investment_number'] = item['investment_number']

        final_data_trust.append(insert1)

    final_data_trust = [i for n, i in enumerate(final_data_trust)
                        if i not in final_data_trust[n + 1:]]

    # POPULATE FINAL TRUST ITEMS WITH INVESTMENTS
    for item in investment_items:
        opportunity_code = item['opportunity_code']
        try:
            filtered_data = list(filter(
                lambda x: x['opportunity_code'] == opportunity_code and x['investor_acc_number'] == item[
                    'investor_acc_number'] and x['investment_number'] == item['investment_number'], final_data_trust))
        except:
            filtered_data = []
        if len(filtered_data) > 0:

            insert2 = {}
            for data in filtered_data:
                for key in data:
                    insert2[key] = data[key]
            insert2['released_investment_amount'] = float(item['investment_amount'])
            insert2['investment_end_date'] = item['end_date']
            insert2['investment_release_date'] = item['release_date']
            insert2['released_investment_number'] = item['investment_number']

            final_data_trust_and_investment.append(insert2)
        # else:
        #     insert3 = {}
        #     for data in filtered_data:
        #         for key in data:
        #             insert3[key] = data[key]
        #     insert3['released_investment_amount'] = 0
        #     insert3['investment_end_date'] = ""
        #     insert3['investment_release_date'] = ""
        #     insert3['opportunity_code'] = item["opportunity_code"]
        #     insert3['investor_acc_number'] = "item["opportunity_code"]"
        #     insert3['released_investment_number'] = 999
        #
        #     final_data_trust_and_investment.append(insert3)

    # print(len(final_data_trust_and_investment))
    # test = []
    # for x in final_data_trust_and_investment:
    #     try:
    #         test.append(x["investor_acc_number"])
    #     except:
    #         print(x)
    # print(len(test))
    # POPULATE WITH UNALLOCATED DATA
    for item in interim_data:
        # try:
        filtered_data = list(filter(
            lambda x: x['opportunity_code'] == item['opportunity_code'], final_data_trust_and_investment))
        # except:
        #     filtered_data = []
        if len(filtered_data) == 0:
            insert = {}
            for key in item:
                insert[key] = item[key]
            insert['investor_acc_number'] = "ZZUN01"
            insert['investor_name'] = "Allocated UN"
            insert['opportunity_code'] = item['opportunity_code']
            final_data_trust_and_investment.append(insert)

    final_data_trust_and_investment = sorted(final_data_trust_and_investment, key=lambda z: (z['opportunity_code'], z['investor_acc_number']))

    return final_data_trust_and_investment


@investor.post("/investorloanagreement")
async def get_investor_for_loan_agreement(investor_acc_number: InvestorAccNumber):
    # token_verification = verify_jwt_token(investor_acc_number.token_received)
    token_verification = "ABC"
    # print(token_verification)
    if token_verification == "Verification Failed":
        return {"error": "User Not Verified"}
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
                              "telefax_number": 1, "trading_name": 1, "vat_number": 1, "pledges": 1,
                              "id": {'$toString': "$_id"}, "_id": 0, }}]))

        # print(result_loan)
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
            investor_id = ""
            for i in result_loan:
                for key in i:
                   # if key == "investor_id_number":

                    if key == "investor_name":
                        investor_name += i[key] + " "
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

                    if key == "investor_id_number":
                        l1.lender_info[6]['text'] = i[key]
                        investor_id = i[key]

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
                                                    investment_interest_rate=investment_interest_rate, investor_id=investor_id)
            return final_doc


@investor.get("/get_loan_agreement")
async def loan_agreement(loan_agreement_name):
    dir_path = "loan_agreements"
    dir_list = os.listdir(dir_path)
    # print("AAWESOMEEEEE")
    if loan_agreement_name in dir_list:
        return FileResponse(f"loan_agreements/{loan_agreement_name}", media_type="application/pdf")
    else:
        return {"ERROR": "File does not exist!!"}

# http://127.0.0.1:8000/get_loan_agreement?loan_agreement_name=Loan_Agreement_Wayne%20Bruton_Heron%20View_HVC102.pdf
