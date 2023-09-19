import os

from bson import ObjectId
from fastapi import APIRouter, Request, BackgroundTasks
from fastapi.responses import FileResponse
from excel_sf_functions.sales_forecast_excel import create_sales_forecast_file
from excel_sf_functions.sales_forecast_excel import create_cash_flow
from config.db import db
import time
from datetime import datetime
from datetime import timedelta
from excel_sf_functions.draw_downs_excel import create_draw_down_file
from excel_sf_functions.draw_history_report import create_draw_history_report
# from excel_sf_functions.investment_list_excel import create_investment_list_file

# from fastapi import File, UploadFile
import smtplib
from email.message import EmailMessage

excel_sales_forecast = APIRouter()

# MONGO COLLECTIONS
investors = db.investors
rates = db.rates
opportunities = db.opportunities
unallocated_investments = db.unallocated_investments
sales_parameters = db.salesParameters
rollovers = db.investorRollovers
drawdowns = db.investor_Draws


@excel_sales_forecast.post("/update-sales-forecast_unallocated")
async def update_unallocated_investments(data: Request):
    request = await data.json()

    try:
        # Update this record in the unallocated_investments collection in the database
        unallocated_investments.update_one(
            {"_id": ObjectId(request['id'])},
            {
                "$set": {
                    "opportunity_code": request['opportunity_code'],
                    "opportunity_amount_required": request['opportunity_amount_required'],
                    "total_investment": request['total_investment'],
                    "unallocated_investment": request['unallocated_investment'],
                    "deposit_date": request['deposit_date'],
                    "release_date": request['release_date'],
                    "project_interest_rate": float(request['project_interest_rate']),
                    "opportunity_end_date": request['opportunity_end_date'],
                    "opportunity_final_transfer_date": request['opportunity_final_transfer_date'],
                    "Category": request['Category'],
                }
            },
        )

        return {"message": "success"}

    except Exception as e:
        print(e)
        return {"message": "error"}


# GET UNALLOCATTED INVESTMENTS
@excel_sales_forecast.post("/sales-forecast_unallocated")
async def get_unallocated_investments():
    try:

        trust_list = []

        investor_list = list(db.investors.find({}))
        for investor in investor_list:
            investor['id'] = str(investor['_id'])
            del investor['_id']
            for trust in investor['trust']:
                insert = {'opportunity_code': trust['opportunity_code'],
                          'investment_amount': float(trust['investment_amount']), 'Category': trust['Category']}
                trust_list.append(insert)

        # Filter out from trust_list where Category is not Southwark
        trust_list = [trust for trust in trust_list if trust['Category'] != 'Southwark']

        calculated_unallocated_investments_list = []

        opportunities_list = list(db.opportunities.find({}))

        for opportunity in opportunities_list:
            # if opportunity['opportunity_end_date'] exists then use it else make it equal to ""
            if 'opportunity_end_date' in opportunity:
                opportunity['opportunity_end_date'] = opportunity['opportunity_end_date']
            else:
                opportunity['opportunity_end_date'] = ""
            # Do the same for opportunity_final_transfer_date
            if 'opportunity_final_transfer_date' in opportunity:
                opportunity['opportunity_final_transfer_date'] = opportunity['opportunity_final_transfer_date']
            else:
                opportunity['opportunity_final_transfer_date'] = ""

            opportunity['id'] = str(opportunity['_id'])
            del opportunity['_id']

            # In a new list filter trust_list where opportunity_code is equal to opportunity['opportunity_code']
            # using list comprehension
            opportunity_trust_list = [trust for trust in trust_list if
                                      trust['opportunity_code'] == opportunity['opportunity_code']]

            # if len(opportunity_trust_list) > 0 sum the investment_amounts as a variable
            if len(opportunity_trust_list) > 0:
                total_investment_amount = sum([float(trust['investment_amount']) for trust in opportunity_trust_list])
            else:
                total_investment_amount = float(0)
            if float(opportunity['opportunity_amount_required']) > total_investment_amount:
                insert = {'opportunity_code': opportunity['opportunity_code'],
                          'opportunity_amount_required': opportunity['opportunity_amount_required'],
                          'total_investment': total_investment_amount, 'unallocated_investment': float(
                        opportunity['opportunity_amount_required']) - total_investment_amount, 'deposit_date': "",
                          'release_date': "", 'project_interest_rate': 18, 'Category': opportunity['Category'],
                          'opportunity_end_date': opportunity['opportunity_end_date'],
                          'opportunity_final_transfer_date': opportunity['opportunity_final_transfer_date']}
                calculated_unallocated_investments_list.append(insert)

        unallocated_investments_list = list(db.unallocated_investments.find({}))
        for unallocated_investment in unallocated_investments_list:
            unallocated_investment['id'] = str(unallocated_investment['_id'])
            del unallocated_investment['_id']

        for item in calculated_unallocated_investments_list:
            # filter unallocated_investments_list where opportunity_code is equal to item['opportunity_code']
            # using list comprehension
            unallocated_investment_to_be_processed = [unallocated_investment for unallocated_investment in
                                                      unallocated_investments_list if
                                                      unallocated_investment['opportunity_code'] == item[
                                                          'opportunity_code']]
            # If unalocated_investment_to_be_processed is not empty then update the item with the deposit_date,
            # release_date and project_interest_rate
            if len(unallocated_investment_to_be_processed) > 0:
                item['deposit_date'] = unallocated_investment_to_be_processed[0]['deposit_date']
                item['release_date'] = unallocated_investment_to_be_processed[0]['release_date']
                item['project_interest_rate'] = unallocated_investment_to_be_processed[0]['project_interest_rate']

        # Sort calculated_unallocated_investments_list by Category and opportunity_code
        # Filter out where Category is Southwark
        calculated_unallocated_investments_list = [item for item in calculated_unallocated_investments_list if
                                                   item['Category'] != 'Southwark']
        calculated_unallocated_investments_list = sorted(calculated_unallocated_investments_list,
                                                         key=lambda k: (k['Category'], k['opportunity_code']))

        # in Mongo delete all from unallocated_investments
        db.unallocated_investments.delete_many({})
        # In Mongo, insert all from calculated_unallocated_investments_list
        db.unallocated_investments.insert_many(calculated_unallocated_investments_list)

        unallocated_investments_list = list(db.unallocated_investments.find({}))
        for unallocated_investment in unallocated_investments_list:
            unallocated_investment['id'] = str(unallocated_investment['_id'])
            del unallocated_investment['_id']

        return unallocated_investments_list

    except Exception as e:
        print("Error", e)
        return e


@excel_sales_forecast.post("/sales-forecast")
async def get_sales_info(background_tasks: BackgroundTasks, data: Request):
    start = time.time()
    request = await data.json()

    # print("request", request['firstName'])

    firstName = request['firstName']

    developments = request['Category']

    if len(developments) > 1:
        filename = f'excel_files/Sales Forecast{developments[0].split(" ")[0]}.xlsx'
    else:
        filename = f'excel_files/Sales Forecast{developments[0]}.xlsx'

    if os.path.exists(filename):
        os.remove(filename)
        print("File Removed!")
    else:
        print("The file does not exist")

    # Get Investors and Manipulate accordingly
    investments = []
    pledges = []
    trust_list = []

    try:

        # investor_list = list(db.investors.find({}))

        # pipeline = [
        #     {
        #         "$match": {
        #             "trust.Category": {"$in": request['Category']}
        #         }
        #     },
        #     {
        #         "$project": {
        #             "_id": 1,
        #             "investor_acc_number": 1,
        #             "investor_name": 1,
        #             "investor_surname": 1,
        #             "pledges": 1,
        #             "trust": {
        #                 "$filter": {
        #                     "input": "$trust",
        #                     "as": "item",
        #                     "cond": {"$in": ["$$item.Category", request['Category']]}
        #                 }
        #             },
        #             "investments": {
        #                 "$filter": {
        #                     "input": "$investments",
        #                     "as": "item",
        #                     "cond": {"$in": ["$$item.Category", request['Category']]}
        #                 }
        #             }
        #         }
        #     },
        #     {
        #         "$match": {
        #             "trust": {"$ne": []}
        #         }
        #     }
        # ]
        #
        # investor_list = list(db.investors.aggregate(pipeline))
        # opportunities_list = list(db.opportunities.find( {"Category": {"$in": request['Category']}}))
        # sales_parameters_list = list(db.salesParameters.find({}))
        # rollovers_list = list(db.investorRollovers.find( {"Category": {"$in": request['Category']}}))
        # rates_list = list(db.rates.find({}))
        # unallocated_investments_list = list(db.unallocated_investments.find(
        #     {"Category": {"$in": request['Category']}}))

        pipeline = [
            {
                "$match": {
                    "trust.Category": {"$in": request['Category']}
                }
            },
            {
                "$project": {
                    "_id": 1,
                    "investor_acc_number": 1,
                    "investor_name": 1,
                    "investor_surname": 1,
                    # "pledges": 1,
                    "trust": {
                        "$filter": {
                            "input": "$trust",
                            "as": "item",
                            "cond": {"$in": ["$$item.Category", request['Category']]}
                        }
                    },
                    "investments": {
                        "$filter": {
                            "input": "$investments",
                            "as": "item",
                            "cond": {"$in": ["$$item.Category", request['Category']]}
                        }
                    },
                    "pledges": {
                        "$filter": {
                            "input": "$pledges",
                            "as": "item",
                            "cond": {"$in": ["$$item.Category", request['Category']]}
                        }
                    },

                }
            },
            {
                "$match": {
                    "trust": {"$ne": []}
                }
            }
        ]
        # query_start = time.time()
        result = {
            "investor_list": list(db.investors.aggregate(pipeline)),
            "opportunities_list": list(db.opportunities.find({"Category": {"$in": request['Category']}})),
            "sales_parameters_list": list(db.salesParameters.find({"Development": {"$in": request['Category']}})),
            "rollovers_list": list(db.investorRollovers.find({"Category": {"$in": request['Category']}})),
            "rates_list": list(db.rates.find({})),
            "unallocated_investments_list": list(
                db.unallocated_investments.find({"Category": {"$in": request['Category']}}))
        }

        investor_list = result['investor_list']
        opportunities_list = result['opportunities_list']
        sales_parameters_list = result['sales_parameters_list']
        rollovers_list = result['rollovers_list']
        rates_list = result['rates_list']
        unallocated_investments_list = result['unallocated_investments_list']

        # query_end = time.time()
        # print("Query Time", query_end - query_start)

        for investor in investor_list:

            investor['id'] = str(investor['_id'])
            del investor['_id']

            # CREATE INSERT DICT TO INSERT INTO RELEVANT LISTS

            if len(investor['pledges']):
                for pledge in investor['pledges']:
                    insert = {
                        'id': investor['id'],
                        'investor_surname': investor['investor_surname'],
                        'investor_name': investor['investor_name'],
                        'investor_acc_number': investor['investor_acc_number'],
                    }
                    for item in pledge:
                        insert[item] = pledge[item]
                    if 'link' in insert:
                        del insert['link']
                    pledges.append(insert)
            # Do the same for investments

            if len(investor['investments']):
                for investment in investor['investments']:
                    insert = {
                        'id': investor['id'],
                        'investor_surname': investor['investor_surname'],
                        'investor_name': investor['investor_name'],
                        'investor_acc_number': investor['investor_acc_number'],
                    }
                    for item in investment:
                        insert[item] = investment[item]
                    investments.append(insert)
            # Do the same for trust
            if len(investor['trust']):
                for trust_item in investor['trust']:
                    insert = {
                        'id': investor['id'],
                        'investor_surname': investor['investor_surname'],
                        'investor_name': investor['investor_name'],
                        'investor_acc_number': investor['investor_acc_number'],
                    }
                    for item in trust_item:
                        insert[item] = trust_item[item]
                    trust_list.append(insert)

        # Get Opportunities and Manipulate accordingly

        # for opportunity in opportunities_list:
        #     opportunity['id'] = str(opportunity['_id'])
        #     del opportunity['_id']

        # using list comprehension to filter out opportunities that are not in the request['Category'] list
        # opportunities_list = [opportunity for opportunity in opportunities_list if
        #                       opportunity['Category'] in request['Category']]

        for opportunity in opportunities_list:
            opportunity['id'] = str(opportunity['_id'])
            del opportunity['_id']
            if opportunity["opportunity_final_transfer_date"] != "":
                tfr_date = datetime.strptime(opportunity["opportunity_final_transfer_date"].replace("-", "/"),
                                             "%Y/%m/%d")
                report_date = datetime.strptime(request['date'].replace("-", "/"), "%Y/%m/%d")
                # print("tfr_date", tfr_date)
                # print("report_date", report_date)
                if tfr_date > report_date:
                    opportunity["opportunity_transferred"] = False
                else:
                    opportunity["opportunity_transferred"] = True
            else:
                opportunity["opportunity_transferred"] = False
                if 'opportunity_end_date' in opportunity:
                    if opportunity['opportunity_end_date'] != '':
                        opportunity['opportunity_final_transfer_date'] = opportunity['opportunity_end_date']
                    else:
                        opportunity['opportunity_final_transfer_date'] = opportunity['opportunity_occupation_date']
                else:
                    opportunity['opportunity_final_transfer_date'] = opportunity['opportunity_occupation_date']

            if not opportunity["opportunity_sold"]:
                opportunity["opportunity_transferred"] = False

            # if opportunity['opportunity_final_transfer_date'] == '':
            #     # if  opportunity['opportunity_end_date'] != '' or does not exist
            #     if 'opportunity_end_date' in opportunity:
            #         if opportunity['opportunity_end_date'] != '':
            #             opportunity['opportunity_final_transfer_date'] = opportunity['opportunity_end_date']
            #         else:
            #             opportunity['opportunity_final_transfer_date'] = opportunity['opportunity_occupation_date']
            #     else:
            #         opportunity['opportunity_final_transfer_date'] = opportunity['opportunity_occupation_date']

        # Get Sales Parameters and Manipulate accordingly

        for sales_parameter in sales_parameters_list:
            sales_parameter['id'] = str(sales_parameter['_id'])
            del sales_parameter['_id']
        # Using list comprehension to filter out sales parameters that are not in the request['Category'] list
        # sales_parameters_list = [sales_parameter for sales_parameter in sales_parameters_list if
        #                          sales_parameter['Development'] in request['Category']]

        # Get Rollovers and Manipulate accordingly

        for rollover in rollovers_list:
            rollover['id'] = str(rollover['_id'])
            del rollover['_id']
        # Using list comprehension to filter out rollovers that are not in the request['Category'] list
        # rollovers_list = [rollover for rollover in rollovers_list if
        #                   rollover['Category'] in request['Category']]

        # Get Rates and Manipulate accordingly

        for rate in rates_list:
            del rate['_id']
        # replace '/' with '-' in Efective_date
        for rate in rates_list:
            rate['Efective_date'] = rate['Efective_date'].replace('/', '-')
            rate['rate'] = float(rate['rate'])
        # sort rates by Efective_date converted to datetime in descending order
        rates_list = sorted(rates_list, key=lambda k: datetime.strptime(k['Efective_date'], '%Y-%m-%d'), reverse=True)

        # Produce an interim list of investors and investor details
        interim_investors_list = []

        for trust in trust_list:
            # if trust['investor_acc_number'] = 'ZCON01' then print trust

            # If trust['project_interest_rate'] does not exist, create it and set it to 0.00
            if 'project_interest_rate' not in trust:
                trust['project_interest_rate'] = 0.00
            # If trust['planned_release_date'] does not exist, create it and set it to ""
            if 'planned_release_date' not in trust:
                trust['planned_release_date'] = ""
            # If trust['investment_interest_rate'] does not exist, create it and set it to 0.00
            if 'investment_interest_rate' not in trust:
                trust['investment_interest_rate'] = 15.00

            insert = {"investor_surname": trust["investor_surname"], "investor_name": trust["investor_name"],
                      "investor_acc_number": trust["investor_acc_number"],
                      "investment_amount": float(trust["investment_amount"]), "deposit_date": trust["deposit_date"],
                      "release_date": trust["release_date"], "opportunity_code": trust["opportunity_code"],
                      "Category": trust["Category"], "investment_number": trust["investment_number"],
                      "project_interest_rate": trust["project_interest_rate"],
                      "planned_release_date": trust["planned_release_date"],
                      "trust_interest": 0.00,
                      "investment_interest_rate": trust['investment_interest_rate'], "investment_end_date": "",
                      }
            # Filter investments by investor_acc_number, opportunity_code and investment_number

            for investment in investments:
                # if investment['investment_number'] does not exist, create it and set it to 0
                if 'investment_number' not in investment:
                    investment['investment_number'] = 0

            final_investment = [investment for investment in investments if
                                investment['investor_acc_number'] == trust['investor_acc_number'] and
                                investment['opportunity_code'] == trust['opportunity_code']
                                and
                                investment['investment_number'] == trust['investment_number']
                                ]

            if len(final_investment):
                insert['investment_interest_rate'] = final_investment[0]["investment_interest_rate"]
                insert['investment_end_date'] = final_investment[0]["end_date"]
                # if final_investment[0]['early_release'] exists, then set insert['early_release'] to equal
                # final_investment[0]['early_release'] else set it to False
                if 'early_release' in final_investment[0]:
                    insert['early_release'] = final_investment[0]['early_release']
                else:
                    insert['early_release'] = False

            interim_investors_list.append(insert)

        # in interim_investors_list, in deposit_date, replace '-' with '/' using list comprehension

        interim_investors_list = [investor for investor in interim_investors_list if
                                  str(investor['deposit_date']).replace('/', '-')]

        # replace in request['date'] '/' with '-'

        request['date'] = request['date'].replace('-', '/')

        for investor in interim_investors_list:
            investor['deposit_date'] = str(investor['deposit_date'])
            investor['deposit_date'] = investor['deposit_date'].replace('-', '/')

        # Filter interim investor list where deposit_date as date is less than or equal to the request['date'] as date
        interim_investors_list = [investor for investor in interim_investors_list if
                                  datetime.strptime(investor['deposit_date'], '%Y/%m/%d') <=
                                  datetime.strptime(request['date'], '%Y/%m/%d')]

        final_investors_list = []

        for item in opportunities_list:

            # Create new list from interim_investors_list where opportunity_code matches item['opportunity_code']
            # using list comprehension
            interim_investors = [investor for investor in interim_investors_list if
                                 investor['opportunity_code'] == item['opportunity_code']]

            item['opportunity_sale_price'] = item[
                'opportunity_sale_price'] if 'opportunity_sale_price' in item else 0
            if len(interim_investors) == 0:
                # if item['opportunity_end_date'] exits then set it to item['opportunity_end_date'] else set it to ""
                item['opportunity_end_date'] = item['opportunity_end_date'] if 'opportunity_end_date' in item else ""

                insert = {"investor_surname": "UnAllocated", "investor_name": "",
                          "investor_acc_number": "ZZUN01",
                          "investment_amount": 0, "deposit_date": "",
                          "release_date": "", "opportunity_code": item['opportunity_code'],
                          "investment_number": 0,
                          "project_interest_rate": trust["project_interest_rate"],
                          "planned_release_date": trust["planned_release_date"],
                          "trust_interest": 0.00,
                          "investment_interest_rate": trust["project_interest_rate"], "investment_end_date": "",
                          "Category": item['Category'],
                          "opportunity_sold": item['opportunity_sold'],
                          "opportunity_end_date": item['opportunity_end_date'],
                          "opportunity_final_transfer_date": item['opportunity_final_transfer_date'],
                          "opportunity_amount_required": float(item['opportunity_amount_required']),
                          "opportunity_sale_price": float(item['opportunity_sale_price']),
                          "investment_interest_today": 0, "released_interest_today": 0, "trust_interest_total": 0,
                          "released_interest_total": 0,
                          "opportunity_transferred": item["opportunity_transferred"],
                          "interest_to_date_still_to_be_raised": 0, "interest_total_still_to_be_raised": 0,
                          'early_release': False
                          }

                final_investors_list.append(insert)
            else:
                # if item['opportunity_end_date'] exits then set it to item['opportunity_end_date'] else set it to ""
                item['opportunity_end_date'] = item['opportunity_end_date'] if 'opportunity_end_date' in item else ""

                # if item['opportunity_final_transfer_date'] exits then set it to item[
                # 'opportunity_final_transfer_date'] else set it to item['opportunity_end_date']
                item['opportunity_final_transfer_date'] = item[
                    'opportunity_final_transfer_date'] if 'opportunity_final_transfer_date' in item else item[
                    'opportunity_end_date']

                for invest in interim_investors:
                    invest["opportunity_sold"] = item['opportunity_sold']
                    invest["opportunity_end_date"] = item['opportunity_end_date']
                    invest["opportunity_final_transfer_date"] = item['opportunity_final_transfer_date']
                    invest["opportunity_amount_required"] = float(item['opportunity_amount_required'])
                    invest["opportunity_sale_price"] = float(item['opportunity_sale_price'])
                    invest['opportunity_transferred'] = item["opportunity_transferred"]
                    invest["interest_to_date_still_to_be_raised"] = 0
                    invest["interest_total_still_to_be_raised"] = 0

                    final_investors_list.append(invest)

        # Convert Rates_list Efective_date to datetime

        for rate in rates_list:
            rate['Efective_date'] = datetime.strptime(rate['Efective_date'].replace('/', '-'), '%Y-%m-%d')
        # order rates_list by Efective_date in descending order
        rates_list = sorted(rates_list, key=lambda k: k['Efective_date'], reverse=True)

        for index, investment in enumerate(final_investors_list):
            # If investment["investment_interest_rate"] is an empty string then set it to 15.5
            investment["investment_interest_rate"] = investment["investment_interest_rate"] \
                if investment["investment_interest_rate"] != "" else 15.5

            # if investment["opportunity_code"] is equal to "HFB313" and the investment['investor_acc_number'] is
            # equal to "ZCON01" then set investment["release_date"] equal to investment['deposit_date']
            if investment["opportunity_code"] == "HFB313" and investment['investor_acc_number'] == "ZCON01":
                investment["release_date"] = investment['deposit_date']

            for sales_parameter in sales_parameters_list:
                for item in sales_parameter:
                    if item == "Description":
                        investment[sales_parameter[item]] = sales_parameter['rate']

            # If investment["opportunity_final_transfer_date"] if investment['opportunity_end_date'] does not exist
            # then set it to investment['opportunity_occupation_date']
            investment['opportunity_end_date'] = investment[
                'opportunity_end_date'] if 'opportunity_end_date' in investment else investment[
                'opportunity_occupation_date']

            if investment["opportunity_final_transfer_date"] == "":
                if investment["investment_end_date"] != "":
                    investment["opportunity_final_transfer_date"] = investment["investment_end_date"]
                else:
                    investment["opportunity_final_transfer_date"] = investment["opportunity_occupation_date"]

            # if investment["release_date"] == "" then investment["planned_release_date"] = deposit_date + 30 days
            # else investment["planned_release_date"] = release_date
            if investment["release_date"] == "" and investment['investor_acc_number'] != "ZZUN01":
                # in investment["deposit_date"] replace '-' with '/'
                investment["deposit_date"] = investment["deposit_date"].replace('-', '/')

                # planned_release_date exists in investment then investment["planned_release_date"] = deposit_date +
                # 30 days else investment["planned_release_date"] = ""
                if "planned_release_date" in investment and investment["planned_release_date"] != "":
                    investment["planned_release_date"] = \
                        str(datetime.strptime(investment["planned_release_date"].replace('-', '/'),
                                              '%Y/%m/%d')).split(" ")[0]
                else:

                    investment["planned_release_date"] = \
                        str(datetime.strptime(investment["deposit_date"].replace('-', '/'),
                                              '%Y/%m/%d') + timedelta(days=30)).split(
                            " ")[
                            0]
            else:
                investment["planned_release_date"] = investment["release_date"]
                # investment["planned_release_date"] = ""

            if investment["investor_acc_number"] != "ZZUN01":
                deposit_date = investment["deposit_date"].replace('-', '/')
                deposit_date = deposit_date.split(" ")[0]
                deposit_date = datetime.strptime(deposit_date, '%Y/%m/%d')

                # if investment["investment_end_date"] != "" replace '/' with '-' in investment[
                # "investment_end_date"] and do the exact same for investment["opportunity_final_transfer_date"]

                if investment["investment_end_date"] != "":
                    investment["investment_end_date"] = investment["investment_end_date"].replace('/', '-')
                    investment["opportunity_final_transfer_date"] = investment[
                        "opportunity_final_transfer_date"].replace(
                        '/', '-')

                planned_release_date = investment["planned_release_date"].replace('-', '/')
                planned_release_date = planned_release_date.split(" ")[0]
                planned_release_date = datetime.strptime(planned_release_date, '%Y/%m/%d')

                opportunity_final_transfer_date = investment["opportunity_final_transfer_date"].replace('-', '/')
                opportunity_final_transfer_date = opportunity_final_transfer_date.split(" ")[0]
                opportunity_final_transfer_date = datetime.strptime(opportunity_final_transfer_date, '%Y/%m/%d')

                # Add a day to deposit_date
                deposit_date = deposit_date + timedelta(days=1)

                investment_interest_total = 0
                released_interest_total = 0

                while deposit_date <= planned_release_date:
                    # Filter rates_list where Efective_date is less than or equal to deposit_date
                    # using list comprehension
                    interim_rate = float(
                        [rate for rate in rates_list if rate['Efective_date'] <= deposit_date][0]['rate'])
                    investment_interest_total = investment_interest_total + (
                            float(investment["investment_amount"]) * (interim_rate - 2.75) / 100 / 365)
                    deposit_date = deposit_date + timedelta(days=1)
                investment["trust_interest_total"] = round(investment_interest_total, 2)

                # Days between planned_release_date and opportunity_final_transfer_date
                days_between = (opportunity_final_transfer_date - planned_release_date).days

                released_interest_total = released_interest_total + (float(investment["investment_amount"]) * float(
                    investment["investment_interest_rate"]) / 100 / 365 * days_between)
                investment["released_interest_total"] = round(released_interest_total, 2)

                investment_interest_today = 0
                released_interest_today = 0

                # Days between deposit_date and today
                # DO TO DATE CALCS
                deposit_date = investment["deposit_date"].replace('-', '/')
                deposit_date = deposit_date.split(" ")[0]
                deposit_date = datetime.strptime(deposit_date, '%Y/%m/%d')
                planned_release_date = investment["planned_release_date"].replace('-', '/')
                planned_release_date = planned_release_date.split(" ")[0]
                planned_release_date = datetime.strptime(planned_release_date, '%Y/%m/%d')

                opportunity_final_transfer_date = investment["opportunity_final_transfer_date"].replace('-', '/')
                opportunity_final_transfer_date = opportunity_final_transfer_date.split(" ")[0]
                opportunity_final_transfer_date = datetime.strptime(opportunity_final_transfer_date, '%Y/%m/%d')

                today_released = datetime.strptime(request['date'], '%Y/%m/%d')
                today_transfer = datetime.strptime(request['date'], '%Y/%m/%d')
                if planned_release_date < today_released:
                    today_released = planned_release_date
                if opportunity_final_transfer_date < today_transfer:
                    today_transfer = opportunity_final_transfer_date

                # Add a day to deposit_date
                deposit_date = deposit_date + timedelta(days=1)

                while deposit_date <= today_released:
                    # Filter rates_list where Efective_date is less than or equal to deposit_date
                    # using list comprehension

                    interim_rate = float(
                        [rate for rate in rates_list if rate['Efective_date'] <= deposit_date][0]['rate'])

                    investment_interest_today = investment_interest_today + (
                            float(investment["investment_amount"]) * interim_rate / 100 / 365)

                    deposit_date = deposit_date + timedelta(days=1)
                investment["investment_interest_today"] = round(investment_interest_today, 2)

                # Days between planned_release_date and opportunity_final_transfer_date
                days_between = (today_transfer - planned_release_date).days

                released_interest_today = released_interest_today + (float(investment["investment_amount"]) * float(
                    investment["investment_interest_rate"]) / 100 / 365 * days_between)
                investment["released_interest_today"] = round(released_interest_today, 2)

        for investment in final_investors_list:
            filtered_rollovers = [rollover for rollover in rollovers_list if
                                  rollover['investor_acc_number'] == investment['investor_acc_number']
                                  and rollover['investment_number'] == investment['investment_number']
                                  and rollover['opportunity_code'] == investment['opportunity_code']]

            if len(filtered_rollovers) > 0:
                investment['rollover_amount'] = filtered_rollovers[0]['rollover_amount']
                investment['rollover_date'] = filtered_rollovers[0]['end_date']
            else:
                investment['rollover_amount'] = 0
                investment['rollover_date'] = ""

        # LOOP THROUGH AND FILL OPPORTUNITIES
        report_date = datetime.strptime(request['date'], '%Y/%m/%d')

        for opportunity in opportunities_list:

            opportunity['opportunity_end_date'] = opportunity['opportunity_end_date'].replace('/', '-')
            # opportunity['opportunity_end_date'] = opportunity['opportunity_end_date'].split(" ")[0]
            opportunity['opportunity_end_date'] = datetime.strptime(opportunity['opportunity_end_date'], '%Y-%m-%d')
            # if opportunity_final_transfer_date in opportunity and if it is not empty then do as above
            if opportunity['opportunity_final_transfer_date'] != "":
                opportunity['opportunity_final_transfer_date'] = opportunity['opportunity_final_transfer_date'].replace(
                    '/', '-')
                # opportunity['opportunity_final_transfer_date'] = opportunity['opportunity_final_transfer_date'].split(
                #     " ")[0]
                opportunity['opportunity_final_transfer_date'] = datetime.strptime(
                    opportunity['opportunity_final_transfer_date'], '%Y-%m-%d')

                if opportunity['opportunity_end_date'] > report_date or opportunity[
                    'opportunity_final_transfer_date'] > report_date:
                    opportunity['opportunity_transferred'] = False

            # Filter final_investors_list where opportunity_code is equal to opportunity['opportunity_code']
            # using list comprehension
            filtered_investors = [investor for investor in final_investors_list if
                                  investor['opportunity_code'] == opportunity['opportunity_code']]

            opportunity_required = float(opportunity['opportunity_amount_required'])

            # sum the investment_amounts in filtered_investors list using list comprehension
            opportunity_invested = sum(float(investor['investment_amount']) for investor in filtered_investors)
            if 0 < opportunity_invested < opportunity_required:
                insert = {"investor_surname": "UnAllocated", "investor_name": "",
                          "investor_acc_number": "ZZUN01",
                          "investment_amount": 0, "deposit_date": "",
                          "release_date": "", "opportunity_code": opportunity['opportunity_code'],
                          "investment_number": 0,
                          "project_interest_rate": opportunity["opportunity_interest_rate"],
                          "planned_release_date": "",
                          "trust_interest": 0.00,
                          "investment_interest_rate": opportunity["opportunity_interest_rate"],
                          "investment_end_date": "", "Category": opportunity['Category'],
                          "opportunity_sold": opportunity['opportunity_sold'],
                          "opportunity_end_date": opportunity['opportunity_end_date'],
                          "opportunity_final_transfer_date": opportunity['opportunity_final_transfer_date'],
                          "opportunity_amount_required": float(opportunity['opportunity_amount_required']),
                          "opportunity_sale_price": float(opportunity['opportunity_sale_price']),
                          "investment_interest_today": 0, "released_interest_today": 0, "trust_interest_total": 0,
                          "released_interest_total": 0,
                          "opportunity_transferred": opportunity["opportunity_transferred"],
                          "raising_commission": 0, "structuring_fee": 0, "commission": 0, "transfer_fees": 0,
                          "bond_registration": 0, "trust_release_fee": 0, "unforseen": 0,
                          "interest_to_date_still_to_be_raised": 0, "interest_total_still_to_be_raised": 0,
                          "early_release": False
                          }

                final_investors_list.append(insert)

        for investment in final_investors_list:

            # filter sales_parameters_list where Development is equal to investment['Category'] using list comprehension

            if investment['raising_commission'] == 0:
                filtered_sales_parameters = [sales_parameter for sales_parameter in sales_parameters_list if
                                             sales_parameter['Development'] == investment['Category'] and
                                             sales_parameter['Description'] == 'raising_commission']
                investment['raising_commission'] = filtered_sales_parameters[0]['rate']
            if investment['structuring_fee'] == 0:
                filtered_sales_parameters = [sales_parameter for sales_parameter in sales_parameters_list if
                                             sales_parameter['Development'] == investment['Category'] and
                                             sales_parameter['Description'] == 'structuring_fee']
                investment['structuring_fee'] = filtered_sales_parameters[0]['rate']

            if investment['commission'] == 0:
                filtered_sales_parameters = [sales_parameter for sales_parameter in sales_parameters_list if
                                             sales_parameter['Development'] == investment['Category'] and
                                             sales_parameter['Description'] == 'commission']
                investment['commission'] = filtered_sales_parameters[0]['rate']
            if investment['transfer_fees'] == 0:
                filtered_sales_parameters = [sales_parameter for sales_parameter in sales_parameters_list if
                                             sales_parameter['Development'] == investment['Category'] and
                                             sales_parameter['Description'] == 'transfer_fees']
                investment['transfer_fees'] = filtered_sales_parameters[0]['rate']
            if investment['bond_registration'] == 0:
                filtered_sales_parameters = [sales_parameter for sales_parameter in sales_parameters_list if
                                             sales_parameter['Development'] == investment['Category'] and
                                             sales_parameter['Description'] == 'bond_registration']
                investment['bond_registration'] = filtered_sales_parameters[0]['rate']
            if investment['trust_release_fee'] == 0:
                filtered_sales_parameters = [sales_parameter for sales_parameter in sales_parameters_list if
                                             sales_parameter['Development'] == investment['Category'] and
                                             sales_parameter['Description'] == 'trust_release_fee']
                investment['trust_release_fee'] = filtered_sales_parameters[0]['rate']
            if investment['unforseen'] == 0:
                filtered_sales_parameters = [sales_parameter for sales_parameter in sales_parameters_list if
                                             sales_parameter['Development'] == investment['Category'] and
                                             sales_parameter['Description'] == 'unforseen']
                investment['unforseen'] = filtered_sales_parameters[0]['rate']

        # if the investor_acc_number = "ZCAM01" and the opportunity_code = "HFA101" and the investment_amount =
        # 400000.0 then filter this record out of final_investors_list

        final_investors_list = [investor for investor in final_investors_list if
                                not (investor['investor_acc_number'] == "ZCAM01" and investor[
                                    'opportunity_code'] == "HFA101" and investor['investment_amount'] == 400000.0)]

        # if investor_acc_number = "ZJHO01" and opportunity_code = "HFA304" and investment_number = 1 then filter this
        # record out of final_investors_list

        final_investors_list = [investor for investor in final_investors_list if
                                not (investor['investor_acc_number'] == "ZJHO01" and investor[
                                    'opportunity_code'] == "HFA304" and investor['investment_number'] == 1)]

        final_investors_list = [investor for investor in final_investors_list if
                                not (investor['investor_acc_number'] == "ZPJB01" and investor[
                                    'opportunity_code'] == "HFA205" and investor['investment_number'] == 1)]

        final_investors_list = [investor for investor in final_investors_list if
                                not (investor['investor_acc_number'] == "ZERA01" and investor[
                                    'opportunity_code'] == "EA205" and investor['investment_number'] == 3)]

        final_investors_list = [investor for investor in final_investors_list if
                                not (investor['investor_acc_number'] == "ZVOL01" and investor[
                                    'opportunity_code'] == "EA103" and investor['investment_number'] == 3)]

        # get unallocated_investments from mongo db where the request['Category'] is in the Category in the DB

        for unallocated_investment in unallocated_investments_list:
            unallocated_investment['id'] = str(unallocated_investment['_id'])
            del unallocated_investment['_id']

        # Filter unallocated_investments_list where deposit_date = "" using list comprehension
        unallocated_investments_list = [unallocated_investment for unallocated_investment in
                                        unallocated_investments_list if
                                        unallocated_investment['deposit_date'] != ""]

        for investor in final_investors_list:

            # Filter unallocated_investments_list where opportunity_code is equal to investor['opportunity_code'] and
            # investor_acc_number is equal to 'ZZUN01' using list comprehension
            if investor['investor_acc_number'] == "ZZUN01":
                filtered_unallocated_investments = [unallocated_investment for unallocated_investment in
                                                    unallocated_investments_list if
                                                    unallocated_investment['opportunity_code'] == investor[
                                                        'opportunity_code']]

                if len(filtered_unallocated_investments) > 0:

                    investor['deposit_date'] = str(filtered_unallocated_investments[0]['deposit_date'])
                    investor['release_date'] = str(filtered_unallocated_investments[0]['release_date'])
                    investor['planned_release_date'] = str(filtered_unallocated_investments[0]['release_date'])

                    # investor['project_interest_rate'] = float(filtered_unallocated_investments[0][
                    # 'project_interest_rate'])
                    investor['project_interest_rate'] = float(
                        filtered_unallocated_investments[0]['project_interest_rate'])

                    interest_to_be_raised_for_momentum = 0
                    interest_to_be_raised_for_released = 0
                    # Do Interest Calcs
                    deposit_date = investor['deposit_date'].replace("-", "/")
                    deposit_date = datetime.strptime(deposit_date, "%Y/%m/%d")

                    # int_release_date = investor['release_date'].replace("-", "/")
                    # int_release_date = datetime.strptime(int_release_date, "%Y/%m/%d")

                    int_planned_release_date = investor['planned_release_date'].replace("-", "/")
                    int_planned_release_date = datetime.strptime(int_planned_release_date, "%Y/%m/%d")

                    # convert investor['opportunity_final_transfer_date'] from a datetime to a string with format
                    # YYYY/MM/DD

                    opportunity_final_transfer_date = str(investor['opportunity_final_transfer_date']).split(" ")[0]

                    # opportunity_final_transfer_date = investor['opportunity_final_transfer_date'].replace("-", "/")
                    opportunity_final_transfer_date = datetime.strptime(opportunity_final_transfer_date, "%Y-%m-%d")

                    # add 1 day to deposit date
                    deposit_date = deposit_date + timedelta(days=1)
                    while deposit_date <= int_planned_release_date:
                        # get rate from rate_list as a float where date less than or equal to deposit_date
                        rate = float([rate for rate in rates_list if rate['Efective_date'] <= deposit_date][0]['rate'])

                        interest_to_be_raised_for_momentum += (filtered_unallocated_investments[0][
                                                                   'unallocated_investment'] * rate / 100) / 365

                        deposit_date = deposit_date + timedelta(days=1)

                    # add 1 day to release date
                    int_planned_release_date = int_planned_release_date + timedelta(days=1)
                    while int_planned_release_date <= opportunity_final_transfer_date:
                        interest_to_be_raised_for_released += (filtered_unallocated_investments[0][
                                                                   'unallocated_investment'] * investor[
                                                                   'project_interest_rate'] / 100) / 365

                        # add a day to int_planned_release_date
                        int_planned_release_date = int_planned_release_date + timedelta(days=1)

                    # investor["interest_to_date_still_to_be_raised"] = 4000
                    investor[
                        "interest_total_still_to_be_raised"] = \
                        interest_to_be_raised_for_released + interest_to_be_raised_for_momentum

            # if investor['opportunity_code'] == "HVG101" and investor['investor_acc_number'] == "ZWIL02":
            #     print(investor)

        # sort final investors list by Category, opportunity_code, investor_acc_number
        final_investors_list = sorted(final_investors_list,
                                      key=lambda k: (k['Category'], k['opportunity_code'], k['investor_acc_number']))

        for investor in final_investors_list:
            # if investor['opportunity_code'] == "HVG101" and investor['investor_acc_number'] == "ZWIL02":
            #     print(investor)

            if 'early_release' not in investor:
                investor['early_release'] = False
            # create a variable called report date and make it the same as request['date'] as a datetime
            report_date = datetime.strptime(request['date'], "%Y/%m/%d")
            if investor['early_release']:
                investor['investment_end_date'] = investor['investment_end_date'].replace("-", "/")

                # datetime then set investor['early_release'] to false
                if datetime.strptime(investor['investment_end_date'], "%Y/%m/%d") >= report_date:
                    investor['early_release'] = False
                investor['opportunity_final_transfer_date'] = investor['opportunity_final_transfer_date'].replace("-",
                                                                                                                  "/")
                if datetime.strptime(investor['opportunity_final_transfer_date'], "%Y/%m/%d") > report_date:
                    investor['opportunity_transferred'] = False

                # if investor['opportunity_transferred']:
                #     investor['early_release'] = False

            for investor in final_investors_list:

                if investor['investor_acc_number'] != "ZZUN01":
                    report_date = datetime.strptime(request['date'], "%Y/%m/%d")
                    if investor['release_date'] != "":
                        start_date_total = datetime.strptime(investor['release_date'].replace("-", "/"), "%Y/%m/%d")
                        start_date_today = datetime.strptime(investor['release_date'].replace("-", "/"), "%Y/%m/%d")
                    else:
                        start_date_total = datetime.strptime(investor['deposit_date'].replace("-", "/"),
                                                             "%Y/%m/%d") + timedelta(days=30)
                        start_date_today = None
                    if investor['investment_end_date'] != "":
                        end_date_total = datetime.strptime(investor['investment_end_date'].replace("-", "/"),
                                                           "%Y/%m/%d")
                        if end_date_total < report_date:
                            end_date_today = end_date_total
                        else:
                            end_date_today = report_date
                    elif investor['opportunity_final_transfer_date'] != "":
                        end_date_total = datetime.strptime(
                            investor['opportunity_final_transfer_date'].replace("-", "/"),
                            "%Y/%m/%d")
                        if end_date_total < report_date:
                            end_date_today = end_date_total
                        else:
                            end_date_today = report_date
                    else:
                        end_date_total = datetime.strptime(investor['opportunity_end_date'].replace("-", "/"),
                                                           "%Y/%m/%d")
                        if end_date_total < report_date:
                            end_date_today = end_date_total
                        else:
                            end_date_today = report_date
                    interest_rate = float(investor['investment_interest_rate'])
                    if start_date_today is not None:
                        released_interest_today = (end_date_today - start_date_today).days * (
                                investor['investment_amount'] * interest_rate / 100) / 365
                    else:
                        released_interest_today = 0
                    if start_date_total is not None:
                        released_interest_total = (end_date_total - start_date_total).days * (
                                investor['investment_amount'] * interest_rate / 100) / 365
                    else:
                        released_interest_total = 0
                    investor['released_interest_today'] = released_interest_today
                    investor['released_interest_total'] = released_interest_total

            for investor in final_investors_list:
                filtered_opps = [opp for opp in opportunities_list if opp['opportunity_code']
                                 == investor['opportunity_code']]

                investor["rental_marked_for_rent"] = filtered_opps[0].get("rental_marked_for_rent", False)

                investor["rental_rented_out"] = filtered_opps[0].get("rental_rented_out", False)
                investor["rental_start_date"] = filtered_opps[0].get("rental_start_date", "")
                investor["rental_end_date"] = filtered_opps[0].get("rental_end_date", "")
                investor["rental_income_to_date"] = float(filtered_opps[0].get("rental_income_to_date", 0))
                investor["rental_income_to_contract_end"] = float(
                    filtered_opps[0].get("rental_income_to_contract_end", 0))
                investor["rental_gross_amount"] = float(filtered_opps[0].get("rental_gross_amount", 0))
                investor["rental_deposit_amount"] = float(filtered_opps[0].get("rental_deposit_amount", 0))
                investor["rental_levy_amount"] = float(filtered_opps[0].get("rental_levy_amount", 0))
                investor["rental_commission"] = float(filtered_opps[0].get("rental_commission", 0))
                investor["rental_rates"] = float(filtered_opps[0].get("rental_rates", 0))
                investor["rental_other_expenses"] = float(filtered_opps[0].get("rental_other_expenses", 0))
                investor["rental_nett_amount"] = float(filtered_opps[0].get("rental_nett_amount", 0))
                if investor["rental_marked_for_rent"]:
                    investor["potential_income"] = investor["rental_nett_amount"]
                    investor["rental_start_date"] = ""
                    investor["rental_end_date"] = ""
                    investor["rental_income_to_date"] = 0
                    investor["rental_income_to_contract_end"] = 0
                    investor["rental_gross_amount"] = 0
                    investor["rental_deposit_amount"] = 0
                    investor["rental_levy_amount"] = 0
                    investor["rental_commission"] = 0
                    investor["rental_rates"] = 0
                    investor["rental_other_expenses"] = 0
                    investor["rental_nett_amount"] = 0
                else:
                    investor["potential_income"] = 0

        for investor in final_investors_list:
            if investor['release_date'] == "" and investor['planned_release_date'] != "":
                final_transfer_date = datetime.strptime(investor['opportunity_final_transfer_date'].replace("-", "/"),
                                                        "%Y/%m/%d")
                planned_release_date = datetime.strptime(investor['planned_release_date'].replace("-", "/"), "%Y/%m/%d")

                days_difference = (final_transfer_date - planned_release_date).days

                investment_amount = float(investor['investment_amount'])
                investment_interest_rate = float(investor['investment_interest_rate'])
                investment_interest = (investment_amount * investment_interest_rate / 100) / 365 * days_difference
                investor['released_interest_total'] = investment_interest

            if investor['investor_acc_number'] == "ZZUN01" and investor["investment_interest_rate"] == 0:
                investor["investment_interest_rate"] = investor['project_interest_rate']

                # end_date_total = datetime.strptime(investor['opportunity_final_transfer_date'].replace("-", "/"), "%Y/%m/%d")

        ## DEVELOPMENT UNITS
        # create a new list called development_list which includes all the records in final_investors_list where the
        # opportunity_code is equal to HB215 or HFB315
        development_list = [opportunity for opportunity in final_investors_list if
                            opportunity['opportunity_code'] == "HFB215" or opportunity['opportunity_code'] == "HFB315"]

        # print("development_list", development_list)

        # print("final_investors_list", final_investors_list[0])

        listData = investment_status(request)

        # append development_list to listData
        # listData.append(development_list)
        for devunit in development_list:
            listData.append(devunit)

        # print("listData", listData[len(listData) - 1])
        # print()
        # print("listData", listData[len(listData) - 2])

        background_tasks.add_task(create_sales_forecast_file, final_investors_list, request, pledges, firstName,
                                  listData, request)

        end = time.time()
        print("Time Taken: ", end - start)

        # print("filename", filename)

        return {"message": "The server is busy processing the data, please be patient.", "filename": f'{filename}'}
        # return {"filename": f'{filename}.xlsx'}
        # return "Time Taken: ", end - start, len(final_investors_list), final_investors_list

    except Exception as e:
        print("Error:", e)
        return {"error": e}


def investment_status(request):
    # request = await data.json()
    # print(request)

    try:
        # start = time.time()

        # developments = request['Category']

        investor_list = list(db.investors.aggregate([
            {
                "$match": {
                    "investments": {
                        "$elemMatch": {
                            "Category": {"$in": request['Category']}
                        }
                    }
                }
            },
            {
                "$addFields": {
                    "investments": {
                        "$filter": {
                            "input": "$investments",
                            "as": "investment",
                            "cond": {
                                "$in": ["$$investment.Category", request['Category']]
                            }
                        }
                    }
                }
            },
            {
                "$match": {
                    "investments": {"$ne": []}
                }
            }
        ]))

        # Get opportunities from the database where Category is included in request['Category']
        opportunities_list = list(db.opportunities.find({"Category": {"$in": request['Category']}}))

        # print("opportunities_list", opportunities_list[0])

        # using list comprehension filter opportunities_list where opportunity['opportunity_code'] is equal to HB215 or HFB315

        # development_list = [opportunity for opportunity in opportunities_list if opportunity['opportunity_code'] == "HFB215" or opportunity['opportunity_code'] == "HFB315"]

        # print("development_list", development_list)

        final_investors_list = []

        report_date = datetime.strptime(request['date'], "%Y/%m/%d")

        for investor in investor_list:
            investor['id'] = str(investor['_id'])
            del investor['_id']
            # del investor['trust']
            del investor['pledges']
            # filter investor['investments'] where investments['Category'] is included in request['Category']
            investor['investments'] = [investment for investment in investor['investments'] if
                                       investment['Category'] in request['Category']]
            for investment in investor['investments']:
                # if investment['end_date'] != "":
                #     end_date = datetime.strptime(investment['end_date'].replace("-", "/"), "%Y/%m/%d")
                #     if end_date > report_date:
                #         investment['end_date'] = ""
                # if investor['investor_acc_number'] == "ZKRU01":
                #     print(investment)

                insert = {}
                # if  investment['investment_number'] exists then set insert['investment_number'] to
                # investment['investment_number'] else set insert['investment_number'] to 0
                if 'investment_number' in investment:
                    insert['investment_number'] = investment['investment_number']
                else:
                    insert['investment_number'] = 0

                if 'early_release' in investment:
                    insert['early_release'] = investment['early_release']
                else:
                    insert['early_release'] = False

                insert['investment_name'] = investor['investment_name']
                insert['investor_acc_number'] = investor['investor_acc_number']
                # do the same as above for Category, opportunity_code, investment_amount, release_date and end_date
                insert['Category'] = investment['Category']
                insert['opportunity_code'] = investment['opportunity_code']
                insert['investment_amount'] = investment['investment_amount']
                insert['release_date'] = investment['release_date']
                insert['end_date'] = investment['end_date']
                insert['investment_interest_rate'] = float(investment['investment_interest_rate'])
                if 'exit_value' in investment:
                    insert['exit_value'] = float(investment['exit_value'].replace(" ", ""))
                else:
                    insert['exit_value'] = 0.00

                # using list comprehension filter investor['trust'] where trust['opportunity_code'] is equal to
                # investment['opportunity_code'] and trust['investment_number'] is equal to investment[
                # 'investment_number'] and trust['Category'] is equal to investment['Category']
                trust_filtered = [trust for trust in investor['trust'] if
                                  trust['opportunity_code'] == investment['opportunity_code'] and
                                  trust['investment_number'] == insert['investment_number']]

                if len(trust_filtered) == 0:
                    insert['deposit_date'] = None

                else:
                    insert['deposit_date'] = trust_filtered[0]['deposit_date']

                final_investors_list.append(insert)

        # filter out where investor_acc_number = "ZCAM01" and investment_number = 0
        final_investors_list = [investment for investment in final_investors_list if
                                not (investment['investor_acc_number'] == "ZCAM01" and investment[
                                    'investment_number'] == 0)]

        # filter out where investor_acc_number = "ZJHO01" and "investment_number": 1
        final_investors_list = [investment for investment in final_investors_list if
                                not (investment['investor_acc_number'] == "ZJHO01" and investment[
                                    'investment_number'] == 1)]

        final_investors_list = [investor for investor in final_investors_list if
                                not (investor['investor_acc_number'] == "ZPJB01" and investor[
                                    'opportunity_code'] == "HFA205" and investor['investment_number'] == 1)]

        final_investors_list = [investor for investor in final_investors_list if
                                not (investor['investor_acc_number'] == "ZERA01" and investor[
                                    'opportunity_code'] == "EA205" and investor['investment_number'] == 3)]

        final_investors_list = [investor for investor in final_investors_list if
                                not (investor['investor_acc_number'] == "ZVOL01" and investor[
                                    'opportunity_code'] == "EA103" and investor['investment_number'] == 3)]
        # print(final_investors_list)
        for investment in final_investors_list:
            # filter opportunities_list where opportunity['opportunity_code'] is equal to investment['opportunity_code']

            opportunity_filtered = [opportunity for opportunity in opportunities_list if
                                    opportunity['opportunity_code'] == investment['opportunity_code']][0]
            # print("opportunity_filtered", opportunity_filtered)

            investment['occupation_date'] = opportunity_filtered['opportunity_occupation_date']
            investment['estimated_transfer_date'] = opportunity_filtered['opportunity_end_date']
            investment['final_transfer_date'] = opportunity_filtered['opportunity_final_transfer_date']
            investment['opportunity_sold'] = opportunity_filtered['opportunity_sold']

        # convert request['date'] to a datetime
        report_date = datetime.strptime(request['date'], "%Y/%m/%d")

        # filter final_investors_list where investment['release_date'] as a dateTime is greater than or equal to
        # report_date
        final_investors_list = [investment for investment in final_investors_list if
                                datetime.strptime(investment['release_date'].replace("-", "/"),
                                                  "%Y/%m/%d") <= report_date]

        rates = list(db.rates.find({}))

        for rate in rates:
            rate['rate'] = float(rate['rate'])
            rate['Efective_date'] = datetime.strptime(rate['Efective_date'].replace("-", "/"), "%Y/%m/%d")

        # sort rates by Efective_date descending
        rates = sorted(rates, key=lambda k: k['Efective_date'], reverse=True)
        # print("rates", rates)

        for investor in final_investors_list:
            investor['report_date'] = report_date.strftime("%Y/%m/%d")
            # make investor['block'] equal to the 4th last character of investor['opportunity_code']
            investor['block'] = investor['opportunity_code'][-4:3]

            if investor['final_transfer_date'] == "":
                exit_deadline = datetime.strptime(investor['release_date'].replace("-", "/"), "%Y/%m/%d")
                # add 730 days to exit_deadline
                exit_deadline = exit_deadline + timedelta(days=730)
                # print("report_date", report_date)
                days_to_exit_deadline = exit_deadline - report_date
                investor['days_to_exit_deadline'] = days_to_exit_deadline.days
            else:
                if datetime.strptime(investor['final_transfer_date'].replace("-", "/"), "%Y/%m/%d") > report_date:
                    investor['final_transfer_date'] = ""
                    exit_deadline = datetime.strptime(investor['release_date'].replace("-", "/"), "%Y/%m/%d")
                    # add 730 days to exit_deadline
                    exit_deadline = exit_deadline + timedelta(days=730)
                    # print("report_date", report_date)
                    days_to_exit_deadline = exit_deadline - report_date
                    investor['days_to_exit_deadline'] = days_to_exit_deadline.days
                else:
                    investor['days_to_exit_deadline'] = 0

            if (investor["final_transfer_date"] is None or investor["final_transfer_date"] == "") and (
                    investor["end_date"] == "" or investor["end_date"] is None):
                investment_interest = 0
                # released_interest = 0
                amount_invested = float(investor['investment_amount'])
                released_interest_rate = float(investor['investment_interest_rate'])
                start_investment_interest_date = datetime.strptime(investor['deposit_date'].replace("-", "/"),
                                                                   "%Y/%m/%d")
                # add 1 day to start_investment_interest_date
                start_investment_interest_date = start_investment_interest_date + timedelta(days=1)
                end_investment_interest_date = datetime.strptime(investor['release_date'].replace("-", "/"), "%Y/%m/%d")
                start_released_interest_date = datetime.strptime(investor['release_date'].replace("-", "/"), "%Y/%m/%d")
                # add 1 day to start_released_interest_date

                end_released_interest_date = datetime.strptime(investor['estimated_transfer_date'].replace("-", "/"),
                                                               "%Y/%m/%d")
                days_released_interest = (end_released_interest_date - start_released_interest_date).days
                while start_investment_interest_date <= end_investment_interest_date:
                    # filter rates where rate['Efective_date'] is equal to start_investment_interest_date
                    rate_filtered = [rate for rate in rates if rate['Efective_date'] <= start_investment_interest_date]
                    rate = rate_filtered[0]['rate']
                    investment_interest += amount_invested * rate / 100 / 365
                    # add 1 day to start_investment_interest_date
                    start_investment_interest_date = start_investment_interest_date + timedelta(days=1)

                released_interest = amount_invested * released_interest_rate / 100 * days_released_interest / 365

                investor['investment_interest'] = amount_invested + investment_interest + released_interest
                investor['exited_by_developer'] = 0
                investor['date_of_exit'] = investor['estimated_transfer_date']

            elif investor["end_date"] != "" and investor["end_date"] is not None and investor[
                "final_transfer_date"] == "":

                end_date = datetime.strptime(investor['end_date'].replace("-", "/"),
                                             "%Y/%m/%d")

                if investor['final_transfer_date'] == "":
                    early_release_date = datetime.strptime(investor['estimated_transfer_date'].replace("-", "/"),
                                                           "%Y/%m/%d")
                else:
                    early_release_date = datetime.strptime(investor['final_transfer_date'].replace("-", "/"),
                                                           "%Y/%m/%d")

                if end_date < early_release_date:

                    investor['exited_by_developer'] = investor['exit_value']
                    investor['investment_interest'] = 0
                    investor['date_of_exit'] = investor['end_date']

                else:
                    # investor['test'] = 'Test'
                    investor['exited_by_developer'] = 0
                    investor['investment_interest'] = 0
                    investor['date_of_exit'] = ""

            else:
                investor['investment_interest'] = 0
                investor['exited_by_developer'] = 0
                investor['date_of_exit'] = ""

        # using list comprehension loop through final_investors_list and append investment['opportunity_code'] to
        # count_of_units if investment['final_transfer_date'] is equal to and then only keep unique values and append
        # to count_of_units
        # count_of_units = [investment['opportunity_code'] for investment in final_investors_list if
        #                   investment['final_transfer_date'] == investor['final_transfer_date']]
        # count_of_units = len(list(set(count_of_units)))
        # print("count_of_units", count_of_units)

        # print("final_investors_list", final_investors_list[0])

        count_of_units = []
        for investment in final_investors_list:
            if investment['final_transfer_date'] == "":
                # print(investment)
                # print()
                count_of_units.append(investment['opportunity_code'])

        count_of_units = len(list(set(count_of_units)))
        # print("count_of_units", count_of_units)

        # sort final investors list by Category, opportunity_code, investor_acc_number
        final_investors_list = sorted(final_investors_list,
                                      key=lambda k: (k['Category'], k['opportunity_code'], k['investor_acc_number']))

        # if row[9].value != "":
        #     reportDate = datetime.strptime(request['date'], '%Y/%m/%d')
        #     exitDate = datetime.strptime(row[9].value, '%Y-%m-%d')
        #     print("reportDate", reportDate)
        #     print("exitDate", exitDate)
        #     if exitDate > reportDate:
        #         row[9].value = ""

        # print("final_investors_list", final_investors_list[0])
        # for investment in final_investors_list:

        # add count_of_units to final_investors_list[0]
        final_investors_list[0]['count_of_units'] = count_of_units

        # end = time.time()

        # return {"done": True, "time_taken": end - start}
        return final_investors_list
    except Exception as e:
        print("Error:", e)
        return []


@excel_sales_forecast.post("/create_cashflow")
async def create_cashflow(data: Request):
    request = await data.json()

    try:
        start = time.time()

        if len(request['Category']) > 1:
            filename = f"excel_files/Cashflow Heron.xlsx"
        else:
            filename = f"excel_files/Cashflow {request['Category'][0]}.xlsx"

        if os.path.exists(filename):
            os.remove(filename)
            print("File Removed!")
            # return {"filename": filename}
        else:
            print("File Does NOT Exist")
        # print("request", request)
        investor_list = list(db.investors.aggregate([
            {
                "$match": {
                    "investments": {
                        "$elemMatch": {
                            "Category": {"$in": request['Category']}
                        }
                    },
                    "trust": {
                        "$elemMatch": {
                            "Category": {"$in": request['Category']}
                        }
                    }
                }
            },
            {
                "$addFields": {
                    "investments": {
                        "$filter": {
                            "input": "$investments",
                            "as": "investment",
                            "cond": {
                                "$in": ["$$investment.Category", request['Category']]
                            }
                        }
                    },
                    "trust": {
                        "$filter": {
                            "input": "$trust",
                            "as": "trust",
                            "cond": {
                                "$in": ["$$trust.Category", request['Category']]
                            }
                        }
                    }
                }
            },
            {
                "$match": {
                    "investments": {"$ne": []}
                }
            }
        ]))

        sales_parameter_list = list(db.salesParameters.find({"Development": {"$in": request['Category']}}))

        opportunities_list = list(db.opportunities.find({"Category": {"$in": request['Category']}}))
        # filter out where opportunity_final_transfer_date is not equal to ""
        opportunities_list = list(
            filter(lambda opportunity_toGet: opportunity_toGet['opportunity_final_transfer_date'] == "",
                   opportunities_list))

        for opportunity in opportunities_list:
            opportunity['id'] = str(opportunity['_id'])
            del opportunity['_id']
            for sale in sales_parameter_list:
                opportunity[sale['Description']] = sale['rate']
            opportunity['nett'] = float(opportunity['opportunity_sale_price']) - \
                                  (((float(opportunity['opportunity_sale_price']) / 1.15) * opportunity['commission']) +
                                   ((float(opportunity['opportunity_sale_price'])) * opportunity['unforseen']) +
                                   opportunity['transfer_fees'] + opportunity['bond_registration'] +
                                   opportunity['trust_release_fee'])

        rates = list(db.rates.find({}))
        for rate in rates:
            rate['Efective_date'] = datetime.strptime(rate['Efective_date'].replace("-", "/"), "%Y/%m/%d")
            rate['rate'] = float(rate['rate'])
        # sort rates by Efective_date desc
        rates = sorted(rates, key=lambda k: k['Efective_date'], reverse=True)

        investments = []
        # print("investor_list", investor_list[0]['investments'])
        for investment in investor_list:
            for inv in investment['investments']:
                if 'investment_number' not in inv:
                    inv['investment_number'] = 0
                    inv['rollover_amount'] = 0

                # if inv['opportunity_code'] == "HFB214" or inv['opportunity_code'] == "HVD201":
                #     print("inv", inv)
                #     print()
                if inv['end_date'] == "":

                    trust = list(filter(lambda trust: trust['opportunity_code'] == inv['opportunity_code'] and
                                                      trust['investment_number'] == inv['investment_number'],
                                        investment['trust']))
                    # create variable from opportunity_list where opportunity_code is equal to inv['opportunity_code']
                    opportunity = list(
                        filter(lambda opportunity: opportunity['opportunity_code'] == inv['opportunity_code'],
                               opportunities_list))
                    if len(trust) > 0 and len(opportunity) > 0:
                        inv['deposit_date'] = trust[0]['deposit_date']
                        inv['interest_end_date'] = opportunity[0]['opportunity_end_date']
                        investment_amount = float(inv['investment_amount'])
                        deposit_date = datetime.strptime(inv['deposit_date'].replace("-", "/"), "%Y/%m/%d")
                        # add one day to deposit_date
                        deposit_date = deposit_date + timedelta(days=1)
                        release_date = datetime.strptime(inv['release_date'].replace("-", "/"), "%Y/%m/%d")
                        interest_end_date = datetime.strptime(inv['interest_end_date'].replace("-", "/"), "%Y/%m/%d")
                        # one day to interest_end_date
                        interest_end_date = interest_end_date + timedelta(days=1)
                        momentum_interest = 0
                        released_interest = 0
                        while deposit_date <= release_date:
                            rate = list(filter(lambda rate: rate['Efective_date'] <= deposit_date, rates))[0]['rate']
                            momentum_interest += investment_amount * rate / 100 / 365
                            # add one day to deposit_date
                            deposit_date = deposit_date + timedelta(days=1)

                        investment_interest_rate = float(inv['investment_interest_rate'])
                        days = (interest_end_date - deposit_date).days
                        released_interest = investment_amount * investment_interest_rate / 100 / 365 * days
                        inv['final_amount_due_to_investor'] = momentum_interest + released_interest + investment_amount

                        investments.append(inv)

        opportunity_end_dates = []

        for opportunity in opportunities_list:
            # Filter investments where opportunity_code is equal to opportunity['opportunity_code']
            opportunity_investments = list(
                filter(lambda investment: investment['opportunity_code'] == opportunity['opportunity_code'],
                       investments))

            opportunity['total_owed_to_investors'] = sum(
                [investment['final_amount_due_to_investor'] for investment in opportunity_investments]) if len(
                opportunity_investments) > 0 else 0

            opportunity['rollover_amount'] = sum(
                [investment.get('rollover_amount', 0) for investment in opportunity_investments]) if len(
                opportunity_investments) > 0 else 0
            # if opportunity['opportunity_code'] == 'HFB214':
            # opportunity['total_investments'] = round(opportunity['total_investments'], 2)
            opportunity['nett_cashflow'] = opportunity['nett'] - opportunity['total_owed_to_investors']
            # if opportunity['opportunity_code'] == 'HVD201' or opportunity['opportunity_code'] == 'HFB214':
            #     print("opportunity['nett_cashflow']", opportunity)
            #     print()
            # convert opportunity['opportunity_end_date'] to datetime and add 5 days
            opportunity['opportunity_end_date'] = datetime.strptime(
                opportunity['opportunity_end_date'].replace("-", "/"),
                "%Y/%m/%d") + timedelta(days=5)
            # opportunity['rollover_amount'] = 0
            # convert opportunity['opportunity_end_date'] to string showing YYYY/MM/DD
            # opportunity['opportunity_end_date'] = opportunity['opportunity_end_date'].strftime("%Y/%m/%d")
            opportunity_end_dates.append(opportunity['opportunity_end_date'])
            # opportunity_end_dates.append(opportunity['nett_cashflow'])

        # print("XXX",opportunity_end_dates[0])
        # only keep unique opportunity_end_dates
        opportunity_end_dates = list(set(opportunity_end_dates))
        # sort opportunity_end_dates
        opportunity_end_dates = sorted(opportunity_end_dates, key=lambda k: k)
        # print("opportunity_end_dates", opportunity_end_dates)

        # sort opportunities_list by opportunity_end_date then by opportunity_code
        opportunities_list = sorted(opportunities_list,
                                    key=lambda k: (k['opportunity_end_date'], k['opportunity_code']))

        final_cashFlow_list = []
        for date in opportunity_end_dates:
            # change date to a datetime object
            date = datetime.strptime(date.strftime("%Y/%m/%d"), "%Y/%m/%d")
            cashFlow = {}
            cashFlow['date'] = date
            cashFlow['total_cashflow'] = sum([opportunity['nett_cashflow'] for opportunity in opportunities_list if
                                              opportunity['opportunity_end_date'] == date])
            cashFlow['rollover_amount'] = sum([opportunity['rollover_amount'] for opportunity in opportunities_list if
                                               opportunity['opportunity_end_date'] == date])
            cashFlow['rollover_date'] = date + timedelta(days=7)
            # filter opportunities_list where opportunity_end_date is equal to date
            filtered_opportunities = list(
                filter(lambda opportunity: opportunity['opportunity_end_date'] == date, opportunities_list))
            units = ''
            for opportunity in filtered_opportunities:
                units += f"{opportunity['opportunity_code']},"
            cashFlow['units'] = units[:-1]
            final_cashFlow_list.append(cashFlow)

        for opportunity in opportunities_list:
            # create rollover_date by adding 7 days to opportunity_end_date
            opportunity['rollover_date'] = opportunity['opportunity_end_date'] + timedelta(days=7)

        # convert request['date'] to a datetime object
        request['date'] = datetime.strptime(request['date'].replace("-", "/"), "%Y/%m/%d")
        # filter final_cashFlow_list where date is greater than request['date']
        final_cashFlow_list = list(
            filter(lambda cashFlow: cashFlow['date'] > request['date'], final_cashFlow_list))
        # do the same for opportunities_list with opportunity_end_date
        opportunities_list = list(
            filter(lambda opportunity: opportunity['opportunity_end_date'] > request['date'], opportunities_list))

        # print(opportunities_list[0])
        # print()
        # print(final_cashFlow_list[0])
        # print()
        # print(request)

        filename = create_cash_flow(final_cashFlow_list, request, opportunities_list)

        # print()
        # print("trust", trust)
        end = time.time()
        print("Time taken: ", end - start)
        return {"filename": filename}

    except Exception as e:
        print("XXXXC", e)
        return {"ERROR": "Something went wrong!!"}
    # return investments


@excel_sales_forecast.get("/get_cashflow")
async def get_cashflow(cashflow_name):
    file_name = cashflow_name
    dir_path = "excel_files"
    dir_list = os.listdir(dir_path)
    if file_name in dir_list:
        return FileResponse(f"{dir_path}/{file_name}", filename=file_name)
    else:
        return {"ERROR": "File does not exist!!"}


@excel_sales_forecast.post("/check_if_file_exists")
async def check_if_file_exists(data: Request):
    request = await data.json()
    filename = request['filename'].split("/")[1]
    if os.path.exists(request['filename']):
        return {"filename": filename}
    else:
        return {"Exists": False}


@excel_sales_forecast.get("/get_sales_forecast")
async def sales_forecast(sales_forecast_name):
    file_name = sales_forecast_name
    dir_path = "excel_files"
    dir_list = os.listdir(dir_path)
    if file_name in dir_list:
        return FileResponse(f"{dir_path}/{file_name}", filename=file_name)
    else:
        return {"ERROR": "File does not exist!!"}


@excel_sales_forecast.post("/send_email")
async def send_email(data: Request):
    request = await data.json()
    smtp_server = "depro8.fcomet.com"
    port = 465  # For starttls
    sender_email = 'omh-app@opportunitymanagement.co.za'
    password = "12071994Wb!"

    message = f"""\
    <html>
      <body>
        <p>Hi Wayne,<br>
           <strong>How are you?</strong><br>
          
           Below is all the code required to send an email with attachments using Python:<br><br>
           I did hard code the file name, but you can easily change that to a variable.<br><br>
           {request['message']}      
              <br><br>
              reply to this waynebruton@icloud.com.
              <br><br>
              I should be twins.
        </p>
      </body>
    </html>
    """

    msg = EmailMessage()
    msg['Subject'] = request['subject']
    msg['From'] = sender_email
    msg['To'] = request['to_email']
    msg.set_content(message, subtype='html')

    with open('excel_files/Sales ForecastHeron.xlsx', 'rb') as f:
        file_content = f.read()
    file_name = os.path.basename('Sales ForecastHeron.xlsx')
    msg.add_attachment(file_content, maintype='application', subtype='octet-stream', filename=file_name)

    try:
        with smtplib.SMTP_SSL(smtp_server, port) as server:
            server.ehlo()
            server.login(sender_email, password=password)
            server.send_message(msg)
            server.quit()
            return {"message": "Email sent successfully"}
    except Exception as e:
        print("Error:", e)
        return {"message": "Email not sent"}


@excel_sales_forecast.post("/create_draw_doc")
async def create_draw_doc(data: Request):
    request = await data.json()

    # print("request", request)
    investor_acc_numbers = []
    for investor in request['drawdowns']:
        # if investor['_id'] exists then delete it
        if '_id' in investor:
            investor['id'] = str(investor['_id'])
            del investor['_id']

        investor_acc_numbers.append(investor['investor_acc_number'])

    for item in request['previous_draws']:
        if '_id' in item:
            item['id'] = str(item['_id'])
            del item['_id']

    # print(investor_acc_numbers)
    # eliminate duplicates in investor_acc_numbers
    investor_acc_numbers = list(dict.fromkeys(investor_acc_numbers))
    # print(investor_acc_numbers)
    # get all investors from db where investor_acc_number in investor_acc_numbers and return on the
    # investor_acc_number and investment_name

    investor_list = list(db.investors.find({
        "investor_acc_number": {"$in": investor_acc_numbers},
    }, {
        "_id": 0,
        "investor_acc_number": 1,
        "investment_name": 1,
        "investor_name": 1,
        "investor_surname": 1,
    }))

    # print(investor_list)
    for investor in request['drawdowns']:
        if 'investment_name' not in investor:
            investor['investment_name'] = ""
        for investor_db in investor_list:
            if investor['investor_acc_number'] == investor_db['investor_acc_number']:
                # check if investment_name exists in investor_db

                if 'investment_name' in investor_db:
                    investor['investment_name'] = investor_db['investment_name']
                else:
                    investor['investment_name'] = investor_db['investor_surname'] + " " + investor_db['investor_name'][
                                                                                          :1]

    app_total = list(db.investors.find({}))

    new_app_total = []

    if request['drawdowns'][0]['Category'] == "Heron View" or request['drawdowns'][0]['Category'] == "Heron Fields":
        Category = "Heron"
    else:
        Category = "Endulini"

    for item in app_total:
        for trust in item['trust']:
            trust['investor_acc_number'] = item['investor_acc_number']
            new_app_total.append(trust)

    # if Category = "Heron" then filter new_app_total where Category is equal to "Heron View" or "Heron Fields"
    # else filter new_app_total where Category is equal to "Endulini"
    if Category == "Heron":
        new_app_total = list(
            filter(lambda trust: trust['Category'] == "Heron View" or trust['Category'] == 'Heron Fields',
                   new_app_total))
    else:
        new_app_total = list(filter(lambda trust: trust['Category'] == "Endulini", new_app_total))

    # filter out of new_app_total where release_date is equal to ""
    new_app_total = list(filter(lambda trust: trust['release_date'] != "", new_app_total))

    new_app_total = [investor for investor in new_app_total if
                     not (investor['investor_acc_number'] == "ZCAM01" and investor[
                         'opportunity_code'] == "HFA101" and investor['investment_number'] == 1)]

    new_app_total = [investor for investor in new_app_total if
                     not (investor['investor_acc_number'] == "ZJHO01" and investor[
                         'opportunity_code'] == "HFA304" and investor['investment_number'] == 1)]

    new_app_total = [investor for investor in new_app_total if
                     not (investor['investor_acc_number'] == "ZPJB01" and investor[
                         'opportunity_code'] == "HFA205" and investor['investment_number'] == 1)]

    new_app_total = [investor for investor in new_app_total if
                     not (investor['investor_acc_number'] == "ZERA01" and investor[
                         'opportunity_code'] == "EA205" and investor['investment_number'] == 3)]

    new_app_total = [investor for investor in new_app_total if
                     not (investor['investor_acc_number'] == "ZVOL01" and investor[
                         'opportunity_code'] == "EA103" and investor['investment_number'] == 3)]

    # save new_app_total to csv file but investment_amount needs to be a number and not a string
    # for investor in new_app_total:
    #     investor['investment_amount'] = float(investor['investment_amount'])
    # df = pd.DataFrame(new_app_total)
    # df.to_csv('excel_files/app_total.csv', index=False)

    # as final_app_total add all the investment_amounts
    final_app_total = sum([float(investor['investment_amount']) for investor in new_app_total])

    print("app_total", final_app_total)

    # print(request['drawdowns'][0])

    # print(request)

    file_name = create_draw_down_file(request=request, app_total=final_app_total)

    file_name = f"{file_name}.xlsx"

    # file_name = sales_forecast_name
    # print(file_name)
    # dir_path = "excel_files"
    # dir_list = os.listdir(dir_path)
    # print(dir_list)
    # if file_name in dir_list:
    #     return FileResponse(f"{dir_path}/{file_name}", filename=file_name)
    # else:
    #     return {"ERROR": "File does not exist!!"}

    # print(test)

    return {"message": "success", "file_name": file_name}


@excel_sales_forecast.get("/get_draw_doc")
async def draw_doc(file_name):
    file_name = file_name
    print(file_name)
    dir_path = "excel_files"
    dir_list = os.listdir(dir_path)
    print(dir_list)
    if file_name in dir_list:
        return FileResponse(f"{dir_path}/{file_name}", filename=file_name)
    else:
        return {"ERROR": "File does not exist!!"}


@excel_sales_forecast.post("/process_draw")
async def process_draw(data: Request):
    request = await data.json()
    request = request['planned_draws']

    try:
        for item in request:
            item['id'] = str(item['_id'])
            id = item['id']
            del item['id']
            del item['_id']
            investor_list = list(db.investors.find({"investor_acc_number": item['investor_acc_number']}))
            investor_id = str(investor_list[0]['_id'])
            del investor_list[0]['_id']

            # if investor_list[0]['investor_acc_number'] == 'ZWIL02':
            #     print("investor_list", investor_list[0]['investments'][-1:])
            #     print()
            for momentum in investor_list[0]['trust']:
                if momentum['opportunity_code'] == item['opportunity_code'] and momentum['release_date'] == "":
                    momentum['release_date'] = item['draw_down_date']
                    momentum['draw'] = f"Draw{item['draw_number']}."

                    insert = {
                        'opportunity_code': momentum['opportunity_code'],
                        'investment_number': momentum['investment_number'],
                        'investment_amount': momentum['investment_amount'], 'draw': momentum['draw'],
                        'release_date': momentum['release_date'], 'email_sent': False, 'end_date': "",
                        'Category': momentum['Category'], "interest": momentum['interest'],
                        'investment_interest_rate': momentum['project_interest_rate'],
                        'dateCreated': momentum['creation_date'], 'interest': '', 'rollover_amount': 0,
                        'rollover_date': '', 'rollover_interest_rate': 0, 'rollover_interest': 0,
                        'interestPaidOnTransfer': 0}

                    investor_list[0]['investments'].append(insert)

            # if investor_list[0]['investor_acc_number'] == 'ZWIL02':
            #     print("trust", investor_list[0]['trust'][-1:])
            #     print()
            #     print("investor_list", investor_list[0]['investments'][-1:])
            #     print()
            #     print("investor_list", investor_list[0])

            # use a transaction to post both the investor and investor_Draws
            with db.client.start_session() as session:
                # Start the transaction
                with session.start_transaction():
                    try:
                        db.investors.update_one({"_id": ObjectId(investor_id)}, {"$set": investor_list[0]})
                        db.investor_Draws.update_one({"_id": ObjectId(id)}, {"$set": item})

                        # If everything goes well, commit the transaction
                        session.commit_transaction()

                    except Exception as e:
                        # If an error occurs, abort the transaction
                        print("Error:", e)
                        session.abort_transaction()

                # End the session (no need to manually commit or end the session)
                # The session context manager automatically handles it

            #

            # print("item", item)
        return {"message": "success"}
    except Exception as e:
        print("Error:", e)
        return {"message": "Error"}


@excel_sales_forecast.post("/get_unallocated_investors")
async def get_unallocated_investors(data: Request):
    request = await data.json()
    # print("request", request)
    try:
        # get investors from db where trust['release_date'] is equal to "" and trust['Category'] is equal to
        # request['developent']
        investors = list(db.investors.aggregate([
            {
                "$match": {
                    "trust": {
                        "$elemMatch": {
                            "release_date": "",
                            "Category": request['development']
                        }
                    }
                }
            },
            {
                "$addFields": {
                    "trust": {
                        "$filter": {
                            "input": "$trust",
                            "as": "trust",
                            "cond": {
                                "$and": [
                                    {"$eq": ["$$trust.release_date", ""]},
                                    {"$eq": ["$$trust.Category", request['development']]}
                                ]
                            }
                        }
                    }
                }
            }
        ]))

        # get investors from db where trust['Category'] is equal to request['developent']
        opportunity_invested = list(db.investors.aggregate([
            {
                "$match": {
                    "trust": {
                        "$elemMatch": {
                            "Category": request['development']
                        }
                    }
                }
            },
            {
                "$addFields": {
                    "trust": {
                        "$filter": {
                            "input": "$trust",
                            "as": "trust",
                            "cond": {
                                "$eq": ["$$trust.Category", request['development']]
                            }
                        }
                    }
                }
            }
        ]))

        # print("opportunity_invested", opportunity_invested[0]['trust'])
        trust_list = []
        for item in opportunity_invested:
            for trust_item in item['trust']:
                insert = {}
                insert['investment_amount'] = float(trust_item['investment_amount'])
                insert['opportunity_code'] = trust_item['opportunity_code']
                trust_list.append(insert)

        # print("trust", trust_list[0:2])

        return_list = []
        for investor in investors:
            investor['id'] = str(investor['_id'])
            del investor['_id']
            for trust in investor['trust']:
                trust['investor_acc_number'] = investor['investor_acc_number']
                trust['investor_name'] = investor['investor_name']
                trust['investor_surname'] = investor['investor_surname']
                trust['investment_name'] = investor['investment_name']
                trust['id'] = investor['id']
                insert = {}
                insert['investor_acc_number'] = trust['investor_acc_number']
                insert['investor_name'] = trust['investor_name']
                insert['investor_surname'] = trust['investor_surname']
                insert['investment_name'] = trust['investment_name']
                insert['id'] = trust['id']
                insert['planned_draw_date'] = trust.get('planned_draw_date', "")
                insert['opportunity_code'] = trust['opportunity_code']
                insert['investment_amount'] = trust['investment_amount']
                insert['investment_number'] = trust['investment_number']
                insert['Category'] = trust['Category']
                insert['deposit_date'] = trust['deposit_date']
                return_list.append(insert)

        # print("return_list", return_list[0:3])

        # sort return_list by deposit_date, opportunity_code, investment_number

        return_list = sorted(return_list,
                             key=lambda k: (k['deposit_date'], k['opportunity_code'], k['investment_number']))

        opportunities = list(db.opportunities.find({"Category": request['development']}))
        # print("opportunities", opportunities[0:2])

        not_fully_invested = []

        for opportunity in opportunities:
            # filter trust as trust_filtered where trust['opportunity_code'] is equal to opportunity['opportunity_code']
            # print("opportunity", opportunity['opportunity_code'], opportunity['opportunity_amount_required'])
            amount_required = float(opportunity['opportunity_amount_required'])
            # using list comprehension filter the trust list where trust['opportunity_code'] is equal to opportunity[
            # 'opportunity_code']
            trust_filtered = [trust for trust in trust_list if
                              trust['opportunity_code'] == opportunity['opportunity_code']]
            if trust_filtered:
                # print("trust_filtered:", trust_filtered[0])
                # sum the investment_amount in trust_filtered
                amount_invested = sum([trust['investment_amount'] for trust in trust_filtered])

            else:
                amount_invested = 0

            # print("amount_invested", amount_invested, amount_required)
            if amount_invested < amount_required:
                insert = {'opportunity_code': opportunity['opportunity_code'],
                          'opportunity_shortfall': amount_required - amount_invested,
                          'opportunity_amount_required': amount_required,
                          'opportunity_invested_to_date': amount_invested,
                          'opportunity_end_date': opportunity['opportunity_end_date'],
                          'opportunity_release_date': "", 'opportunity_deposit_date': "",
                          'Category': opportunity['Category'],
                          'opportunity_end_date': opportunity['opportunity_end_date'],
                          'opportunity_final_transfer_date': opportunity['opportunity_final_transfer_date'],
                          'opportunity_interest_rate': float(opportunity['opportunity_interest_rate'])}

                not_fully_invested.append(insert)

        # print("not_fully_invested", not_fully_invested[0:2])

        # print("investors", investors[0:2])

        # sort not_fully_invested by opportunity_end_date, opprtunity_shortfall, opportunity_code
        not_fully_invested = sorted(not_fully_invested,
                                    key=lambda k: (k['opportunity_end_date'], k['opportunity_shortfall'],
                                                   k['opportunity_code']))

        future_draws = list(db.future_cf_requirements.find({}))
        # print("future_draws:",future_draws)
        for draw in future_draws:
            draw['id'] = str(draw['_id'])
            del draw['_id']

        list_to_return = {
            "not Drawn": return_list,
            "not fully invested": not_fully_invested,
            "futureDraws": future_draws,
            "success": "success message"
        }
        return list_to_return
        # return return_list

    except Exception as e:
        print("Error:", e)
        return {"message": "Error"}


@excel_sales_forecast.post("/process_unallocated_investors")
async def process_unallocated_investors(data: Request):
    request = await data.json()
    # create a list called unallocated_investors and loop through request['draws'] and append to
    # unallocated_investors where src = "unallocated" using list comprehension
    unallocated_investors = [draw for draw in request['draws'] if draw['src'] == "unallocated"]
    # loop through unallocated_investors and copy draw_date to release_date then convert draw_date to datetime and
    # deduct 30 days and convert back to string and put it to deposit_date
    if unallocated_investors:
        for investor in unallocated_investors:
            investor['release_date'] = investor['draw_date']
            investor['draw_date'] = datetime.strptime(investor['draw_date'].replace("-", "/"), "%Y/%m/%d") - timedelta(
                days=30)
            investor['deposit_date'] = investor['draw_date'].strftime("%Y/%m/%d")
            investor['total_investment'] = investor['opportunity_invested_to_date']
            investor['unallocated_investment'] = investor['opportunity_shortfall']
            investor['project_interest_rate'] = investor['opportunity_interest_rate']
            del investor['investment_amount']
            del investor['opportunity_release_date']
            del investor['opportunity_deposit_date']
            del investor['opportunity_shortfall']
            del investor['opportunity_invested_to_date']
            del investor['opportunity_interest_rate']
            del investor['src']
            del investor['draw_date']
            del investor['draw_number']

    # create a list called unreleased_investors and loop through request['draws'] and append to unreleased_investors
    # where draw['src'] = "draw" using list comprehension print() print("draws", request['draws'])
    # unreleased_investors = [draw for draw in request['draws'] if draw['src'] == "draw"]
    unreleased_investors = [draw for draw in request['draws'] if 'src' in draw and draw['src'] == "draw"]

    for draw in unreleased_investors:
        draw['planned_release_date'] = draw['draw_date']
        draw['draw'] = f"Draw{draw['draw_number']}."

    try:
        db.future_cf_requirements.delete_many({})

        unallocated_investments.delete_many({})

        if unallocated_investors:
            insert = unallocated_investments.insert_many(unallocated_investors)

        # get investor from db where investor_acc_number is equal to draw['investor_acc_number']
        if unreleased_investors:
            for draw in unreleased_investors:
                investor = list(db.investors.find({"investor_acc_number": draw['investor_acc_number']}))

                id = str(investor[0]['_id'])
                del investor[0]['_id']
                for trust_item in investor[0]['trust']:
                    if trust_item['opportunity_code'] == draw['opportunity_code'] and trust_item[
                        'release_date'] == "" and \
                            trust_item['investment_number'] == draw['investment_number']:
                        trust_item['draw'] = draw['draw']
                        trust_item['planned_release_date'] = draw['planned_release_date']
                        print(trust_item)

                db.investors.update_one({"_id": ObjectId(id)}, {"$set": investor[0]})

        return {"message": "success"}

    except Exception as e:
        print("Error:", e)
        return {"message": "Error"}


@excel_sales_forecast.post("/update_future_cf_requirements")
async def update_future_cf_requirements(data: Request):
    request = await data.json()
    print("request", request['futureDraws'])
    future_draws_list = request['futureDraws']
    try:
        db.future_cf_requirements.delete_many({})
        db.future_cf_requirements.insert_many(request['futureDraws'])
        return {"message": "success"}
    except Exception as e:
        print("Error:", e)
        return {"message": "Error"}


@excel_sales_forecast.get("/draw_history")
async def draw_history():
    # get info from investors and project investor_acc_number, investor_name, investor_surname, investment_name and trust array

    # request = await data.json()
    # print("request", request)
    try:
        draw_history = list(db.investors.aggregate([
            {
                "$project": {
                    "investor_acc_number": 1,
                    "Category": 1,
                    "investor_name": 1,
                    "investor_surname": 1,
                    "investment_name": 1,
                    "trust": 1,
                    "pledges": 1,
                    "_id": 0
                }
            }
        ]))

        # remove records where trust is empty
        final_draw_history = []
        pledges_history = [item for item in draw_history if len(item['pledges']) > 0 ]
        # print("pledges_history", pledges_history[0]['pledges'])
        draw_history = [item for item in draw_history if item['trust']]
        # print("draw_history", draw_history[0])
        for draw in draw_history:
            for trust in draw['trust']:
                insert = {'investor_acc_number': draw['investor_acc_number'],
                          'investment_name': draw['investment_name'], 'opportunity_code': trust['opportunity_code'],
                          'investment_amount': float(trust['investment_amount']), 'draw_date': trust['release_date'],
                          'investment_date': trust['deposit_date'], 'Category': trust['Category']}
                if insert['draw_date'] == "" or insert['draw_date'] == None:
                    insert['drawn_to_date'] = 0
                    insert['available_to_draw'] = float(insert['investment_amount'])
                else:
                    insert['drawn_to_date'] = float(insert['investment_amount'])
                    insert['available_to_draw'] = 0

                final_draw_history.append(insert)

        # filter out of final_draw_history where Category is equal to "Southwark

        final_draw_history = list(filter(lambda draw: draw['Category'] != "Southwark", final_draw_history))

        # sort final_draw_history by investment_date, investor_acc_number, opportunity_code
        final_draw_history = sorted(final_draw_history,
                                    key=lambda k: (k['investment_date'], k['investor_acc_number'],
                                                   k['opportunity_code']))

        final_pledges_history = []
        for draw in pledges_history:
            for pledge in draw['pledges']:
                if float(pledge['investment_amount']) > 0:
                    insert = {'investor_acc_number': draw['investor_acc_number'],
                              'investment_name': draw['investment_name'], 'opportunity_code': pledge['opportunity_code'],
                              'investment_amount': float(pledge['investment_amount']), 'Category': pledge['Category']}


                final_pledges_history.append(insert)

        # print("final_pledges_history", final_pledges_history[0:2])

        # remove duplicates from final_pledges_history if investment_amount and opportunity_code and investor_acc_number is the same
        final_pledges_history = [dict(t) for t in {tuple(d.items()) for d in final_pledges_history}]
        # sort by  opportunity_code , investor_acc_number, investment_amount
        final_pledges_history = sorted(final_pledges_history,
                                        key=lambda k: (k['opportunity_code'], k['investor_acc_number'],
                                                        k['investment_amount']))



        report_data = create_draw_history_report(final_draw_history, final_pledges_history)

        return report_data
    except Exception as e:
        print("Error:", e)
        return {"message": "Error"}


@excel_sales_forecast.get("/get_draw_history_report")
async def sales_forecast(draw_report_name):
    file_name = draw_report_name.split("/")[1]
    dir_path = "excel_files"
    dir_list = os.listdir(dir_path)
    print(f"{dir_path}/{file_name}")
    if file_name in dir_list:
        return FileResponse(f"{dir_path}/{file_name}", filename=file_name)
    else:
        return {"ERROR": "File does not exist!!"}
