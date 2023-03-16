import os

from bson import ObjectId
from fastapi import APIRouter, Request
from fastapi.responses import FileResponse
from excel_functions.sales_forecast_excel import create_sales_forecast_file
from config.db import db
import time
from datetime import datetime
from datetime import timedelta

excel_sales_forecast = APIRouter()

# MONGO COLLECTIONS
investors = db.investors
rates = db.rates
opportunities = db.opportunities
unallocated_investments = db.unallocated_investments
sales_parameters = db.salesParameters
rollovers = db.investorRollovers


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
                insert = {}
                insert['opportunity_code'] = opportunity['opportunity_code']
                insert['opportunity_amount_required'] = opportunity['opportunity_amount_required']
                insert['total_investment'] = total_investment_amount
                insert['unallocated_investment'] = float(
                    opportunity['opportunity_amount_required']) - total_investment_amount
                insert['deposit_date'] = ""
                insert['release_date'] = ""
                insert['project_interest_rate'] = 0
                insert['Category'] = opportunity['Category']
                insert['opportunity_end_date'] = opportunity['opportunity_end_date']
                insert['opportunity_final_transfer_date'] = opportunity['opportunity_final_transfer_date']
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
async def get_sales_info(data: Request):
    start = time.time()
    request = await data.json()
    print("request", request)

    # Get Investors and Manipulate accordingly
    investments = []
    pledges = []
    trust_list = []

    try:

        investor_list = list(db.investors.find({}))
        for investor in investor_list:
            investor['id'] = str(investor['_id'])
            del investor['_id']

            # Filter Pledges, Investments and Trust by request['Category'] items only using list comprehension
            investor['pledges'] = [pledge for pledge in investor['pledges'] if
                                   pledge['Category'] in request['Category']]
            investor['investments'] = [investment for investment in investor['investments'] if
                                       investment['Category'] in request['Category']]
            investor['trust'] = [trust for trust in investor['trust'] if trust['Category'] in request['Category']]

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
        opportunities_list = list(db.opportunities.find({}))
        for opportunity in opportunities_list:
            opportunity['id'] = str(opportunity['_id'])
            del opportunity['_id']

        # using list comprehension to filter out opportunities that are not in the request['Category'] list
        opportunities_list = [opportunity for opportunity in opportunities_list if
                              opportunity['Category'] in request['Category']]

        for opportunity in opportunities_list:
            if opportunity["opportunity_final_transfer_date"] != "":
                opportunity["opportunity_transferred"] = True
            else:
                opportunity["opportunity_transferred"] = False
            if opportunity["opportunity_sold"] == False:
                opportunity["opportunity_transferred"] = False

            if opportunity['opportunity_final_transfer_date'] == '':
                # if  opportunity['opportunity_end_date'] != '' or does not exist
                if 'opportunity_end_date' in opportunity:
                    if opportunity['opportunity_end_date'] != '':
                        opportunity['opportunity_final_transfer_date'] = opportunity['opportunity_end_date']
                    else:
                        opportunity['opportunity_final_transfer_date'] = opportunity['opportunity_occupation_date']
                else:
                    opportunity['opportunity_final_transfer_date'] = opportunity['opportunity_occupation_date']

        # Get Sales Parameters and Manipulate accordingly
        sales_parameters_list = list(db.salesParameters.find({}))
        for sales_parameter in sales_parameters_list:
            sales_parameter['id'] = str(sales_parameter['_id'])
            del sales_parameter['_id']
        # Using list comprehension to filter out sales parameters that are not in the request['Category'] list
        sales_parameters_list = [sales_parameter for sales_parameter in sales_parameters_list if
                                 sales_parameter['Development'] in request['Category']]

        # Get Rollovers and Manipulate accordingly
        rollovers_list = list(db.investorRollovers.find({}))
        for rollover in rollovers_list:
            rollover['id'] = str(rollover['_id'])
            del rollover['_id']
        # Using list comprehension to filter out rollovers that are not in the request['Category'] list
        rollovers_list = [rollover for rollover in rollovers_list if
                          rollover['Category'] in request['Category']]

        # Get Rates and Manipulate accordingly
        rates_list = list(db.rates.find({}))
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
            # If trust['project_interest_rate'] does not exist, create it and set it to 0.00
            if 'project_interest_rate' not in trust:
                trust['project_interest_rate'] = 0.00
            # If trust['planned_release_date'] does not exist, create it and set it to ""
            if 'planned_release_date' not in trust:
                trust['planned_release_date'] = ""

            insert = {"investor_surname": trust["investor_surname"], "investor_name": trust["investor_name"],
                      "investor_acc_number": trust["investor_acc_number"],
                      "investment_amount": float(trust["investment_amount"]), "deposit_date": trust["deposit_date"],
                      "release_date": trust["release_date"], "opportunity_code": trust["opportunity_code"],
                      "Category": trust["Category"], "investment_number": trust["investment_number"],
                      "project_interest_rate": trust["project_interest_rate"],
                      "planned_release_date": trust["planned_release_date"],
                      "trust_interest": 0.00,
                      "investment_interest_rate": 15.00, "investment_end_date": "",
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
            rate['Efective_date'] = datetime.strptime(rate['Efective_date'], '%Y-%m-%d')
        # order rates_list by Efective_date in descending order
        rates_list = sorted(rates_list, key=lambda k: k['Efective_date'], reverse=True)

        for index, investment in enumerate(final_investors_list):
            # If investment["investment_interest_rate"] is an empty string then set it to 15.5
            investment["investment_interest_rate"] = investment["investment_interest_rate"] if investment[
                                                                                                   "investment_interest_rate"] != "" else 15.5

            # if investment["opportunity_code"] is equal to "HFB313" and the investment['investor_acc_number'] is equal to "ZCON01" then set investment["release_date"] equal to investment['deposit_date']
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

                investment["planned_release_date"] = str(datetime.strptime(investment["deposit_date"],
                                                                           '%Y/%m/%d') + timedelta(days=30)).split(" ")[
                    0]
            else:
                investment["planned_release_date"] = investment["release_date"]
                # investment["planned_release_date"] = ""

            if investment["investor_acc_number"] != "ZZUN01":
                deposit_date = investment["deposit_date"].replace('-', '/')
                deposit_date = deposit_date.split(" ")[0]
                deposit_date = datetime.strptime(deposit_date, '%Y/%m/%d')
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
                            float(investment["investment_amount"]) * interim_rate / 100 / 365)
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

        # bring rollovers into the mix

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
        for opportunity in opportunities_list:
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
                          }

                final_investors_list.append(insert)

        # if the investor_acc_number = "ZCAM01" and the opportunity_code = "HFA101" and the investment_amount =
        # 400000.0 then filter this record out of final_investors_list
        final_investors_list = [investor for investor in final_investors_list if
                                not (investor['investor_acc_number'] == "ZCAM01" and investor[
                                    'opportunity_code'] == "HFA101" and investor['investment_amount'] == 400000.0)]

        # get unallocated_investments from mongo db where the request['Category'] is in the Category in the DB
        unallocated_investments_list = list(db.unallocated_investments.find(
            {"Category": {"$in": request['Category']}}))

        for unallocated_investment in unallocated_investments_list:
            unallocated_investment['id'] = str(unallocated_investment['_id'])
            del unallocated_investment['_id']
            # THE BELOW IS FOR TESTING ONLY
            # if unallocated_investment['opportunity_code'] == "EA203":
            #     unallocated_investment['deposit_date'] = "2022/01/01"
            #     unallocated_investment['release_date'] = "2022/01/15"
            #     unallocated_investment['planned_release_date'] = "2022/01/15"
            #     unallocated_investment['project_interest_rate'] = float(15)

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

                    int_release_date = investor['release_date'].replace("-", "/")
                    int_release_date = datetime.strptime(int_release_date, "%Y/%m/%d")

                    int_planned_release_date = investor['planned_release_date'].replace("-", "/")
                    int_planned_release_date = datetime.strptime(int_planned_release_date, "%Y/%m/%d")

                    # todays_date = request['date'].replace("-", "/")
                    # todays_date = datetime.strptime(todays_date, "%Y/%m/%d")

                    # convert investor['opportunity_final_transfer_date'] to datetime
                    opportunity_final_transfer_date = investor['opportunity_final_transfer_date'].replace("-", "/")
                    opportunity_final_transfer_date = datetime.strptime(opportunity_final_transfer_date, "%Y/%m/%d")

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
                        "interest_total_still_to_be_raised"] = interest_to_be_raised_for_released + interest_to_be_raised_for_momentum

        # sort final investors list by Category, opportunity_code, investor_acc_number
        final_investors_list = sorted(final_investors_list,
                                      key=lambda k: (k['Category'], k['opportunity_code'], k['investor_acc_number']))

        filename = create_sales_forecast_file(final_investors_list, request, pledges)
        end = time.time()
        print("Time Taken: ", end - start)

        return {"filename": f'{filename}.xlsx'}
        # return "Time Taken: ", end - start, len(final_investors_list), final_investors_list

    except Exception as e:
        print("Error:", e)
        return []


# create a route to return the file in filename above to the user

@excel_sales_forecast.get("/get_sales_forecast")
async def sales_forecast(sales_forecast_name):
    print("file_name", sales_forecast_name)
    file_name = sales_forecast_name
    
    dir_path = "excel_files"
    dir_list = os.listdir(dir_path)
    print("dir_list", dir_list)
    if file_name in dir_list:
        return FileResponse(f"{dir_path}/{file_name}", filename=file_name)
    else:
        return {"ERROR": "File does not exist!!"}












# @excel_sales_forecast.get("/get_sales_forecast")
# async def get_file_to_return(file_name):
#
#     print("file_name",file_name)# File Name incl path.
#     is_exists = os.path.exists(file_name)
#     print("is_exists",is_exists)
#     if is_exists:
#         return FileResponse(f"{file_name}", media_type="multipart/form-data")
#     else:
#         return {"ERROR": "File does not exist!!"}
