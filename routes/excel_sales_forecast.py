from fastapi import APIRouter, Request
from fastapi.responses import FileResponse
from config.db import db
import os
import time
from datetime import datetime
from datetime import timedelta

excel_sales_forecast = APIRouter()

# MONGO COLLECTIONS
investors = db.investors
rates = db.rates
opportunities = db.opportunities
sales_parameters = db.salesParameters
rollovers = db.investorRollovers


@excel_sales_forecast.post("/sales-forecast")
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
        },
        {
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
        },
        {
            '$project': {
                'isAvailable': {
                    '$ne': [
                        '$available_date', ''
                    ]
                },
                'id': {
                    '$toString': '$_id'
                },
                '_id': 0,
                'investor_acc_number': 1,
                'opportunity_code': 1,
                'rollover_amount': 1,
                'investment_number': 1
            }
        }, {
            '$match': {
                'isAvailable': True
            }
        }
    ]))

    print(project_rollovers)

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

    final_data = []
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
        rate_item["Efective_date"] = rate_item["Efective_date"].replace("/", "-")
        rate_item["Efective_date"] = datetime.strptime(rate_item["Efective_date"], '%Y-%m-%d')
    rates_retrieved = sorted(rates_retrieved, key=lambda d: d['Efective_date'], reverse=True)
    # print(rates_retrieved)

    # FILTER COSTS
    costs = [x for x in costs if x["Development"] == request["Category"][0]]

    # CREATE FINAL DATA (INITIAL)
    for opportunity in opportunities_listed:
        opportunity_code = opportunity["opportunity_code"]

        if "opportunity_end_date" not in opportunity:
            opportunity["opportunity_end_date"] = ""
        # ADD BASIC UNIT INFO TO FINAL DATA
        insert = {"opportunity_code": opportunity["opportunity_code"],
                  "opportunity_amount_required": opportunity["opportunity_amount_required"],
                  "opportunity_end_date": opportunity["opportunity_end_date"],
                  "opportunity_final_transfer_date": opportunity["opportunity_final_transfer_date"],
                  "opportunity_sale_price": opportunity["opportunity_sale_price"],
                  "opportunity_sold": opportunity["opportunity_sold"]}

        if insert["opportunity_end_date"] == "":
            insert["opportunity_end_date"] = request["date"]

        if opportunity["opportunity_final_transfer_date"] == "":
            insert["opportunity_transferred"] = False
            opportunity["opportunity_final_transfer_date"] = opportunity["opportunity_end_date"]
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
            insert2['investment_interest_rate'] = float(item['investment_interest_rate'])

            final_data_trust_and_investment.append(insert2)

    # POPULATE WITH UNALLOCATED DATA
    for item in interim_data:
        filtered_data = list(filter(
            lambda x: x['opportunity_code'] == item['opportunity_code'], final_data_trust_and_investment))
        if len(filtered_data) == 0:
            insert = {}
            for key in item:
                insert[key] = item[key]
            insert['investor_acc_number'] = "ZZUN01"
            insert['investor_name'] = "Allocated UN"
            insert['opportunity_code'] = item['opportunity_code']
            insert['investment_end_date'] = ""
            insert['deposit_date'] = ""
            insert['release_date'] = ""
            insert['investment_release_date'] = ""
            insert['investment_number'] = 999
            insert['released_investment_number'] = 999
            insert['investment_amount'] = 0
            insert['released_investment_amount'] = 0
            insert['investment_release_date'] = ""
            final_data_trust_and_investment.append(insert)

    final_data_trust_and_investment = sorted(final_data_trust_and_investment,
                                             key=lambda z: (z['opportunity_code'], z['investor_acc_number']))

    # ENSURE FINAL TRANSFER DATE HAS A VALUE
    for inv in final_data_trust_and_investment:
        if inv["opportunity_end_date"] == "":
            inv["opportunity_end_date"] = "2022/03/01"
        if inv["opportunity_final_transfer_date"] == "":
            inv["opportunity_final_transfer_date"] = inv["opportunity_end_date"]

    # CALCULATE INVESTMENT INTEREST
    for inv in final_data_trust_and_investment:
        # UNALLOCATED INTEREST
        if inv["investor_acc_number"] == "ZZUN01":
            inv["investment_interest_total"] = 0
            inv["investment_interest_to_date"] = 0
            inv["released_interest_total"] = 0
            inv["released_interest_to_date"] = 0
        # INVESTMENT INTEREST
        if inv["investor_acc_number"] != "ZZUN01":
            final_transfer_date = inv["opportunity_final_transfer_date"].replace("/", "-")
            final_transfer_date = datetime.strptime(final_transfer_date, '%Y-%m-%d')
            report_date = inv["report_date"].replace("/", "-")
            report_date = datetime.strptime(report_date, '%Y-%m-%d')
            deposit_date = inv["deposit_date"].replace("/", "-")
            deposit_date = datetime.strptime(deposit_date, '%Y-%m-%d')
            deposit_date = deposit_date + timedelta(days=1)

            deposit_date_to_date = inv["deposit_date"].replace("/", "-")
            deposit_date_to_date = datetime.strptime(deposit_date_to_date, '%Y-%m-%d')
            deposit_date_to_date = deposit_date_to_date + timedelta(days=1)
            # "opportunity_final_transfer_date"

            if inv["release_date"] != "":
                release_date = inv["release_date"].replace("/", "-")
                release_date = datetime.strptime(release_date, '%Y-%m-%d')
                investment_interest = 0
                while deposit_date <= release_date:
                    filtered_rate = [x for x in rates_retrieved if x["Efective_date"] <= deposit_date][0]["rate"]
                    # print(filtered_rate)
                    investment_interest += inv["investment_amount"] * float(filtered_rate) / 100 / 365
                    deposit_date = deposit_date + timedelta(days=1)
                inv["investment_interest_total"] = investment_interest
                investment_interest_to_date = 0
                if release_date < report_date:
                    while deposit_date_to_date <= release_date:
                        filtered_rate = [x for x in rates_retrieved if x["Efective_date"] <= deposit_date_to_date][0][
                            "rate"]
                        # print(filtered_rate)
                        investment_interest_to_date += inv["investment_amount"] * float(filtered_rate) / 100 / 365
                        deposit_date_to_date = deposit_date_to_date + timedelta(days=1)
                    inv["investment_interest_to_date"] = investment_interest_to_date

            else:
                investment_interest = 0
                while deposit_date <= final_transfer_date:
                    filtered_rate = [x for x in rates_retrieved if x["Efective_date"] <= deposit_date][0]["rate"]
                    investment_interest += inv["investment_amount"] * float(filtered_rate) / 100 / 365
                    deposit_date = deposit_date + timedelta(days=1)
                inv["investment_interest_total"] = investment_interest
                investment_interest_to_date = 0
                while deposit_date <= report_date:
                    filtered_rate = [x for x in rates_retrieved if x["Efective_date"] <= deposit_date][0]["rate"]
                    investment_interest_to_date += inv["investment_amount"] * float(filtered_rate) / 100 / 365
                    deposit_date_to_date = deposit_date_to_date + timedelta(days=1)
                inv["investment_interest_to_date"] = investment_interest_to_date

            # if inv["release_date"] != "" and inv["release_date"]

        # RELEASED INTEREST

    return final_data_trust_and_investment
