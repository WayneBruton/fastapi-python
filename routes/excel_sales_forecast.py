from fastapi import APIRouter, Request
# from fastapi.responses import FileResponse
from excel_functions.sales_forecast import create_sales_forecast_file
from config.db import db
# import os
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
    start = time.time()
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
                'investment_number': 1,
                'available_date': 1,
            }
        }, {
            '$match': {
                'isAvailable': True
            }
        }
    ]))

    # return request

    # FILTER OUT NON RELEVANT OPPORTUNITIES
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

    # SEPARATE INVESTMENTS AND TRUST
    trust_items = []
    investment_items = []
    interim_data = []

    for inv in initial_forecast_data:
        for t in inv['trust']:
            t['investor_acc_number'] = inv['investor_acc_number']
            t["investor_name"] = f"{inv['investor_surname']} {inv['investor_name'][0:1]}"
            t["source"] = "Trust"
            trust_items.append(t)

        for i in inv['investments']:
            i['investor_acc_number'] = inv['investor_acc_number']
            i["investor_name"] = f"{inv['investor_surname']} {inv['investor_name'][0:1]}"
            i["source"] = "XInvestment"
            investment_items.append(i)

    investment_items = [x for x in investment_items if
                        x.get('early_release') is None or x.get('early_release') == False]

    # SORT RATES BY DATE DESCENDING
    for rate_item in rates_retrieved:
        rate_item["Efective_date"] = rate_item["Efective_date"].replace("/", "-")
        rate_item["Efective_date"] = datetime.strptime(rate_item["Efective_date"], '%Y-%m-%d')
    rates_retrieved = sorted(rates_retrieved, key=lambda d: d['Efective_date'], reverse=True)

    # FILTER COSTS
    costs = [x for x in costs if x["Development"] == request["Category"][0]]

    # print("ZZZZZ",opportunities_listed[0])

    new_trust_items = []
    for trust in trust_items:
        filtered = [x for x in investment_items if x["opportunity_code"] == trust['opportunity_code']
                    and x['investor_acc_number'] == trust['investor_acc_number']
                    and x['investment_number'] == trust['investment_number']]
        if len(filtered) == 1:
            filtered[0]['deposit_date'] = trust['deposit_date']
            filtered[0]['investment_end_date'] = filtered[0]['end_date']
            new_trust_items.append(filtered[0])
        else:
            # opportunity = [x for x in opportunities_listed if x['opportunity_code'] == trust['opportunity_code']]
            trust['investment_interest_rate'] = 18
            trust['investment_end_date'] = ""
            trust['rollover_amount'] = 0
            new_trust_items.append(trust)

    # print(trust_items[0])
    # print()
    # print(investment_items[0])
    # print()
    # print(new_trust_items[0])

    # trust_items = [y for x in [trust_items, investment_items] for y in x]
    investment_items = []

    # trust_items = sorted(trust_items, key=lambda k: (k['investor_acc_number'], k['opportunity_code'],
    #                                                  k['investment_number'], k['source']), reverse=False)

    # new_filter = []
    # new_trust_items = []

    # for obj in trust_items:
    #     new_filter.append(obj)

    # for item_a in trust_items:
    #     filter_a = [x for x in new_filter if x['investor_acc_number'] == item_a['investor_acc_number']
    #                 and x['opportunity_code'] == item_a['opportunity_code']
    #                 and x['investment_number'] == item_a['investment_number']]
    #
    #     if len(filter_a) == 1:
    #         filter_a[0]["investment_interest_rate"] = 18
    #         new_trust_items.append(filter_a[0])
    #
    #     elif len(filter_a) == 2:
    #
    #         filter_b = [x for x in filter_a if x['source'] == "XInvestment"]
    #
    #         trust_filter = [x for x in filter_a if x['source'] == "Trust"]
    #         filter_b[0]["deposit_date"] = trust_filter[0]["deposit_date"]
    #
    #         new_trust_items.append(filter_b[0])

    trust_items = new_trust_items

    # for x in trust_items:
    #     if x["investor_acc_number"] == "ZBRU01":
    #         print("ZZZZZZ:::::", x)

    # for item in trust_items:
    #     if item["investor_acc_number"] == "ZJER01" and item['opportunity_code'] == "EB102":
    #         print(item)
    #
    # print()
    # print()

    # print(len(opportunities_listed))

    # CREATE FINAL DATA (INITIAL)
    for opportunity in opportunities_listed:
        # if opportunity['opportunity_code'] == 'EA101':
        #     print(opportunity)
        opportunity_code = opportunity["opportunity_code"]

        if "opportunity_end_date" not in opportunity:
            opportunity["opportunity_end_date"] = ""
        # ADD BASIC UNIT INFO TO FINAL DATA
        insert = {"opportunity_code": opportunity["opportunity_code"],
                  "opportunity_amount_required": float(opportunity["opportunity_amount_required"]),
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

        costs = [x for x in costs if x.get('Description') is not None]
        for cost in costs:
            key = cost["Description"]
            rate = cost["rate"]
            insert[key] = rate

        interim_data.append(insert)

    # print(interim_data[0])
    # print(len(interim_data))

    final_data_trust = []
    final_data_trust_and_investment = []

    # print(interim_data[0])
    # print()
    # print(len(interim_data))
    # print()
    # print(trust_items[0])

    print(interim_data[0])

    # REMOVE DUPLICATES
    interim_data = [i for n, i in enumerate(interim_data)
                    if i not in interim_data[n + 1:]]

    # POPULATE FINAL TRUST ITEMS
    for item in trust_items:
        opportunity_code = item['opportunity_code']
        filtered_data = [x for x in interim_data if x['opportunity_code'] == item['opportunity_code']]

        # print("filtered_data",filtered_data[0])
        # print()
        # print("item", item)
        # print()
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
        insert1['investment_interest_rate'] = float(item['investment_interest_rate'])
        insert1['released_investment_amount'] = float(item['investment_amount'])
        insert1['investment_end_date'] = item['investment_end_date']
        insert1['investment_release_date'] = item['release_date']
        insert1['released_investment_number'] = item['investment_number']
        insert1['rollover_amount'] = 0


        # print(insert)

        final_data_trust.append(insert1)
        final_data_trust_and_investment.append(insert1)

        # print("INDEX::::", final_data_trust)
        # print()
        # print()

    # final_data_trust = [i for n, i in enumerate(final_data_trust)
    #                     if i not in final_data_trust[n + 1:]]

    # POPULATE FINAL TRUST ITEMS WITH INVESTMENTS
    # for item in investment_items:
    #     opportunity_code = item['opportunity_code']
    #     try:
    #         filtered_data = list(filter(
    #             lambda x: x['opportunity_code'] == opportunity_code and x['investor_acc_number'] == item[
    #                 'investor_acc_number'] and x['investment_number'] == item['investment_number'], final_data_trust))
    #     except:
    #         filtered_data = []
    #
    #     if len(filtered_data) > 0:
    #
    #         insert2 = {}
    #         for data in filtered_data:
    #             for key in data:
    #                 insert2[key] = data[key]
    #                 # print(insert2[key])
    #         insert2['released_investment_amount'] = float(item['investment_amount'])
    #         insert2['investment_end_date'] = item['end_date']
    #         insert2['investment_release_date'] = item['release_date']
    #         insert2['released_investment_number'] = item['investment_number']
    #         insert2['investment_interest_rate'] = float(item['investment_interest_rate'])
    #
    #         final_data_trust_and_investment.append(insert2)

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
            insert["investment_interest_rate"] = 18
            final_data_trust_and_investment.append(insert)

    final_data_trust_and_investment = sorted(final_data_trust_and_investment,
                                             key=lambda z: (z['opportunity_code'], z['investor_acc_number']))

    # ENSURE FINAL TRANSFER DATE HAS A VALUE
    for inv in final_data_trust_and_investment:
        if inv["opportunity_end_date"] == "":
            inv["opportunity_end_date"] = "2022/03/01"
        if inv["opportunity_final_transfer_date"] == "":
            inv["opportunity_final_transfer_date"] = inv["opportunity_end_date"]
        filtered_rollovers = [x for x in project_rollovers if
                              x['opportunity_code'] == inv['opportunity_code'] and x["investor_acc_number"] == inv[
                                  "investor_acc_number"]]
        if not filtered_rollovers:
            inv['rollover_amount'] = 0
            inv['rollover_date'] = ""
        else:
            inv["rollover_amount"] = filtered_rollovers[0]["rollover_amount"]
            inv["rollover_date"] = filtered_rollovers[0]["available_date"]

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
            # if inv['investor_acc_number'] == "ZAFR01":
            #     print("deposit_date_to_date", deposit_date_to_date, inv['investor_acc_number'], inv)
            #     break

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
                while deposit_date_to_date <= report_date:
                    # print("deposit_date_to_date", deposit_date_to_date, inv['investor_acc_number'])
                    filtered_rate = [x for x in rates_retrieved if x["Efective_date"] <= deposit_date][0]["rate"]
                    investment_interest_to_date += inv["investment_amount"] * float(filtered_rate) / 100 / 365
                    deposit_date_to_date = deposit_date_to_date + timedelta(days=1)
                inv["investment_interest_to_date"] = investment_interest_to_date

            # RELEASED INTEREST
            if inv["release_date"] != "":
                from_date = inv["release_date"].replace("/", "-")
                from_date = datetime.strptime(from_date, '%Y-%m-%d')
                from_date = from_date + timedelta(days=1)

                from_date_to_date = inv["release_date"].replace("/", "-")
                from_date_to_date = datetime.strptime(from_date_to_date, '%Y-%m-%d')
                from_date_to_date = from_date_to_date + timedelta(days=1)

                # RELEASED INTEREST TILL TRANSFER
                released_interest = 0

                while from_date <= final_transfer_date:
                    released_interest += inv["investment_amount"] * inv["investment_interest_rate"] / 100 / 365
                    from_date = from_date + timedelta(days=1)
                inv["released_interest_total"] = released_interest

                # RELEASED INTEREST TO DATE
                released_interest_to_date = 0

                if report_date <= final_transfer_date:
                    while from_date_to_date <= report_date:
                        released_interest_to_date += inv["investment_amount"] * inv[
                            "investment_interest_rate"] / 100 / 365
                        from_date_to_date = from_date_to_date + timedelta(days=1)
                    inv["released_interest_to_date"] = released_interest_to_date
                else:
                    inv["released_interest_to_date"] = released_interest
    # start = time.time()
    create_sales_forecast_file(final_data_trust_and_investment, request)
    end = time.time()
    print("Time Taken: ", end - start)

    return final_data_trust_and_investment
