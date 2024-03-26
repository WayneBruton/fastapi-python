import os
import re
from datetime import datetime, timedelta

from fastapi import APIRouter, Request, BackgroundTasks
from fastapi.responses import FileResponse
from configuration.db import db
from bson.objectid import ObjectId
import time
from cashflow_excel_functions.cashflow_projection_nsst import cashflow_projections

cashflow = APIRouter()


@cashflow.post("/construction_cashflow")
async def construction_cashflow(data: Request):
    request = await data.json()
    data = request['data']
    try:
        db.cashflow_construction.delete_many({})
        result = db.cashflow_construction.insert_many(data)
        return {"success": True}
    except Exception as e:
        return {"success": False, "error": str(e)}


def update_cashflow_dates():
    new_dates = [
        {
            "opportunity_code": "EA107",
            "forecast_transfer_date": "2024/05/30"
        },
        {
            "opportunity_code": "EA201",
            "forecast_transfer_date": "2024/05/30"
        },
        {
            "opportunity_code": "EA205",
            "forecast_transfer_date": "2024/05/30"
        },
        {
            "opportunity_code": "EA206",
            "forecast_transfer_date": "2024/05/30"
        },
        {
            "opportunity_code": "EA207",
            "forecast_transfer_date": "2024/06/13"
        },
        {
            "opportunity_code": "EA302",
            "forecast_transfer_date": "2024/07/05"
        },
        {
            "opportunity_code": "EA305",
            "forecast_transfer_date": "2024/07/05"
        },
        {
            "opportunity_code": "EB102",
            "forecast_transfer_date": "2024/08/01"
        },
        {
            "opportunity_code": "EB201",
            "forecast_transfer_date": "2024/07/05"
        },
        {
            "opportunity_code": "EB202",
            "forecast_transfer_date": "2024/07/05"
        },
        {
            "opportunity_code": "ED203",
            "forecast_transfer_date": "2024/03/25"
        },
        {
            "opportunity_code": "HFB206",
            "forecast_transfer_date": "2024/04/26"
        },
        {
            "opportunity_code": "HFB207",
            "forecast_transfer_date": "2024/04/26"
        },
        {
            "opportunity_code": "HFB209",
            "forecast_transfer_date": "2024/04/26"
        },
        {
            "opportunity_code": "HFB210",
            "forecast_transfer_date": "2024/05/27"
        },
        {
            "opportunity_code": "HFB211",
            "forecast_transfer_date": "2024/05/27"
        },
        {
            "opportunity_code": "HFB212",
            "forecast_transfer_date": "2024/05/27"
        },
        {
            "opportunity_code": "HFB213",
            "forecast_transfer_date": "2024/06/28"
        },
        {
            "opportunity_code": "HFB215",
            "forecast_transfer_date": "2024/06/28"
        },
        {
            "opportunity_code": "HVC202",
            "forecast_transfer_date": "2024/07/12"
        },
        {
            "opportunity_code": "HVC204",
            "forecast_transfer_date": "2024/07/12"
        },
        {
            "opportunity_code": "HVC205",
            "forecast_transfer_date": "2024/04/12"
        },
        {
            "opportunity_code": "HVC206",
            "forecast_transfer_date": "2024/04/12"
        },
        {
            "opportunity_code": "HVC302",
            "forecast_transfer_date": "2024/05/17"
        },
        {
            "opportunity_code": "HVC304",
            "forecast_transfer_date": "2024/06/06"
        },
        {
            "opportunity_code": "HVC305",
            "forecast_transfer_date": "2024/06/28"
        },
        {
            "opportunity_code": "HVC306",
            "forecast_transfer_date": "2024/04/19"
        },
        {
            "opportunity_code": "HVD102",
            "forecast_transfer_date": "2024/02/19"
        },
        {
            "opportunity_code": "HVD202",
            "forecast_transfer_date": "2024/05/03"
        },
        {
            "opportunity_code": "HVD203",
            "forecast_transfer_date": "2024/04/02"
        },
        {
            "opportunity_code": "HVD301",
            "forecast_transfer_date": "2024/02/19"
        },
        {
            "opportunity_code": "HVD302",
            "forecast_transfer_date": "2024/05/16"
        },
        {
            "opportunity_code": "HVD303",
            "forecast_transfer_date": "2024/04/25"
        },
        {
            "opportunity_code": "HVD304",
            "forecast_transfer_date": "2024/05/16"
        },
        {
            "opportunity_code": "HVE101",
            "forecast_transfer_date": "2025/01/28"
        },
        {
            "opportunity_code": "HVE102",
            "forecast_transfer_date": "2025/01/28"
        },
        {
            "opportunity_code": "HVE103",
            "forecast_transfer_date": "2025/01/28"
        },
        {
            "opportunity_code": "HVE104",
            "forecast_transfer_date": "2025/01/28"
        },
        {
            "opportunity_code": "HVE201",
            "forecast_transfer_date": "2025/01/28"
        },
        {
            "opportunity_code": "HVE202",
            "forecast_transfer_date": "2025/01/28"
        },
        {
            "opportunity_code": "HVE203",
            "forecast_transfer_date": "2025/01/28"
        },
        {
            "opportunity_code": "HVE204",
            "forecast_transfer_date": "2025/01/28"
        },
        {
            "opportunity_code": "HVE301",
            "forecast_transfer_date": "2025/01/28"
        },
        {
            "opportunity_code": "HVE302",
            "forecast_transfer_date": "2025/01/28"
        },
        {
            "opportunity_code": "HVE303",
            "forecast_transfer_date": "2025/01/28"
        },
        {
            "opportunity_code": "HVE304",
            "forecast_transfer_date": "2025/01/28"
        },
        {
            "opportunity_code": "HVF101",
            "forecast_transfer_date": "2025/03/06"
        },
        {
            "opportunity_code": "HVF102",
            "forecast_transfer_date": "2025/03/06"
        },
        {
            "opportunity_code": "HVF103",
            "forecast_transfer_date": "2025/03/06"
        },
        {
            "opportunity_code": "HVF104",
            "forecast_transfer_date": "2025/03/06"
        },
        {
            "opportunity_code": "HVF201",
            "forecast_transfer_date": "2025/03/06"
        },
        {
            "opportunity_code": "HVF202",
            "forecast_transfer_date": "2025/03/06"
        },
        {
            "opportunity_code": "HVF203",
            "forecast_transfer_date": "2025/03/06"
        },
        {
            "opportunity_code": "HVF204",
            "forecast_transfer_date": "2025/03/06"
        },
        {
            "opportunity_code": "HVG101",
            "forecast_transfer_date": "2024/08/19"
        },
        {
            "opportunity_code": "HVG102",
            "forecast_transfer_date": "2024/08/19"
        },
        {
            "opportunity_code": "HVG103",
            "forecast_transfer_date": "2024/08/19"
        },
        {
            "opportunity_code": "HVG104",
            "forecast_transfer_date": "2024/08/19"
        },
        {
            "opportunity_code": "HVG201",
            "forecast_transfer_date": "2024/08/19"
        },
        {
            "opportunity_code": "HVG202",
            "forecast_transfer_date": "2024/08/19"
        },
        {
            "opportunity_code": "HVG203",
            "forecast_transfer_date": "2024/08/19"
        },
        {
            "opportunity_code": "HVG204",
            "forecast_transfer_date": "2024/08/19"
        },
        {
            "opportunity_code": "HVG301",
            "forecast_transfer_date": "2024/08/19"
        },
        {
            "opportunity_code": "HVG302",
            "forecast_transfer_date": "2024/08/19"
        },
        {
            "opportunity_code": "HVG303",
            "forecast_transfer_date": "2024/08/19"
        },
        {
            "opportunity_code": "HVG304",
            "forecast_transfer_date": "2024/08/19"
        },
        {
            "opportunity_code": "HVH101",
            "forecast_transfer_date": "2024/07/17"
        },
        {
            "opportunity_code": "HVH102",
            "forecast_transfer_date": "2024/07/17"
        },
        {
            "opportunity_code": "HVH103",
            "forecast_transfer_date": "2024/07/17"
        },
        {
            "opportunity_code": "HVH104",
            "forecast_transfer_date": "2024/07/17"
        },
        {
            "opportunity_code": "HVH201",
            "forecast_transfer_date": "2024/07/17"
        },
        {
            "opportunity_code": "HVH202",
            "forecast_transfer_date": "2024/07/17"
        },
        {
            "opportunity_code": "HVH203",
            "forecast_transfer_date": "2024/07/17"
        },
        {
            "opportunity_code": "HVH204",
            "forecast_transfer_date": "2024/07/17"
        },
        {
            "opportunity_code": "HVI101",
            "forecast_transfer_date": "2024/07/18"
        },
        {
            "opportunity_code": "HVI102",
            "forecast_transfer_date": "2024/07/18"
        },
        {
            "opportunity_code": "HVI103",
            "forecast_transfer_date": "2024/07/18"
        },
        {
            "opportunity_code": "HVI104",
            "forecast_transfer_date": "2024/07/18"
        },
        {
            "opportunity_code": "HVI201",
            "forecast_transfer_date": "2024/07/18"
        },
        {
            "opportunity_code": "HVI202",
            "forecast_transfer_date": "2024/07/18"
        },
        {
            "opportunity_code": "HVI203",
            "forecast_transfer_date": "2024/07/18"
        },
        {
            "opportunity_code": "HVI204",
            "forecast_transfer_date": "2024/07/18"
        },
        {
            "opportunity_code": "HVJ101",
            "forecast_transfer_date": "2024/09/17"
        },
        {
            "opportunity_code": "HVJ102",
            "forecast_transfer_date": "2024/09/17"
        },
        {
            "opportunity_code": "HVJ103",
            "forecast_transfer_date": "2024/09/17"
        },
        {
            "opportunity_code": "HVJ201",
            "forecast_transfer_date": "2024/09/17"
        },
        {
            "opportunity_code": "HVJ202",
            "forecast_transfer_date": "2024/09/17"
        },
        {
            "opportunity_code": "HVJ203",
            "forecast_transfer_date": "2024/09/17"
        },
        {
            "opportunity_code": "HVJ301",
            "forecast_transfer_date": "2024/09/17"
        },
        {
            "opportunity_code": "HVJ302",
            "forecast_transfer_date": "2024/09/17"
        },
        {
            "opportunity_code": "HVJ303",
            "forecast_transfer_date": "2024/09/17"
        },
        {
            "opportunity_code": "HVJ401",
            "forecast_transfer_date": "2024/09/17"
        },
        {
            "opportunity_code": "HVJ402",
            "forecast_transfer_date": "2024/09/17"
        },
        {
            "opportunity_code": "HVJ403",
            "forecast_transfer_date": "2024/09/17"
        },
        {
            "opportunity_code": "HVK101",
            "forecast_transfer_date": "2024/08/06"
        },
        {
            "opportunity_code": "HVK102",
            "forecast_transfer_date": "2024/08/06"
        },
        {
            "opportunity_code": "HVK103",
            "forecast_transfer_date": "2024/08/06"
        },
        {
            "opportunity_code": "HVK104",
            "forecast_transfer_date": "2024/08/06"
        },
        {
            "opportunity_code": "HVK105",
            "forecast_transfer_date": "2024/08/06"
        },
        {
            "opportunity_code": "HVK106",
            "forecast_transfer_date": "2024/08/06"
        },
        {
            "opportunity_code": "HVK201",
            "forecast_transfer_date": "2024/08/06"
        },
        {
            "opportunity_code": "HVK202",
            "forecast_transfer_date": "2024/08/06"
        },
        {
            "opportunity_code": "HVK203",
            "forecast_transfer_date": "2024/08/06"
        },
        {
            "opportunity_code": "HVK204",
            "forecast_transfer_date": "2024/08/06"
        },
        {
            "opportunity_code": "HVK205",
            "forecast_transfer_date": "2024/08/06"
        },
        {
            "opportunity_code": "HVK206",
            "forecast_transfer_date": "2024/08/06"
        },
        {
            "opportunity_code": "HVK301",
            "forecast_transfer_date": "2024/08/06"
        },
        {
            "opportunity_code": "HVK302",
            "forecast_transfer_date": "2024/08/06"
        },
        {
            "opportunity_code": "HVK303",
            "forecast_transfer_date": "2024/08/06"
        },
        {
            "opportunity_code": "HVK304",
            "forecast_transfer_date": "2024/08/06"
        },
        {
            "opportunity_code": "HVK305",
            "forecast_transfer_date": "2024/08/06"
        },
        {
            "opportunity_code": "HVK306",
            "forecast_transfer_date": "2024/08/06"
        },
        {
            "opportunity_code": "HVK401",
            "forecast_transfer_date": "2024/08/06"
        },
        {
            "opportunity_code": "HVK402",
            "forecast_transfer_date": "2024/08/06"
        },
        {
            "opportunity_code": "HVK403",
            "forecast_transfer_date": "2024/08/06"
        },
        {
            "opportunity_code": "HVK404",
            "forecast_transfer_date": "2024/08/06"
        },
        {
            "opportunity_code": "HVK405",
            "forecast_transfer_date": "2024/08/06"
        },
        {
            "opportunity_code": "HVK406",
            "forecast_transfer_date": "2024/08/06"
        },
        {
            "opportunity_code": "HVL101",
            "forecast_transfer_date": "2024/08/12"
        },
        {
            "opportunity_code": "HVL102",
            "forecast_transfer_date": "2024/08/12"
        },
        {
            "opportunity_code": "HVL103",
            "forecast_transfer_date": "2024/08/12"
        },
        {
            "opportunity_code": "HVL104",
            "forecast_transfer_date": "2024/08/12"
        },
        {
            "opportunity_code": "HVL201",
            "forecast_transfer_date": "2024/08/12"
        },
        {
            "opportunity_code": "HVL202",
            "forecast_transfer_date": "2024/08/12"
        },
        {
            "opportunity_code": "HVL203",
            "forecast_transfer_date": "2024/08/12"
        },
        {
            "opportunity_code": "HVL204",
            "forecast_transfer_date": "2024/08/12"
        },
        {
            "opportunity_code": "HVM101",
            "forecast_transfer_date": "2024/06/07"
        },
        {
            "opportunity_code": "HVM102",
            "forecast_transfer_date": "2024/06/07"
        },
        {
            "opportunity_code": "HVM103",
            "forecast_transfer_date": "2024/06/07"
        },
        {
            "opportunity_code": "HVM104",
            "forecast_transfer_date": "2024/06/07"
        },
        {
            "opportunity_code": "HVM201",
            "forecast_transfer_date": "2024/06/07"
        },
        {
            "opportunity_code": "HVM202",
            "forecast_transfer_date": "2024/06/07"
        },
        {
            "opportunity_code": "HVM203",
            "forecast_transfer_date": "2024/06/07"
        },
        {
            "opportunity_code": "HVM204",
            "forecast_transfer_date": "2024/06/07"
        },
        {
            "opportunity_code": "HVN102",
            "forecast_transfer_date": "2024/05/16"
        },
        {
            "opportunity_code": "HVN104",
            "forecast_transfer_date": "2024/03/20"
        },
        {
            "opportunity_code": "HVN301",
            "forecast_transfer_date": "2024/06/21"
        },
        {
            "opportunity_code": "HVN302",
            "forecast_transfer_date": "2024/05/10"
        },
        {
            "opportunity_code": "HVN303",
            "forecast_transfer_date": "2024/06/21"
        },
        {
            "opportunity_code": "HVN304",
            "forecast_transfer_date": "2024/06/21"
        },
        {
            "opportunity_code": "HVO101",
            "forecast_transfer_date": "2024/08/14"
        },
        {
            "opportunity_code": "HVO102",
            "forecast_transfer_date": "2024/08/14"
        },
        {
            "opportunity_code": "HVO103",
            "forecast_transfer_date": "2024/08/14"
        },
        {
            "opportunity_code": "HVO104",
            "forecast_transfer_date": "2024/08/14"
        },
        {
            "opportunity_code": "HVO105",
            "forecast_transfer_date": "2024/08/14"
        },
        {
            "opportunity_code": "HVO201",
            "forecast_transfer_date": "2024/08/14"
        },
        {
            "opportunity_code": "HVO202",
            "forecast_transfer_date": "2024/08/14"
        },
        {
            "opportunity_code": "HVO203",
            "forecast_transfer_date": "2024/08/14"
        },
        {
            "opportunity_code": "HVO204",
            "forecast_transfer_date": "2024/08/14"
        },
        {
            "opportunity_code": "HVO205",
            "forecast_transfer_date": "2024/08/14"
        },
        {
            "opportunity_code": "HVO301",
            "forecast_transfer_date": "2024/08/14"
        },
        {
            "opportunity_code": "HVO302",
            "forecast_transfer_date": "2024/08/14"
        },
        {
            "opportunity_code": "HVO303",
            "forecast_transfer_date": "2024/08/14"
        },
        {
            "opportunity_code": "HVO304",
            "forecast_transfer_date": "2024/08/14"
        },
        {
            "opportunity_code": "HVO305",
            "forecast_transfer_date": "2024/08/14"
        },
        {
            "opportunity_code": "HVP201",
            "forecast_transfer_date": "2024/05/17"
        },
        {
            "opportunity_code": "HVP203",
            "forecast_transfer_date": "2024/06/13"
        },
        {
            "opportunity_code": "HVP303",
            "forecast_transfer_date": "2024/06/13"
        }
    ]
    try:
        sales = list(db.cashflow_sales.find({}))
        print("sales", sales[0])
        print("sales", len(sales))
        for item in sales:
            item["_id"] = str(item["_id"])
            id = item["_id"]
            del item["_id"]
            print(item['opportunity_code'])
            new_dates_filtered = list(
                filter(lambda x: x['opportunity_code'] == item['opportunity_code'] and item['transferred'] == False,
                       new_dates))
            print(new_dates_filtered)
            if len(new_dates_filtered) > 0:
                item['forecast_transfer_date'] = new_dates_filtered[0]['forecast_transfer_date']
                db.cashflow_sales.update_one({"_id": ObjectId(id)}, {
                    "$set": {"forecast_transfer_date": new_dates_filtered[0]['forecast_transfer_date']}})
            else:
                item['forecast_transfer_date'] = "2025/01/31"
                db.cashflow_sales.update_one({"_id": ObjectId(id)}, {"$set": {"forecast_transfer_date": "2025/01/31"}})

        return {"success": True}
    except Exception as e:
        print("ERROR", e)
        return {"success": False}

        # result = db.opportunities.insert_many(data)
        # return {"success": True}


# update_cashflow_dates()
@cashflow.post("/trial-balance_cashflow")
async def trial_balance_cashflow(data: Request):
    request = await data.json()
    data = request['data']
    try:
        db.cashflow_trial_balance.delete_many({})
        result = db.cashflow_trial_balance.insert_many(data)
        return {"success": True}
    except Exception as e:
        return {"success": False, "error": str(e)}


@cashflow.get("/construction_cashflow")
async def get_construction_cashflow():
    try:
        data = list(db.cashflow_construction.find({}))
        # print(data[0])
        for item in data:
            item['_id'] = str(item['_id'])
        # data = list(data)
        # sort data by Whitebox-Able in descending order

        data = sorted(data, key=lambda x: (x['Whitebox-Able']), reverse=True)

        return {"success": True, "data": data}
    except Exception as e:
        print("ERROR", e)
        return {"success": False, "error": str(e)}


@cashflow.get("/trial-balance_cashflow")
async def get_trial_balance_cashflow():
    try:
        data = list(db.cashflow_trial_balance.find({}))
        for item in data:
            item['_id'] = str(item['_id'])
        # data = list(data)
        return {"success": True, "data": data}
    except Exception as e:
        return {"success": False, "error": str(e)}


@cashflow.post("/update_construction_data")
async def update_construction_data(data: Request):
    request = await data.json()
    data = request['data']
    _id = data["_id"]
    del data["_id"]
    try:
        db.cashflow_construction.update_one({"_id": ObjectId(_id)}, {"$set": data})
        return {"success": True}
    except Exception as e:
        return {"success": False, "error": str(e)}


# SALES HELPER FUNCTIONS
def get_sales_parameters():
    try:
        sales_parameters = list(db.salesParameters.find({}, {"_id": 0}))
        return sales_parameters
    except Exception as e:
        print("Error getting sales parameters")
        return {"success": False, "error": str(e)}


def get_rates():
    try:
        rates = list(db.rates.find({}, {"_id": 0}))
        # print("Rates", rates)
        for rate in rates:
            rate['rate'] = float(rate['rate'])
            rate['Efective_date'] = rate['Efective_date'].replace("-", "/")
            rate['Efective_date'] = datetime.strptime(rate['Efective_date'], '%Y/%m/%d')
        return rates
    except Exception as e:
        print("Error getting rates")
        return {"success": False, "error": str(e)}


def get_opportunities(data):
    try:
        opportunities = list(db.opportunities.find({"Category": {"$in": data}},
                                                   {"opportunity_code": 1, "Category": 1, "opportunity_end_date": 1,
                                                    "opportunity_final_transfer_date": 1, "opportunity_sale_price": 1,
                                                    "opportunity_sold": 1, "_id": 0}))
        for opportunity in opportunities:
            if opportunity.get('opportunity_final_transfer_date', "") == "":
                opportunity['transferred'] = False
            else:
                opportunity['transferred'] = True
        return opportunities
    except Exception as e:
        print("Error getting opportunities")
        return {"success": False, "error": str(e)}


def get_investors(data):
    try:
        final_investors = []

        investors = list(db.investors.find({}, {"_id": 0}))
        for investor in investors:
            investor_trust = list(
                filter(lambda x: x['Category'] in data and x['release_date'] == '', investor['trust']))
            investor_investments = list(
                filter(lambda x: x['Category'] in data, investor['investments']))
            if len(investor_trust) > 0:
                for item in investor_trust:
                    # print()
                    insert = {
                        "investor_acc_number": investor['investor_acc_number'],
                        "investor_name": investor['investor_name'],
                        "investor_surname": investor['investor_surname'],
                        "Category": item['Category'],
                        "opportunity_code": item['opportunity_code'],
                        "Block": item['opportunity_code'][-4],
                        "deposit_date": item['deposit_date'],
                        "release_date": item['release_date'],
                        "end_date": item.get('end_date', ""),
                        "investment_number": item.get('investment_number', 0),
                        "investment_amount": float(item['investment_amount']),
                        "investment_interest_rate": float(item['investment_interest_rate']),
                        "early_release": item.get('early_release', False),
                    }
                    final_investors.append(insert)

            for item in investor_investments:
                filtered_trust = list(filter(
                    lambda x: x['opportunity_code'] == item['opportunity_code'] and float(
                        x['investment_amount']) == float(item['investment_amount']), investor['trust']))

                insert = {
                    "investor_acc_number": investor['investor_acc_number'],
                    "investor_name": investor['investor_name'],
                    "investor_surname": investor['investor_surname'],
                    "Category": item['Category'],
                    "opportunity_code": item['opportunity_code'],
                    "Block": item['opportunity_code'][-4],
                    "deposit_date": filtered_trust[0]['deposit_date'],
                    "release_date": item['release_date'],
                    "end_date": item.get('end_date', ""),
                    "investment_number": item.get('investment_number', 0),
                    "investment_amount": float(item['investment_amount']),
                    "investment_interest_rate": float(item['investment_interest_rate']),
                    "early_release": item.get('early_release', False),
                }

                final_investors.append(insert)

        final_investors = [investor for investor in final_investors if
                           not (investor['investor_acc_number'] == "ZCAM01" and investor[
                               'opportunity_code'] == "HFA101" and investor['investment_amount'] == 400000.0)]

        final_investors = [investor for investor in final_investors if
                           not (investor['investor_acc_number'] == "ZJHO01" and investor[
                               'opportunity_code'] == "HFA304" and investor['investment_number'] == 1)]

        final_investors = [investor for investor in final_investors if
                           not (investor['investor_acc_number'] == "ZPJB01" and investor[
                               'opportunity_code'] == "HFA205" and investor['investment_number'] == 1)]

        final_investors = [investor for investor in final_investors if
                           not (investor['investor_acc_number'] == "ZERA01" and investor[
                               'opportunity_code'] == "EA205" and investor['investment_number'] == 3)]

        final_investors = [investor for investor in final_investors if
                           not (investor['investor_acc_number'] == "ZVOL01" and investor[
                               'opportunity_code'] == "EA103" and investor['investment_number'] == 3)]

        final_investors = [investor for investor in final_investors if
                           not (investor['investor_acc_number'] == "ZLEW03" and investor[
                               'opportunity_code'] == "EA205" and investor['investment_number'] == 1)]

        # sort final investors first by opportunity_code then by investor_acc_number
        final_investors = sorted(final_investors, key=lambda x: (x['opportunity_code'], x['investor_acc_number']))

        return final_investors
    except Exception as e:
        print("Error getting investors")
        return {"success": False, "error": str(e)}


def get_construction_costs():
    try:
        construction_costs = list(db.cashflow_construction.find({}, {"Complete Build": 1,
                                                                     "Whitebox-Able": 1, "Blocks": 1, "_id": 0}))
        # filter construction costs to only include whiteboxe-Able

        construction_costs = list(filter(lambda x: x['Whitebox-Able'] == True, construction_costs))
        construction_costs = list(filter(lambda x: x['Complete Build'] == True, construction_costs))
        return construction_costs
    except Exception as e:
        print("Error getting construction costs")
        return {"success": False, "error": str(e)}


def get_sales_for_cashflow(data):
    try:
        sales = list(db.sales_processed.find({"development": {"$in": data}},
                                             {"development": 1, "opportunity_code": 1, "opportunity_sales_date": 1,
                                              "opportunity_potential_reg_date": 1, "opportunity_actual_reg_date": 1,
                                              "opportunity_bond_registration": 1, "opportunity_commission": 1,
                                              "opportunity_transfer_fees": 1, "opportunity_trust_release_fee": 1,
                                              "opportunity_unforseen": 1, "_id": 0}))
        return sales
    except Exception as e:
        print("Error getting sales")
        return {"success": False, "error": str(e)}


def calculate_cashflow_investors(final_investors, opportunities, sales_data):
    # CALCULATE INITIAL DUE TO INVESTOR - final_investors, opportunities, rates
    final_interest_calculations = []
    opportunity_codes = []
    rates = get_rates()
    # print("final_investors", final_investors[0])
    try:
        for investor in final_investors:
            # investor['_id'] = str(investor['_id'])

            if investor['release_date'] == "":
                investor['not_released'] = True
                # convert investor['deposit_date'] to datetime
                investor['deposit_date'] = (investor['deposit_date'].replace("-", "/"))
                investor['deposit_date'] = datetime.strptime(investor['deposit_date'], '%Y/%m/%d')
                investor['release_date'] = investor['deposit_date'] + timedelta(days=30)
                filtered_opportunity = list(
                    filter(lambda x: x['opportunity_code'] == investor['opportunity_code'], opportunities))
                filtered_sales_data = list(
                    filter(lambda x: x['opportunity_code'] == investor['opportunity_code'], sales_data))
                if len(filtered_opportunity) > 0:
                    if filtered_sales_data[0]['forecast_transfer_date'] == "":
                        investor['end_date'] = filtered_opportunity[0]['opportunity_end_date']
                    else:
                        investor['end_date'] = filtered_sales_data[0]['forecast_transfer_date']
                    investor['end_date'] = (investor['end_date'].replace("-", "/"))
                    investor['end_date'] = datetime.strptime(investor['end_date'], '%Y/%m/%d')

            else:
                investor['not_released'] = False
                investor['deposit_date'] = (investor['deposit_date'].replace("-", "/"))
                investor['deposit_date'] = datetime.strptime(investor['deposit_date'], '%Y/%m/%d')
                investor['release_date'] = (investor['release_date'].replace("-", "/"))
                investor['release_date'] = datetime.strptime(investor['release_date'], '%Y/%m/%d')
                if investor['end_date'] == "":
                    filtered_opportunity = list(
                        filter(lambda x: x['opportunity_code'] == investor['opportunity_code'], opportunities))
                    filtered_sales_data = list(
                        filter(lambda x: x['opportunity_code'] == investor['opportunity_code'], sales_data))
                    if len(filtered_opportunity) > 0:
                        if filtered_sales_data[0]['forecast_transfer_date'] == "":
                            investor['end_date'] = filtered_opportunity[0]['opportunity_end_date']
                        else:
                            investor['end_date'] = filtered_sales_data[0]['forecast_transfer_date']

                investor['end_date'] = (investor['end_date'].replace("-", "/"))
                investor['end_date'] = datetime.strptime(investor['end_date'], '%Y/%m/%d')

            momentum_start_date = investor['deposit_date'] + timedelta(days=1)
            momentum_end_date = investor['release_date']

            interest = 0
            momentum_interest = 0
            investment_interest = 0
            while momentum_start_date <= momentum_end_date:
                rate = list(filter(lambda x: x['Efective_date'] <= momentum_start_date, rates))
                # sort by Efective_date descending
                rate = sorted(rate, key=lambda x: x['Efective_date'], reverse=True)
                rate = rate[0]['rate'] / 100
                momentum_interest += investor['investment_amount'] * rate / 365
                momentum_start_date += timedelta(days=1)
            investor['momentum_interest'] = momentum_interest
            investment_start_date = investor['release_date']
            investment_end_date = investor['end_date']
            # I NEED TO CHANGE THIS TO THE FORECAST DATE DATES ARE AN ISSUE
            # filtered_sales_data = list(
            #     filter(lambda x: x['opportunity_code'] == investor['opportunity_code'], sales_data))
            # if len(filtered_sales_data) > 0:
            #     trial_investment_end_date = filtered_sales_data[0]['forecast_transfer_date']
            #     if trial_investment_end_date != "":
            #         trial_investment_end_date = trial_investment_end_date.replace("-", "/")
            #         trial_investment_end_date = datetime.strptime(trial_investment_end_date, '%Y/%m/%d')
            #     if investor['opportunity_code'] == "HVC306":
            #         print("trial_investment_end_date", trial_investment_end_date)

            investment_interest += investor['investment_amount'] * investor['investment_interest_rate'] / 100 / 365 * (
                    investment_end_date - investment_start_date).days
            investor['investment_interest'] = investment_interest
            interest = momentum_interest + investment_interest
            investor['interest'] = interest

            opportunity_codes.append(investor['opportunity_code'])

        # Get unique values from opportunity_codes
        opportunity_codes = list(sorted(set(opportunity_codes)))
        # print("opportunity_codes",opportunity_codes)
        for opportunity in opportunity_codes:
            filtered_investors = list(filter(lambda x: x['opportunity_code'] == opportunity, final_investors))
            # Total_due_to_investors = sum of investment_amount + interest
            total_due_to_investors = sum([x['investment_amount'] + x['interest'] for x in filtered_investors])
            insert = {
                "opportunity_code": opportunity,
                "total_due_to_investors": total_due_to_investors,
            }
            final_interest_calculations.append(insert)

            # print(investor)
        # print(len(final_investors))
        # filter out from investors where early_release is true
        # final_investors = list(filter(lambda x: x['early_release'] == False, final_investors))
        # print(len(final_investors))
        # print("final_investors", final_investors[0])
        # print("final_interest_calculations", final_interest_calculations[0])
        return final_interest_calculations
    except Exception as e:
        print("Error calculating cashflow investors", e)
        return {"success": False, "error": str(e)}

    # I AM HERE


@cashflow.post("/get_sales_cashflow_initial")
async def get_sales_cashflow_initial(data: Request):
    request = await data.json()
    # print(request)
    data = request['data']
    try:
        start = time.time()
        opportunities = get_opportunities(data)

        sales_data = list(db.cashflow_sales.find({}))

        # print("sales_data", sales_data[0])
        # print()

        construction = get_construction_costs()
        for item in construction:
            item['block'] = item['Blocks'].replace("Block ", "")
            # remove leading and trailing whitespace from item['block']
            item['block'] = item['block'].strip()

            del item['Blocks']

        if len(sales_data) == len(opportunities):

            final_investors = get_investors(data)
            # print("final_investors", final_investors)
            investors_with_interest = calculate_cashflow_investors(final_investors, opportunities, sales_data)

            items_to_update = []
            for item in sales_data:

                if item['refinanced'] == False and item['VAT'] == 0:
                    final_investors_filtered = list(
                        filter(lambda x: x['opportunity_code'] == item['opportunity_code'], investors_with_interest))
                    # print("final_investors_filtered", final_investors_filtered)
                    if len(final_investors_filtered) > 0:
                        # print("final_investors_filtered", final_investors_filtered)
                        # print()
                        # print("item", item)
                        item['due_to_investors'] = final_investors_filtered[0]['total_due_to_investors']

                    else:
                        item['due_to_investors'] = 0

                    item['VAT'] = float(item['sale_price']) / 1.15 * 0.15
                    item['nett'] = float(item['sale_price']) / 1.15

                    item['transfer_income'] = float(item['sale_price']) - float(
                        item['opportunity_transfer_fees']) - float(item[
                                                                       'opportunity_trust_release_fee']) - float(
                        item['opportunity_unforseen']) - float(item[
                                                                   'opportunity_bond_registration']) - float(
                        item['opportunity_commission'])
                    item['profit_loss'] = item['transfer_income'] - item['due_to_investors']
                    items_to_update.append(item)

                item['_id'] = str(item['_id'])
                filtered_construction = list(
                    filter(lambda x: x['block'] == item['block'], construction))

                # print("filtered_construction", filtered_construction)

                if len(filtered_construction) > 0:

                    if item['complete_build']:
                        item['complete_build'] = True
                        items_to_update.append(item)  # I AM HERE Now I need to update the sales data
                else:

                    if not item['complete_build']:
                        item['complete_build'] = False

                        items_to_update.append(item)

                # DO SALES PRICE, PROFIT AND INVESTOR INTEREST CALCULATIONS ONLY IF DIFFERENT I AM HERE
                filtered_investors = list(filter(lambda x: x['opportunity_code'] == item['opportunity_code'],
                                                 investors_with_interest))
                if len(filtered_investors) > 0:
                    if item['due_to_investors'] == 0 or float(item['due_to_investors']) != float(filtered_investors[0][
                                                                                                     'total_due_to_investors']):
                        item['due_to_investors'] = filtered_investors[0]['total_due_to_investors']
                        item['profit_loss'] = item['transfer_income'] - item['due_to_investors']
                        items_to_update.append(item)
                filtered_opportunities = list(
                    filter(lambda x: x['opportunity_code'] == item['opportunity_code'], opportunities))
                if len(filtered_opportunities) > 0:
                    if float(item['sale_price']) != float(filtered_opportunities[0]['opportunity_sale_price']):
                        item['sale_price'] = filtered_opportunities[0]['opportunity_sale_price']
                        item['VAT'] = float(item['sale_price']) / 1.15 * 0.15
                        item['nett'] = float(item['sale_price']) / 1.15
                        item['transfer_income'] = (float(item['sale_price']) - item['opportunity_transfer_fees'] -
                                                   item['opportunity_trust_release_fee'] - item[
                                                       'opportunity_unforseen'] - item[
                                                       'opportunity_commission'] * float(item['sale_price']) -
                                                   item['opportunity_bond_registration'])
                        items_to_update.append(item)

            if len(items_to_update) > 0:

                items_to_update = list({v['_id']: v for v in items_to_update}.values())
                try:
                    for index, item in enumerate(items_to_update):
                        id1 = item['_id']

                        del item['_id']
                        db.cashflow_sales.update_one({"_id": ObjectId(id1)}, {"$set": item})
                        item['_id'] = id1
                except Exception as e:
                    print("Error updating sales data", e, index, item['opportunity_code'])
                    # print()
                    return {"success": False, "error": str(e)}

            end = time.time()
            for item in sales_data:
                # convert item['sale_price'] to currency format with R symbol
                item['sale_price_nice'] = "R" + "{:,.2f}".format(float(item['sale_price']))
                item['profit_loss_nice'] = "R" + "{:,.2f}".format(float(item['profit_loss']))

            # sort sales_data by transferred then by opportunity_code
            sales_data = sorted(sales_data, key=lambda x: (x['transferred'], x['opportunity_code']))

            # for index, data in enumerate(sales_data):
            #     if index == 0:
            #         print("sales_data", data['_id'])

            return {"success": True, "data": sales_data, "time": end - start}

        else:

            sales_parameters = get_sales_parameters()

            # rates = get_rates()

            sales = get_sales_for_cashflow(data)

            opportunities_sold = list(filter(lambda x: x['opportunity_sold'] == True, opportunities))

            sales = list(
                filter(lambda x: x['opportunity_code'] in [item['opportunity_code'] for item in opportunities_sold],
                       sales))

            construction = get_construction_costs()

            final_investors = get_investors(data)

            # Calculate INITIAL DATA
            final_sales = []
            for unit in opportunities:
                filtered_construction = list(
                    filter(lambda x: x['Blocks'] == "Block " + unit['opportunity_code'][-4], construction))
                if len(filtered_construction) > 0:
                    print("filtered_construction", filtered_construction[0])
                    complete_build = filtered_construction[0]['Complete Build']
                    # If whiteboxed is true, then make it false and vica a versa
                    complete_build = not complete_build
                else:
                    complete_build = False

                if unit['opportunity_final_transfer_date'] == "":
                    original_planned_transfer_date = unit['opportunity_end_date']
                    forecast_transfer_date = ""
                else:
                    original_planned_transfer_date = unit['opportunity_final_transfer_date']
                    forecast_transfer_date = unit['opportunity_final_transfer_date']
                filtered_sales = list(filter(lambda x: x['opportunity_code'] == unit['opportunity_code'], sales))
                if len(filtered_sales) > 0:
                    # print("Filtered Sales: ",filtered_sales)
                    opportunity_transfer_fees = float(filtered_sales[0].get('opportunity_transfer_fees', 0))
                    opportunity_trust_release_fee = float(filtered_sales[0].get('opportunity_trust_release_fee', 0))
                    opportunity_unforseen = float(filtered_sales[0].get('opportunity_unforseen', 0)) * float(
                        unit['opportunity_sale_price'])
                    opportunity_commission = float(filtered_sales[0].get('opportunity_commission', 0)) / 1.15 * float(
                        unit['opportunity_sale_price'])
                    opportunity_bond_registration = float(filtered_sales[0].get('opportunity_bond_registration', 0))
                    if opportunity_transfer_fees == 0:
                        # if unit['opportunity_code'] == "HVC306":
                        #     print("Filtered Sales: ", filtered_sales)
                        #     print()

                        filtered_sales_parameters = list(
                            filter(
                                lambda x: x['Development'] == unit['Category'] and x['Description'] == "transfer_fees",
                                sales_parameters))

                        opportunity_transfer_fees = float(filtered_sales_parameters[0]['rate'])
                        filtered_sales_parameters = list(
                            filter(
                                lambda x: x['Development'] == unit['Category'] and x[
                                    'Description'] == "trust_release_fee",
                                sales_parameters))
                        opportunity_trust_release_fee = float(filtered_sales_parameters[0]['rate'])
                        filtered_sales_parameters = list(
                            filter(
                                lambda x: x['Development'] == unit['Category'] and x['Description'] == "unforseen",
                                sales_parameters))

                        opportunity_unforseen = float(filtered_sales_parameters[0]['rate']) * float(
                            unit['opportunity_sale_price'])
                        filtered_sales_parameters = list(
                            filter(
                                lambda x: x['Development'] == unit['Category'] and x['Description'] == "commission",
                                sales_parameters))
                        # if unit['opportunity_code'] == "HVC306":
                        #     print("filtered_sales_parameters", filtered_sales_parameters)
                        opportunity_commission = float(filtered_sales_parameters[0]['rate']) / 1.15 * float(
                            unit['opportunity_sale_price'])
                        filtered_sales_parameters = list(
                            filter(
                                lambda x: x['Development'] == unit['Category'] and x[
                                    'Description'] == "bond_registration",
                                sales_parameters))
                        opportunity_bond_registration = float(filtered_sales_parameters[0]['rate'])

                opportunity_unforseen = float(
                    unit['opportunity_sale_price']) * .005

                opportunity_commission = float(
                    unit['opportunity_sale_price']) * .05

                insert = {
                    "Category": unit['Category'],
                    "block": unit['opportunity_code'][-4],
                    "opportunity_code": unit['opportunity_code'],
                    "sold": unit['opportunity_sold'],
                    "transferred": unit['transferred'],
                    "complete_build": complete_build,
                    "original_planned_transfer_date": original_planned_transfer_date,
                    "forecast_transfer_date": forecast_transfer_date,
                    "sale_price": float(unit['opportunity_sale_price']),
                    "VAT": float(unit['opportunity_sale_price']) / 1.15 * 0.15,
                    "nett": float(unit['opportunity_sale_price']) / 1.15,
                    "opportunity_transfer_fees": opportunity_transfer_fees,
                    "opportunity_trust_release_fee": opportunity_trust_release_fee,
                    "opportunity_unforseen": opportunity_unforseen,
                    "opportunity_commission": opportunity_commission,
                    "opportunity_bond_registration": opportunity_bond_registration,
                    "transfer_income": float(unit[
                                                 'opportunity_sale_price']) - opportunity_transfer_fees - opportunity_trust_release_fee - opportunity_unforseen - opportunity_commission - opportunity_bond_registration,
                    "due_to_investors": 0,
                    "profit_loss": 0,
                    "refinanced": False,
                }
                final_sales.append(insert)

            # print("final_sales", len(final_sales))

            if len(sales_data) == 0:
                try:
                    db.cashflow_sales.insert_many(final_sales)
                except Exception as e:
                    print("Error inserting sales data", e)
                    return {"success": False, "error": str(e)}
            else:
                if len(sales_data) < len(final_sales):
                    for sale in final_sales:
                        if sale not in sales_data:
                            try:
                                db.cashflow_sales.insert_one(sale)
                            except Exception as e:
                                print("Error inserting sales data", e)
                                return {"success": False, "error": str(e)}

            # print("insert", insert)
            # print()

        # print("final_investors", final_investors[10])
        # print("final_investors", len(final_investors))
        # print("opportunities", opportunities[0])
        # print("sales_data", sales_data)
        # print("sales_parameters", sales_parameters)
        # print("rates", rates)
        # print("investors", investors[11])
        # print("sales", sales[0])
        end = time.time()
        print("Time taken", end - start)
        return {"success": True, "data": sales_data, "time": end - start}
    except Exception as e:
        print("Error getting sales cashflow initial", e)
        return {"success": False, "error": str(e)}
    # return {"success": True, "data": data}


@cashflow.post("/update_sales_data")
async def update_sales_data(data: Request):
    request = await data.json()
    data = request['data']
    for item in data:
        _id = item["_id"]
        del item["_id"]
        try:
            db.cashflow_sales.update_one({"_id": ObjectId(_id)}, {"$set": item})
            print("Updated")
        except Exception as e:
            print("Error updating sales data", e)
            return {"success": False, "error": str(e)}
    return {"success": True}


def calculate_vat_due(sale_date):
    vat_periods = {
        1: "03/31",
        2: "03/31",
        3: "05/31",
        4: "05/31",
        5: "07/31",
        6: "07/31",
        7: "09/30",
        8: "09/30",
        9: "11/30",
        10: "11/30",
        11: "01/31",
        12: "01/31",
    }

    sale_date = datetime.strptime(sale_date.replace("-", "/"), '%Y/%m/%d')

    sale_month = sale_date.month
    sale_year = sale_date.year
    if sale_month > 10:
        vat_year = sale_year + 1
    else:
        vat_year = sale_year

    vat_date = f"{vat_year}/{vat_periods[sale_month]}"
    # print("vat_date", vat_date)
    return vat_date


# FUNCTIONS TO GENERATE REPORT


def investors_new_cashflow_nsst_report():
    # start = time.time()
    investors = list(db.investors.find({}, {"_id": 0}))
    for investor in investors:
        # print("filtered_opportunityInvest", investor['investments'])
        # filter investor["Trust"] where Category = "Heron Fields" or Category = "Heron View"
        investor["trust"] = list(
            filter(lambda x: x["Category"] == "Heron Fields" or x["Category"] == "Heron View", investor["trust"]))
        # filter investor["Investments"] where Category = "Heron Fields" or Category = "Heron View"

        investor["investments"] = list(
            filter(lambda x: x["Category"] == "Heron Fields" or x["Category"] == "Heron View", investor["investments"]))
        # filter investor["pledges"] where Category = "Heron Fields" or Category = "Heron View"
        investor["pledges"] = list(
            filter(lambda x: x["Category"] == "Heron Fields" or x["Category"] == "Heron View", investor["pledges"]))

        # filter out from investor["trust"] where release_date is not empty
    # filter out of investors where trust is empty
    investors = list(filter(lambda x: len(x["trust"]) > 0, investors))

    # filter out of investors where investments is empty
    # I AM HERE
    # if investor['investor_acc_number'] == "ZMAN01":
    #     print("filtered_opportunityTrust", investor['trust'])
    #     print()

    # investors = list(filter(lambda x: len(x["investments"]) > 0, investors))

    opportunities = list(db.opportunities.find({}, {"_id": 0}))

    opportunities = list(
        filter(lambda x: x["Category"] == "Heron Fields" or x["Category"] == "Heron View", opportunities))

    final_investors = []

    for investor in investors:
        # filter out from investor["trust"] where release_date is not empty

        trust_filtered = list(filter(lambda x: x.get("release_date", "") == "", investor["trust"]))
        if len(trust_filtered) > 0:
            for trust in trust_filtered:
                filtered_opportunity = list(
                    filter(lambda x: x["opportunity_code"] == trust["opportunity_code"], opportunities))

                # sold = False
                # transferred = False
                # if len(filtered_opportunity) > 0:
                # print("filtered_opportunity", filtered_opportunity[0])
                sold = filtered_opportunity[0]["opportunity_sold"]
                if filtered_opportunity[0]["opportunity_final_transfer_date"] != "":
                    transferred = True
                else:
                    transferred = False

                insert = {
                    "investor_acc_number": investor["investor_acc_number"],
                    "investor_name": investor["investor_name"],
                    "investor_surname": investor["investor_surname"],
                    "Category": trust["Category"],
                    "opportunity_code": trust["opportunity_code"],
                    "sold": sold,
                    "tansferred": transferred,
                    "Block": trust["opportunity_code"][-4],
                    "deposit_date": trust["deposit_date"],
                    "release_date": trust["release_date"],
                    "end_date": trust.get("end_date", ""),
                    "investment_number": trust.get("investment_number", 0),
                    "investment_amount": float(trust["investment_amount"]),
                    "investment_interest_rate": float(trust["investment_interest_rate"]),
                    "early_release": trust.get("early_release", False),
                    "still_pledged": False,
                }
                final_investors.append(insert)

        if len(investor["investments"]) > 0:
            for investment in investor["investments"]:
                filtered_trust = list(filter(
                    lambda x: x["opportunity_code"] == investment["opportunity_code"] and float(
                        x["investment_amount"]) == float(investment["investment_amount"]), investor["trust"]))

                filtered_opportunity = list(
                    filter(lambda x: x["opportunity_code"] == investment["opportunity_code"], opportunities))

                sold = filtered_opportunity[0]["opportunity_sold"]
                # if filtered_opportunity[0]["opportunity_code"] == "HVC106":
                #     print("filtered_opportunity", filtered_opportunity)
                #     print()
                if filtered_opportunity[0]["opportunity_final_transfer_date"] != "":
                    transferred = True
                else:
                    transferred = False

                insert = {
                    "investor_acc_number": investor["investor_acc_number"],
                    "investor_name": investor["investor_name"],
                    "investor_surname": investor["investor_surname"],
                    "Category": investment["Category"],
                    "opportunity_code": investment["opportunity_code"],
                    "sold": sold,
                    "tansferred": transferred,
                    "Block": investment["opportunity_code"][-4],
                    "deposit_date": filtered_trust[0]["deposit_date"],
                    "release_date": investment["release_date"],
                    "end_date": investment.get("end_date", ""),
                    "investment_number": investment.get("investment_number", 0),
                    "investment_amount": float(investment["investment_amount"]),
                    "investment_interest_rate": float(investment["investment_interest_rate"]),
                    "early_release": investment.get("early_release", False),
                    "still_pledged": False,
                }
                final_investors.append(insert)

        if len(investor["pledges"]) > 0:
            for pledge in investor["pledges"]:

                filtered_opportunity = list(
                    filter(lambda x: x["opportunity_code"] == pledge["opportunity_code"], opportunities))
                sold = filtered_opportunity[0]["opportunity_sold"]

                # print("filtered_opportunity", filtered_opportunity[0])
                if filtered_opportunity[0]["opportunity_final_transfer_date"] != "":
                    transferred = True
                else:
                    transferred = False

                insert = {
                    "investor_acc_number": investor["investor_acc_number"],
                    "investor_name": investor["investor_name"],
                    "investor_surname": investor["investor_surname"],
                    "Category": pledge["Category"],
                    "opportunity_code": pledge["opportunity_code"],
                    "sold": sold,
                    "tansferred": transferred,
                    "Block": pledge["opportunity_code"][-4],
                    "deposit_date": pledge["deposit_date"],
                    "release_date": pledge["release_date"],
                    "end_date": pledge.get("end_date", ""),
                    "investment_number": pledge.get("investment_number", 0),
                    "investment_amount": float(pledge["investment_amount"]),
                    "investment_interest_rate": float(pledge["investment_interest_rate"]),
                    "early_release": pledge.get("early_release", False),
                    "still_pledged": True,
                }
                final_investors.append(insert)

    rates = list(db.rates.find({}, {"_id": 0}))
    for rate in rates:
        rate["rate"] = float(rate["rate"])
        rate["Efective_date"] = rate["Efective_date"].replace("-", "/")
        rate["Efective_date"] = datetime.strptime(rate["Efective_date"], "%Y/%m/%d")

    # sort rates by Efective_date decending
    rates = sorted(rates, key=lambda x: x['Efective_date'], reverse=True)

    sales_parameters = list(db.salesParameters.find({}, {"_id": 0}))
    sales_parameters = list(
        filter(lambda x: x["Development"] == "Heron Fields" or x["Development"] == "Heron View", sales_parameters))
    for param in sales_parameters:
        param["rate"] = float(param["rate"])
        param["Effective_date"] = param["Effective_date"].replace("-", "/")
        param["Effective_date"] = datetime.strptime(param["Effective_date"], "%Y/%m/%d")

    sales = list(db.sales_processed.find({}, {"_id": 0}))
    sales = list(filter(lambda x: x["development"] == "Heron Fields" or x["development"] == "Heron View", sales))
    # print("sales", sales[0])

    # FUNNIES

    final_investors = [investor for investor in final_investors if
                       not (investor['investor_acc_number'] == "ZCAM01" and investor[
                           'opportunity_code'] == "HFA101" and investor['investment_amount'] == 400000.0)]

    final_investors = [investor for investor in final_investors if
                       not (investor['investor_acc_number'] == "ZJHO01" and investor[
                           'opportunity_code'] == "HFA304" and investor['investment_number'] == 1)]

    final_investors = [investor for investor in final_investors if
                       not (investor['investor_acc_number'] == "ZPJB01" and investor[
                           'opportunity_code'] == "HFA205" and investor['investment_number'] == 1)]

    final_investors = [investor for investor in final_investors if
                       not (investor['investor_acc_number'] == "ZERA01" and investor[
                           'opportunity_code'] == "EA205" and investor['investment_number'] == 3)]

    final_investors = [investor for investor in final_investors if
                       not (investor['investor_acc_number'] == "ZVOL01" and investor[
                           'opportunity_code'] == "EA103" and investor['investment_number'] == 3)]

    final_investors = [investor for investor in final_investors if
                       not (investor['investor_acc_number'] == "ZLEW03" and investor[
                           'opportunity_code'] == "EA205" and investor['investment_number'] == 1)]

    for final in final_investors:
        if final['end_date'] == "":
            filtered_opportunity = list(
                filter(lambda x: x['opportunity_code'] == final['opportunity_code'], opportunities))
            if len(filtered_opportunity) > 0:
                if filtered_opportunity[0]['opportunity_final_transfer_date'] != "":
                    final['end_date'] = filtered_opportunity[0]['opportunity_final_transfer_date']
                else:
                    final['end_date'] = filtered_opportunity[0]['opportunity_end_date']

        final['deposit_date'] = final['deposit_date'].replace("-", "/")
        if final['deposit_date'] == "":
            # final['deposit_date'] = today
            final['deposit_date'] = datetime.now()
            # final['deposit_date'] = data['deposit_date']
            # print("final['deposit_date']", final['deposit_date'])
            # print("final", final)
        else:
            final['deposit_date'] = datetime.strptime(final['deposit_date'], '%Y/%m/%d')
        if final['release_date'] != "":
            final['release_date'] = final['release_date'].replace("-", "/")
            final['release_date'] = datetime.strptime(final['release_date'], '%Y/%m/%d')

        final['end_date'] = final['end_date'].replace("-", "/")
        final['end_date'] = datetime.strptime(final['end_date'], '%Y/%m/%d')

        momentum_start = final['deposit_date'] + timedelta(days=1)
        momentum_end = final['release_date']
        # released_start = momentum_end.replace("-", "/")
        end_date = final['end_date']
        if momentum_end != "":
            momentum_end = momentum_end
        else:
            momentum_end = final['deposit_date'] + timedelta(days=30)

        momentum_interest = 0
        while momentum_start <= momentum_end:
            rate = list(filter(lambda x: x['Efective_date'] <= momentum_start, rates))
            rate = sorted(rate, key=lambda x: x['Efective_date'], reverse=True)
            rate = rate[0]['rate'] / 100
            momentum_interest += final['investment_amount'] * rate / 365
            momentum_start += timedelta(days=1)

        # print("momentum_interest", momentum_interest)
        investment_interest = 0
        investment_start = momentum_end  # I AM HERE
        investment_end = final['end_date']
        investment_interest += final['investment_amount'] * final['investment_interest_rate'] / 100 / 365 * (
                investment_end - investment_start).days
        # print("investment_interest", investment_interest)

        final['momentum_interest'] = momentum_interest
        final['investment_interest'] = investment_interest
        final['interest'] = momentum_interest + investment_interest

    # sort by Category then by opportunity_code then by investor_acc_number
    final_investors = sorted(final_investors,
                             key=lambda x: (x['Category'], x['opportunity_code'], x['investor_acc_number']))
    # print("final", final_investors[158])

    # end = time.time()

    return final_investors


# investors_new_cashflow_nsst_report()

def get_construction_costs():
    # try:
    construction_costs = list(db.cashflow_construction.find({}))
    for cost in construction_costs:
        if cost['Complete Build']:
            cost['Complete Build'] = 1
        else:
            cost['Complete Build'] = 0
        del [cost['_id']]
        for key, value in cost.items():
            # if value begins with R then remove the R and all asci characters and convert to float
            if str(value).startswith("R"):
                try:
                    # print(key, cost[key])
                    cleaned_string = re.sub(r'[^\d.,]', '', value)
                    float_value = float(cleaned_string.replace(',', '.'))
                    cost[key] = float_value
                    # cost[key] = cost[key].replace(" ", "")

                    # print(key, cost[key])
                    # print()
                    # cost[key] = float(re.sub(r'\D.', '', value))
                    # print(key,cost[key])
                    # print()
                except Exception as e:
                    # print("Error converting construction cost to float", e)
                    # print(cost)
                    continue
                    # cost[key] = value where the "R" is replaced with "", spaces are replaced with "" and "," is replaced with "."




                except Exception as e:
                    # print("Error converting construction cost to float", e)
                    # print(cost)
                    continue

    # sort by whitebox-able then by Blocks
    # print("construction_costs", construction_costs[0])
    construction_costs = sorted(construction_costs, key=lambda x: (x['Whitebox-Able'], x['Blocks']))

    # print("construction_costs", construction_costs[0])
    return construction_costs
    # except Exception as e:
    #     print("Error getting construction costs")
    #     return {"success": False, "error": str(e)}


def get_sales_data():
    try:
        construction_data = list(db.cashflow_construction.find({}, {"_id": 0}))
        for item in construction_data:
            item['block'] = item['Blocks'].replace("Block ", "")
            # remove leading and trailing whitespace from item['block']
            item['block'] = item['block'].strip()
            del item['Blocks']
        # filter construction_data by "Whitebox-Able" is True

        construction_data = list(filter(lambda x: x['Whitebox-Able'] == True, construction_data))

        # print("construction_data", construction_data[0])
        sales_data = list(db.cashflow_sales.find({}, {"_id": 0}))
        # print()
        # print("sales_data", sales_data[0])
        for sale in sales_data:
            construction_data_filtered = list(filter(lambda x: x['block'] == sale['block'], construction_data))
            # print(sale['block'],construction_data_filtered )
            if len(construction_data_filtered) > 0:
                if construction_data_filtered[0]['Complete Build'] == False and construction_data_filtered[0][
                    'Whitebox-Able'] == True:
                    sale['complete_build'] = False
                elif construction_data_filtered[0]['Complete Build'] == True and construction_data_filtered[0][
                    'Whitebox-Able'] == True:
                    sale['complete_build'] = True
            else:
                sale['complete_build'] = True

            if sale['complete_build']:
                sale['complete_build'] = 1
            else:
                sale['complete_build'] = 0

            # print("sales_data", sales_data[0])
            if 'profit_loss_nice' in sale:
                del sale['profit_loss_nice']
            if 'sale_price_nice' in sale:
                del sale['sale_price_nice']
            if sale['forecast_transfer_date'] == "":
                sale['vat_date'] = calculate_vat_due(sale['original_planned_transfer_date'])
            else:
                sale['vat_date'] = calculate_vat_due(sale['forecast_transfer_date'])
            # convert original_planned_transfer_date to datetime
            sale['original_planned_transfer_date'] = sale['original_planned_transfer_date'].replace("-", "/")
            sale['original_planned_transfer_date'] = datetime.strptime(sale['original_planned_transfer_date'],
                                                                       '%Y/%m/%d')
            # if forecast_transfer_date is not empty then convert to datetime
            if sale['forecast_transfer_date'] != "":
                sale['forecast_transfer_date'] = sale['forecast_transfer_date'].replace("-", "/")
                sale['forecast_transfer_date'] = datetime.strptime(sale['forecast_transfer_date'], '%Y/%m/%d')
            # convert vat_date to datetime
            sale['vat_date'] = sale['vat_date'].replace("/", "-")
            sale['vat_date'] = datetime.strptime(sale['vat_date'], '%Y-%m-%d')

            # del sale['profit_loss_nice']
            # del sale['sale_price_nice']

        return sales_data
    except Exception as e:
        print("Error getting sales data")
        return {"success": False, "error": str(e)}


def get_operational_costs():
    try:
        operational_costs = list(db.cashflow_trial_balance.find({}, {"_id": 0}))
        for cost in operational_costs:

            for key, value in cost.items():
                # if value begins with R then remove the R and all asci characters and convert to float
                if str(value).startswith("R"):
                    try:
                        cost[key] = float(re.sub(r'[^\d.]', '', value)) / 100
                    except Exception as e:
                        # print("Error converting operational cost to float", e)
                        # print(cost)
                        continue
        # print("operational_costs", operational_costs)

        return operational_costs
    except Exception as e:
        print("Error getting operational costs", e)
        return {"success": False, "error": str(e)}


def get_xero_tbs():
    try:
        xero_tbs = list(db.cashflow_xero_tb.find({}, {"_id": 0}))

        # sort by ReportDate, AccountCode, AccountName, ReportTitle
        xero_tbs = sorted(xero_tbs,
                          key=lambda x: (x['ReportDate'], x['AccountCode'], x['AccountName'], x['ReportTitle']))
        # print("xero_tbs", xero_tbs[0])
        # if 'ReportTitle' not 'Heron Fields (Pty) Ltd' or 'Cape Projects Construction (Pty) Ltd' then remove all other items if 'Category' is not 'Assets'
        xero_tbs1 = list(
            filter(
                lambda x: x['ReportTitle']
                          in [
                              'Heron Fields (Pty) Ltd',
                              'Heron View (Pty) Ltd',
                              'Cape Projects Construction (Pty) Ltd',
                          ],
                xero_tbs,
            )
        )
        xero_tbs2 = list(
            filter(
                lambda x: x['ReportTitle']
                          not in [
                              'Heron Fields (Pty) Ltd',
                              'Heron View (Pty) Ltd',
                              'Cape Projects Construction (Pty) Ltd',
                          ],
                xero_tbs,
            )
        )
        # filter from xero_tbs2 where Category is not 'Assets'
        xero_tbs2 = list(filter(lambda x: x['Category'] == 'Assets', xero_tbs2))
        # filter out of xero_tbs2 where AccountCode does not begin with 84
        xero_tbs2 = list(filter(lambda x: str(x['AccountCode']).startswith("84"), xero_tbs2))
        xero_tbs2 = list(filter(lambda x: not x['AccountCode'].startswith("8480"), xero_tbs2))
        # print("xero_tbs1", xero_tbs1[0])
        # print()
        # print("xero_tbs2", xero_tbs2[0])

        xero_tbs = xero_tbs1 + xero_tbs2
        # print()
        # print("xero_tbs", xero_tbs[0:2])
        for item in xero_tbs:
            # convert ReportDate to datetime
            item['ReportDate'] = item['ReportDate'].replace("-", "/")
            item['ReportDate'] = datetime.strptime(item['ReportDate'], '%Y/%m/%d')

        return xero_tbs
    except Exception as e:
        print("Error getting xero tbs", e)
        return {"success": False, "error": str(e)}


def get_opportunities():
    try:
        # get opportunities from db where Category equals "Heron Fields" or "Heron View"
        opportunities = list(db.opportunities.find({}, {"_id": 0}))
        opportunities = list(
            filter(lambda x: x["Category"] == "Heron Fields" or x["Category"] == "Heron View", opportunities))
        # sort by Category then by opportunity_code
        opportunities = sorted(opportunities, key=lambda x: (x['Category'], x['opportunity_code']))
        # print("opportunities", opportunities[0])
        # print()
        # get investors from db
        investors = list(db.investors.find({}, {"_id": 0}))
        for investor in investors:
            # filter investor["Trust"] where Category = "Heron Fields" or Category = "Heron View"
            investor["trust"] = list(
                filter(lambda x: x["Category"] == "Heron Fields" or x["Category"] == "Heron View", investor["trust"]))
            # filter investor["Investments"] where Category = "Heron Fields" or Category = "Heron View"
            # investor["investments"] = list(
            #     filter(lambda x: x["Category"] == "Heron Fields" or x["Category"] == "Heron View",
            #            investor["investments"]))
            # filter investor["pledges"] where Category = "Heron Fields" or Category = "Heron View"
            investor["pledges"] = list(
                filter(lambda x: x["Category"] == "Heron Fields" or x["Category"] == "Heron View", investor["pledges"]))

            # REMEMBER

        # filter out of investors where trust is empty
        investors = list(filter(lambda x: len(x["trust"]) > 0, investors))
        final_investors = []
        for index, investor in enumerate(investors):
            for amount in investor['trust']:
                insert = {
                    "investment_amount": float(amount['investment_amount']),
                    "opportunity_code": amount['opportunity_code'],
                    "Category": amount['Category'],
                    "investor_acc_number": investor['investor_acc_number'],
                    "investment_number": amount.get("investment_number", 0),
                    # "pledged": False
                }
                final_investors.append(insert)
            # if len(investor['pledges']) > 0:
            #     for amount in investor['pledges']:
            #         insert = {
            #             "investment_amount": float(amount['investment_amount']),
            #             "opportunity_code": amount['opportunity_code'],
            #             "Category": amount['Category'],
            #             "investor_acc_number": investor['investor_acc_number'],
            #             "investment_number": amount.get("investment_number", 0),
            #             "pledged": True
            #         }
            #         final_investors.append(insert)

        final_investors = [investor for investor in final_investors if
                           not (investor['investor_acc_number'] == "ZCAM01" and investor[
                               'opportunity_code'] == "HFA101" and investor['investment_amount'] == 400000.0)]

        final_investors = [investor for investor in final_investors if
                           not (investor['investor_acc_number'] == "ZJHO01" and investor[
                               'opportunity_code'] == "HFA304" and investor['investment_number'] == 1)]

        final_investors = [investor for investor in final_investors if
                           not (investor['investor_acc_number'] == "ZPJB01" and investor[
                               'opportunity_code'] == "HFA205" and investor['investment_number'] == 1)]

        # print("final_investors", final_investors[0])
        # print()
        # print("length", len(final_investors))

        final_opportunities = []

        for opportunity in opportunities:
            # convert opportunity_end_date to datetime
            opportunity['opportunity_end_date'] = opportunity['opportunity_end_date'].replace("-", "/")
            opportunity['opportunity_end_date'] = datetime.strptime(opportunity['opportunity_end_date'], '%Y/%m/%d')
            # convert opportunity_final_transfer_date to datetime
            if opportunity['opportunity_final_transfer_date'] != "":
                opportunity['transferred'] = True
                opportunity['opportunity_final_transfer_date'] = opportunity['opportunity_final_transfer_date'].replace(
                    "-",
                    "/")
                opportunity['opportunity_final_transfer_date'] = datetime.strptime(
                    opportunity['opportunity_final_transfer_date'], '%Y/%m/%d')
            else:
                opportunity['transferred'] = False
            insert = {
                "Category": opportunity['Category'],
                "opportunity_code": opportunity['opportunity_code'],
                "sold": opportunity['opportunity_sold'],
                "transferred": opportunity['transferred'],
                "opportunity_amount_required": float(opportunity['opportunity_amount_required']),
                "opportunity_end_date": opportunity['opportunity_end_date'],
                "opportunity_final_transfer_date": opportunity['opportunity_final_transfer_date'],
            }
            # filter final_investors where opportunity_code is equal to opportunity['opportunity_code']
            filtered_investors = list(
                filter(lambda x: x['opportunity_code'] == opportunity['opportunity_code'], final_investors))
            # sum the investment_amounts in filtered_investors
            investment_amount = sum([x['investment_amount'] for x in filtered_investors])
            insert['investment_amount'] = investment_amount
            final_opportunities.append(insert)

        # print("final_opportunities", final_opportunities[0])
        # print()
        # print("final_opportunities", final_opportunities[1])
        # print()
        # print("final_opportunities", final_opportunities[2])
        # print()
        # print("final_opportunities", final_opportunities[4])

        return final_opportunities
    except Exception as e:
        print("Error getting opportunities", e)
        return {"success": False, "error": str(e)}





@cashflow.post("/generate_investors_new_cashflow_nsst_report")
async def generate_investors_new_cashflow_nsst_report(data: Request):
    request = await data.json()
    date = request['date']
    # print(date)

    try:
        start = time.time()
        invest = investors_new_cashflow_nsst_report()
        construction = get_construction_costs()
        sales = get_sales_data()
        operational_costs = get_operational_costs()
        xero = get_xero_tbs()
        opportunities = get_opportunities()
        result = cashflow_projections(invest, construction, sales, operational_costs, xero,opportunities, date)
        # result = "Awesome"
        end = time.time()
        print("Time taken", end - start)
        return {"success": True, "Result": result, "time": end - start}
    except Exception as e:
        print("Error generating investors new cashflow nsst report", e)
        return {"success": False, "error": str(e)}


def adjust_sales():
    # get cashflow_sales from db
    cashflow_sales = list(db.cashflow_sales.find({}))
    for sale in cashflow_sales:
        sale['_id'] = str(sale['_id'])
    # filter cashflow sales where transferred is false and forecast_transfer_date = ""
    filtered_sales = list(
        filter(lambda x: x['transferred'] == False and x['forecast_transfer_date'] == "", cashflow_sales))
    for sale in filtered_sales:
        sale['forecast_transfer_date'] = sale['original_planned_transfer_date']
        try:
            id = sale["_id"]
            del sale["_id"]
            db.cashflow_sales.update_one({"_id": ObjectId(id)}, {"$set": sale})
            print("Success")
        except Exception as e:
            print(e)
    # print(filtered_sales[0])

    # print(cashflow_sales[0])


@cashflow.get("/get_last_date")
async def get_last_date():
    # print("Hello")
    try:
        dates = list(db.cashflow_xero_tb.find({}, {"_id": 0}))
        dates_array = []
        for date in dates:
            # convert date['ReportDate'] to datetime
            date['ReportDate'] = date['ReportDate'].replace("-", "/")
            date['ReportDate'] = datetime.strptime(date['ReportDate'], '%Y/%m/%d')

            dates_array.append(date['ReportDate'])

        dates_array = sorted(dates_array, reverse=True)
        # print("dates_array", dates_array)
        final_date = dates_array[0]

        # print("dates", dates_array)

        # cashflow_sales = sorted(cashflow_sales, key=lambda x: x['forecast_transfer_date'], reverse=True)
        # last_date = cashflow_sales[0]['forecast_transfer_date']
        return {"success": True, "last_date": final_date}
    except Exception as e:
        print("Error getting last date", e)
        return {"success": False, "error": str(e)}


@cashflow.get("/get_cash_projection")
async def get_cash_projection(file_name):
    try:
        print(file_name)
        file_directory = "cashflow_p&l_files"
        # file_name = f"{file_directory}/{file_name}"
        print(file_name)

        # file_name = file_name + ".xlsx"
        # # file_name = file_name.replace("_", " ")
        # # file_name = file_name.replace("xx", "&")
        #
        # dir_path = "early_releases_excel_generation"
        dir_list = os.listdir(file_directory)
        print(dir_list)

        if file_name in dir_list:

            return FileResponse(f"{file_directory}/{file_name}", filename=file_name)
        else:
            return {"ERROR": "File does not exist!!"}
    except Exception as e:
        print(e)
        return {"ERROR": "Please Try again"}

# @cashflow.post("/get_sales_data")
# async def get_sales_data(data: Request):
#     request = await data.json()
#     date = request['date']
#     try:
#         sales_data = list(db.cashflow_sales.find({}, {"_id": 0}))
#         for sale in sales_data:
#             sale['original_planned_transfer_date'] = sale['original_planned_transfer_date'].replace("-", "/")
#             sale['original_planned_transfer_date'] = datetime.strptime(sale['original_planned_transfer_date'], '%Y/%m/%d')
#             if sale['forecast_transfer_date'] != "":
#                 sale['forecast_transfer_date'] = sale['forecast_transfer_date'].replace("-", "/")
#                 sale['forecast_transfer_date'] = datetime.strptime(sale['forecast_transfer_date'], '%Y/%m/%d')
#             sale['vat_date'] = calculate_vat_due(sale['forecast_transfer_date'])
#             sale['vat_date'] = sale['vat_date'].replace("/", "-")
#             sale['vat_date'] = datetime.strptime(sale['vat_date'], '%Y-%m-%d')
#         return {"success": True, "data": sales_data}
#     except Exception as e:
#         print("Error getting sales data", e)
#         return {"success": False, "error": str(e)}

# adjust_sales()
