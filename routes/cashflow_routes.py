import copy
import json
import os
import re
from datetime import datetime, timedelta

import pandas as pd
from fastapi import APIRouter, Request, BackgroundTasks
from fastapi.responses import FileResponse
from configuration.db import db
from bson.objectid import ObjectId
import time
from cashflow_excel_functions.cashflow_projection_nsst import cashflow_projections
from cashflow_excel_functions.daily_cashflow import daily_cashflow

cashflow = APIRouter()


def recalc_interest():
    try:
        sales = list(db.sales_processed.find({}, {"_id": 0, "development": 1, "opportunity_code": 1,
                                                  "opportunity_potential_reg_date": 1,
                                                  "opportunity_actual_reg_date": 1}))
        # sort sales where development is Heron Fields or Heron View
        sales = list(filter(lambda x: x['development'] == "Heron Fields" or x['development'] == "Heron View", sales))
        for sale in sales:
            if sale['opportunity_potential_reg_date'] != "" and sale['opportunity_potential_reg_date'] != None:
                sale['opportunity_potential_reg_date'] = datetime.strptime(
                    sale['opportunity_potential_reg_date'].replace("-", "/"), "%Y/%m/%d")
            else:
                sale['opportunity_potential_reg_date'] = ""
            if sale['opportunity_actual_reg_date'] != "" and sale['opportunity_actual_reg_date'] != None:
                sale['opportunity_actual_reg_date'] = datetime.strptime(
                    sale['opportunity_actual_reg_date'].replace("-", "/"), "%Y/%m/%d")
            else:
                sale['opportunity_actual_reg_date'] = ""
        print("SALES", sales[0])
        print()
        rates = list(db.rates.find({}, {"_id": 0}))
        for rate in rates:
            rate['rate'] = float(rate['rate'])
            rate['Efective_date'] = datetime.strptime(rate['Efective_date'].replace("-", "/"), "%Y/%m/%d")
        # sort rates by Efective_date descending
        rates = sorted(rates, key=lambda x: x['Efective_date'], reverse=True)
        # print("RATES",rates)

        opportunities = list(db.opportunities.find({}, {"_id": 0, "Category": 1, "opportunity_code": 1,
                                                        "opportunity_end_date": 1, "opportunity_final_transfer_date": 1,
                                                        "opportunity_sold": 1}))
        # filter opporunities to include only Heron Fields and Heron View
        opportunities = list(
            filter(lambda x: x['Category'] == "Heron Fields" or x['Category'] == "Heron View", opportunities))
        print("OPPORTUNITIES", opportunities[0])
        investors = list(db.investors.find({}))
        investments = []
        released = []
        for investor in investors:
            trust = list(filter(lambda x: x['Category'] == "Heron Fields" or x['Category'] == "Heron View",
                                investor['trust']))

            investment = list(filter(lambda x: x['Category'] == "Heron Fields" or x['Category'] == "Heron View",
                                     investor['investments']))
            # print()
            if len(trust) > 0:
                # print("TRUST", trust[0])
                for item in trust:
                    # convert item['deposit_date'] to datetime object
                    if item['deposit_date'] != "" and item['deposit_date'] != None:
                        deposit_date = datetime.strptime(item['deposit_date'], "%Y/%m/%d")
                    else:
                        deposit_date = ""
                    if item['release_date'] != "" and item['release_date'] != None:
                        release_date = datetime.strptime(item['release_date'], "%Y/%m/%d")
                    else:
                        release_date = ""
                    item['end_date'] = item.get('end_date', "")
                    if item['end_date'] != "" and item['end_date'] != None:
                        end_date = datetime.strptime(item['end_date'], "%Y/%m/%d")
                    else:
                        end_date = ""
                    insert = {
                        'investor_acc_number': investor['investor_acc_number'],
                        'investor_name': investor['investor_name'],
                        'investor_surname': investor['investor_surname'],
                        'investment_amount': float(item['investment_amount']),
                        # 'deposit_date': deposit_date,
                        'release_date': release_date,
                        'end_date': end_date,
                        'opportunity_code': item['opportunity_code'],
                        'Category': item['Category'],
                        'rate': item['rate'],
                        'investment_number': item['investment_number'],
                    }
                    investments.append(insert)

            # print()
            if len(investment) > 0:
                for item in investment:
                    # convert item['deposit_date'] to datetime object
                    # if item['deposit_date'] != "" and item['deposit_date'] != None:
                    #     deposit_date = datetime.strptime(item['deposit_date'], "%Y/%m/%d")
                    # else:
                    #     deposit_date = ""
                    if item['release_date'] != "" and item['release_date'] != None:
                        release_date = datetime.strptime(item['release_date'], "%Y/%m/%d")
                    else:
                        release_date = ""
                    if item['end_date'] != "" and item['end_date'] != None:
                        end_date = datetime.strptime(item['end_date'], "%Y/%m/%d")
                    else:
                        end_date = ""
                    insert = {
                        'investor_acc_number': investor['investor_acc_number'],
                        'investor_name': investor['investor_name'],
                        'investor_surname': investor['investor_surname'],
                        'investment_amount': float(item['investment_amount']),
                        # 'deposit_date': deposit_date,
                        'release_date': release_date,
                        'end_date': end_date,
                        'opportunity_code': item['opportunity_code'],
                        'Category': item['Category'],
                        'rate': float(item['investment_interest_rate']),
                        'investment_number': item.get('investment_number', 0),
                    }
                    released.append(insert)

        print("INVESTMENTS", investments[0])
        print()
        print("RELEASED", released[0])

        for item in investments:
            filtered_released = list(filter(
                lambda x: x['investor_acc_number'] == item['investor_acc_number'] and x['opportunity_code'] == item[
                    'opportunity_code'] and x['investment_number'] == item['investment_number'], released))
            if len(filtered_released) > 0:
                item['end_date'] = filtered_released[0]['end_date']
                item['rate'] = filtered_released[0]['rate']

        print()
        print("INVESTMENTS", investments[0])
        print()
        print("RELEASED", released[0])
        print("RELEASED", len(released))
        print("INVESTMENTS", len(investments))

        # for item in investments:

        # print("INVESTMENT", investment[0])
    except Exception as e:
        print("ERROR", e)


# recalc_interest()


@cashflow.post("/construction_cashflow")
async def construction_cashflow(data: Request):
    request = await data.json()
    data = request['data']

    for item in data:
        # loop through each individual item dictionary
        # convert all values to float
        for key, value in item.items():
            if key != "Whitebox-Able" and key != "Complete Build" and key != "Blocks" and key != "Option" and key != "Development":
                # item[key] = float(value)
                value = value.replace("R\xa0", "")
                value = value.replace(",", "")
                value = value.replace(" ", "")
                value = value.replace("R", "")
                value = float(value)
                item[key] = value
    #
    #
    #         # print(f"{key}: {value}")
    print("DATA", data[0])

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


@cashflow.get("/construction_cashflowFetch")
async def get_construction_cashflow():
    try:
        data = list(db.cashflow_construction.find({}))

        print("DATAXX", data[0])
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


def get_opportunities_1(data):
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

    data = request['data']

    try:
        start = time.time()

        get_date = list(db.cashflow_xero_tb.find({}, {"_id": 0, "ReportDate": 1}))
        dates = []
        for d in get_date:
            dates.append(d['ReportDate'])

        dates = list(sorted(set(dates)))

        date = dates[-1]

        sales = get_sales_data(date)
        # sort sales by transferred and opportunity_code
        sales = sorted(sales, key=lambda x: (x['transferred'], x['opportunity_code']))

        opportunities = list(db.opportunities.find({}, {"_id": 0}))

        sales_data = list(db.cashflow_sales.find({}))

        # filter sales_data where transferred is False
        sales_to_update = []
        sales_data_filtered = list(filter(lambda x: x['transferred'] == False, sales_data))
        # print(len(sales_data_filtered))
        for sale in sales_data_filtered:
            sale['_id'] = str(sale['_id'])
            opportunities_filtered = list(
                filter(lambda x: x['opportunity_code'] == sale['opportunity_code'], opportunities))
            if len(opportunities_filtered) > 0:
                if opportunities_filtered[0]['opportunity_final_transfer_date'] != "" and sale['transferred'] == False:
                    # print("sale", sale)
                    # print()
                    # print("opportunities_filtered", opportunities_filtered[0])
                    # print()
                    sale['forecast_transfer_date'] = opportunities_filtered[0]['opportunity_final_transfer_date']
                    sale['sold'] = True
                    sale['transferred'] = True
                    sales_to_update.append(sale)
                elif opportunities_filtered[0]['opportunity_sold'] == True and sale['sold'] == False:
                    sale['sold'] = True
                    sales_to_update.append(sale)
                    # print("sale sold", sale)
                    # print()
                    # print("opportunities_filtered sold", opportunities_filtered[0])
                    # print()

        # print("sales_to_update", len(sales_to_update))
        # print()
        # print("sales_to_update", sales_to_update)

        # sort sales to update by opportunity_code
        sales_to_update = sorted(sales_to_update, key=lambda x: x['opportunity_code'])

        if len(sales_to_update) > 0:
            for sale in sales_to_update:
                try:
                    id = sale['_id']
                    del sale['_id']
                    db.cashflow_sales.update_one({"_id": ObjectId(id)}, {"$set": sale})
                except Exception as e:
                    print("Error updating sales", e)

            # print("sale", sale['opportunity_code'])
            sales_data = list(db.cashflow_sales.find({}))

        # print("sales_data_filtered", sales_data_filtered[0])

        construction = get_construction_costs()
        for item in construction:
            item['Blocks'] = item['Blocks'].replace("Block ", "")
            # remove leading and trailing whitespace from item['block']
            item['Blocks'] = item['Blocks'].strip()

        for sale in sales:
            filtered_sales_data = list(filter(lambda x: x['opportunity_code'] == sale['opportunity_code'], sales_data))
            filtered_construction_data = list(filter(
                lambda x: x['Blocks'] == sale['block'] and x['Whitebox-Able'] == True and sale[
                    'Category'] == 'Heron View', construction))
            # print("filtered_construction_data", len(filtered_construction_data))
            if len(filtered_construction_data) > 0:
                if filtered_construction_data[0]['Complete Build'] == 1:
                    sale['complete_build'] = True
                else:
                    sale['complete_build'] = False
            else:
                sale['complete_build'] = True

            # print("filtered_sales_data", len(filtered_sales_data))
            sale['_id'] = str(filtered_sales_data[0]['_id'])
            sale['profit_loss'] = float(sale['profit_loss'])
            sale['sale_price'] = float(sale['sale_price'])
            sale['profit_loss_nice'] = "R{:,.2f}".format(sale['profit_loss'])
            sale['sale_price_nice'] = "R{:,.2f}".format(sale['sale_price'])
            del sale['vat_date']
            # convert sale['original_planned_transfer_date'] to a string in the format yyyy-mm-dd
            sale['original_planned_transfer_date'] = sale['original_planned_transfer_date'].strftime('%Y-%m-%d')
            # convert sale['forecast_transfer_date'] to a string in the format yyyy-mm-dd
            sale['forecast_transfer_date'] = sale['forecast_transfer_date'].strftime('%Y-%m-%d')

        end = time.time()

        return {"success": True, "data": sales, "time": end - start}

    except Exception as e:
        print("Error getting sales cashflow initial", e)
        return {"success": False, "error": str(e)}


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
        investor["trust"] = list(
            filter(lambda x: x["Category"] != "Endulini" and x["Category"] != "Southwark" and x[
                "Category"] != "NGAH", investor["trust"]))
        # filter investor["Investments"] where Category = "Heron Fields" or Category = "Heron View"

        investor["investments"] = list(
            filter(lambda x: x["Category"] != "Endulini" and x["Category"] != "Southwark" and x[
                "Category"] != "NGAH", investor["investments"]))
        # filter investor["pledges"] where Category = "Heron Fields" or Category = "Heron View"
        investor["pledges"] = list(
            filter(lambda x: x["Category"] != "Endulini" and x["Category"] != "Southwark" and x[
                "Category"] != "NGAH", investor["pledges"]))

        # filter out from investor["trust"] where release_date is not empty
    # filter out of investors where trust is empty
    investors = list(filter(lambda x: len(x["trust"]) > 0, investors))



    opportunities = list(db.opportunities.find({}, {"_id": 0}))

    opportunities = list(
        filter(
            lambda x: x["Category"] != "Endulini" and x["Category"] != "Southwark" and x[
                "Category"] != "NGAH",
            opportunities))

    final_investors = []

    for investor in investors:
        # filter out from investor["trust"] where release_date is not empty

        trust_filtered = list(filter(lambda x: x.get("release_date", "") == "", investor["trust"]))
        if len(trust_filtered) > 0:
            for trust in trust_filtered:
                filtered_opportunity = list(
                    filter(lambda x: x["opportunity_code"] == trust["opportunity_code"], opportunities))

                if len(filtered_opportunity) > 0:

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

                if len(filtered_opportunity) > 0:

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
        # print(opportunities)
        if len(investor["pledges"]) > 0:
            for pledge in investor["pledges"]:
                if pledge["Category"] != "Unallocated":

                    filtered_opportunity = list(
                        filter(lambda x: x["opportunity_code"] == pledge["opportunity_code"], opportunities))

                    if len(filtered_opportunity) > 0:
                        # print(pledge)
                        # print("filtered_opportunity", filtered_opportunity)
                        # print("Opportunity Code:",pledge['opportunity_code'])
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
    # sales_parameters = list(
    #     filter(lambda x: x["Development"] == "Heron Fields" or x["Development"] == "Heron View", sales_parameters))
    for param in sales_parameters:
        param["rate"] = float(param["rate"])
        param["Effective_date"] = param["Effective_date"].replace("-", "/")
        param["Effective_date"] = datetime.strptime(param["Effective_date"], "%Y/%m/%d")

    sales = list(db.sales_processed.find({}, {"_id": 0}))
    # sales = list(filter(lambda x: x["development"] == "Heron Fields" or x["development"] == "Heron View", sales))
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

    final_investors = [investor for investor in final_investors if
                       not (investor['investor_acc_number'] == "ZGEC01")]

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

def get_construction_costsA():
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


def get_sales_data(report_date):
    # print("report_date", report_date)
    report_date = report_date.replace("-", "/")
    report_date = datetime.strptime(report_date, "%Y/%m/%d")
    # print("report_date", report_date)
    try:
        construction_data = list(db.cashflow_construction.find({}, {"_id": 0}))
        for item in construction_data:
            item['block'] = item['Blocks'].replace("Block ", "")
            # remove leading and trailing whitespace from item['block']
            item['block'] = item['block'].strip()
            del item['Blocks']
        # filter construction_data by "Whitebox-Able" is True

        construction_data = list(filter(lambda x: x['Whitebox-Able'] == True, construction_data))

        opportunities = list(db.opportunities.find({}, {"_id": 0}))
        # print("opportunities", opportunities[0])

        opportunities = list(
            filter(lambda x: x["Category"] != "Endulini" and x["Category"] != "Southwark" and x[
                "Category"] != "NGAH",
                   opportunities))

        # Get actual sales data from sales_processed where development is "Heron Fields" or "Heron View"
        sales_data_actual = list(
            db.sales_processed.find({}, {"_id": 0, "development": 1, "opportunity_code": 1, "opportunity_sales_date": 1,
                                         "opportunity_actual_reg_date": 1, "opportunity_potential_reg_date": 1,
                                         "opportunity_contract_price": 1}))

        # filter out where development is not Heron Fields or Heron View
        # print("sales_data_actual A", len(sales_data_actual))
        # for sale in sales_data_actual:
        #     print(sale)
        # sales_data_actual = list(
        #     filter(lambda x: x["development"] == "Heron View" or x["development"] == "Heron Fields" or x[
        #         "development"] == "Goodwood",
        #            sales_data_actual))

        sales_data_actual = list(
            filter(lambda x: x["development"] != "Endulini" and x["development"] != "Southwark" and x[
                "development"] != "NGAH",
                   sales_data_actual))

        # print("sales_data_actual B", sales_data_actual[0])

        # print(len(sales_data_actual))
        # print(sales_data_actual[160])
        # print()
        # print(sales_data_actual[161])

        for index, sale in enumerate(sales_data_actual):

            # convert opportunity_sales_date to datetime
            # if sale['opportunity_sales_date'] == None:
            # print(sale['opportunity_code'], sale['opportunity_sales_date'], sale['opportunity_actual_reg_date'])
            if 'opportunity_sales_date' in sale:
                sale['opportunity_sales_date'] = sale['opportunity_sales_date'].replace("-", "/")
                sale['opportunity_sales_date'] = datetime.strptime(sale['opportunity_sales_date'], '%Y/%m/%d')
            else:
                sale['opportunity_sales_date'] = sale['opportunity_sale_date'].replace("-", "/")
                sale['opportunity_sales_date'] = datetime.strptime(sale['opportunity_sales_date'], '%Y/%m/%d')

            # convert opportunity_actual_reg_date to datetime
            # print("Got this far!!!!", index)
            try:

                if sale['opportunity_actual_reg_date'] != "" and sale['opportunity_actual_reg_date'] is not None:
                    # print("ACTUAL",sale['opportunity_actual_reg_date'], sale['opportunity_code'])
                    sale['opportunity_actual_reg_date'] = sale['opportunity_actual_reg_date'].replace("-", "/")
                    sale['opportunity_actual_reg_date'] = datetime.strptime(sale['opportunity_actual_reg_date'],
                                                                            '%Y/%m/%d')
                elif sale['opportunity_potential_reg_date'] != "" and sale[
                    'opportunity_potential_reg_date'] is not None:
                    # print("POTENTIAL",sale['opportunity_actual_reg_date'], sale['opportunity_code'])
                    sale['opportunity_actual_reg_date'] = sale['opportunity_potential_reg_date'].replace("-", "/")
                    sale['opportunity_actual_reg_date'] = datetime.strptime(sale['opportunity_potential_reg_date'],
                                                                            '%Y/%m/%d')
                else:

                    opportunities_filtered = list(
                        filter(lambda x: x['opportunity_code'] == sale['opportunity_code'], opportunities))
                    # opportunity_end_date
                    # print()
                    # print("opportunities_filtered", opportunities_filtered[0])

                    sale['opportunity_actual_reg_date'] = opportunities_filtered[0]['opportunity_end_date'].replace("-",
                                                                                                                    "/")
                    # print("END DATE",sale['opportunity_actual_reg_date'], sale['opportunity_code'])
                    # print()
                    sale['opportunity_actual_reg_date'] = datetime.strptime(sale['opportunity_actual_reg_date'],
                                                                            '%Y/%m/%d')
                    # print("Processed END DATE",sale['opportunity_actual_reg_date'], sale['opportunity_code'])
                    # print()
                    # print(sale['opportunity_actual_reg_date'], sale['opportunity_code'])
                    # print("Hello",sale['opportunity_actual_reg_date'], sale['opportunity_code'])

                    # sale['opportunity_actual_reg_date'] = "None"

                # if sale["development"] == "Goodwood":
                #     print("Goodwood", sale)
            except Exception as e:
                print("Error converting opportunity_actual_reg_date to datetime", e, sale['opportunity_code'],
                      sale['opportunity_potential_reg_date'], sale['opportunity_code'])
                # print(sale['opportunity_actual_reg_date'], sale['opportunity_code'])

                continue

        # print("sales_data", sales_data_actual[0])

        # print("construction_data", construction_data[0])


        sales_data = list(db.cashflow_sales.find({}, {"_id": 0}))

        # filter out of sales data where Category is Endulini, Southwark or NGAH
        sales_data = list(
            filter(lambda x: x["Category"] != "Endulini" and x["Category"] != "Southwark" and x["Category"] != "NGAH",
                   sales_data))

        # print("opportunities", opportunities[0])

        for opp in opportunities:
            filtered_sales_data = list(filter(lambda x: x['opportunity_code'] == opp['opportunity_code'], sales_data))
            if len(filtered_sales_data) == 0:
                # print("opp", opp['opportunity_code'])
                sale_price = float(opp['opportunity_sale_price'])
                if opp['opportunity_final_transfer_date'] != "":
                    transferred = True
                else:
                    transferred = False
                insert = {
                    "Category": opp['Category'],
                    "block": opp['opportunity_code'][-4],
                    "opportunity_code": opp['opportunity_code'],
                    "sold": opp['opportunity_sold'],
                    "transferred": transferred,
                    "complete_build": True,
                    "original_planned_transfer_date": opp['opportunity_end_date'],
                    "forecast_transfer_date": opp['opportunity_end_date'],
                    "sale_price": opp['opportunity_sale_price'],
                    "VAT": sale_price / 1.15 * 0.15,
                    "nett": sale_price / 1.15,
                    "opportunity_transfer_fees": 0,
                    "opportunity_trust_release_fee": 1789,
                    "opportunity_unforseen": sale_price * 0.005,
                    "opportunity_commission": sale_price * 0.05,
                    "opportunity_bond_registration": 3500,
                    "transfer_income": (sale_price / 1.15) - 1789 - (sale_price * 0.005) - (sale_price * 0.05) - 3500,
                    "due_to_investors": 0,
                    "profit_loss": (sale_price / 1.15) - 1789 - (sale_price * 0.005) - (sale_price * 0.05) - 3500,
                    "refinanced": False,
                }
                sales_data.append(insert)
                # print("insert", insert)
                # print()

        # for sale in sales_data_actual:
        #     filtered_sales_data = list(filter(lambda x: x['opportunity_code'] == sale['opportunity_code'], sales_data))
        #     if len(filtered_sales_data) == 0:
        #
        #         print("sale", sale)
        # print()
        # sales_data_actual.remove(sale

        for index, sale in enumerate(sales_data):
            #

            # if index == 18:
            # print("sale", sale['opportunity_code'], index)

            construction_data_filtered = list(filter(lambda x: x['block'] == sale['block'], construction_data))


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

            # print("sales_data 18", sales_data[18])
            # if index == 18:
            #     print("Got here 1")

            # convert original_planned_transfer_date to datetime
            sale['original_planned_transfer_date'] = sale['original_planned_transfer_date'].replace("-", "/")
            sale['original_planned_transfer_date'] = datetime.strptime(sale['original_planned_transfer_date'],
                                                                       '%Y/%m/%d')

            # print("Got here 2")
            # if forecast_transfer_date is not empty then convert to datetime
            sales_data_actual_filtered = list(
                filter(lambda x: x['opportunity_code'] == sale['opportunity_code'], sales_data_actual))

            # print("Got here 3")
            # print("sales_data_actual_filtered", sales_data_actual_filtered[0])
            # print("sales_data_actual_filtered", len(sales_data_actual_filtered))

            if len(sales_data_actual_filtered) > 0:
                sale['forecast_transfer_date'] = sales_data_actual_filtered[0]['opportunity_actual_reg_date']
                sale['forecast_transfer_date'] = sale['forecast_transfer_date'].strftime('%Y-%m-%d')
                sale['sale_price'] = sales_data_actual_filtered[0]['opportunity_contract_price']
            else:
                opportunities_filtered = list(
                    filter(lambda x: x['opportunity_code'] == sale['opportunity_code'], opportunities))
                if len(opportunities_filtered) > 0:
                    sale['forecast_transfer_date'] = opportunities_filtered[0]['opportunity_end_date']

                else:
                    sale['forecast_transfer_date'] = "ISSUE HERE"

            # print("Got here 4")
            # print("XXXXX",sale['forecast_transfer_date'])
            # convert forecast_transfer_date to string in the format yyyy-mm-dd

            if sale['forecast_transfer_date'] == "" or sale['forecast_transfer_date'] == None:
                # if index == 18:
                # print("17",sales_data[17])
                # print()
                # print("18",sale)

                # convert sale['original_planned_transfer_date'] to string in the format yyyy/mm/dd
                vat_date = sale['original_planned_transfer_date'].strftime('%Y/%m/%d')
                sale['vat_date'] = calculate_vat_due(vat_date)
            else:
                sale['vat_date'] = calculate_vat_due(sale['forecast_transfer_date'])
            # print("Got here 4A")
            if sale['forecast_transfer_date'] != "" and sale['forecast_transfer_date'] != None:
                sale['forecast_transfer_date'] = sale['forecast_transfer_date'].replace("-", "/")
                sale['forecast_transfer_date'] = datetime.strptime(sale['forecast_transfer_date'], '%Y/%m/%d')

            # convert vat_date to datetime
            sale['vat_date'] = sale['vat_date'].replace("/", "-")
            sale['vat_date'] = datetime.strptime(sale['vat_date'], '%Y-%m-%d')
            # print("Got here 5")
            filtered_sales_data_actual = list(
                filter(lambda x: x['opportunity_code'] == sale['opportunity_code'], sales_data_actual))
            # print("Got here 4","len", len(filtered_sales_data_actual), filtered_sales_data_actual[0]['opportunity_sales_date'], report_date)
            if len(filtered_sales_data_actual) > 0:
                # print("Got here 6")

                if filtered_sales_data_actual[0]['opportunity_sales_date'] > report_date:
                    sale['sold'] = False
                # print("Got here 6")

                # print(sale)
                # print()

                if filtered_sales_data_actual[0]['opportunity_actual_reg_date'] != "None":
                    if filtered_sales_data_actual[0]['opportunity_actual_reg_date'] != None:
                        if filtered_sales_data_actual[0]['opportunity_actual_reg_date'] > report_date:
                            sale['transferred'] = False
                # print("Got here 7", sale['opportunity_code'])
            # else:
            # print("Got here 8")

            # del sale['profit_loss_nice']
            # del sale['sale_price_nice']
        # print("sales_data", sales_data[0])


        for item in sales_data:

            filtered_opportunities = list(
                filter(lambda x: x['opportunity_code'] == item['opportunity_code'], opportunities))
            # print("item", item)
            # opportunity_sold
            if not filtered_opportunities[0]['opportunity_sold']:
                item['sold'] = False
                item['transferred'] = False

        # filter out of sales_data where Category is Endulini
        sales_data = list(filter(lambda x: x['Category'] != "Endulini", sales_data))
        # filter out of sales_data where Category is Heron Fields
        # sales_data = list(filter(lambda x: x['Category'] != "Heron Fields", sales_data))

        for item in sales_data:
            if item['Category'] == "Goodwood":
                item['block'] = "R"
                # print(item)

        # sort sales_data by Category, block, opportunity_code
        sales_data = sorted(sales_data, key=lambda x: (x['Category'], x['block'], x['opportunity_code']))

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
        # filter xero_tbs returning only hen AccountCode exists
        # filter out of xero_tbs where ReportTitle is "Purple Blok Projects (Pty) Ltd"

        # xero_tbs = list(filter(lambda x: x['ReportTitle'] != "Purple Blok Projects (Pty) Ltd", xero_tbs))
        xero_tbs = list(filter(lambda x: 'AccountCode' in x, xero_tbs))
        # for item in xero_tbs:
        #     if not 'AccountCode' in item:
        #         del item
        #         # print(item)

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
        # xero_tbs2 = list(filter(lambda x: x['Category'] == 'Assets', xero_tbs2))
        # filter out of xero_tbs2 where AccountCode does not begin with 84
        # xero_tbs2 = list(filter(lambda x: str(x['AccountCode']).startswith("84"), xero_tbs2))
        # xero_tbs2 = list(filter(lambda x: not x['AccountCode'].startswith("8480"), xero_tbs2))
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
            filter(lambda x: x["Category"] != "Endulini" and x["Category"] != "Southwark" and x[
                "Category"] != "NGAH", opportunities))
        # sort by Category then by opportunity_code
        opportunities = sorted(opportunities, key=lambda x: (x['Category'], x['opportunity_code']))

        investors = list(db.investors.find({}, {"_id": 0}))
        for investor in investors:
            # filter investor["Trust"] where Category = "Heron Fields" or Category = "Heron View"
            investor["trust"] = list(
                filter(lambda x: x["Category"] != "Endulini" and x["Category"] != "Southwark" and x[
                    "Category"] != "NGAH", investor["trust"]))

            investor["pledges"] = list(
                filter(lambda x: x["Category"] != "Endulini" and x["Category"] != "Southwark" and x[
                    "Category"] != "NGAH", investor["pledges"]))

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

        final_investors = [investor for investor in final_investors if
                           not (investor['investor_acc_number'] == "ZCAM01" and investor[
                               'opportunity_code'] == "HFA101" and investor['investment_amount'] == 400000.0)]

        final_investors = [investor for investor in final_investors if
                           not (investor['investor_acc_number'] == "ZJHO01" and investor[
                               'opportunity_code'] == "HFA304" and investor['investment_number'] == 1)]

        final_investors = [investor for investor in final_investors if
                           not (investor['investor_acc_number'] == "ZPJB01" and investor[
                               'opportunity_code'] == "HFA205" and investor['investment_number'] == 1)]

        final_opportunities = []
        print("opportunities", len(opportunities))
        # filter opportunities so that Category is Heron View
        # test1 = list(filter(lambda x: x['Category'] == "Heron View", opportunities))
        # test2 = list(filter(lambda x: x['Category'] == "Heron Fields", opportunities))
        # test3 = list(filter(lambda x: x['Category'] == "Goodwood", opportunities))
        # test4 = list(filter(lambda x: x['Category'] == "TEST", opportunities))
        # print("test1", len(test1))
        # print("test2", len(test2))
        # print("test3", len(test3))
        # print("test4", len(test4))

        # for opp in opportunities:
        #     print("opp", opp)
        #     print()

        for index, opportunity in enumerate(opportunities):
            # print("opportunity", opportunity, index)
            # convert opportunity_end_date to datetime
            opportunity['opportunity_end_date'] = opportunity['opportunity_end_date'].replace("-", "/")
            opportunity['opportunity_end_date'] = datetime.strptime(opportunity['opportunity_end_date'], '%Y/%m/%d')
            # convert opportunity_final_transfer_date to datetime
            # print("opportunity", opportunity, index)
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

        return final_opportunities
    except Exception as e:
        print("Error getting opportunities", e)
        return {"success": False, "error": str(e)}


@cashflow.post("/generate_investors_new_cashflow_nsst_report")
async def generate_investors_new_cashflow_nsst_report(data: Request, background_tasks: BackgroundTasks):
    request = await data.json()
    date = request['date']
    # print(date)

    try:
        start = time.time()
        if os.path.exists("cashflow_p&l_files/cashflow_projection.xlsx"):
            os.remove("cashflow_p&l_files/cashflow_projection.xlsx")
        invest = investors_new_cashflow_nsst_report()

        construction = get_construction_costsA()
        sales = get_sales_data(date)
        for sale in sales:
            if sale['sale_price'] == "" or sale['sale_price'] == None:
                sale['sale_price'] = 0
            sale['sale_price'] = float(sale['sale_price'])
        print("sales", sales[0])

        operational_costs = get_operational_costs()
        xero = get_xero_tbs()
        opportunities = get_opportunities()
        # DO FOR INVESTOR EXIT REPORT THEN CREATE A FUNCTION
        # print("invest", invest[0])
        investor_exit = []
        current_report_date = date.replace("-", "/")
        current_report_date = datetime.strptime(current_report_date, '%Y/%m/%d')
        for investor in invest:
            insert = {
                "total_units": investor['investor_acc_number'],
                "investment_number": investor['investment_number'],
                "unit_number": investor['opportunity_code'],
                "block": investor['Block'],
                "Investor": f"{investor['investor_surname']} {investor['investor_name']}",
                "capital_amount": 0,
                "fund_release_date": investor['release_date'],
                "unit_sold_status": investor['sold'],
                "unit_transferred_status": investor['tansferred'],
                "estimated_transfer_date": investor['end_date'],
                "actual_transfer_date": "",
                "Exit Deadline (730 Days)": "",
                "current_report_date": current_report_date,
                "days_to_contract_exit": 0,
                "days_to_estimated_exit": 0,
                "Investor Contract expiry exit": 0,
                "Capital & Interest to be Exited": 0,
                "Investor Exit Value On Sales": 0,
                "Exited by Developer": 0,
                "Date of Exit": "",
                "Early Release": investor['early_release'],
                "Investor pay Back On transfer": 0,
                "Developer & Unbonded": 0,
                "still_pledged": investor['still_pledged'],
            }
            investor_exit.append(insert)

        for investor in investor_exit:
            if investor['unit_transferred_status'] and investor['estimated_transfer_date'] > current_report_date:
                investor['unit_transferred_status'] = False

        momentum = list(db.cashflow_momentum.find({}, {"_id": 0}))
        for item in momentum:
            item['MONTH'] = item['MONTH'].replace("-", "/")
            item['MONTH'] = datetime.strptime(item['MONTH'], '%Y/%m/%d')
            # item['_id'] = str(item['_id'])

        # print()
        # print("investor_exit", investor_exit[0])
        # print()

        # result = cashflow_projections(invest, construction, sales, operational_costs, xero, opportunities,
        #                               investor_exit, momentum, date)

        # for item in construction:
        #     print(item)
        #     print()
        # insert = {
        #     "Whitebox-Able": True,
        #     "Blocks": "Block R",
        #     'Development': 'Goodwood',
        #     "Complete Build": 0,
        #     'Option': 0,
        #     'Remaining As Per Options': 0.0,
        #     'To Complete After Option Achieved': 0.0,
        #     ' Total Cost To Complete ': 2019943.35,
        #     '2024/03/31 Actual': 0.0,
        #     '30-Apr-24': 0.0,
        #     '30-May-24': 0.0,
        #     '29-Jun-24': 0.0,
        #     '29-Jul-24': 211008.5,
        #     '28-Aug-24': 302718.75,
        #     '27-Sep-24': 211710.25,
        #     '31-Oct-24': 5437.5,
        #     '30-Nov-24': 2718.75,
        #     '31-Dec-24': 2718.75,
        #     '31-Jan-25': 2718.75,
        #     '28-Feb-25': 2718.75,
        #
        # }
        # construction.append(insert)
        #
        # # head = [0, 1, 2]
        # # k = 3
        # # print("rotat= ", k % len(head))
        # # k = k % len(head)
        # # if k > 0:
        # #     for i in range(1, k + 1):
        # #         # move the last item in the list to the first position
        # #         head.insert(0, head.pop())
        # # print("Head",head)
        #
        # # print("operatonal_costs", operational_costs[0])
        # insert = {
        #     ' Company ': 'Goodwood',
        #     ' Account ': 'Professional & Other Fees',
        #     ' Month1 ': 0,
        #     ' Month2 ': 0,
        #     ' Month3 ': 0,
        #     ' Month4 ': 0,
        #     ' Month5 ': 295387.0425,
        #     ' Month6 ': 361192.3975,
        #     ' Month7 ': 248820,
        #     ' Month8 ': 238645,
        #     ' Month9 ': 228470,
        #     ' Month10 ': 215455.614,
        #     ' Month11 ': 97526,
        #     ' Operating Expenses ': 240785.1506
        # }
        # operational_costs.append(insert)

        for item in invest:
            if item["Category"] == "Goodwood":
                item['Block'] = "R"

        for item in investor_exit:
            # if the unit_number begins with GW then make block = R
            if item['unit_number'].startswith("GW"):
                item['block'] = "R"
                # print()
            # if index == 0:
            #     print(item)
            # item['block'] = "R"
            # print()

        background_tasks.add_task(cashflow_projections, invest, construction, sales, operational_costs, xero,
                                  opportunities,
                                  investor_exit, momentum, date)

        result = "cashflow_p&l_files/cashflow_projection.xlsx"

        # result = "Awesome"
        end = time.time()
        print("Time taken", end - start)
        return {"success": True, "Result": result, "time": end - start}
    except Exception as e:
        print("Error generating investors new cashflow nsst report", e)
        return {"success": False, "error": str(e)}


@cashflow.get("/check_if_file_exists")
async def check_if_file_exists():
    file = "cashflow_p&l_files/cashflow_projection.xlsx"
    # if os.path.exists("cashflow_p&l_files/cashflow_projection.xlsx"):
    #     os.remove("cashflow_p&l_files/cashflow_projection.xlsx")
    try:
        if os.path.exists(file):
            return {"success": True}
        else:
            return {"success": False}
    except Exception as e:
        print("Error checking if file exists", e)
        return {"success": False, "error": str(e)}


# def import_momentum():
#     try:
#         # get momentum.json
#         data = [
#  {
#    "MONTH": "2021-08-31",
#    "INTEREST": 0.69,
#    "ADVICE FEES (Exc Vat)": 2.88,
#    "ONGOING ADVICE FEES (Inc Vat)": 0
#  },
#  {
#    "MONTH": "2021-09-30",
#    "INTEREST": 6153.47,
#    "ADVICE FEES (Exc Vat)": 20700,
#    "ONGOING ADVICE FEES (Inc Vat)": 142.83
#  },
#  {
#    "MONTH": "2021-10-31",
#    "INTEREST": 27493.93,
#    "ADVICE FEES (Exc Vat)": 39100,
#    "ONGOING ADVICE FEES (Inc Vat)": 1963.75
#  },
#  {
#    "MONTH": "2021-11-30",
#    "INTEREST": 39490.47,
#    "ADVICE FEES (Exc Vat)": 6612.5,
#    "ONGOING ADVICE FEES (Inc Vat)": 4603.54
#  },
#  {
#    "MONTH": "2021-12-31",
#    "INTEREST": 34809.07,
#    "ADVICE FEES (Exc Vat)": 0,
#    "ONGOING ADVICE FEES (Inc Vat)": 4987.02
#  },
#  {
#    "MONTH": "2022-01-31",
#    "INTEREST": 25737.44,
#    "ADVICE FEES (Exc Vat)": 0,
#    "ONGOING ADVICE FEES (Inc Vat)": 3416.75
#  },
#  {
#    "MONTH": "2022-02-28",
#    "INTEREST": 17923.77,
#    "ADVICE FEES (Exc Vat)": 2012.5,
#    "ONGOING ADVICE FEES (Inc Vat)": 2547.7
#  },
#  {
#    "MONTH": "2022-03-31",
#    "INTEREST": 7694.95,
#    "ADVICE FEES (Exc Vat)": 0,
#    "ONGOING ADVICE FEES (Inc Vat)": 1329.87
#  },
#  {
#    "MONTH": "2022-04-30",
#    "INTEREST": 33565.07,
#    "ADVICE FEES (Exc Vat)": 22274.94,
#    "ONGOING ADVICE FEES (Inc Vat)": 561.9
#  },
#  {
#    "MONTH": "2022-05-31",
#    "INTEREST": 19911.29,
#    "ADVICE FEES (Exc Vat)": 56298.31,
#    "ONGOING ADVICE FEES (Inc Vat)": 551.12
#  },
#  {
#    "MONTH": "2022-06-30",
#    "INTEREST": 57111.6,
#    "ADVICE FEES (Exc Vat)": 80343.79,
#    "ONGOING ADVICE FEES (Inc Vat)": 3852.47
#  },
#  {
#    "MONTH": "2022-07-31",
#    "INTEREST": 49997.61,
#    "ADVICE FEES (Exc Vat)": 55413.57,
#    "ONGOING ADVICE FEES (Inc Vat)": 4675.82
#  },
#  {
#    "MONTH": "2022-08-31",
#    "INTEREST": 49010.03,
#    "ADVICE FEES (Exc Vat)": 30982.39,
#    "ONGOING ADVICE FEES (Inc Vat)": 4493.24
#  },
#  {
#    "MONTH": "2022-09-30",
#    "INTEREST": 30938.76,
#    "ADVICE FEES (Exc Vat)": 10350,
#    "ONGOING ADVICE FEES (Inc Vat)": 3425.4
#  },
#  {
#    "MONTH": "2022-10-31",
#    "INTEREST": 31855.81,
#    "ADVICE FEES (Exc Vat)": 0,
#    "ONGOING ADVICE FEES (Inc Vat)": 1362.36
#  },
#  {
#    "MONTH": "2022-11-30",
#    "INTEREST": 57998.29,
#    "ADVICE FEES (Exc Vat)": 0,
#    "ONGOING ADVICE FEES (Inc Vat)": 3649
#  },
#  {
#    "MONTH": "2022-12-31",
#    "INTEREST": 49641.63,
#    "ADVICE FEES (Exc Vat)": 0,
#    "ONGOING ADVICE FEES (Inc Vat)": 2707.2
#  },
#  {
#    "MONTH": "2023-01-31",
#    "INTEREST": 71919.64,
#    "ADVICE FEES (Exc Vat)": 0,
#    "ONGOING ADVICE FEES (Inc Vat)": 1766.93
#  },
#  {
#    "MONTH": "2023-02-28",
#    "INTEREST": 116840.82,
#    "ADVICE FEES (Exc Vat)": 0,
#    "ONGOING ADVICE FEES (Inc Vat)": 6530.48
#  },
#  {
#    "MONTH": "2023-03-31",
#    "INTEREST": 224301.31,
#    "ADVICE FEES (Exc Vat)": 0,
#    "ONGOING ADVICE FEES (Inc Vat)": 5945.09
#  },
#  {
#    "MONTH": "2023-04-30",
#    "INTEREST": 106203.74,
#    "ADVICE FEES (Exc Vat)": 0,
#    "ONGOING ADVICE FEES (Inc Vat)": 9664.42
#  },
#  {
#    "MONTH": "2023-05-31",
#    "INTEREST": 222519.64,
#    "ADVICE FEES (Exc Vat)": 0,
#    "ONGOING ADVICE FEES (Inc Vat)": 6054.16
#  },
#  {
#    "MONTH": "2023-06-30",
#    "INTEREST": 393896.21,
#    "ADVICE FEES (Exc Vat)": 0,
#    "ONGOING ADVICE FEES (Inc Vat)": 12101.97
#  },
#  {
#    "MONTH": "2023-07-31",
#    "INTEREST": 259590.93,
#    "ADVICE FEES (Exc Vat)": 0,
#    "ONGOING ADVICE FEES (Inc Vat)": 17468.09
#  },
#  {
#    "MONTH": "2023-08-31",
#    "INTEREST": 241364.67,
#    "ADVICE FEES (Exc Vat)": 0,
#    "ONGOING ADVICE FEES (Inc Vat)": 14897.28
#  },
#  {
#    "MONTH": "2023-09-30",
#    "INTEREST": 291385.94,
#    "ADVICE FEES (Exc Vat)": 0,
#    "ONGOING ADVICE FEES (Inc Vat)": 17422.06
#  },
#  {
#    "MONTH": "2023-10-31",
#    "INTEREST": 342067.18,
#    "ADVICE FEES (Exc Vat)": 0,
#    "ONGOING ADVICE FEES (Inc Vat)": 21887.62
#  },
#  {
#    "MONTH": "2023-11-30",
#    "INTEREST": 289283.88,
#    "ADVICE FEES (Exc Vat)": 0,
#    "ONGOING ADVICE FEES (Inc Vat)": 20676.71
#  },
#  {
#    "MONTH": "2023-12-31",
#    "INTEREST": 158795.4,
#    "ADVICE FEES (Exc Vat)": 0,
#    "ONGOING ADVICE FEES (Inc Vat)": 13562.87
#  },
#  {
#    "MONTH": "2024-01-31",
#    "INTEREST": 125751.22,
#    "ADVICE FEES (Exc Vat)": 0,
#    "ONGOING ADVICE FEES (Inc Vat)": 8270.97
#  }
# ]
#         # with open('momentum.json') as f:
#         #     data = json.load(f)
#         # print(data)
#         try:
#             result = db.cashflow_momentum.insert_many(data)
#             print("RESULT", result)
#         except Exception as e:
#             print("Error", e)
#
#     except:
#         print("Error")

# import_momentum()

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


@cashflow.get("/get_cashflow_momentum")
async def get_cashflow_momentum():
    try:
        momentum = list(db.cashflow_momentum.find({}))
        for item in momentum:
            item['_id'] = str(item['_id'])
            # format all fields except _id and "Month" as currency with an R symbol
            item["INTEREST_R"] = f"R{item['INTEREST']:,.2f}"
            item["ADVICE FEES (Exc Vat)_R"] = f"R{item['ADVICE FEES (Exc Vat)']:,.2f}"
            item["ONGOING ADVICE FEES (Inc Vat)_R"] = f"R{item['ONGOING ADVICE FEES (Inc Vat)']:,.2f}"
            # item[key] = f"R{value:,.2f}"

        return {"success": True, "data": momentum}
    except Exception as e:
        print("Error getting cashflow momentum", e)
        return {"success": False, "error": str(e)}


@cashflow.post("/add_cashflow_momentum")
async def add_cashflow_momentum(data: Request):
    request = await data.json()
    print(request)
    data = request['data']
    data['INTEREST'] = float(data['INTEREST'])
    data['ADVICE FEES (Exc Vat)'] = float(data['ADVICE FEES (Exc Vat)'])
    data['ONGOING ADVICE FEES (Inc Vat)'] = float(data['ONGOING ADVICE FEES (Inc Vat)'])
    # return {"success": True, "data": request}
    try:
        result = db.cashflow_momentum.insert_one(data)
        return {"success": True, "data": str(result.inserted_id)}
    except Exception as e:
        print("Error adding cashflow momentum", e)
        return {"success": False, "error": str(e)}


@cashflow.post("/delete_cashflow_momentum")
async def delete_cashflow_momentum(data: Request):
    request = await data.json()
    print(request)
    _id = request['data']
    print(_id)
    # return {"success": True, "data": _id}
    try:
        result = db.cashflow_momentum.delete_one({"_id": ObjectId(_id)})
        return {"success": True, "data": str(result.deleted_count)}
    except Exception as e:
        print("Error deleting cashflow momentum", e)
        return {"success": False, "error": str(e)}


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


# import pandas as pd
# def calculate_interest_daily():
#     try:
#         investors = list(db.investors.find({}, {"_id": 0}))
#
#         # filter out of investors where investments is empty
#         final_investors = []
#         investors = list(filter(lambda x: len(x["investments"]) > 0, investors))
#         for investor in investors:
#             for investment in investor['investments']:
#                 if investment['end_date'] == "":
#                     insert = {
#                         "investment_amount": float(investment['investment_amount']),
#                         "opportunity_code": investment['opportunity_code'],
#                         "Category": investment['Category'],
#                         "investor_acc_number": investor['investor_acc_number'],
#                         "investment_number": investment.get("investment_number", 0),
#                         "investment_interest_rate": float(investment.get("investment_interest_rate", 0)) / 100,
#                         "daily_interest": float(investment['investment_amount']) * (float(investment.get("investment_interest_rate", 0)) / 100) / 365,
#                     }
#                     final_investors.append(insert)
#
#         print("investors", final_investors[0])
#         # save final_investors as a csv file
#         df = pd.DataFrame(final_investors)
#         df.to_csv("daily_interest.csv", index=False)
#
#
#
#     except Exception as e:
#         print("Error getting investors", e)

# calculate_interest_daily()

# def get_correct_investors():
#     investors = list(db.investors.find({}, {"_id": 0}))
#     # filter out of investors where investments is empty
#     final_investors = []
#     investors = list(filter(lambda x: len(x["investments"]) > 0, investors))
#     # filter out of investments where Category is not "Heron View"
#     for investor in investors:
#             investor["investments"] = list(
#                 filter(lambda x: x["Category"] == "Heron Fields" or x["Category"] == "Heron View", investor["investments"]))
#
#             for investment in investor['investments']:
#                 if investment['end_date'] == "":
#                     insert = {
#                         "investment_amount": float(investment['investment_amount']),
#                         "opportunity_code": investment['opportunity_code'],
#                         "Category": investment['Category'],
#                         "investment_number": investment.get("investment_number", 0),
#                         "investor_acc_number": investor['investor_acc_number'],
#                         "investment_number": investment.get("investment_number", 0),
#                         "release_date": datetime.strptime(investment.get("release_date", "").replace("/","-"), "%Y-%m-%d"),
#                     }
#                     final_investors.append(insert)
#
#
#     print(len(final_investors))
#     print(final_investors[0])
#     # save final_investors as a Excel file
#     df = pd.DataFrame(final_investors)
#     df.to_excel("correct_investors.xlsx", index=False)


# get_correct_investors()


def rollover_investors(effective_date):
    effective_date = datetime.strptime(effective_date.replace("/", "-"), "%Y-%m-%d")
    print("effective_date", effective_date)
    try:
        investors = list(db.investors.find({},
                                           {"_id": 0, "investor_acc_number": 1, "investments": 1, "investor_name": 1,
                                            "investor_surname": 1, "trust": 1}))

        db.rollovers_investors.delete_many({})
        rollovers_investors = list(db.rollovers_investors.find({}))
        if len(rollovers_investors) > 0:
            for rollover in rollovers_investors:
                rollover['_id'] = str(rollover['_id'])
        # print("rollovers_investors", rollovers_investors)
        rates = list(db.rates.find({}, {"_id": 0}))

        for rate in rates:
            rate['Efective_date'] = rate['Efective_date'].replace("-", "/")
            rate['Efective_date'] = datetime.strptime(rate['Efective_date'], "%Y/%m/%d")
            rate['rate'] = float(rate['rate']) / 100

        # sort rates by Efective_date in descending order
        rates = sorted(rates, key=lambda x: x['Efective_date'], reverse=True)

        opportunities = list(db.opportunities.find({"Category": "Heron View"}, {"_id": 0, "Category": 1,
                                                                                "opportunity_amount_required": 1,
                                                                                "opportunity_code": 1}))
        for opportunity in opportunities:
            opportunity['block'] = opportunity['opportunity_code'][2]
            opportunity['opportunity_amount_required'] = float(opportunity['opportunity_amount_required'])
            del opportunity['opportunity_code']
            del opportunity['Category']

        # remove duplicates from opportunities
        opportunities = [dict(t) for t in {tuple(d.items()) for d in opportunities}]

        sold_status = list(db.cashflow_sales.find({}, {"_id": 0}))
        # print("sold_status", sold_status[0])

        # print("opportunities", opportunities)
        # print("opportunities", len(opportunities))

        previous_investments = []

        rollovers_investors_to_edit = []
        final_investors = []
        # filter out of investments where Category is not "Heron View"
        for investor in investors:
            investor["investments"] = list(
                filter(lambda x: x["Category"] == "Heron View", investor["investments"]))
            previous_blocks = []
            for investment in investor['investments']:
                block = investment['opportunity_code'][2]
                previous_blocks.append(block)

            if len(previous_blocks) > 0:
                # remove duplicates from previous_blocks
                previous_blocks = list(set(previous_blocks))
                insert = {
                    "investor": investor['investor_acc_number'],
                    "previous_blocks": previous_blocks,
                }
                previous_investments.append(insert)

            # filter out of investments where end_date is not empty
            investor["investments"] = list(
                filter(lambda x: x["end_date"] == "", investor["investments"]))

            investor["trust"] = list(
                filter(lambda x: x["Category"] == "Heron View", investor["trust"]))

        investors = list(filter(lambda x: len(x["investments"]) > 0, investors))

        # print("previous_investments", previous_investments)
        for investor in investors:

            for investment in investor['investments']:
                filtered_trust = list(filter(lambda x: x['opportunity_code'] == investment['opportunity_code']
                                                       and x['investment_number'] == investment[
                                                           'investment_number'],
                                             investor['trust']))

                filterered_opportunities = list(filter(lambda x: x['block'] == investment['opportunity_code'][2],
                                                       opportunities))

                filtered_sold_status = list(
                    filter(lambda x: x['opportunity_code'] == investment['opportunity_code'], sold_status))
                if len(filtered_sold_status) > 0:
                    sold = filtered_sold_status[0]['sold']
                    transferred = filtered_sold_status[0]['transferred']
                    complete_build = filtered_sold_status[0]['complete_build']
                    forecast_transfer_date = filtered_sold_status[0]['forecast_transfer_date']
                    # convert forecast_transfer_date to datetime
                    if forecast_transfer_date != "":
                        forecast_transfer_date = forecast_transfer_date.replace("-", "/")
                        forecast_transfer_date = datetime.strptime(forecast_transfer_date, '%Y/%m/%d')

                # print("filterered_opportunities", filterered_opportunities)
                filtered_previous_investments = list(
                    filter(lambda x: x['investor'] == investor['investor_acc_number'], previous_investments))
                previously_invested = ""
                if len(filtered_previous_investments) > 0:
                    for index, block in enumerate(filtered_previous_investments[0]['previous_blocks']):
                        if index == len(filtered_previous_investments[0]['previous_blocks']) - 1:
                            previously_invested += block
                        else:
                            previously_invested += block + ", "

                insert = {
                    "investment_amount": float(investment['investment_amount']),
                    "max_investment_amount": filterered_opportunities[0]['opportunity_amount_required'],
                    "opportunity_code": investment['opportunity_code'],
                    # create field called block which is the 3rd character of opportunity_code
                    "Block": investment['opportunity_code'][2],
                    "previously_invested": previously_invested,
                    "sold": sold,
                    "transferred": transferred,
                    "complete_build": complete_build,
                    "forecast_transfer_date": forecast_transfer_date,
                    "Category": investment['Category'],
                    "investor_acc_number": investor['investor_acc_number'],
                    "investor_name": investor['investor_name'],
                    "investor_surname": investor['investor_surname'],
                    "investment_number": investment.get("investment_number", 0),
                    "investment_interest_rate": float(investment.get("investment_interest_rate", 0)) / 100,
                    "deposit_date": datetime.strptime(filtered_trust[0].get("deposit_date", "").replace("/", "-"),
                                                      "%Y-%m-%d"),
                    "release_date": datetime.strptime(investment.get("release_date", "").replace("/", "-"),
                                                      "%Y-%m-%d"),
                    "contract_expiry_date": datetime.strptime(investment.get("release_date", "").replace("/", "-"),
                                                              "%Y-%m-%d") + timedelta(days=731),
                    "rollover_date": effective_date,
                    "roll_from": "",
                    "roll_to": "",
                }
                deposit_date = insert['deposit_date']
                release_date = insert['release_date']
                investment_amount = insert['investment_amount']
                momentum_interest = 0
                while deposit_date < release_date:
                    # filter rates where Efective_date is less than or to the deposit_date
                    filtered_rates = list(filter(lambda x: x['Efective_date'] <= deposit_date, rates))
                    rate = filtered_rates[0]['rate']
                    # calculate interest
                    interest = investment_amount * rate / 365
                    momentum_interest += interest
                    deposit_date += timedelta(days=1)

                insert['momentum_interest'] = momentum_interest
                insert['investment_interest'] = investment_amount * insert['investment_interest_rate'] / 365 * (
                        effective_date - release_date).days
                insert['total_value'] = investment_amount + insert['investment_interest'] + momentum_interest

                # print("insert", insert['investor_acc_number'])

                final_investors.append(insert)
        # print("final_investors", final_investors[0])
        # print("final_investors", len(final_investors))
        # filter out of final_investors where transferred is true
        final_investors = list(filter(lambda x: x['transferred'] == False, final_investors))
        # filter out of final_investors where sold is true and forecast_transfer_date < contract_expiry_date
        # final_investors = list(filter(lambda x: x['sold'] == False and x['forecast_transfer_date'] > x['contract_expiry_date'], final_investors))
        for investor in final_investors:
            if investor['sold'] == True and investor['forecast_transfer_date'] < investor['contract_expiry_date']:
                investor['include'] = False
            elif investor['sold'] == True and investor['forecast_transfer_date'] > investor['contract_expiry_date']:
                investor['include'] = True
            else:
                investor['include'] = True
                # print("investor", investor)

        # filter out of final_investors where include is false
        final_investors = list(filter(lambda x: x['include'] == True, final_investors))
        final_investors_to_insert = []

        # make a deep copy of final_investors
        final_investors_copy = copy.deepcopy(final_investors)
        # filter out
        # sort by contract_expiry_date
        final_investors_copy = sorted(final_investors_copy, key=lambda x: x['contract_expiry_date'])
        for investor in final_investors:
            # if investor['opportunity_code'] == "HVD302":
            #     print("investor", investor)
            filtered_final_investors_copy = list(filter(lambda x: x['Block'] == investor['Block'],

                                                        final_investors_copy))
            investor['earliest_contract_expiry_date'] = filtered_final_investors_copy[0]['contract_expiry_date']

        # sort by earliest_contract_expiry_date and then by opportunity_code
        final_investors = sorted(final_investors,
                                 key=lambda x: (x['earliest_contract_expiry_date'], x['opportunity_code']))

        for investor in final_investors:
            filtered_rollovers_investors = list(
                filter(lambda x: x['investor_acc_number'] == investor['investor_acc_number']
                                 and x['opportunity_code'] == investor['opportunity_code']
                                 and x['investment_number'] == investor['investment_number'],
                       rollovers_investors))
            if len(filtered_rollovers_investors) > 0:
                filtered_rollovers_investors[0]['investment_interest'] = investor['investment_interest']
                rollovers_investors_to_edit.append(filtered_rollovers_investors[0])
            else:
                final_investors_to_insert.append(investor)

        print("rollovers_investors_to_edit", len(rollovers_investors_to_edit))
        print("final_investors_to_insert", len(final_investors_to_insert))

        if len(rollovers_investors_to_edit) > 0:
            for rollover in rollovers_investors_to_edit:
                try:
                    id = rollover["_id"]
                    del rollover["_id"]
                    db.rollovers_investors.update_one({"_id": ObjectId(id)}, {"$set": rollover})
                    print("Success A")
                except Exception as e:
                    print(e)
        if len(final_investors_to_insert) > 0:
            # save this as an Excel file
            df = pd.DataFrame(final_investors_to_insert)
            df.to_excel("rollover_investors.xlsx", index=False)

            try:
                db.rollovers_investors.insert_many(final_investors_to_insert)
                print("Success B")
            except Exception as e:
                print(e)

        print("DONE!!")
        return {"success": True}


    except Exception as e:
        print("Error getting investors", e)
        return {"success": False, "error": str(e)}


# rollover_investors("2024/06/01")

@cashflow.post("/get_daily_cashflow")
async def get_daily_cashflow(request: Request):
    request = await request.json()
    print("request", request)
    date = request['date']
    if request['development'] == "Heron":
        development = ["Heron View", "Heron Fields"]
    else:
        development = [request['development']]

    print("development", development)

    try:
        # if file exists delete it
        if os.path.exists("cashflow_p&l_files/daily_cashflow.xlsx"):
            os.remove("cashflow_p&l_files/daily_cashflow.xlsx")


        investors = list(db.investors.find({}, {"_id": 0, "investor_acc_number": 1, "investments": 1, "trust": 1,
                                                "investor_name": 1, "investor_surname": 1}))
        # filter out of investors where trust is empty
        for investor in investors:
            investor["trust"] = list(filter(lambda x: x["Category"] in development, investor["trust"]))
            investor["investments"] = list(filter(lambda x: x["Category"] in development, investor["investments"]))
        investors = list(filter(lambda x: len(x["trust"]) > 0, investors))

        investor_exit = []

        for investor in investors:
            for trust in investor['trust']:
                filtered_investments = list(filter(lambda x: x['opportunity_code'] == trust['opportunity_code']
                                                   ,
                                                   investor['investments']))
                if len(filtered_investments) > 0:
                    investment_interest_rate = float(filtered_investments[0].get("investment_interest_rate", 0)) / 100
                    if filtered_investments[0].get("end_date", "") != "":
                        end_date = datetime.strptime(filtered_investments[0].get("end_date", "").replace("/", "-"),
                                                     "%Y-%m-%d")
                    else:
                        end_date = datetime.strptime(trust.get("release_date", "").replace("/", "-"),
                                                     "%Y-%m-%d") + timedelta(days=731)
                else:
                    investment_interest_rate = 0
                    end_date = ""
                if trust.get("release_date", "") != "":
                    release_date = datetime.strptime(trust.get("release_date", "").replace("/", "-"), "%Y-%m-%d")
                else:
                    release_date = ""

                if trust["Category"] == "Goodwood":
                    block = "R"
                else:
                    block = trust['opportunity_code'][-4:-3]

                insert = {
                    "investor_acc_number": investor['investor_acc_number'],
                    "investor_name": investor['investor_name'],
                    "investor_surname": investor['investor_surname'],
                    "Category": trust['Category'],
                    "block": block,
                    "investment_number": trust.get('investment_number', 0),
                    "unit_number": trust['opportunity_code'],
                    "deposit_date": datetime.strptime(trust.get("deposit_date", "").replace("/", "-"), "%Y-%m-%d"),
                    "release_date": release_date,
                    "end_date": end_date,
                    "investment_amount": float(trust['investment_amount']),
                    "investment_interest_rate": investment_interest_rate,
                    "trust_interest": 0,
                    "released_interest": 0,
                    "total_interest": 0,
                    "current_exit_date": "",
                }
                investor_exit.append(insert)

        sales = get_sales_data(date)

        # filter sales where Category is in development
        sales = list(filter(lambda x: x['Category'] in development, sales))

        for sale in sales:
            try:
                sale['sale_price'] = float(sale['sale_price'])
            except:
                sale['sale_price'] = 0.0

        opportunities = get_opportunities()
        # filter opportunities where Category is in development
        opportunities = list(filter(lambda x: x['Category'] in development, opportunities))
        # DO FOR INVESTOR EXIT REPORT THEN CREATE A FUNCTION

        current_report_date = date.replace("-", "/")
        current_report_date = datetime.strptime(current_report_date, '%Y/%m/%d')

        # result = background_tasks.add_task(daily_cashflow(sales, investor_exit,date))
        result = daily_cashflow(sales, investor_exit, date)
        print("result", result)
        #
        # return {"sales": sales, "investors": investor_exit, "opportunities": opportunities}
        return {"filename": result['filename'], "success": True}

    except Exception as e:
        print("Error getting daily cashflow", e)
        return {"success": False, "error": str(e)}

@cashflow.get("/get_dailycashflow")
async def get_dailycashflow(file_name):
    try:
        print(file_name)
        file_directory = "cashflow_p&l_files"

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


