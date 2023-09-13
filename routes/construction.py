import os

# from bson import ObjectId
from fastapi import APIRouter, Request, BackgroundTasks
# from fastapi.responses import FileResponse
# from excel_sf_functions.sales_forecast_excel import create_sales_forecast_file
# from excel_sf_functions.sales_forecast_excel import create_cash_flow
from config.db import db
# import time
# from datetime import datetime
# from datetime import timedelta
# from excel_sf_functions.draw_downs_excel import create_draw_down_file
# # from excel_sf_functions.investment_list_excel import create_investment_list_file
from construction_payment_advice_files.valuations_excel_file import create_valuations_file

#
# # from fastapi import File, UploadFile
# import smtplib
# from email.message import EmailMessage

construction = APIRouter()

# MONGO COLLECTIONS
# investors = db.investors
# rates = db.rates
# opportunities = db.opportunities
# unallocated_investments = db.unallocated_investments
# sales_parameters = db.salesParameters
# rollovers = db.investorRollovers
# drawdowns = db.investor_Draws
valuations = db.constructionValuations


@construction.post("/construction_valuations")
async def create_construction_valuation(request: Request):
    data = await request.json()
    print(data)
    # get valuations from db where development = data['development'] and block = data['block']
    valuations_to_date = list(valuations.find({"development": data['development'], "block": data['block']}))
    # print(valuations_to_date)
    for valuation in valuations_to_date:
        valuation['_id'] = str(valuation['_id'])
        filtered_tasks = [task for task in valuation['tasks'] if task['approved']]
        valuation['qty'] = int(valuation['qty'])
        valuation['rate'] = float(valuation['rate'])
        # if index % 2 == 0:
        valuation['terms'] = valuation.get('terms', "30 Days")
        # else:
        #     valuation['terms'] = valuation.get('terms', "End of Month")
        if len(filtered_tasks) > 0:
            valuation['percent_complete'] = filtered_tasks[-1]['currentProgress']
            valuation['amount_complete'] = valuation['amount'] * (valuation['percent_complete'] / 100)
        else:
            valuation['percent_complete'] = 0
            valuation['amount_complete'] = 0
        valuation['tasks'].reverse()
    # only unique values remain
    subcontractors = list(set([valuation['subcontractor'] for valuation in valuations_to_date]))
    # sort subcontractors alphabetically
    subcontractors.sort()

    for index, subcontractor in enumerate(subcontractors):
        for valuation in valuations_to_date:
            if valuation['subcontractor'] == subcontractor and valuation['subcontractor'] != "ZZWayne":
                if index % 2 == 0:
                    valuation['terms'] = "30 Days"
                else:
                    valuation['terms'] = "End of Month"

    file_created = create_valuations_file(valuations_to_date, subcontractors)
    print(file_created)

    # print(subcontractors)

    return {"status": "ok", "length": len(valuations_to_date), "valuations": valuations_to_date}


@construction.post("/construction_valuations_submit")
async def submit_construction_valuation(request: Request):
    data = await request.json()
    print(data)
    valuations_to_approve = list(valuations.find({"development": data['development'], "block": data['block']}))
    for valuation in valuations_to_approve:
        valuation['_id'] = str(valuation['_id'])

    return {"status": "ok", "length": len(data), "valuations": data}