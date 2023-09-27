import os
import copy

from bson import ObjectId
# from bson import ObjectId
from fastapi import APIRouter, Request, BackgroundTasks
from fastapi.responses import FileResponse

from config.db import db

from construction_payment_advice_files.valuations_excel_file import create_valuations_file

construction = APIRouter()

# MONGO COLLECTIONS

valuations = db.constructionValuations


@construction.get("/available_valuation_data")
async def get_available_valuation_data():
    try:
        pipeline = [
            {
                "$group": {
                    "_id": {
                        "development": "$development",
                        "block": "$block"
                    }
                }
            },
            {
                "$project": {
                    "_id": 0,
                    "development": "$_id.development",
                    "block": "$_id.block"
                }
            }
        ]
        available_data = list(valuations.aggregate(pipeline))

        available_data.sort(key=lambda x: (x['development'], x['block']))
        # print(available_data)
        return {"status": "ok", "available_data": available_data}
    except Exception as e:
        print("XXXX",e)
        return {"status": "error", "error": e}


@construction.post("/construction_valuations")
async def create_construction_valuation(request: Request):
    try:
        data = await request.json()
        # print(data[0])
        # get valuations from db where development = data['development'] and block = data['block']
        valuations_to_date = list(valuations.find({"development": data['development'], "block": data['block']}))

        # print(valuations_to_date)
        if len(valuations_to_date) > 0:

            for index, valuation in enumerate(valuations_to_date):
                valuation['_id'] = str(valuation['_id'])
                filtered_tasks = [task for task in valuation['tasks'] if task['approved']]
                valuation['qty'] = int(valuation['qty'])
                valuation['rate'] = float(valuation['rate'])
                # if index % 2 == 0:
                valuation['taskCategory'] = valuation.get('taskCategory','Normal')
                # if valuation['taskCategory'] == "ATCV":
                #     valuation['terms'] = valuations_to_date[index - 1]['terms']
                #     if valuation['subcontractor'] == "Taaibosch Development & Projects Pty Ltd - Heron View Block D":
                #
                #         print(valuations_to_date[index - 1])
                #         print()
                #         print(valuation['terms'])
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
                for index2, valuation in enumerate(valuations_to_date):
                    if valuation['subcontractor'] == subcontractor and valuation['taskCategory'] == "ATCV":
                        filtered_valuations_to_date = [valuation for valuation in valuations_to_date if
                                                         valuation['subcontractor'] == subcontractor]
                        valuation['terms'] = filtered_valuations_to_date[0]['terms']
                        valuation['works'] = filtered_valuations_to_date[0]['works']
                        # if index % 2 == 0:
                        #     valuation['terms'] = "30 Days"
                        # else:
                        #     valuation['terms'] = "End of Month"

            file_created = create_valuations_file(valuations_to_date, subcontractors)
            # print(file_created)
            file_created = file_created.split("/")[1]

            # print(subcontractors)

            return {"status": "ok", "filename": file_created }

        else:
            return {"status": "error", "error": "No Valuations to date"}
    except Exception as e:
        print("ERROR",e)
        return {"status": "error", "error": e}


@construction.post("/construction_valuations_submit")
async def submit_construction_valuation(request: Request):
    data = await request.json()
    # print(data)

    try:
        valuations_to_approve = list(valuations.find({"development": data['development'], "block": data['block']}))
        subcontractors_needing_approval = []
        for valuation in valuations_to_approve:
            valuation['_id'] = str(valuation['_id'])
            subcontractors_needing_approval.append(valuation['subcontractor'])

        subcontractors_needing_approval = list(set(subcontractors_needing_approval))
        final_valuations_jobs = []
        for subcontractor in subcontractors_needing_approval:
            valuation_required = 0
            valuations_to_approve_filtered = [valuation for valuation in valuations_to_approve if
                                              valuation['subcontractor'] == subcontractor]
            for valuation in valuations_to_approve_filtered:
                # print(subcontractor,valuation['tasks'])
                if len(valuation['tasks']) > 0:
                    valuation_required += valuation['tasks'][-1:][0]['currentProgress'] - valuation['tasks'][-1:][0][
                        'initialProgress']
            # print(subcontractor, valuation_required)
            if valuation_required > 0:
                final_valuations_jobs.append(subcontractor)

        # print(final_valuations_jobs)
        # sort final_valuations_jobs alphabetically
        final_valuations_jobs.sort()

        # sort valuations_to_approve alphabetically by subcontractor
        valuations_to_approve.sort(key=lambda x: x['subcontractor'])

        for subcontractor in final_valuations_jobs:
            for index, job_to_approve in enumerate(valuations_to_approve, 1):
                if subcontractor == job_to_approve['subcontractor']:
                    # print(subcontractor, job_to_approve)
                    id = job_to_approve['_id']
                    del job_to_approve['_id']

                    insert = copy.deepcopy(job_to_approve['tasks'][-1:])
                    job_to_approve['tasks'][-1:][0]['approved'] = True
                    job_to_approve['tasks'][-1:][0]['status'] = "Approved"
                    insert[0]['initialProgress'] = insert[0]['currentProgress']
                    if insert[0]['paymentAdviceNumber'] == "" or insert[0]['paymentAdviceNumber'] == None:
                        insert[0]['paymentAdviceNumber'] = "0"
                    else:
                        paNumber = insert[0]['paymentAdviceNumber'].rsplit("-", 1)[1]
                    initial_past_of_number = insert[0]['paymentAdviceNumber'].rsplit("-", 1)[0]
                    # convert paNumber to int, add 1, convert back to string of three characters and add to
                    # initial_past_of_number
                    paNumber = str(int(paNumber) + 1).zfill(3)
                    insert[0]['paymentAdviceNumber'] = initial_past_of_number + "-" + paNumber
                    job_to_approve['tasks'][-1:][0]['date'] = data['date']
                    insert[0]['date'] = data['date']

                    job_to_approve['tasks'].append(insert[0])

                    # update db with job_to_approve
                    response = valuations.update_one({"_id": ObjectId(id)}, {"$set": job_to_approve})
                    # print(response)

        return {"status": "ok", "length": len(data), "valuations": data}
    except Exception as e:
        print(e)
        return {"status": "error", "error": e}


@construction.post("/construction_valuations_to_approve")
async def get_construction_valuations_to_approve(request: Request):

    data = await request.json()
    # print(data)
    # get valuations from db where development = data['development'] and block = data['block']
    valuations_to_approve = list(valuations.find({"development": data['development'], "block": data['block']}))
    # print(valuations_to_approve)
    # print(valuations_to_approve[0])
    # try:

    if len(valuations_to_approve) > 0:

        total_valuation_30_days = 0
        total_valuation_end_of_month = 0
        for valuation in valuations_to_approve:
            # if len(valuation['tasks']) == 0:
            #     print("VALUATION", valuation)
            #     pass

            retention = valuation.get('retention', 0)
            if valuation['retention'] == None:
                valuation['retention'] = 0
            if retention == None:
                retention = 0
            # print("Retention", retention)
            valuation['_id'] = str(valuation['_id'])
            task_filtered = [task for task in valuation['tasks'] if not task['approved']]
            valuation['terms'] = valuation.get('terms', "30 Days")
            # print("XXXX", valuation['subcontractor'])
            # print("XXXX", valuation['amount'])
            if valuation['terms'] == "30 Days":
                if valuation['vatable'] == "Yes":

                    total_valuation_30_days += ((float(valuation['amount']) * (task_filtered[0]["currentProgress"] -
                                                                               task_filtered[0][
                                                                                   "initialProgress"]) / 100) - (
                                                        float(retention) / 100 * (float(valuation['amount']) * (
                                                        (task_filtered[0]["currentProgress"] -
                                                         task_filtered[0]["initialProgress"]) / 100)))) * 1.15


                else:
                    total_valuation_30_days += (float(valuation['amount']) * (task_filtered[0]["currentProgress"] -
                                                                              task_filtered[0][
                                                                                  "initialProgress"]) / 100) - (
                                                       float(retention) / 100 * (float(valuation['amount']) * (
                                                       (task_filtered[0]["currentProgress"] -
                                                        task_filtered[0]["initialProgress"]) / 100)))
            else:
                if valuation['vatable'] == "Yes":
                    total_valuation_end_of_month += ((float(valuation['amount']) * (
                                task_filtered[0]["currentProgress"] -
                                task_filtered[0][
                                    "initialProgress"]) / 100) - (
                                                             float(retention) / 100 * (float(valuation['amount']) * (
                                                             (task_filtered[0]["currentProgress"] -
                                                              task_filtered[0]["initialProgress"]) / 100)))) * 1.15


                else:
                    total_valuation_end_of_month += (float(valuation['amount']) * (task_filtered[0]["currentProgress"] -
                                                                                   task_filtered[0][
                                                                                       "initialProgress"]) / 100) - (
                                                            float(retention) / 100 * (float(valuation['amount']) * (
                                                            (task_filtered[0]["currentProgress"] -
                                                             task_filtered[0]["initialProgress"]) / 100)))

        total_approvals = total_valuation_30_days + total_valuation_end_of_month
        # make all three values currency with two decimal places and symbol R
        total_valuation_30_days = f"R{total_valuation_30_days:,.2f}"
        total_valuation_end_of_month = f"R{total_valuation_end_of_month:,.2f}"
        total_approvals = f"R{total_approvals:,.2f}"
        # print("XXXX",total_valuation_30_days, total_valuation_end_of_month, total_approvals)

        return {"30 Days": total_valuation_30_days, "total EOM": total_valuation_end_of_month,
                "Total Approvals": total_approvals}
    # except Exception as e:
    #     print("ER",e)
    #     return {"status": "error", "error": e}


@construction.get("/get_valuation_report")
async def sales_forecast(valuation_report_name: str):
    file_name = valuation_report_name
    dir_path = "excel_files"
    dir_list = os.listdir(dir_path)
    # print(f"{dir_path}/{file_name}")
    if file_name in dir_list:
        return FileResponse(f"{dir_path}/{file_name}", filename=file_name)
    else:
        return {"ERROR": "File does not exist!!"}
