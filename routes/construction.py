import csv
import os
import copy
import time

import boto3
from botocore.exceptions import ClientError
from decouple import config

from bson import ObjectId
# from bson import ObjectId
from fastapi import APIRouter, Request, BackgroundTasks
from fastapi.responses import FileResponse

from configuration.db import db

from docxtpl import DocxTemplate

from construction_payment_advice_files.valuations_excel_file import create_valuations_file

construction = APIRouter()

# MONGO COLLECTIONS

valuations = db.constructionValuations
investor_draws = db.investor_Draws

# AWS BUCKET INFO - ENSURE IN VARIABLES ON HEROKU
AWS_BUCKET_NAME = config("AWS_BUCKET_NAME")
AWS_BUCKET_REGION = config("AWS_BUCKET_REGION")
AWS_ACCESS_KEY_ID = config("AWS_ACCESS_KEY_ID")
AWS_SECRET_ACCESS_KEY = config("AWS_SECRET_ACCESS_KEY")

# CONNECT TO S3 SERVICE
s3 = boto3.client(
    "s3",
    aws_access_key_id=AWS_ACCESS_KEY_ID,
    aws_secret_access_key=AWS_SECRET_ACCESS_KEY
)


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
        print("XXXX", e)
        return {"status": "error", "error": e}


@construction.post("/construction_valuations")
async def create_construction_valuation(request: Request):
    try:
        data = await request.json()
        # print("data",data)
        # print(data[0])
        # get valuations from db where development = data['development'] and block = data['block']
        valuations_to_date = list(valuations.find({"development": data['development'], "block": data['block']}))


        # print("valuations_to_date::",valuations_to_date)
        if len(valuations_to_date) > 0:

            for index, valuation in enumerate(valuations_to_date):
                # if valuation['measure'] != '%':
                #     print("valuation", valuation)
                #     print()
                valuation['_id'] = str(valuation['_id'])
                filtered_tasks = [task for task in valuation['tasks'] if task['approved']]
                valuation['qty'] = int(valuation['qty'])
                valuation['rate'] = float(valuation['rate'])
                # if index % 2 == 0:
                valuation['taskCategory'] = valuation.get('taskCategory', 'Normal')

                if len(filtered_tasks) > 0:
                    valuation['percent_complete'] = filtered_tasks[-1]['currentProgress']
                    if valuation['measure'] == '%':
                        valuation['amount_complete'] = valuation['amount'] * (valuation['percent_complete'] / 100)
                    else:
                        valuation['amount_complete'] = valuation['rate'] * (valuation['percent_complete'])
                else:
                    valuation['percent_complete'] = 0
                    valuation['amount_complete'] = 0
                valuation['tasks'].reverse()
            # only unique values remain
            subcontractors = list(set([valuation['subcontractor'] for valuation in valuations_to_date]))
            # sort subcontractors alphabetically
            subcontractors.sort()

            for index, subcontractor in enumerate(subcontractors):
                # if subcontractor
                # print("subcontractor", subcontractor)
                # print()
                for index2, valuation in enumerate(valuations_to_date):
                    # if valuation['subcontractor'] == "ZZWayne4" and valuation['measure'] != "%":
                    #     print("valuation", valuation)
                    #     print()
                    # print("valuation", valuation)
                    # print()
                    if valuation['subcontractor'] == subcontractor and valuation['taskCategory'] == "ATCV":
                        filtered_valuations_to_date = [valuation for valuation in valuations_to_date if
                                                       valuation['subcontractor'] == subcontractor]
                        # print("XXXX", filtered_valuations_to_date[0])
                        if len(filtered_valuations_to_date) > 0:
                            valuation['terms'] = filtered_valuations_to_date[0].get('terms', "30 Days")
                            # print("Hello:", valuation['terms'])
                        else:
                            valuation['terms'] = "30 Days"
                            # print("Goodbye", valuation['terms'])
                        valuation['works'] = filtered_valuations_to_date[0]['works']
                        # print("valuation", valuation)
                        # if index % 2 == 0:
                        #     valuation['terms'] = "30 Days"
                        # else:
                        #     valuation['terms'] = "End of Month"

            file_created = create_valuations_file(valuations_to_date, subcontractors)
            print(file_created)
            file_created = file_created.split("/")[1]

            # print(subcontractors)

            return {"status": "ok", "filename": file_created}

        else:
            return {"status": "error", "error": "No Valuations to date"}
    except Exception as e:
        print("ERROR", e)
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
    # print("valuations_to_approve", valuations_to_approve)
    final_valuations_jobs = []
    for index, valuation in enumerate(valuations_to_approve):
        valuation['_id'] = str(valuation['_id'])

        for task in valuation['tasks']:

            insert = {}
            insert['development'] = valuation['development']
            insert['block'] = valuation['block']
            insert['subcontractor'] = valuation['subcontractor']
            insert['qty'] = valuation['qty']
            insert['rate'] = valuation['rate']
            insert['measure'] = valuation.get('measure', "%")
            insert['amount'] = valuation['amount']
            insert['taskCategory'] = valuation.get('taskCategory', "Normal")
            insert['initialProgress'] = task['initialProgress']
            insert['currentProgress'] = task['currentProgress']
            if valuation['vatable'] == "Yes":
                insert['vatable'] = 1.15
            else:
                insert['vatable'] = 1.0
            # insert['vatable'] = valuation['vatable']
            insert['retention'] = valuation.get('retention', 0)
            insert['approved'] = task['approved']
            insert['status'] = task.get('status', "Pending")
            insert['link'] = task.get('link', "")
            if task['paymentAdviceNumber'] == "" or task['paymentAdviceNumber'] == None:
                # filter valuations_to_approve where subcontractor = subcontractor
                valuations_to_approve_filtered = [valuation for valuation in valuations_to_approve if
                                                  valuation['subcontractor'] == insert['subcontractor']]
                # print("valuations_to_approve_filtered", valuations_to_approve_filtered[0])
                # print()
                # print("valuation id:", valuation['_id'])

                previous_task = valuations_to_approve_filtered[0]['tasks'][-1:][0]['paymentAdviceNumber']

                to_update = valuations.find_one({"_id": ObjectId(valuation['_id'])})
                # print("to_update", to_update)
                to_update['_id'] = str(to_update['_id'])
                for index, task in enumerate(to_update['tasks'], 1):
                    if index == len(to_update['tasks']):
                        task['paymentAdviceNumber'] = previous_task
                        task['status'] = "Pending"

                # print("to_update", to_update)
                id = to_update['_id']
                del to_update['_id']
                valuations.update_one({"_id": ObjectId(id)}, {"$set": to_update})

                # print("previous_task", previous_task)

                insert['paymentAdviceNumber'] = previous_task
            else:
                insert['paymentAdviceNumber'] = task['paymentAdviceNumber']
                # get the last paymentAdviceNumber from previous task for the same subcontractor

                # if len(previous_task) > 0:
                #     insert['paymentAdviceNumber'] = previous_task[-1:][0]['paymentAdviceNumber']
                # else:
                #     insert['paymentAdviceNumber'] = "0"

                # insert['paymentAdviceNumber'] = task.get('paymentAdviceNumber',

            final_valuations_jobs.append(insert)

    if len(valuations_to_approve) > 0:

        total_valuation_30_days = 0
        total_valuation_end_of_month = 0
        for valuation in valuations_to_approve:

            retention = valuation.get('retention', 0)
            valuation['link'] = valuation.get('link', "")
            if valuation['retention'] is None:
                valuation['retention'] = 0
            if retention is None:
                retention = 0

            valuation['_id'] = str(valuation['_id'])
            task_filtered = [task for task in valuation['tasks'] if not task['approved']]

            if len(task_filtered) > 0:
                if valuation.get('terms', "30 Days") == "30 Days":
                    # print(valuation['measure'])

                    if valuation['vatable'] == "Yes":

                        if valuation['measure'] == '%':

                            total_valuation_30_days += ((float(valuation['amount']) * (
                                        task_filtered[0]["currentProgress"] -
                                        task_filtered[0][
                                            "initialProgress"]) / 100) - (
                                                                float(retention) / 100 * (float(valuation['amount']) * (
                                                                (task_filtered[0]["currentProgress"] -
                                                                 task_filtered[0]["initialProgress"]) / 100)))) * 1.15
                            # print("total_valuation_30_days %", total_valuation_30_days)

                        else:
                            total_valuation_30_days += ((float(valuation['rate']) * (
                                    task_filtered[0]["currentProgress"] -
                                    task_filtered[0][
                                        "initialProgress"])) - (
                                                                float(retention) / 100 * (float(valuation['rate']) * (
                                                            (task_filtered[0]["currentProgress"] -
                                                             task_filtered[0]["initialProgress"]))))) * 1.15
                            # print("total_valuation_30_days Unit", ((float(valuation['rate']) * (
                            #         task_filtered[0]["currentProgress"] -
                            #         task_filtered[0][
                            #             "initialProgress"])) - (
                            #                                     float(retention) / 100 * (float(valuation['rate']) * (
                            #                                 (task_filtered[0]["currentProgress"] -
                            #                                  task_filtered[0]["initialProgress"]))))) * 1.15)


                    else:
                        if valuation['measure'] == '%':
                            total_valuation_30_days += (float(valuation['amount']) * (
                                        task_filtered[0]["currentProgress"] -
                                        task_filtered[0][
                                            "initialProgress"]) / 100) - (
                                                               float(retention) / 100 * (float(valuation['amount']) * (
                                                               (task_filtered[0]["currentProgress"] -
                                                                task_filtered[0]["initialProgress"]) / 100)))
                        else:
                            total_valuation_30_days += (float(valuation['rate']) * (
                                    task_filtered[0]["currentProgress"] -
                                    task_filtered[0][
                                        "initialProgress"])) - (
                                                               float(retention) / 100 * (float(valuation['rate']) * (
                                                           (task_filtered[0]["currentProgress"] -
                                                            task_filtered[0]["initialProgress"]))))

                else:
                    if valuation['vatable'] == "Yes":
                        if valuation['measure'] == '%':
                            total_valuation_end_of_month += ((float(valuation['amount']) * (
                                    task_filtered[0]["currentProgress"] -
                                    task_filtered[0][
                                        "initialProgress"]) / 100) - (
                                                                     float(retention) / 100 * (
                                                                     float(valuation['amount']) * (
                                                                     (task_filtered[0]["currentProgress"] -
                                                                      task_filtered[0][
                                                                          "initialProgress"]) / 100)))) * 1.15
                        else:
                            total_valuation_end_of_month += ((float(valuation['rate']) * (
                                    task_filtered[0]["currentProgress"] -
                                    task_filtered[0][
                                        "initialProgress"])) - (
                                                                     float(retention) / 100 * (
                                                                     float(valuation['rate']) * (
                                                                     (task_filtered[0]["currentProgress"] -
                                                                      task_filtered[0][
                                                                          "initialProgress"]))))) * 1.15


                    else:
                        if valuation['measure'] == '%':
                            total_valuation_end_of_month += (float(valuation['amount']) * (
                                    task_filtered[0]["currentProgress"] -
                                    task_filtered[0][
                                        "initialProgress"]) / 100) - (
                                                                    float(retention) / 100 * (float(valuation['amount']) * (
                                                                    (task_filtered[0]["currentProgress"] -
                                                                     task_filtered[0]["initialProgress"]) / 100)))
                        else:
                            total_valuation_end_of_month += (float(valuation['rate']) * (
                                    task_filtered[0]["currentProgress"] -
                                    task_filtered[0][
                                        "initialProgress"])) - (
                                                                    float(retention) / 100 * (
                                                                        float(valuation['rate']) * (
                                                                    (task_filtered[0]["currentProgress"] -
                                                                     task_filtered[0]["initialProgress"]))))

        total_approvals = total_valuation_30_days + total_valuation_end_of_month
        # make all three values currency with two decimal places and symbol R
        total_valuation_30_days = f"R{total_valuation_30_days:,.2f}"
        total_valuation_end_of_month = f"R{total_valuation_end_of_month:,.2f}"
        total_approvals = f"R{total_approvals:,.2f}"

        # remove duplicates from final_valuations_jobs using list comprehension
        final_valuations_jobs = [dict(t) for t in {tuple(d.items()) for d in final_valuations_jobs}]
        # sort final_valuations_jobs alphabetically by subcontractor, then by paymentAdviceNumber, then by status
        final_valuations_jobs.sort(key=lambda x: (x['subcontractor'], x['paymentAdviceNumber'], x['status']))

        payment_advice_numbers = []

        for final_valuation in final_valuations_jobs:
            payment_advice_numbers.append(final_valuation['paymentAdviceNumber'])

        # show only unique payment_advice_numbers
        payment_advice_numbers = list(set(payment_advice_numbers))
        final_data = []
        # print("payment_advice_numbers", payment_advice_numbers)
        for number in payment_advice_numbers:
            filtered_jobs = [job for job in final_valuations_jobs if job['paymentAdviceNumber'] == number]

            insert = {}
            insert['paymentAdviceNumber'] = number
            insert['subcontractor'] = filtered_jobs[0]['subcontractor']
            insert['link'] = filtered_jobs[0]['link']
            value = 0.0
            for job in filtered_jobs:
                # if job['measure'] != "%":
                    # print("job", job)
                    # print()
                if job['retention'] is None:
                    job['retention'] = 0
                if job['measure'] == "%":
                    value += ((float(job['amount']) * (job['currentProgress'] - job['initialProgress']) / 100) * (
                            1 - job.get('retention', 0.0) / 100)) * job['vatable']
                else:
                    value += ((float(job['rate']) * (job['currentProgress'] - job['initialProgress'])) * (
                            1 - job.get('retention', 0.0) / 100)) * job['vatable']
                # print("value", value)
                # print()

                # if job['vatable'] == "Yes":
                #     value = value * 1.15

            insert['value'] = value
            insert['status'] = filtered_jobs[0]['status']
            final_data.append(insert)

        # print("final_data", final_data)
        # sort final_data alphabetically by subcontractor, then by paymentAdviceNumber, then by status
        final_data.sort(key=lambda x: (x['subcontractor'], x['paymentAdviceNumber'], x['status']))
        # filter out records with a value of 0
        final_data = [data for data in final_data if data['value'] > 0]
        for data in final_data:
            data['value_nice'] = f"R{data['value']:,.2f}"

        return {"30 Days": total_valuation_30_days, "total EOM": total_valuation_end_of_month,
                "Total Approvals": total_approvals, "valuations": final_data}
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


@construction.post("/create_payment_advice")
async def create_payment_advice(request: Request):
    data = await request.json()

    try:
        start_time = time.time()
        # get valuations from db where subcontractor = data['subcontractor']
        valuations_to_approve = list(valuations.find({"subcontractor": data['subcontractor']}))
        # print("valuations_to_approve", valuations_to_approve[0])
        final_valuations_jobs = []

        for valuation in valuations_to_approve:
            valuation['_id'] = str(valuation['_id'])
            # filter valuation['tasks'] where valuation['tasks']['paymentAdviceNumber'] == data['paymentAdviceNumber']
            filtered_tasks = [task for task in valuation['tasks'] if
                              task['paymentAdviceNumber'] == data['paymentAdviceNumber']]
            if len(filtered_tasks) > 0:
                measure = valuation.get('measure', "%")
                if measure == "%":
                    calculable_amount = valuation['amount']
                    initialProgress = filtered_tasks[0]['initialProgress'] / 100
                    currentProgress = filtered_tasks[0]['currentProgress'] / 100
                else:
                    calculable_amount = valuation['rate']
                    initialProgress = filtered_tasks[0]['initialProgress']
                    currentProgress = filtered_tasks[0]['currentProgress']
                insert = {}
                insert['development'] = valuation['development']
                insert['block'] = valuation['block']
                insert['subcontractor'] = valuation['subcontractor']
                insert['qty'] = valuation['qty']
                insert['rate'] = valuation['rate']
                insert['amount'] = valuation['amount']
                insert['calculable_amount'] = calculable_amount
                insert['works'] = valuation.get('works', "")
                insert['taskCategory'] = valuation.get('taskCategory', "Normal")
                insert['initialProgress'] = initialProgress
                insert['currentProgress'] = currentProgress
                insert['measure'] = valuation.get('measure', "%")
                insert['is_vatable'] = valuation['vatable']
                insert['retention'] = valuation.get('retention', 0)
                insert['approved'] = filtered_tasks[0]['approved']
                insert['status'] = filtered_tasks[0].get('status', "Pending")
                insert['paymentAdviceNumber'] = filtered_tasks[0]['paymentAdviceNumber']
                final_valuations_jobs.append(insert)

        # print("final_valuations_jobs", final_valuations_jobs[0])
        # print()
        # filter out records with a measure of "%"
        final_valuations_jobs_test = [valuation for valuation in final_valuations_jobs if valuation['measure'] != "%"]
        # print("final_valuations_jobs_test", final_valuations_jobs_test[0])

        retention = float(final_valuations_jobs[0].get('retention', 0)) / 100

        # total_contract_value is the sum of all the amounts in the final_valuations_jobs using list comprehension
        # excluding taskCategory == "ATCV"
        total_contract_value = sum(
            [valuation['amount'] for valuation in final_valuations_jobs if valuation['taskCategory'] != "ATCV"])
        # total_valuation_current is the sum of the amounts multiplied by currentProgress / 100 in the
        # final_valuations_jobs using list comprehension excluding taskCategory == "ATCV"
        # print("Measure: ",valuation['measure'])

        # if valuation['measure'] == "%":
        total_valuation_current = sum(
            [(valuation['calculable_amount'] * valuation['currentProgress']) for valuation in final_valuations_jobs if
             valuation['taskCategory'] != "ATCV"])
        # total_valuation_previous is the sum of the amounts multiplied by initialProgress / 100 in the
        # final_valuations_jobs using list comprehension excluding taskCategory == "ATCV"
        total_valuation_previous = sum(
            [(valuation['calculable_amount'] * (valuation['initialProgress'])) for valuation in final_valuations_jobs if
             valuation['taskCategory'] != "ATCV"]) - (
                                           sum([(valuation['calculable_amount'] * (valuation['initialProgress'])) for
                                                valuation in final_valuations_jobs if
                                                valuation['taskCategory'] != "ATCV"]) * retention)
        # else:
        #     total_valuation_current = sum(
        #         [(valuation['rate'] * (valuation['currentProgress'] / 100)) for valuation in final_valuations_jobs if
        #          valuation['taskCategory'] != "ATCV"])
        #     print("total_valuation_current", total_valuation_current)
        #     # total_valuation_previous is the sum of the amounts multiplied by initialProgress / 100 in the
        #     # final_valuations_jobs using list comprehension excluding taskCategory == "ATCV"
        #     total_valuation_previous = sum(
        #         [(valuation['rate'] * (valuation['initialProgress'] / 100)) for valuation in final_valuations_jobs if
        #          valuation['taskCategory'] != "ATCV"]) - (
        #                                        sum([(valuation['amount'] * (valuation['initialProgress'] / 100)) for
        #                                             valuation in final_valuations_jobs if
        #                                             valuation['taskCategory'] != "ATCV"]) * retention)

        # total_atcv_to_date is the sum of the amounts multiplied by currentProgress / 100 in the
        # final_valuations_jobs using list comprehension where taskCategory == "ATCV"
        total_atcv_to_date = sum(
            [(valuation['calculable_amount'] * (valuation['currentProgress'])) for valuation in final_valuations_jobs if
             valuation['taskCategory'] == "ATCV"])
        # total_atcv_previous is the sum of the amounts multiplied by initialProgress / 100 in the
        # final_valuations_jobs using list comprehension where taskCategory == "ATCV"
        total_atcv_previous = sum(
            [(valuation['calculable_amount'] * (valuation['initialProgress'] / 100)) for valuation in
             final_valuations_jobs if
             valuation['taskCategory'] == "ATCV"])

        total_work_done = total_valuation_current + total_atcv_to_date
        total_retention = total_valuation_current * retention

        if retention == 0:
            isRetention = "No"
        else:
            isRetention = "Yes"

        if final_valuations_jobs[0]['is_vatable'] == "Yes":
            vat = 0.15
        else:
            vat = 0.0

        doc = DocxTemplate("payment_advice/Payment Advice template.docx")

        context = {
            "paymentAdviceNumber": data['paymentAdviceNumber'].rsplit("-")[-1],
            "interim": "Interim",
            "subcontractor": final_valuations_jobs[0]['subcontractor'],
            "works": final_valuations_jobs[0]['works'],
            "total_contract_value": f"R{total_contract_value:,.2f}",
            "total_atcvs": f"R{total_atcv_to_date:,.2f}",
            "current_subcontract_value": f"R{(total_contract_value + total_atcv_to_date):,.2f}",
            "total_valuation_current": f"R{total_valuation_current:,.2f}",
            "total_retention": f"R{total_retention:,.2f}",
            "retention": f"{retention * 100}%",
            "isRetention": isRetention,
            "retention_amount": f"R{(total_valuation_current * retention):,.2f}",
            "total_atcv_due": f"R{total_atcv_to_date:,.2f}",
            "gross_amount_certified": f"R{((total_valuation_current - (total_valuation_current * retention)) + total_atcv_to_date):,.2f}",
            "previously_certified": f"R{(total_valuation_previous + total_atcv_previous):,.2f}",
            "nett_amount_certified": f"R{((total_valuation_current - (total_valuation_current * retention)) + total_atcv_to_date) - (total_valuation_previous + total_atcv_previous):,.2f}",
            "vat": f"R{(((total_valuation_current - (total_valuation_current * retention)) + total_atcv_to_date) - (total_valuation_previous + total_atcv_previous)) * vat:,.2f}",
            "nett": f"R{(((total_valuation_current - (total_valuation_current * retention)) + total_atcv_to_date) - (total_valuation_previous + total_atcv_previous)) + ((((total_valuation_current - (total_valuation_current * retention)) + total_atcv_to_date) - (total_valuation_previous + total_atcv_previous)) * vat):,.2f}",
            "execution": f"{((total_valuation_current / total_contract_value) * 100).__round__(2)}%",
        }

        doc.render(context)
        doc.save(f"payment_advice/{data['paymentAdviceNumber']}.docx")

        file = f"{data['paymentAdviceNumber']}.docx"

        try:
            s3.upload_file(
                f"payment_advice/{file}",
                AWS_BUCKET_NAME,
                f"{file}",
            )
            link = f"https://{AWS_BUCKET_NAME}.s3.{AWS_BUCKET_REGION}.amazonaws.com/{file}"
            for valuation in valuations_to_approve:
                id = valuation['_id']
                del valuation['_id']
                for task in valuation['tasks']:
                    if task['paymentAdviceNumber'] == data['paymentAdviceNumber']:
                        task['link'] = link
                        # task['status'] = "Sent"
                        # print("task", task)
                result = valuations.update_one({"_id": ObjectId(id)}, {"$set": valuation})

            # delete all files in payment_advice folder except the template "Payment Advice template.docx"
            dir_path = "payment_advice"
            dir_list = os.listdir(dir_path)
            for file in dir_list:
                if file != "Payment Advice template.docx":
                    os.remove(f"{dir_path}/{file}")

            print("SUCCESS")

        except ClientError as err:
            print("Credentials are incorrect")
            print(err)
            link = ""

        end_time = time.time()
        print("Time taken to create payment advice:", (end_time - start_time).__round__(2))

        return {"link": link, "status": "ok", "valuation": data['paymentAdviceNumber'],
                "time": (end_time - start_time).__round__(2)}

    except Exception as e:
        print("ERROR", e)
        return {"status": "error", "error": e}
