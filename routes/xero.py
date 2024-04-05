import copy
import csv
import math
import os
from datetime import datetime, timedelta
import xmltodict
import requests
import time

from dateutil.relativedelta import relativedelta
from fastapi import APIRouter, Request
import httpx
from base64 import b64encode
from urllib.parse import urlencode
from fastapi.responses import FileResponse
from fastapi.responses import RedirectResponse
from configuration.db import db
from bson.objectid import ObjectId
from decouple import config
import time

from cashflow_excel_functions.cpc_profit_loss_files import insert_data_from_xero_profit_loss
import cashflow_excel_functions.cpc_data_fields as cpc_data_fields
from cashflow_excel_functions.cashflow_hf_hv_functions import cashflow_hf_hv

xero = APIRouter()
# MONGODB DOCUMENTS
xeroCredentials = db.xeroCredentials
xeroTenants = db.xeroTenants
# XERO CREDENTIALS
clientId = config("clientId")
xerosecret = config("xerosecret")
redirectURI = config("redirectURI")


@xero.get("/authorizeApp")
async def authorize_app():
    # get refresh_expires from xeroCredentials

    # Construct the authorization URL
    authorization_url = (
        f"https://login.xero.com/identity/connect/authorize?"
        f"response_type=code&"
        f"client_id={clientId}&"
        f"redirect_uri={redirectURI}&"
        f"scope=accounting.reports.read offline_access&"
        # f"scope=finance.statements.read offline_access&"
        f"state=123"
    )

    credentials = xeroCredentials.find_one({})
    refresh_expires = credentials["refresh_expires"]
    # if refresh_expires is a string like so 2023-11-19 11:58:10, then convert to a datetime object
    if isinstance(refresh_expires, str):
        refresh_expires = datetime.strptime(refresh_expires, "%Y-%m-%d %H:%M:%S")

    # is refresh_expires in the future?
    if refresh_expires > datetime.now():
        return {"response": authorization_url, "refresh_valid": True}
    else:
        return {"response": authorization_url, "refresh_valid": False}


@xero.get("/callback")
async def xero_callback(request: Request, code: str):
    # Extract the authorization code from the query parameter
    code = code.split("&")[0].split("=")[-1]

    # Create the base64-encoded token for Basic Authentication
    token = b64encode(f"{clientId}:{xerosecret}".encode()).decode()

    # Prepare the request data
    data = {
        "grant_type": "authorization_code",
        "code": code,
        "redirect_uri": redirectURI,
    }

    # Make a POST request to exchange the authorization code for an access token
    async with httpx.AsyncClient() as client:
        response = await client.post(
            "https://identity.xero.com/connect/token",
            headers={
                "Content-Type": "application/x-www-form-urlencoded",
                "Authorization": f"Basic {token}",
            },
            data=urlencode(data),
        )

        response.raise_for_status()

        token_data = response.json()

        # Extract the access token and other data
        access_token = token_data["access_token"]
        expires_in = token_data["expires_in"]
        refresh_token = token_data["refresh_token"]

        # Make a GET request to the connections endpoint
        async with httpx.AsyncClient() as client:
            response = await client.get(
                "https://api.xero.com/connections",
                headers={
                    "Content-Type": "application/json",
                    "Authorization": f"Bearer {access_token}",
                },
            )

            response.raise_for_status()

            tenant_id = response.json()

        # refresh_expires in now +  60 days, convert this and create a variable with it
        refresh_expires = datetime.now() + timedelta(days=60)

        insert = {
            "access_token": access_token,
            "expires_in": expires_in,
            "refresh_token": refresh_token,
            "refresh_expires": refresh_expires,
        }

        # delete all documents in the collection
        xeroCredentials.delete_many({})
        # insert the new document
        xeroCredentials.insert_one(insert)
        # delete all the documents in xeroTenants
        xeroTenants.delete_many({})
        # insert the tenant id into xeroTenants
        xeroTenants.insert_many(tenant_id)

    # return {"authorized": True}
    return RedirectResponse(url="http://localhost:8080/finance")


# @xero.post("/get_profit_and_loss")
# async def get_profit_and_loss(data: Request):
#     request = await data.json()
#
#     year = request['from_date'].split("-")[0]
#     month = request['from_date'].split("-")[1]
#
#     try:
#         start_time = time.time()
#
#         if int(month) < 10:
#             current_year = int(year)
#         else:
#             current_year = int(year) - 1
#
#         periods_to_report = [
#             {
#                 "period": 1,
#                 "from_date": f"{current_year}-03-01",
#                 "to_date": f"{current_year}-03-31"
#             },
#             {
#                 "period": 2,
#                 "from_date": f"{current_year}-04-01",
#                 "to_date": f"{current_year}-04-30"
#             },
#             {
#                 "period": 3,
#                 "from_date": f"{current_year}-05-01",
#                 "to_date": f"{current_year}-05-31"
#             },
#             {
#                 "period": 4,
#                 "from_date": f"{current_year}-06-01",
#                 "to_date": f"{current_year}-06-30"
#             },
#             {
#                 "period": 5,
#                 "from_date": f"{current_year}-07-01",
#                 "to_date": f"{current_year}-07-31"
#             },
#             {
#                 "period": 6,
#                 "from_date": f"{current_year}-08-01",
#                 "to_date": f"{current_year}-08-31"
#             },
#             {
#                 "period": 7,
#                 "from_date": f"{current_year}-09-01",
#                 "to_date": f"{current_year}-09-30"
#             },
#             {
#                 "period": 8,
#                 "from_date": f"{current_year}-10-01",
#                 "to_date": f"{current_year}-10-31"
#             },
#             {
#                 "period": 9,
#                 "from_date": f"{current_year}-11-01",
#                 "to_date": f"{current_year}-11-30"
#             },
#             {
#                 "period": 10,
#                 "from_date": f"{current_year}-12-01",
#                 "to_date": f"{current_year}-12-31"
#             },
#             {
#                 "period": 11,
#                 "from_date": f"{current_year + 1}-01-01",
#                 "to_date": f"{current_year + 1}-01-31"
#             },
#             {
#                 "period": 12,
#                 "from_date": f"{current_year + 1}-02-01",
#                 "to_date": f"{current_year + 1}-02-29"
#             }
#         ]
#
#         data_from_xero = []
#         data_from_xero_HF_PandL = []
#         data_from_xero_HV_PandL = []
#
#         base_data = cpc_data_fields.base_data.copy()
#
#         for record in base_data:
#             for i in range(request['period']):
#                 record['Amount'][i] = 0.0
#
#         base_data_HF_PandL = cpc_data_fields.base_data_HF_PandL.copy()
#
#         for record in base_data_HF_PandL:
#             for i in range(request['period']):
#                 record['Amount'][i] = 0.0
#
#         base_data_HV_PandL = cpc_data_fields.base_data_HV_PandL.copy()
#
#         for record in base_data_HV_PandL:
#             for i in range(0, request['period']):
#                 # print(request['period'])
#                 record['Amount'][i] = 0.0
#
#         credentials_string = f"{clientId}:{xerosecret}".encode("utf-8")
#
#         # Encode the combined string in base64
#         encoded_credentials = b64encode(credentials_string).decode("utf-8")
#
#         # get data from xeroCredentials and save to variable credentials
#         credentials = xeroCredentials.find_one({})
#         id = str(credentials["_id"])
#         del credentials["_id"]
#
#         access_token = credentials["access_token"]
#         refresh_token = credentials["refresh_token"]
#
#         if isinstance(credentials["refresh_expires"], str):
#             credentials["refresh_expires"] = datetime.strptime(credentials["refresh_expires"], "%Y-%m-%d %H:%M:%S")
#         if credentials["refresh_expires"] > datetime.now():
#
#             url_refresh = "https://identity.xero.com/connect/token"
#
#             # Define the headers
#             headers = {
#                 "Content-Type": "application/x-www-form-urlencoded",
#                 "Authorization": f"Basic {encoded_credentials}",  # Replace token with your base64-encoded credentials
#             }
#
#             # Define the form data
#             data = {
#                 "grant_type": "refresh_token",
#                 "refresh_token": refresh_token,
#             }
#
#             # Send the POST request
#             response = requests.post(url_refresh, headers=headers, data=data)
#
#             # Check the response
#             if response.status_code == 200:
#                 # Request was successful, you can access response data like JSON content
#                 response_data = response.json()
#                 access_token = response_data.get("access_token")
#                 expires_in = response_data.get("expires_in")
#                 refresh_token = response_data.get("refresh_token")
#                 refresh_expires = datetime.now() + timedelta(days=60)
#                 insert = {
#                     "access_token": access_token,
#                     "expires_in": expires_in,
#                     "refresh_token": refresh_token,
#                     "refresh_expires": refresh_expires,
#                 }
#                 # update the document in the database
#                 xeroCredentials.update_one({"_id": ObjectId(id)}, {"$set": insert})
#
#             else:
#                 # Request failed, handle the error
#                 print(f"Error: {response.status_code} - {response.text}")
#
#         else:
#             print("refresh_expires is in the past")
#
#         tenant_id = "30b5d5a0-cf38-4bdb-baa1-9dda35b278a2"
#         tenant_id_HF = "9c4ba92b-93b0-4358-9ff8-141aa0718242"
#
#         tenant_id_HV = "4af624e3-6de5-4cc7-9123-36d63d2acbb4"
#
#         tenants = [tenant_id, tenant_id_HF, tenant_id_HV]
#
#         tenants_tb = [tenant_id_HF, tenant_id_HV]
#
#         hv_tb = []
#         hf_tb = []
#
#         for index, tenant in enumerate(tenants_tb):
#             report_date = request['to_date']
#
#             try:
#                 async with httpx.AsyncClient() as client:
#                     response = await client.get(
#                         f"https://api.xero.com/api.xro/2.0/Reports/TrialBalance?date={report_date}",
#                         headers={
#                             "Content-Type": "application/xml",
#                             "Authorization": f"Bearer {access_token}",
#                             "xero-tenant-id": tenant,
#                         },
#                     )
#             except httpx.HTTPError as exc:
#                 print(f"HTTP Error: {exc}")
#                 return {"Success": False, "Error": "HTTP Error"}
#
#             # Check the HTTP status code
#             if response.status_code != 200:
#                 return {"Success": False, "Error": "Non-200 Status Code"}
#
#             # print("Hello")
#
#             python_dict = xmltodict.parse(response.text)
#
#             python_dict = python_dict['Response']['Reports']['Report']['Rows']['Row']
#
#             python_dict = list(filter(lambda x: x['RowType'] == 'Section', python_dict))
#
#             for item in python_dict:
#                 new_item = item['Rows']['Row']
#                 for cells in new_item:
#
#                     if isinstance(cells, dict):
#                         insert = {}
#
#                         for new_index, cell in enumerate(cells['Cells']['Cell']):
#
#                             # put all the above as key value pairs in a dictionary
#                             if new_index == 0:
#                                 clean_account = cell.get('Value', 0)
#                                 if clean_account != 0:
#                                     clean_account_list = clean_account.split("(")
#                                     clean_account_list[0] = clean_account_list[0].strip()
#                                     clean_account_list[1] = clean_account_list[1].replace(")", "")
#                                     clean_account_list[1] = clean_account_list[1].strip()
#
#                                 insert['Account Code'] = clean_account_list[1]
#                                 insert['Account'] = clean_account_list[0]
#                                 insert['Account Type'] = ''
#
#                             if new_index == 3:
#                                 insert['debit_ytd'] = float(cell.get('Value', 0))
#                             if new_index == 4:
#                                 insert['credit_ytd'] = float(cell.get('Value', 0))
#
#                         if index == 0:
#                             hf_tb.append(insert)
#                         elif index == 1:
#                             hv_tb.append(insert)
#
#         short_months = [2, 4, 6, 8, 12]
#         comparison_data_cpc = []
#         comparison_data_hf = []
#         comparison_data_hv = []
#
#         # Make a GET request to the connections endpoint
#         for index1, tenant in enumerate(tenants):
#
#             if request['period'] in short_months:
#
#                 for period in range(request['period'], request['period'] + 1):
#
#                     filtered_xero = list(filter(lambda x: x['period'] == period, periods_to_report))
#
#                     report_from_date = filtered_xero[0]['from_date']
#                     report_to_date = filtered_xero[0]['to_date']
#                     period = filtered_xero[0]['period']
#                     try:
#                         async with httpx.AsyncClient() as client:
#                             response = await client.get(
#                                 f"https://api.xero.com/api.xro/2.0/Reports/ProfitAndLoss?fromDate={report_from_date}&toDate="
#                                 f"{report_to_date}",
#                                 headers={
#                                     "Content-Type": "application/xml",
#                                     "Authorization": f"Bearer {access_token}",
#                                     "xero-tenant-id": tenant,
#                                 },
#                             )
#                     except httpx.HTTPError as exc:
#                         print(f"HTTP Error: {exc}")
#                         return {"Success": False, "Error": "HTTP Error"}
#
#                     # Check the HTTP status code
#                     if response.status_code != 200:
#                         # print("HTTP Status Code:", response.status_code)
#                         return {"Success": False, "Error": "Non-200 Status Code"}
#
#                     python_dict = xmltodict.parse(response.text)
#
#                     python_dict = python_dict['Response']['Reports']['Report']['Rows']['Row']
#
#                     for item in python_dict:
#
#                         if 'Rows' in item:
#                             if isinstance(item['Rows']['Row'], list):
#                                 data = item['Rows']['Row']
#                                 for row in data:
#                                     if row['RowType'] == 'Row':
#
#                                         final_line_data = row['Cells']['Cell']
#                                         insert = {}
#                                         for index, line_item in enumerate(final_line_data):
#
#                                             try:
#                                                 insert[index] = float(line_item['Value'])
#                                             except:
#                                                 insert[index] = line_item['Value']
#                                         insert['period'] = period
#                                         if index1 == 0:
#                                             data_from_xero.append(insert)
#
#                                         elif index1 == 1:
#                                             data_from_xero_HF_PandL.append(insert)
#
#                                         elif index1 == 2:
#                                             data_from_xero_HV_PandL.append(insert)
#
#                 period = request['period'] - 1
#
#                 filtered_xero = list(filter(lambda x: x['period'] == period, periods_to_report))
#
#                 report_from_date = filtered_xero[0]['from_date']
#                 report_to_date = filtered_xero[0]['to_date']
#                 period = filtered_xero[0]['period']
#
#                 try:
#                     async with httpx.AsyncClient() as client:
#                         response = await client.get(
#                             f"https://api.xero.com/api.xro/2.0/Reports/ProfitAndLoss?fromDate={report_from_date}&"
#                             f"toDate={report_to_date}&periods={period - 1}&timeframe=MONTH",
#                             headers={
#                                 "Content-Type": "application/xml",
#                                 "Authorization": f"Bearer {access_token}",
#                                 "xero-tenant-id": tenant,
#                             },
#                         )
#                 except httpx.HTTPError as exc:
#                     print(f"HTTP Error: {exc}")
#                     return {"Success": False, "Error": "HTTP Error"}
#
#                 # Check the HTTP status code
#                 if response.status_code != 200:
#                     return {"Success": False, "Error": "Non-200 Status Code"}
#
#                 python_dict = xmltodict.parse(response.text)
#
#                 python_dict = python_dict['Response']['Reports']['Report']['Rows']['Row']
#
#                 python_dict = list(filter(lambda x: x['RowType'] == 'Section', python_dict))
#
#                 for item in python_dict:
#
#                     if 'Rows' in item:
#                         if isinstance(item['Rows']['Row'], list):
#                             data = item['Rows']['Row']
#                             for row in data:
#                                 # print("row", row)
#                                 if row['RowType'] == 'Row':
#
#                                     insert = {}
#                                     values = []
#                                     if isinstance(row['Cells']['Cell'], list):
#                                         newData = row['Cells']['Cell']
#                                         for index, line_item in enumerate(newData):
#
#                                             if index == 0:
#                                                 insert['Account'] = line_item['Value']
#                                             else:
#                                                 values.append(line_item['Value'])
#                                     insert['Amount'] = values
#                                     if index1 == 0:
#                                         comparison_data_cpc.append(insert)
#
#                                     elif index1 == 1:
#                                         comparison_data_hf.append(insert)
#
#                                     elif index1 == 2:
#                                         comparison_data_hv.append(insert)
#
#
#             else:
#
#                 period = request['period']
#
#                 filtered_xero = list(filter(lambda x: x['period'] == period, periods_to_report))
#
#                 report_from_date = filtered_xero[0]['from_date']
#                 report_to_date = filtered_xero[0]['to_date']
#                 period = filtered_xero[0]['period']
#
#                 try:
#                     async with httpx.AsyncClient() as client:
#                         response = await client.get(
#                             f"https://api.xero.com/api.xro/2.0/Reports/ProfitAndLoss?fromDate={report_from_date}&toDate={report_to_date}&periods={period - 1}&timeframe=MONTH",
#                             headers={
#                                 "Content-Type": "application/xml",
#                                 "Authorization": f"Bearer {access_token}",
#                                 "xero-tenant-id": tenant,
#                             },
#                         )
#                 except httpx.HTTPError as exc:
#                     print(f"HTTP Error: {exc}")
#                     return {"Success": False, "Error": "HTTP Error"}
#
#                 # Check the HTTP status code
#                 if response.status_code != 200:
#                     return {"Success": False, "Error": "Non-200 Status Code"}
#
#                 python_dict = xmltodict.parse(response.text)
#
#                 python_dict = python_dict['Response']['Reports']['Report']['Rows']['Row']
#
#                 python_dict = list(filter(lambda x: x['RowType'] == 'Section', python_dict))
#
#                 for item in python_dict:
#
#                     if 'Rows' in item:
#                         if isinstance(item['Rows']['Row'], list):
#                             data = item['Rows']['Row']
#                             for row in data:
#
#                                 if row['RowType'] == 'Row':
#
#                                     insert = {}
#                                     values = []
#                                     if isinstance(row['Cells']['Cell'], list):
#                                         newData = row['Cells']['Cell']
#                                         for index, line_item in enumerate(newData):
#
#                                             if index == 0:
#                                                 insert['Account'] = line_item['Value']
#                                             else:
#                                                 values.append(line_item['Value'])
#                                     insert['Amount'] = values
#                                     if index1 == 0:
#                                         comparison_data_cpc.append(insert)
#
#                                     elif index1 == 1:
#                                         comparison_data_hf.append(insert)
#
#                                     elif index1 == 2:
#                                         comparison_data_hv.append(insert)
#                                         print("comparison_data_hv", comparison_data_hv)
#
#         for item in comparison_data_cpc:
#             item['Amount'].reverse()
#             insert = {}
#             for idx, value in enumerate(item['Amount'], start=1):
#                 insert[0] = item['Account']
#                 insert[1] = float(value)
#                 insert['period'] = idx
#
#                 data_from_xero.append(insert)
#                 insert = {}
#
#         for item in comparison_data_hf:
#             item['Amount'].reverse()
#             insert = {}
#             for idx, value in enumerate(item['Amount'], start=1):
#                 insert[0] = item['Account']
#                 insert[1] = float(value)
#                 insert['period'] = idx
#
#                 data_from_xero_HF_PandL.append(insert)
#                 insert = {}
#
#         # print("comparison_data_hv", comparison_data_hv)
#         for item in comparison_data_hv:
#             item['Amount'].reverse()
#             insert = {}
#             for idx, value in enumerate(item['Amount'], start=1):
#                 insert[0] = item['Account']
#                 insert[1] = float(value)
#                 insert['period'] = idx
#
#                 data_from_xero_HV_PandL.append(insert)
#                 insert = {}
#
#         data_from_xero.sort(key=lambda x: (x[0], x['period']))
#
#         data_from_xero_HF_PandL.sort(key=lambda x: (x[0], x['period']))
#
#         data_from_xero_HV_PandL.sort(key=lambda x: (x[0], x['period']))
#
#         for final_record in base_data:
#
#             filtered_data_from_xero = list(filter(lambda x: x[0] == final_record['Account'], data_from_xero))
#
#             for record in filtered_data_from_xero:
#                 final_record['Amount'][record['period'] - 1] = record[1]
#
#             final_record['Amount'].reverse()
#
#         for final_record in base_data_HF_PandL:
#
#             filtered_data_from_xero = list(filter(lambda x: x[0] == final_record['Account'], data_from_xero_HF_PandL))
#
#             for record in filtered_data_from_xero:
#                 final_record['Amount'][record['period'] - 1] = record[1]
#
#             final_record['Amount'].reverse()
#
#         for record in base_data_HF_PandL:
#             if record['Account'] == 'Accounting Fees':
#
#                 filtered_data_from_xero = list(
#                     filter(lambda x: x['Account'] == 'Accounting - CIPC', base_data_HF_PandL))
#
#                 for index, item in enumerate(record['Amount']):
#                     record['Amount'][index] = record['Amount'][index] + filtered_data_from_xero[0]['Amount'][index]
#
#             if record['Account'] == 'Staff welfare':
#
#                 filtered_data_from_xero1 = list(
#                     filter(lambda x: x['Account'] == 'Entertainment Expenses', base_data_HF_PandL))
#
#                 filtered_data_from_xero2 = list(
#                     filter(lambda x: x['Account'] == 'General Expenses', base_data_HF_PandL))
#
#                 for index, item in enumerate(record['Amount']):
#                     record['Amount'][index] = filtered_data_from_xero1[0]['Amount'][index] + \
#                                               filtered_data_from_xero2[0]['Amount'][index]
#
#             if record['Account'] == 'Repairs _AND_ Maintenance':
#                 # print("record", record)
#                 filtered_data_from_xero = list(
#                     filter(lambda x: x['Account'] == 'Motor Vehicle Expenses', base_data_HF_PandL))
#
#                 for index, item in enumerate(record['Amount']):
#                     record['Amount'][index] = record['Amount'][index] + filtered_data_from_xero[0]['Amount'][index]
#
#         for final_record in base_data_HV_PandL:
#
#             filtered_data_from_xero = list(filter(lambda x: x[0] == final_record['Account'], data_from_xero_HV_PandL))
#
#             for record in filtered_data_from_xero:
#                 final_record['Amount'][record['period'] - 1] = record[1]
#
#             final_record['Amount'].reverse()
#
#         # sort data_from_xero_HF_PandL first by '0' then by 'period'
#         data_from_xero_HF_PandL.sort(key=lambda x: (x[0], x['period']))
#         data_from_xero_HV_PandL.sort(key=lambda x: (x[0], x['period']))
#
#         returned_data = insert_data_from_xero_profit_loss(base_data, base_data_HF_PandL, base_data_HV_PandL, hf_tb,
#                                                           hv_tb, request['to_date'])
#
#         end_time = time.time()
#
#         if returned_data['Success']:
#             return {"Success": True, "time_taken": end_time - start_time}
#         else:
#             return {"Success": False, "Error": returned_data['Error']}
#
#
#
#     except Exception as e:
#         print("Error", e)
#         return {"Success": False, "Error": str(e)}
#
#
# @xero.get("/get_profit_and_loss")
# async def get_profit_and_loss(file_name):
#     try:
#         file_name = file_name + ".xlsx"
#         file_name = file_name.replace("_", " ")
#         file_name = file_name.replace("xx", "&")
#
#         dir_path = "cashflow_p&l_files"
#         dir_list = os.listdir(dir_path)
#
#         if file_name in dir_list:
#
#             return FileResponse(f"{dir_path}/{file_name}", filename=file_name)
#         else:
#             return {"ERROR": "File does not exist!!"}
#     except Exception as e:
#         print(e)
#         return {"ERROR": "Please Try again"}


@xero.post("/get_bank_transactions")
async def get_bank_transactions(data: Request):
    request = await data.json()
    print("Request", request)

    # CONNECT TO XERO
    credentials_string = f"{clientId}:{xerosecret}".encode("utf-8")

    # Encode the combined string in base64
    encoded_credentials = b64encode(credentials_string).decode("utf-8")

    # get data from xeroCredentials and save to variable credentials
    credentials = xeroCredentials.find_one({})
    id = str(credentials["_id"])
    del credentials["_id"]

    access_token = credentials["access_token"]
    refresh_token = credentials["refresh_token"]

    if isinstance(credentials["refresh_expires"], str):
        credentials["refresh_expires"] = datetime.strptime(credentials["refresh_expires"], "%Y-%m-%d %H:%M:%S")
    if credentials["refresh_expires"] > datetime.now():

        url_refresh = "https://identity.xero.com/connect/token"

        # Define the headers
        headers = {
            "Content-Type": "application/x-www-form-urlencoded",
            "Authorization": f"Basic {encoded_credentials}",  # Replace token with your base64-encoded credentials
        }

        # Define the form data
        data = {
            "grant_type": "refresh_token",
            "refresh_token": refresh_token,
        }

        # Send the POST request
        response = requests.post(url_refresh, headers=headers, data=data)

        # Check the response
        if response.status_code == 200:
            # Request was successful, you can access response data like JSON content
            response_data = response.json()
            access_token = response_data.get("access_token")
            expires_in = response_data.get("expires_in")
            refresh_token = response_data.get("refresh_token")
            refresh_expires = datetime.now() + timedelta(days=60)
            insert = {
                "access_token": access_token,
                "expires_in": expires_in,
                "refresh_token": refresh_token,
                "refresh_expires": refresh_expires,
            }
            # update the document in the database
            xeroCredentials.update_one({"_id": ObjectId(id)}, {"$set": insert})

        else:
            # Request failed, handle the error
            print(f"Error: {response.status_code} - {response.text}")

    else:
        print("refresh_expires is in the past")

    # get XERO TENANTS
    tenants = list(xeroTenants.find({}))
    # Manipulate request dates
    start_date = request['start_date'].replace("-", ", ")
    end_date = request['end_date'].replace("-", ", ")

    bank_transactions = []

    for index, tenant in enumerate(tenants):
        try:
            async with httpx.AsyncClient() as client:
                response = await client.get(
                    f"https://api.xero.com/api.xro/2.0/BankTransactions?where=Date >= DateTime({start_date}) && Date < DateTime({end_date})",
                    headers={
                        "Content-Type": "application/xml",
                        "Authorization": f"Bearer {access_token}",
                        "xero-tenant-id": tenant['tenantId'],
                    },
                )
        except httpx.HTTPError as exc:
            print(f"HTTP Error: {exc}")
            return {"Success": False, "Error": "HTTP Error"}

        # Check the HTTP status code
        if response.status_code != 200:
            print("HTTP Status Code:", response.status_code)
            return {"Success": False, "Error": "Non-200 Status Code - First"}

        python_dict = xmltodict.parse(response.text)

        if 'BankTransactions' in python_dict['Response']:

            for transaction in python_dict['Response']['BankTransactions']['BankTransaction']:
                print("Transaction", transaction)
                print()

                try:
                    contact = transaction['Contact']['Name']

                except KeyError:
                    contact = ""

                insert = {
                    "Company": tenant['tenantName'],
                    "TenantID": tenant['tenantId'],
                    "Date": transaction['Date'],
                    "Subtotal": transaction.get('SubTotal', 0),
                    "TotalTax": transaction.get('TotalTax', 0),
                    "Amount": transaction.get('Total', 0),
                    "Status": transaction['Status'],
                    "Type": transaction['Type'],
                    "isReconciled": transaction['IsReconciled'],
                    "BankAccount": transaction['BankAccount']['Code'],
                    "Contact": contact,
                    "BankTransactionID": transaction['BankTransactionID'],
                }

                bank_transactions.append(insert)

    interim_bank_transactions = []

    print("Length of bank_transactions", len(bank_transactions))

    no_of_requests = len(bank_transactions) / 50
    # round up to the nearest whole number
    no_of_requests = math.ceil(no_of_requests)

    this_request = 0
    while this_request <= no_of_requests:
        start = this_request * 50
        end = ((this_request + 1) * 50)
        filtered_bank_transactions = bank_transactions[start:end]

        if len(filtered_bank_transactions) > 0:
            for transaction in filtered_bank_transactions:
                try:
                    async with httpx.AsyncClient() as client:
                        transaction_id = transaction['BankTransactionID']
                        tenant_id = transaction['TenantID']
                        response = await client.get(
                            f"https://api.xero.com/api.xro/2.0/BankTransactions/{transaction_id}",
                            headers={
                                "Content-Type": "application/xml",
                                "Authorization": f"Bearer {access_token}",
                                "xero-tenant-id": tenant_id,
                            },
                        )
                except httpx.HTTPError as exc:
                    print(f"HTTP Error: {exc}")
                    return {"Success": False, "Error": "HTTP Error"}

                    # Check the HTTP status code
                if response.status_code != 200:
                    return {"Success": False, "Error": "Non-200 Status Code - Second"}

                python_dict = xmltodict.parse(response.text)

                python_dict['Response']['Company'] = transaction['Company']
                interim_bank_transactions.append(python_dict['Response'])

        this_request += 1
        if len(filtered_bank_transactions) < 50:
            break
        time.sleep(60)

    final_data = []
    for tr in interim_bank_transactions:
        insert_a = {
            "Company": tr['Company'],
            "Type": tr['BankTransactions']['BankTransaction']['Type'],
            "isReconciled": tr['BankTransactions']['BankTransaction']['IsReconciled'],
            "Date": tr['BankTransactions']['BankTransaction']['Date'],
            "Status": tr['BankTransactions']['BankTransaction']['Status'],
            "AccountCode": tr['BankTransactions']['BankTransaction']['LineItems']['LineItem']['AccountCode'],
            "Description": tr['BankTransactions']['BankTransaction']['LineItems']['LineItem'].get('Description',
                                                                                                  "Not Found"),
            "Amount": tr['BankTransactions']['BankTransaction']['Total'],
            "BankAccountCode": tr['BankTransactions']['BankTransaction']['BankAccount']['Code'],
            "BankAccountName": tr['BankTransactions']['BankTransaction']['BankAccount']['Name'],
            "BankTransactionID": tr.get('Id', "Cannot Find")
        }

        final_data.append(insert_a)

    #     return {"Success": True, "length": len(test), "BankTransactions": interim_bank_transactions}
    print("complete!!!")
    return {"Success": True, "length": len(final_data), "Final Data": final_data, "BankTransactions": bank_transactions}
    # return {"Success": True, "length": len(bank_transactions), "BankTransactions": bank_transactions}


@xero.post("/get_bank_payments")
async def get_bank_payments(data: Request):
    request = await data.json()
    print("Request", request)

    # CONNECT TO XERO
    credentials_string = f"{clientId}:{xerosecret}".encode("utf-8")

    # Encode the combined string in base64
    encoded_credentials = b64encode(credentials_string).decode("utf-8")

    # get data from xeroCredentials and save to variable credentials
    credentials = xeroCredentials.find_one({})
    id = str(credentials["_id"])
    del credentials["_id"]

    access_token = credentials["access_token"]
    refresh_token = credentials["refresh_token"]

    if isinstance(credentials["refresh_expires"], str):
        credentials["refresh_expires"] = datetime.strptime(credentials["refresh_expires"], "%Y-%m-%d %H:%M:%S")
    if credentials["refresh_expires"] > datetime.now():

        url_refresh = "https://identity.xero.com/connect/token"

        # Define the headers
        headers = {
            "Content-Type": "application/x-www-form-urlencoded",
            "Authorization": f"Basic {encoded_credentials}",  # Replace token with your base64-encoded credentials
        }

        # Define the form data
        data = {
            "grant_type": "refresh_token",
            "refresh_token": refresh_token,
        }

        # Send the POST request
        response = requests.post(url_refresh, headers=headers, data=data)

        # Check the response
        if response.status_code == 200:
            # Request was successful, you can access response data like JSON content
            response_data = response.json()
            access_token = response_data.get("access_token")
            expires_in = response_data.get("expires_in")
            refresh_token = response_data.get("refresh_token")
            refresh_expires = datetime.now() + timedelta(days=60)
            insert = {
                "access_token": access_token,
                "expires_in": expires_in,
                "refresh_token": refresh_token,
                "refresh_expires": refresh_expires,
            }
            # update the document in the database
            xeroCredentials.update_one({"_id": ObjectId(id)}, {"$set": insert})

        else:
            # Request failed, handle the error
            print(f"Error: {response.status_code} - {response.text}")

    else:
        print("refresh_expires is in the past")

    # get XERO TENANTS
    tenants = list(xeroTenants.find({}))
    # Manipulate request dates
    start_date = request['start_date'].replace("-", ", ")
    end_date = request['end_date'].replace("-", ", ")

    payment_transactions = []

    for index, tenant in enumerate(tenants):
        try:
            async with httpx.AsyncClient() as client:
                response = await client.get(
                    f"https://api.xero.com/api.xro/2.0/Payments?where=Date >= DateTime({start_date}) && Date < DateTime({end_date})",
                    headers={
                        "Content-Type": "application/xml",
                        "Authorization": f"Bearer {access_token}",
                        "xero-tenant-id": tenant['tenantId'],
                    },
                )
        except httpx.HTTPError as exc:
            print(f"HTTP Error: {exc}")
            return {"Success": False, "Error": "HTTP Error"}

        # Check the HTTP status code
        if response.status_code != 200:
            print("HTTP Status Code:", response.status_code)
            return {"Success": False, "Error": "Non-200 Status Code - First"}

        python_dict = xmltodict.parse(response.text)

        # print("Python Dict", python_dict)

        if 'Payments' in python_dict['Response']:
            if 'Payment' in python_dict['Response']['Payments']:
                for payment in python_dict['Response']['Payments']['Payment']:
                    print("Payment", payment)
                    print()
                    print("Type", type(payment))
                    print()
                    # try:
                    if type(payment) == dict:
                        insert = {
                            "PaymentID": payment.get('PaymentID', "None"),
                            "Date": payment['Date'],
                            "BankAmount": payment['BankAmount'],
                            "Company": tenant['tenantName'],
                            "TenantID": tenant['tenantId'],

                        }
                        payment_transactions.append(insert)
                    # except KeyError:
                    #     print("Key Error", payment)
                    # print("Payment Transactions", payment_transactions)
                # print("Payment", payment)
                # print()

    #     if 'BankTransactions' in python_dict['Response']:
    #
    #         for transaction in python_dict['Response']['BankTransactions']['BankTransaction']:
    #             print("Transaction", transaction)
    #             print()
    #
    #             try:
    #                 contact = transaction['Contact']['Name']
    #
    #             except KeyError:
    #                 contact = ""
    #
    #             insert = {
    #                 "Company": tenant['tenantName'],
    #                 "TenantID": tenant['tenantId'],
    #                 "Date": transaction['Date'],
    #                 "Subtotal": transaction.get('SubTotal', 0),
    #                 "TotalTax": transaction.get('TotalTax', 0),
    #                 "Amount": transaction.get('Total', 0),
    #                 "Status": transaction['Status'],
    #                 "Type": transaction['Type'],
    #                 "isReconciled": transaction['IsReconciled'],
    #                 "BankAccount": transaction['BankAccount']['Code'],
    #                 "Contact": contact,
    #                 "BankTransactionID": transaction['BankTransactionID'],
    #             }
    #
    #             bank_transactions.append(insert)
    #
    interim_bank_transactions = []

    print("Length of bank_transactions", len(payment_transactions))

    no_of_requests = len(payment_transactions) / 50
    # round up to the nearest whole number
    no_of_requests = math.ceil(no_of_requests)

    this_request = 0
    while this_request <= no_of_requests:
        start = this_request * 50
        end = ((this_request + 1) * 50)
        filtered_payment_transactions = payment_transactions[start:end]

        if len(filtered_payment_transactions) > 0:
            for transaction in filtered_payment_transactions:
                try:
                    async with httpx.AsyncClient() as client:
                        transaction_id = transaction['PaymentID']
                        tenant_id = transaction['TenantID']
                        response = await client.get(
                            f"https://api.xero.com/api.xro/2.0/Payments/{transaction_id}",
                            headers={
                                "Content-Type": "application/xml",
                                "Authorization": f"Bearer {access_token}",
                                "xero-tenant-id": tenant_id,
                            },
                        )
                except httpx.HTTPError as exc:
                    print(f"HTTP Error: {exc}")
                    return {"Success": False, "Error": "HTTP Error"}

                    # Check the HTTP status code
                if response.status_code != 200:
                    return {"Success": False, "Error": "Non-200 Status Code - Second"}

                python_dict = xmltodict.parse(response.text)

                # print("Python Dict", python_dict)

                python_dict['Response']['Company'] = transaction['Company']
                interim_bank_transactions.append(python_dict['Response'])

        this_request += 1
        if len(filtered_payment_transactions) < 50:
            break
        time.sleep(60)

    final_data = []
    for tr in interim_bank_transactions:
        # print(type(tr['Payments']['Payment']['Invoice']['LineItems']['LineItem']))
        if type(tr['Payments']['Payment']['Invoice']['LineItems']['LineItem']) == list:
            paid_account = tr['Payments']['Payment']['Invoice']['LineItems']['LineItem'][0].get('AccountCode',
                                                                                                "Unknown")
        else:
            paid_account = tr['Payments']['Payment']['Invoice']['LineItems']['LineItem']['AccountCode']

        insert_a = {
            "PaymentID": tr['Payments']['Payment']['PaymentID'],
            "Date": tr['Payments']['Payment']['Date'],
            "Company": tr['Company'],
            "Amount": tr['Payments']['Payment']['Amount'],
            "Reference": tr['Payments']['Payment'].get('Reference', "None"),
            "IsReconciled": tr['Payments']['Payment']['IsReconciled'],
            "AccountCode": tr['Payments']['Payment']['Account']['Code'],
            "AccountName": tr['Payments']['Payment']['Account']['Name'],
            "paid_account": paid_account,
        }

        final_data.append(insert_a)

    #     return {"Success": True, "length": len(test), "BankTransactions": interim_bank_transactions}
    print("complete!!!")
    return {"Success": True, "Length": len(final_data), "PaymentTransactions": final_data}
    # return {"Success": True, "length": len(bank_transactions), "BankTransactions": bank_transactions}


@xero.post("/process_profit_and_loss")
async def process_profit_and_loss(data: Request):
    global count_received, count_received4, count_received3, count_received2, count_received1, count_exited4, \
        count_exited3, count_exited2, count_exited1, count_units4, count_units3, count_units2, count_units1, \
        transferred_units4, transferred_units3, transferred_units2, transferred_units1, sold_units4, sold_units3, \
        sold_units2, sold_units1, remaining_units4, remaining_units3, remaining_units2, remaining_units1, \
        transferred_units_sold_value4, transferred_units_sold_value3, transferred_units_sold_value2, \
        transferred_units_sold_value1, value_exited4, current_investor_capital_deployed4, value_still_to_be_raised4, \
        value_exited3, current_investor_capital_deployed3, value_still_to_be_raised3, value_exited2, \
        current_investor_capital_deployed2, value_exited1, current_investor_capital_deployed1, \
        value_still_to_be_raised1, value_still_to_be_raised2, total_units_sales_value4, total_units_sales_value3, \
        total_units_sales_value2, total_units_sales_value1, remaining_units_value4, remaining_units_value3, \
        remaining_units_value2, remaining_units_value1, sold_units_value4, sold_units_value3, sold_units_value2, \
        sold_units_value1, remaining_units_commission4, remaining_units_commission3, remaining_units_commission2, \
        remaining_units_commission1, total_units_commission4, transferred_units_commission4, sold_units_commission4, \
        total_units_commission3, transferred_units_commission3, sold_units_commission3, total_units_commission2, \
        transferred_units_commission2, sold_units_commission2, total_units_commission1, transferred_units_commission1, \
        sold_units_commission1, sold_units_transfer_fees4, sold_units_transfer_fees3, sold_units_transfer_fees2, \
        sold_units_transfer_fees1, total_units_transfer_fees4, remaining_units_transfer_fees4, \
        total_units_transfer_fees3, total_units_transfer_fees2, total_units_transfer_fees2, \
        remaining_units_transfer_fees2, remaining_units_transfer_fees2, total_units_transfer_fees1, \
        remaining_units_transfer_fees1, transferred_units_transfer_fees4, transferred_units_transfer_fees3, \
        transferred_units_transfer_fees2, transferred_units_transfer_fees1, remaining_units_transfer_fees3, \
        sold_units_bond_registration4, sold_units_bond_registration3, sold_units_bond_registration2, \
        sold_units_bond_registration1, total_units_bond_registration4, remaining_units_bond_registration4, \
        total_units_bond_registration3, remaining_units_bond_registration3, total_units_bond_registration2, \
        remaining_units_bond_registration2, total_units_bond_registration1, remaining_units_bond_registration1, \
        transferred_units_bond_registration4, transferred_units_bond_registration3, \
        transferred_units_bond_registration2, transferred_units_bond_registration1, sold_units_trust_release_fee4, \
        sold_units_trust_release_fee3, sold_units_trust_release_fee2, sold_units_trust_release_fee1, \
        total_units_trust_release_fee4, remaining_units_trust_release_fee4, total_units_trust_release_fee3, \
        remaining_units_trust_release_fee3, total_units_trust_release_fee2, remaining_units_trust_release_fee2, \
        total_units_trust_release_fee1, remaining_units_trust_release_fee1, transferred_units_trust_release_fee4, \
        transferred_units_trust_release_fee3, transferred_units_trust_release_fee2, \
        transferred_units_trust_release_fee1, remaining_units_unforseen4, remaining_units_unforseen3, \
        remaining_units_unforseen2, remaining_units_unforseen1, total_units_unforseen4, total_units_unforseen3, \
        total_units_unforseen2, total_units_unforseen1, transferred_units_unforseen4, transferred_units_unforseen3, \
        transferred_units_unforseen2, transferred_units_unforseen1, sold_units_unforseen4, sold_units_unforseen3, \
        sold_units_unforseen2, sold_units_unforseen1, interest_paid_to_date4, interest_paid_to_date3, \
        interest_paid_to_date2, interest_paid_to_date1
    request = await data.json()

    # for index, item in enumerate(cpc_data_fields.base_data,1):
    #     # make all the values in item['Amount'] list equal 0
    #     item['Amount'] = [0.00] * 12
    #
    # for index, item in enumerate(cpc_data_fields.base_data_HF_PandL,1):
    #     # make all the values in item['Amount'] list equal 0
    #     item['Amount'] = [0.00] * 12
    #
    # for index, item in enumerate(cpc_data_fields.base_data_HV_PandL,1):
    #     # make all the values in item['Amount'] list equal 0
    #     item['Amount'] = [0.00] * 12

    # Was 12
    for i in range(0, 12):
        for data in cpc_data_fields.base_data:
            data["Amount"].append(0.00)
        for data in cpc_data_fields.base_data_HF_PandL:
            data["Amount"].append(0.00)
        for data in cpc_data_fields.base_data_HV_PandL:
            data["Amount"].append(0.00)

    # print(cpc_data_fields.base_data)

    print("2nd Print: ", cpc_data_fields.base_data_HV_PandL[len(cpc_data_fields.base_data_HV_PandL) - 2]['Amount'])

    year = request['from_date'].split("-")[0]
    month = request['from_date'].split("-")[1]
    # print("Request", request)
    # print("month", month)

    # try:
    # remove cashflow_p&l_files/cashflow_hf_hv.xlsx if it exists
    if os.path.exists("cashflow_p&l_files/cashflow_hf_hv.xlsx"):
        os.remove("cashflow_p&l_files/cashflow_hf_hv.xlsx")

    start_time = time.time()
    if int(month) - 2 < 11:
        current_year = int(year)
        short_year = int(str(current_year)[2:4])

    else:
        current_year = int(year) - 1
        short_year = int(str(current_year)[2:4])

    periods_to_report = [
        {
            "period": 1,
            "from_date": f"{current_year}-03-01",
            "to_date": f"{current_year}-03-31",
            "Month": f"Mar-{short_year}"
        },
        {
            "period": 2,
            "from_date": f"{current_year}-04-01",
            "to_date": f"{current_year}-04-30",
            "Month": f"Apr-{short_year}"
        },
        {
            "period": 3,
            "from_date": f"{current_year}-05-01",
            "to_date": f"{current_year}-05-31",
            "Month": f"May-{short_year}"
        },
        {
            "period": 4,
            "from_date": f"{current_year}-06-01",
            "to_date": f"{current_year}-06-30",
            "Month": f"Jun-{short_year}"
        },
        {
            "period": 5,
            "from_date": f"{current_year}-07-01",
            "to_date": f"{current_year}-07-31",
            "Month": f"Jul-{short_year}"

        },
        {
            "period": 6,
            "from_date": f"{current_year}-08-01",
            "to_date": f"{current_year}-08-31",
            "Month": f"Aug-{short_year}"
        },
        {
            "period": 7,
            "from_date": f"{current_year}-09-01",
            "to_date": f"{current_year}-09-30",
            "Month": f"Sep-{short_year}"
        },
        {
            "period": 8,
            "from_date": f"{current_year}-10-01",
            "to_date": f"{current_year}-10-31",
            "Month": f"Oct-{short_year}"
        },
        {
            "period": 9,
            "from_date": f"{current_year}-11-01",
            "to_date": f"{current_year}-11-30",
            "Month": f"Nov-{short_year}"
        },
        {
            "period": 10,
            "from_date": f"{current_year}-12-01",
            "to_date": f"{current_year}-12-31",
            "Month": f"Dec-{short_year}"
        },
        {
            "period": 11,
            "from_date": f"{current_year + 1}-01-01",
            "to_date": f"{current_year + 1}-01-31",
            "Month": f"Jan-{short_year + 1}"
        },
        {
            "period": 12,
            "from_date": f"{current_year + 1}-02-01",
            "to_date": f"{current_year + 1}-02-29",
            "Month": f"Feb-{short_year + 1}"
        }
    ]

    credentials_string = f"{clientId}:{xerosecret}".encode("utf-8")

    # Encode the combined string in base64
    encoded_credentials = b64encode(credentials_string).decode("utf-8")

    # get data from xeroCredentials and save to variable credentials
    credentials = xeroCredentials.find_one({})
    id = str(credentials["_id"])
    del credentials["_id"]

    access_token = credentials["access_token"]
    refresh_token = credentials["refresh_token"]

    if isinstance(credentials["refresh_expires"], str):
        credentials["refresh_expires"] = datetime.strptime(credentials["refresh_expires"], "%Y-%m-%d %H:%M:%S")
    if credentials["refresh_expires"] > datetime.now():

        url_refresh = "https://identity.xero.com/connect/token"

        # Define the headers
        headers = {
            "Content-Type": "application/x-www-form-urlencoded",
            "Authorization": f"Basic {encoded_credentials}",  # Replace token with your base64-encoded credentials
        }

        # Define the form data
        data = {
            "grant_type": "refresh_token",
            "refresh_token": refresh_token,
        }

        # Send the POST request
        response = requests.post(url_refresh, headers=headers, data=data)

        # Check the response
        if response.status_code == 200:
            # Request was successful, you can access response data like JSON content
            response_data = response.json()
            access_token = response_data.get("access_token")
            expires_in = response_data.get("expires_in")
            refresh_token = response_data.get("refresh_token")
            refresh_expires = datetime.now() + timedelta(days=60)
            insert = {
                "access_token": access_token,
                "expires_in": expires_in,
                "refresh_token": refresh_token,
                "refresh_expires": refresh_expires,
            }
            # update the document in the database
            xeroCredentials.update_one({"_id": ObjectId(id)}, {"$set": insert})

            # print("SUCCESS")

        else:
            # Request failed, handle the error
            print(f"Error: {response.status_code} - {response.text}")

    else:
        print("refresh_expires is in the past")

    tenant_id = "30b5d5a0-cf38-4bdb-baa1-9dda35b278a2"
    tenant_id_HF = "9c4ba92b-93b0-4358-9ff8-141aa0718242"

    tenant_id_HV = "4af624e3-6de5-4cc7-9123-36d63d2acbb4"

    tenants = [tenant_id, tenant_id_HF, tenant_id_HV]

    data_from_xero = []
    data_from_xero_HF_PandL = []
    data_from_xero_HV_PandL = []

    short_months = [2, 4, 7, 9, 12]
    comparison_data_cpc = []
    comparison_data_hf = []
    comparison_data_hv = []

    for index1, tenant in enumerate(tenants):

        if request['period'] in short_months:

            for period in range(request['period'], request['period'] + 1):

                filtered_xero = list(filter(lambda x: x['period'] == period, periods_to_report))

                report_from_date = filtered_xero[0]['from_date']
                report_to_date = filtered_xero[0]['to_date']
                period = filtered_xero[0]['period']
                try:
                    async with httpx.AsyncClient() as client:
                        response = await client.get(
                            f"https://api.xero.com/api.xro/2.0/Reports/ProfitAndLoss?fromDate={report_from_date}&toDate="
                            f"{report_to_date}",
                            headers={
                                "Content-Type": "application/xml",
                                "Authorization": f"Bearer {access_token}",
                                "xero-tenant-id": tenant,
                            },
                        )
                except httpx.HTTPError as exc:
                    print(f"HTTP Error: {exc}")
                    return {"Success": False, "Error": "HTTP Error"}

                # Check the HTTP status code
                if response.status_code != 200:
                    # print("HTTP Status Code:", response.status_code)
                    return {"Success": False, "Error": "Non-200 Status Code"}

                python_dict = xmltodict.parse(response.text)

                python_dict = python_dict['Response']['Reports']['Report']['Rows']['Row']

                for item in python_dict:

                    if 'Rows' in item:
                        if isinstance(item['Rows']['Row'], list):
                            data = item['Rows']['Row']
                            for row in data:
                                if row['RowType'] == 'Row':

                                    final_line_data = row['Cells']['Cell']
                                    insert = {}
                                    for index, line_item in enumerate(final_line_data):

                                        try:
                                            insert[index] = float(line_item['Value'])
                                        except:
                                            insert[index] = line_item['Value']
                                    insert['period'] = period
                                    if index1 == 0:
                                        data_from_xero.append(insert)
                                        # if insert['period'] == 1 or insert['period'] == 2:
                                        # print("insert 1", insert)

                                    elif index1 == 1:
                                        data_from_xero_HF_PandL.append(insert)

                                    elif index1 == 2:
                                        # print("Insert A", insert)
                                        data_from_xero_HV_PandL.append(insert)

            period = request['period'] - 1

            filtered_xero = list(filter(lambda x: x['period'] == period, periods_to_report))

            report_from_date = filtered_xero[0]['from_date']
            report_to_date = filtered_xero[0]['to_date']
            period = filtered_xero[0]['period']

            try:
                async with httpx.AsyncClient() as client:
                    response = await client.get(
                        f"https://api.xero.com/api.xro/2.0/Reports/ProfitAndLoss?fromDate={report_from_date}&"
                        f"toDate={report_to_date}&periods={period - 1}&timeframe=MONTH",
                        headers={
                            "Content-Type": "application/xml",
                            "Authorization": f"Bearer {access_token}",
                            "xero-tenant-id": tenant,
                        },
                    )
            except httpx.HTTPError as exc:
                print(f"HTTP Error: {exc}")
                return {"Success": False, "Error": "HTTP Error"}

            # Check the HTTP status code
            if response.status_code != 200:
                return {"Success": False, "Error": "Non-200 Status Code"}

            python_dict = xmltodict.parse(response.text)

            python_dict = python_dict['Response']['Reports']['Report']['Rows']['Row']

            python_dict = list(filter(lambda x: x['RowType'] == 'Section', python_dict))

            for item in python_dict:

                if 'Rows' in item:
                    if isinstance(item['Rows']['Row'], list):
                        data = item['Rows']['Row']
                        for row in data:
                            # print("row", row)
                            if row['RowType'] == 'Row':

                                insert = {}
                                values = []
                                if isinstance(row['Cells']['Cell'], list):

                                    newData = row['Cells']['Cell']

                                    for index, line_item in enumerate(newData):

                                        if index == 0:
                                            insert['Account'] = line_item['Value']
                                        else:
                                            values.append(line_item['Value'])
                                insert['Amount'] = values
                                # print(len(values))

                                if index1 == 0:
                                    # if insert['period'] == 1 or insert['period'] == 2:
                                    #     print("insertA", insert)

                                    comparison_data_cpc.append(insert)
                                    # print("insert 2", insert)


                                elif index1 == 1:

                                    comparison_data_hf.append(insert)


                                elif index1 == 2:
                                    # print("Insert B", insert)
                                    comparison_data_hv.append(insert)



        else:
            # print('XXXXX')
            period = request['period'] - 1
            # print("request", request)

            filtered_xero = list(filter(lambda x: x['period'] == period, periods_to_report))

            report_from_date = request['from_date']
            report_to_date = request['to_date']
            period = request['period']
            # print("Report From Date", report_from_date)
            # print("Report To Date", report_to_date)
            # print("Period", period)

            try:
                async with httpx.AsyncClient() as client:
                    response = await client.get(
                        f"https://api.xero.com/api.xro/2.0/Reports/ProfitAndLoss?fromDate={report_from_date}&"
                        f"toDate={report_to_date}&periods={period - 1}&timeframe=MONTH",
                        headers={
                            "Content-Type": "application/xml",
                            "Authorization": f"Bearer {access_token}",
                            "xero-tenant-id": tenant,
                        },
                    )
            except httpx.HTTPError as exc:
                print(f"HTTP Error: {exc}")
                return {"Success": False, "Error": "HTTP Error"}

            # Check the HTTP status code
            if response.status_code != 200:
                return {"Success": False, "Error": "Non-200 Status Code"}

            python_dict = xmltodict.parse(response.text)

            python_dict = python_dict['Response']['Reports']['Report']['Rows']['Row']

            python_dict = list(filter(lambda x: x['RowType'] == 'Section', python_dict))

            for item in python_dict:
                # print("Item", item)
                # print()

                if 'Rows' in item:
                    if isinstance(item['Rows']['Row'], list):
                        data = item['Rows']['Row']
                        for row in data:
                            # print("row", row)
                            if row['RowType'] == 'Row':

                                insert = {}
                                values = []
                                if isinstance(row['Cells']['Cell'], list):

                                    newData = row['Cells']['Cell']

                                    for index, line_item in enumerate(newData):

                                        if index == 0:
                                            insert['Account'] = line_item['Value']
                                        else:
                                            values.append(line_item['Value'])
                                insert['Amount'] = values
                                # print(len(values))

                                if index1 == 0:
                                    # if insert['period'] == 1 or insert['period'] == 2:
                                    #     print("insertA", insert)

                                    comparison_data_cpc.append(insert)
                                    # print("insert 2", insert)


                                elif index1 == 1:

                                    comparison_data_hf.append(insert)


                                elif index1 == 2:
                                    # print("Insert Long", insert)
                                    comparison_data_hv.append(insert)

    # filter comparison_data_cpc by Account contains 'Heron'

    # print("data_from_xero", data_from_xero)
    # filter out of data_from_xero where Account does not contain 'Heron'
    data_from_xero = list(filter(lambda x: 'Heron' in x[0], data_from_xero))

    data_from_xero = list(filter(lambda x: 'Fees - Construction' not in x[0], data_from_xero))

    comparison_data_cpc = list(filter(lambda x: 'Heron' in x['Account'], comparison_data_cpc))
    # filter out of comparison_data_cpc where Account contains 'Fees - Construction'
    comparison_data_cpc = list(filter(lambda x: 'Fees - Construction' not in x['Account'], comparison_data_cpc))

    if len(data_from_xero) > 0:
        for item in data_from_xero:
            # print("item XX", item)
            comparison_data_cpc_filtered = list(filter(lambda x: x['Account'] == item[0], comparison_data_cpc))
            # print("comparison_data_cpc_filtered", comparison_data_cpc_filtered)
            # if len(comparison_data_cpc_filtered) is greater than 0 then find the index of the item in
            # comparison_data_cpc and insert index 1 of item into comparison_data_cpc[index]['Amount'] at the
            # beginning
            if len(comparison_data_cpc_filtered) > 0:
                index = comparison_data_cpc.index(comparison_data_cpc_filtered[0])
                comparison_data_cpc[index]['Amount'].insert(0, item[1])
                # JUST TO CHECK
                # comparison_data_cpc_filtered = list(filter(lambda x: x['Account'] == item[0], comparison_data_cpc))
                # print("comparison_data_cpc_filtered 2", comparison_data_cpc_filtered)
            # else:

    else:
        print("data_from_xero is empty")

    if len(data_from_xero_HF_PandL) > 0:
        global amount
        amount = 0
        for item in data_from_xero_HF_PandL:
            # print("item YY", item)
            comparison_data_hf_filtered = list(filter(lambda x: x['Account'] == item[0], comparison_data_hf))
            # print("comparison_data_hf_filtered", comparison_data_hf_filtered)
            if len(comparison_data_hf_filtered) == 0 and item[0] == 'Security':
                amount = item[1]
                # print("Amount", amount)

            if len(comparison_data_hf_filtered) > 0:

                if comparison_data_hf_filtered[0]['Account'] == 'Security - ADT':
                    # print("Amount", amount)

                    item[1] = item[1] + amount
                    # print("item[1]", item[1])

                index = comparison_data_hf.index(comparison_data_hf_filtered[0])
                comparison_data_hf[index]['Amount'].insert(0, item[1])
                # JUST TO CHECK
                # comparison_data_hf_filtered = list(filter(lambda x: x['Account'] == item[0], comparison_data_hf))
                # print("comparison_data_hf_filtered 2", comparison_data_hf_filtered)
                # print()
    else:
        print("data_from_xero is empty")

    # print(data_from_xero_HV_PandL)

    # print("comparison_data_hv", comparison_data_hv)

    if len(data_from_xero_HV_PandL) > 0:
        # global amount
        # amount = 0
        for item in data_from_xero_HV_PandL:

            # print("item YY", item)
            comparison_data_hv_filtered = list(filter(lambda x: x['Account'] == item[0], comparison_data_hv))
            # print("comparison_data_hV_filtered", comparison_data_hv_filtered)
            # if len(comparison_data_hv_filtered) == 0 and item[0] == 'Security':
            #     amount = item[1]
            #     print("Amount", amount)

            if len(comparison_data_hv_filtered) > 0:
                # if comparison_data_hf_filtered[0]['Account'] == 'Security - ADT':
                #     print("Amount", amount)
                #
                #     item[1] = item[1] + amount
                #     print("item[1]", item[1])

                index = comparison_data_hv.index(comparison_data_hv_filtered[0])
                comparison_data_hv[index]['Amount'].insert(0, item[1])
                # JUST TO CHECK
                comparison_data_hv_filtered = list(filter(lambda x: x['Account'] == item[0], comparison_data_hv))
                # print("comparison_data_hf_filtered 2", comparison_data_hv_filtered)
                # print()
    else:
        print("data_from_xero is empty")

    # create a variable called profit_loss and fetch all the data from the profit_and_loss collection
    profit_loss = list(db.profit_and_loss.find({}))
    for item in profit_loss:
        item['id'] = str(item['_id'])
        del item['_id']

    # create a new variable from periods_to_report and filter it by period where period is less than or equal to
    # request['period'] - 2 this will be used to get the data from xero for the comparison data
    periods_to_report_comparison = list(filter(lambda x: x['period'] <= request['period'], periods_to_report))
    # reverse periods_to_report_comparison
    periods_to_report_comparison.reverse()

    final_data_for_profit_loss_to_update = []
    final_data_for_profit_loss_to_insert = []

    for item in comparison_data_hv:
        # print("item", item)
        # print()
        # print("periods_to_report_comparison", periods_to_report_comparison)

        for index, value in enumerate(periods_to_report_comparison):

            profit_loss_filtered = list(filter(
                lambda x: x['Account'] == item['Account'] and x['Development'] == 'Heron View' and x['Month'] ==
                          value['Month'], profit_loss))

            if len(profit_loss_filtered) > 0:
                # insert = {}
                insert = profit_loss_filtered[0]
                try:
                    if float(insert['Actual']) != float(item['Amount'][index]):
                        insert['Actual'] = float(item['Amount'][index])
                        # print("update", insert)
                        # print("Got this far!!!", item)
                        final_data_for_profit_loss_to_update.append(insert)
                except IndexError:
                    insert['Actual'] = 0
                    # insert[]
                    final_data_for_profit_loss_to_update.append(insert)

                    # print("item", item)
                    # print("insert", insert)
                    # print("index", index)
                    # print("profit_loss_filtered", profit_loss_filtered)

                    # break

            else:
                if 'Category' not in item:
                    if item['Account'].startswith("COS"):
                        item['Category'] = 'COS'

                    elif item['Account'].startswith("Bond") or item['Account'].startswith("Sale"):
                        item['Category'] = 'Trading Income'

                    elif item['Account'].startswith("Rental") or item['Account'].startswith("Interest Received"):
                        item['Category'] = 'Other Income'

                    else:
                        item['Category'] = 'Operating Expenses'

                insert = {'Account': item['Account'], 'Actual': item['Amount'][index], 'Forecast': 0,
                          'Development': 'Heron View', 'Applicable_dev': 'Heron View', 'Month': value['Month'],
                          "Category": item['Category']}
                # print("insert", insert)

                final_data_for_profit_loss_to_insert.append(insert)

    for item in comparison_data_hf:

        for index, value in enumerate(periods_to_report_comparison):

            profit_loss_filtered = list(filter(
                lambda x: x['Account'] == item['Account'] and x['Development'] == 'Heron Fields' and x[
                    'Month'] == value[
                              'Month'], profit_loss))

            if len(profit_loss_filtered) > 0:
                # insert = {}
                insert = profit_loss_filtered[0]
                try:
                    if float(insert['Actual']) != float(item['Amount'][index]):
                        # print("profit_loss_filtered", profit_loss_filtered)
                        insert['Actual'] = float(item['Amount'][index])
                        final_data_for_profit_loss_to_update.append(insert)
                except IndexError:
                    insert['Actual'] = 0
                    final_data_for_profit_loss_to_update.append(insert)

            else:
                if 'Category' not in item:
                    if item['Account'].startswith("COS"):
                        item['Category'] = 'COS'

                    elif item['Account'].startswith("Bond") or item['Account'].startswith("Sale"):
                        item['Category'] = 'Trading Income'

                    elif item['Account'].startswith("Rental") or item['Account'].startswith("Interest Received"):
                        item['Category'] = 'Other Income'

                    else:
                        item['Category'] = 'Operating Expenses'
                # print("item", item)
                insert = {'Account': item['Account'], 'Actual': item['Amount'][index], 'Forecast': 0,
                          'Development': 'Heron Fields', 'Applicable_dev': 'Heron Fields', 'Month': value['Month'],
                          "Category": item['Category']}

                final_data_for_profit_loss_to_insert.append(insert)

    for item in comparison_data_cpc:

        for index, value in enumerate(periods_to_report_comparison):

            profit_loss_filtered = list(filter(
                lambda x: x['Account'] == item['Account'] and x['Development'] == 'CPC' and x[
                    'Month'] == value[
                              'Month'], profit_loss))

            # print()
            # print("profit_loss_filtered", profit_loss_filtered)

            if len(profit_loss_filtered) > 0:

                insert = profit_loss_filtered[0]
                # if insert['Account'] == 'COS - Heron View - Construction':
                #     print("insert XX", insert)
                #     print("item XX", item)
                #     print("index XX", index)
                #     print("value XX", value)
                try:
                    if float(insert['Actual']) != float(item['Amount'][index]):
                        insert['Actual'] = float(item['Amount'][index])
                        final_data_for_profit_loss_to_update.append(insert)
                        # if insert['Account'] == 'COS - Heron View - Construction':
                        #     print("insert", insert)
                        #     print("item", item)
                        #     print("index", index)
                        #     print("value", value)
                except IndexError:
                    # if insert['Account'] == 'COS - Heron View - Construction':
                    #     print("insertXX", insert)
                    #     print("itemXX", item)
                    #     print("indexXX", index)
                    #     print("valueXX", value)
                    #     insert['Actual'] = 0
                    final_data_for_profit_loss_to_update.append(insert)

            else:

                if 'Category' not in item:
                    if item['Account'].startswith("COS"):
                        item['Category'] = 'COS'

                    elif item['Account'].startswith("Bond") or item['Account'].startswith("Sale"):
                        item['Category'] = 'Trading Income'

                    elif item['Account'].startswith("Rental") or item['Account'].startswith("Interest Received"):
                        item['Category'] = 'Other Income'

                    else:
                        item['Category'] = 'Operating Expenses'

                insert = {'Account': item['Account'], 'Actual': item['Amount'][index], 'Forecast': 0,
                          'Development': 'CPC', 'Applicable_dev': 'Heron View', 'Month': value['Month'],
                          "Category": item['Category']}

                final_data_for_profit_loss_to_insert.append(insert)

    # update the data in the profit_and_loss collection with the data in final_data_for_profit_loss_to_update

    # print("final_data_for_profit_loss_to_update", final_data_for_profit_loss_to_update)

    if len(final_data_for_profit_loss_to_update) > 0:
        for item in final_data_for_profit_loss_to_update:
            id = item['id']
            del item['id']
            db.profit_and_loss.update_one({"_id": ObjectId(id)}, {"$set": item})

    # insert the data in the profit_and_loss collection with the data in final_data_for_profit_loss_to_insert
    if len(final_data_for_profit_loss_to_insert) > 0:
        db.profit_and_loss.insert_many(final_data_for_profit_loss_to_insert)

    # GET FINAL DATA FROM DB
    profit_and_loss = list(db.profit_and_loss.find({}))

    # GET PRINT FORM DATA
    # get opportunities from db where Category is equal to 'Heron Fields' or 'Heron View'
    opportunities = list(db.opportunities.find({"Category": {"$in": ["Heron Fields", "Heron View"]}}))
    # sales_processed from the db where development is equal to 'Heron Fields' or 'Heron View', project only
    # opportunity_code and opportunity_sales_date
    sales_processed = list(db.sales_processed.find({"development": {"$in": ["Heron Fields", "Heron View"]}},
                                                   {"opportunity_code": 1, "opportunity_sales_date": 1,
                                                    "opportunity_transfer_fees": 1,
                                                    "opportunity_bond_registration": 1,
                                                    "opportunity_trust_release_fee": 1, "opportunity_unforseen": 1,
                                                    "opportunity_commission": 1}))

    # create a variable called sum_to_raise which is the sum of
    opportunity_amount_required = sum([float(item['opportunity_amount_required']) for item in opportunities])
    # do the same as the above, but only where Category = 'Heron Fields'
    opportunity_amount_required_hf = sum(
        [float(item['opportunity_amount_required']) for item in opportunities if
         item['Category'] == 'Heron Fields'])
    # print("opportunity_amount_required_hf", opportunity_amount_required_hf)
    # print("opportunity_amount_required", opportunity_amount_required)
    # from investors in the db get all investors where 'Category' in the trust array is equal to 'Heron Fields'
    # or 'Heron View'
    for sales in sales_processed:
        sales['id'] = str(sales['_id'])
        del sales['_id']
        sales['opportunity_bond_registration'] = float(sales.get('opportunity_bond_registration', 0))
        sales['opportunity_transfer_fees'] = float(sales.get('opportunity_transfer_fees', 0))
        sales['opportunity_trust_release_fee'] = float(sales.get('opportunity_trust_release_fee', 0))
        sales['opportunity_unforseen'] = float(sales.get('opportunity_unforseen', 0))
        sales['opportunity_commission'] = float(sales.get('opportunity_commission', 0))
        # print("sales", sales)

    # print("to_date", request['to_date'])
    # convert request['to_date'] to datetime as a variable called to_date
    to_date = datetime.strptime(request['to_date'], "%Y-%m-%d")
    dates = []
    for i in range(-3, 0):
        dates.append(to_date + relativedelta(months=i))
    dates.append(to_date)

    trust_investments_received = []
    released_investments = []

    investors = list(db.investors.find({}))

    for investor in investors:
        investor['id'] = str(investor['_id'])
        del investor['_id']
        if investor['investor_acc_number'] == "ZCAM01":
            # filter out the first item in the trust array and the first item in the investments array
            investor['trust'] = investor['trust'][1:]
            investor['investments'] = investor['investments'][1:]
        if investor['investor_acc_number'] == "ZJHO01":
            # filter out of trust and investments where opportunity_code is equal to 'HFA205' and
            # investment_number is equal to 1
            investor['trust'] = list(filter(lambda x: not (x['investment_number'] == 1), investor['trust']))
            investor['investments'] = list(
                filter(lambda x: not (x['investment_number'] == 1), investor['investments']))
        if investor['investor_acc_number'] == "ZPJB01":
            # filter out of trust and investments where opportunity_code is equal to 'HFA205' and
            # investment_number is equal to 1
            investor['trust'] = list(
                filter(lambda x: not (x['opportunity_code'] == "HFA205" and x['investment_number'] == 1),
                       investor['trust']))
            investor['investments'] = list(
                filter(lambda x: not (x['opportunity_code'] == "HFA205" and x['investment_number'] == 1),
                       investor['investments']))
        if investor['investor_acc_number'] == "ZERA01":
            # filter out of trust and investments where opportunity_code is equal to 'EA205' and
            # investment_number is equal to 3
            investor['trust'] = list(
                filter(lambda x: not (x['opportunity_code'] == "EA205" and x['investment_number'] == 3),
                       investor['trust']))
            investor['investments'] = list(
                filter(lambda x: not (x['opportunity_code'] == "EA205" and x['investment_number'] == 3),
                       investor['investments']))
        if investor['investor_acc_number'] == "ZVOL01":
            # filter out of trust and investments where opportunity_code is equal to 'EA103' and
            # investment_number is equal to 3
            investor['trust'] = list(
                filter(lambda x: not (x['opportunity_code'] == "EA103" and x['investment_number'] == 3),
                       investor['trust']))
            investor['investments'] = list(
                filter(lambda x: not (x['opportunity_code'] == "EA103" and x['investment_number'] == 3),
                       investor['investments']))
        if investor['investor_acc_number'] == "ZLEW03":
            # filter out of trust and investments where opportunity_code is equal to 'EA205' and
            # investment_number is equal to 1
            investor['trust'] = list(
                filter(lambda x: not (x['opportunity_code'] == "EA205" and x['investment_number'] == 1),
                       investor['trust']))
            investor['investments'] = list(
                filter(lambda x: not (x['opportunity_code'] == "EA205" and x['investment_number'] == 1),
                       investor['investments']))

        # filter trust array by 'Category' in the trust array is equal to 'Heron Fields' or 'Heron View'
        investor['trust'] = list(
            filter(lambda x: x['Category'] in ['Heron Fields', 'Heron View'], investor['trust']))
        # filter investments array by 'Category' in the trust array is equal to 'Heron Fields' or 'Heron View'
        investor['investments'] = list(
            filter(lambda x: x['Category'] in ['Heron Fields', 'Heron View'], investor['investments']))

        # loop through trust array and convert 'deposit_date' to datetime
        if len(investor['trust']) > 0:
            for item in investor['trust']:
                item['deposit_date'] = datetime.strptime(item['deposit_date'], "%Y/%m/%d")
                item['release_date'] = item.get('release_date', "")
                if item['release_date'] != "":
                    item['release_date'] = datetime.strptime(item['release_date'], "%Y/%m/%d")
                    item['planned_release_date'] = item['release_date']

                else:
                    item['release_date'] = "2099/12/31"
                    item['release_date'] = datetime.strptime(item['release_date'], "%Y/%m/%d")
                    item['planned_release_date'] = datetime.strptime(
                        item.get('planned_release_date', "2099/12/31").replace("-", "/"), "%Y/%m/%d")

                opportunities_filtered = list(filter(lambda x: x['opportunity_code'] == item['opportunity_code'],
                                                     opportunities))
                if opportunities_filtered[0]['opportunity_final_transfer_date'] == "":
                    estimated_final_date = opportunities_filtered[0]['opportunity_end_date'].replace("-", "/")
                    estimated_final_date = datetime.strptime(estimated_final_date, "%Y/%m/%d")
                else:
                    estimated_final_date = opportunities_filtered[0]['opportunity_final_transfer_date'].replace(
                        "-", "/")
                    estimated_final_date = datetime.strptime(estimated_final_date, "%Y/%m/%d")

                insert = {}

                insert['opportunity_code'] = item['opportunity_code']
                insert['investor_acc_number'] = investor['investor_acc_number']
                insert['planned_release_date'] = item['planned_release_date']
                # print("insert['planned_release_date']", insert['planned_release_date'])
                insert['investment_number'] = item['investment_number']
                insert['deposit_date'] = item['deposit_date']
                insert['release_date'] = item['release_date']
                insert['amount'] = float(item['investment_amount'])
                insert['estimated_final_date'] = estimated_final_date
                insert['interest_rate'] = float(item.get('investment_interest_rate', 18))

                filtered_opportunities = list(filter(lambda x: x['opportunity_code'] == item['opportunity_code'],
                                                     opportunities))
                insert['amount_required'] = float(filtered_opportunities[0]['opportunity_amount_required'])

                trust_investments_received.append(insert)

        if len(investor['investments']) > 0:
            for index, item in enumerate(investor['investments']):
                opportunities_filtered = list(filter(lambda x: x['opportunity_code'] == item['opportunity_code'],
                                                     opportunities))
                # print("opportunities_filtered", opportunities_filtered[0])
                item['sold'] = opportunities_filtered[0]['opportunity_sold']

                if opportunities_filtered[0]['opportunity_sold'] == True:
                    # print(item['opportunity_code'], index, investor['investor_acc_number'])
                    sales_processed_filtered = list(
                        filter(lambda x: x['opportunity_code'] == item['opportunity_code'],
                               sales_processed))

                    item['sale_date'] = datetime.strptime(
                        sales_processed_filtered[0]['opportunity_sales_date'].replace('-', '/'), "%Y/%m/%d")

                    if opportunities_filtered[0]['opportunity_final_transfer_date'] != "":
                        item['transferred'] = True
                        item['transfer_date'] = datetime.strptime(
                            opportunities_filtered[0]['opportunity_final_transfer_date'].replace('-', '/'),
                            "%Y/%m/%d")
                    else:
                        item['transferred'] = False
                        item['transfer_date'] = "2099/12/31"
                        item['transfer_date'] = datetime.strptime(item['transfer_date'], "%Y/%m/%d")
                else:
                    item['transferred'] = False
                    item['transfer_date'] = "2099/12/31"
                    item['transfer_date'] = datetime.strptime(item['transfer_date'], "%Y/%m/%d")
                    item['sale_date'] = "2099/12/31"
                    item['sale_date'] = datetime.strptime(item['sale_date'], "%Y/%m/%d")

                # item['deposit_date'] = datetime.strptime(item['deposit_date'], "%Y/%m/%d")
                item['end_date'] = item.get('end_date', "2099/12/31")
                item['early_release'] = item.get('early_release', False)
                item['release_date'] = item.get('release_date', "2099/12/31")
                if item['release_date'] != "":
                    item['release_date'] = datetime.strptime(item['release_date'], "%Y/%m/%d")
                else:
                    item['release_date'] = "2099/12/31"
                    item['release_date'] = datetime.strptime(item['release_date'], "%Y/%m/%d")
                if item['end_date'] != "":
                    item['end_date'] = datetime.strptime(item['end_date'], "%Y/%m/%d")
                else:
                    item['end_date'] = "2099/12/31"
                    item['end_date'] = datetime.strptime(item['end_date'], "%Y/%m/%d")

                if opportunities_filtered[0]['opportunity_final_transfer_date'] == "":
                    estimated_final_date = opportunities_filtered[0]['opportunity_end_date'].replace("-", "/")
                    estimated_final_date = datetime.strptime(estimated_final_date, "%Y/%m/%d")
                else:
                    estimated_final_date = opportunities_filtered[0]['opportunity_final_transfer_date'].replace(
                        "-", "/")
                    estimated_final_date = datetime.strptime(estimated_final_date, "%Y/%m/%d")

                insert = {}
                insert['opportunity_code'] = item['opportunity_code']
                insert['investor_acc_number'] = investor['investor_acc_number']
                insert['investment_number'] = item['investment_number']
                insert['early_release'] = item['early_release']
                insert['sold'] = item['sold']
                insert['sale_date'] = item['sale_date']
                insert['transferred'] = item['transferred']
                insert['transfer_date'] = item['transfer_date']
                insert['interest_rate'] = float(item['investment_interest_rate'])
                insert['estimated_final_date'] = estimated_final_date

                insert['end_date'] = item['end_date']
                insert['release_date'] = item['release_date']
                insert['amount'] = float(item['investment_amount'])

                filtered_opportunities = list(filter(lambda x: x['opportunity_code'] == item['opportunity_code'],
                                                     opportunities))
                insert['amount_required'] = float(filtered_opportunities[0]['opportunity_amount_required'])

                released_investments.append(insert)

    # sort trust_investments_received by opportunity_code, investor_acc_number
    trust_investments_received = sorted(trust_investments_received,
                                        key=lambda i: (i['opportunity_code'], i['investor_acc_number']))
    # sort released_investments by opportunity_code, investor_acc_number
    released_investments = sorted(released_investments,
                                  key=lambda i: (i['opportunity_code'], i['investor_acc_number']))

    # print("trust_investments_received", trust_investments_received[0], len(trust_investments_received))
    # print()
    # print("released_investments", released_investments[0], len(released_investments))
    # print()

    sales_units = []

    for investment in released_investments:
        insert = {}
        opportunities_filtered = list(filter(lambda x: x['opportunity_code'] == investment['opportunity_code'],
                                             opportunities))

        insert['opportunity_code'] = investment['opportunity_code']
        insert['sold'] = investment['sold']
        insert['sale_date'] = investment['sale_date']
        insert['sales_price'] = float(opportunities_filtered[0]['opportunity_sale_price'])
        insert['transferred'] = investment['transferred']
        insert['transfer_date'] = investment['transfer_date']
        sales_units.append(insert)

    # remove all duplicates from sales_units
    sales_units = [i for n, i in enumerate(sales_units) if i not in sales_units[n + 1:]]

    for opportunity in opportunities:
        sales_units_filtered = list(
            filter(lambda x: x['opportunity_code'] == opportunity['opportunity_code'], sales_units))

        if len(sales_units_filtered) == 0:

            insert = {}
            insert['opportunity_code'] = opportunity['opportunity_code']
            insert['sold'] = opportunity['opportunity_sold']
            if opportunity['opportunity_sold'] == True:
                sales_processed_filtered = list(
                    filter(lambda x: x['opportunity_code'] == item['opportunity_code'],
                           sales_processed))
                if len(sales_processed_filtered) > 0 and sales_processed_filtered[0]['opportunity_sales_date'] != None:
                    sale_date = sales_processed_filtered[0]['opportunity_sales_date'].replace('-', '/')
                else:
                    sale_date = "2099/12/31"
            else:
                sale_date = "2099/12/31"
            if opportunity['opportunity_final_transfer_date'] != "":
                insert['transferred'] = True
                transfer_date = opportunity['opportunity_final_transfer_date']
            else:
                insert['transferred'] = False
                transfer_date = "2099/12/31"
            insert['sale_date'] = datetime.strptime(sale_date.replace('-', '/'), "%Y/%m/%d")
            insert['sales_price'] = float(opportunity['opportunity_sale_price'])

            insert['transfer_date'] = datetime.strptime(transfer_date.replace('-', '/'), "%Y/%m/%d")
            sales_units.append(insert)

    # sort sales_units by opportunity_code
    sales_units = sorted(sales_units, key=lambda i: i['opportunity_code'])

    print()
    month_4 = 0
    month_3 = 0
    month_2 = 0
    month_1 = 0

    month_4released = 0
    month_3released = 0
    month_2released = 0
    month_1released = 0

    month_4mom = 0
    month_3mom = 0
    month_2mom = 0
    month_1mom = 0

    # SALES_PARAMETERS
    # get sales_parameters from db where Development is equal to 'Heron View'
    sales_parameters = list(db.salesParameters.find({"Development": "Heron View"}))
    for sale in sales_parameters:
        sale['id'] = str(sale['_id'])
        del sale['_id']
        # print("sale", sale)

    for sale in sales_units:
        for param in sales_parameters:
            # make the sale key equal to the param['Description'] and the value equal to the param['rate'] as a
            # float
            sale[param['Description']] = float(param['rate'])

    for sale in sales_units:
        sales_processed_filtered = list(
            filter(lambda x: x['opportunity_code'] == sale['opportunity_code'],
                   sales_processed))
        if len(sales_processed_filtered) > 0:
            # print("sales_processed_filtered", sales_processed_filtered[0])
            if sales_processed_filtered[0]['opportunity_sales_date'] != None:
                sale_date = sales_processed_filtered[0]['opportunity_sales_date'].replace('-', '/')
            else:
                sale_date = "2024/02/26"
            sale['sale_date'] = datetime.strptime(sale_date, "%Y/%m/%d")
            sale['sold'] = True

            if sales_processed_filtered[0]['opportunity_trust_release_fee'] != 0:
                sale['trust_release_fee'] = sales_processed_filtered[0]['opportunity_trust_release_fee']
            if sales_processed_filtered[0]['opportunity_bond_registration'] != 0:
                sale['bond_registration'] = sales_processed_filtered[0]['opportunity_bond_registration']
            if sales_processed_filtered[0]['opportunity_transfer_fees'] != 0:
                sale['transfer_fees'] = sales_processed_filtered[0]['opportunity_transfer_fees']
            if sales_processed_filtered[0]['opportunity_unforseen'] != 0:
                sale['unforseen'] = sales_processed_filtered[0]['opportunity_unforseen']
            if sales_processed_filtered[0]['opportunity_commission'] != 0:
                sale['commission'] = sales_processed_filtered[0]['opportunity_commission']

    value_to_be_raised = []

    # the above again using list comprehension
    units = [investment['opportunity_code'] for investment in opportunities]
    # eliminate duplicates from units which is a list of strings
    units = list(dict.fromkeys(units))
    # sort units by opportunity_code
    units = sorted(units)

    for index, unit in enumerate(units):
        opportunities_filtered = list(
            filter(lambda x: x['opportunity_code'] == unit, opportunities))
        insert = {}
        insert['opportunity_code'] = unit
        insert['amount'] = 0
        insert['amount_required'] = float(opportunities_filtered[0]['opportunity_amount_required'])
        if opportunities_filtered[0]['opportunity_final_transfer_date'] == "":
            insert['transferred'] = False
            insert['transfer_date'] = "2099/12/31"
            insert['transfer_date'] = datetime.strptime(insert['transfer_date'], "%Y/%m/%d")
        else:
            insert['transferred'] = True
            insert['transfer_date'] = datetime.strptime(
                opportunities_filtered[0]['opportunity_final_transfer_date'].replace('-', '/'), "%Y/%m/%d")
        trust_investments_received_filtered = list(
            filter(lambda x: x['opportunity_code'] == unit, trust_investments_received))
        if len(trust_investments_received_filtered) > 0:
            insert['amount'] = sum([item['amount'] for item in trust_investments_received_filtered])
        else:
            released_investments_filtered = list(
                filter(lambda x: x['opportunity_code'] == unit, released_investments))
            if len(released_investments_filtered) > 0:
                insert['amount'] = sum([item['amount'] for item in released_investments_filtered])
        value_to_be_raised.append(insert)

    # get rates from the db
    rates = list(db.rates.find({}))

    for rate in rates:
        rate['id'] = str(rate['_id'])
        del rate['_id']
        # convert rate['Efective_date'] to datetime
        rate['Efective_date'] = datetime.strptime(rate['Efective_date'].replace('-', '/'), "%Y/%m/%d")
        rate['rate'] = float(rate['rate']) / 100
    # sort rates by Efective_date descending
    rates = sorted(rates, key=lambda i: i['Efective_date'], reverse=True)

    # SALES_UNITS COSTS
    # print()
    # print("sales_units", sales_units[0])
    # print()

    for item in trust_investments_received:
        released_investments_filtered = list(
            filter(lambda x: x['opportunity_code'] == item['opportunity_code'] and
                             x['investor_acc_number'] == item['investor_acc_number'] and
                             x['investment_number'] == item['investment_number'],
                   released_investments))
        if len(released_investments_filtered) > 0:
            item['transferred'] = released_investments_filtered[0]['transferred']
            item['transfer_date'] = released_investments_filtered[0]['transfer_date']
            item['interest_rate'] = float(released_investments_filtered[0]['interest_rate'])
            if released_investments_filtered[0]['early_release'] == True:
                item['estimated_final_date'] = released_investments_filtered[0]['end_date']
        else:
            item['transferred'] = False
            item['transfer_date'] = "2099/12/31"
            item['transfer_date'] = datetime.strptime(item['transfer_date'], "%Y/%m/%d")

    # get unallocated_investments from the db
    unallocated_investments = list(db.unallocated_investments.find({}))
    for item in unallocated_investments:
        item['id'] = str(item['_id'])
        del item['_id']

        if item['opportunity_final_transfer_date'] == "":
            item['end_date'] = item['opportunity_end_date']
        else:
            item['end_date'] = item['opportunity_final_transfer_date']
        item['end_date'] = datetime.strptime(item['end_date'].replace('-', '/'), "%Y/%m/%d")
        item['deposit_date'] = datetime.strptime(item['deposit_date'].replace('-', '/'), "%Y/%m/%d")
        item['release_date'] = datetime.strptime(item['release_date'].replace('-', '/'), "%Y/%m/%d")

    for index, date in enumerate(dates, 1):

        unallocated_investment_released_interest = 0
        unallocated_investment_capital = 0
        unallocated_investment_momentum_interest = 0

        for item in unallocated_investments:
            unallocated_investment_capital += item['unallocated_investment']
            unallocated_investment_released_interest += (item['unallocated_investment'] * item[
                'project_interest_rate'] / 100 / 365) * (item['end_date'] - item['release_date']).days
            momentum_interest = 0
            start_date = item['deposit_date']
            start_date += timedelta(days=1)
            end_date = item['release_date']
            while start_date <= end_date:
                rate = list(filter(lambda x: x['Efective_date'] <= start_date, rates))[0]['rate']
                momentum_interest += item['unallocated_investment'] * rate / 365
                start_date += timedelta(days=1)
            unallocated_investment_momentum_interest += momentum_interest

        total_unallocated_interest = unallocated_investment_released_interest + unallocated_investment_momentum_interest

        for item in trust_investments_received:
            amount = item['amount']
            start_date = item['deposit_date']
            start_date += timedelta(days=1)
            if item['transferred'] == True and item['transfer_date'] <= date:

                if item['release_date'] != "":
                    if item['release_date'] > date:
                        end_date = date
                    else:
                        end_date = item['release_date']
                else:
                    end_date = date

                interest = 0
                interest_all_mom = 0
                while start_date <= end_date:
                    rate = list(filter(lambda x: x['Efective_date'] <= start_date, rates))[0]['rate']
                    interest += amount * rate / 365
                    start_date += timedelta(days=1)
                item['interest'] = interest
                item['interest_other'] = 0
            else:
                item['interest'] = 0
                item['interest_other'] = 0
                interest = 0
                if item['release_date'] >= item['planned_release_date']:
                    end_date = item['planned_release_date']
                else:
                    end_date = item['release_date']

                while start_date <= end_date:
                    rate = list(filter(lambda x: x['Efective_date'] <= start_date, rates))[0]['rate']
                    interest += amount * rate / 365
                    start_date += timedelta(days=1)
                item['interest'] = 0
                item['interest_other'] = interest
            released_interest = 0
            new_start_date = end_date
            # new_start_date += timedelta(days=1)
            new_end_date = item['estimated_final_date']
            released_interest = amount * item['interest_rate'] / 100 / 365 * (new_end_date - new_start_date).days
            item['released_interest'] = released_interest

        interest_mom = sum([item['interest'] for item in trust_investments_received])
        # print("interestMom", interest_mom)
        interest_mom_other = sum([item['interest_other'] for item in trust_investments_received])
        # print("interestMomOther", interest_mom_other)
        released_interest = sum([item['released_interest'] for item in trust_investments_received])
        # print("releasedInterest", released_interest)

        for item in released_investments:
            interest = 0
            if item['transferred'] == True and item['transfer_date'] <= date:
                amount = item['amount']
                start_date = item['release_date']
                if item['end_date'] != "":
                    end_date = item['end_date']
                else:
                    end_date = item['transfer_date']
                interest = 0
                interest = (amount * item['interest_rate'] / 100) / 365 * (end_date - start_date).days
                item['interest'] = interest
            else:
                item['interest'] = 0
            # if item['opportunity_code'] == 'HFA101':
            #     print("item", item)

        interest_released = sum([item['interest'] for item in released_investments])
        # print("interestReleased", interest_released)

        interest_paid_to_date = interest_mom + interest_released

        total_estimated_interest = interest_mom_other + released_interest + interest_mom + total_unallocated_interest
        # print("totalEstimatedInterest", total_estimated_interest + total_unallocated_interest)
        # print("interest", interest)

        # sum all the amounts in trust_investments_received where deposit_date is less than or equal to date
        month_received = sum(
            [item['amount'] for item in trust_investments_received if item['deposit_date'] <= date])

        # do exactly as above but only where the opportunity_code begins with 'HF'
        month_hf_received = sum([item['amount'] for item in trust_investments_received if
                                 item['deposit_date'] <= date and item['opportunity_code'].startswith('HF')])

        # print("month_hf", month_hf_received)
        # print("Capital_not Raised", opportunity_amount_required_hf - month_hf_received)
        capital_not_raised = opportunity_amount_required_hf - month_hf_received

        # sum all the amounts in trust_investments_received where release_date is less than or equal to date if
        # released_date is not equal to ""
        month_released = sum(
            [item['amount'] for item in trust_investments_received if item['release_date'] != "" and
             item['release_date'] <= date])
        # sum all the amounts in trust_investments_received where deposit_date is less than or equal to date AND
        # release_date is greater than to date or release_date is equal to ""
        month_mom = sum(
            [item['amount'] for item in trust_investments_received if
             item['deposit_date'] <= date and (item['release_date'] > date or item['release_date'] == "")])

        count_received = len(
            [item['amount'] for item in trust_investments_received if item['deposit_date'] != "" and
             item['deposit_date'] <= date])

        count_exited = len(
            [item['amount'] for item in released_investments if item['transfer_date'] != "" and
             item['transfer_date'] <= date])

        value_exited = sum(
            [item['amount'] for item in released_investments if item['transfer_date'] != "" and
             item['transfer_date'] <= date])

        early_release = sum(
            [item['amount'] for item in released_investments if
             item['early_release'] == True and item['transferred'] == False and
             item['end_date'] <= date])

        current_investor_capital_deployed = month_released - value_exited - early_release

        # make a deep copy of value_to_be_raised
        value_to_be_raised_list = copy.deepcopy(value_to_be_raised)

        for item in value_to_be_raised_list:
            if item['transferred']:

                if item['transfer_date'] > date:
                    item['transferred'] = False

            # value_to_be_raised_list.append(insert)

        value_still_to_be_raised = sum(
            [item['amount_required'] - item['amount'] for item in value_to_be_raised_list if
             item['transferred'] == False and item['transfer_date'] > date])

        # get the length of sales_units
        count_units = len(sales_units)

        # transferred_units is length of sales_units where transferred is equal to True and transfer_date is less
        # than or equal to date
        transferred_units = len(
            [item['sales_price'] for item in sales_units if
             item['transferred'] == True and item['sold'] == True and item['transfer_date'] <= date])

        transferred_units_sold_value = sum(
            [item['sales_price'] for item in sales_units if
             item['transferred'] == True and item['sold'] == True and item['transfer_date'] <= date]) / 1.15

        transferred_units_commission = sum(
            [(item['sales_price'] / 1.15) * (item['commission'] / 1.15) for item in sales_units if
             item['transferred'] == True and item['sold'] == True and item['transfer_date'] <= date])

        transferred_units_transfer_fees = sum(
            [item['transfer_fees'] for item in sales_units if
             item['transferred'] == True and item['sold'] == True and item['transfer_date'] <= date]) / 1.15

        transferred_units_bond_registration = sum(
            [item['bond_registration'] for item in sales_units if
             item['transferred'] == True and item['sold'] == True and item['transfer_date'] <= date]) / 1.15

        transferred_units_trust_release_fee = sum(
            [(item['trust_release_fee'] / 1.15) for item in sales_units if
             item['transferred'] == True and item['sold'] == True and item['transfer_date'] <= date])

        transferred_units_unforseen = sum(
            [(item['sales_price'] / 1.15) * (item['unforseen']) for item in sales_units if
             item['transferred'] == True and item['sold'] == True and item['transfer_date'] <= date])

        total_sale_value = sum(
            [item['sales_price'] for item in sales_units if
             item['sale_date'] <= date]) / 1.15

        completely_unsold_units = len(
            [item['sales_price'] for item in sales_units if
             item['sold'] == False and item['transferred'] == False])

        sold_after_date = len(
            [item['sales_price'] for item in sales_units if
             item['sold'] == True and item['sale_date'] > date])

        total_units_sales_value = sum(
            [item['sales_price'] for item in sales_units]) / 1.15

        total_units_commission = sum(
            [(item['sales_price'] / 1.15) * (item['commission'] / 1.15) for item in sales_units])

        total_units_transfer_fees = sum(
            [item['transfer_fees'] for item in sales_units]) / 1.15

        total_units_bond_registration = sum(
            [item['bond_registration'] for item in sales_units]) / 1.15

        total_units_trust_release_fee = sum(
            [(item['trust_release_fee'] / 1.15) for item in sales_units])

        total_units_unforseen = sum(
            [(item['sales_price'] / 1.15) * (item['unforseen']) for item in sales_units])

        remaining_units_value = completely_unsold_units + sold_after_date

        # create a variable called adjusted_sales_units and deep copy sales_units
        adjusted_sales_units = copy.deepcopy(sales_units)

        sold_units_filterred = []
        for item in adjusted_sales_units:
            insert = item
            if item['sold'] == True and item['sale_date'] <= date:

                insert['sold'] = True
            else:

                insert['sold'] = False
            if item['transferred'] == True and item['transfer_date'] <= date:
                insert['transferred'] = True
            else:
                insert['transferred'] = False
            sold_units_filterred.append(insert)

        sold_units = len(
            [item['sales_price'] for item in sold_units_filterred if
             item['sold'] == True and item['transferred'] == False])

        sold_units_value = sum(
            [item['sales_price'] for item in sold_units_filterred if
             item['sold'] == True and item['transferred'] == False]) / 1.15

        sold_units_commission = sum(
            [(item['sales_price'] / 1.15) * (item['commission'] / 1.15) for item in sold_units_filterred if
             item['sold'] == True and item['transferred'] == False])

        sold_units_transfer_fees = sum(
            [item['transfer_fees'] for item in sold_units_filterred if
             item['sold'] == True and item['transferred'] == False]) / 1.15

        sold_units_bond_registration = sum(
            [item['bond_registration'] for item in sold_units_filterred if
             item['sold'] == True and item['transferred'] == False]) / 1.15

        sold_units_trust_release_fee = sum(
            [(item['trust_release_fee'] / 1.15) for item in sold_units_filterred if
             item['sold'] == True and item['transferred'] == False])

        sold_units_unforseen = sum(
            [(item['sales_price'] / 1.15) * (item['unforseen']) for item in sold_units_filterred if
             item['sold'] == True and item['transferred'] == False])

        remaining_units = count_units - transferred_units - sold_units

        remaining_units_value = total_units_sales_value - transferred_units_sold_value - sold_units_value

        remaining_units_commission = total_units_commission - transferred_units_commission - sold_units_commission

        remaining_units_transfer_fees = total_units_transfer_fees - transferred_units_transfer_fees - sold_units_transfer_fees

        remaining_units_bond_registration = total_units_bond_registration - transferred_units_bond_registration - sold_units_bond_registration

        remaining_units_trust_release_fee = total_units_trust_release_fee - transferred_units_trust_release_fee - sold_units_trust_release_fee

        remaining_units_unforseen = total_units_unforseen - transferred_units_unforseen - sold_units_unforseen

        if index == 1:
            month_1_received = month_received
            month_1released = month_released
            capital_not_raised1 = capital_not_raised

            month_1mom = month_mom
            count_received1 = count_received
            count_exited1 = count_exited
            count_units1 = count_units

            value_exited1 = value_exited

            current_investor_capital_deployed1 = current_investor_capital_deployed

            value_still_to_be_raised1 = value_still_to_be_raised

            transferred_units1 = transferred_units
            sold_units1 = sold_units
            remaining_units1 = remaining_units
            transferred_units_sold_value1 = transferred_units_sold_value
            sold_units_value1 = sold_units_value
            remaining_units_value1 = remaining_units_value
            total_sale_value1 = total_sale_value
            total_units_sales_value1 = total_units_sales_value

            total_units_commission1 = total_units_commission
            transferred_units_commission1 = transferred_units_commission
            sold_units_commission1 = sold_units_commission
            remaining_units_commission1 = remaining_units_commission

            total_units_transfer_fees1 = total_units_transfer_fees
            transferred_units_transfer_fees1 = transferred_units_transfer_fees
            sold_units_transfer_fees1 = sold_units_transfer_fees
            remaining_units_transfer_fees1 = remaining_units_transfer_fees

            total_units_bond_registration1 = total_units_bond_registration
            transferred_units_bond_registration1 = transferred_units_bond_registration
            sold_units_bond_registration1 = sold_units_bond_registration
            remaining_units_bond_registration1 = remaining_units_bond_registration

            total_units_trust_release_fee1 = total_units_trust_release_fee
            transferred_units_trust_release_fee1 = transferred_units_trust_release_fee
            sold_units_trust_release_fee1 = sold_units_trust_release_fee
            remaining_units_trust_release_fee1 = remaining_units_trust_release_fee

            total_units_unforseen1 = total_units_unforseen
            transferred_units_unforseen1 = transferred_units_unforseen
            sold_units_unforseen1 = sold_units_unforseen
            remaining_units_unforseen1 = remaining_units_unforseen

            total_estimated_interest1 = total_estimated_interest
            interest_paid_to_date1 = interest_paid_to_date

        if index == 2:
            month_2_received = month_received
            month_2released = month_released
            capital_not_raised2 = capital_not_raised
            month_2mom = month_mom
            count_received2 = count_received
            count_exited2 = count_exited
            count_units2 = count_units
            value_still_to_be_raised2 = value_still_to_be_raised

            value_exited2 = value_exited
            current_investor_capital_deployed2 = current_investor_capital_deployed
            transferred_units2 = transferred_units
            sold_units2 = sold_units
            remaining_units2 = remaining_units
            transferred_units_sold_value2 = transferred_units_sold_value
            sold_units_value2 = sold_units_value
            remaining_units_value2 = remaining_units_value
            total_sale_value2 = total_sale_value
            total_units_sales_value2 = total_units_sales_value

            total_units_commission2 = total_units_commission
            transferred_units_commission2 = transferred_units_commission
            sold_units_commission2 = sold_units_commission
            remaining_units_commission2 = remaining_units_commission

            total_units_transfer_fees2 = total_units_transfer_fees
            transferred_units_transfer_fees2 = transferred_units_transfer_fees
            sold_units_transfer_fees2 = sold_units_transfer_fees
            remaining_units_transfer_fees2 = remaining_units_transfer_fees

            total_units_bond_registration2 = total_units_bond_registration
            transferred_units_bond_registration2 = transferred_units_bond_registration
            sold_units_bond_registration2 = sold_units_bond_registration
            remaining_units_bond_registration2 = remaining_units_bond_registration

            total_units_trust_release_fee2 = total_units_trust_release_fee
            transferred_units_trust_release_fee2 = transferred_units_trust_release_fee
            sold_units_trust_release_fee2 = sold_units_trust_release_fee
            remaining_units_trust_release_fee2 = remaining_units_trust_release_fee

            total_units_unforseen2 = total_units_unforseen
            transferred_units_unforseen2 = transferred_units_unforseen
            sold_units_unforseen2 = sold_units_unforseen
            remaining_units_unforseen2 = remaining_units_unforseen

            interest_paid_to_date2 = interest_paid_to_date
            total_estimated_interest2 = total_estimated_interest

        if index == 3:
            month_3_received = month_received
            month_3released = month_released
            capital_not_raised3 = capital_not_raised
            month_3mom = month_mom
            count_received3 = count_received
            count_exited3 = count_exited
            count_units3 = count_units
            value_still_to_be_raised3 = value_still_to_be_raised

            value_exited3 = value_exited
            current_investor_capital_deployed3 = current_investor_capital_deployed
            transferred_units3 = transferred_units
            sold_units3 = sold_units
            remaining_units3 = remaining_units
            transferred_units_sold_value3 = transferred_units_sold_value
            sold_units_value3 = sold_units_value
            remaining_units_value3 = remaining_units_value
            total_sale_value3 = total_sale_value
            total_units_sales_value3 = total_units_sales_value

            total_units_commission3 = total_units_commission
            transferred_units_commission3 = transferred_units_commission
            sold_units_commission3 = sold_units_commission
            remaining_units_commission3 = remaining_units_commission

            total_units_transfer_fees3 = total_units_transfer_fees
            transferred_units_transfer_fees3 = transferred_units_transfer_fees
            sold_units_transfer_fees3 = sold_units_transfer_fees
            remaining_units_transfer_fees3 = remaining_units_transfer_fees

            total_units_bond_registration3 = total_units_bond_registration
            transferred_units_bond_registration3 = transferred_units_bond_registration
            sold_units_bond_registration3 = sold_units_bond_registration
            remaining_units_bond_registration3 = remaining_units_bond_registration

            total_units_trust_release_fee3 = total_units_trust_release_fee
            transferred_units_trust_release_fee3 = transferred_units_trust_release_fee
            sold_units_trust_release_fee3 = sold_units_trust_release_fee
            remaining_units_trust_release_fee3 = remaining_units_trust_release_fee

            total_units_unforseen3 = total_units_unforseen
            transferred_units_unforseen3 = transferred_units_unforseen
            sold_units_unforseen3 = sold_units_unforseen
            remaining_units_unforseen3 = remaining_units_unforseen

            interest_paid_to_date3 = interest_paid_to_date
            total_estimated_interest3 = total_estimated_interest

        if index == 4:
            month_4_received = month_received
            month_4released = month_released
            capital_not_raised4 = capital_not_raised
            month_4mom = month_mom
            count_received4 = count_received
            count_exited4 = count_exited
            count_units4 = count_units
            value_still_to_be_raised4 = value_still_to_be_raised

            value_exited4 = value_exited
            current_investor_capital_deployed4 = current_investor_capital_deployed
            transferred_units4 = transferred_units
            sold_units4 = sold_units
            remaining_units4 = remaining_units
            transferred_units_sold_value4 = transferred_units_sold_value
            sold_units_value4 = sold_units_value
            remaining_units_value4 = remaining_units_value
            total_sale_value4 = total_sale_value
            total_units_sales_value4 = total_units_sales_value

            total_units_commission4 = total_units_commission
            transferred_units_commission4 = transferred_units_commission
            sold_units_commission4 = sold_units_commission
            remaining_units_commission4 = remaining_units_commission

            total_units_transfer_fees4 = total_units_transfer_fees
            transferred_units_transfer_fees4 = transferred_units_transfer_fees
            sold_units_transfer_fees4 = sold_units_transfer_fees
            remaining_units_transfer_fees4 = remaining_units_transfer_fees

            total_units_bond_registration4 = total_units_bond_registration
            transferred_units_bond_registration4 = transferred_units_bond_registration
            sold_units_bond_registration4 = sold_units_bond_registration
            remaining_units_bond_registration4 = remaining_units_bond_registration

            total_units_trust_release_fee4 = total_units_trust_release_fee
            transferred_units_trust_release_fee4 = transferred_units_trust_release_fee
            sold_units_trust_release_fee4 = sold_units_trust_release_fee
            remaining_units_trust_release_fee4 = remaining_units_trust_release_fee

            total_units_unforseen4 = total_units_unforseen
            transferred_units_unforseen4 = transferred_units_unforseen
            sold_units_unforseen4 = sold_units_unforseen
            remaining_units_unforseen4 = remaining_units_unforseen

            interest_paid_to_date4 = interest_paid_to_date
            total_estimated_interest4 = total_estimated_interest

    # filter investors where trust array is not empty
    investors = list(filter(lambda x: len(x['trust']) > 0, investors))

    row1 = ["", "", "", "", "C.3_d", "", "", "", "", "C.3_d", "", "", "", "", "C.3_d", "", "", "", "", "C.3_d"]
    row2 = ["NSST HERON PROJECT REPORT", "", "", "", "", "NSST HERON PROJECT REPORT", "", "", "", "",
            "NSST HERON PROJECT REPORT", "", "", "", "", "NSST HERON PROJECT REPORT", "", "", "", ""]
    row3 = ["Report Date", dates[3], "", "", "", "Report Date", dates[2], "", "", "", "Report Date", dates[1], "",
            "", "", "Report Date", dates[0], "", "", ""]
    row4 = ["Development", "Heron Fields and Heron View", "", "", "", "Development", "Heron Fields and Heron View",
            "", "", "", "Development", "Heron Fields and Heron View", "", "", "", "Development",
            "Heron Fields and Heron View", "", "", ""]
    row5 = ["CAPITAL", "", "", "", "", "CAPITAL", "", "", "", "", "CAPITAL", "", "", "", "", "CAPITAL", "", "", "",
            ""]
    row6 = ["Total Investment capital to be raised (Estimated)", opportunity_amount_required, "", "", "",
            "Total Investment capital to be raised (Estimated)", opportunity_amount_required, "", "", "",
            "Total Investment capital to be raised (Estimated)", opportunity_amount_required, "", "", "",
            "Total Investment capital to be raised (Estimated)", opportunity_amount_required, "", "", ""]
    row7 = ["Total Investment capital received", month_4_received, "", "", "", "Total Investment capital received",
            month_3_received, "", "", "", "Total Investment capital received", month_2_received, "", "", "",
            "Total Investment capital received", month_1_received, "", "", ""]
    row8 = ["Total Funds Drawn Down into Development", month_4released, "", "", "",
            "Total Funds Drawn Down into Development", month_3released, "", "", "",
            "Total Funds Drawn Down into Development", month_2released, "", "", "",
            "Total Funds Drawn Down into Development", month_1released, "", "", ""]
    row9 = ["Momentum Investment Account", month_4mom, "", "", "", "Momentum Investment Account", month_3mom, "",
            "", "", "Momentum Investment Account", month_2mom, "", "", "", "Momentum Investment Account",
            month_1mom, "", "", ""]
    row10 = ["Capital not Raised", capital_not_raised4, "", "", "", "Capital not Raised", capital_not_raised3, "",
             "", "", "Capital not Raised", capital_not_raised2, "", "", "", "Capital not Raised",
             capital_not_raised1, "", "", ""]
    row11 = ["Available to be raised (Estimated)", value_still_to_be_raised4, "", "", "",
             "Available to be raised (Estimated)", value_still_to_be_raised3, "", "", "",
             "Available to be raised (Estimated)", value_still_to_be_raised2, "", "", "",
             "Available to be raised (Estimated)", value_still_to_be_raised1, "", "", ""]
    row12 = ["Capital repaid", value_exited4, "", "", "", "Capital repaid", value_exited3, "", "", "",
             "Capital repaid", value_exited2, "", "", "", "Capital repaid", value_exited1, "", "", ""]
    row13 = ["Current Investor Capital deployed", current_investor_capital_deployed4, "", "", "",
             "Current Investor Capital deployed", current_investor_capital_deployed3, "", "", "",
             "Current Investor Capital deployed", current_investor_capital_deployed2, "", "", "",
             "Current Investor Capital deployed", current_investor_capital_deployed1, "", "", ""]
    row14 = ["INVESTMENTS", "", "", "", "", "INVESTMENTS", "", "", "", "", "INVESTMENTS", "", "", "", "",
             "INVESTMENTS", "", "", "", ""]
    row15 = ["No. of Capital Investments received", count_received4, "", "", "",
             "No. of Capital Investments received", count_received3, "", "", "",
             "No. of Capital Investments received", count_received2, "", "", "",
             "No. of Capital Investments received", count_received1, "", "", ""]
    row16 = ["No. Investments exited to date", count_exited4, "", "", "", "No. Investments exited to date",
             count_exited3, "", "", "", "No. Investments exited to date", count_exited2, "", "", "",
             "No. Investments exited to date", count_exited1, "", "", ""]
    row17 = ["No. Investments still in Development", count_received4 - count_exited4, "", "", "",
             "No. Investments still in Development", count_received3 - count_exited3, "", "", "",
             "No. Investments still in Development", count_received2 - count_exited2, "", "", "",
             "No. Investments still in Development", count_received1 - count_exited1, "", "", ""]
    row18 = ["SALES INCOME", "", "", "", "", "SALES INCOME", "", "", "", "", "SALES INCOME", "", "", "", "",
             "SALES INCOME", "", "", "", ""]
    row19 = ["", "Total", "Transferred", "Sold", "Remaining", "", "Total", "Transferred", "Sold", "Remaining", "",
             "Total", "Transferred", "Sold", "Remaining", "", "Total", "Transferred", "Sold", "Remaining"]
    row20 = ["Units", count_units4, transferred_units4, sold_units4, remaining_units4, "Units", count_units3,
             transferred_units3, sold_units3, remaining_units3, "Units", count_units2, transferred_units2,
             sold_units2, remaining_units2, "Units", count_units1, transferred_units1, sold_units1,
             remaining_units1]
    row21 = ["Sales Income", total_units_sales_value4, transferred_units_sold_value4, sold_units_value4,
             remaining_units_value4, "Sales Income", total_units_sales_value3, transferred_units_sold_value3,
             sold_units_value3, remaining_units_value3, "Sales Income", total_units_sales_value2,
             transferred_units_sold_value2, sold_units_value2, remaining_units_value2, "Sales Income",
             total_units_sales_value1, transferred_units_sold_value1, sold_units_value1, remaining_units_value1]
    row22 = ["Commission", total_units_commission4, transferred_units_commission4, sold_units_commission4,
             remaining_units_commission4, "Commission", total_units_commission3, transferred_units_commission3,
             sold_units_commission3, remaining_units_commission3, "Commission", total_units_commission2,
             transferred_units_commission2, sold_units_commission2, remaining_units_commission2, "Commission",
             total_units_commission1, transferred_units_commission1, sold_units_commission1,
             remaining_units_commission1]
    row23 = ["Transfer Fees", total_units_transfer_fees4, transferred_units_transfer_fees4,
             sold_units_transfer_fees4, remaining_units_transfer_fees4, "Transfer Fees", total_units_transfer_fees3,
             transferred_units_transfer_fees3, sold_units_transfer_fees3, remaining_units_transfer_fees3,
             "Transfer Fees", total_units_transfer_fees2, transferred_units_transfer_fees2,
             sold_units_transfer_fees2, remaining_units_transfer_fees2, "Transfer Fees", total_units_transfer_fees1,
             transferred_units_transfer_fees1, sold_units_transfer_fees1, remaining_units_transfer_fees1]
    row24 = ["Bond Registration", total_units_bond_registration4, transferred_units_bond_registration4,
             sold_units_bond_registration4, remaining_units_bond_registration4, "Bond Registration",
             total_units_bond_registration3, transferred_units_bond_registration3, sold_units_bond_registration3,
             remaining_units_bond_registration3, "Bond Registration", total_units_bond_registration2,
             transferred_units_bond_registration2, sold_units_bond_registration2,
             remaining_units_bond_registration2, "Bond Registration", total_units_bond_registration1,
             transferred_units_bond_registration1, sold_units_bond_registration1,
             remaining_units_bond_registration1]
    row25 = ["Security Release Fee", total_units_trust_release_fee4, transferred_units_trust_release_fee4,
             sold_units_trust_release_fee4, remaining_units_trust_release_fee4, "Security Release Fee",
             total_units_trust_release_fee3, transferred_units_trust_release_fee3, sold_units_trust_release_fee3,
             remaining_units_trust_release_fee3, "Security Release Fee", total_units_trust_release_fee2,
             transferred_units_trust_release_fee2, sold_units_trust_release_fee2,
             remaining_units_trust_release_fee2, "Security Release Fee", total_units_trust_release_fee1,
             transferred_units_trust_release_fee1, sold_units_trust_release_fee1,
             remaining_units_trust_release_fee1]
    row26 = ["Unforseen (0.05%)", total_units_unforseen4, transferred_units_unforseen4, sold_units_unforseen4,
             remaining_units_unforseen4, "Unforseen", total_units_unforseen3, transferred_units_unforseen3,
             sold_units_unforseen3, remaining_units_unforseen3, "Unforseen", total_units_unforseen2,
             transferred_units_unforseen2, sold_units_unforseen2, remaining_units_unforseen2, "Unforseen",
             total_units_unforseen1, transferred_units_unforseen1, sold_units_unforseen1,
             remaining_units_unforseen1]
    row27 = ["Transfer Income", 0, 0, 0, 0, "Transfer Income", 0, 0, 0, 0, "Transfer Income", 0, 0, 0, 0,
             "Transfer Income", 0, 0, 0, 0]
    row28 = ["INTEREST", "", "", "", "", "INTEREST", "", "", "", "", "INTEREST", "", "", "", "", "INTEREST", "", "",
             "", ""]
    row29 = ["Total Estimated Interest", f"={total_estimated_interest4}+C48", "", "", "",
             "Total Estimated Interest",
             total_estimated_interest3, "", "", "", "Total Estimated Interest", total_estimated_interest2, "", "",
             "", "Total Estimated Interest", total_estimated_interest1, "", "", ""]
    row30 = ["Interest Paid to Date", interest_paid_to_date4, "", "", "", "Interest Paid to Date",
             interest_paid_to_date3, "", "", "", "Interest Paid to Date", interest_paid_to_date2, "", "", "",
             "Interest Paid to Date", interest_paid_to_date1, "", "", ""]
    row31 = ["FUNDING AVAILABLE", "", "", "", "", "FUNDING AVAILABLE", "", "", "", "", "FUNDING AVAILABLE", "", "",
             "", "", "FUNDING AVAILABLE", "", "", "", ""]
    row32 = ["Total Draw funds available", 0, "", "", "", "Total Draw funds available", 0, "", "", "",
             "Total Draw funds available", 0, "", "", "", "Total Draw funds available", 0, "", "", ""]
    row33 = ["Projected Heron Projects Income", 0, "", "", "", "Projected Heron Projects Income", 0, "", "", "",
             "Projected Heron Projects Income", 0, "", "", "", "Projected Heron Projects Income", 0, "", "", ""]
    row34 = ["Total Funds Available", 0, "", "", "", "Total Funds Available", 0, "", "", "",
             "Total Funds Available", 0, "", "", "", "Total Funds Available", 0, "", "", ""]
    row35 = ["Total Cost to Complete", 0, "", "", "", "Total Cost to Complete", 0, "", "", "",
             "Total Cost to Complete", 0, "", "", "", "Total Cost to Complete", 0, "", "", ""]
    row36 = ["Total funds (required)/Surplus", 0, "", "", "", "Total funds (required)/Surplus", 0, "", "", "",
             "Total funds (required)/Surplus", 0, "", "", "", "Total funds (required)/Surplus", 0, "", "", ""]
    row37 = ["PROJECTED PROFIT", "", "", "", "", "PROJECTED PROFIT", "", "", "", "", "PROJECTED PROFIT", "", "", "",
             "", "PROJECTED PROFIT", "", "", "", ""]
    row38 = ["Projected Nett Revenue", total_units_sales_value4, "", "", "", "Projected Nett Revenue",
             total_units_sales_value3, "", "", "", "Projected Nett Revenue", total_units_sales_value2, "", "", "",
             "Projected Nett Revenue", total_units_sales_value1, "", "", ""]
    row39 = ["Other Income (interest received)", 0, "", "", "", "Other Income (interest received)", 0, "", "", "",
             "Other Income (interest received)", 0, "", "", "", "Other Income (interest received)", 0, "", "", ""]
    row40 = ["Total Estimated Development Cost", 0, "", "", "", "Total Estimated Development Cost", 0, "", "", "",
             "Total Estimated Development Cost", 0, "", "", "", "Total Estimated Development Cost", 0, "", "", ""]
    row41 = ["Interest Expense", 0, "", "", "", "Interest Expense", 0, "", "", "", "Interest Expense", 0, "", "",
             "", "Interest Expense", 0, "", "", ""]
    row42 = ["Profit", 0, "", "", "", "Profit", 0, "", "", "", "Profit", 0, "", "", "", "Profit", 0, "", "", ""]
    row43 = ["Sales Increase", 0, 0, 0]
    row44 = ["CPC Construction", 0, 0, 0]
    row45 = ["Rent Salaries and wages", 0, 0, 0]
    row46 = ["CPSD", 0, 0, 0]
    row47 = ["Opp Invest", 0, 0, 0]
    row48 = ["investor interest", 0, 0, 0]
    row49 = ["Commissions", 0, 0, 0]
    row50 = ["Unforseen", 0, 0, 0]
    row51 = ["Cash to flow to Heron from Quinate early exits", 0, 0, 0]

    # row44 = ["", 0, "CPC Construction", 0, "", "", 0, "CPC Construction", 0, "", "", 0, "CPC Construction", 0, "",
    #          "", 0, "CPC Construction", 0, ""]
    # row45 = ["", 0, " Rent Salaries and wages", 0, "", "", 0, " Rent Salaries and wages", 0, "", "", 0,
    #          " Rent Salaries and wages", 0, "", "", 0, " Rent Salaries and wages", 0, "", ""]
    # row46 = ["", 0, "CPSD", 0, "", "", 0, "CPSD", 0, "", "", 0, "CPSD", 0, "", "", 0, "CPSD", 0, "", ""]
    # row47 = ["", 0, "Opp Invest", 0, "", "", 0, "Opp Invest", 0, "", "", 0, "Opp Invest", 0, "", "", 0,
    #          "Opp Invest", 0, "", ""]
    # row48 = ["", 0, "investor interest", 0, "", "", 0, "investor interest", 0, "", "", 0, "investor interest", 0,
    #          "", "", 0, "investor interest", 0, ""]
    # row49 = ["", 0, "Commissions", 0, "", "", 0, "Commissions", 0, "", "", 0, "Commissions", 0, "", "", 0,
    #          "Commissions", 0, ""]
    # row50 = ["", 0, "Unforseen", 0, "", "", 0, "Unforseen", 0, "", "", 0, "Unforseen", 0, "", "", 0, "Unforseen", 0,
    #          "", ""]
    # row51 = ["", 0, "Cash to flow to Heron from Quinate early exits", "", 0, "", 0,
    #          "Cash to flow to Heron from Quinate early exits", "", 0, "", 0,
    #          "Cash to flow to Heron from Quinate early exits", "", 0, "", 0,
    #          "Cash to flow to Heron from Quinate early exits", "", 0]
    row52 = ["", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""]
    row53 = ["Sales Income", total_units_sales_value4, transferred_units_sold_value4, sold_units_value4,
             remaining_units_value4, "Sales Income", total_units_sales_value3, transferred_units_sold_value3,
             sold_units_value3, remaining_units_value3, "Sales Income", total_units_sales_value2,
             transferred_units_sold_value2, sold_units_value2, remaining_units_value2, "Sales Income",
             total_units_sales_value1, transferred_units_sold_value1, sold_units_value1, remaining_units_value1]

    # append all rows to nsst_print_report
    nsss_print_report = [row1, row2, row3, row4, row5, row6, row7, row8, row9, row10, row11,
                         row12, row13, row14, row15, row16, row17, row18, row19, row20,
                         row21, row22, row23, row24, row25, row26, row27, row28, row29,
                         row30, row31, row32, row33, row34, row35, row36, row37, row38, row39, row40, row41, row42,
                         row43, row44, row45, row46, row47, row48, row49, row50, row51, row52, row53]

    # print(len(investors))
    report_date = request['to_date']

    # CREATE SPREADSHEET

    # print("profit_and_loss", profit_and_loss)
    result = cashflow_hf_hv(profit_and_loss, nsss_print_report, report_date)
    print("result", result)

    end_time = time.time()
    return {"Success": True, "time_taken": end_time - start_time, "file": "cashflow_p&l_files/cashflow_hf_hv.xlsx"}

    # except Exception as e:
    #     print(e)
    #     return {"ERROR": "Please Try again"}


@xero.get("/get_profit_and_loss_data")
def get_profit_and_loss_data():
    try:
        # print("Hello")
        # get all profit and loss from the db
        profit_and_loss = list(db.profit_and_loss.find({}))
        for item in profit_and_loss:
            item['_id'] = str(item['_id'])
        # print(profit_and_loss)
        # print(len(profit_and_loss))
        return {"Success": True, "profit_and_loss": profit_and_loss}
    except Exception as e:
        print(e)
        return {"ERROR": "Please Try again"}


@xero.post("/update_profit_and_loss")
async def update_profit_and_loss(data: Request):
    request = await data.json()
    request = request['data']
    # print("request", request)
    try:
        id = request['_id']
        del request['_id']
        # update profit and loss with request
        db.profit_and_loss.update_one({"_id": ObjectId(id)}, {"$set": request})
        return {"status": "Success"}
    except Exception as e:
        print(e)
        return {"status": "Please Try again"}


@xero.post("/insert_profit_and_loss")
async def insert_profit_and_loss(data: Request):
    request = await data.json()
    request = request['data']
    # print("request", request)

    try:
        # insert profit and loss with request
        db.profit_and_loss.insert_one(request)
        return {"status": "Success"}
    except Exception as e:
        print(e)
        return {"status": "Please Try again"}


@xero.get("/get_profit_and_loss_report")
async def get_profit_and_loss_report(profit_and_loss_name):
    file_name = profit_and_loss_name
    # print("file_name", file_name)
    dir_path = "cashflow_p&l_files"
    dir_list = os.listdir(dir_path)
    # print("dir_list", dir_list)
    if file_name in dir_list:
        # print("file exists")
        return FileResponse(f"{dir_path}/{file_name}", filename=file_name)
    else:
        return {"ERROR": "File does not exist!!"}


# def get_json_file():
#     import json
#
#     # Specify the path to your JSON file
#     file_path = 'HeronFieldsP&L2023A.json'
#
#     # Open the JSON file for reading
#     try:
#         with open(file_path, 'r') as json_file:
#             data = json.load(json_file)
#     except (FileNotFoundError, json.JSONDecodeError) as e:
#         print(f"Error opening or parsing the JSON file: {e}")
#
#     try:
#         # input data into profit_and_loss in the db
#         db.profit_and_loss.insert_many(data)
#         print("Data inserted successfully")
#     except Exception as e:
#         print(e)
#         return {"ERROR": "Please Try again"}
#
#     print(len(data))
#
#
# # get_json_file()


# TEMP FUNCTION TO UPDATE PROFIT AND LOSS
def update_profit_and_loss_from_cf_file():
    input_p_and_l = [
        {"Account": "Sales - Heron Fields", "Category": "Trading Income", "Development": "Heron Fields""Heron Fields",
         "Applicable_dev": "Heron Fields", "Month": "Apr-24", "Actual": 0, "Forecast": 1217304.34782609},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron Fields",
         "Applicable_dev": "Heron Fields", "Month": "Apr-24", "Actual": 0, "Forecast": 39515.45},
        {"Account": "COS - Commission HF Units", "Category": "COS", "Development": "Heron Fields",
         "Applicable_dev": "Heron Fields", "Month": "Apr-24", "Actual": 0, "Forecast": 60865.2173913044},
        {"Account": "Sales - Heron Fields", "Category": "Trading Income", "Development": "Heron Fields""Heron Fields",
         "Applicable_dev": "Heron Fields", "Month": "Apr-24", "Actual": 0, "Forecast": 1260782.60869565},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron Fields",
         "Applicable_dev": "Heron Fields", "Month": "Apr-24", "Actual": 0, "Forecast": 39515.45},
        {"Account": "COS - Commission HF Units", "Category": "COS", "Development": "Heron Fields",
         "Applicable_dev": "Heron Fields", "Month": "Apr-24", "Actual": 0, "Forecast": 63039.1304347826},
        {"Account": "Sales - Heron Fields", "Category": "Trading Income", "Development": "Heron Fields""Heron Fields",
         "Applicable_dev": "Heron Fields", "Month": "Apr-24", "Actual": 0, "Forecast": 1304260.86956522},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron Fields",
         "Applicable_dev": "Heron Fields", "Month": "Apr-24", "Actual": 0, "Forecast": 39515.45},
        {"Account": "COS - Commission HF Units", "Category": "COS", "Development": "Heron Fields",
         "Applicable_dev": "Heron Fields", "Month": "Apr-24", "Actual": 0, "Forecast": 65213.0434782609},
        {"Account": "Sales - Heron Fields", "Category": "Trading Income", "Development": "Heron Fields""Heron Fields",
         "Applicable_dev": "Heron Fields", "Month": "Apr-24", "Actual": 0, "Forecast": 1304260.86956522},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron Fields",
         "Applicable_dev": "Heron Fields", "Month": "Apr-24", "Actual": 0, "Forecast": 39515.45},
        {"Account": "COS - Commission HF Units", "Category": "COS", "Development": "Heron Fields",
         "Applicable_dev": "Heron Fields", "Month": "Apr-24", "Actual": 0, "Forecast": 65213.0434782609},
        {"Account": "Sales - Heron Fields", "Category": "Trading Income", "Development": "Heron Fields""Heron Fields",
         "Applicable_dev": "Heron Fields", "Month": "Apr-24", "Actual": 0, "Forecast": 1304260.86956522},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron Fields",
         "Applicable_dev": "Heron Fields", "Month": "Apr-24", "Actual": 0, "Forecast": 39515.45},
        {"Account": "COS - Commission HF Units", "Category": "COS", "Development": "Heron Fields",
         "Applicable_dev": "Heron Fields", "Month": "Apr-24", "Actual": 0, "Forecast": 65213.0434782609},
        {"Account": "Sales - Heron Fields", "Category": "Trading Income", "Development": "Heron Fields""Heron Fields",
         "Applicable_dev": "Heron Fields", "Month": "Apr-24", "Actual": 0, "Forecast": 1252086.95652174},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron Fields",
         "Applicable_dev": "Heron Fields", "Month": "Apr-24", "Actual": 0, "Forecast": 39515.45},
        {"Account": "COS - Commission HF Units", "Category": "COS", "Development": "Heron Fields",
         "Applicable_dev": "Heron Fields", "Month": "Apr-24", "Actual": 0, "Forecast": 62604.347826087},
        {"Account": "Sales - Heron Fields", "Category": "Trading Income", "Development": "Heron Fields""Heron Fields",
         "Applicable_dev": "Heron Fields", "Month": "Apr-24", "Actual": 0, "Forecast": 1260782.60869565},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron Fields",
         "Applicable_dev": "Heron Fields", "Month": "Apr-24", "Actual": 0, "Forecast": 39515.45},
        {"Account": "COS - Commission HF Units", "Category": "COS", "Development": "Heron Fields",
         "Applicable_dev": "Heron Fields", "Month": "Apr-24", "Actual": 0, "Forecast": 63039.1304347826},
        {"Account": "Sales - Heron Fields", "Category": "Trading Income", "Development": "Heron Fields""Heron Fields",
         "Applicable_dev": "Heron Fields", "Month": "Apr-24", "Actual": 0, "Forecast": 1260782.60869565},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron Fields",
         "Applicable_dev": "Heron Fields", "Month": "Apr-24", "Actual": 0, "Forecast": 39515.45},
        {"Account": "COS - Commission HF Units", "Category": "COS", "Development": "Heron Fields",
         "Applicable_dev": "Heron Fields", "Month": "Apr-24", "Actual": 0, "Forecast": 63039.1304347826},
        {"Account": "Sales - Heron View Sales", "Category": "Trading Income", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "May-24", "Actual": 0, "Forecast": 1304260.86956522},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron View", "Applicable_dev": "Heron View",
         "Month": "May-24", "Actual": 0, "Forecast": 39515.45},
        {"Account": "COS - Commission HV Units", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "May-24", "Actual": 0, "Forecast": 65213.0434782609},
        {"Account": "Sales - Heron View Sales", "Category": "Trading Income", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "May-24", "Actual": 0, "Forecast": 1243391.30434783},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron View", "Applicable_dev": "Heron View",
         "Month": "May-24", "Actual": 0, "Forecast": 39515.45},
        {"Account": "COS - Commission HV Units", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "May-24", "Actual": 0, "Forecast": 62169.5652173913},
        {"Account": "Sales - Heron View Sales", "Category": "Trading Income", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Mar-24", "Actual": 0, "Forecast": 1295565.2173913},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron View", "Applicable_dev": "Heron View",
         "Month": "Mar-24", "Actual": 0, "Forecast": 1789},
        {"Account": "COS - Commission HV Units", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Mar-24", "Actual": 0, "Forecast": 64778.2608695652},
        {"Account": "Sales - Heron View Sales", "Category": "Trading Income", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Mar-24", "Actual": 0, "Forecast": 1312956.52173913},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron View", "Applicable_dev": "Heron View",
         "Month": "Mar-24", "Actual": 0, "Forecast": 1789},
        {"Account": "COS - Commission HV Units", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Mar-24", "Actual": 0, "Forecast": 65647.8260869565},
        {"Account": "Sales - Heron View Sales", "Category": "Trading Income", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "May-24", "Actual": 0, "Forecast": 1312956.52173913},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron View", "Applicable_dev": "Heron View",
         "Month": "May-24", "Actual": 0, "Forecast": 39515.45},
        {"Account": "COS - Commission HV Units", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "May-24", "Actual": 0, "Forecast": 65647.8260869565},
        {"Account": "Sales - Heron View Sales", "Category": "Trading Income", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Jun-24", "Actual": 0, "Forecast": 1278173.91304348},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron View", "Applicable_dev": "Heron View",
         "Month": "Jun-24", "Actual": 0, "Forecast": 39515.45},
        {"Account": "COS - Commission HV Units", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Jun-24", "Actual": 0, "Forecast": 63908.6956521739},
        {"Account": "Sales - Heron View Sales", "Category": "Trading Income", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Jun-24", "Actual": 0, "Forecast": 1312956.52173913},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron View", "Applicable_dev": "Heron View",
         "Month": "Jun-24", "Actual": 0, "Forecast": 39515.45},
        {"Account": "COS - Commission HV Units", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Jun-24", "Actual": 0, "Forecast": 65647.8260869565},
        {"Account": "Sales - Heron View Sales", "Category": "Trading Income", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Jun-24", "Actual": 0, "Forecast": 1321652.17391304},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron View", "Applicable_dev": "Heron View",
         "Month": "Jun-24", "Actual": 0, "Forecast": 39515.45},
        {"Account": "COS - Commission HV Units", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Jun-24", "Actual": 0, "Forecast": 66082.6086956522},
        {"Account": "Sales - Heron View Sales", "Category": "Trading Income", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Feb-24", "Actual": 0, "Forecast": 1478173.91304348},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron View", "Applicable_dev": "Heron View",
         "Month": "Feb-24", "Actual": 0, "Forecast": 39515.45},
        {"Account": "COS - Commission HV Units", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Feb-24", "Actual": 0, "Forecast": 73908.6956521739},
        {"Account": "Sales - Heron View Sales", "Category": "Trading Income", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Jun-24", "Actual": 0, "Forecast": 1321652.17391304},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron View", "Applicable_dev": "Heron View",
         "Month": "Jun-24", "Actual": 0, "Forecast": 39515.45},
        {"Account": "COS - Commission HV Units", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Jun-24", "Actual": 0, "Forecast": 66082.6086956522},
        {"Account": "Sales - Heron View Sales", "Category": "Trading Income", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Jun-24", "Actual": 0, "Forecast": 1321652.17391304},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron View", "Applicable_dev": "Heron View",
         "Month": "Jun-24", "Actual": 0, "Forecast": 1789},
        {"Account": "COS - Commission HV Units", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Jun-24", "Actual": 0, "Forecast": 66082.6086956522},
        {"Account": "Sales - Heron View Sales", "Category": "Trading Income", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Feb-24", "Actual": 0, "Forecast": 1391217.39130435},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron View", "Applicable_dev": "Heron View",
         "Month": "Feb-24", "Actual": 0, "Forecast": 39515.45},
        {"Account": "COS - Commission HV Units", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Feb-24", "Actual": 0, "Forecast": 69560.8695652174},
        {"Account": "Sales - Heron View Sales", "Category": "Trading Income", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Jun-24", "Actual": 0, "Forecast": 1339043.47826087},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron View", "Applicable_dev": "Heron View",
         "Month": "Jun-24", "Actual": 0, "Forecast": 39515.45},
        {"Account": "COS - Commission HV Units", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Jun-24", "Actual": 0, "Forecast": 66952.1739130435},
        {"Account": "Sales - Heron View Sales", "Category": "Trading Income", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Jun-24", "Actual": 0, "Forecast": 1339043.47826087},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron View", "Applicable_dev": "Heron View",
         "Month": "Jun-24", "Actual": 0, "Forecast": 39515.45},
        {"Account": "COS - Commission HV Units", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Jun-24", "Actual": 0, "Forecast": 66952.1739130435},
        {"Account": "Sales - Heron View Sales", "Category": "Trading Income", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Jun-24", "Actual": 0, "Forecast": 1391217.39130435},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron View", "Applicable_dev": "Heron View",
         "Month": "Jun-24", "Actual": 0, "Forecast": 39515.45},
        {"Account": "COS - Commission HV Units", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Jun-24", "Actual": 0, "Forecast": 69560.8695652174},
        {"Account": "Sales - Heron View Sales", "Category": "Trading Income", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Jul-24", "Actual": 0, "Forecast": 1470433.91304348},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron View", "Applicable_dev": "Heron View",
         "Month": "Jul-24", "Actual": 0, "Forecast": 1789},
        {"Account": "COS - Commission HV Units", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Jul-24", "Actual": 0, "Forecast": 73521.6956521739},
        {"Account": "Sales - Heron View Sales", "Category": "Trading Income", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Jul-24", "Actual": 0, "Forecast": 1452173.04347826},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron View", "Applicable_dev": "Heron View",
         "Month": "Jul-24", "Actual": 0, "Forecast": 1789},
        {"Account": "COS - Commission HV Units", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Jul-24", "Actual": 0, "Forecast": 72608.6521739131},
        {"Account": "Sales - Heron View Sales", "Category": "Trading Income", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Jul-24", "Actual": 0, "Forecast": 1452173.04347826},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron View", "Applicable_dev": "Heron View",
         "Month": "Jul-24", "Actual": 0, "Forecast": 1789},
        {"Account": "COS - Commission HV Units", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Jul-24", "Actual": 0, "Forecast": 72608.6521739131},
        {"Account": "Sales - Heron View Sales", "Category": "Trading Income", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Jul-24", "Actual": 0, "Forecast": 1470433.91304348},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron View", "Applicable_dev": "Heron View",
         "Month": "Jul-24", "Actual": 0, "Forecast": 1789},
        {"Account": "COS - Commission HV Units", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Jul-24", "Actual": 0, "Forecast": 73521.6956521739},
        {"Account": "Sales - Heron View Sales", "Category": "Trading Income", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Jul-24", "Actual": 0, "Forecast": 1485216.52173913},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron View", "Applicable_dev": "Heron View",
         "Month": "Jul-24", "Actual": 0, "Forecast": 1789},
        {"Account": "COS - Commission HV Units", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Jul-24", "Actual": 0, "Forecast": 74260.8260869565},
        {"Account": "Sales - Heron View Sales", "Category": "Trading Income", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Jul-24", "Actual": 0, "Forecast": 1470433.91304348},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron View", "Applicable_dev": "Heron View",
         "Month": "Jul-24", "Actual": 0, "Forecast": 1789},
        {"Account": "COS - Commission HV Units", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Jul-24", "Actual": 0, "Forecast": 73521.6956521739},
        {"Account": "Sales - Heron View Sales", "Category": "Trading Income", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Jul-24", "Actual": 0, "Forecast": 1470433.91304348},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron View", "Applicable_dev": "Heron View",
         "Month": "Jul-24", "Actual": 0, "Forecast": 1789},
        {"Account": "COS - Commission HV Units", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Jul-24", "Actual": 0, "Forecast": 73521.6956521739},
        {"Account": "Sales - Heron View Sales", "Category": "Trading Income", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Jul-24", "Actual": 0, "Forecast": 1485216.52173913},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron View", "Applicable_dev": "Heron View",
         "Month": "Jul-24", "Actual": 0, "Forecast": 1789},
        {"Account": "COS - Commission HV Units", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Jul-24", "Actual": 0, "Forecast": 74260.8260869565},
        {"Account": "Sales - Heron View Sales", "Category": "Trading Income", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Jun-24", "Actual": 0, "Forecast": 1530347.82608696},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron View", "Applicable_dev": "Heron View",
         "Month": "Jun-24", "Actual": 0, "Forecast": 1789},
        {"Account": "COS - Commission HV Units", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Jun-24", "Actual": 0, "Forecast": 76517.3913043478},
        {"Account": "Sales - Heron View Sales", "Category": "Trading Income", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Jun-24", "Actual": 0, "Forecast": 1512956.52173913},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron View", "Applicable_dev": "Heron View",
         "Month": "Jun-24", "Actual": 0, "Forecast": 39515.45},
        {"Account": "COS - Commission HV Units", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Jun-24", "Actual": 0, "Forecast": 75647.8260869565},
        {"Account": "Sales - Heron View Sales", "Category": "Trading Income", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Jun-24", "Actual": 0, "Forecast": 1512956.52173913},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron View", "Applicable_dev": "Heron View",
         "Month": "Jun-24", "Actual": 0, "Forecast": 39515.45},
        {"Account": "COS - Commission HV Units", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Jun-24", "Actual": 0, "Forecast": 75647.8260869565},
        {"Account": "Sales - Heron View Sales", "Category": "Trading Income", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Jul-24", "Actual": 0, "Forecast": 1434695.65217391},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron View", "Applicable_dev": "Heron View",
         "Month": "Jul-24", "Actual": 0, "Forecast": 39515.45},
        {"Account": "COS - Commission HV Units", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Jul-24", "Actual": 0, "Forecast": 71734.7826086957},
        {"Account": "Sales - Heron View Sales", "Category": "Trading Income", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Jul-24", "Actual": 0, "Forecast": 1434695.65217391},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron View", "Applicable_dev": "Heron View",
         "Month": "Jul-24", "Actual": 0, "Forecast": 39515.45},
        {"Account": "COS - Commission HV Units", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Jul-24", "Actual": 0, "Forecast": 71734.7826086957},
        {"Account": "Sales - Heron View Sales", "Category": "Trading Income", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Jul-24", "Actual": 0, "Forecast": 1434695.65217391},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron View", "Applicable_dev": "Heron View",
         "Month": "Jul-24", "Actual": 0, "Forecast": 39515.45},
        {"Account": "COS - Commission HV Units", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Jul-24", "Actual": 0, "Forecast": 71734.7826086957},
        {"Account": "Sales - Heron View Sales", "Category": "Trading Income", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Aug-24", "Actual": 0, "Forecast": 1460782.60869565},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron View", "Applicable_dev": "Heron View",
         "Month": "Aug-24", "Actual": 0, "Forecast": 39515.45},
        {"Account": "COS - Commission HV Units", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Aug-24", "Actual": 0, "Forecast": 73039.1304347826},
        {"Account": "Sales - Heron View Sales", "Category": "Trading Income", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Aug-24", "Actual": 0, "Forecast": 1460782.60869565},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron View", "Applicable_dev": "Heron View",
         "Month": "Aug-24", "Actual": 0, "Forecast": 39515.45},
        {"Account": "COS - Commission HV Units", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Aug-24", "Actual": 0, "Forecast": 73039.1304347826},
        {"Account": "Sales - Heron View Sales", "Category": "Trading Income", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Aug-24", "Actual": 0, "Forecast": 1460782.60869565},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron View", "Applicable_dev": "Heron View",
         "Month": "Aug-24", "Actual": 0, "Forecast": 39515.45},
        {"Account": "COS - Commission HV Units", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Aug-24", "Actual": 0, "Forecast": 73039.1304347826},
        {"Account": "Sales - Heron View Sales", "Category": "Trading Income", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Sep-24", "Actual": 0, "Forecast": 1478173.91304348},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron View", "Applicable_dev": "Heron View",
         "Month": "Sep-24", "Actual": 0, "Forecast": 39515.45},
        {"Account": "COS - Commission HV Units", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Sep-24", "Actual": 0, "Forecast": 73908.6956521739},
        {"Account": "Sales - Heron View Sales", "Category": "Trading Income", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Sep-24", "Actual": 0, "Forecast": 1478173.91304348},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron View", "Applicable_dev": "Heron View",
         "Month": "Sep-24", "Actual": 0, "Forecast": 39515.45},
        {"Account": "COS - Commission HV Units", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Sep-24", "Actual": 0, "Forecast": 73908.6956521739},
        {"Account": "Sales - Heron View Sales", "Category": "Trading Income", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Sep-24", "Actual": 0, "Forecast": 1478173.91304348},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron View", "Applicable_dev": "Heron View",
         "Month": "Sep-24", "Actual": 0, "Forecast": 39515.45},
        {"Account": "COS - Commission HV Units", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Sep-24", "Actual": 0, "Forecast": 73908.6956521739},
        {"Account": "Sales - Heron View Sales", "Category": "Trading Income", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Sep-24", "Actual": 0, "Forecast": 1512956.52173913},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron View", "Applicable_dev": "Heron View",
         "Month": "Sep-24", "Actual": 0, "Forecast": 39515.45},
        {"Account": "COS - Commission HV Units", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Sep-24", "Actual": 0, "Forecast": 75647.8260869565},
        {"Account": "Sales - Heron View Sales", "Category": "Trading Income", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Sep-24", "Actual": 0, "Forecast": 1512956.52173913},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron View", "Applicable_dev": "Heron View",
         "Month": "Sep-24", "Actual": 0, "Forecast": 39515.45},
        {"Account": "COS - Commission HV Units", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Sep-24", "Actual": 0, "Forecast": 75647.8260869565},
        {"Account": "Sales - Heron View Sales", "Category": "Trading Income", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Sep-24", "Actual": 0, "Forecast": 1512956.52173913},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron View", "Applicable_dev": "Heron View",
         "Month": "Sep-24", "Actual": 0, "Forecast": 39515.45},
        {"Account": "COS - Commission HV Units", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Sep-24", "Actual": 0, "Forecast": 75647.8260869565},
        {"Account": "Sales - Heron View Sales", "Category": "Trading Income", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Sep-24", "Actual": 0, "Forecast": 1512956.52173913},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron View", "Applicable_dev": "Heron View",
         "Month": "Sep-24", "Actual": 0, "Forecast": 39515.45},
        {"Account": "COS - Commission HV Units", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Sep-24", "Actual": 0, "Forecast": 75647.8260869565},
        {"Account": "Sales - Heron View Sales", "Category": "Trading Income", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Sep-24", "Actual": 0, "Forecast": 1512956.52173913},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron View", "Applicable_dev": "Heron View",
         "Month": "Sep-24", "Actual": 0, "Forecast": 39515.45},
        {"Account": "COS - Commission HV Units", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Sep-24", "Actual": 0, "Forecast": 75647.8260869565},
        {"Account": "Sales - Heron View Sales", "Category": "Trading Income", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Sep-24", "Actual": 0, "Forecast": 1512956.52173913},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron View", "Applicable_dev": "Heron View",
         "Month": "Sep-24", "Actual": 0, "Forecast": 39515.45},
        {"Account": "COS - Commission HV Units", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Sep-24", "Actual": 0, "Forecast": 75647.8260869565},
        {"Account": "Sales - Heron View Sales", "Category": "Trading Income", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Oct-24", "Actual": 0, "Forecast": 1304260.86956522},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron View", "Applicable_dev": "Heron View",
         "Month": "Oct-24", "Actual": 0, "Forecast": 39515.45},
        {"Account": "COS - Commission HV Units", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Oct-24", "Actual": 0, "Forecast": 65213.0434782609},
        {"Account": "Sales - Heron View Sales", "Category": "Trading Income", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Oct-24", "Actual": 0, "Forecast": 1304260.86956522},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron View", "Applicable_dev": "Heron View",
         "Month": "Oct-24", "Actual": 0, "Forecast": 39515.45},
        {"Account": "COS - Commission HV Units", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Oct-24", "Actual": 0, "Forecast": 65213.0434782609},
        {"Account": "Sales - Heron View Sales", "Category": "Trading Income", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Oct-24", "Actual": 0, "Forecast": 1478173.91304348},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron View", "Applicable_dev": "Heron View",
         "Month": "Oct-24", "Actual": 0, "Forecast": 39515.45},
        {"Account": "COS - Commission HV Units", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Oct-24", "Actual": 0, "Forecast": 73908.6956521739},
        {"Account": "Sales - Heron View Sales", "Category": "Trading Income", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Oct-24", "Actual": 0, "Forecast": 1304260.86956522},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron View", "Applicable_dev": "Heron View",
         "Month": "Oct-24", "Actual": 0, "Forecast": 39515.45},
        {"Account": "COS - Commission HV Units", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Oct-24", "Actual": 0, "Forecast": 65213.0434782609},
        {"Account": "Sales - Heron View Sales", "Category": "Trading Income", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Oct-24", "Actual": 0, "Forecast": 1304260.86956522},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron View", "Applicable_dev": "Heron View",
         "Month": "Oct-24", "Actual": 0, "Forecast": 39515.45},
        {"Account": "COS - Commission HV Units", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Oct-24", "Actual": 0, "Forecast": 65213.0434782609},
        {"Account": "Sales - Heron View Sales", "Category": "Trading Income", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Oct-24", "Actual": 0, "Forecast": 1304260.86956522},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron View", "Applicable_dev": "Heron View",
         "Month": "Oct-24", "Actual": 0, "Forecast": 39515.45},
        {"Account": "COS - Commission HV Units", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Oct-24", "Actual": 0, "Forecast": 65213.0434782609},
        {"Account": "Sales - Heron View Sales", "Category": "Trading Income", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Nov-24", "Actual": 0, "Forecast": 1330347.82608696},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron View", "Applicable_dev": "Heron View",
         "Month": "Nov-24", "Actual": 0, "Forecast": 39515.45},
        {"Account": "COS - Commission HV Units", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Nov-24", "Actual": 0, "Forecast": 66517.3913043478},
        {"Account": "Sales - Heron View Sales", "Category": "Trading Income", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Nov-24", "Actual": 0, "Forecast": 1330347.82608696},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron View", "Applicable_dev": "Heron View",
         "Month": "Nov-24", "Actual": 0, "Forecast": 39515.45},
        {"Account": "COS - Commission HV Units", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Nov-24", "Actual": 0, "Forecast": 66517.3913043478},
        {"Account": "Sales - Heron View Sales", "Category": "Trading Income", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Nov-24", "Actual": 0, "Forecast": 1504260.86956522},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron View", "Applicable_dev": "Heron View",
         "Month": "Nov-24", "Actual": 0, "Forecast": 39515.45},
        {"Account": "COS - Commission HV Units", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Nov-24", "Actual": 0, "Forecast": 75213.0434782609},
        {"Account": "Sales - Heron View Sales", "Category": "Trading Income", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Nov-24", "Actual": 0, "Forecast": 1330347.82608696},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron View", "Applicable_dev": "Heron View",
         "Month": "Nov-24", "Actual": 0, "Forecast": 39515.45},
        {"Account": "COS - Commission HV Units", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Nov-24", "Actual": 0, "Forecast": 66517.3913043478},
        {"Account": "Sales - Heron View Sales", "Category": "Trading Income", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Nov-24", "Actual": 0, "Forecast": 1330347.82608696},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron View", "Applicable_dev": "Heron View",
         "Month": "Nov-24", "Actual": 0, "Forecast": 39515.45},
        {"Account": "COS - Commission HV Units", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Nov-24", "Actual": 0, "Forecast": 66517.3913043478},
        {"Account": "Sales - Heron View Sales", "Category": "Trading Income", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Nov-24", "Actual": 0, "Forecast": 1330347.82608696},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron View", "Applicable_dev": "Heron View",
         "Month": "Nov-24", "Actual": 0, "Forecast": 39515.45},
        {"Account": "COS - Commission HV Units", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Nov-24", "Actual": 0, "Forecast": 66517.3913043478},
        {"Account": "Sales - Heron View Sales", "Category": "Trading Income", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Dec-24", "Actual": 0, "Forecast": 1304260.86956522},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron View", "Applicable_dev": "Heron View",
         "Month": "Dec-24", "Actual": 0, "Forecast": 39515.45},
        {"Account": "COS - Commission HV Units", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Dec-24", "Actual": 0, "Forecast": 65213.0434782609},
        {"Account": "Sales - Heron View Sales", "Category": "Trading Income", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Dec-24", "Actual": 0, "Forecast": 1304260.86956522},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron View", "Applicable_dev": "Heron View",
         "Month": "Dec-24", "Actual": 0, "Forecast": 39515.45},
        {"Account": "COS - Commission HV Units", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Dec-24", "Actual": 0, "Forecast": 65213.0434782609},
        {"Account": "Sales - Heron View Sales", "Category": "Trading Income", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Dec-24", "Actual": 0, "Forecast": 1521652.17391304},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron View", "Applicable_dev": "Heron View",
         "Month": "Dec-24", "Actual": 0, "Forecast": 39515.45},
        {"Account": "COS - Commission HV Units", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Dec-24", "Actual": 0, "Forecast": 76082.6086956522},
        {"Account": "Sales - Heron View Sales", "Category": "Trading Income", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Dec-24", "Actual": 0, "Forecast": 1304260.86956522},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron View", "Applicable_dev": "Heron View",
         "Month": "Dec-24", "Actual": 0, "Forecast": 39515.45},
        {"Account": "COS - Commission HV Units", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Dec-24", "Actual": 0, "Forecast": 65213.0434782609},
        {"Account": "Sales - Heron View Sales", "Category": "Trading Income", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Dec-24", "Actual": 0, "Forecast": 1304260.86956522},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron View", "Applicable_dev": "Heron View",
         "Month": "Dec-24", "Actual": 0, "Forecast": 39515.45},
        {"Account": "COS - Commission HV Units", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Dec-24", "Actual": 0, "Forecast": 65213.0434782609},
        {"Account": "Sales - Heron View Sales", "Category": "Trading Income", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Dec-24", "Actual": 0, "Forecast": 1304260.86956522},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron View", "Applicable_dev": "Heron View",
         "Month": "Dec-24", "Actual": 0, "Forecast": 39515.45},
        {"Account": "COS - Commission HV Units", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Dec-24", "Actual": 0, "Forecast": 65213.0434782609},
        {"Account": "Sales - Heron View Sales", "Category": "Trading Income", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Jul-24", "Actual": 0, "Forecast": 1460782.60869565},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron View", "Applicable_dev": "Heron View",
         "Month": "Jul-24", "Actual": 0, "Forecast": 39515.45},
        {"Account": "COS - Commission HV Units", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Jul-24", "Actual": 0, "Forecast": 73039.1304347826},
        {"Account": "Sales - Heron View Sales", "Category": "Trading Income", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Jul-24", "Actual": 0, "Forecast": 1434695.65217391},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron View", "Applicable_dev": "Heron View",
         "Month": "Jul-24", "Actual": 0, "Forecast": 39515.45},
        {"Account": "COS - Commission HV Units", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Jul-24", "Actual": 0, "Forecast": 71734.7826086957},
        {"Account": "Sales - Heron View Sales", "Category": "Trading Income", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Jul-24", "Actual": 0, "Forecast": 1478173.91304348},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron View", "Applicable_dev": "Heron View",
         "Month": "Jul-24", "Actual": 0, "Forecast": 39515.45},
        {"Account": "COS - Commission HV Units", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Jul-24", "Actual": 0, "Forecast": 73908.6956521739},
        {"Account": "Sales - Heron View Sales", "Category": "Trading Income", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Jul-24", "Actual": 0, "Forecast": 1460782.60869565},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron View", "Applicable_dev": "Heron View",
         "Month": "Jul-24", "Actual": 0, "Forecast": 39515.45},
        {"Account": "COS - Commission HV Units", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Jul-24", "Actual": 0, "Forecast": 73039.1304347826},
        {"Account": "Sales - Heron View Sales", "Category": "Trading Income", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Aug-24", "Actual": 0, "Forecast": 1486869.56521739},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron View", "Applicable_dev": "Heron View",
         "Month": "Aug-24", "Actual": 0, "Forecast": 39515.45},
        {"Account": "COS - Commission HV Units", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Aug-24", "Actual": 0, "Forecast": 74343.4782608696},
        {"Account": "Sales - Heron View Sales", "Category": "Trading Income", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Aug-24", "Actual": 0, "Forecast": 1460782.60869565},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron View", "Applicable_dev": "Heron View",
         "Month": "Aug-24", "Actual": 0, "Forecast": 39515.45},
        {"Account": "COS - Commission HV Units", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Aug-24", "Actual": 0, "Forecast": 73039.1304347826},
        {"Account": "Sales - Heron View Sales", "Category": "Trading Income", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Aug-24", "Actual": 0, "Forecast": 1504260.86956522},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron View", "Applicable_dev": "Heron View",
         "Month": "Aug-24", "Actual": 0, "Forecast": 39515.45},
        {"Account": "COS - Commission HV Units", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Aug-24", "Actual": 0, "Forecast": 75213.0434782609},
        {"Account": "Sales - Heron View Sales", "Category": "Trading Income", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Aug-24", "Actual": 0, "Forecast": 1478173.91304348},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron View", "Applicable_dev": "Heron View",
         "Month": "Aug-24", "Actual": 0, "Forecast": 39515.45},
        {"Account": "COS - Commission HV Units", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Aug-24", "Actual": 0, "Forecast": 73908.6956521739},
        {"Account": "Sales - Heron View Sales", "Category": "Trading Income", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Jun-24", "Actual": 0, "Forecast": 1469478.26086957},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron View", "Applicable_dev": "Heron View",
         "Month": "Jun-24", "Actual": 0, "Forecast": 39515.45},
        {"Account": "COS - Commission HV Units", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Jun-24", "Actual": 0, "Forecast": 73473.9130434783},
        {"Account": "Sales - Heron View Sales", "Category": "Trading Income", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Mar-24", "Actual": 0, "Forecast": 1478173.91304348},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron View", "Applicable_dev": "Heron View",
         "Month": "Mar-24", "Actual": 0, "Forecast": 1789},
        {"Account": "COS - Commission HV Units", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Mar-24", "Actual": 0, "Forecast": 73908.6956521739},
        {"Account": "Sales - Heron View Sales", "Category": "Trading Income", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Jun-24", "Actual": 0, "Forecast": 1304260.86956522},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron View", "Applicable_dev": "Heron View",
         "Month": "Jun-24", "Actual": 0, "Forecast": 39515.45},
        {"Account": "COS - Commission HV Units", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Jun-24", "Actual": 0, "Forecast": 65213.0434782609},
        {"Account": "Sales - Heron View Sales", "Category": "Trading Income", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Jun-24", "Actual": 0, "Forecast": 1295565.2173913},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron View", "Applicable_dev": "Heron View",
         "Month": "Jun-24", "Actual": 0, "Forecast": 39515.45},
        {"Account": "COS - Commission HV Units", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Jun-24", "Actual": 0, "Forecast": 64778.2608695652},
        {"Account": "Sales - Heron View Sales", "Category": "Trading Income", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Jun-24", "Actual": 0, "Forecast": 1295565.2173913},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron View", "Applicable_dev": "Heron View",
         "Month": "Jun-24", "Actual": 0, "Forecast": 39515.45},
        {"Account": "COS - Commission HV Units", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Jun-24", "Actual": 0, "Forecast": 64778.2608695652},
        {"Account": "Sales - Heron View Sales", "Category": "Trading Income", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Jun-24", "Actual": 0, "Forecast": 1304260.86956522},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron View", "Applicable_dev": "Heron View",
         "Month": "Jun-24", "Actual": 0, "Forecast": 39515.45},
        {"Account": "COS - Commission HV Units", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Jun-24", "Actual": 0, "Forecast": 65213.0434782609},
        {"Account": "Sales - Heron View Sales", "Category": "Trading Income", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Jun-24", "Actual": 0, "Forecast": 1460782.60869565},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron View", "Applicable_dev": "Heron View",
         "Month": "Jun-24", "Actual": 0, "Forecast": 1789},
        {"Account": "COS - Commission HV Units", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Jun-24", "Actual": 0, "Forecast": 73039.1304347826},
        {"Account": "Sales - Heron View Sales", "Category": "Trading Income", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Jun-24", "Actual": 0, "Forecast": 1443391.30434783},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron View", "Applicable_dev": "Heron View",
         "Month": "Jun-24", "Actual": 0, "Forecast": 39515.45},
        {"Account": "COS - Commission HV Units", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Jun-24", "Actual": 0, "Forecast": 72169.5652173913},
        {"Account": "Sales - Heron View Sales", "Category": "Trading Income", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Jun-24", "Actual": 0, "Forecast": 1443391.30434783},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron View", "Applicable_dev": "Heron View",
         "Month": "Jun-24", "Actual": 0, "Forecast": 39515.45},
        {"Account": "COS - Commission HV Units", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Jun-24", "Actual": 0, "Forecast": 72169.5652173913},
        {"Account": "Sales - Heron View Sales", "Category": "Trading Income", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Jun-24", "Actual": 0, "Forecast": 1443391.30434783},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron View", "Applicable_dev": "Heron View",
         "Month": "Jun-24", "Actual": 0, "Forecast": 39515.45},
        {"Account": "COS - Commission HV Units", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Jun-24", "Actual": 0, "Forecast": 72169.5652173913},
        {"Account": "Sales - Heron View Sales", "Category": "Trading Income", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Jun-24", "Actual": 0, "Forecast": 1512956.52173913},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron View", "Applicable_dev": "Heron View",
         "Month": "Jun-24", "Actual": 0, "Forecast": 1789},
        {"Account": "COS - Commission HV Units", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Jun-24", "Actual": 0, "Forecast": 75647.8260869565},
        {"Account": "Sales - Heron View Sales", "Category": "Trading Income", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Jul-24", "Actual": 0, "Forecast": 1469478.26086957},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron View", "Applicable_dev": "Heron View",
         "Month": "Jul-24", "Actual": 0, "Forecast": 39515.45},
        {"Account": "COS - Commission HV Units", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Jul-24", "Actual": 0, "Forecast": 73473.9130434783},
        {"Account": "Sales - Heron View Sales", "Category": "Trading Income", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Jul-24", "Actual": 0, "Forecast": 1460782.60869565},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron View", "Applicable_dev": "Heron View",
         "Month": "Jul-24", "Actual": 0, "Forecast": 39515.45},
        {"Account": "COS - Commission HV Units", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Jul-24", "Actual": 0, "Forecast": 73039.1304347826},
        {"Account": "Sales - Heron View Sales", "Category": "Trading Income", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Jul-24", "Actual": 0, "Forecast": 1460782.60869565},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron View", "Applicable_dev": "Heron View",
         "Month": "Jul-24", "Actual": 0, "Forecast": 39515.45},
        {"Account": "COS - Commission HV Units", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Jul-24", "Actual": 0, "Forecast": 73039.1304347826},
        {"Account": "Sales - Heron View Sales", "Category": "Trading Income", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Jul-24", "Actual": 0, "Forecast": 1460782.60869565},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron View", "Applicable_dev": "Heron View",
         "Month": "Jul-24", "Actual": 0, "Forecast": 39515.45},
        {"Account": "COS - Commission HV Units", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Jul-24", "Actual": 0, "Forecast": 73039.1304347826},
        {"Account": "Sales - Heron View Sales", "Category": "Trading Income", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Jul-24", "Actual": 0, "Forecast": 1521652.17391304},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron View", "Applicable_dev": "Heron View",
         "Month": "Jul-24", "Actual": 0, "Forecast": 39515.45},
        {"Account": "COS - Commission HV Units", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Jul-24", "Actual": 0, "Forecast": 76082.6086956522},
        {"Account": "Sales - Heron View Sales", "Category": "Trading Income", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Aug-24", "Actual": 0, "Forecast": 1478173.91304348},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron View", "Applicable_dev": "Heron View",
         "Month": "Aug-24", "Actual": 0, "Forecast": 39515.45},
        {"Account": "COS - Commission HV Units", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Aug-24", "Actual": 0, "Forecast": 73908.6956521739},
        {"Account": "Sales - Heron View Sales", "Category": "Trading Income", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Aug-24", "Actual": 0, "Forecast": 1469478.26086957},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron View", "Applicable_dev": "Heron View",
         "Month": "Aug-24", "Actual": 0, "Forecast": 39515.45},
        {"Account": "COS - Commission HV Units", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Aug-24", "Actual": 0, "Forecast": 73473.9130434783},
        {"Account": "Sales - Heron View Sales", "Category": "Trading Income", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Aug-24", "Actual": 0, "Forecast": 1469478.26086957},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron View", "Applicable_dev": "Heron View",
         "Month": "Aug-24", "Actual": 0, "Forecast": 39515.45},
        {"Account": "COS - Commission HV Units", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Aug-24", "Actual": 0, "Forecast": 73473.9130434783},
        {"Account": "Sales - Heron View Sales", "Category": "Trading Income", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Aug-24", "Actual": 0, "Forecast": 1469478.26086957},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron View", "Applicable_dev": "Heron View",
         "Month": "Aug-24", "Actual": 0, "Forecast": 39515.45},
        {"Account": "COS - Commission HV Units", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Aug-24", "Actual": 0, "Forecast": 73473.9130434783},
        {"Account": "Sales - Heron View Sales", "Category": "Trading Income", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Aug-24", "Actual": 0, "Forecast": 1530347.82608696},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron View", "Applicable_dev": "Heron View",
         "Month": "Aug-24", "Actual": 0, "Forecast": 39515.45},
        {"Account": "COS - Commission HV Units", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Aug-24", "Actual": 0, "Forecast": 76517.3913043478},
        {"Account": "Sales - Heron View Sales", "Category": "Trading Income", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Jul-24", "Actual": 0, "Forecast": 1478173.91304348},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron View", "Applicable_dev": "Heron View",
         "Month": "Jul-24", "Actual": 0, "Forecast": 39515.45},
        {"Account": "COS - Commission HV Units", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Jul-24", "Actual": 0, "Forecast": 73908.6956521739},
        {"Account": "Sales - Heron View Sales", "Category": "Trading Income", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Jul-24", "Actual": 0, "Forecast": 1504260.86956522},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron View", "Applicable_dev": "Heron View",
         "Month": "Jul-24", "Actual": 0, "Forecast": 39515.45},
        {"Account": "COS - Commission HV Units", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Jul-24", "Actual": 0, "Forecast": 75213.0434782609},
        {"Account": "Sales - Heron View Sales", "Category": "Trading Income", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Jul-24", "Actual": 0, "Forecast": 1521652.17391304},
        {"Account": "COS - Legal Fees", "Category": "COS", "Development": "Heron View", "Applicable_dev": "Heron View",
         "Month": "Jul-24", "Actual": 0, "Forecast": 39515.45},
        {"Account": "COS - Commission HV Units", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Jul-24", "Actual": 0, "Forecast": 76082.6086956522},

        {"Account": "COS - Heron View - Construction", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Feb-24", "Actual": 0, "Forecast": 5707656.42066916},
        {"Account": "COS - Heron View - Construction", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Mar-24", "Actual": 0, "Forecast": 10095497.4878337},
        {"Account": "COS - Heron View - Construction", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Apr-24", "Actual": 0, "Forecast": 6644697.30778797},
        {"Account": "COS - Heron View - Construction", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "May-24", "Actual": 0, "Forecast": 9067916.29306447},
        {"Account": "COS - Heron View - Construction", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Jun-24", "Actual": 0, "Forecast": 7381385.27235665},
        {"Account": "COS - Heron View - Construction", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Jul-24", "Actual": 0, "Forecast": 3648699.27027617},
        {"Account": "COS - Heron View - Construction", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Aug-24", "Actual": 0, "Forecast": 1000594.94489614},
        {"Account": "COS - Heron View - Construction", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Sep-24", "Actual": 0, "Forecast": 512305.573497741},
        {"Account": "COS - Heron View - Construction", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Oct-24", "Actual": 0, "Forecast": 754779.014486066},
        {"Account": "COS - Heron View - Construction", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Nov-24", "Actual": 0, "Forecast": 0},
        {"Account": "COS - Heron View - Construction", "Category": "COS", "Development": "Heron View",
         "Applicable_dev": "Heron View", "Month": "Dec-24", "Actual": 0, "Forecast": 0},
    ]
    try:
        current = list(db.profit_and_loss.find())
        # try:
        #     db.profit_and_loss.update_many({}, {"$set": {"Actual": 0, "Forecast": 0}})
        #     print("Awesome")
        # except Exception as e:
        #     print(e)
        # for item in current:
        #     del item['_id']
        #     # Change item['Actual'] to 0 and item['Forecast'] to 0 in current and update the database accordingly
        #     item['Actual'] = 0
        #     item['Forecast'] = 0
        #     try:
        #         db.profit_and_loss.update_one({"Account": item['Account'], "Month": item['Month']}, {"$set": item})
        #         print("Data updated successfully")
        #     except Exception as e:
        #         print(e)

        updated_p_and_l = []
        for item in input_p_and_l:
            insert = {
                "Account": item['Account'],
                "Category": item['Category'],
                "Development": item['Development'],
                "Applicable_dev": item['Applicable_dev'],
                "Month": item['Month'],
            }
            updated_p_and_l.append(insert)
        print(len(updated_p_and_l))
        print(len(input_p_and_l))

        # remove duplicates from updated_p_and_l
        updated_p_and_l = [dict(t) for t in {tuple(d.items()) for d in updated_p_and_l}]
        print(len(updated_p_and_l))

        for item in updated_p_and_l:
            # filter input_p_and_l where Account, Category, Development, Applicable_dev and Month are equal to item['Account'], item['Category'], item['Development'], item['Applicable_dev'] and item['Month'] respectively
            filtered_input_p_and_l = list(filter(
                lambda x: x['Account'] == item['Account'] and x['Category'] == item['Category'] and x['Development'] ==
                          item['Development'] and x['Applicable_dev'] == item['Applicable_dev'] and x['Month'] == item[
                              'Month'], input_p_and_l))
            # sum the Forecast of the filtered_input_p_and_l and make item['Forecast'] equal to the sum
            item['Forecast'] = sum([x['Forecast'] for x in filtered_input_p_and_l])
            # round item['Forecast'] to 2 decimal places
            item['Forecast'] = round(item['Forecast'], 2)
            item['Actual'] = sum([x['Actual'] for x in filtered_input_p_and_l])
            # round item['Actual'] to 2 decimal places
            item['Actual'] = round(item['Actual'], 2)
            # print("Item", item)

        for item in updated_p_and_l:
            # round item['Forecast'] to 2 decimal places
            item['Forecast'] = round(item['Forecast'], 2)
            filtered_current = list(
                filter(lambda x: x['Account'] == item['Account'] and x['Month'] == item['Month'], current))
            if len(filtered_current) == 0:
                # insert item into the database
                try:
                    db.profit_and_loss.insert_one(item)
                    print("Data inserted successfully")
                except Exception as e:
                    print(e)
            else:
                try:
                    db.profit_and_loss.update_one({"Account": item['Account'], "Month": item['Month']}, {"$set": item})
                    print("Data updated successfully")
                except Exception as e:
                    print(e)
        print("Data updated successfully AA")
    except Exception as e:
        print(e)
        # return {"ERROR": "Please Try again"}


# update_profit_and_loss_from_cf_file()


def get_investors_for_interest_for_cashflow():
    # print("Hello")
    try:
        investors = list(db.investors.find())

        investments_to_keep = []
        final_opportunities = []
        # final_rates = []
        # rates = list(db.rates.find())
        # print(rates)
        # print()

        print(investors[10]["investments"][0])
        print()
        print(investors[10]["trust"][0])

        investors = list(filter(lambda x: len(x['investments']) > 0, investors))
        for investment in investors:
            for item in investment['investments']:
                insert = {
                    "investor_acc_number": investment['investor_acc_number'],
                    "Category": item['Category'],
                    "opportunity_code": item['opportunity_code'],
                    "investment_amount": float(item['investment_amount']),
                    "deposit_date": item['deposit_date'],
                    "release_date": item['release_date'],
                    "end_date": item['end_date'],
                    "investment_interest_rate": float(item['investment_interest_rate']),

                }
                investments_to_keep.append(insert)

        investments_to_keep = list(filter(lambda x: x['Category'] != "Southwark", investments_to_keep))
        investments_to_keep = list(filter(lambda x: x['Category'] != "Goodwood", investments_to_keep))
        investments_to_keep = sorted(investments_to_keep, key=lambda x: x['opportunity_code'])

        opportunities = list(db.opportunities.find())
        for opp in opportunities:
            insert = {
                "Category": opp['Category'],
                "opportunity_code": opp['opportunity_code'],
                "opportunity_end_date": opp['opportunity_end_date'],
                "opportunity_final_transfer_date": opp['opportunity_final_transfer_date'],
                "opportunity_sold": opp['opportunity_sold'],
            }
            final_opportunities.append(insert)

        final_opportunities = list(filter(lambda x: x['Category'] != "Southwark", final_opportunities))
        final_opportunities = list(filter(lambda x: x['Category'] != "Goodwood", final_opportunities))
        final_opportunities = sorted(final_opportunities, key=lambda x: x['opportunity_code'])

        # export investment to keep to a csv file
        # with open('investments_to_keep.csv', 'w', newline='') as file:
        #     writer = csv.writer(file)
        #     writer.writerow(["investor_acc_number", "Category", "opportunity_code", "investment_amount", "release_date", "end_date", "investment_interest_rate"])
        #     for item in investments_to_keep:
        #         writer.writerow([item['investor_acc_number'], item['Category'], item['opportunity_code'], item['investment_amount'], item['release_date'], item['end_date'], item['investment_interest_rate']])
        # # export final opportunities to a csv file
        # with open('final_opportunities.csv', 'w', newline='') as file:
        #     writer = csv.writer(file)
        #     writer.writerow(["Category", "opportunity_code", "opportunity_end_date", "opportunity_final_transfer_date", "opportunity_sold"])
        #     for item in final_opportunities:
        #         writer.writerow([item['Category'], item['opportunity_code'], item['opportunity_end_date'], item['opportunity_final_transfer_date'], item['opportunity_sold']])


    except Exception as e:
        print(e)
        return {"ERROR": "Please Try again"}
    print()
    print("Data retrieved successfully")


@xero.post("/get_trial_balances")
async def get_trial_balances(request: Request):
    request = await request.json()
    credentials_string = f"{clientId}:{xerosecret}".encode("utf-8")

    # Encode the combined string in base64
    encoded_credentials = b64encode(credentials_string).decode("utf-8")

    # get data from xeroCredentials and save to variable credentials
    credentials = xeroCredentials.find_one({})

    id = str(credentials["_id"])
    del credentials["_id"]

    access_token = credentials["access_token"]
    refresh_token = credentials["refresh_token"]

    if isinstance(credentials["refresh_expires"], str):
        credentials["refresh_expires"] = datetime.strptime(credentials["refresh_expires"], "%Y-%m-%d %H:%M:%S")
    if credentials["refresh_expires"] > datetime.now():

        url_refresh = "https://identity.xero.com/connect/token"

        # Define the headers
        headers = {
            "Content-Type": "application/x-www-form-urlencoded",
            "Authorization": f"Basic {encoded_credentials}",  # Replace token with your base64-encoded credentials
        }

        # Define the form data
        data = {
            "grant_type": "refresh_token",
            "refresh_token": refresh_token,
        }

        # print(data)

        # Send the POST request
        response = requests.post(url_refresh, headers=headers, data=data)

        # Check the response
        if response.status_code == 200:
            # Request was successful, you can access response data like JSON content
            response_data = response.json()
            access_token = response_data.get("access_token")
            expires_in = response_data.get("expires_in")
            refresh_token = response_data.get("refresh_token")
            refresh_expires = datetime.now() + timedelta(days=60)
            insert = {
                "access_token": access_token,
                "expires_in": expires_in,
                "refresh_token": refresh_token,
                "refresh_expires": refresh_expires,
            }
            # update the document in the database
            xeroCredentials.update_one({"_id": ObjectId(id)}, {"$set": insert})

            print("SUCCESS")

        else:
            # Request failed, handle the error
            print(f"ErrorXX: {response.status_code} - {response.text} - {response.json()}")

    else:
        print("refresh_expires is in the past")

    tenants1 = list(xeroTenants.find({}))
    # print("tenants1",tenants1)

    tenant_id = "30b5d5a0-cf38-4bdb-baa1-9dda35b278a2"
    tenant_id_HF = "9c4ba92b-93b0-4358-9ff8-141aa0718242"

    tenant_id_HV = "4af624e3-6de5-4cc7-9123-36d63d2acbb4"

    # tenants = [
    #     # tenant_id,
    #     tenant_id_HF,
    #     # tenant_id_HV
    # ]

    tenants = [tenant['tenantId'] for tenant in tenants1]

    data = []

    for index1, tenant in enumerate(tenants):

        try:
            async with httpx.AsyncClient() as client:
                response = await client.get(
                    f"https://api.xero.com/api.xro/2.0/Reports/TrialBalance?date={request['date']}",
                    headers={
                        "Content-Type": "application/xml",
                        "Authorization": f"Bearer {access_token}",
                        "xero-tenant-id": tenant,
                    },
                )
        except httpx.HTTPError as exc:
            print(f"HTTP Error: {exc}")
            return {"Success": False, "Error": "HTTP Error"}

        # Check the HTTP status code
        if response.status_code != 200:
            return {"Success": False, "Error": "Non-200 Status Code"}

        python_dict = xmltodict.parse(response.text)

        report_title = python_dict['Response']['Reports']['Report']['ReportTitles']['ReportTitle'][1]

        report_date = request['date']

        python_dict = python_dict['Response']['Reports']['Report']['Rows']['Row']

        header = []

        for index, item in enumerate(python_dict):
            if index == 0:
                for x in item['Cells']['Cell']:
                    header.append(x['Value'])

            else:
                for y in item:
                    if "Title" in y:
                        title = item['Title']
                        # print(title)
                        for index1, val in enumerate(item['Rows']['Row']):

                            if "Cells" in val:
                                values = []

                                try:

                                    for index2, items2 in enumerate(val['Cells']['Cell']):
                                        try:
                                            values.append(float(items2.get('Value', 0)))
                                        except ValueError:
                                            values.append(items2.get('Value', 0))
                                    insert = {
                                        "ReportTitle": report_title,
                                        "ReportDate": report_date,
                                        "Category": title,
                                        "Values": values,
                                        "Header": header
                                    }
                                    data.append(insert)

                                except TypeError:
                                    test = item['Rows']['Row']['Cells']['Cell']

                                    for index3, item3 in enumerate(test):
                                        try:
                                            values.append(float(item3.get('Value', 0)))
                                        except ValueError:
                                            values.append(item3.get('Value', 0))
                                        insert = {
                                            "ReportTitle": report_title,
                                            "ReportDate": report_date,
                                            "Category": title,
                                            "Values": values,
                                            "Header": header
                                        }
                                        data.append(insert)

                                    continue

                                except KeyError:
                                    print("KeyError", val, index1)
                                    continue

        final_data = []
        for item in data:
            insert = {
                "ReportTitle": item['ReportTitle'],
                "ReportDate": item['ReportDate'],
                "Category": item['Category'],
            }
            for index, val in enumerate(item['Header']):
                if index == 0:
                    account_details = item['Values'][index].split("(")
                    account_name = account_details[0].strip()
                    # print(account_details)
                    if len(account_details) > 1:
                        account_code = account_details[1].split(")")[0]
                        insert["AccountCode"] = account_code
                        insert["AccountName"] = account_name
                        # print(account_details)

                    # account_code = account_details[1].split(")")[0]
                    # insert["AccountCode"] = account_code
                    # insert["AccountName"] = account_name
                    # print(account_details)
                else:
                    insert[val] = item['Values'][index]

            final_data.append(insert)

    for item in final_data:
        item['Current'] = item['Debit'] - item['Credit']
        item['YTD'] = item['YTD Debit'] - item['YTD Credit']
        item['ReportDate'] = item['ReportDate'].replace("/", "-")
        del item['Debit']
        del item['Credit']
        del item['YTD Debit']
        del item['YTD Credit']

    print("final_data", len(final_data))

    cashflow_xero_tb = list(db.cashflow_xero_tb.find({"ReportDate": report_date}))
    for item in cashflow_xero_tb:
        item["_id"] = str(item["_id"])

    # filter cashflow_xero_tb where "ReportDate" is equal to report_date
    # cashflow_xero_tb = list(filter(lambda x: x['ReportDate'] == report_date, cashflow_xero_tb))

    print("cashflow_xero_tb:", len(cashflow_xero_tb))

    # remove all duplicate data from final_data
    final_data = [dict(t) for t in {tuple(d.items()) for d in final_data}]

    print("final_data", len(final_data))
    # print("final_data",final_data[0])

    report_title_list = []
    for item in final_data:
        report_title_list.append(item['ReportTitle'])

    # remove all duplicate data from report_title_list
    # report_title_list = list(set(report_title_list))
    #
    # print(report_title_list)
    # print(len(report_title_list))

    if len(cashflow_xero_tb) == 0:
        try:
            db.cashflow_xero_tb.insert_many(final_data)
        except Exception as e:
            print("XXXXXXX", e)
            return {"ERROR": "Please Try again"}
    else:
        for item in final_data:
            try:
                filtered_cashflow_xero_tb = list(filter(
                    lambda z: z.get('AccountCode',"") == item.get('AccountCode',"000") and z['ReportDate'] == item['ReportDate'] and z[
                        'AccountName'] == item['AccountName'] and z['ReportTitle'] == item['ReportTitle'],
                    cashflow_xero_tb))
                if len(filtered_cashflow_xero_tb) == 0:
                    try:
                        db.cashflow_xero_tb.insert_one(item)
                        print("Successful Insert")
                    except Exception as e:
                        print("Error", e)
                        continue

                elif len(filtered_cashflow_xero_tb) == 1:
                    item_to_update = filtered_cashflow_xero_tb[0]
                    # print("OKAY")

                    if int(item_to_update['Current'] - item['Current'] + item_to_update['YTD'] - item['YTD']) != 0:
                        # update the Current and YTD in the database where the _id is equal to item_to_update['_id']
                        # and Current equals item['Current'] and YTD equals item['YTD']. Only update those two fields

                        print("Update Needed")
                        print("item_to_update CURRENT",
                              item_to_update['Current'] - item['Current'] + item_to_update['YTD'] - item['YTD'], item,
                              item_to_update)
                        print()

                        try:
                            _id = item_to_update['_id']
                            del item_to_update['_id']

                            result = db.cashflow_xero_tb.update_one({"_id": ObjectId(_id)}, {"$set": item})

                            print("Update Successful")
                        except Exception as e:
                            print("Error", e, "item_to_update", item_to_update, "item", item, "_id", _id, "Result",result)
                            continue

                else:
                    print("More than one item found", item_to_update, item)

            except Exception as e:
                print("Error 2", e, "item_to_update", item_to_update, "item", item)
                # return {"ERROR": "Please Try again"}
                continue

    # create list called bank_accounts and filter only where "AccountCode" begins with 84

    # bank_accounts = list(filter(lambda x: x['AccountCode'].startswith("84"), final_data))
    # # filter bank accounts and exclude all where the "AccountCode" begins with 8480
    # bank_accounts = list(filter(lambda x: not x['AccountCode'].startswith("8480"), bank_accounts))
    # print("Length Banks", len(bank_accounts))
    #
    # for acc in bank_accounts:
    #     if '_id' in acc:
    #         del acc['_id']
    #     # print(acc)

    bank_accounts = list(db.cashflow_xero_tb.find({"AccountCode": {"$regex": "^84"}}, {"_id": 0}))
    bank_accounts = list(filter(lambda x: not x['AccountCode'].startswith("8480"), bank_accounts))
    print("Length Banks", len(bank_accounts))

    bank_periods = []
    for item in bank_accounts:
        bank_periods.append(item['ReportDate'].replace("/", "-"))

    bank_periods = list(set(bank_periods))

    # print(bank_periods)

    bank_periods = sorted(bank_periods, key=lambda x: datetime.strptime(x, "%Y-%m-%d"))

    bank_summary = []

    for period in bank_periods:
        bank_accounts_filtered = list(filter(lambda x: x['ReportDate'].replace("/", "-") == period, bank_accounts))

        bank_balances = sum([x['YTD'] for x in bank_accounts_filtered if x['AccountName'] != "Momentum Investors Account RU502229930"])
        momentum_balances = sum([x['YTD'] for x in bank_accounts_filtered if x['AccountName'] == "Momentum Investors Account RU502229930"])
        print(bank_balances, momentum_balances, period)
        bank_balances = f"R {bank_balances:,.2f}"
        momentum_balances = f"R {momentum_balances:,.2f}"
        insert = {
            "Period": period,
            "BankBalances": bank_balances,
            "MomentumBalances": momentum_balances
        }
        bank_summary.append(insert)





    # print(bank_periods)

    sum_current = sum([x['Current'] for x in bank_accounts])

    sum_YTD = sum([x['YTD'] for x in bank_accounts])

    # format the above as currency with a symbol of R
    sum_current = f"R {sum_current:,.2f}"
    sum_YTD = f"R {sum_YTD:,.2f}"
    # sum_YTD_debits = f"R {sum_YTD_debits:,.2f}"
    # sum_YTD_credits = f"R {sum_YTD_credits:,.2f}"

    print(sum_current, sum_YTD)

    # get all records from cashflow_xero_tb where the AccountCode begins with 84

    # latest_data = list(db.cashflow_xero_tb.find({}))
    # for item in latest_data:
    #     item["_id"] = str(item["_id"])

    return {"summary": bank_summary, "final_data": bank_accounts}

# def get_investor_id_numbers():
#     # get all investors from the database and project investor_acc_number, investor_id_number, investor_name and investor_surname, investment_name, investor_email, investor_mobile and investor_landline
#     investors = list(db.investors.find({}, {"investor_acc_number": 1, "investor_id_number": 1, "investor_name": 1, "investor_surname": 1, "investment_name": 1, "investor_email": 1, "investor_mobile": 1, "investor_landline": 1,"investor_organisation":1, "_id": 0}))
#     print(investors[0])
#
#     # export the investors to an excel file and ensure investor_id_number is a string
#
#     with open('investors.csv', 'w', newline='') as file:
#         writer = csv.writer(file)
#         writer.writerow(["investor_acc_number", "investor_id_number", "investor_name", "investor_surname", "investment_name", "investor_email", "investor_mobile", "investor_landline", "investor_organisation"])
#         for item in investors:
#             writer.writerow([item['investor_acc_number'],"'" + item['investor_id_number'], item['investor_name'], item['investor_surname'], item.get('investment_name',""), item['investor_email'], item['investor_mobile'], item['investor_landline'], item.get('investor_organisation',"")])
#
#     return {"Success": True}
#
# get_investor_id_numbers()
