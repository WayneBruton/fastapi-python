import os
from datetime import datetime, timedelta
import xmltodict
import requests
import time
from fastapi import APIRouter, Request
import httpx
from base64 import b64encode
from urllib.parse import urlencode
from fastapi.responses import FileResponse
from fastapi.responses import RedirectResponse
from config.db import db
from bson.objectid import ObjectId
from decouple import config

from cashflow_excel_functions.cpc_profit_loss_files import insert_data_from_xero_profit_loss
import cashflow_excel_functions.cpc_data_fields as cpc_data_fields

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


@xero.post("/get_profit_and_loss")
async def get_profit_and_loss(data: Request):
    request = await data.json()

    year = request['from_date'].split("-")[0]
    month = request['from_date'].split("-")[1]

    try:
        start_time = time.time()

        if int(month) < 10:
            current_year = int(year)
        else:
            current_year = int(year) - 1

        periods_to_report = [
            {
                "period": 1,
                "from_date": f"{current_year}-03-01",
                "to_date": f"{current_year}-03-31"
            },
            {
                "period": 2,
                "from_date": f"{current_year}-04-01",
                "to_date": f"{current_year}-04-30"
            },
            {
                "period": 3,
                "from_date": f"{current_year}-05-01",
                "to_date": f"{current_year}-05-31"
            },
            {
                "period": 4,
                "from_date": f"{current_year}-06-01",
                "to_date": f"{current_year}-06-30"
            },
            {
                "period": 5,
                "from_date": f"{current_year}-07-01",
                "to_date": f"{current_year}-07-31"
            },
            {
                "period": 6,
                "from_date": f"{current_year}-08-01",
                "to_date": f"{current_year}-08-31"
            },
            {
                "period": 7,
                "from_date": f"{current_year}-09-01",
                "to_date": f"{current_year}-09-30"
            },
            {
                "period": 8,
                "from_date": f"{current_year}-10-01",
                "to_date": f"{current_year}-10-31"
            },
            {
                "period": 9,
                "from_date": f"{current_year}-11-01",
                "to_date": f"{current_year}-11-30"
            },
            {
                "period": 10,
                "from_date": f"{current_year}-12-01",
                "to_date": f"{current_year}-12-31"
            },
            {
                "period": 11,
                "from_date": f"{current_year + 1}-01-01",
                "to_date": f"{current_year + 1}-01-31"
            },
            {
                "period": 12,
                "from_date": f"{current_year + 1}-02-01",
                "to_date": f"{current_year + 1}-02-29"
            }
        ]

        data_from_xero = []
        data_from_xero_HF_PandL = []
        data_from_xero_HV_PandL = []

        base_data = cpc_data_fields.base_data.copy()

        for record in base_data:
            for i in range(request['period']):
                record['Amount'][i] = 0.0

        base_data_HF_PandL = cpc_data_fields.base_data_HF_PandL.copy()

        for record in base_data_HF_PandL:
            for i in range(request['period']):
                record['Amount'][i] = 0.0

        base_data_HV_PandL = cpc_data_fields.base_data_HV_PandL.copy()

        for record in base_data_HV_PandL:
            for i in range(0, request['period']):
                # print(request['period'])
                record['Amount'][i] = 0.0

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

        tenant_id = "30b5d5a0-cf38-4bdb-baa1-9dda35b278a2"
        tenant_id_HF = "9c4ba92b-93b0-4358-9ff8-141aa0718242"

        tenant_id_HV = "4af624e3-6de5-4cc7-9123-36d63d2acbb4"

        tenants = [tenant_id, tenant_id_HF, tenant_id_HV]

        tenants_tb = [tenant_id_HF, tenant_id_HV]

        hv_tb = []
        hf_tb = []

        for index, tenant in enumerate(tenants_tb):
            report_date = request['to_date']

            try:
                async with httpx.AsyncClient() as client:
                    response = await client.get(
                        f"https://api.xero.com/api.xro/2.0/Reports/TrialBalance?date={report_date}",
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

            # print("Hello")

            python_dict = xmltodict.parse(response.text)

            python_dict = python_dict['Response']['Reports']['Report']['Rows']['Row']

            python_dict = list(filter(lambda x: x['RowType'] == 'Section', python_dict))

            for item in python_dict:
                new_item = item['Rows']['Row']
                for cells in new_item:

                    if isinstance(cells, dict):
                        insert = {}

                        for new_index, cell in enumerate(cells['Cells']['Cell']):

                            # put all the above as key value pairs in a dictionary
                            if new_index == 0:
                                clean_account = cell.get('Value', 0)
                                if clean_account != 0:
                                    clean_account_list = clean_account.split("(")
                                    clean_account_list[0] = clean_account_list[0].strip()
                                    clean_account_list[1] = clean_account_list[1].replace(")", "")
                                    clean_account_list[1] = clean_account_list[1].strip()

                                insert['Account Code'] = clean_account_list[1]
                                insert['Account'] = clean_account_list[0]
                                insert['Account Type'] = ''

                            if new_index == 3:
                                insert['debit_ytd'] = float(cell.get('Value', 0))
                            if new_index == 4:
                                insert['credit_ytd'] = float(cell.get('Value', 0))

                        if index == 0:
                            hf_tb.append(insert)
                        elif index == 1:
                            hv_tb.append(insert)

        short_months = [2, 4, 7, 9, 12]
        comparison_data_cpc = []
        comparison_data_hf = []
        comparison_data_hv = []

        # Make a GET request to the connections endpoint
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

                                        elif index1 == 1:
                                            data_from_xero_HF_PandL.append(insert)

                                        elif index1 == 2:
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
                                    if index1 == 0:
                                        comparison_data_cpc.append(insert)

                                    elif index1 == 1:
                                        comparison_data_hf.append(insert)

                                    elif index1 == 2:
                                        comparison_data_hv.append(insert)

            else:

                period = request['period']

                filtered_xero = list(filter(lambda x: x['period'] == period, periods_to_report))

                report_from_date = filtered_xero[0]['from_date']
                report_to_date = filtered_xero[0]['to_date']
                period = filtered_xero[0]['period']

                try:
                    async with httpx.AsyncClient() as client:
                        response = await client.get(
                            f"https://api.xero.com/api.xro/2.0/Reports/ProfitAndLoss?fromDate={report_from_date}&toDate={report_to_date}&periods={period - 1}&timeframe=MONTH",
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
                                    if index1 == 0:
                                        comparison_data_cpc.append(insert)

                                    elif index1 == 1:
                                        comparison_data_hf.append(insert)

                                    elif index1 == 2:
                                        comparison_data_hv.append(insert)

        for item in comparison_data_cpc:
            item['Amount'].reverse()
            insert = {}
            for idx, value in enumerate(item['Amount'], start=1):
                insert[0] = item['Account']
                insert[1] = float(value)
                insert['period'] = idx

                data_from_xero.append(insert)
                insert = {}

        for item in comparison_data_hf:
            item['Amount'].reverse()
            insert = {}
            for idx, value in enumerate(item['Amount'], start=1):
                insert[0] = item['Account']
                insert[1] = float(value)
                insert['period'] = idx

                data_from_xero_HF_PandL.append(insert)
                insert = {}

        for item in comparison_data_hv:
            item['Amount'].reverse()
            insert = {}
            for idx, value in enumerate(item['Amount'], start=1):
                insert[0] = item['Account']
                insert[1] = float(value)
                insert['period'] = idx

                data_from_xero_HV_PandL.append(insert)
                insert = {}

        data_from_xero.sort(key=lambda x: (x[0], x['period']))

        data_from_xero_HF_PandL.sort(key=lambda x: (x[0], x['period']))

        data_from_xero_HV_PandL.sort(key=lambda x: (x[0], x['period']))

        for final_record in base_data:

            filtered_data_from_xero = list(filter(lambda x: x[0] == final_record['Account'], data_from_xero))

            for record in filtered_data_from_xero:
                final_record['Amount'][record['period'] - 1] = record[1]

            final_record['Amount'].reverse()

        for final_record in base_data_HF_PandL:

            filtered_data_from_xero = list(filter(lambda x: x[0] == final_record['Account'], data_from_xero_HF_PandL))

            for record in filtered_data_from_xero:
                final_record['Amount'][record['period'] - 1] = record[1]

            final_record['Amount'].reverse()

        for record in base_data_HF_PandL:
            if record['Account'] == 'Accounting Fees':

                filtered_data_from_xero = list(
                    filter(lambda x: x['Account'] == 'Accounting - CIPC', base_data_HF_PandL))

                for index, item in enumerate(record['Amount']):
                    record['Amount'][index] = record['Amount'][index] + filtered_data_from_xero[0]['Amount'][index]

            if record['Account'] == 'Staff welfare':

                filtered_data_from_xero1 = list(
                    filter(lambda x: x['Account'] == 'Entertainment Expenses', base_data_HF_PandL))

                filtered_data_from_xero2 = list(
                    filter(lambda x: x['Account'] == 'General Expenses', base_data_HF_PandL))

                for index, item in enumerate(record['Amount']):
                    record['Amount'][index] = filtered_data_from_xero1[0]['Amount'][index] + \
                                              filtered_data_from_xero2[0]['Amount'][index]

            if record['Account'] == 'Repairs _AND_ Maintenance':
                # print("record", record)
                filtered_data_from_xero = list(
                    filter(lambda x: x['Account'] == 'Motor Vehicle Expenses', base_data_HF_PandL))

                for index, item in enumerate(record['Amount']):
                    record['Amount'][index] = record['Amount'][index] + filtered_data_from_xero[0]['Amount'][index]

        for final_record in base_data_HV_PandL:

            filtered_data_from_xero = list(filter(lambda x: x[0] == final_record['Account'], data_from_xero_HV_PandL))

            for record in filtered_data_from_xero:
                final_record['Amount'][record['period'] - 1] = record[1]

            final_record['Amount'].reverse()

        # sort data_from_xero_HF_PandL first by '0' then by 'period'
        data_from_xero_HF_PandL.sort(key=lambda x: (x[0], x['period']))
        data_from_xero_HV_PandL.sort(key=lambda x: (x[0], x['period']))
        # print("data_from_xero_HF_PandL", data_from_xero_HV_PandL)

        returned_data = insert_data_from_xero_profit_loss(base_data, base_data_HF_PandL, base_data_HV_PandL, hf_tb,
                                                          hv_tb, request['to_date'])

        end_time = time.time()

        if returned_data['Success']:
            return {"Success": True, "time_taken": end_time - start_time}
        else:
            return {"Success": False, "Error": returned_data['Error']}



    except Exception as e:
        print("Error", e)
        return {"Success": False, "Error": str(e)}


@xero.get("/get_profit_and_loss")
async def get_profit_and_loss(file_name):
    try:
        file_name = file_name + ".xlsx"
        file_name = file_name.replace("_", " ")
        file_name = file_name.replace("xx", "&")

        dir_path = "cashflow_p&l_files"
        dir_list = os.listdir(dir_path)

        if file_name in dir_list:

            return FileResponse(f"{dir_path}/{file_name}", filename=file_name)
        else:
            return {"ERROR": "File does not exist!!"}
    except Exception as e:
        print(e)
        return {"ERROR": "Please Try again"}
