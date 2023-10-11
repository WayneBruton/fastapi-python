import os

from datetime import datetime, timedelta

import xmltodict

import requests

from fastapi import APIRouter, Request, Depends, HTTPException
import httpx
from base64 import b64encode
from urllib.parse import urlencode
import openpyxl
# from fastapi.responses import FileResponse
from config.db import db

from bson.objectid import ObjectId

from decouple import config

import time

# from typing import Annotated

# from sales_python_files.sales_excel_sheet import create_excel_file
# from sales_client_onboarding_creation_file.onboarding import print_onboarding_pdf
# from sales_client_onboarding_creation_file.offer_to_purchase import print_otp_pdf
from cashflow_excel_functions.cpc_profit_loss_files import insert_data_from_xero_profit_loss


xero = APIRouter()

# investors = db.investors
# opportunityCategories = db.opportunityCategories
# opportunities = db.opportunities
# sales_parameters = db.salesParameters
# sales_processed = db.sales_processed
# sales_agents = db.sales_agents
# mortgage_brokers = db.sales_mortgage_brokers

xeroCredentials = db.xeroCredentials
xeroTenants = db.xeroTenants

clientId = config("clientId")
xerosecret = config("xerosecret")
redirectURI = config("redirectURI")


# url = f"https://login.xero.com/identity/connect/authorize?response_type=code&client_id={clientId}&redirect_uri={
# redirectURI}&scope=accounting.reports.read&state=123"


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
        f"state=123"
    )

    credentials = xeroCredentials.find_one({})
    refresh_expires = credentials["refresh_expires"]
    # is refresh_expires in the future?
    if refresh_expires > datetime.now():
        return {"response": authorization_url, "refresh_valid": True}
    else:
        return {"response": authorization_url, "refresh_valid": False}


@xero.get("/callback")
async def xero_callback(request: Request, code: str):
    # Extract the authorization code from the query parameter
    code = code.split("&")[0].split("=")[-1]

    # print("code", code)

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

    return {"authorized": True}


@xero.post("/get_profit_and_loss")
async def get_profit_and_loss(data: Request):
    request = await data.json()

    # print(request)
    # get the year and month from request['from_date']
    year = request['from_date'].split("-")[0]
    month = request['from_date'].split("-")[1]
    # print(year, month)

    try:
        start_time = time.time()
        # if int(month) < 10 then current_year = year else current_year = year - 1
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

        # print("periods_2024", periods_to_report)

        data_from_xero = []

        base_data = [{'Account': 'Fees - Construction - Heron',
                      'Amount': [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]},
                     {'Account': 'COS - Bakhoven',
                      'Amount': [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]},
                     {'Account': 'COS - Electricity Cost Endulini',
                      'Amount': [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]},
                     {'Account': 'COS - Electricity Cost Heron Field',
                      'Amount': [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]},
                     {'Account': 'COS - Endulini - Telephone & Internet',
                      'Amount': [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]},
                     {'Account': 'COS - Endulini Construction',
                      'Amount': [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]},
                     {'Account': 'COS - Endulini P and G',
                      'Amount': [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]},
                     {'Account': 'COS - Heron - Internet',
                      'Amount': [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]},
                     {'Account': 'COS - Heron Fields - Construction',
                      'Amount': [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]},
                     {'Account': 'COS - Heron Fields - Health & Safety',
                      'Amount': [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]},
                     {'Account': 'COS - Heron Fields - P & G',
                      'Amount': [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]},
                     {'Account': 'COS - Heron Fields - Printing & Stationary',
                      'Amount': [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]},
                     {'Account': 'COS - Heron View - Construction',
                      'Amount': [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]},
                     {'Account': 'COS - Heron View - P&G',
                      'Amount': [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]},
                     {'Account': 'COS - Heron View - Printing & Stationary',
                      'Amount': [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]},
                     {'Account': 'COS - Insurance',
                      'Amount': [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]},
                     {'Account': 'COS - Repairs & Maintenance - SH Soho',
                      'Amount': [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]},
                     {'Account': 'COS - Repairs & Maintenance - SW Southwark',
                      'Amount': [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]},
                     {'Account': 'COS - Security - Guarding',
                      'Amount': [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]},
                     {'Account': 'Cost of Sales - SouthWark Project',
                      'Amount': [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]},
                     {'Account': 'Interest Received - FNB',
                      'Amount': [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]},
                     {'Account': 'Interest received - Momentum',
                      'Amount': [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]},
                     {'Account': 'Rent - CT Office',
                      'Amount': [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]},
                     {'Account': 'Accounting Fees - Audit',
                      'Amount': [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]},
                     {'Account': 'Accounting Fees - Other',
                      'Amount': [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]},
                     {'Account': 'Advertising - Design Fees',
                      'Amount': [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]},
                     {'Account': 'Advertising - Other',
                      'Amount': [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]},
                     {'Account': 'Advertising _AND_ Promotions',
                      'Amount': [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]},
                     {'Account': 'Bank Charges',
                      'Amount': [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]},
                     {'Account': 'BIBC Company Contribution',
                      'Amount': [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]},
                     {'Account': 'BIBC Employee Contribution',
                      'Amount': [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]},
                     {'Account': 'Cleaning', 'Amount': [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]},
                     {'Account': 'Computer Exp - IT, Internet/Hosting Fee',
                      'Amount': [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]},
                     {'Account': 'Computer Expenses',
                      'Amount': [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]},
                     {'Account': 'Consulting Fees - Admin and Finance',
                      'Amount': [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]},
                     {'Account': 'Courier _AND_ Postage',
                      'Amount': [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]},
                     {'Account': 'Depreciation - Computer Equipment',
                      'Amount': [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]},
                     {'Account': 'Depreciation - Furniture and Fittings',
                      'Amount': [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]},
                     {'Account': 'Depreciation - Generator Fixed Asset',
                      'Amount': [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]},
                     {'Account': 'Electricity _AND_ Water',
                      'Amount': [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]},
                     {'Account': 'Insurance - Santam',
                      'Amount': [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]},
                     {'Account': 'Legal Fees', 'Amount': [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]},
                     {'Account': 'Motor Vehicle - Insurance _AND_ Licence',
                      'Amount': [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]},
                     {'Account': 'Motor Vehicle - Petrol _AND_ Oil',
                      'Amount': [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]},
                     {'Account': 'Motor Vehicle Expenses',
                      'Amount': [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]},
                     {'Account': 'PAYE Contributions',
                      'Amount': [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]},
                     {'Account': 'Printing - Printer rental',
                      'Amount': [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]},
                     {'Account': 'Printing _AND_ Stationery',
                      'Amount': [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]},
                     {'Account': 'Rates', 'Amount': [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]},
                     {'Account': 'Rent Paid', 'Amount': [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]},
                     {'Account': 'Repairs _AND_ Maintenance',
                      'Amount': [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]},
                     {'Account': 'Salaries & Wages - WCA',
                      'Amount': [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]},
                     {'Account': 'Salaries _AND_ Wages',
                      'Amount': [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]},
                     {'Account': 'SDL Contributions',
                      'Amount': [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]},
                     {'Account': 'Secretarial fees - CIPC',
                      'Amount': [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]},
                     {'Account': 'Security', 'Amount': [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]},
                     {'Account': 'Small Assets',
                      'Amount': [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]},
                     {'Account': 'Staff Welfare _AND_ Refreshmts',
                      'Amount': [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]},
                     {'Account': 'Subscriptions - Smartsheet',
                      'Amount': [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]},
                     {'Account': 'Subscriptions & Licenses - Caseware',
                      'Amount': [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]},
                     {'Account': 'Subscriptions & Licenses - Sage payroll',
                      'Amount': [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]},
                     {'Account': 'Subscriptions & Licenses - Xero',
                      'Amount': [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]},
                     {'Account': 'Subscriptions / Licenses',
                      'Amount': [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]},
                     {'Account': 'UIF Company Contributions',
                      'Amount': [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]},
                     {'Account': 'UIF Employee Contribution',
                      'Amount': [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]}]

        credentials_string = f"{clientId}:{xerosecret}".encode("utf-8")

        # Encode the combined string in base64
        encoded_credentials = b64encode(credentials_string).decode("utf-8")

        # print("encoded_credentials", encoded_credentials)

        # get data from xeroCredentials and save to variable credentials
        credentials = xeroCredentials.find_one({})
        id = str(credentials["_id"])
        del credentials["_id"]

        access_token = credentials["access_token"]
        refresh_token = credentials["refresh_token"]

        # refresh_expires = credentials["refresh_expires"], it is datetime object, ensure it is in the future
        # if it is in the future, use the refresh token to get a new access token
        if credentials["refresh_expires"] > datetime.now():
            # print("refresh_expires is in the future")

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

        # Make a GET request to the connections endpoint
        for period in range(1, request['period'] + 1):
            filtered_xero = list(filter(lambda x: x['period'] == period, periods_to_report))
            # print("filtered_xero", filtered_xero)
            report_from_date = filtered_xero[0]['from_date']
            report_to_date = filtered_xero[0]['to_date']
            period = filtered_xero[0]['period']

            async with httpx.AsyncClient() as client:
                response = await client.get(
                    f"https://api.xero.com/api.xro/2.0/Reports/ProfitAndLoss?fromDate={report_from_date}&toDate="
                    f"{report_to_date}",
                    headers={
                        "Content-Type": "application/xml",
                        "Authorization": f"Bearer {access_token}",
                        "xero-tenant-id": tenant_id,
                    },
                )

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
                                    # print(row['Cells']['Cell'])
                                    final_line_data = row['Cells']['Cell']
                                    insert = {}
                                    for index, line_item in enumerate(final_line_data):

                                        # print(index, ':', line_item['Value'])
                                        try:
                                            insert[index] = float(line_item['Value'])
                                        except:
                                            insert[index] = line_item['Value']
                                    insert['period'] = period
                                    data_from_xero.append(insert)

        # sort data_from_xero first by 0 then by period
        data_from_xero.sort(key=lambda x: (x[0], x['period']))
        for final_record in base_data:

            filtered_data_from_xero = list(filter(lambda x: x[0] == final_record['Account'], data_from_xero))

            for record in filtered_data_from_xero:
                final_record['Amount'][record['period'] - 1] = record[1]

            final_record['Amount'].reverse()
        # print("base_data", base_data)
        returned_data = insert_data_from_xero_profit_loss(base_data)
        end_time = time.time()
        print("time taken", end_time - start_time)
        # print("returned_data", returned_data)
        if returned_data['Success']:
            return {"Success": True}
        else:
            return {"Success": False, "Error": returned_data['Error']}




        # return {"Success": True}

    except Exception as e:
        print("Error", e)
        return {"Success": False, "Error": str(e)}
