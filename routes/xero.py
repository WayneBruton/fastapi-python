# import os
from datetime import datetime, timedelta
import xmltodict
import pprint

from fastapi import APIRouter, Request, Depends, HTTPException
import httpx
from base64 import b64encode
from urllib.parse import urlencode

from fastapi.responses import FileResponse
from config.db import db

# from bson.objectid import ObjectId
# from decouple import config
# from typing import Annotated

# from sales_python_files.sales_excel_sheet import create_excel_file
# from sales_client_onboarding_creation_file.onboarding import print_onboarding_pdf
# from sales_client_onboarding_creation_file.offer_to_purchase import print_otp_pdf

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

clientId = "38209D80951344669758F26B5F119227"
xerosecret = "2HPTGXHRFsuA7Jk4jmy7LSC65FHDVn_LOxGVI0rjcJM-0ePZ"
redirectURI = "http://localhost:8000/callback"

url = f"https://login.xero.com/identity/connect/authorize?response_type=code&client_id={clientId}&redirect_uri={redirectURI}&scope=accounting.reports.read&state=123"


@xero.get("/authorizeApp")
async def authorize_app():
    # Construct the authorization URL
    authorization_url = (
        f"https://login.xero.com/identity/connect/authorize?"
        f"response_type=code&"
        f"client_id={clientId}&"
        f"redirect_uri={redirectURI}&"
        f"scope=accounting.reports.read offline_access&"
        f"state=123"
    )

    return {"response": authorization_url}


@xero.get("/callback")
async def xero_callback(request: Request, code: str):
    # Extract the authorization code from the query parameter
    code = code.split("&")[0].split("=")[-1]

    print("code", code)

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

        # Your logic to store the tokens in the database goes here
        # You can use your database client and schema to perform these operations
        # For simplicity, I'll print the tokens
        print("Access Token:", access_token)
        print("Expires In:", expires_in)
        print("Refresh Token:", refresh_token)

        # url: "https://api.xero.com/connections",
        #           method: "GET",
        #           headers: {
        #             "Content-Type": "application/json",
        #             Authorization: `Bearer ${access_token}`,
        #           },

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

            # Extract the tenant ID from the response
            # print("response.json()", response.json())
            tenant_id = response.json()

            # print("Tenant ID:", tenant_id)

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

        # Use the access token to make requests to Xero's API endpoints
        # Your logic to interact with Xero goes here

    # Redirect to your desired URL after handling the tokens
    # Replace "process.env.reroute_URI" with your actual redirection URL
    return {"message": "Callback completed successfully"}


@xero.post("/get_profit_and_loss")
async def get_profit_and_loss(data: Request):
    request = await data.json()
    print("request", request)
    try:

        # get the access token from the database
        access_token = xeroCredentials.find_one({})["access_token"]
        refresh_token = xeroCredentials.find_one({})["refresh_token"]

        tenant_id = "30b5d5a0-cf38-4bdb-baa1-9dda35b278a2"

        # Make a GET request to the connections endpoint
        async with httpx.AsyncClient() as client:
            response = await client.get(
                f"https://api.xero.com/api.xro/2.0/Reports/ProfitAndLoss?fromDate={request['from_date']}&toDate={request['to_date']}&periods={request['period']}",
                headers={
                    "Content-Type": "application/json",
                    "Authorization": f"Bearer {access_token}",
                    "xero-tenant-id": tenant_id,
                },
            )

            # response.raise_for_status()

            # Check the HTTP status code
            if response.status_code != 200:
                print("HTTP Status Code:", response.status_code)
                return {"Success": False, "Error": "Non-200 Status Code"}

            # print("response.text", response.text)

            python_dict = xmltodict.parse(response.text)

            python_dict = python_dict['Response']['Reports']['Report']['Rows']['Row']
            pprint.pprint(python_dict)

            return {"Success": True}

    except Exception as e:
        print("Error", e)
        return {"Success": False, "Error": str(e)}
