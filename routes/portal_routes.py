import copy
import csv
import os
import secrets
from datetime import datetime
# from random import random

# from pprint import pprint

from bson import ObjectId
from fastapi import APIRouter, Request
from fastapi.responses import FileResponse
from portal_statement_files.portal_statement_create import create_pdf
# import pandas_datareader.data as web
import pandas as pd
import datetime
import yfinance as yf
# import investpy


import time

import smtplib
from email.message import EmailMessage

from config.db import db

import vonage

# API_KEY = "5f79bce5"
# API_SECRET = "dtNN0iQpgit3I6bZ"

# GET API_KEY AND API_SECRET FROM ENVIRONMENT VARIABLES
API_KEY = os.environ.get("VONAGE_API_KEY")
API_SECRET = os.environ.get("VONAGE_API_SECRET")

portal_info = APIRouter()

# MONGO COLLECTIONS
portalUsers = db.portalUsers
investors = db.investors

rates = db.rates


# opportunities = db.opportunities
# sales_parameters = db.salesParameters
# rollovers = db.investorRollovers


@portal_info.post("/create_portal_statements")
async def create_portal_statements(data: Request):
    start_time = time.time()
    request = await data.json()
    data = request["data"]
    # make a deep copy of the data
    data1 = copy.deepcopy(data)
    # print(data)

    type1 = create_pdf(1, data)
    type2 = create_pdf(2, data1)
    # return create_pdf(data)
    end_time = time.time()

    # show time taken in seconds rounded to two decimal places

    print(f"Time taken: {round(end_time - start_time, 2)} seconds")

    return {"type1": type1, "type2": type2, "time": round(end_time - start_time, 2)}


@portal_info.get("/get_investor_statement")
async def sales_forecast(investor_statement_name):
    try:
        file_name = investor_statement_name
        dir_path = "portal_statements"
        dir_list = os.listdir(dir_path)
        if file_name in dir_list:
            return FileResponse(f"{dir_path}/{file_name}", filename=file_name)
        else:
            return {"ERROR": "File does not exist!!"}
    except Exception as e:
        print(e)
        return {"ERROR": "Please Try again"}


@portal_info.post("/generate_otp")
async def generate_otp(data: Request):
    request = await data.json()
    # print(request)
    pin = secrets.randbelow(10 ** 6)
    # keep repeating the above until the pin is 6 digits
    while len(str(pin)) != 6:
        pin = secrets.randbelow(10 ** 6)

    if request['method'] == "email":

        email = request['email'].split(":")[1]
        # trim all white spaces
        email = email.strip()

        smtp_server = "depro8.fcomet.com"
        port = 465  # For starttls
        sender_email = 'omh-app@opportunitymanagement.co.za'
        password = "12071994Wb!"

        message = f"""\
            <html>
              <body>
                <p>Good Day,<br> 
                <br /><br />       
                    <strong>{pin}</strong> is your OTP to change your password on the OMH Portal. <br><br>             
              </body>
            </html>
            """

        msg = EmailMessage()
        msg['Subject'] = "OMH Portal - OTP"
        msg['From'] = sender_email
        msg['To'] = email
        msg.set_content(message, subtype='html')

        try:
            with smtplib.SMTP_SSL(smtp_server, port) as server:
                server.ehlo()
                server.login(sender_email, password=password)
                server.send_message(msg)
                server.quit()
                return {"message": "Email sent successfully", "otp": str(pin)}
        except Exception as e:
            print("Error:", e)
            return {"message": "Email not sent"}

    elif request['method'] == "SMS":
        try:
            to_number = request['mobile'].split(":")[1]
            # trim all white spaces
            to_number = to_number.strip()
            # remove all non-numeric characters, remove the first 0 and insert 27 at the beginning
            to_number = "27" + to_number[1:]
            # remove all spaces from to_number
            to_number = to_number.replace(" ", "")
            # print(to_number)

            client = vonage.Client(key=API_KEY, secret=API_SECRET)
            sms = vonage.Sms(client)

            responseData = sms.send_message(
                {
                    "from": "OMH",
                    # "from": "Vonage APIs",
                    "to": to_number,
                    "text": f"OMH OTP - {pin}",
                }
            )

            if responseData["messages"][0]["status"] == "0":
                return {"message": "SMS sent successfully", "otp": str(pin)}
                # print("Message sent successfully.")
                # print("RESPONSE DATA", responseData['messages'][0])
            else:
                # print(f"Message failed with error: {responseData['messages'][0]['error-text']}")
                return {"message": "SMS not sent XXX"}
            # print(responseData)
        except Exception as e:
            print("Error:", e)
            return {"message": "SMS not sent"}


@portal_info.post("/add_investor")
async def add_investor(data: Request):
    request = await data.json()
    # print(request)
    try:
        # get investor from investors collection where _id = request['id'] as an objectId and project only the
        # investor_mobile field
        investor = investors.find_one({"_id": ObjectId(request['id'])}, {"investor_mobile": 1, "_id": 0})

        mobile = investor['investor_mobile']

        # print(mobile)

        # create a variable called password made up of 20 random characters
        password = secrets.token_urlsafe(20)

        # print(password)

        insert = {
            "name": request['name'],
            "surname": request['surname'],
            "email": request['email'],
            "password": password,
            "investor_id": request['id'],
            "role": "INVESTOR",
            "mobile": mobile,
            # add "created_at" and "updated_at" fields to the document with the current date formatted as YYYY-MM-DD
            "created_at": datetime.datetime.now().strftime("%Y-%m-%d"),
            "updated_at": datetime.datetime.now().strftime("%Y-%m-%d"),
            "investor_acc_number": request['acc_number']
        }

        # print(insert)

        # insert the document into the portalUsers collection
        response = portalUsers.insert_one(insert)
        print(response)

        # send email to the investor
        smtp_server = "depro8.fcomet.com"
        port = 465  # For starttls
        sender_email = 'omh-app@opportunitymanagement.co.za'
        password = "12071994Wb!"

        message = f"""\
        <html>
          <body>
            <p>Dear {request['name']},<br><br>
               As a valued Investor, you have access to our "Portal" at <a href="https://www.opconlineportal.co.za">
               https://www.opconlineportal.co.za</a>.<br><br>
               
               I would strongly suggest saving the link to your favourites.<br><br>
              
               Please go to the above link, click on the red button to change your password and follow the 
               instructions.<br><br>
               
               You will then be re-directed to the login page where you can enter your email address and the
                password you chose.<br><br>
                
                The Portal is a secure site and you can see "Live" information on your investments.<br><br>
                   
                  <br><br>
                  <u>Please do not reply to this email as it is not monitored</u>.<br><br>
                  <br /><br />
                  Regards<br />
                  <strong>The OMH Team</strong><br />
            </p>
          </body>
        </html>
        """

        msg = EmailMessage()
        msg['Subject'] = "OMH Portal - Login Details"
        msg['From'] = sender_email
        msg['To'] = request['email']
        # copy in nick@opportunity.co.za, wynand@capeprojects.co.za, debbie@opportunity.co.za and
        # dirk@cpconstruction.co.za to the email
        # msg['Cc'] = "nick@opportunity.co.za, wynand@capeprojects.co.za, debbie@opportunity.co.za, " \
        #             "dirk@cpconstruction.co.za"

        msg.set_content(message, subtype='html')

        try:
            with smtplib.SMTP_SSL(smtp_server, port) as server:
                server.ehlo()
                server.login(sender_email, password=password)
                server.send_message(msg)
                server.quit()
                return {"message": "Email sent successfully"}
        except Exception as e:
            print("Error:", e)
            return {"message": "Email not sent"}

        # return {"awesome": "It Works"}
    except Exception as e:

        return {"ERROR": f"Please Try again: {e}"}


@portal_info.post("/investment_termination")
async def investment_termination(data: Request):
    request = await data.json()
    # print(request)
    try:
        # get investor from investors collection where _id = request['id'] as an objectId and project only the
        # investor_mobile field
        investor = investors.find_one({"_id": ObjectId(request['_id'])},
                                      {"investor_email": 1, "investor_name": 1, "investor_surname": 1, "_id": 0})
        #
        #     mobile = investor['investor_mobile']
        #
        # print("investor", investor)

        investor_name = f"{investor['investor_name']} {investor['investor_surname']}"

        investor_email = f"{investor['investor_email']}, debbie@opportunity.co.za; leandri@opportunity.co.za"

        option_chosen = ""
        exit_amount = ""
        rollover_amount = ""

        if request['full_exit']:
            option_chosen = "Full Exit"
            # variable exit_amount = request['float_balance'] formatted as currency with 2 decimal places and a space
            # between the currency symbol and the amount. e.g. R 1 000.00
            exit_amount = f"R {request['float_balance']:,.2f}"
            # rollover_amount = 0 formatted as above
            rollover_amount = f"R 0.00"
        elif request['full_rollover']:
            option_chosen = "Full Rollover"
            # exit_amount = 0 formatted as above
            exit_amount = f"R 0.00"
            # rollover_amount = request['float_balance'] formatted as above
            rollover_amount = f"R {request['float_balance']:,.2f}"
        else:
            if request['partial_exit']:
                option_chosen = "Partial Exit"
                # exit_amount = request['float_investment'] formatted as above
                exit_amount = f"R {request['float_exit']:,.2f}"
                # rollover_amount = request['float_balance'] - request['float_exit'] formatted as above
                rollover_amount = f"R {request['float_balance'] - request['float_exit']:,.2f}"

        # send email to the investor
        smtp_server = "depro8.fcomet.com"
        port = 465  # For starttls
        sender_email = 'omh-app@opportunitymanagement.co.za'
        password = "12071994Wb!"
        #
        message = f"""\
        <html>
          <body>
            <p>Dear {investor_name},<br><br>
            
                We have received your prefered options when your investment in {request['opportunity_code']} is 
                transferred.<br><br>
                
                Based on Today's date, your investment balance is R {request['float_balance']:,.2f}.<br><br>
                
                {request['opportunity_code']}<br><br>
                <strong>Option Chosen:</strong> {option_chosen}<br><br>
                <strong>Exit Amount:</strong> {exit_amount}<br><br>
                <strong>Rollover Amount:</strong> {rollover_amount}<br><br>
                
                We will be in contact with you shortly to finalise the transfer.<br><br>
                
                
                <strong>Fixed Investor Returns:</strong> The projected returns below are applicable to the next 
                investment cycle. Please note that these are subject to change and should this occur, Opportunity 
                will communicate this to you. <br /><br /> 
                
                <ol> 
                <li>Investments of R 100 000 – R 499 000 @ 14% p.a</li> 
                <li>Investments of R 500 000 – R 999 000 @ 16% p.a</li> 
                <li>Investments of R 1 000 000 upwards @ 18% p.a</li>
                </ol>
                
                <br /><br />
                
                Please do not reply to this email as it is not monitored. Any questions can be directed to Debbie 
                Landsberg at debbie@opportunity.co.za<br /><br />
                
                Kind Regards<br />
                
                <strong>The OMH Team</strong><br />
                
            </p>
          </body>
        </html>
        """

        msg = EmailMessage()
        msg['Subject'] = "OMH Portal - Investment Termination provisional confirmation"
        msg['From'] = sender_email
        msg['To'] = investor_email
        # copy in nick@opportunity.co.za, wynand@capeprojects.co.za, debbie@opportunity.co.za and
        # dirk@cpconstruction.co.za to the email
        # msg['Cc'] = "nick@opportunity.co.za, wynand@capeprojects.co.za, debbie@opportunity.co.za, " \
        #             "dirk@cpconstruction.co.za"

        msg.set_content(message, subtype='html')

        try:
            with smtplib.SMTP_SSL(smtp_server, port) as server:
                server.ehlo()
                server.login(sender_email, password=password)
                server.send_message(msg)
                server.quit()
                return {"message": "Email sent successfully"}
        except Exception as e:
            print("Error:", e)
            return {"message": "Email not sent"}

    except Exception as e:

        return {"ERROR": f"Please Try again: {e}"}


# EARLY RELEASES TO CSV
@portal_info.post("/early_releases")
async def early_releases():
    try:
        # get all documents from investors collection
        investors_list = investors.find()
        # print(investors_list)
        # create an empty list
        early_releases_list = []
        # loop through the investors_list
        for investor in investors_list:
            # loop through investments in investor
            for investment in investor['investments']:
                investment['investor_acc_number'] = investor['investor_acc_number']
                investment['investor_name'] = investor['investor_name']
                investment['investor_surname'] = investor['investor_surname']

                # insert investment into early_releases list
                early_releases_list.append(investment)

        # filter out all documents where the "early_release" field is not true or does not exist
        early_releases_list = list(filter(lambda x: x.get('early_release', False), early_releases_list))

        new_list = []

        for release in early_releases_list:
            insert = {
                "investor_acc_number": release['investor_acc_number'],
                "investor_name": release['investor_name'],
                "investor_surname": release['investor_surname'],
                "Category": release['Category'],
                "investment_amount": release['investment_amount'],
                "investment_exit_rollover": release['investment_exit_rollover'],
                "opportunity_code": release['opportunity_code'],
                "release_date": release['release_date'],
                "end_date": release['end_date'],
                "interest": release['interest'],
                "exit_value": release['exit_value'],
                "investment_number": release['investment_number'],
                "exit_amount": release['exit_amount'],
                "rollover_amount": release['rollover_amount'],
            }
            new_list.append(insert)

        # add early releases to a csv file called early_releases.csv
        with open('early_releases.csv', 'w', newline='') as csvfile:

            fieldnames = ['investor_acc_number', 'investor_name', 'investor_surname', 'Category', 'investment_amount',
                          'investment_exit_rollover', 'opportunity_code', 'release_date', 'end_date', 'interest',
                          'exit_value', 'investment_number', 'exit_amount', 'rollover_amount']

            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            writer.writeheader()
            for early_release in new_list:
                writer.writerow(early_release)

        return {"early_releases": early_releases_list}
    except Exception as e:
        return {"ERROR": f"Please Try again: {e}"}


# INVESTMENTS TO CSV
@portal_info.post("/investments_draws_taken")
async def investments_draws():
    try:
        # get all documents from investors collection
        investors_list = list(investors.find())
        # print(investors_list)
        for investor in investors_list:
            investor['_id'] = str(investor['_id'])

        # create an empty list
        investments_draws_list = []
        # loop through the investors_list
        for investor in investors_list:
            # loop through investments in investor
            for investment in investor['trust']:
                investment['investor_acc_number'] = investor['investor_acc_number']
                investment['investor_name'] = investor['investor_name']
                investment['investor_surname'] = investor['investor_surname']

                # insert investment into investments_draws list
                investments_draws_list.append(investment)

        # filter out all documents where the Category field equals "Southwark" or "SouthWark"
        investments_draws_list = list(filter(lambda x: x.get('Category', False) != "Southwark", investments_draws_list))
        investments_draws_list = list(filter(lambda x: x.get('Category', False) != "SouthWark", investments_draws_list))
        # print(investments_draws_list)
        print(len(investments_draws_list))
        new_list = []

        # loop through investments_draws_list and print each item followed by a blank line
        for investment in investments_draws_list:
            # print(investment)
            # print()
            insert = {
                "investor_acc_number": investment['investor_acc_number'],
                "investor_name": investment['investor_name'],
                "investor_surname": investment['investor_surname'],
                "Category": investment['Category'],
                "investment_amount": float(investment['investment_amount']),

                "opportunity_code": investment['opportunity_code'],
                "deposit_date": investment['deposit_date'],
                "release_date": investment['release_date'],

            }
            new_list.append(insert)

        # print(new_list)
        # add investments to a csv file called investments.csv
        with open('investments_drawn.csv', 'w', newline='') as csvfile:

            # create variable feildnames and assign it a list of all the feilds in insert above
            fieldnames = ['investor_acc_number', 'investor_name', 'investor_surname', 'Category', 'investment_amount',
                          'opportunity_code', 'deposit_date', 'release_date']

            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            writer.writeheader()
            for investment in new_list:
                writer.writerow(investment)

        return {"investments": new_list}
    except Exception as e:
        return {"ERROR": "Please Try again", "Error": e}


@portal_info.get("/stock_market")
async def stock_market():
    try:
        indexes = get_indices()
        # pprint("indexes", indexes)
        currencies = get_currency_data()
        # pprint("currencies", currencies)
        commodities = get_commodity_data()
        # pprint("commodities", commodities)

        # join the indexes, currencies and commodities lists
        stock_market_figures_retrieved = indexes + currencies + commodities
        # pprint("stock_market", stock_market)

        return {"stock_market": stock_market_figures_retrieved}

    except Exception as e:

        return {"ERROR": "Please Try again", "Error": e}


@portal_info.post("/getChartData")
async def get_chart_data(request: Request):
    try:

        data = await request.json()
        data = data['chartData']
        rates_list = list(rates.find())
        # print(data)

        # loop through rates_list and delete _id field, replace '-' with '/' in Efective_date and convert Efective_date
        # to datetime
        for rate in rates_list:
            rate['_id'] = str(rate['_id'])
            rate['Efective_date'] = rate['Efective_date'].replace('-', '/')
            rate['Efective_date'] = datetime.datetime.strptime(rate['Efective_date'], "%Y/%m/%d")
            # convert rate['rate'] to float divide by 100
            rate['rate'] = float(rate['rate']) / 100

        # sort rates_list by Efective_date descending
        rates_list = sorted(rates_list, key=lambda x: x['Efective_date'], reverse=True)

        investments_to_chart = []
        for investment in data:
            insert = {
                "opportunity_code": investment['opportunity_code'],
                "investment_number": investment['investment_number'],
            }
            investments_to_chart.append(insert)

        investments_to_chart_opportunity_codes = []
        for investment in investments_to_chart:
            investments_to_chart_opportunity_codes.append(investment['opportunity_code'])
        # Ensure only unique opportunity codes are in the list
        investments_to_chart_opportunity_codes = list(set(investments_to_chart_opportunity_codes))

        # get info from investors collection where investor_acc_number equals data[0]['investor_acc_number'],
        # project only the investments array
        investors_list = list(investors.aggregate([
            {"$match": {"investor_acc_number": data[0]['investor_acc_number']}},
            {"$project": {"investments": 1, "trust": 1, "_id": 0}}
        ]))[0]
        # print("trust_list", investors_list['trust'])
        trust_list = investors_list['trust']
        investors_list = investors_list['investments']

        # filter investors_list to only include investments where the opportunity_code is in
        # investments_to_chart_opportunity_codes
        investors_list = list(
            filter(lambda x: x.get('opportunity_code', False) in investments_to_chart_opportunity_codes,
                   investors_list))

        # filter trust_list to only include investments where the opportunity_code is in
        # investments_to_chart_opportunity_codes
        trust_list = list(
            filter(lambda x: x.get('opportunity_code', False) in investments_to_chart_opportunity_codes,
                   trust_list))

        # loop through investments_list
        for investment in investors_list:
            # filter trust_list to only include investments where the opportunity_code is equal to
            # investment['opportunity_code'] and the investment_number is equal to investment['investment_number'] and
            # save the result to a variable called filtered_trust_list
            filtered_trust_list = list(
                filter(lambda x: x.get('opportunity_code', False) == investment['opportunity_code'] and x.get(
                    'investment_number', False) == investment['investment_number'],
                       trust_list))
            # make investment['deposit_date'] equal to the deposit_date of the first item in filtered_trust_list
            investment['deposit_date'] = filtered_trust_list[0]['deposit_date']

        finalised_chart_data = []

        # loop through investments_to_chart
        for investment in investments_to_chart:
            # filter investors_list to only include investments where the opportunity_code is equal to
            # investment['opportunity_code'] and the investment_number is equal to investment['investment_number'] and
            # save the result to a variable called filtered_investments_list
            filtered_investments_list = list(
                filter(lambda x: x.get('opportunity_code', False) == investment['opportunity_code'] and x.get(
                    'investment_number', False) == investment['investment_number'],
                       investors_list))
            # print("filtered_investments_list", filtered_investments_list)
            # create a variable called investment_amount and assign it the value of the investment_amount in
            # filtered_investments_list as a float
            investment_amount = float(filtered_investments_list[0]['investment_amount'])

            # create a variable called investment_interest_rate and assign it the value of the
            # investment_interest_rate as a float divided by 100
            investment_interest_rate = float(filtered_investments_list[0]['investment_interest_rate']) / 100
            # convert investment['deposit_date'] to a datetime object and save it to a variable called deposit_date
            # replace '-' with '/' in investment['deposit_date']

            deposit_date = filtered_investments_list[0]['deposit_date'].replace('-', '/')
            deposit_date = datetime.datetime.strptime(deposit_date, "%Y/%m/%d")

            # do the same as above for release_date and end_date
            release_date = filtered_investments_list[0]['release_date'].replace('-', '/')
            release_date = datetime.datetime.strptime(release_date, "%Y/%m/%d")

            end_date = filtered_investments_list[0]['end_date'].replace('-', '/')
            end_date = datetime.datetime.strptime(end_date, "%Y/%m/%d")

            momentum_interest = 0.0
            total_days_invested = end_date - deposit_date

            total_released_days = end_date - release_date

            total_released_interest = investment_amount * investment_interest_rate * (total_released_days.days / 365)

            # add 1 day to the deposit_date
            deposit_date = deposit_date + datetime.timedelta(days=1)

            while deposit_date <= release_date:
                # filter rates_list to only include rates where the date is less than or equal to deposit_date
                filtered_rates_list = list(filter(lambda x: x.get('Efective_date', False) <= deposit_date, rates_list))
                # create a variable called rate and assign it the value of the rate of the first item in
                # filtered_rates_list
                rate = filtered_rates_list[0]['rate']
                # calculate the days interest and add it to momentum_interest
                momentum_interest += investment_amount * rate * (1 / 365)
                # add 1 day to deposit_date
                deposit_date = deposit_date + datetime.timedelta(days=1)

            final_interest = total_released_interest + momentum_interest

            # annualised_interest = final_interest / total_days_invested.days * 365

            return_on_investment = final_interest / investment_amount * 100

            annualised_interest_rate = return_on_investment / total_days_invested.days * 365
            # round annualised_interest_rate to 1 decimal place
            # annualised_interest_rate = round(annualised_interest_rate, 1)

            insert = {
                "opportunity_code": investment['opportunity_code'],
                "investment_number": investment['investment_number'],
                "investment_amount": investment_amount,
                "final_interest": final_interest,
                "return_on_investment": round(return_on_investment, 1),
                "annualised_interest_rate": round(annualised_interest_rate, 1),
            }
            finalised_chart_data.append(insert)

        return {"final_chart_data": finalised_chart_data}

    except Exception as e:
        return {"ERROR": "Please Try again", "Error": e}


def get_currency_data():
    try:

        jse_tickers2 = ['ZAR=X', 'ZARGBP=X', 'ZAREUR=X', 'ZARAUD=X', 'ZARJPY=X', 'ZARCNY=X']

        jse_data = yf.download(jse_tickers2)
        # jse_data2 = jse_data.T

        # print("Currency Data", jse_data2)

        if pd.isna(jse_data['Close']['ZAR=X'][-1]):
            usd_zar_rate = jse_data['Close']['ZAR=X'][-2]
            if jse_data['Close']['ZAR=X'][-3] > jse_data['Close']['ZAR=X'][-2]:
                icon = "arrow_drop_up"
                icon_color = "green"
            else:
                icon = "arrow_drop_down"
                icon_color = "white"
        else:
            usd_zar_rate = jse_data['Close']['ZAR=X'][-1]
            if jse_data['Close']['ZAR=X'][-2] > jse_data['Close']['ZAR=X'][-1]:
                icon = "arrow_drop_up"
                icon_color = "green"
            else:
                icon = "arrow_drop_down"
                icon_color = "white"

        if pd.isna(jse_data['Close']['ZARGBP=X'][-1]):
            gbp_zar_rate = jse_data['Close']['ZARGBP=X'][-2]
            if jse_data['Close']['ZARGBP=X'][-3] < jse_data['Close']['ZARGBP=X'][-2]:
                icon2 = "arrow_drop_up"
                icon_color2 = "green"
            else:
                icon2 = "arrow_drop_down"
                icon_color2 = "white"
        else:
            gbp_zar_rate = jse_data['Close']['ZARGBP=X'][-1]
            if jse_data['Close']['ZARGBP=X'][-2] < jse_data['Close']['ZARGBP=X'][-1]:
                icon2 = "arrow_drop_up"
                icon_color2 = "green"
            else:
                icon2 = "arrow_drop_down"
                icon_color2 = "white"
        gbp_zar_rate = 1 / gbp_zar_rate

        if pd.isna(jse_data['Close']['ZAREUR=X'][-1]):
            eur_zar_rate = jse_data['Close']['ZAREUR=X'][-2]
            if jse_data['Close']['ZAREUR=X'][-3] < jse_data['Close']['ZAREUR=X'][-2]:
                icon3 = "arrow_drop_up"
                icon_color3 = "green"
            else:
                icon3 = "arrow_drop_down"
                icon_color3 = "white"

        else:
            eur_zar_rate = jse_data['Close']['ZAREUR=X'][-1]
            if jse_data['Close']['ZAREUR=X'][-2] < jse_data['Close']['ZAREUR=X'][-1]:
                icon3 = "arrow_drop_up"
                icon_color3 = "green"
            else:
                icon3 = "arrow_drop_down"
                icon_color3 = "white"
        eur_zar_rate = 1 / eur_zar_rate

        if pd.isna(jse_data['Close']['ZARAUD=X'][-1]):
            aud_zar_rate = jse_data['Close']['ZARAUD=X'][-2]
            if jse_data['Close']['ZARAUD=X'][-3] < jse_data['Close']['ZARAUD=X'][-2]:
                icon4 = "arrow_drop_up"
                icon_color4 = "green"
            else:
                icon4 = "arrow_drop_down"
                icon_color4 = "white"

        else:
            aud_zar_rate = jse_data['Close']['ZARAUD=X'][-1]
            if jse_data['Close']['ZARAUD=X'][-2] < jse_data['Close']['ZARAUD=X'][-1]:
                icon4 = "arrow_drop_up"
                icon_color4 = "green"
            else:
                icon4 = "arrow_drop_down"
                icon_color4 = "white"
        aud_zar_rate = 1 / aud_zar_rate

        if pd.isna(jse_data['Close']['ZARJPY=X'][-1]):
            jpy_zar_rate = jse_data['Close']['ZARJPY=X'][-2]
            if jse_data['Close']['ZARJPY=X'][-3] < jse_data['Close']['ZARJPY=X'][-2]:
                icon5 = "arrow_drop_up"
                icon_color5 = "green"
            else:
                icon5 = "arrow_drop_down"
                icon_color5 = "white"
        else:
            jpy_zar_rate = jse_data['Close']['ZARJPY=X'][-1]
            if jse_data['Close']['ZARJPY=X'][-2] < jse_data['Close']['ZARJPY=X'][-1]:
                icon5 = "arrow_drop_up"
                icon_color5 = "green"
            else:
                icon5 = "arrow_drop_down"
                icon_color5 = "white"
        jpy_zar_rate = 1 / jpy_zar_rate

        if pd.isna(jse_data['Close']['ZARCNY=X'][-1]):
            cny_zar_rate = jse_data['Close']['ZARCNY=X'][-2]
            if jse_data['Close']['ZARCNY=X'][-3] < jse_data['Close']['ZARCNY=X'][-2]:
                icon6 = "arrow_drop_up"
                icon_color6 = "green"
            else:
                icon6 = "arrow_drop_down"
                icon_color6 = "white"
        else:
            cny_zar_rate = jse_data['Close']['ZARCNY=X'][-1]
            if jse_data['Close']['ZARCNY=X'][-2] < jse_data['Close']['ZARCNY=X'][-1]:
                icon6 = "arrow_drop_up"
                icon_color6 = "green"
            else:
                icon6 = "arrow_drop_down"
                icon_color6 = "white"
        cny_zar_rate = 1 / cny_zar_rate

        rates_returned = [
            {'Description': 'USD', 'price': usd_zar_rate.round(4), 'color': 'rgb(150, 0, 0)', 'icon': icon,
             'icon_color': icon_color},
            {'Description': 'GBP', 'price': gbp_zar_rate.round(4), 'color': 'rgb(170, 0, 0)', 'icon': icon2,
             'icon_color': icon_color2},
            {'Description': 'EUR', 'price': eur_zar_rate.round(4), 'color': 'rgb(190, 0, 0)', 'icon': icon3,
             'icon_color': icon_color3},
            {'Description': 'AUD', 'price': aud_zar_rate.round(4), 'color': 'rgb(210, 0, 0)', 'icon': icon4,
             'icon_color': icon_color4},
            {'Description': 'JPY', 'price': jpy_zar_rate.round(4), 'color': 'rgb(230, 0, 0)', 'icon': icon5,
             'icon_color': icon_color5},
            {'Description': 'CNY', 'price': cny_zar_rate.round(4), 'color': 'rgb(250, 0, 0)', 'icon': icon6,
             'icon_color': icon_color6}]

        return rates_returned

    except Exception as e:
        print("Error getting exchange rates", e)
        return []


def get_commodity_data():
    try:
        tickers = ['GC=F', 'SI=F', 'PL=F', 'BZ=F']  # Tickers for gold, silver, platinum, palladium, and Brent crude
        commodities = yf.download(tickers, period='10d',
                                  interval='1d')  # Download commodity data for the past day at 1-minute

        gold_price = commodities['Close']['GC=F'][-1]
        if commodities['Open']['GC=F'][-1] < commodities['Close']['GC=F'][-1]:
            icon = "arrow_drop_up"
            icon_color = "white"
        else:
            icon = "arrow_drop_down"
            icon_color = "red"

        silver_price = commodities['Close']['SI=F'][-1]
        if commodities['Open']['SI=F'][-1] < commodities['Close']['SI=F'][-1]:
            icon2 = "arrow_drop_up"
            icon_color2 = "white"
        else:
            icon2 = "arrow_drop_down"
            icon_color2 = "red"
        platinum_price = commodities['Close']['PL=F'][-1]
        if commodities['Open']['PL=F'][-1] < commodities['Close']['PL=F'][-1]:
            icon3 = "arrow_drop_up"
            icon_color3 = "white"
        else:
            icon3 = "arrow_drop_down"
            icon_color3 = "red"
        brent_price = commodities['Close']['BZ=F'][-1]
        if commodities['Open']['BZ=F'][-1] < commodities['Close']['BZ=F'][-1]:
            icon4 = "arrow_drop_up"
            icon_color4 = "white"
        else:
            icon4 = "arrow_drop_down"
            icon_color4 = "red"

        # commodities2 = commodities.T
        # print("commodities2", commodities2)

        commodities_collected = [
            {'Description': 'Gold', 'price': f"${gold_price:,.2f}", 'color': "rgb(0, 100, 0)", 'icon': icon,
             'icon_color': icon_color},
            {'Description': 'Silver', 'price': f"${silver_price:,.2f}", 'color': "rgb(0, 130, 0)", 'icon': icon2,
             'icon_color': icon_color2},
            {'Description': 'Platinum', 'price': f"${platinum_price:,.2f}",
             'color': "rgb(0, 160, 0)", 'icon': icon3, 'icon_color': icon_color3},
            {'Description': 'Brent Crude', 'price': f"${brent_price:,.2f}",
             'color': "rgb(0, 190, 0)", 'icon': icon4, 'icon_color': icon_color4}]

        return commodities_collected

    except Exception as e:
        print("Error getting commodity data", e)
        return []


# get_commodity_data()


def get_indices():
    try:
        start_date = datetime.datetime.today() - datetime.timedelta(days=30)

        end_date = datetime.datetime.today()

        indices = ['^FTSE', '^DJI', '^GSPC', '^HSI', '^IXIC', 'JSE.JO']

        indices_collected = []

        # Retrieve the historical data for each index and store it in a dictionary
        index_data = {}
        for index in indices:
            data = yf.download(index, start=start_date, end=end_date)
            index_data[index] = data

        # Print the closing prices for each index
        # range_start = 13
        for index, data in index_data.items():
            # start at a number and let it descend from say 13 to 1
            if index == 'JSE.JO':
                description = 'JSE'
                color = "rgb(0, 80, 160)"

            elif index == '^FTSE':
                description = 'FTSE 100'
                color = "rgb(0, 80, 70)"
            elif index == '^DJI':
                description = 'Dow Jones Industrial Average'
                color = "rgb(0, 80, 90)"
            elif index == '^GSPC':
                description = 'S&P 500'
                color = "rgb(0, 80, 120)"
            elif index == '^HSI':
                description = 'Hang Seng'
                color = "rgb(0, 80, 150)"
            elif index == '^IXIC':
                description = 'NASDAQ'
                color = "rgb(0, 80, 180)"
            else:
                description = index
                color = "rgb(0, 80, 160)"

            change = data['Close'][-1] - data['Open'][-1]

            if change > 0:
                icon = "arrow_drop_up"
                icon_color = "green"
            elif change < 0:
                icon = "arrow_drop_down"
                icon_color = "red"
            else:
                icon = "minimize"
                icon_color = "grey"

            price = f"{data['Close'][-1]:,.2f}"

            indices_collected.append({'index': index, 'Description': description, 'price': price, 'change': change,
                                      'color': color, 'icon': icon, 'icon_color': icon_color})
            # range_start -= 1

        return indices_collected

    except Exception as e:
        print("Error getting indices", e)
        return []

# get_stock_list()


# get_stock_data()
# get_cpi()

# indexes = get_indices()
# print(indexes)
# currencies = get_currency_data()
# print(currencies)
# commodities = get_commodity_data()
# print(commodities)

