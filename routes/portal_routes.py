import copy
import csv
import os
import secrets
from datetime import datetime

from bson import ObjectId
from fastapi import APIRouter, Request
from fastapi.responses import FileResponse
from portal_statement_files.portal_statement_create import create_pdf

import time

import smtplib
from email.message import EmailMessage

from config.db import db

portal_info = APIRouter()

# MONGO COLLECTIONS
portalUsers = db.portalUsers
investors = db.investors


# rates = db.rates
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

    pin = secrets.randbelow(10 ** 6)
    # keep repeating the above until the pin is 6 digits
    while len(str(pin)) != 6:
        pin = secrets.randbelow(10 ** 6)
    # print(pin)
    # print(request['email'])
    email = request['email'].split(":")[1]
    # trim all white spaces
    email = email.strip()
    print("email", email)

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

    # with open('excel_files/Sales ForecastHeron.xlsx', 'rb') as f:
    #     file_content = f.read()
    # file_name = os.path.basename('Sales ForecastHeron.xlsx')
    # msg.add_attachment(file_content, maintype='application', subtype='octet-stream', filename=file_name)

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

    # return {"otp": str(pin)}


@portal_info.post("/add_investor")
async def add_investor(data: Request):
    request = await data.json()
    # print(request)
    try:
        # get investor from investors collection where _id = request['id'] as an objectId and project only the
        # investor_mobile field
        investor = investors.find_one({"_id": ObjectId(request['id'])}, {"investor_mobile": 1, "_id": 0})

        mobile = investor['investor_mobile']

        # create a variable called password made up of 20 random characters
        password = secrets.token_urlsafe(20)

        insert = {
            "name": request['name'],
            "surname": request['surname'],
            "email": request['email'],
            "password": password,
            "investor_id": request['id'],
            "role": "INVESTOR",
            "mobile": mobile,
            # add "created_at" and "updated_at" fields to the document with the current date formatted as YYYY-MM-DD
            "created_at": datetime.now().strftime("%Y-%m-%d"),
            "updated_at": datetime.now().strftime("%Y-%m-%d"),
            "investor_acc_number": request['acc_number']
        }

        # insert the document into the portalUsers collection
        response = portalUsers.insert_one(insert)

        # send email to the investor
        smtp_server = "depro8.fcomet.com"
        port = 465  # For starttls
        sender_email = 'omh-app@opportunitymanagement.co.za'
        password = "12071994Wb!"

        message = f"""\
        <html>
          <body>
            <p>Dear {request['name']},<br><br>
               As a valued Investor, you have access to our "Portal" at <a href="https://www.eccentrictoad.com">
               https://www.eccentrictoad.com</a>.<br><br>
               
               I would strongly suggest saving the link to your favourites.<br><br>
              
               Please go to the above link, click on the red button to change your password and follow the instructions.<br><br>
               
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

        return {"ERROR": "Please Try again"}


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
        return {"ERROR": "Please Try again"}


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
        return {"ERROR": "Please Try again"}
