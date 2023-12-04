from bson.objectid import ObjectId
from fastapi import APIRouter, Request, BackgroundTasks
# from fastapi.encoders import jsonable_encoder
from apscheduler.schedulers.background import BackgroundScheduler

from config.db import db

import random

import smtplib
from email.message import EmailMessage
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import imaplib
import email
import os
from decouple import config
from datetime import datetime
# from distutils.command.clean import clean
from email.header import decode_header
import re

leads = APIRouter()


# get AWS_BUCKET_NAME from .env file
# SMTP_SERVER = config("SMTP_SERVER")
# print("SMTP_SERVER", SMTP_SERVER)


@leads.post("/post_sales_lead_form")
async def post_sales_lead_form(background_tasks: BackgroundTasks, data: Request):
    request = await data.json()
    sales_people = list(db.lead_sales_people.find({"active": True}))
    for person in sales_people:
        person["_id"] = str(person["_id"])

    if last_leads_generated := list(
            db.leads_sales.find().sort("created_at", -1).limit(len(sales_people))
    ):
        sales_people_ids_with_recent_leads = {lead["sales_person_id"] for lead in last_leads_generated}
        # Choose from sales people who haven't received a lead recently, or the last one if all have.
        sales_people_to_choose = [person for person in sales_people if
                                  person["_id"] not in sales_people_ids_with_recent_leads] or [sales_people[-1]]
        sales_person = random.choice(sales_people_to_choose)
    else:
        sales_person = random.choice(sales_people)

    request.update({
        'origin': request['source']['source'],
        'type': request['source']['type'],
        'created_at': datetime.now(),
        "sales_person": f"{sales_person['name']} {sales_person['surname']}",
        "sales_person_id": sales_person["_id"]
    })
    del request['source']

    db.leads_sales.insert_one(request)
    background_tasks.add_task(send_email_to_sales_person, sales_person, request)
    background_tasks.add_task(send_email_to_sales_lead, sales_person, request)

    return {"message": "success", "sales_person": request["sales_person"]}


@leads.post("/post_investments_lead_form")
async def post_investments_lead_form(background_tasks: BackgroundTasks, data: Request):
    request = await data.json()
    active_consultants_cursor = db.lead_investment_consultants.find({"active": True})
    consultants = [{**person, "_id": str(person["_id"])} for person in active_consultants_cursor]

    last_leads_cursor = db.leads_investments.find().sort("created_at", -1).limit(len(consultants))
    if last_leads_generated := list(last_leads_cursor):
        consultants_to_not_choose = {lead["consultant_id"] for lead in last_leads_generated}
        consultants_people_to_choose = [person for person in consultants if
                                        person["_id"] not in consultants_to_not_choose] or [consultants[-1]]
        consultant_person = random.choice(consultants_people_to_choose)
    else:
        consultant_person = random.choice(consultants)

    request.update({
        'origin': request['source']['source'],
        'type': request['source']['type'],
        'created_at': datetime.now(),
        "consultant": f"{consultant_person['name']} {consultant_person['surname']}",
        "consultant_id": consultant_person["_id"]
    })

    db.leads_investments.insert_one(request)
    background_tasks.add_task(send_email_to_consultant, consultant_person, request)
    background_tasks.add_task(send_email_to_investment_lead, consultant_person, request)

    return {"message": "success", "consultant": request["consultant"]}


@leads.post("/opportunity_contact_form")
async def opportunity_contact_form(data: Request):
    request = await data.form()
    print("request", request)

    try:

        consultants = list(db.lead_investment_consultants.find({"active": True}))
        length = len(consultants)
        consultants_to_not_choose = []
        for person in consultants:
            person["id"] = str(person["_id"])
            del person["_id"]

        last_leads_generated = list(db.leads_investments.find().sort("created_at", -1).limit(length))
        for lead in last_leads_generated:
            lead["id"] = str(lead["_id"])
            del lead["_id"]

        if len(last_leads_generated) == 0:
            consultant_person = random.choice(consultants)
        elif length > len(last_leads_generated) > 0:
            for lead in last_leads_generated:
                consultants_to_not_choose.append(lead["consultant_id"])
            consultants_people_to_choose = list \
                (filter(lambda x: x["id"] not in consultants_to_not_choose, consultants))
            consultant_person = random.choice(consultants_people_to_choose)
        else:
            consultants_people_to_choose = []
            for person in consultants:
                # create variable called filtered_leads where the sales_person_id is equal to the person id
                # filtered_leads = list(filter(lambda x: x["consultant_id"] == person["_id"], last_leads_generated))
                filtered_leads = list(filter(lambda x: x["consultant_id"] == person["id"], last_leads_generated))
                if len(filtered_leads) == 0:
                    consultants_people_to_choose.append(person["id"])
            if len(consultants_people_to_choose) == 0:
                consultants_people_to_choose.append \
                    (last_leads_generated[len(last_leads_generated) - 1]["consultant_id"])
            consultant_person = list(filter(lambda x: x["id"] in consultants_people_to_choose, consultants))[0]

        name = request.getlist("First Name")[0]
        surname = request.getlist("Last Name")[0]
        email = request.getlist("Email")[0]
        contact = request.getlist("Phone")[0]
        min_value = ""
        investment_amount = ""
        investment_choice = request.getlist('LEADCF10')[0]
        submission_date = datetime.now()
        # format submission_date as yyyy-mm-dd hh:mm:ss
        submission_date = submission_date.strftime("%Y-%m-%d %H:%M:%S")
        message = request["Description"]
        origin = 'Opportunity Website'
        type = 'Investments'
        created_at = datetime.now()
        data_investments = {
            "name": name,
            "surname": surname,
            "email": email,
            "contact": contact,
            "investment_choice": investment_choice,
            "min_value": min_value,
            "investment_amount": investment_amount,
            "submission_date": submission_date,
            "message": message,
            "origin": origin,
            "type": type,
            "created_at": created_at,
            "consultant": consultant_person["name"] + " " + consultant_person["surname"],
            "consultant_id": consultant_person["id"]
        }
        print("data", data_investments)

        db.leads_investments.insert_one(data_investments)

        consultant_email = send_email_to_consultant(consultant_person, data_investments)

        client_email = send_email_to_investment_lead(consultant_person, data_investments)

        return {"message": "success"}

        # return {"message": "success", "request": data_investments}
    except Exception as e:
        print("Error:", e)
        return {"message": "Lead not created"}


# @leads.post("/opportunityprop_contact_form")
# async def opportunityprop_contact_form(backgound_tasks: BackgroundTasks, data: Request):
#     request = await data.form()
#     print("request", request)
#
#     sales_people = list(db.lead_sales_people.find({"active": True}))
#     length = len(sales_people)
#     sales_people_to_not_choose = []
#     for person in sales_people:
#         person["_id"] = str(person["_id"])
#
#     last_leads_generated = list(db.leads_sales.find().sort("created_at", -1).limit(length))
#     if len(last_leads_generated) == 0:
#         sales_person = random.choice(sales_people)
#     elif length > len(last_leads_generated) > 0:
#         for lead in last_leads_generated:
#             sales_people_to_not_choose.append(lead["sales_person_id"])
#         sales_people_to_choose = list(filter(lambda x: x["_id"] not in sales_people_to_not_choose, sales_people))
#         sales_person = random.choice(sales_people_to_choose)
#     else:
#         sales_people_to_choose = []
#         for person in sales_people:
#             # create variable called filtered_leads where the sales_person_id is equal to the person id
#             filtered_leads = list(filter(lambda x: x["sales_person_id"] == person["_id"], last_leads_generated))
#             if len(filtered_leads) == 0:
#                 sales_people_to_choose.append(person["_id"])
#         if len(sales_people_to_choose) == 0:
#             sales_people_to_choose.append(last_leads_generated[len(last_leads_generated) - 1]["sales_person_id"])
#         sales_person = list(filter(lambda x: x["_id"] in sales_people_to_choose, sales_people))[0]
#
#     name = request.getlist("form_fields[name]")[0]
#     surname = ""
#     contact = request.getlist("form_fields[field_c045127]")[0]
#     email = request.getlist("form_fields[email]")[0]
#     development = ""
#     submission_date = datetime.now()
#     # format submission_date as yyyy-mm-dd hh:mm:ss
#     submission_date = submission_date.strftime("%Y-%m-%d %H:%M:%S")
#     contact_time = "ASAP"
#     message = request.getlist("form_fields[field_16ad61e]")[0]
#     origin = 'OpportunityProp Website'
#     type = 'sales'
#     created_at = datetime.now()
#
#     data = {
#         "name": name,
#         "surname": surname,
#         "email": email,
#         "contact": contact,
#         "development": development,
#         "submission_date": submission_date,
#         "contact_time": contact_time,
#         "message": message,
#         "origin": origin,
#         "type": type,
#         "created_at": created_at,
#         "sales_person": sales_person["name"] + " " + sales_person["surname"],
#         "sales_person_id": sales_person["_id"]
#     }
#
#     # request = data
#
#     db.leads_sales.insert_one(data)
#
#     # sp_email = send_email_to_sales_person(sales_person, data)
#     # send email to sales person as background task
#     backgound_tasks.add_task(send_email_to_sales_person, sales_person, data)
#
#     # client_email = send_email_to_sales_lead(sales_person, data)
#     # send email to client as background task
#     backgound_tasks.add_task(send_email_to_sales_lead, sales_person, data)
#
#     print("opportunityprop_contact_form ZZ")
#     return {"message": "success"}
#

@leads.get('/get_sales_leads')
async def get_sales_leads():
    leads = list(db.leads_sales.find())
    default_values = {
        'action_taken': "",
        'action_taken_date_time': "",
        'comments': "",
        'previous_actions': [],
        'rating': 0,
        'purchased': "",
        'purchased_date': "",
        'purchased_unit': ""
    }

    for lead in leads:
        lead["_id"] = str(lead["_id"])
        lead["created_at"] = lead["created_at"].strftime("%Y-%m-%d %H:%M:%S")
        for key, default in default_values.items():
            lead[key] = lead.get(key, default)

    leads = sorted(leads, key=lambda x: x['created_at'], reverse=True)

    return {"message": "success", "leads": leads}


def format_datetime(dt):
    return dt.strftime("%Y-%m-%d %H:%M:%S")


@leads.get('/get_sales_people')
async def get_sales_people():
    sales_people = [
        {
            **person,
            "_id": str(person["_id"]),
            "created_at": format_datetime(person["created_at"]),
            "updated_at": format_datetime(person["updated_at"]),
            # "unavailable_from": format_datetime(person["unavailable_from"]),
            # "unavailable_to": format_datetime(person["unavailable_to"]),
        }
        for person in db.lead_sales_people.find()
    ]

    return {"message": "success", "sales_people": sales_people}


@leads.get('/get_consultants')
async def get_sales_people():
    consultants = list(db.lead_investment_consultants.find())
    for person in consultants:
        person["_id"] = str(person["_id"])
        # if person

        person["created_at"] = person["created_at"].strftime("%Y-%m-%d %H:%M:%S")
        person["updated_at"] = person["updated_at"].strftime("%Y-%m-%d %H:%M:%S")
        # person["unavailable_from"] = person["unavailable_from"].strftime("%Y-%m-%d %H:%M:%S")
        # person["unavailable_to"] = person["unavailable_to"].strftime("%Y-%m-%d %H:%M:%S")

    return {"message": "success", "sales_people": consultants}


@leads.get('/get_investment_leads')
async def get_investment_leads():
    leads = list(db.leads_investments.find())
    for lead in leads:
        # if lead['created_at'] is a str type then print the lead
        if type(lead['created_at']) == str:
            # print("lead", lead)
            # convert created_at to datetime object
            lead['created_at'] = datetime.fromisoformat(lead['created_at'])

        lead["_id"] = str(lead["_id"])
        lead["created_at"] = lead["created_at"].strftime("%Y-%m-%d %H:%M:%S")
        lead['action_taken'] = lead.get('action_taken', "")
        lead['action_taken_date_time'] = lead.get('action_taken_date_time', "")
        lead['comments'] = lead.get('comments', "")
        lead['previous_actions'] = lead.get('previous_actions', [])
        lead['rating'] = lead.get('rating', 0)
        lead['category'] = lead.get('category', "")
        if lead['min_value'] == "No":
            lead['investment_amount'] = "< R100 000"
        # lead['purchased'] = lead.get('purchased', "")
        # lead['purchased_date'] = lead.get('purchased_date', "")
        # lead['purchased_unit'] = lead.get('purchased_unit', "")

    # sort leads by created_at descending
    leads = sorted(leads, key=lambda x: x['created_at'], reverse=True)

    return {"message": "success", "leads": leads}


@leads.post('/edit_sales_lead')
async def edit_sales_lead(data: Request):
    request = await data.json()

    try:
        lead_id = request['_id']
        del request['_id']
        request['created_at'] = datetime.strptime(request['created_at'], "%Y-%m-%d %H:%M:%S")
        db.leads_sales.update_one({"_id": ObjectId(lead_id)}, {"$set": request})

        return {"message": "success"}

    except Exception as e:
        print("Error:", e)
        return {"message": "Lead not updated"}


@leads.post('/edit_investment_lead')
async def edit_investment_lead(background_tasks: BackgroundTasks, data: Request):
    request = await data.json()
    if request['category'] not in ["", "Category 1"]:
        background_tasks.add_task(send_email_to_leandri, request)

    try:
        lead_id = request['_id']
        del request['_id']
        # convert created_at to datetime object
        request['created_at'] = datetime.strptime(request['created_at'], "%Y-%m-%d %H:%M:%S")
        db.leads_investments.update_one({"_id": ObjectId(lead_id)}, {"$set": request})

        return {"message": "success"}

    except Exception as e:
        print("Error:", e)
        return {"message": "Lead not updated"}


def send_email_to_sales_person(sales_person, lead):
    email_sp = sales_person['email'].strip()
    smtp_server = config('SMTP_SERVER')
    port = config('SMTP_PORT')
    sender_email = config('SENDER_EMAIL')
    password = config('EMAIL_PASSWORD')

    message = f"""
                <html>
                  <body>
                    <p>Dear {sales_person['name']},<br>
                    <br>
                    A new sales lead has been generated. Please see the details below:<br>
                    <br>
                    <b>Lead Details</b><br>
                    <br>
                    <b>First Name:</b> {lead['name']}<br>
                    <b>Last Name:</b> {lead['surname']}<br>
                    <b>Cell:</b> {lead['contact']}<br>
                    <b>Email:</b> {lead['email']}<br>
                    <b>Origin:</b> {lead['origin']}<br>
                    <b>Development:</b> {lead['development']}<br>
                    <b>Best Time:</b> {lead['contact_time']}<br>
                    <b>Message:</b> {lead['message']}<br>
                    <br>
                    Please contact the lead as soon as possible and follow up by inputting into the app.<br>
                    <br>
                    Kind Regards,<br>
                    OMH App<br>
                  </body>
                </html>
                """

    msg = EmailMessage()
    msg['Subject'] = "Sales Lead"
    msg['From'] = sender_email
    msg['To'] = email_sp
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


def send_email_to_leandri(lead):
    email_sp = 'leandri@opportunity.co.za'
    # email_sp = 'wayne@opportunity.co.za'
    smtp_server = config('SMTP_SERVER')
    port = config('SMTP_PORT')
    sender_email = config('SENDER_EMAIL')
    password = config('EMAIL_PASSWORD')

    message = f"""
                <html>
                  <body>
                    <p>Dear Leandri,<br>
                    <br>
                    A new investment lead has been categorised as {lead['category']}:<br>
                    <br>
                    <b>Lead Details</b><br>
                    <br>
                    <b>First Name:</b> {lead['name']}<br>
                    <b>Last Name:</b> {lead['surname']}<br>
                    <b>Cell:</b> {lead['contact']}<br>
                    <b>Email:</b> {lead['email']}<br>
                    <b>Message:</b> {lead['message']}<br>
                    <br>
                    Kind Regards,<br>
                    OMH App<br>
                  </body>
                </html>
                """

    msg = EmailMessage()
    msg['Subject'] = "Investment Lead"
    msg['From'] = sender_email
    msg['To'] = email_sp
    msg.set_content(message, subtype='html')

    try:
        with smtplib.SMTP_SSL(smtp_server, port) as server:
            server.ehlo()
            server.login(sender_email, password=password)
            server.send_message(msg)
            return {"message": "Email sent successfully"}
    except Exception as e:
        print("Error:", e)
        return {"message": "Email not sent"}


def send_email_to_sales_lead(sales_person, lead):
    email_lead = lead['email']
    # # trim all white spaces
    email_lead = email_lead.strip()
    #
    smtp_server = config('SMTP_SERVER')
    port = config('SMTP_PORT')
    sender_email = config('SENDER_EMAIL')
    password = config('EMAIL_PASSWORD')

    image_attached = False

    message = f"""\
                    <html>
                      <body>
                        <p>Good Day {lead['name']},<br>
                        <br /><br />
                        Thank you for your enquiry about {lead['development']}.<br />
                        <br /><br />
                        <b>The following agent will be in contact with you shortly:</b><br />
                        <br /><br />
                        <b>Agent Details</b><br />
                        <b>First Name:</b> {sales_person['name']}<br />
                        <b>Last Name:</b> {sales_person['surname']}<br />
                        <b>Cell:</b> {sales_person['cell']}<br />
                        <b>Email:</b> {sales_person['email']}<br />
                        

                        <br /><br />
                        Please keep a lookout for their call.<br />

                        <br /><br />
                        Kind Regards,<br />
                        Opportunity Property<br />
                        

                      </body>
                    </html>
                    """

    # msg = EmailMessage()
    msg = MIMEMultipart()
    msg['Subject'] = "Opportunity Property"
    msg['From'] = sender_email
    msg['To'] = email_lead
    msg.attach(MIMEText(message, 'html'))
    # msg.set_content(message, subtype='html')

    with open("annexures/Signature OppProp.jpg", "rb") as image_file:
        image_data = image_file.read()
        image = MIMEImage(image_data, _subtype="jpg")

        image.add_header("Content-ID", "<image1>")
        msg.attach(image)

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


def send_email_to_consultant(consultant_person, lead):
    email_consultant = consultant_person['email']
    # # trim all white spaces
    email_consultant = email_consultant.strip()

    smtp_server = config('SMTP_SERVER')
    port = config('SMTP_PORT')
    sender_email = config('SENDER_EMAIL')
    password = config('EMAIL_PASSWORD')

    message = f"""\
                <html>
                  <body>
                    <p>Dear {consultant_person['name']},<br>
                    <br /><br />
                    A new investment lead has been generated. Please see the details below:<br />
                    <br /><br />
                    <b>Lead Details</b><br />
                    <br /><br />
                    <b>First Name:</b> {lead['name']}<br />
                    <b>Last Name:</b> {lead['surname']}<br />
                    <b>Cell:</b> {lead['contact']}<br />
                    <b>Email:</b> {lead['email']}<br />
                    <b>Origin:</b> {lead['origin']}<br /> 
                    <b>Interested in:</b> {lead['investment_choice']}<br />
                    
                    <b>Has R100 0000 minimum to invest:</b> {lead['min_value']}<br />
                    <b>Range keen to invest:</b> {lead['investment_amount']}<br />

                    <br /><br />
                    Please contact the lead as soon as possible and follow up by inputting into the app.<br />

                    <br /><br />
                    Kind Regards,<br />
                    OMH App<br />

                  </body>
                </html>
                """

    msg = EmailMessage()
    msg['Subject'] = "Investment Lead"
    msg['From'] = sender_email
    msg['To'] = email_consultant
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


def send_email_to_investment_lead(consultant_person, lead):
    email_lead = lead['email']
    # # trim all white spaces
    email_lead = email_lead.strip()
    #
    smtp_server = config('SMTP_SERVER')
    port = config('SMTP_PORT')
    sender_email = config('SENDER_EMAIL')
    password = config('EMAIL_PASSWORD')

    message = f"""\
                    <html>
                      <body>
                        <p>Good Day {lead['name']},<br>
                        <br /><br />
                        Thank you for your interest in this investment offering.<br />
                        <br /><br />
                        <b>Our investment consultant will be in touch with you shortly.</b><br />
                        <br /><br />
                        Please keep a lookout for their call.<br />
                        <br /><br />
                        Thank You,<br /><br />
                        021 919 9944<br />
                        <a href=mailTo:invest@opportunity.co.za>invest@opportunity.co.za</a><br />
                        <a href=www.opportunity.co.za>www.opportunity.co.za</a><br />

                      </body>
                    </html>
                    """

    # msg = EmailMessage()
    msg = MIMEMultipart()
    msg['Subject'] = "Opportunity Private Capital"
    msg['From'] = sender_email
    msg['To'] = email_lead
    msg.attach(MIMEText(message, 'html'))
    # msg.set_content(message, subtype='html')

    with open("annexures/image002.png", "rb") as image_file:
        image_data = image_file.read()
        image = MIMEImage(image_data, _subtype="png")

        image.add_header("Content-ID", "<image1>")
        msg.attach(image)

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


def create_sales_lead_collection():
    sales_people = [
        {
            "name": "Morne",
            "surname": "Willemse",
            "email": "morne@opportunityprop.co.za",
            "cell": "0837169898",
            "active": True,
            "created_at": datetime.now(),
            "updated_at": datetime.now(),
            "unavailable_from": None,
            "unavailable_to": None,
        },
        {
            "name": "Minette",
            "surname": "du Plessis",
            "email": "minette@opportunityprop.co.za",
            "cell": "0827752178",
            "active": True,
            "created_at": datetime.now(),
            "updated_at": datetime.now(),
            "unavailable_from": None,
            "unavailable_to": None
        },
        {
            "name": "Yvette",
            "surname": "Mostert",
            "email": "yvette@opportunityprop.co.za",
            "cell": "0836161216",
            "active": True,
            "created_at": datetime.now(),
            "updated_at": datetime.now(),
            "unavailable_from": None,
            "unavailable_to": None,
        },
    ]
    db.lead_sales_people.insert_many(sales_people)


# create_sales_lead_collection()


def create_investment_lead_collection():
    sales_people = [
        {
            "name": "Francois",
            "surname": "Geldenhuys",
            "email": "FrancoisG@opportunity.co.za",
            "cell": "0824482928",
            "active": False,
            "created_at": datetime.now(),
            "updated_at": datetime.now(),
            "unavailable_from": None,
            "unavailable_to": None,
        },
        {
            "name": "Leandri",
            "surname": "Kriel",
            "email": "leandri@opportunity.co.za",
            "cell": "0767828558",
            "active": True,
            "created_at": datetime.now(),
            "updated_at": datetime.now(),
            "unavailable_from": None,
            "unavailable_to": None
        },
    ]

    db.lead_investment_consultants.insert_many(sales_people)


# create_investment_lead_collection()

def check_emails_p24():
    smtp_server = config("SMTP_SERVER")
    email_address = config("SENDER_EMAIL")
    password = config("EMAIL_PASSWORD")

    # Connect to the mail server
    mail = imaplib.IMAP4_SSL(f"imap.{smtp_server}")

    # Login to the email account
    mail.login(email_address, password)

    # Select the mailbox (inbox, for example)

    mail.select("inbox")

    # CHANGE TO "UNSEEN" RATHER THAN ALL - Below is for testing ALL
    status, messages = mail.search(None, "UNSEEN")

    # Get the list of email IDs
    email_ids = messages[0].split()

    # Loop through the email IDs
    for email_id in email_ids:
        # Fetch the email by ID
        status, msg_data = mail.fetch(email_id, "(RFC822)")

        # Get the email content
        raw_email = msg_data[0][1]
        msg = email.message_from_bytes(raw_email)

        # Extract relevant information (e.g., subject and sender)
        subject, encoding = decode_header(msg["Subject"])[0]
        if isinstance(subject, bytes):
            subject = subject.decode(encoding or "utf-8")

        sender = msg.get("From")
        if sender:
            sender = sender.split()[-1]
            sender = sender.split("<")[-1].strip(">").strip()

        # if sender == "wayne@opportunity.co.za":
        if sender == "no-reply@property24.com":

            enquiry_by, contact_number, email_address_in_mail, message, address, development, body \
                = "", "", "", "", "", "", ""

            original_date_string = msg["Date"]

            # Convert to datetime object
            original_date = datetime.strptime(original_date_string, "%a, %d %b %Y %H:%M:%S %z")

            # Format the date as "yyyy-mm-dd h:mm:ss"
            formatted_date = original_date.strftime("%Y-%m-%d %H:%M:%S")

            # get the message body
            if msg.is_multipart():
                for part in msg.walk():
                    # extract content type of email
                    content_type = part.get_content_type()
                    content_disposition = str(part.get("Content-Disposition"))
                    try:
                        # get the email body
                        body = part.get_payload(decode=True).decode()
                    except Exception:
                        pass
                    if content_type == "text/plain" and "attachment" not in content_disposition:

                        email_body = body

                        address_match = re.search(r"Address:(.+?)Web ref:", email_body, re.DOTALL)
                        enquiry_by_match = re.search(r"Enquiry by:(.+?)Contact Number:", email_body, re.DOTALL)
                        contact_number_match = re.search(r"Contact Number:(.+?)\(", email_body, re.DOTALL)
                        email_match = re.search(r"Email Address:(.+?)Message:", email_body, re.DOTALL)
                        message_match = re.search(r"Message:(.+?)Regards", email_body, re.DOTALL)

                        # Print extracted information
                        if message_match:
                            message = message_match.group(1).strip()
                            # print(f"Message: {message}")

                        if email_match:
                            email_address_in_mail = email_match.group(1).strip()
                            email_address_in_mail = email_address_in_mail.split("mailto:")[-1].strip()
                            email_address_in_mail = email_address_in_mail.replace(">", "")
                            # make email_address_in_mail lowercase
                            email_address_in_mail = email_address_in_mail.lower()
                            # print(f"Email Address: {email_address_in_mail}")

                        if contact_number_match:
                            contact_number = contact_number_match.group(1).strip()
                            contact_number = contact_number.replace(" ", "")
                            # print(f"Contact Number: {contact_number}")

                        if enquiry_by_match:
                            enquiry_by = enquiry_by_match.group(1).strip()
                            # print(f"Enquiry By: {enquiry_by}")

                        if address_match:
                            address = address_match.group(1).strip()

                            if "Heron View" in address:
                                development = "Heron View"
                            elif "Heron Fields" in address:
                                development = "Heron Fields"
                            elif "Endulini" in address:
                                development = "Endulini"

                        data = {
                            "name": enquiry_by,
                            "surname": "",
                            "contact": contact_number,
                            "email": email_address_in_mail,
                            "message": f"{message}[{address}]",
                            "development": development,
                            "origin": "Property 24",
                            "type": "sales",
                            "submission_date": formatted_date,
                            "contact_time": "ASAP",
                        }

                        print()
                        for item in data:
                            print(item, ":", data[item])
                        print(data)

                        # UNCOMMENT BELOW to UPDATE DB
                        process_property_24_leads(data)
            else:
                # extract content type of email

                content_type = msg.get_content_type()
                print("content_type", content_type)
                # print("msg", msg)
                # get the email body
                body = msg.get_payload(decode=True).decode()
                print("BODY P24", body)
                # if content_type == "text/plain":
                if content_type == "text/html":
                    # print only text email parts
                    # print()
                    print("BODY2 P24", body)
                    # body = body.split("<br>")
                    # # filter out empty strings
                    # body = list(filter(None, body))
                    # # filter out "---" string
                    # body = list(filter(lambda x: x != "---", body))
                    # # filter out where string starts with "Date:
                    # body = list(filter(lambda x: not x.startswith("Date:"), body))
                    # # filter out where string contand \r
                    # body = list(filter(lambda x: not x.startswith("\r"), body))
                    # for item in body:
                    #     # print("item", item)
                    #     if item.startswith("Name:"):
                    #         enquiry_by = item.split("Name:")[-1].strip()
                    #         # print("enquiry_by", enquiry_by)
                    #     elif item.startswith("Mobile:"):
                    #         contact_number = item.split("Mobile:")[-1].strip()
                    #         # print("contact_number", contact_number)
                    #     elif item.startswith("Email:"):
                    #         email_address_in_mail = item.split("Email:")[-1].strip()
                    #         # print("email_address_in_mail", email_address_in_mail)
                    #     elif item.startswith("Message:"):
                    #         message = item.split("Message:")[-1].strip()
                    #         # print("message", message)
                    #     elif item.startswith("Address:"):
                    #         address = item.split("Address:")[-1].strip()
                    #         # print("address", address)
                    #     elif item.startswith("Web ref:"):
                    #         development = item.split("Web ref:")[-1].strip()
                    # print("development", development)
                    # print("BODY3", body)
                    # data = {
                    #
                    #     "name": enquiry_by,
                    #     "surname": "",
                    #     "contact": contact_number,
                    #     "email": email_address_in_mail,
                    #     "message": message,
                    #     "development": "",
                    #     "origin": "opportunityProp",
                    #     "type": "sales",
                    #     "submission_date": formatted_date,
                    #     "contact_time": "ASAP"
                    # }
                    # process_property_24_leads(data)
        # if sender == "waynebruton@icloud.com":
        # if sender == "webmaster@opportunityprop.co.za":
        # if sender is "webmaster@opportunityprop.co.za" and the subject is "Message via website"
        if sender == "webmaster@opportunityprop.co.za" and subject == "Message via website":

            enquiry_by, contact_number, email_address_in_mail, message, address, development, body \
                = "", "", "", "", "", "", ""

            # Print or process the email information
            # print(f"Subject: {subject}")
            # print(f"From: {sender}")
            # print(f"Date: {msg['Date']}")
            # # # print msg recipient
            # print(f"To: {msg['To']}")

            original_date_string = msg["Date"]

            # Convert to datetime object
            original_date = datetime.strptime(original_date_string, "%a, %d %b %Y %H:%M:%S %z")

            # Format the date as "yyyy-mm-dd h:mm:ss"
            formatted_date = original_date.strftime("%Y-%m-%d %H:%M:%S")

            # print(msg)
            if msg.is_multipart():
                for part in msg.walk():
                    # extract content type of email
                    content_type = part.get_content_type()
                    content_disposition = str(part.get("Content-Disposition"))
                    try:
                        # get the email body
                        body = part.get_payload(decode=True).decode()
                    except Exception:
                        pass
                    if content_type == "text/plain" and "attachment" not in content_disposition:
                        # print text/plain emails and skip attachments
                        print("BODY1", body)
                        # print()

                        email_body = body

                        address_match = re.search(r"Address:(.+?)Web ref:", email_body, re.DOTALL)
                        enquiry_by_match = re.search(r"Name:(.+?)Mobile:", email_body,
                                                     re.DOTALL)
                        contact_number_match = re.search(r"Mobile:(.+?)Email:", email_body,
                                                         re.DOTALL)
                        # contact_number_match = re.search(r"Mobile:(.+?)\(", email_body, re.DOTALL)
                        email_match = re.search(r"Email:(.+?)Message:", email_body, re.DOTALL)
                        message_match = re.search(r"Message:(.+?)---", email_body, re.DOTALL)
                        #
                        # # Print extracted information
                        if message_match:
                            message = message_match.group(1).strip()
                            # print(f"Message: {message}")

                        if email_match:
                            email_address_in_mail = email_match.group(1).strip()
                            email_address_in_mail = email_address_in_mail.split("mailto:")[-1].strip()
                            email_address_in_mail = email_address_in_mail.replace(">", "")
                            # make email_address_in_mail lowercase
                            email_address_in_mail = email_address_in_mail.lower()
                            # print(f"Email Address: {email_address_in_mail}")

                        if contact_number_match:
                            contact_number = contact_number_match.group(1).strip()
                            contact_number = contact_number.replace(" ", "")
                            # print(f"Contact Number: {contact_number}")

                        if enquiry_by_match:
                            enquiry_by = enquiry_by_match.group(1).strip()
                            # print(f"Enquiry By: {enquiry_by}")

                        # if address_match:
                        #     address = address_match.group(1).strip()
                        #     # print(f"Address: {address}")
                        #     # if address contains 'Heron View' then development = 'Heron View', if address contains
                        #     # 'Heron Fields' then development = 'Heron Fields', if address contains 'Endulini' then
                        #     # development = 'Endulini'
                        #     if "Heron View" in address:
                        #         development = "Heron View"
                        #     elif "Heron Fields" in address:
                        #         development = "Heron Fields"
                        #     elif "Endulini" in address:
                        #         development = "Endulini"
                        #
                        data = {

                            "name": enquiry_by,
                            "surname": "",
                            "contact": contact_number,
                            "email": email_address_in_mail,
                            "message": message,
                            "development": "",
                            "origin": "opportunityProp",
                            "type": "sales",
                            "submission_date": formatted_date,
                            "contact_time": "ASAP"
                        }

                        # print()
                        # for item in data:
                        #     print(item, ":", data[item])
                        # # print(data)

                        # UNCOMMENT BELOW to UPDATE DB
                        process_property_24_leads(data)

            else:
                # extract content type of email

                content_type = msg.get_content_type()

                body = msg.get_payload(decode=True).decode()
                # print("BODY", body)
                # if content_type == "text/plain":
                if content_type == "text/html":

                    # print only text email parts
                    # print()
                    # print("BODY2", body)
                    body = body.split("<br>")
                    # filter out empty strings
                    body = list(filter(None, body))
                    # filter out "---" string
                    body = list(filter(lambda x: x != "---", body))
                    # filter out where string starts with "Date:
                    body = list(filter(lambda x: not x.startswith("Date:"), body))
                    # filter out where string contand \r
                    body = list(filter(lambda x: not x.startswith("\r"), body))
                    for item in body:
                        # print("item", item)
                        if item.startswith("Name:"):
                            enquiry_by = item.split("Name:")[-1].strip()
                            # print("enquiry_by", enquiry_by)
                        elif item.startswith("Mobile:"):
                            contact_number = item.split("Mobile:")[-1].strip()
                            # print("contact_number", contact_number)
                        elif item.startswith("Email:"):
                            email_address_in_mail = item.split("Email:")[-1].strip()
                            # print("email_address_in_mail", email_address_in_mail)
                        elif item.startswith("Message:"):
                            message = item.split("Message:")[-1].strip()
                            # print("message", message)
                        elif item.startswith("Address:"):
                            address = item.split("Address:")[-1].strip()
                            # print("address", address)
                        elif item.startswith("Web ref:"):
                            development = item.split("Web ref:")[-1].strip()
                            # print("development", development)
                    # print("BODY3", body)
                    data = {

                        "name": enquiry_by,
                        "surname": "",
                        "contact": contact_number,
                        "email": email_address_in_mail,
                        "message": message,
                        "development": "",
                        "origin": "opportunityProp",
                        "type": "sales",
                        "submission_date": formatted_date,
                        "contact_time": "ASAP"
                    }
                    process_property_24_leads(data)
                    # print("data",data)

    # Logout from the email account
    mail.logout()
    print("Done")


def select_sales_person(sales_people, last_leads_generated):
    sales_people_ids = {person["_id"] for person in sales_people}
    last_leads_ids = {lead["sales_person_id"] for lead in last_leads_generated}
    if available_sales_people_ids := sales_people_ids - last_leads_ids:
        return next(person for person in sales_people if person["_id"] in available_sales_people_ids)
    else:
        return next(person for person in sales_people if person["_id"] == last_leads_generated[-1]["sales_person_id"])


def process_property_24_leads(data):
    sales_people = [{**person, "_id": str(person["_id"])} for person in db.lead_sales_people.find({"active": True})]
    if last_leads_generated := list(
            db.leads_sales.find().sort("created_at", -1).limit(len(sales_people))
    ):
        sales_person = select_sales_person(sales_people, last_leads_generated)
    else:
        sales_person = random.choice(sales_people)

    data['created_at'] = datetime.now()
    data["sales_person"] = sales_person["name"] + " " + sales_person["surname"]
    data["sales_person_id"] = sales_person["_id"]

    db.leads_sales.insert_one(data)

    sp_email = send_email_to_sales_person(sales_person, data)
    client_email = send_email_to_sales_lead(sales_person, data)

    return {"message": "success"}

# SET UP CRON JOB FOR BELOW
# check_emails_p24()
scheduler = BackgroundScheduler()
scheduler.add_job(check_emails_p24, 'interval', minutes=5)
scheduler.start()


# Shut down the scheduler when exiting the app

@leads.on_event("shutdown")
def shutdown_event():
    scheduler.shutdown()
