import json
from time import sleep, time

from bson.objectid import ObjectId
from fastapi import APIRouter, Request, BackgroundTasks
from fastapi.responses import FileResponse
from collections import Counter
# from fastapi.encoders import jsonable_encoder
from apscheduler.schedulers.background import BackgroundScheduler, BlockingScheduler

import random
import pandas as pd
import smtplib
from email.message import EmailMessage
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import imaplib
import email
from email.header import decode_header
import os
from decouple import config
from datetime import datetime, timedelta
# from distutils.command.clean import clean
from email.header import decode_header
from configuration.db import db
import re

import vonage

API_KEY_VONAGE = config("VONAGE_API_KEY")
API_SECRET_VONAGE = config("VONAGE_API_SECRET")

leads = APIRouter()


# get AWS_BUCKET_NAME from .env file
# SMTP_SERVER = configuration("SMTP_SERVER")
# print("SMTP_SERVER", SMTP_SERVER)

def fair_allocate_lead(sales_people, last_leads_generated):
    sales_person_allocations = Counter(lead["sales_person_id"] for lead in last_leads_generated[-3:])

    sales_person_chosen = min(sales_people, key=lambda person: sales_person_allocations.get(person["_id"], 0))

    return sales_person_chosen["_id"]


@leads.post("/post_sales_lead_form")
async def post_sales_lead_form(background_tasks: BackgroundTasks, data: Request):
    request = await data.json()
    sales_people = list(db.lead_sales_people.find({"active": True}))
    print("sales_people", sales_people)
    for person in sales_people:
        person["_id"] = str(person["_id"])

    # generate a list of leads limited to the number of sales people

    last_leads_generated = list(db.leads_sales.find().sort("created_at", -1).limit(len(sales_people)))
    print("last_leads_generated", len(last_leads_generated))
    print("last_leads_generated", last_leads_generated)
    print()

    sales_person_chosen = fair_allocate_lead(sales_people, last_leads_generated)
    # print("last_leads_generated", last_leads_generated)
    # for lead_generated in last_leads_generated:
    #     lead_generated["_id"] = str(lead_generated["_id"])
    #     print("lead_generated", lead_generated)
    #     print()

    # for lead in last_leads_generated:
    #     lead["_id"] = str(lead["_id"])
    #     print("lead", lead["sales_person_id"])
    #     print("sales_people", lead['sales_person'])
    #     print()

    # loop through sales_people and see if the sales_person_id is in the last_leads_generated list excluding the last
    # lead

    # Flag to track if a suitable salesperson has been found
    # salesperson_found = False
    #
    # for person in sales_people:
    #     if person["_id"] not in [lead["sales_person_id"] for lead in last_leads_generated[-3:]]:
    #         sales_person_chosen = person["_id"]
    #         salesperson_found = True
    #         break
    #
    # if not salesperson_found:
    #     # If all salespeople are in the last 3 leads, fairly allocate the third last lead
    #     sales_person_chosen = last_leads_generated[-3]["sales_person_id"]

    # for person in sales_people:
    #     # print("person", person["_id"])
    #     # print()
    #     # if the person id is not in the last_leads_generated list
    #     if person["_id"] not in [lead["sales_person_id"] for lead in last_leads_generated[:-1]]:
    #         print("person 1", person["_id"])
    #         print("person 1 name", person["name"])
    #         # print()
    #         sales_person_chosen = person["_id"]
    #         # print("sales_person_chosen_id", sales_person_chosen)
    #         # get the sales_person_chosen from the sales_people list
    #         # sales_person = list(filter(lambda x: x["_id"] == sales_person_chosen, sales_people))[0]
    #         # print("sales_person_chosen", sales_person)
    #         break
    #     else:
    #         sales_person_chosen = last_leads_generated[len(last_leads_generated) - 2]["sales_person_id"]
    #         print("person 2", last_leads_generated[len(last_leads_generated) - 2]["sales_person"])
    #         break
    #         # print("person", person["name"])

    # get the sales_person who is in the last record of the last_leads_generated list

    # get the sales_person_chosen from the sales_people list
    sales_person = list(filter(lambda x: x["_id"] == sales_person_chosen, sales_people))[0]

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
    name = request['name'] + " " + request['surname']
    # send sms as background task
    background_tasks.add_task(send_sms, sales_person['cell'], name)

    return {"message": "success", "sales_person": request["sales_person"]}


def send_sms(number, lead):
    print("number", number)
    client = vonage.Client(key=API_KEY_VONAGE, secret=API_SECRET_VONAGE)
    sms = vonage.Sms(client)
    to_number = re.sub(r"[ \-()]", "", number)
    # replace only leading 0 with +27
    to_number = re.sub(r"^0", "+27", to_number)

    responseData = sms.send_message(
        {
            "from": "OMH-APP",
            # "from": "Vonage APIs",
            "to": to_number,
            "text": f"New lead assigned to you, Log into https://www.opportunitymanagement.co.za to view details. "
                    f"Lead Name: {lead}",
        }
    )

    if responseData["messages"][0]["status"] == "0":
        return {"message": "SMS sent successfully"}

    else:
        return {"message": "SMS not sent XXX"}


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

    name = request['name'] + " " + request['surname']
    # send sms as background task
    background_tasks.add_task(send_sms, consultant_person['cell'], name)

    return {"message": "success", "consultant": request["consultant"]}


@leads.post("/opportunity_contact_form")
async def opportunity_contact_form(background_tasks: BackgroundTasks, data: Request):
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

        # send email to consultant as background task
        background_tasks.add_task(send_email_to_consultant, consultant_person, data_investments)
        # send email to client as background task
        background_tasks.add_task(send_email_to_investment_lead, consultant_person, data_investments)

        name = data_investments['name'] + " " + data_investments['surname']
        background_tasks.add_task(send_sms, consultant_person['cell'], name)

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

@leads.post('/get_sales_leads')
async def get_sales_leads(request: Request):
    start = time()
    request = await request.json()

    print("request", request)
    if request["user"] != None:
        leads = list(db.leads_sales.find({"sales_person_id": request["sales_person_id"]}))
    else:
        create_sales_lead_excel_sheet()
        leads = list(db.leads_sales.find())
    # opportunities =list(db.opportunities.find()) get all opportunities and project only 'opportunity_code' and
    # 'Category' but exclude where Category = "NGAH" or "Endulini" or "Southwark" or "Goodwood"

    opportunities = list(db.opportunities.find({"Category": {"$nin": ["NGAH", "Endulini", "Goodwood", "Southwark"]}},
                                               {"opportunity_code": 1, "Category": 1, "_id": 0}))

    end = time()
    # leads = list(db.leads_sales.find())
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
    print(len(leads))
    for index, lead in enumerate(leads):
        lead["_id"] = str(lead["_id"])
        lead["created_at"] = lead["created_at"].strftime("%Y-%m-%d %H:%M:%S")
        lead['rental_enquiry'] = lead.get('rental_enquiry', False)

        for key, default in default_values.items():
            lead[key] = lead.get(key, default)

    leads = sorted(leads, key=lambda x: x['created_at'], reverse=True)

    print("Time taken to get sales leads:", end - start)

    return {"message": "success", "leads": leads, "opportunities": opportunities}


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

    check_unanswered_leads()

    return {"message": "success", "leads": leads}


@leads.post('/edit_sales_lead')
async def edit_sales_lead(background_task: BackgroundTasks, data: Request):
    request = await data.json()

    try:
        # print("request",request)

        lead_id = request['_id']
        # find the lead by id
        old_lead = db.leads_sales.find_one({"_id": ObjectId(lead_id)})
        # print("old_lead", old_lead)
        if (old_lead.get('rental_enquiry', False) != request['rental_enquiry']) and old_lead.get('rental_enquiry',
                                                                                                 False) == False and \
                request['rental_enquiry'] == True:
            # print("send email")
            # print()
            # print(request)
            # send_email_to_mario(request)
            # send as background task
            background_task.add_task(send_email_to_mario, request)
            # send email to sales person as background task
            # background_tasks.add_task(send_email_to_sales_person, old_lead, request)
            # # send email to client as background task
            # background_tasks.add_task(send_email_to_sales_lead, old_lead, request)

        del request['_id']
        request['created_at'] = datetime.strptime(request['created_at'], "%Y-%m-%d %H:%M:%S")
        test = db.leads_sales.update_one({"_id": ObjectId(lead_id)}, {"$set": request})
        print("test", test)

        return {"message": "success"}

    except Exception as e:
        print("Error:", e)
        return {"message": "Lead not updated"}


@leads.post('/add_sales_lead')
async def add_sales_lead(data: Request):
    try:
        request = await data.json()

        created_at = datetime.now()
        print("created_at", created_at)
        request['created_at'] = created_at
        print("request", request)
        #
        db.leads_sales.insert_one(request)
        return {"message": "success"}
    except Exception as e:
        print("Error:", e)
        return {"message": "Lead not added"}


@leads.post('/delete_sales_lead')
async def delete_sales_lead(data: Request):
    request = await data.json()
    try:
        print(request)
        lead_id = request['id']
        # remove the document from the collection
        db.leads_sales.delete_one({"_id": ObjectId(lead_id)})

        return {"message": "success"}
    except Exception as e:
        print("Error:", e)
        return {"message": "Lead not deleted"}


@leads.post('/add_investment_lead')
async def add_investment_lead(data: Request):
    try:
        request = await data.json()
        # print("request", request)
        created_at = datetime.now()
        print("created_at", created_at)
        request['created_at'] = created_at
        print("request", request)
        #
        db.leads_investments.insert_one(request)
        return {"message": "success"}
    except Exception as e:
        print("Error:", e)
        return {"message": "Lead not added"}


@leads.post('/delete_investment_lead')
async def delete_investment_lead(data: Request):
    request = await data.json()
    try:
        print(request)
        lead_id = request['id']
        # remove the document from the collection
        db.leads_investments.delete_one({"_id": ObjectId(lead_id)})

        return {"message": "success"}
    except Exception as e:
        print("Error:", e)
        return {"message": "Lead not deleted"}


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
    print("email_sp", email_sp)
    smtp_server = config('SMTP_SERVER')
    port = config('SMTP_PORT')
    sender_email = config('SENDER_EMAIL')
    password = config('EMAIL_PASSWORD')

    plain_text = f"Dear {sales_person['name']},\n\nA new sales lead has been generated. Please see the details below:\n\nLead Details\n\nFirst Name: {lead['name']}\nLast Name: {lead['surname']}\nCell: {lead['contact']}\nEmail: {lead['email']}\nOrigin: {lead['origin']}\nDevelopment: {lead['development']}\nBest Time: {lead['contact_time']}\nMessage: {lead['message']}\n\nPlease contact the lead as soon as possible and follow up by inputting into the app.\n\nKind Regards,\nOMH App"

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
    msg.set_content(plain_text)
    msg.add_alternative(message, subtype='html')

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

    plain_text = f"Dear Leandri,\n\nA new investment lead has been categorised as {lead['category']}.\n\nLead Details\n\nFirst Name: {lead['name']}\nLast Name: {lead['surname']}\nCell: {lead['contact']}\nEmail: {lead['email']}\nMessage: {lead['message']}\n\nKind Regards,\nOMH App"

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
    msg.set_content(plain_text)
    msg.add_alternative(message, subtype='html')

    try:
        with smtplib.SMTP_SSL(smtp_server, port) as server:
            server.ehlo()
            server.login(sender_email, password=password)
            server.send_message(msg)
            return {"message": "Email sent successfully"}
    except Exception as e:
        print("Error:", e)
        return {"message": "Email not sent"}


def send_email_to_mario(lead):
    email_sp = 'mario@opportunityprop.co.za'
    # email_sp = 'wayne@opportunity.co.za'
    smtp_server = config('SMTP_SERVER')
    port = config('SMTP_PORT')
    sender_email = config('SENDER_EMAIL')
    password = config('EMAIL_PASSWORD')

    plain_text = f"Dear Mario,\n\nA new investment lead has been categorised as Rental.\n\nLead Details\n\nFirst Name: {lead['name']}\nLast Name: {lead['surname']}\nCell: {lead['contact']}\nEmail: {lead['email']}\nMessage: {lead['message']}\n\nKind Regards,\nOMH App"

    message = f"""
                <html>
                  <body>
                    <p>Dear Mario,<br>
                    <br>
                    A new investment lead has been categorised as Rental:<br>
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
    msg.set_content(plain_text)
    msg.add_alternative(message, subtype='html')

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

    plain_text = f"Good Day {lead['name']},\n\nThank you for your enquiry.\n\nThe following agent will be in contact with you shortly:\n\nAgent Details\n\nFirst Name: {sales_person['name']}\nLast Name: {sales_person['surname']}\nCell: {sales_person['cell']}\nEmail: {sales_person['email']}\n\nPlease keep a lookout for their call.\n\nKind Regards,\nOpportunity Property"

    message = f"""\
                    <html>
                      <body>
                        <p>Good Day {lead['name']},<br>
                        <br /><br />
                        Thank you for your enquiry.<br />
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
    msg.set_content(plain_text)

    # msg.set_content(message, subtype='html')

    with open("annexures/Signature OppProp.jpg", "rb") as image_file:
        image_data = image_file.read()
        image = MIMEImage(image_data, _subtype="jpg")

        image.add_header("Content-ID", "<image1>")
        msg.attach(image)

    msg.add_alternative(message, 'html')

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

    plain_text = f"Dear {consultant_person['name']},\n\nA new investment lead has been generated. Please see the details below:\n\nLead Details\n\nFirst Name: {lead['name']}\nLast Name: {lead['surname']}\nCell: {lead['contact']}\nEmail: {lead['email']}\nOrigin: {lead['origin']}\nInterested in: {lead['investment_choice']}\nHas R100 0000 minimum to invest: {lead['min_value']}\nRange keen to invest: {lead['investment_amount']}\n\nPlease contact the lead as soon as possible and follow up by inputting into the app.\n\nKind Regards,\nOMH App"

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
    msg.set_content(plain_text)
    msg.add_alternative(message, subtype='html')

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


def send_email_to_consultant_unanswered(consultant_person, lead):
    email_consultant = consultant_person['email']
    # # trim all white spaces
    email_consultant = email_consultant.strip()

    smtp_server = config('SMTP_SERVER')
    port = config('SMTP_PORT')
    sender_email = config('SENDER_EMAIL')
    password = config('EMAIL_PASSWORD')

    plain_text = f"Dear {consultant_person['name']},\n\nThis potential client did not answer, perhaps you want to follow up? It has been around 48 hours. Please see the details below:\n\nLead Details\n\nFirst Name: {lead['name']}\nLast Name: {lead['surname']}\nCell: {lead['contact']}\nEmail: {lead['email']}\nOrigin: {lead['origin']}\nInterested in: {lead['investment_choice']}\nHas R100 0000 minimum to invest: {lead['min_value']}\nRange keen to invest: {lead['investment_amount']}\n\nPlease contact the lead as soon as possible and follow up by inputting into the app.\n\nKind Regards,\nOMH App"

    message = f"""\
                <html>
                  <body>
                    <p>Dear {consultant_person['name']},<br>
                    <br /><br />
                    This potential client did not answer, perhaps you want to follow up? It has been around 48 hours. Please see the details below:<br />
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
    msg.set_content(plain_text)
    msg.add_alternative(message, subtype='html')

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

    # msg = EmailMessage()
    msg = MIMEMultipart()
    msg['Subject'] = "Opportunity Private Capital"
    msg['From'] = sender_email
    msg['To'] = email_lead

    # msg.set_content(message, subtype='html')

    with open("annexures/image002.png", "rb") as image_file:
        image_data = image_file.read()
        image = MIMEImage(image_data, _subtype="png")

        image.add_header("Content-ID", "<image1>")
        msg.attach(image)

    plain_text = f"Good Day {lead['name']},\n\nThank you for your interest in this investment offering.\n\nThe following consultant will be in contact with you shortly:\n\nConsultant Details\n\nFirst Name: {consultant_person['name']}\nLast Name: {consultant_person['surname']}\nCell: {consultant_person['cell']}\nEmail: {consultant_person['email']}\n\nPlease keep a lookout for their call.\n\nKind Regards,\nOpportunity Private Capital"

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
                            <img src="cid:image1" alt="Investment Image"><br />

                          </body>
                        </html>
                        """

    msg.set_content(plain_text)
    msg.add_alternative(message, 'html')

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

    mail = imaplib.IMAP4_SSL(f"imap.{smtp_server}")

    mail.login(email_address, password)

    mail.select("inbox")

    # CHANGE TO "UNSEEN" RATHER THAN ALL - Below is for testing ALL
    status, messages = mail.search(None, "UNSEEN")

    # Get the list of email IDs
    email_ids = messages[0].split()

    final_data = []
    done = False

    # Loop through the email IDs
    for email_id in email_ids:

        # Fetch the email by ID
        status, msg_data = mail.fetch(email_id, "(RFC822)")

        # Get the email content
        raw_email = msg_data[0][1]
        # msg = email.message_from_bytes(raw_email)
        # msg = email.message_from_string(raw_email.decode())
        if isinstance(raw_email, int):
            raw_email = str(raw_email).encode()

        msg = email.message_from_bytes(raw_email)
        # print("msg", msg)
        subject = msg["Subject"]
        if subject is not None:
            subject, encoding = decode_header(subject)[0]
            if isinstance(subject, bytes):
                subject = subject.decode(encoding or "utf-8")

        # Extract relevant information (e.g., subject and sender)
        # subject, encoding = decode_header(msg["Subject"])[0]
        if isinstance(subject, bytes):
            subject = subject.decode(encoding or "utf-8")

        sender = msg.get("From")
        if sender:
            sender = sender.split()[-1]
            sender = sender.split("<")[-1].strip(">").strip()

        # if sender == "wayne@opportunity.co.za" and subject contains "Contact Request"
        if sender == "no-reply@property24.com" and "Contact Request" in subject:

            enquiry_by, contact_number, email_address_in_mail, message, address, development, body \
                = "", "", "", "", "", "", ""

            original_date_string = msg["Date"]

            # Convert to datetime object
            original_date = datetime.strptime(original_date_string, "%a, %d %b %Y %H:%M:%S %z")

            # Format the date as "yyyy-mm-dd h:mm:ss"
            formatted_date = original_date.strftime("%Y-%m-%d %H:%M:%S")

            content_type = msg.get_content_type()

            body = msg.get_payload(decode=True).decode()

            if content_type == "text/html":

                html = body

                # Define a regular expression pattern
                enquiry_by = r'<strong>Enquiry by:</strong>\s*<\/td>\s*<td[^>]*>\s*(.*?)\s*<\/td>'
                address = r"<strong>Address:</strong>\s*<\/td>\s*<td[^>]*>\s*(.*?)\s*<\/td>"
                contact_number = r'<strong>Contact Number:</strong>\s*<\/td>\s*<td[^>]*>\s*(.*?)\s*<\/td>'
                email_address = r'<strong>Email Address:</strong>\s*<\/td>\s*<td[^>]*>\s*(.*?)\s*<\/td>'
                message = r"<strong>Message:</strong>\s*<\/td>\s*<td[^>]*>\s*(.*?)\s*<\/td>"

                # Search for the pattern in the HTML
                enquiry_by_match = re.search(enquiry_by, html, re.DOTALL)
                address_match = re.search(address, html, re.DOTALL)
                contact_number_match = re.search(contact_number, html, re.DOTALL)
                email_address_match = re.search(email_address, html, re.DOTALL)
                message_match = re.search(message, html, re.DOTALL)

                if enquiry_by_match:
                    enquiry_by = enquiry_by_match.group(1).strip()
                    # print(f'Enquiry By: {enquiry_by}')
                else:
                    print('Enquiry By information not found in the HTML.')

                if address_match:
                    raw_address = address_match.group(1).strip()
                    # decoded_address = html.unescape(raw_address)
                    address = raw_address.replace("&#xD;", "")
                    address = address.replace("&#xA;", "")
                    address = address.replace("&#x27;", "'")
                    # print(f'Address: {address}')
                else:
                    print('Address information not found in the HTML.')

                if contact_number_match:
                    contact_number = contact_number_match.group(0).strip()
                    # split contact_number by <td and get the last item in the list
                    contact_number = contact_number.split("<td")[-1].strip()
                    # split contact number by > and get the last item in the list
                    contact_number = contact_number.split(">")[1].strip()
                    # split contact number by ( and get the first item in the list
                    contact_number = contact_number.split("(")[0].strip()
                    # print(f'Contact Number: {contact_number}')
                else:
                    print('Contact Number information not found in the HTML.')

                if email_address_match:
                    email_address = email_address_match.group(0).strip()
                    # split email_address by <td and get the last item in the list
                    email_address = email_address.split("<td")[-1].strip()
                    # # split email_address by > and get the last item in the list
                    email_address = email_address.split(">")[1].strip()
                    # # split email_address by < and get the first item in the list
                    email_address = email_address.split("<")[0].strip()
                    # print(f'Email Address: {email_address}')
                else:
                    print('Email Address information not found in the HTML.')

                if message_match:
                    message = message_match.group(0).strip()
                    # split message by <td and get the last item in the list
                    message = message.split("<td")[-1].strip()
                    # split message by > and get the last item in the list
                    message = message.split(">")[1].strip()
                    # split message by < and get the first item in the list
                    message = message.split("<")[0].strip()
                    message = message.replace("&#x27;", "'")
                    message = message.replace("&#xD;", "")
                    message = message.replace("&#xA;", "")


                else:
                    print('Message information not found in the HTML.')

                data = {
                    "email_id": email_id,
                    "name": enquiry_by,
                    "surname": "",
                    "contact": contact_number,
                    "email": email_address,
                    "message": f"{message} [{address}]",
                    "development": "",
                    "origin": "Property24",
                    "type": "sales",
                    "submission_date": formatted_date,
                    "contact_time": "ASAP"
                }

                final_data.append(data)

        if sender == "webmaster@opportunityprop.co.za" and subject == "Message via website":

            enquiry_by, contact_number, email_address_in_mail, message, address, development, body \
                = "", "", "", "", "", "", ""

            original_date_string = msg["Date"]

            # Convert to datetime object
            original_date = datetime.strptime(original_date_string, "%a, %d %b %Y %H:%M:%S %z")

            # Format the date as "yyyy-mm-dd h:mm:ss"
            formatted_date = original_date.strftime("%Y-%m-%d %H:%M:%S")

            content_type = msg.get_content_type()

            body = msg.get_payload(decode=True).decode()

            if content_type == "text/html":

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

                data = {
                    "email_id": email_id,
                    "name": enquiry_by,
                    "surname": "",
                    "contact": contact_number,
                    "email": email_address_in_mail,
                    "message": message,
                    "development": "",
                    "origin": "OpportunityProp",
                    "type": "sales",
                    "submission_date": formatted_date,
                    "contact_time": "ASAP"
                }

                # if data['message'] contains <br> then do nothing, otherwise append to final_data
                if "<br>" in data['message']:
                    pass
                else:
                    final_data.append(data)

        if sender == "noreply@opportunityprop.co.za" and subject == "Inquiry from https://opportunityprop.co.za/":

            enquiry_by, contact_number, email_address_in_mail, message, address, development, body \
                = "", "", "", "", "", "", ""

            original_date_string = msg["Date"]

            # Convert to datetime object
            original_date = datetime.strptime(original_date_string, "%a, %d %b %Y %H:%M:%S %z")

            # Format the date as "yyyy-mm-dd h:mm:ss"
            formatted_date = original_date.strftime("%Y-%m-%d %H:%M:%S")

            content_type = msg.get_content_type()

            body = msg.get_payload(decode=True).decode()

            if content_type == "text/html":

                body = body.split("<br>")
                # filter out empty strings
                body = list(filter(None, body))

                # filter out "---" string
                body = list(filter(lambda x: x != "---", body))
                # # filter out where string starts with "Date:
                body = list(filter(lambda x: not x.startswith("Date:"), body))
                # # filter out where string contand \r
                body = list(filter(lambda x: not x.startswith("\r"), body))

                # remove all html tags
                body = [re.sub('<[^<]+?>', '', item) for item in body]
                # print("body", body)
                for index, item in enumerate(body):
                    item = item.replace("<strong>", "")
                    item = item.replace("</strong>", "")
                    # print("item", item, index)
                    # if item contains "Name:" then split by "Name:" and get the last item in the list
                    if "Name:" in item:
                        enquiry_by = item.split("Name:")[-1].strip()
                        # print("enquiry_by", enquiry_by)
                    elif item.startswith("Phone:"):
                        contact_number = item.split("Phone:")[-1].strip()
                        # print("contact_number", contact_number)
                    elif item.startswith("Email:"):
                        email_address_in_mail = item.split("Email:")[-1].strip()
                        # print("email_address_in_mail", email_address_in_mail)
                    elif item.startswith("Message:") and index + 1 < len(body):
                        message = body[index + 1].strip()
                        # print("message", message)
                    elif 'Sent From:' in item:
                        address = item.split("Sent From:")[-1].strip()
                        # address = item.split(\r)[0].strip()
                        # right split address on last '/'
                        address = address.rsplit('/', 1)[0].strip()
                        # print("sent from:", address)

                data = {
                    "email_id": email_id,
                    "name": enquiry_by,
                    "surname": "",
                    "contact": contact_number,
                    "email": email_address_in_mail,
                    "message": message + " [" + address + "]",
                    "development": "",
                    "origin": "OpportunityProp",
                    "type": "sales",
                    "submission_date": formatted_date,
                    "contact_time": "ASAP"
                }

                # print("data", data)
                #
                final_data.append(data)

        # I AM HERE

        # if sender == "leads@syte.co.za" and subject == "Inquiry from https://opportunityprop.co.za/":
        if sender == "leads@syte.co.za":

            enquiry_by, contact_number, email_address_in_mail, message, address, development, body \
                = "", "", "", "", "", "", ""

            original_date_string = msg["Date"]

            # Convert to datetime object
            original_date = datetime.strptime(original_date_string, "%a, %d %b %Y %H:%M:%S %z")

            # Format the date as "yyyy-mm-dd h:mm:ss"
            formatted_date = original_date.strftime("%Y-%m-%d %H:%M:%S")

            content_type = msg.get_content_type()

            # print("content_type", content_type)

            body = msg.get_payload(decode=True).decode()

            if content_type == "text/plain":

                # print("body", body)
                message = ""
                for line in body.split("\n"):

                    # print("line", line)
                    if "name:" in line:
                        enquiry_by = line.split("name:")[-1].strip()
                        # print("enquiry_by", enquiry_by)
                    elif "Name:" in line:
                        enquiry_by = line.split("Name:")[-1].strip()
                        # print("enquiry_by", enquiry_by)
                    elif "Phone Number:" in line:
                        contact_number = line.split("Phone Number:")[-1].strip()
                        # print("contact_number", contact_number)
                    elif "Email:" in line:
                        email_address_in_mail = line.split("Email:")[-1].strip()
                        # print("email_address_in_mail", email_address_in_mail)
                    elif "Campaign name" in line:
                        development = line.split("[")[-1].strip()
                        development = development.split("]")[0].strip()
                        # print("development",development)
                        # message += line + "\n"
                    else:
                        message += line + "\n"

                data = {
                    "email_id": email_id,
                    "name": enquiry_by,
                    "surname": "",
                    "contact": contact_number,
                    "email": email_address_in_mail,
                    "message": message,
                    "development": development,
                    "origin": "Meta(FB-IG)",
                    "type": "sales",
                    "submission_date": formatted_date,
                    "contact_time": "ASAP"
                }
                #
                # print("data", data)
                # #
                final_data.append(data)
    # Logout from the email account
    mail.logout()

    if final_data:
        while not done:
            process_property_24_leads(final_data)
            done = True
            final_data = []
            print("Done")
        # exit mail

    # print("final_data", final_data)
    # print("final_data", len(final_data))


def select_sales_person(sales_people, last_leads_generated):
    sales_people_ids = {person["_id"] for person in sales_people}
    last_leads_ids = {lead["sales_person_id"] for lead in last_leads_generated}
    if available_sales_people_ids := sales_people_ids - last_leads_ids:
        return next(person for person in sales_people if person["_id"] in available_sales_people_ids)
    else:
        return next(person for person in sales_people if person["_id"] == last_leads_generated[-1]["sales_person_id"])


def process_property_24_leads(data):
    # in data list, eliminate duplicates based on 'email_id'
    data = [dict(t) for t in {tuple(d.items()) for d in data}]
    # print("data", data)
    # print("len(data)", len(data))
    # print()

    for index, email_data in enumerate(data):

        name = email_data["name"]
        submission_date = email_data["submission_date"]
        email = email_data["email"]
        origin = email_data["origin"]
        sleep(3)
        result = db.leads_sales.find_one(
            {"name": name, "submission_date": submission_date, "email": email, "origin": origin})

        if result is not None:
            print(f"Lead already exists: {index}")
            print("result", result)
            continue

        # print("email_id", email_data["email_id"])
        print(f"Processing lead: {index}")
        del email_data["email_id"]

        sales_people = [{**person, "_id": str(person["_id"])} for person in db.lead_sales_people.find({"active": True})]
        sales_person = (
            select_sales_person(sales_people, last_leads_generated)
            if (
                last_leads_generated := list(
                    db.leads_sales.find()
                    .sort("created_at", -1)
                    .limit(len(sales_people))
                )
            )
            else random.choice(sales_people)
        )
        # print("data", data)

        email_data['created_at'] = datetime.now()
        email_data["sales_person"] = sales_person["name"] + " " + sales_person["surname"]
        email_data["sales_person_id"] = sales_person["_id"]

        db.leads_sales.insert_one(email_data)

        send_email_to_sales_person(sales_person, email_data)
        # client_email = send_email_to_sales_lead(sales_person, email_data)

        name = email_data['name'] + " " + email_data['surname']
        send_sms(sales_person['cell'], name)

        # send_sms(email_data)

    return {"message": "success"}


def check_unanswered_leads():
    leads = list(db.leads_investments.find({"action_taken": "Called - No Answer"}))
    for lead in leads:
        action_taken_date_time = datetime.strptime(lead.get('action_taken_date_time', ""), "%Y-%m-%d %H:%M:%S")
        now = datetime.now()
        if action_taken_date_time != "":
            difference = now - action_taken_date_time
        else:
            # difference = now less 48 hours
            difference = now - timedelta(hours=48)
        # if difference > timedelta(minutes=5):
        if difference > timedelta(hours=47):
            consultant = db.lead_investment_consultants.find_one({"_id": ObjectId(lead.get('consultant_id', ""))})
            send_email_to_consultant_unanswered(consultant, lead)


@leads.post('/check_emails_omh_app')
async def check_emails_omh_app():
    try:
        check_emails_p24()
        # check_unanswered_leads()
        return {"message": "success Checked"}
    except Exception as e:
        print("Error:", e)
        return {"message": "Emails not checked"}


def insert_lead_sales():
    # get all the documents from the lead_sales collection
    lead_sales = list(db.leads_sales.find())
    for lead in lead_sales:
        lead['_id'] = str(lead['_id'])
        lead['rental_enquiry'] = lead.get('rental_enquiry', False)
        id = lead['_id']
        del lead['_id']
        # update the document in the lead_sales collection with the new data
        db.leads_sales.update_one({'_id': ObjectId(id)}, {'$set': lead})
    print("lead_sales", lead_sales)


# insert_lead_sales()


def create_sales_lead_excel_sheet():
    lead_sales = list(db.leads_sales.find({}, {"_id": 0}))
    # order by created_at in descending order
    lead_sales = sorted(lead_sales, key=lambda x: x['created_at'], reverse=True)
    # print("lead_sales", lead_sales[0])
    df = pd.DataFrame(lead_sales)
    df.to_excel('sales_leads.xlsx', index=False)


# create_sales_lead_excel_sheet()

@leads.get("/get_sales_lead_spreadsheet")
async def get_sales_lead_spreadsheet(file_name):
    try:
        # print("file_name", file_name)
        file_name = file_name + ".xlsx"
        try:
            return FileResponse(f"{file_name}", filename=file_name)
        except FileNotFoundError:
            return {"ERROR": "File does not exist!!"}
    except Exception as e:
        print(e)
        return {"ERROR": "Please Try again"}
