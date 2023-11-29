from bson.objectid import ObjectId
from fastapi import APIRouter, Request
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
# import os
from datetime import datetime
# from distutils.command.clean import clean
from email.header import decode_header
import re

leads = APIRouter()


@leads.post("/post_sales_lead_form")
async def post_sales_lead_form(data: Request):
    request = await data.json()
    # get lead_sales_people collection from db where active is True
    sales_people = list(db.lead_sales_people.find({"active": True}))
    length = len(sales_people)
    sales_people_to_not_choose = []
    for person in sales_people:
        person["_id"] = str(person["_id"])

    last_leads_generated = list(db.leads_sales.find().sort("created_at", -1).limit(length))
    if len(last_leads_generated) == 0:
        sales_person = random.choice(sales_people)
    elif length > len(last_leads_generated) > 0:
        for lead in last_leads_generated:
            sales_people_to_not_choose.append(lead["sales_person_id"])
        sales_people_to_choose = list(filter(lambda x: x["_id"] not in sales_people_to_not_choose, sales_people))
        sales_person = random.choice(sales_people_to_choose)
    else:
        sales_people_to_choose = []
        for person in sales_people:
            # create variable called filtered_leads where the sales_person_id is equal to the person id
            filtered_leads = list(filter(lambda x: x["sales_person_id"] == person["_id"], last_leads_generated))
            if len(filtered_leads) == 0:
                sales_people_to_choose.append(person["_id"])
        if len(sales_people_to_choose) == 0:
            sales_people_to_choose.append(last_leads_generated[len(last_leads_generated) - 1]["sales_person_id"])
        sales_person = list(filter(lambda x: x["_id"] in sales_people_to_choose, sales_people))[0]

    request['origin'] = request['source']['source']
    request['type'] = request['source']['type']
    del request['source']
    request['created_at'] = datetime.now()
    request["sales_person"] = sales_person["name"] + " " + sales_person["surname"]
    request["sales_person_id"] = sales_person["_id"]

    db.leads_sales.insert_one(request)

    sp_email = send_email_to_sales_person(sales_person, request)

    client_email = send_email_to_sales_lead(sales_person, request)

    return {"message": "success",
            "sales_person": request["sales_person"]
            }


@leads.post("/post_investments_lead_form")
async def post_investments_lead_form(data: Request):
    request = await data.json()

    # get lead_sales_people collection from db where active is True
    consultants = list(db.lead_investment_consultants.find({"active": True}))
    length = len(consultants)
    consultants_to_not_choose = []
    for person in consultants:
        person["_id"] = str(person["_id"])

    last_leads_generated = list(db.leads_investments.find().sort("created_at", -1).limit(length))
    if len(last_leads_generated) == 0:
        consultant_person = random.choice(consultants)
    elif length > len(last_leads_generated) > 0:
        for lead in last_leads_generated:
            consultants_to_not_choose.append(lead["sales_person_id"])
        consultants_people_to_choose = list(filter(lambda x: x["_id"] not in consultants_to_not_choose, consultants))
        consultant_person = random.choice(consultants_people_to_choose)
    else:
        consultants_people_to_choose = []
        for person in consultants:
            # create variable called filtered_leads where the sales_person_id is equal to the person id
            filtered_leads = list(filter(lambda x: x["consultant_id"] == person["_id"], last_leads_generated))
            if len(filtered_leads) == 0:
                consultants_people_to_choose.append(person["_id"])
        if len(consultants_people_to_choose) == 0:
            consultants_people_to_choose.append(last_leads_generated[len(last_leads_generated) - 1]["consultant_id"])
        consultant_person = list(filter(lambda x: x["_id"] in consultants_people_to_choose, consultants))[0]

    request['origin'] = request['source']['source']
    request['type'] = request['source']['type']
    del request['source']
    request['created_at'] = datetime.now()
    request["consultant"] = consultant_person["name"] + " " + consultant_person["surname"]
    request["consultant_id"] = consultant_person["_id"]

    db.leads_investments.insert_one(request)

    consultant_email = send_email_to_consultant(consultant_person, request)

    client_email = send_email_to_investment_lead(consultant_person, request)

    return {"message": "success",
            "consultant": request["consultant"]
            }


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


@leads.post("/opportunityprop_contact_form")
async def opportunityprop_contact_form(data: Request):
    request = await data.form()
    print("request", request)

    sales_people = list(db.lead_sales_people.find({"active": True}))
    length = len(sales_people)
    sales_people_to_not_choose = []
    for person in sales_people:
        person["_id"] = str(person["_id"])

    last_leads_generated = list(db.leads_sales.find().sort("created_at", -1).limit(length))
    if len(last_leads_generated) == 0:
        sales_person = random.choice(sales_people)
    elif length > len(last_leads_generated) > 0:
        for lead in last_leads_generated:
            sales_people_to_not_choose.append(lead["sales_person_id"])
        sales_people_to_choose = list(filter(lambda x: x["_id"] not in sales_people_to_not_choose, sales_people))
        sales_person = random.choice(sales_people_to_choose)
    else:
        sales_people_to_choose = []
        for person in sales_people:
            # create variable called filtered_leads where the sales_person_id is equal to the person id
            filtered_leads = list(filter(lambda x: x["sales_person_id"] == person["_id"], last_leads_generated))
            if len(filtered_leads) == 0:
                sales_people_to_choose.append(person["_id"])
        if len(sales_people_to_choose) == 0:
            sales_people_to_choose.append(last_leads_generated[len(last_leads_generated) - 1]["sales_person_id"])
        sales_person = list(filter(lambda x: x["_id"] in sales_people_to_choose, sales_people))[0]

    name = request.getlist("form_fields[name]")[0]
    surname = ""
    contact = request.getlist("form_fields[field_c045127]")[0]
    email = request.getlist("form_fields[email]")[0]
    development = ""
    submission_date = datetime.now()
    # format submission_date as yyyy-mm-dd hh:mm:ss
    submission_date = submission_date.strftime("%Y-%m-%d %H:%M:%S")
    contact_time = "ASAP"
    message = request.getlist("form_fields[field_16ad61e]")[0]
    origin = 'OpportunityProp Website'
    type = 'sales'
    created_at = datetime.now()

    data = {
        "name": name,
        "surname": surname,
        "email": email,
        "contact": contact,
        "development": development,
        "submission_date": submission_date,
        "contact_time": contact_time,
        "message": message,
        "origin": origin,
        "type": type,
        "created_at": created_at,
        "sales_person": sales_person["name"] + " " + sales_person["surname"],
        "sales_person_id": sales_person["_id"]
    }

    # request = data

    db.leads_sales.insert_one(data)

    sp_email = send_email_to_sales_person(sales_person, data)

    client_email = send_email_to_sales_lead(sales_person, data)


    print("opportunityprop_contact_form ZZ")
    return {"message": "success"}


@leads.get('/get_sales_leads')
async def get_sales_leads():
    leads = list(db.leads_sales.find())
    for lead in leads:
        lead["_id"] = str(lead["_id"])
        lead["created_at"] = lead["created_at"].strftime("%Y-%m-%d %H:%M:%S")
        lead['action_taken'] = lead.get('action_taken', "")
        lead['action_taken_date_time'] = lead.get('action_taken_date_time', "")
        lead['comments'] = lead.get('comments', "")
        lead['previous_actions'] = lead.get('previous_actions', [])
        lead['rating'] = lead.get('rating', 0)
        lead['purchased'] = lead.get('purchased', "")
        lead['purchased_date'] = lead.get('purchased_date', "")
        lead['purchased_unit'] = lead.get('purchased_unit', "")

    # sort leads by created_at descending
    leads = sorted(leads, key=lambda x: x['created_at'], reverse=True)

    return {"message": "success", "leads": leads}


@leads.get('/get_sales_people')
async def get_sales_people():
    sales_people = list(db.lead_sales_people.find())
    for person in sales_people:
        person["_id"] = str(person["_id"])
        person["created_at"] = person["created_at"].strftime("%Y-%m-%d %H:%M:%S")
        person["updated_at"] = person["updated_at"].strftime("%Y-%m-%d %H:%M:%S")
        # person["unavailable_from"] = person["unavailable_from"].strftime("%Y-%m-%d %H:%M:%S")
        # person["unavailable_to"] = person["unavailable_to"].strftime("%Y-%m-%d %H:%M:%S")

    return {"message": "success", "sales_people": sales_people}


@leads.post('/edit_sales_lead')
async def edit_sales_lead(data: Request):
    request = await data.json()

    try:
        lead_id = request['_id']
        del request['_id']
        # convert created_at to datetime object
        request['created_at'] = datetime.strptime(request['created_at'], "%Y-%m-%d %H:%M:%S")

        # obj_instance = ObjectId(lead_id)
        # update leads_sales in db
        db.leads_sales.update_one({"_id": ObjectId(lead_id)}, {"$set": request})

        return {"message": "success"}

    except Exception as e:
        print("Error:", e)
        return {"message": "Lead not updated"}


def send_email_to_sales_person(sales_person, lead):
    email_sp = sales_person['email']
    # # trim all white spaces
    email_sp = email_sp.strip()
    #
    smtp_server = "depro8.fcomet.com"
    port = 465  # For starttls
    sender_email = 'omh-app@opportunitymanagement.co.za'
    password = "12071994Wb!"

    message = f"""\
                <html>
                  <body>
                    <p>Dear {sales_person['name']},<br>
                    <br /><br />
                    A new sales lead has been generated. Please see the details below:<br />
                    <br /><br />
                    <b>Lead Details</b><br />
                    <br /><br />
                    <b>First Name:</b> {lead['name']}<br />
                    <b>Last Name:</b> {lead['surname']}<br />
                    <b>Cell:</b> {lead['contact']}<br />
                    <b>Email:</b> {lead['email']}<br />
                    <b>Origin:</b> {lead['origin']}<br /> 
                    <b>Development:</b> {lead['development']}<br />
                    <b>Best Time:</b> {lead['contact_time']}<br />
                    <b>Message:</b> {lead['message']}<br />
                    
                    <br /><br />
                    Please contact the lead as soon as possible and follow up by inputting into the app.<br />
                    
                    <br /><br />
                    Kind Regards,<br />
                    OMH App<br />
                         
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


def send_email_to_sales_lead(sales_person, lead):
    email_lead = lead['email']
    # # trim all white spaces
    email_lead = email_lead.strip()
    #
    smtp_server = "depro8.fcomet.com"
    port = 465  # For starttls
    sender_email = 'omh-app@opportunitymanagement.co.za'
    password = "12071994Wb!"

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

    smtp_server = "depro8.fcomet.com"
    port = 465  # For starttls
    sender_email = 'omh-app@opportunitymanagement.co.za'
    password = "12071994Wb!"

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
    smtp_server = "depro8.fcomet.com"
    port = 465  # For starttls
    sender_email = 'omh-app@opportunitymanagement.co.za'
    password = "12071994Wb!"

    message = f"""\
                    <html>
                      <body>
                        <p>Good Day {lead['name']},<br>
                        <br /><br />
                        Thank you for your interest in this investment offering.<br />
                        <br /><br />
                        <b>Our investment consultant will be in touch with you shortly.</b><br />
                        <br /><br />
                        <b>Consultant Details</b><br />
                        {consultant_person['name']} {consultant_person['surname']}.<br />
                        
                         {consultant_person['cell']}<br />
                         <a href=mailTo:{consultant_person['email']}>{consultant_person['email']}</a><br />


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
    # global enquiry_by, contact_number, email_address_in_mail, message, address, development, body
    smtp_server = "depro8.fcomet.com"
    email_address = "omh-app@opportunitymanagement.co.za"
    password = "12071994Wb!"

    # Connect to the mail server
    mail = imaplib.IMAP4_SSL("imap.depro8.fcomet.com")

    # Login to the email account
    mail.login(email_address, password)

    # Select the mailbox (inbox, for example)
    mail.select("inbox")

    # Search for all unseen (new) emails
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

        # if sender == "no-reply@property24.com" and msg['To'] != "Mario Stoop <mario@opportunityprop.co.za>": NOT
        # SURE IF I NEED TO WORRY ABOUT MARIO
        # if sender == "wayne@opportunity.co.za":
        if sender == "no-reply@property24.com":

            enquiry_by, contact_number, email_address_in_mail, message, address, development, body \
                = "", "", "", "", "", "", ""

            # Print or process the email information
            print(f"Subject: {subject}")
            print(f"From: {sender}")
            print(f"Date: {msg['Date']}")
            # print msg recipient
            print(f"To: {msg['To']}")

            original_date_string = msg["Date"]

            # Convert to datetime object
            original_date = datetime.strptime(original_date_string, "%a, %d %b %Y %H:%M:%S %z")

            # Format the date as "yyyy-mm-dd h:mm:ss"
            formatted_date = original_date.strftime("%Y-%m-%d %H:%M:%S")

            # Print the result
            print("formatted_date", formatted_date)
            # convert msg['Date'] into a datetime object and format it to be in the format yyyy-mm-dd hh:mm:ss called
            # submission_date

            print()
            # get the message body
            if msg.is_multipart():
                for part in msg.walk():
                    # extract content type of email
                    content_type = part.get_content_type()
                    content_disposition = str(part.get("Content-Disposition"))
                    try:
                        # get the email body
                        body = part.get_payload(decode=True).decode()
                    except:
                        pass
                    if content_type == "text/plain" and "attachment" not in content_disposition:
                        # print text/plain emails and skip attachments
                        # print("BODY1",body)
                        # print()

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
                            # print(f"Address: {address}")
                            # if address contains 'Heron View' then development = 'Heron View', if address contains
                            # 'Heron Fields' then development = 'Heron Fields', if address contains 'Endulini' then
                            # development = 'Endulini'
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
                            "message": message + "[" + address + "]",
                            "development": development,
                            "origin": "Property 24",
                            "type": "sales",
                            "submission_date": formatted_date,
                            "contact_time": "ASAP"
                        }

                        print()
                        for item in data:
                            print(item, ":", data[item])
                        # print(data)

                        # UNCOMMENT BELOW to UPDATE DB
                        process_property_24_leads(data)

                    # elif "attachment" in content_disposition:
                    #     # download attachment
                    #     filename = part.get_filename()
                    #     if filename:
                    #         folder_name = clean(subject)
                    #         if not os.path.isdir(folder_name):
                    #             # make a folder for this email (named after the subject)
                    #             os.mkdir(folder_name)
                    #         filepath = os.path.join(folder_name, filename)
                    #         # download attachment and save it
                    #         open(filepath, "wb").write(part.get_payload(decode=True))
            # else:
            #     # extract content type of email
            #     content_type = msg.get_content_type()
            #     # get the email body
            #     body = msg.get_payload(decode=True).decode()
            #     if content_type == "text/plain":
            #         # print only text email parts
            #         # print()
            #         print("BODY2", body)

    # Logout from the email account
    mail.logout()


def process_property_24_leads(data):
    sales_people = list(db.lead_sales_people.find({"active": True}))
    length = len(sales_people)
    sales_people_to_not_choose = []
    for person in sales_people:
        person["_id"] = str(person["_id"])

    last_leads_generated = list(db.leads_sales.find().sort("created_at", -1).limit(length))
    if len(last_leads_generated) == 0:
        sales_person = random.choice(sales_people)
    elif length > len(last_leads_generated) > 0:
        for lead in last_leads_generated:
            sales_people_to_not_choose.append(lead["sales_person_id"])
        sales_people_to_choose = list(filter(lambda x: x["_id"] not in sales_people_to_not_choose, sales_people))
        sales_person = random.choice(sales_people_to_choose)
    else:
        sales_people_to_choose = []
        for person in sales_people:
            # create variable called filtered_leads where the sales_person_id is equal to the person id
            filtered_leads = list(filter(lambda x: x["sales_person_id"] == person["_id"], last_leads_generated))
            if len(filtered_leads) == 0:
                sales_people_to_choose.append(person["_id"])
        if len(sales_people_to_choose) == 0:
            sales_people_to_choose.append(last_leads_generated[len(last_leads_generated) - 1]["sales_person_id"])
        sales_person = list(filter(lambda x: x["_id"] in sales_people_to_choose, sales_people))[0]

    # data['origin'] = request['source']['source']
    # request['type'] = request['source']['type']
    # del request['source']
    data['created_at'] = datetime.now()
    data["sales_person"] = sales_person["name"] + " " + sales_person["surname"]
    data["sales_person_id"] = sales_person["_id"]
    # TEMPORARY
    # data['email'] = 'waynebruton@icloud.com'

    db.leads_sales.insert_one(data)

    sp_email = send_email_to_sales_person(sales_person, data)
    # print()
    # print("sp_email", sp_email)
    client_email = send_email_to_sales_lead(sales_person, data)
    # print()
    # print("client_email", client_email)

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
