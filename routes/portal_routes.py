import copy
import os
import secrets

from fastapi import APIRouter, Request
from fastapi.responses import FileResponse
from portal_statement_files.portal_statement_create import create_pdf

import time

import smtplib
from email.message import EmailMessage

portal_info = APIRouter()


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
    file_name = investor_statement_name
    dir_path = "portal_statements"
    dir_list = os.listdir(dir_path)
    if file_name in dir_list:
        return FileResponse(f"{dir_path}/{file_name}", filename=file_name)
    else:
        return {"ERROR": "File does not exist!!"}


@portal_info.post("/generate_otp")
async def generate_otp(data: Request):
    request = await data.json()
    pin = secrets.randbelow(10 ** 6)
    print(pin)
    print(request['email'])
    email = request['email'].split(":")[1]
    # trim all white spaces
    email = email.strip()
    print("email",email)



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


@portal_info.post("/send_email")
async def send_email(data: Request):
    request = await data.json()
    smtp_server = "depro8.fcomet.com"
    port = 465  # For starttls
    sender_email = 'omh-app@opportunitymanagement.co.za'
    password = "12071994Wb!"

    message = f"""\
    <html>
      <body>
        <p>Hi Wayne,<br>
           <strong>How are you?</strong><br>

           Below is all the code required to send an email with attachments using Python:<br><br>
           I did hard code the file name, but you can easily change that to a variable.<br><br>
           {request['message']}      
              <br><br>
              reply to this waynebruton@icloud.com.
              <br><br>
              I should be twins.
        </p>
      </body>
    </html>
    """

    msg = EmailMessage()
    msg['Subject'] = request['subject']
    msg['From'] = sender_email
    msg['To'] = request['to_email']
    msg.set_content(message, subtype='html')

    with open('excel_files/Sales ForecastHeron.xlsx', 'rb') as f:
        file_content = f.read()
    file_name = os.path.basename('Sales ForecastHeron.xlsx')
    msg.add_attachment(file_content, maintype='application', subtype='octet-stream', filename=file_name)

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
