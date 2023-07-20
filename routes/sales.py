import os
import boto3
from botocore.exceptions import ClientError
from fastapi import APIRouter, Request
from fastapi.responses import FileResponse
from config.db import db
from bson.objectid import ObjectId
from decouple import config
# from typing import Annotated

from sales_python_files.sales_excel_sheet import create_excel_file
from sales_client_onboarding_creation_file.onboarding import print_onboarding_pdf
from sales_client_onboarding_creation_file.offer_to_purchase import print_otp_pdf

sales = APIRouter()

# MONGO COLLECTIONS
investors = db.investors
opportunityCategories = db.opportunityCategories
opportunities = db.opportunities
sales_parameters = db.salesParameters
sales_processed = db.sales_processed
sales_agents = db.sales_agents
mortgage_brokers = db.sales_mortgage_brokers

# AWS BUCKET INFO - ENSURE IN VARIABLES ON HEROKU
AWS_BUCKET_NAME = config("AWS_BUCKET_NAME")
AWS_BUCKET_REGION = config("AWS_BUCKET_REGION")
AWS_ACCESS_KEY_ID = config("AWS_ACCESS_KEY_ID")
AWS_SECRET_ACCESS_KEY = config("AWS_SECRET_ACCESS_KEY")

# CONNECT TO S3 SERVICE
s3 = boto3.client(
    "s3",
    aws_access_key_id=AWS_ACCESS_KEY_ID,
    aws_secret_access_key=AWS_SECRET_ACCESS_KEY
)


def get_files_required_for_sales():
    file_list_items = []
    results = list(sales_processed.find(({})))

    try:
        for result in results:
            if 'opportunity_otp' in result:
                file_list_items.append(result['opportunity_otp'])
            if 'opportunity_uploadId' in result:
                file_list_items.append(result['opportunity_uploadId'])
            if 'opportunity_addressproof' in result:
                file_list_items.append(result['opportunity_addressproof'])
            if 'opportunity_uploadId_sec' in result:
                file_list_items.append(result['opportunity_uploadId_sec'])
            if 'opportunity_addressproof_sec' in result:
                file_list_items.append(result['opportunity_addressproof_sec'])
            if 'opportunity_upload_deposite' in result:
                file_list_items.append(result['opportunity_upload_deposite'])
            if 'opportunity_uploadId_3rd' in result:
                file_list_items.append(result['opportunity_uploadId_3rd'])
            if 'opportunity_uploadId_4th' in result:
                file_list_items.append(result['opportunity_uploadId_4th'])
            if 'opportunity_uploadId_5th' in result:
                file_list_items.append(result['opportunity_uploadId_5th'])
            if 'opportunity_uploadId_6th' in result:
                file_list_items.append(result['opportunity_uploadId_6th'])
            if 'opportunity_uploadId_7th' in result:
                file_list_items.append(result['opportunity_uploadId_7th'])
            if 'opportunity_uploadId_8th' in result:
                file_list_items.append(result['opportunity_uploadId_8th'])
            if 'opportunity_uploadId_9th' in result:
                file_list_items.append(result['opportunity_uploadId_9th'])
            if 'opportunity_uploadId_10th' in result:
                file_list_items.append(result['opportunity_uploadId_10th'])
            if 'opportunity_addressproof_3rd' in result:
                file_list_items.append(result['opportunity_addressproof_3rd'])
            if 'opportunity_addressproof_4th' in result:
                file_list_items.append(result['opportunity_addressproof_4th'])
            if 'opportunity_addressproof_5th' in result:
                file_list_items.append(result['opportunity_addressproof_5th'])
            if 'opportunity_addressproof_6th' in result:
                file_list_items.append(result['opportunity_addressproof_6th'])
            if 'opportunity_addressproof_7th' in result:
                file_list_items.append(result['opportunity_addressproof_7th'])
            if 'opportunity_addressproof_8th' in result:
                file_list_items.append(result['opportunity_addressproof_8th'])
            if 'opportunity_addressproof_9th' in result:
                file_list_items.append(result['opportunity_addressproof_9th'])
            if 'opportunity_addressproof_10th' in result:
                file_list_items.append(result['opportunity_addressproof_10th'])
            if 'opportunity_upload_annexure' in result:
                file_list_items.append(result['opportunity_upload_annexure'])
            if 'opportunity_upload_company_docs' in result:
                file_list_items.append(result['opportunity_upload_company_docs'])
            if 'opportunity_upload_company_addressproof' in result:
                file_list_items.append(result['opportunity_upload_company_addressproof'])

        file_list_items = [x for x in file_list_items if x is not None]
        file_list_items = [x for x in file_list_items if x != '']
        file_list_items = [x.split("/")[1] for x in file_list_items if type(x) == str]

        list_of_filenames = os.listdir("sales_documents")

        final_list_to_download = [x for x in file_list_items if x not in list_of_filenames]

        if final_list_to_download:
            for file in final_list_to_download:
                try:
                    s3.download_file(AWS_BUCKET_NAME, file, f"./sales_documents/{file}")
                except ClientError as err:
                    print("Credentials are incorrect")
                    print("XX", err)
                except Exception as err:
                    print("YY", err)
        else:
            print("List is empty")
    except Exception as err:
        print("ZZ", err)


# GET DEVELOPMENTS
@sales.post("/getDevelopmentForSales")
async def get_developments_for_sales():
    result_developments = list(opportunityCategories.aggregate([{"$project": {
        "id2": {'$toString': "$_id"}, "_id": 0, "Description": 1, "blocked": 1
    }}]))
    # print("Hello")
    # filter out if Descrption = 'Southwark' or blocked = 1
    result_developments = [x for x in result_developments if x['Description'] != 'Southwark' and x['blocked'] != 1]

    # for development in result_developments:
    #     print(development)
    get_files_required_for_sales()
    return result_developments


@sales.post("/get_new_sale_info")
async def get_new_sale_info(data: Request):
    request = await data.json()
    # print(request)
    result = list(opportunities.aggregate([
        {'$match': {'opportunity_code': request['opportunity_code']}}
    ])
    )
    for item in result:
        item['_id'] = str(item['_id'])
    # print(result)
    return result


# GET UNITS FOR SALE
@sales.post("/getUnitsForSales")
async def get_units_for_sales(data: Request):
    request = await data.json()
    # print(request)
    result = list(opportunities.aggregate([

        {
            '$match': {
                'Category': {"$in": request['development']},
                'opportunity_sold': False
            }
        },

        {
            '$project': {
                '_id': 0,
                'opportunity_code': 1,
                'opportunity_sale_price': 1,
                'opportunity_name': 1,
                'opportunity_sold': 1,

            }
        }
    ]))
    # print(len(result))
    # in result, format the opportunity_sale_price as currency with two decimal places and an R in front and a
    # thousand separator
    for unit in result:
        unit['opportunity_sale_price'] = float(unit['opportunity_sale_price'])
        unit['opportunity_sale_price'] = f"R{unit['opportunity_sale_price']:,.2f}"

    # print(result)
    # print(len(result))
    return result


@sales.get("/sales_agents")
async def get_sales_agents():
    result = list(sales_agents.aggregate([
        {
            '$project': {
                '_id': 0,
                'agent_name': 1,
            }
        }
    ]))
    return result


# GET SOLD UNITS
@sales.post("/get_units_sold")
async def get_all_units_sold(data: Request):
    request = await data.json()
    # print(request)
    response = list(sales_processed.aggregate([
        {
            '$match': {
                'development': {"$in": request['development']}
            }
        }, {
            '$project': {
                '_id': 0,
                'id': 1,
                'development': 1,
                'opportunity_code': 1,
                'opportunity_firstname': 1,
                'opportunity_lastname': 1,
                'opportunity_otp': 1,
                'opportunity_deposite_date': 1,
                'opportunity_bond_instruction_date': 1,
                'opportunity_actual_lodgement': 1,
                'opportunity_actual_reg_date': 1,

                'opportunity_bond_originator': 1,
                'opportunity_pay_type': 1,
                'opportunity_client_type': 1,
                'opportunity_client_no': 1,

                'opportunity_uploadId': 1,
                'opportunity_uploadId_sec': 1,
                'opportunity_uploadId_3rd': 1,
                'opportunity_uploadId_4th': 1,
                'opportunity_uploadId_5th': 1,
                'opportunity_uploadId_6th': 1,
                'opportunity_uploadId_7th': 1,
                'opportunity_uploadId_8th': 1,
                'opportunity_uploadId_9th': 1,
                'opportunity_uploadId_10th': 1,
                'opportunity_addressproof': 1,
                'opportunity_addressproof_sec': 1,
                'opportunity_addressproof_3rd': 1,
                'opportunity_addressproof_4th': 1,
                'opportunity_addressproof_5th': 1,
                'opportunity_addressproof_6th': 1,
                'opportunity_addressproof_7th': 1,
                'opportunity_addressproof_8th': 1,
                'opportunity_addressproof_9th': 1,
                'opportunity_addressproof_10th': 1,
                'opportunity_upload_deposite': 1,
                'opportunity_upload_annexure': 1,
                'opportunity_upload_company_docs': 1,
                'opportunity_upload_company_addressproof': 1,
            }
        }
    ]))
    # print(response)
    for sale in response:

        if 'opportunity_otp' not in sale:
            sale['opportunity_otp'] = None
        if 'opportunity_deposite_date' not in sale:
            sale['opportunity_deposite_date'] = None
        if 'opportunity_bond_instruction_date' not in sale:
            sale['opportunity_bond_instruction_date'] = None
        if 'opportunity_actual_lodgement' not in sale:
            sale['opportunity_actual_lodgement'] = None
        if 'opportunity_actual_reg_date' not in sale:
            sale['opportunity_actual_reg_date'] = None
        if 'opportunity_uploadId' not in sale:
            sale['opportunity_uploadId'] = None
        if 'opportunity_uploadId_sec' not in sale:
            sale['opportunity_uploadId_sec'] = None
        if 'opportunity_uploadId_3rd' not in sale:
            sale['opportunity_uploadId_3rd'] = None
        if 'opportunity_uploadId_4th' not in sale:
            sale['opportunity_uploadId_4th'] = None
        if 'opportunity_uploadId_5th' not in sale:
            sale['opportunity_uploadId_5th'] = None
        if 'opportunity_uploadId_6th' not in sale:
            sale['opportunity_uploadId_6th'] = None
        if 'opportunity_uploadId_7th' not in sale:
            sale['opportunity_uploadId_7th'] = None
        if 'opportunity_uploadId_8th' not in sale:
            sale['opportunity_uploadId_8th'] = None
        if 'opportunity_uploadId_9th' not in sale:
            sale['opportunity_uploadId_9th'] = None
        if 'opportunity_uploadId_10th' not in sale:
            sale['opportunity_uploadId_10th'] = None
        if 'opportunity_addressproof' not in sale:
            sale['opportunity_addressproof'] = None
        if 'opportunity_addressproof_sec' not in sale:
            sale['opportunity_addressproof_sec'] = None
        if 'opportunity_addressproof_3rd' not in sale:
            sale['opportunity_addressproof_3rd'] = None
        if 'opportunity_addressproof_4th' not in sale:
            sale['opportunity_addressproof_4th'] = None
        if 'opportunity_addressproof_5th' not in sale:
            sale['opportunity_addressproof_5th'] = None
        if 'opportunity_addressproof_6th' not in sale:
            sale['opportunity_addressproof_6th'] = None
        if 'opportunity_addressproof_7th' not in sale:
            sale['opportunity_addressproof_7th'] = None
        if 'opportunity_addressproof_8th' not in sale:
            sale['opportunity_addressproof_8th'] = None
        if 'opportunity_addressproof_9th' not in sale:
            sale['opportunity_addressproof_9th'] = None
        if 'opportunity_addressproof_10th' not in sale:
            sale['opportunity_addressproof_10th'] = None
        if 'opportunity_upload_deposite' not in sale:
            sale['opportunity_upload_deposite'] = None
        if 'opportunity_upload_annexure' not in sale:
            sale['opportunity_upload_annexure'] = None
        if 'opportunity_upload_company_docs' not in sale:
            sale['opportunity_upload_company_docs'] = None
        if 'opportunity_upload_company_addressproof' not in sale:
            sale['opportunity_upload_company_addressproof'] = None
        if 'opportunity_pay_type' not in sale:
            sale['opportunity_pay_type'] = None

    for sale in response:
        if sale['opportunity_client_type'] == 'Company':
            total_docs_required = 9
        elif int(sale['opportunity_client_no'].split(" ")[0]) == 1:
            total_docs_required = 5
        elif int(sale['opportunity_client_no'].split(" ")[0]) == 2:
            total_docs_required = 7
        elif int(sale['opportunity_client_no'].split(" ")[0]) == 3:
            total_docs_required = 9
        elif int(sale['opportunity_client_no'].split(" ")[0]) == 4:
            total_docs_required = 11
        elif int(sale['opportunity_client_no'].split(" ")[0]) == 5:
            total_docs_required = 13
        elif int(sale['opportunity_client_no'].split(" ")[0]) == 6:
            total_docs_required = 15
        elif int(sale['opportunity_client_no'].split(" ")[0]) == 7:
            total_docs_required = 17
        elif int(sale['opportunity_client_no'].split(" ")[0]) == 8:
            total_docs_required = 19
        elif int(sale['opportunity_client_no'].split(" ")[0]) == 9:
            total_docs_required = 21
        elif int(sale['opportunity_client_no'].split(" ")[0]) == 10:
            total_docs_required = 23

        if sale['opportunity_bond_originator'] == 'Cash' or sale['opportunity_pay_type'] == 'Cash':
            total_docs_required -= 1

        total_docs_uploaded = 0
        if sale['opportunity_otp'] is not None and sale['opportunity_otp'] != "":
            total_docs_uploaded += 1
        if sale['opportunity_uploadId'] is not None and sale['opportunity_uploadId'] != "":
            total_docs_uploaded += 1
        if sale['opportunity_uploadId_sec'] is not None and sale['opportunity_uploadId_sec'] != "":
            total_docs_uploaded += 1
        if sale['opportunity_uploadId_3rd'] is not None and sale['opportunity_uploadId_3rd'] != "":
            total_docs_uploaded += 1
        if sale['opportunity_uploadId_4th'] is not None and sale['opportunity_uploadId_4th'] != "":
            total_docs_uploaded += 1
        if sale['opportunity_uploadId_5th'] is not None and sale['opportunity_uploadId_5th'] != "":
            total_docs_uploaded += 1
        if sale['opportunity_uploadId_6th'] is not None and sale['opportunity_uploadId_6th'] != "":
            total_docs_uploaded += 1
        if sale['opportunity_uploadId_7th'] is not None and sale['opportunity_uploadId_7th'] != "":
            total_docs_uploaded += 1
        if sale['opportunity_uploadId_8th'] is not None and sale['opportunity_uploadId_8th'] != "":
            total_docs_uploaded += 1
        if sale['opportunity_uploadId_9th'] is not None and sale['opportunity_uploadId_9th'] != "":
            total_docs_uploaded += 1
        if sale['opportunity_uploadId_10th'] is not None and sale['opportunity_uploadId_10th'] != "":
            total_docs_uploaded += 1
        if sale['opportunity_addressproof'] is not None and sale['opportunity_addressproof'] != "":
            total_docs_uploaded += 1
        if sale['opportunity_addressproof_sec'] is not None and sale['opportunity_addressproof_sec'] != "":
            total_docs_uploaded += 1
        if sale['opportunity_addressproof_3rd'] is not None and sale['opportunity_addressproof_3rd'] != "":
            total_docs_uploaded += 1
        if sale['opportunity_addressproof_4th'] is not None and sale['opportunity_addressproof_4th'] != "":
            total_docs_uploaded += 1
        if sale['opportunity_addressproof_5th'] is not None and sale['opportunity_addressproof_5th'] != "":
            total_docs_uploaded += 1
        if sale['opportunity_addressproof_6th'] is not None and sale['opportunity_addressproof_6th'] != "":
            total_docs_uploaded += 1
        if sale['opportunity_addressproof_7th'] is not None and sale['opportunity_addressproof_7th'] != "":
            total_docs_uploaded += 1
        if sale['opportunity_addressproof_8th'] is not None and sale['opportunity_addressproof_8th'] != "":
            total_docs_uploaded += 1
        if sale['opportunity_addressproof_9th'] is not None and sale['opportunity_addressproof_9th'] != "":
            total_docs_uploaded += 1
        if sale['opportunity_addressproof_10th'] is not None and sale['opportunity_addressproof_10th'] != "":
            total_docs_uploaded += 1
        if sale['opportunity_upload_deposite'] is not None and sale['opportunity_upload_deposite'] != "":
            total_docs_uploaded += 1
        if sale['opportunity_upload_annexure'] is not None and sale['opportunity_upload_annexure'] != "":
            total_docs_uploaded += 1
        if sale['opportunity_upload_company_docs'] is not None and sale['opportunity_upload_company_docs'] != "":
            total_docs_uploaded += 1
        if sale['opportunity_upload_company_addressproof'] is not None and sale[
            'opportunity_upload_company_addressproof'] != "":
            total_docs_uploaded += 1

        sale['total_docs_required'] = total_docs_required
        sale['total_docs_uploaded'] = total_docs_uploaded

        if sale['opportunity_otp'] is None or sale['opportunity_otp'] == "":
            sale['reserved'] = True
        else:
            sale['reserved'] = False

        if (sale['opportunity_deposite_date'] is None or sale['opportunity_deposite_date'] == "") and \
                sale['reserved'] is False:
            sale['pending'] = True
            sale['reserved'] = False
        else:
            sale['pending'] = False
        if sale['opportunity_deposite_date'] is not None and sale['opportunity_deposite_date'] != "" and sale[
            'pending']:
            sale['sold'] = True
        else:
            sale['sold'] = False
        if sale['opportunity_bond_instruction_date'] is not None or sale['opportunity_bond_originator'] == 'Cash':
            sale['bond_approval'] = True
        else:
            sale['bond_approval'] = False
        if sale['opportunity_actual_lodgement'] is not None and sale['opportunity_actual_lodgement'] != "":
            sale['lodged'] = True
        else:
            sale['lodged'] = False
        if sale['opportunity_actual_reg_date'] is not None and sale['opportunity_actual_reg_date'] != "":
            sale['registered'] = True
        else:
            sale['registered'] = False

        del sale['opportunity_otp']
        del sale['opportunity_deposite_date']
        del sale['opportunity_bond_instruction_date']
        del sale['opportunity_actual_lodgement']
        del sale['opportunity_actual_reg_date']
        del sale['opportunity_upload_company_addressproof']
        del sale['opportunity_upload_company_docs']
        del sale['opportunity_upload_annexure']
        del sale['opportunity_upload_deposite']
        del sale['opportunity_addressproof_10th']
        del sale['opportunity_addressproof_9th']
        del sale['opportunity_addressproof_8th']
        del sale['opportunity_addressproof_7th']
        del sale['opportunity_addressproof_6th']
        del sale['opportunity_addressproof_5th']
        del sale['opportunity_addressproof_4th']
        del sale['opportunity_addressproof_3rd']
        del sale['opportunity_addressproof_sec']
        del sale['opportunity_addressproof']
        del sale['opportunity_uploadId_10th']
        del sale['opportunity_uploadId_9th']
        del sale['opportunity_uploadId_8th']
        del sale['opportunity_uploadId_7th']
        del sale['opportunity_uploadId_6th']
        del sale['opportunity_uploadId_5th']
        del sale['opportunity_uploadId_4th']
        del sale['opportunity_uploadId_3rd']
        del sale['opportunity_uploadId_sec']
        del sale['opportunity_uploadId']

        if sale['registered']:
            sale['lodged'] = True
        if sale['lodged']:
            sale['bond_approval'] = True
        if sale['bond_approval']:
            sale['sold'] = True
        # if sale['sold']:
        #     sale['pending'] = True
        # if sale['pending']:
        #     sale['reserved'] = True

    # print(response[0])

    return response


@sales.post("/deleteSale")
async def delete_sale(data: Request):
    request = await data.json()
    try:
        print(request)
        result = sales_processed.delete_one({"opportunity_code": request['opportunity_code']})
        result2 = opportunities.update_one({"opportunity_code": request['opportunity_code']},
                                           {"$set": {"opportunity_sold": False, "opportunity_final_transfer_date": ""}})
        # print(result.deleted_count)

        return {"message": "success"}
    except Exception as e:
        print(e)
        return {"message": "error"}


@sales.post("/getLinkedInvestors")
async def get_investors_linked_to_sale(data: Request):
    request = await data.json()
    result = list(investors.aggregate([
        {
            '$unwind': {
                'path': '$investments',
                'preserveNullAndEmptyArrays': False
            }
        }, {
            '$match': {
                'investments.opportunity_code': request['opportunity_code']
            }
        }, {
            '$set': {
                'opportunity_code': '$investments.opportunity_code',
                'investment_amount': '$investments.investment_amount'
            }
        }, {
            '$project': {
                'investor_acc_number': 1,
                'investor_name': 1,
                'investor_surname': 1,
                'opportunity_code': 1,
                'investment_amount': 1,
                'id': {
                    '$toString': '$_id'
                },
                '_id': 0
            }
        }
    ]))

    return result


# SAVE A SALE
@sales.post("/saveSale")
async def save_sale(data: Request):
    try:
        request = await data.json()
        # print(request)

        # Update the opportunities collection
        opportunity_code = request['formData']['opportunity_code']

        if (request['formData']['opportunity_pay_type'] == "Cash" or request['formData']['opportunity_bank'] == "Cash"
                or request['formData']['opportunity_bond_originator'] == "Cash"
                or (request['formData']['opportunity_bond_instruction_date'] != ""
                    and request['formData']['opportunity_bond_instruction_date'] is not None)):
            opportunity_sold = True
            result = opportunities.update_one({"opportunity_code": opportunity_code}, {"$set": {
                "opportunity_sold": opportunity_sold}})
        else:
            opportunity_sold = False
            result = opportunities.update_one({"opportunity_code": opportunity_code}, {"$set": {
                "opportunity_sold": opportunity_sold}})

        if request['formData']['opportunity_actual_reg_date'] != "" and request['formData'][
            'opportunity_actual_reg_date'] is not None:
            opportunity_sold = True
            result = opportunities.update_one({"opportunity_code": opportunity_code}, {"$set": {
                "opportunity_sold": opportunity_sold,
                "opportunity_final_transfer_date": request['formData']['opportunity_actual_reg_date']}})
        else:
            result = opportunities.update_one({"opportunity_code": opportunity_code}, {"$set": {
                "opportunity_final_transfer_date": ""}})

        if request['formData']['id'] != "":
            id_to_update = request['formData']['id']
            obj_instance = ObjectId(id_to_update)
            sales_processed.update_one({"_id": obj_instance}, {"$set": request['formData']})
            return {"updated": "Success"}

        else:
            obj_instance = sales_processed.insert_one(request['formData']).inserted_id
            sales_processed.update_one({"_id": obj_instance}, {"$set": {"id": str(obj_instance)}})
            return {"id": str(obj_instance)}
    except Exception as e:
        print(e)
        return {"error": str(e)}


# DELETE A SALE AND ASSOCIATED UPLOADED DOCUMENTS
# @sales.post("/delete_sale")
# async def delete_sale(data: Request):
#     request = await data.json()
#     obj_instance = ObjectId(request['id'])
#     result = sales_processed.delete_one({"_id": obj_instance})
#     # print(result)
#     unit = request['unit']
#     list_of_filenames = os.listdir("sales_documents")
#     # DO AWS REMOVE FILE
#     for file in list_of_filenames:
#         length = len(unit)
#         if file[:length] == unit:
#             os.remove(f"sales_documents/{file}")
#     return "Deleted"


# RETRIEVE AN ALREADY SOLD UNIT
@sales.post("/get_sold_unit")
async def get_sold_unit(data: Request):
    request = await data.json()
    result = sales_processed.find_one(
        {"opportunity_code": request['opportunity_code'], "development": request['development']})
    # print(result)
    if result:
        result["id"] = str(result["_id"])
        del result["_id"]
    return result


# UPLOAD A DOCUMENT
@sales.post("/upload_sales_file")
async def upload_file(data: Request):
    form = await data.form()
    print("form", form['fileName'])
    filename = form['fileName']
    contents = await form['doc'].read()
    with open(f"sales_documents/{filename}", 'wb') as f:
        f.write(contents)
    try:
        s3.upload_file(
            f"sales_documents/{filename}",
            AWS_BUCKET_NAME,
            f"{filename}",
        )
    except ClientError as err:
        print("Credentials are incorrect")
        print(err)
    except Exception as err:
        print(err)
    return {"fileName": f"sales_documents/{filename}"}


# DELETE AN UPLOADED DOCUMENT
@sales.post("/delete_uploaded_file")
async def delete_uploaded_file(data: Request):
    try:
        request = await data.json()
        # print(request)
        is_exists = os.path.exists(request['fileToDelete'])
        # print(is_exists)
        if is_exists:
            os.remove(request['fileToDelete'])
            return {"file_deleted": True}
        else:
            return {"file_deleted": False}
    except Exception as err:
        print(err)
        return {"file_deleted": False}


# RETRIEVE AN UPLOADED DOCUMENT
@sales.get("/get_uploaded_document")
async def get_uploaded_file(file_name):
    try:  # File Name incl path.
        is_exists = os.path.exists(file_name)
        if is_exists:
            return FileResponse(f"{file_name}", media_type="application/pdf")
        else:
            return {"ERROR": "File does not exist!!"}
    except Exception as err:
        print(err)
        return {"ERROR": "File does not exist!!"}


# GET ALL SALES and when returning ensure _id is a string and project all fields, then call create_excel_file
@sales.get("/get_all_sales")
async def get_all_sales():
    try:

        # if excel_files/unit_sales.xlsx exists, delete it
        is_exists = os.path.exists("excel_files/unit_sales.xlsx")
        if is_exists:
            os.remove("excel_files/unit_sales.xlsx")
            print("file deleted")
        else:
            print("file does not exist")

        result = list(sales_processed.find())

        for item in result:
            item['_id'] = str(item['_id'])
            if 'opportunity_bond_amount' not in item:
                item['opportunity_bond_amount'] = 0
            if 'opportunity_parking_cost' not in item:
                item['opportunity_parking_cost'] = 0

            if 'opportunity_stove_cost' not in item:
                item['opportunity_stove_cost'] = 0

            if 'opportunity_pay_type' not in item:
                item['opportunity_pay_type'] = ""
        # print(result)

        create_excel_file(result, "unit_sales.xlsx")
        # return result
        return {"done": "excel_files/unit_sales.xlsx"}
    except Exception as err:
        print(err)
        return {"error": "error"}


@sales.get("/get_salesFile")
async def get_uploaded_file(file_name):
    print(file_name)

    try:  # File Name incl path.
        is_exists = os.path.exists(file_name)
        print(is_exists)
        if is_exists:
            return FileResponse(f"{file_name}", filename=file_name)
        else:
            return {"ERROR": "File does not exist!!"}
    except Exception as err:
        print(err)
        return {"ERROR": "File does not exist!!"}


@sales.post("/get_salesAgents")
async def get_salesAgents():
    try:
        result = list(sales_agents.find())
        for item in result:
            item['_id'] = str(item['_id'])

        return result
    except Exception as err:
        print(err)
        return {"error": "error"}


@sales.post("/get_morgtage_brokers")
async def get_morgtage_brokers():
    try:
        result = list(mortgage_brokers.find())
        for item in result:
            item['_id'] = str(item['_id'])

        return result
    except Exception as err:
        print(err)
        return {"error": "error"}


@sales.post("/post_newAgent")
async def post_newAgent(data: Request):
    request = await data.json()
    print(request)
    try:
        sales_agents.insert_one(request)
        return {"done": True}
    except Exception as err:
        print(err)
        return {"done": False}
    # return {"done": True}


@sales.post("/post_new_broker")
async def post_new_broker(data: Request):
    request = await data.json()
    print(request)
    try:
        mortgage_brokers.insert_one(request)
        return {"done": True}
    except Exception as err:
        print(err)
        return {"done": False}
    # return {"done": True}


@sales.post("/delete_agent")
async def delete_agent(data: Request):
    request = await data.json()
    print(request)
    # return {"done": True}
    try:
        sales_agents.delete_one({"_id": ObjectId(request['_id'])})
        return {"done": True}
    except Exception as err:
        print(err)
        return {"done": False}


@sales.post("/delete_broker")
async def delete_broker(data: Request):
    request = await data.json()
    print(request)
    # return {"done": True}
    try:
        mortgage_brokers.delete_one({"_id": ObjectId(request['_id'])})
        return {"done": True}
    except Exception as err:
        print(err)
        return {"done": False}


@sales.post("/print_onboarding_doc")
async def print_onboarding_doc(data: Request):
    request = await data.json()
    newData = request['data']

    doc_name = f"sales_client_onboarding_docs/{newData['opportunity_code']}-onboarding.pdf"
    is_exists = os.path.exists(doc_name)

    if is_exists:
        os.remove(doc_name)
        print("Successfully Removed")
    else:
        print("No such Document")

    try:
        result = print_onboarding_pdf(newData)

        # print("RESULT", result)
        return {"fileName": result}
    except Exception as err:
        print(err)
        return {"done": False}


@sales.get("/get_onboarding_doc")
async def get_uploaded_file(file_name):
    print(file_name)

    try:  # File Name incl path.
        is_exists = os.path.exists(f"sales_client_onboarding_docs/{file_name}")
        print(is_exists)
        if is_exists:
            return FileResponse(f"sales_client_onboarding_docs/{file_name}", filename=file_name)
        else:
            return {"ERROR": "File does not exist!!"}
    except Exception as err:
        print(err)
        return {"ERROR": "File does not exist!!"}


@sales.post("/print_otp_doc")
async def print_otp_doc(data: Request):
    request = await data.json()
    newData = request['data']

    doc_name = f"sales_client_onboarding_docs/{newData['opportunity_code']}-OTP.pdf"
    print("doc_name", doc_name)
    is_exists = os.path.exists(doc_name)

    if is_exists:
        os.remove(doc_name)
        print("Successfully Removed")
    else:
        print("No such Document")

    # try:
    # print("newData", newData)
    result = print_otp_pdf(newData)

    print("RESULT", result)
    return {"fileName": result}
    # except Exception as err:
    #     print("XXXXX", err)
    #     return {"done": False}
