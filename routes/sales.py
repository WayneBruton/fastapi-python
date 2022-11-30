import os
from fastapi import APIRouter, Request
from fastapi.responses import FileResponse
from config.db import db
from bson.objectid import ObjectId

sales = APIRouter()

# MONGO COLLECTIONS
investors = db.investors
opportunityCategories = db.opportunityCategories
opportunities = db.opportunities
sales_parameters = db.salesParameters
sales_processed = db.sales_processed


@sales.post("/getDevelopmentForSales")
async def get_developments_for_sales():
    result_developments = list(opportunityCategories.aggregate([{"$project": {
        "id2": {'$toString': "$_id"}, "_id": 0, "Description": 1, "blocked": 1
    }}]))
    return result_developments


@sales.post("/getUnitsForSales")
async def get_units_for_sales(data: Request):
    request = await data.json()
    print(request)
    result = list(opportunities.aggregate([
        {
            '$match': {
                'Category': request['development']
            }
        }, {
            '$project': {
                '_id': 0,
                'opportunity_code': 1
            }
        }
    ]))
    return result


@sales.post("/get_units_sold")
async def get_all_units_sold(data: Request):
    request = await data.json()
    print(request)
    response = list(sales_processed.aggregate([
        {
            '$match': {
                'development': request['development']
            }
        }, {
            '$project': {
                '_id': 0,
                'id': 1,
                'development': 1,
                'opportunity_code': 1,
                'opportunity_firstname': 1,
                'opportunity_lastname': 1
            }
        }
    ]))
    print(response)
    return response


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


@sales.post("/saveSale")
async def save_sale(data: Request):
    request = await data.json()
    if request['formData']['id'] != "":
        id_to_update = request['formData']['id']
        obj_instance = ObjectId(id_to_update)
        sales_processed.update_one({"_id": obj_instance}, {"$set": request['formData']})
        return {"updated": "Success"}

    else:
        obj_instance = sales_processed.insert_one(request['formData']).inserted_id
        sales_processed.update_one({"_id": obj_instance}, {"$set": {"id": str(obj_instance)}})
        return {"id": str(obj_instance)}


@sales.post("/delete_sale")
async def delete_sale(data: Request):
    request = await data.json()
    obj_instance = ObjectId(request['id'])
    result = sales_processed.delete_one({"_id": obj_instance})
    print(result)
    unit = request['unit']
    list_of_filenames = os.listdir("sales_documents")
    for file in list_of_filenames:
        length = len(unit)
        if file[:length] == unit:
            os.remove(f"sales_documents/{file}")
    return "Deleted"


@sales.post("/get_sold_unit")
async def get_sold_unit(data: Request):
    request = await data.json()
    result = sales_processed.find_one(
        {"opportunity_code": request['opportunity_code'], "development": request['development']})
    if result:
        del result["_id"]
    return result


@sales.post("/upload_file")
async def upload_file(data: Request):
    form = await data.form()
    filename = form['fileName']
    contents = await form['doc'].read()
    with open(f"sales_documents/{filename}", 'wb') as f:
        f.write(contents)
    return {"fileName": f"sales_documents/{filename}"}


@sales.post("/delete_uploaded_file")
async def delete_uploaded_file(data: Request):
    request = await data.json()
    print(request)
    is_exists = os.path.exists(request['fileToDelete'])
    print(is_exists)
    if is_exists:
        os.remove(request['fileToDelete'])
        return {"file_deleted": True}
    else:
        return {"file_deleted": False}


@sales.get("/get_uploaded_document")
async def get_uploaded_file(file_name):  # File Name incl path.
    is_exists = os.path.exists(file_name)
    if is_exists:
        return FileResponse(f"{file_name}", media_type="application/pdf")
    else:
        return {"ERROR": "File does not exist!!"}
