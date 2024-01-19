import os
from bson import ObjectId
from fastapi import APIRouter, Request, BackgroundTasks
from fastapi.responses import FileResponse
from excel_sf_functions.sales_forecast_excel import create_sales_forecast_file
from excel_sf_functions.sales_forecast_excel import create_cash_flow
from configuration.db import db
import time
from datetime import datetime
from datetime import timedelta
from excel_sf_functions.draw_downs_excel import create_draw_down_file
from rental_excel_functions.rental_excel_function import rental_report

rentals = APIRouter()

# MONGO COLLECTIONS
investors = db.investors
rates = db.rates
opportunities = db.opportunities
unallocated_investments = db.unallocated_investments
sales_parameters = db.salesParameters
rollovers = db.investorRollovers


@rentals.get("/drag_developments")
async def draw_developments():
    try:
        # get Category field from the opportunities collection and return a list of unique values
        developments = list(opportunities.find({}, {"Category": 1, "_id": 0}))
        # for development in developments:
        #     development["_id"] = str(development["_id"])

        print(developments)
        return developments

    except Exception as e:
        print(e)
        return {"error": str(e)}


@rentals.get("/rental_developments")
async def rental_developments():
    try:
        # get Category field from the opportunities collection and return a list of unique values
        developments = list(opportunities.find({}, {"Category": 1, "_id": 1, "opportunity_final_transfer_date": 1,
                                                    "opportunity_code": 1, "opportunity_sold": 1,
                                                    "rental_marked_for_rent": 1, "rental_rented_out": 1,
                                                    "rental_gross_amount": 1, "rental_deposit_amount": 1,
                                                    "rental_levy_amount": 1, "rental_commission": 1, "rental_rates": 1,
                                                    "rental_other_expenses": 1, "rental_nett_amount": 1,
                                                    "rental_start_date": 1, "rental_end_date": 1,
                                                    "rental_income_to_date": 1, "rental_income_to_contract_end": 1,
                                                    "rental_name": 1, "rental_agent": 1, "rental_placement_fee": 1,
                                                    "rental_admin_fee": 1, "rental_commission_percent": 1}))
        for development in developments:
            development["_id"] = str(development["_id"])
            development["rental_marked_for_rent"] = development.get("rental_marked_for_rent", False)
            development["rental_rented_out"] = development.get("rental_rented_out", False)
            development["rental_start_date"] = development.get("rental_start_date", "")
            development["rental_end_date"] = development.get("rental_end_date", "")
            development["rental_income_to_date"] = development.get("rental_income_to_date", 0)
            development["rental_income_to_contract_end"] = development.get("rental_income_to_contract_end", 0)
            development["rental_gross_amount"] = development.get("rental_gross_amount", 0)
            development["rental_deposit_amount"] = development.get("rental_deposit_amount", 0)
            development["rental_levy_amount"] = development.get("rental_levy_amount", 0)
            development["rental_commission"] = development.get("rental_commission", 0)
            development["rental_rates"] = development.get("rental_rates", 0)
            development["rental_other_expenses"] = development.get("rental_other_expenses", 0)
            development["rental_nett_amount"] = development.get("rental_nett_amount", 0)
            development["rental_name"] = development.get("rental_name", "")
            development["rental_agent"] = development.get("rental_agent", "")
            development["rental_placement_fee"] = development.get("rental_placement_fee", 0)
            development["rental_admin_fee"] = development.get("rental_admin_fee", 0)
            development["rental_commission_percent"] = development.get("rental_commission_percent", 0)

        # filter out from developments where opportunity_final_transfer_date does not equal to "" or None
        developments = [development for development in developments if
                        development["opportunity_final_transfer_date"] is None or development[
                            "opportunity_final_transfer_date"] == ""]

        return developments

    except Exception as e:
        print(e)
        return {"error": str(e)}

    # rental_developments = list(opportunities.find({}))
    # for rental in rentals:
    #     rental["_id"] = str(rental["_id"])


@rentals.post('/add_rental')
async def add_rental(data: Request):
    insert = await data.json()
    # print(insert)
    id = insert["_id"]
    del insert["_id"]
    del insert["Category"]
    del insert["opportunity_final_transfer_date"]
    del insert["opportunity_code"]
    del insert["opportunity_sold"]

    # loop insert and update opportunity collection with the fields in insert and an _id of id
    for key, value in insert.items():
        opportunities.update_one({"_id": ObjectId(id)}, {"$set": {key: value}})
    # print(insert)
    return {"res": "ok"}


@rentals.get('/generate_rental_report')
async def generate_rental_report():
    # get today's date and format it as YYYY-MM-DD
    # today = datetime.now().strftime("%Y-%m-%d")
    # print(today)
    # from the opportunities in the database get all the records where rental_rented_out field exists and is equal to
    # True
    final_rentals = []
    try:
        results = list(opportunities.find({"rental_rented_out": True}))
        for result in results:
            del result['_id']
            insert = {}
            insert['Category'] = result['Category']
            insert['opportunity_code'] = result['opportunity_code']
            insert['opportunity_sold'] = result.get('opportunity_sold', False)
            insert['rental_marked_for_rent'] = result.get('rental_marked_for_rent', False)
            insert['rental_rented_out'] = result.get('rental_rented_out', False)
            insert['rental_start_date'] = result.get('rental_start_date', "")
            insert['rental_end_date'] = result.get('rental_end_date', "")
            insert['rental_name'] = result.get('rental_name', "")
            insert['rental_agent'] = result.get('rental_agent', "")
            insert['rental_gross_amount'] = float(result.get('rental_gross_amount', 0))
            insert['rental_deposit_amount'] = float(result.get('rental_deposit_amount', 0))
            insert['rental_levy_amount'] = float(result.get('rental_levy_amount', 0))
            insert['rental_commission'] = float(result.get('rental_commission', 0))
            insert['rental_rates'] = float(result.get('rental_rates', 0))
            insert['rental_other_expenses'] = float(result.get('rental_other_expenses', 0))
            insert['rental_nett_amount'] = float(result.get('rental_nett_amount', 0))
            insert['rental_income_to_date'] = float(result.get('rental_income_to_date', 0))
            insert['rental_income_to_contract_end'] = float(result.get('rental_income_to_contract_end', 0))
            insert['rental_placement_fee'] = float(result.get('rental_placement_fee', 0))
            insert['rental_admin_fee'] = float(result.get('rental_admin_fee', 0))
            insert['rental_commission_percent'] = float(result.get('rental_commission_percent', 0)) / 100
            final_rentals.append(insert)
        # sort final_results by Category, opportunity_code, rental_start_date
        final_rentals = sorted(final_rentals, key=lambda i: (i['Category'], i['opportunity_code'],
                                                             i['rental_start_date']))
        report_done = rental_report(final_rentals)
        print(report_done)
        return {"res": "ok"}
    except Exception as e:
        print(e)
        results = []
        return {"error": str(e)}


@rentals.get("/get_rental_report")
async def get_rental_report(file_name):
    try:
        file_name = file_name + ".xlsx"

        dir_path = "rental_excel_functions"
        dir_list = os.listdir(dir_path)

        if file_name in dir_list:

            return FileResponse(f"{dir_path}/{file_name}", filename=file_name)
        else:
            return {"ERROR": "File does not exist!!"}
    except Exception as e:
        print(e)
        return {"ERROR": "Please Try again"}