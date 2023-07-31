import os
from bson import ObjectId
from fastapi import APIRouter, Request, BackgroundTasks
from fastapi.responses import FileResponse
from excel_sf_functions.sales_forecast_excel import create_sales_forecast_file
from excel_sf_functions.sales_forecast_excel import create_cash_flow
from config.db import db
import time
from datetime import datetime
from datetime import timedelta
from excel_sf_functions.draw_downs_excel import create_draw_down_file

rentals = APIRouter()

# MONGO COLLECTIONS
investors = db.investors
rates = db.rates
opportunities = db.opportunities
unallocated_investments = db.unallocated_investments
sales_parameters = db.salesParameters
rollovers = db.investorRollovers


@rentals.get("/rental_developments")
async def rental_developments():
    # get Category field from the opportunities collection and return a list of unique values
    developments = list(opportunities.find({}, {"Category": 1, "_id": 1, "opportunity_final_transfer_date": 1,
                                                "opportunity_code": 1, "opportunity_sold": 1,
                                                "rental_marked_for_rent": 1, "rental_rented_out": 1,
                                                "rental_gross_amount": 1, "rental_deposit_amount": 1,
                                                "rental_levy_amount": 1, "rental_commission": 1, "rental_rates": 1,
                                                "rental_other_expenses": 1, "rental_nett_amount": 1,
                                                "rental_start_date": 1, "rental_end_date": 1,
                                                "rental_income_to_date": 1, "rental_income_to_contract_end": 1}))
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

    # filter out from developments where opportunity_final_transfer_date does not equal to "" or None
    developments = [development for development in developments if
                    development["opportunity_final_transfer_date"] is None or development[
                        "opportunity_final_transfer_date"] == ""]






    return developments

    # rental_developments = list(opportunities.find({}))
    # for rental in rentals:
    #     rental["_id"] = str(rental["_id"])

    return rentals
