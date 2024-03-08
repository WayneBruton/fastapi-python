from datetime import datetime, timedelta

from fastapi import APIRouter, Request, BackgroundTasks
from fastapi.responses import FileResponse
from configuration.db import db
from bson.objectid import ObjectId
import time

cashflow = APIRouter()


@cashflow.post("/construction_cashflow")
async def construction_cashflow(data: Request):
    request = await data.json()
    data = request['data']
    try:
        db.cashflow_construction.delete_many({})
        result = db.cashflow_construction.insert_many(data)
        return {"success": True}
    except Exception as e:
        return {"success": False, "error": str(e)}


@cashflow.post("/trial-balance_cashflow")
async def trial_balance_cashflow(data: Request):
    request = await data.json()
    data = request['data']
    try:
        db.cashflow_trial_balance.delete_many({})
        result = db.cashflow_trial_balance.insert_many(data)
        return {"success": True}
    except Exception as e:
        return {"success": False, "error": str(e)}


@cashflow.get("/construction_cashflow")
async def get_construction_cashflow():
    try:
        data = list(db.cashflow_construction.find({}))
        # print(data[0])
        for item in data:
            item['_id'] = str(item['_id'])
        # data = list(data)
        # sort data by Whitebox-Able in descending order

        data = sorted(data, key=lambda x: (x['Whitebox-Able']), reverse=True)

        return {"success": True, "data": data}
    except Exception as e:
        print("ERROR", e)
        return {"success": False, "error": str(e)}


@cashflow.get("/trial-balance_cashflow")
async def get_trial_balance_cashflow():
    try:
        data = list(db.cashflow_trial_balance.find({}))
        for item in data:
            item['_id'] = str(item['_id'])
        # data = list(data)
        return {"success": True, "data": data}
    except Exception as e:
        return {"success": False, "error": str(e)}


@cashflow.post("/update_construction_data")
async def update_construction_data(data: Request):
    request = await data.json()
    data = request['data']
    _id = data["_id"]
    del data["_id"]
    try:
        db.cashflow_construction.update_one({"_id": ObjectId(_id)}, {"$set": data})
        return {"success": True}
    except Exception as e:
        return {"success": False, "error": str(e)}


# SALES HELPER FUNCTIONS
def get_sales_parameters():
    try:
        sales_parameters = list(db.salesParameters.find({}, {"_id": 0}))
        return sales_parameters
    except Exception as e:
        print("Error getting sales parameters")
        return {"success": False, "error": str(e)}


def get_rates():
    try:
        rates = list(db.rates.find({}, {"_id": 0}))
        # print("Rates", rates)
        for rate in rates:
            rate['rate'] = float(rate['rate'])
            rate['Efective_date'] = rate['Efective_date'].replace("-", "/")
            rate['Efective_date'] = datetime.strptime(rate['Efective_date'], '%Y/%m/%d')
        return rates
    except Exception as e:
        print("Error getting rates")
        return {"success": False, "error": str(e)}


def get_opportunities(data):
    try:
        opportunities = list(db.opportunities.find({"Category": {"$in": data}},
                                                   {"opportunity_code": 1, "Category": 1, "opportunity_end_date": 1,
                                                    "opportunity_final_transfer_date": 1, "opportunity_sale_price": 1,
                                                    "opportunity_sold": 1, "_id": 0}))
        for opportunity in opportunities:
            if opportunity.get('opportunity_final_transfer_date', "") == "":
                opportunity['transferred'] = False
            else:
                opportunity['transferred'] = True
        return opportunities
    except Exception as e:
        print("Error getting opportunities")
        return {"success": False, "error": str(e)}


def get_investors(data):
    try:
        final_investors = []

        investors = list(db.investors.find({}, {"_id": 0}))
        for investor in investors:
            investor_trust = list(
                filter(lambda x: x['Category'] in data and x['release_date'] == '', investor['trust']))
            investor_investments = list(
                filter(lambda x: x['Category'] in data, investor['investments']))
            if len(investor_trust) > 0:
                for item in investor_trust:
                    # print()
                    insert = {
                        "investor_acc_number": investor['investor_acc_number'],
                        "investor_name": investor['investor_name'],
                        "investor_surname": investor['investor_surname'],
                        "Category": item['Category'],
                        "opportunity_code": item['opportunity_code'],
                        "Block": item['opportunity_code'][-4],
                        "deposit_date": item['deposit_date'],
                        "release_date": item['release_date'],
                        "end_date": item.get('end_date', ""),
                        "investment_number": item.get('investment_number', 0),
                        "investment_amount": float(item['investment_amount']),
                        "investment_interest_rate": float(item['investment_interest_rate']),
                        "early_release": item.get('early_release', False),
                    }
                    final_investors.append(insert)

            for item in investor_investments:
                filtered_trust = list(filter(
                    lambda x: x['opportunity_code'] == item['opportunity_code'] and float(
                        x['investment_amount']) == float(item['investment_amount']), investor['trust']))

                insert = {
                    "investor_acc_number": investor['investor_acc_number'],
                    "investor_name": investor['investor_name'],
                    "investor_surname": investor['investor_surname'],
                    "Category": item['Category'],
                    "opportunity_code": item['opportunity_code'],
                    "Block": item['opportunity_code'][-4],
                    "deposit_date": filtered_trust[0]['deposit_date'],
                    "release_date": item['release_date'],
                    "end_date": item.get('end_date', ""),
                    "investment_number": item.get('investment_number', 0),
                    "investment_amount": float(item['investment_amount']),
                    "investment_interest_rate": float(item['investment_interest_rate']),
                    "early_release": item.get('early_release', False),
                }

                final_investors.append(insert)

        final_investors = [investor for investor in final_investors if
                           not (investor['investor_acc_number'] == "ZCAM01" and investor[
                               'opportunity_code'] == "HFA101" and investor['investment_amount'] == 400000.0)]

        final_investors = [investor for investor in final_investors if
                           not (investor['investor_acc_number'] == "ZJHO01" and investor[
                               'opportunity_code'] == "HFA304" and investor['investment_number'] == 1)]

        final_investors = [investor for investor in final_investors if
                           not (investor['investor_acc_number'] == "ZPJB01" and investor[
                               'opportunity_code'] == "HFA205" and investor['investment_number'] == 1)]

        final_investors = [investor for investor in final_investors if
                           not (investor['investor_acc_number'] == "ZERA01" and investor[
                               'opportunity_code'] == "EA205" and investor['investment_number'] == 3)]

        final_investors = [investor for investor in final_investors if
                           not (investor['investor_acc_number'] == "ZVOL01" and investor[
                               'opportunity_code'] == "EA103" and investor['investment_number'] == 3)]

        final_investors = [investor for investor in final_investors if
                           not (investor['investor_acc_number'] == "ZLEW03" and investor[
                               'opportunity_code'] == "EA205" and investor['investment_number'] == 1)]

        # sort final investors first by opportunity_code then by investor_acc_number
        final_investors = sorted(final_investors, key=lambda x: (x['opportunity_code'], x['investor_acc_number']))

        return final_investors
    except Exception as e:
        print("Error getting investors")
        return {"success": False, "error": str(e)}


def get_construction_costs():
    try:
        construction_costs = list(db.cashflow_construction.find({}, {"Complete Build": 1,
                                                                     "Whitebox-Able": 1, "Blocks": 1, "_id": 0}))
        # filter construction costs to only include whiteboxe-Able

        construction_costs = list(filter(lambda x: x['Whitebox-Able'] == True, construction_costs))
        construction_costs = list(filter(lambda x: x['Complete Build'] == True, construction_costs))
        return construction_costs
    except Exception as e:
        print("Error getting construction costs")
        return {"success": False, "error": str(e)}


def get_sales_for_cashflow(data):
    try:
        sales = list(db.sales_processed.find({"development": {"$in": data}},
                                             {"development": 1, "opportunity_code": 1, "opportunity_sales_date": 1,
                                              "opportunity_potential_reg_date": 1, "opportunity_actual_reg_date": 1,
                                              "opportunity_bond_registration": 1, "opportunity_commission": 1,
                                              "opportunity_transfer_fees": 1, "opportunity_trust_release_fee": 1,
                                              "opportunity_unforseen": 1, "_id": 0}))
        return sales
    except Exception as e:
        print("Error getting sales")
        return {"success": False, "error": str(e)}


def calculate_cashflow_investors(final_investors, opportunities, sales_data):
    # CALCULATE INITIAL DUE TO INVESTOR - final_investors, opportunities, rates
    final_interest_calculations = []
    opportunity_codes = []
    rates = get_rates()
    # print("final_investors", final_investors[0])
    try:
        for investor in final_investors:
            # investor['_id'] = str(investor['_id'])

            if investor['release_date'] == "":
                investor['not_released'] = True
                # convert investor['deposit_date'] to datetime
                investor['deposit_date'] = (investor['deposit_date'].replace("-", "/"))
                investor['deposit_date'] = datetime.strptime(investor['deposit_date'], '%Y/%m/%d')
                investor['release_date'] = investor['deposit_date'] + timedelta(days=30)
                filtered_opportunity = list(
                    filter(lambda x: x['opportunity_code'] == investor['opportunity_code'], opportunities))
                filtered_sales_data = list(
                    filter(lambda x: x['opportunity_code'] == investor['opportunity_code'], sales_data))
                if len(filtered_opportunity) > 0:
                    if filtered_sales_data[0]['forecast_transfer_date'] == "":
                        investor['end_date'] = filtered_opportunity[0]['opportunity_end_date']
                    else:
                        investor['end_date'] = filtered_sales_data[0]['forecast_transfer_date']
                    investor['end_date'] = (investor['end_date'].replace("-", "/"))
                    investor['end_date'] = datetime.strptime(investor['end_date'], '%Y/%m/%d')

            else:
                investor['not_released'] = False
                investor['deposit_date'] = (investor['deposit_date'].replace("-", "/"))
                investor['deposit_date'] = datetime.strptime(investor['deposit_date'], '%Y/%m/%d')
                investor['release_date'] = (investor['release_date'].replace("-", "/"))
                investor['release_date'] = datetime.strptime(investor['release_date'], '%Y/%m/%d')
                if investor['end_date'] == "":
                    filtered_opportunity = list(
                        filter(lambda x: x['opportunity_code'] == investor['opportunity_code'], opportunities))
                    filtered_sales_data = list(
                        filter(lambda x: x['opportunity_code'] == investor['opportunity_code'], sales_data))
                    if len(filtered_opportunity) > 0:
                        if filtered_sales_data[0]['forecast_transfer_date'] == "":
                            investor['end_date'] = filtered_opportunity[0]['opportunity_end_date']
                        else:
                            investor['end_date'] = filtered_sales_data[0]['forecast_transfer_date']

                investor['end_date'] = (investor['end_date'].replace("-", "/"))
                investor['end_date'] = datetime.strptime(investor['end_date'], '%Y/%m/%d')

            momentum_start_date = investor['deposit_date'] + timedelta(days=1)
            momentum_end_date = investor['release_date']

            interest = 0
            momentum_interest = 0
            investment_interest = 0
            while momentum_start_date <= momentum_end_date:
                rate = list(filter(lambda x: x['Efective_date'] <= momentum_start_date, rates))
                # sort by Efective_date descending
                rate = sorted(rate, key=lambda x: x['Efective_date'], reverse=True)
                rate = rate[0]['rate'] / 100
                momentum_interest += investor['investment_amount'] * rate / 365
                momentum_start_date += timedelta(days=1)
            investor['momentum_interest'] = momentum_interest
            investment_start_date = investor['release_date']
            investment_end_date = investor['end_date']
            # I NEED TO CHANGE THIS TO THE FORECAST DATE DATES ARE AN ISSUE
            # filtered_sales_data = list(
            #     filter(lambda x: x['opportunity_code'] == investor['opportunity_code'], sales_data))
            # if len(filtered_sales_data) > 0:
            #     trial_investment_end_date = filtered_sales_data[0]['forecast_transfer_date']
            #     if trial_investment_end_date != "":
            #         trial_investment_end_date = trial_investment_end_date.replace("-", "/")
            #         trial_investment_end_date = datetime.strptime(trial_investment_end_date, '%Y/%m/%d')
            #     if investor['opportunity_code'] == "HVC306":
            #         print("trial_investment_end_date", trial_investment_end_date)

            investment_interest += investor['investment_amount'] * investor['investment_interest_rate'] / 100 / 365 * (
                    investment_end_date - investment_start_date).days
            investor['investment_interest'] = investment_interest
            interest = momentum_interest + investment_interest
            investor['interest'] = interest

            opportunity_codes.append(investor['opportunity_code'])

        # Get unique values from opportunity_codes
        opportunity_codes = list(sorted(set(opportunity_codes)))
        # print("opportunity_codes",opportunity_codes)
        for opportunity in opportunity_codes:
            filtered_investors = list(filter(lambda x: x['opportunity_code'] == opportunity, final_investors))
            # Total_due_to_investors = sum of investment_amount + interest
            total_due_to_investors = sum([x['investment_amount'] + x['interest'] for x in filtered_investors])
            insert = {
                "opportunity_code": opportunity,
                "total_due_to_investors": total_due_to_investors,
            }
            final_interest_calculations.append(insert)

            # print(investor)
        # print(len(final_investors))
        # filter out from investors where early_release is true
        # final_investors = list(filter(lambda x: x['early_release'] == False, final_investors))
        # print(len(final_investors))
        # print("final_investors", final_investors[0])
        # print("final_interest_calculations", final_interest_calculations[0])
        return final_interest_calculations
    except Exception as e:
        print("Error calculating cashflow investors", e)
        return {"success": False, "error": str(e)}

    # I AM HERE


@cashflow.post("/get_sales_cashflow_initial")
async def get_sales_cashflow_initial(data: Request):
    request = await data.json()
    # print(request)
    data = request['data']
    try:
        start = time.time()
        opportunities = get_opportunities(data)

        sales_data = list(db.cashflow_sales.find({}))

        # print("sales_data", sales_data[0])
        # print()

        construction = get_construction_costs()
        for item in construction:
            item['block'] = item['Blocks'].replace("Block ", "")
            # remove leading and trailing whitespace from item['block']
            item['block'] = item['block'].strip()

            del item['Blocks']

        if len(sales_data) == len(opportunities):

            final_investors = get_investors(data)
            # print("final_investors", final_investors)
            investors_with_interest = calculate_cashflow_investors(final_investors, opportunities, sales_data)

            items_to_update = []
            for item in sales_data:

                if item['refinanced'] == False and item['VAT'] == 0:
                    final_investors_filtered = list(
                        filter(lambda x: x['opportunity_code'] == item['opportunity_code'], investors_with_interest))
                    # print("final_investors_filtered", final_investors_filtered)
                    if len(final_investors_filtered) > 0:
                        # print("final_investors_filtered", final_investors_filtered)
                        # print()
                        # print("item", item)
                        item['due_to_investors'] = final_investors_filtered[0]['total_due_to_investors']

                    else:
                        item['due_to_investors'] = 0

                    item['VAT'] = float(item['sale_price']) / 1.15 * 0.15
                    item['nett'] = float(item['sale_price']) / 1.15

                    item['transfer_income'] = float(item['sale_price']) - float(
                        item['opportunity_transfer_fees']) - float(item[
                                                                       'opportunity_trust_release_fee']) - float(
                        item['opportunity_unforseen']) - float(item[
                                                                   'opportunity_bond_registration']) - float(
                        item['opportunity_commission'])
                    item['profit_loss'] = item['transfer_income'] - item['due_to_investors']
                    items_to_update.append(item)

                item['_id'] = str(item['_id'])
                filtered_construction = list(
                    filter(lambda x: x['block'] == item['block'], construction))

                # print("filtered_construction", filtered_construction)

                if len(filtered_construction) > 0:

                    if item['complete_build']:
                        item['complete_build'] = False
                        items_to_update.append(item) # I AM HERE Now I need to update the sales data
                else:

                    if not item['complete_build']:
                        item['complete_build'] = True

                        items_to_update.append(item)

                # DO SALES PRICE, PROFIT AND INVESTOR INTEREST CALCULATIONS ONLY IF DIFFERENT I AM HERE
                filtered_investors = list(filter(lambda x: x['opportunity_code'] == item['opportunity_code'],
                                                 investors_with_interest))
                if len(filtered_investors) > 0:
                    if item['due_to_investors'] == 0 or float(item['due_to_investors']) != float(filtered_investors[0][
                                                                                                     'total_due_to_investors']):
                        item['due_to_investors'] = filtered_investors[0]['total_due_to_investors']
                        item['profit_loss'] = item['transfer_income'] - item['due_to_investors']
                        items_to_update.append(item)
                filtered_opportunities = list(
                    filter(lambda x: x['opportunity_code'] == item['opportunity_code'], opportunities))
                if len(filtered_opportunities) > 0:
                    if float(item['sale_price']) != float(filtered_opportunities[0]['opportunity_sale_price']):
                        item['sale_price'] = filtered_opportunities[0]['opportunity_sale_price']
                        item['VAT'] = float(item['sale_price']) / 1.15 * 0.15
                        item['nett'] = float(item['sale_price']) / 1.15
                        item['transfer_income'] = (float(item['sale_price']) - item['opportunity_transfer_fees'] -
                                                   item['opportunity_trust_release_fee'] - item[
                                                       'opportunity_unforseen'] - item[
                                                       'opportunity_commission'] * float(item['sale_price']) -
                                                   item['opportunity_bond_registration'])
                        items_to_update.append(item)

            if len(items_to_update) > 0:

                items_to_update = list({v['_id']: v for v in items_to_update}.values())
                try:
                    for index, item in enumerate(items_to_update):
                        id1 = item['_id']

                        del item['_id']
                        db.cashflow_sales.update_one({"_id": ObjectId(id1)}, {"$set": item})
                        item['_id'] = id1
                except Exception as e:
                    print("Error updating sales data", e, index, item['opportunity_code'])
                    # print()
                    return {"success": False, "error": str(e)}

            end = time.time()
            for item in sales_data:
                # convert item['sale_price'] to currency format with R symbol
                item['sale_price_nice'] = "R" + "{:,.2f}".format(float(item['sale_price']))
                item['profit_loss_nice'] = "R" + "{:,.2f}".format(float(item['profit_loss']))

            # sort sales_data by transferred then by opportunity_code
            sales_data = sorted(sales_data, key=lambda x: (x['transferred'], x['opportunity_code']))

            # for index, data in enumerate(sales_data):
            #     if index == 0:
            #         print("sales_data", data['_id'])

            return {"success": True, "data": sales_data, "time": end - start}

        else:

            sales_parameters = get_sales_parameters()

            # rates = get_rates()

            sales = get_sales_for_cashflow(data)

            opportunities_sold = list(filter(lambda x: x['opportunity_sold'] == True, opportunities))

            sales = list(
                filter(lambda x: x['opportunity_code'] in [item['opportunity_code'] for item in opportunities_sold],
                       sales))

            construction = get_construction_costs()

            final_investors = get_investors(data)

            # Calculate INITIAL DATA
            final_sales = []
            for unit in opportunities:
                filtered_construction = list(
                    filter(lambda x: x['Blocks'] == "Block " + unit['opportunity_code'][-4], construction))
                if len(filtered_construction) > 0:
                    print("filtered_construction", filtered_construction[0])
                    complete_build = filtered_construction[0]['Complete Build']
                    # If whiteboxed is true, then make it false and vica a versa
                    complete_build = not complete_build
                else:
                    complete_build = False

                if unit['opportunity_final_transfer_date'] == "":
                    original_planned_transfer_date = unit['opportunity_end_date']
                    forecast_transfer_date = ""
                else:
                    original_planned_transfer_date = unit['opportunity_final_transfer_date']
                    forecast_transfer_date = unit['opportunity_final_transfer_date']
                filtered_sales = list(filter(lambda x: x['opportunity_code'] == unit['opportunity_code'], sales))
                if len(filtered_sales) > 0:
                    # print("Filtered Sales: ",filtered_sales)
                    opportunity_transfer_fees = float(filtered_sales[0].get('opportunity_transfer_fees', 0))
                    opportunity_trust_release_fee = float(filtered_sales[0].get('opportunity_trust_release_fee', 0))
                    opportunity_unforseen = float(filtered_sales[0].get('opportunity_unforseen', 0)) * float(
                        unit['opportunity_sale_price'])
                    opportunity_commission = float(filtered_sales[0].get('opportunity_commission', 0)) / 1.15 * float(
                        unit['opportunity_sale_price'])
                    opportunity_bond_registration = float(filtered_sales[0].get('opportunity_bond_registration', 0))
                    if opportunity_transfer_fees == 0:
                        # if unit['opportunity_code'] == "HVC306":
                        #     print("Filtered Sales: ", filtered_sales)
                        #     print()

                        filtered_sales_parameters = list(
                            filter(
                                lambda x: x['Development'] == unit['Category'] and x['Description'] == "transfer_fees",
                                sales_parameters))

                        opportunity_transfer_fees = float(filtered_sales_parameters[0]['rate'])
                        filtered_sales_parameters = list(
                            filter(
                                lambda x: x['Development'] == unit['Category'] and x[
                                    'Description'] == "trust_release_fee",
                                sales_parameters))
                        opportunity_trust_release_fee = float(filtered_sales_parameters[0]['rate'])
                        filtered_sales_parameters = list(
                            filter(
                                lambda x: x['Development'] == unit['Category'] and x['Description'] == "unforseen",
                                sales_parameters))

                        opportunity_unforseen = float(filtered_sales_parameters[0]['rate']) * float(
                            unit['opportunity_sale_price'])
                        filtered_sales_parameters = list(
                            filter(
                                lambda x: x['Development'] == unit['Category'] and x['Description'] == "commission",
                                sales_parameters))
                        # if unit['opportunity_code'] == "HVC306":
                        #     print("filtered_sales_parameters", filtered_sales_parameters)
                        opportunity_commission = float(filtered_sales_parameters[0]['rate']) / 1.15 * float(
                            unit['opportunity_sale_price'])
                        filtered_sales_parameters = list(
                            filter(
                                lambda x: x['Development'] == unit['Category'] and x[
                                    'Description'] == "bond_registration",
                                sales_parameters))
                        opportunity_bond_registration = float(filtered_sales_parameters[0]['rate'])

                opportunity_unforseen = float(
                    unit['opportunity_sale_price']) * .005

                opportunity_commission = float(
                    unit['opportunity_sale_price']) * .05

                insert = {
                    "Category": unit['Category'],
                    "block": unit['opportunity_code'][-4],
                    "opportunity_code": unit['opportunity_code'],
                    "sold": unit['opportunity_sold'],
                    "transferred": unit['transferred'],
                    "complete_build": complete_build,
                    "original_planned_transfer_date": original_planned_transfer_date,
                    "forecast_transfer_date": forecast_transfer_date,
                    "sale_price": float(unit['opportunity_sale_price']),
                    "VAT": float(unit['opportunity_sale_price']) / 1.15 * 0.15,
                    "nett": float(unit['opportunity_sale_price']) / 1.15,
                    "opportunity_transfer_fees": opportunity_transfer_fees,
                    "opportunity_trust_release_fee": opportunity_trust_release_fee,
                    "opportunity_unforseen": opportunity_unforseen,
                    "opportunity_commission": opportunity_commission,
                    "opportunity_bond_registration": opportunity_bond_registration,
                    "transfer_income": float(unit[
                                                 'opportunity_sale_price']) - opportunity_transfer_fees - opportunity_trust_release_fee - opportunity_unforseen - opportunity_commission - opportunity_bond_registration,
                    "due_to_investors": 0,
                    "profit_loss": 0,
                    "refinanced": False,
                }
                final_sales.append(insert)

            # print("final_sales", len(final_sales))

            if len(sales_data) == 0:
                try:
                    db.cashflow_sales.insert_many(final_sales)
                except Exception as e:
                    print("Error inserting sales data", e)
                    return {"success": False, "error": str(e)}
            else:
                if len(sales_data) < len(final_sales):
                    for sale in final_sales:
                        if sale not in sales_data:
                            try:
                                db.cashflow_sales.insert_one(sale)
                            except Exception as e:
                                print("Error inserting sales data", e)
                                return {"success": False, "error": str(e)}

            # print("insert", insert)
            # print()

        # print("final_investors", final_investors[10])
        # print("final_investors", len(final_investors))
        # print("opportunities", opportunities[0])
        # print("sales_data", sales_data)
        # print("sales_parameters", sales_parameters)
        # print("rates", rates)
        # print("investors", investors[11])
        # print("sales", sales[0])
        end = time.time()
        print("Time taken", end - start)
        return {"success": True, "data": sales_data, "time": end - start}
    except Exception as e:
        print("Error getting sales cashflow initial", e)
        return {"success": False, "error": str(e)}
    # return {"success": True, "data": data}


@cashflow.post("/update_sales_data")
async def update_sales_data(data: Request):
    request = await data.json()
    data = request['data']
    for item in data:
        _id = item["_id"]
        del item["_id"]
        try:
            db.cashflow_sales.update_one({"_id": ObjectId(_id)}, {"$set": item})
            print("Updated")
        except Exception as e:
            print("Error updating sales data", e)
            return {"success": False, "error": str(e)}
    return {"success": True}


def calculate_vat_due(sale_date):
    vat_periods = {
        1: "03/31",
        2: "03/31",
        3: "05/31",
        4: "05/31",
        5: "07/31",
        6: "07/31",
        7: "09/30",
        8: "09/30",
        9: "11/30",
        10: "11/30",
        11: "01/31",
        12: "01/31",
    }

    sale_date = datetime.strptime(sale_date.replace("-", "/"), '%Y/%m/%d')

    sale_month = sale_date.month
    sale_year = sale_date.year
    if sale_month > 10:
        vat_year = sale_year + 1
    else:
        vat_year = sale_year

    vat_date = f"{vat_year}/{vat_periods[sale_month]}"
    return vat_date

# result = calculate_vat_due("2022-12-01")
# print(result)
