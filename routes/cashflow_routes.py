from fastapi import APIRouter, Request, BackgroundTasks
from fastapi.responses import FileResponse
from configuration.db import db
from bson.objectid import ObjectId

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
        for item in data:
            item['_id'] = str(item['_id'])
        # data = list(data)
        return {"success": True, "data": data}
    except Exception as e:
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
