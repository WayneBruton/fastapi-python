from fastapi import APIRouter, Request
# from fastapi.responses import FileResponse
from config.db import db
from deepdiff import DeepDiff
from datetime import datetime

audit = APIRouter()

# MONGO COLLECTIONS
audit_trail = db.audit_trail


# INSERT INTO AUDIT TRAIL WHEN USER LOGS ON
@audit.post("/update_audit")
async def add_to_audit_trail(data: Request):
    request = await data.json()
    posted = audit_trail.insert_one(request)
    print("Request", request)
    print(posted.inserted_id)
    return f"new trail item inserted: Id = {posted.inserted_id}"


# UPDATE INVESTOR CHANGES - BLOCK AND DELETE TO FOLLOW
@audit.post("/update_audit_investor")
async def add_to_audit_trail_investor(data: Request):
    request = await data.json()
    current = request['currentInvestor']
    original = request['original_investor']
    task = ""
    ddiff = DeepDiff(original, current)

    if ddiff.get('dictionary_item_removed') is not None:
        del ddiff['dictionary_item_removed']
    if ddiff.get('values_changed') is not None:
        for item in ddiff['values_changed']:
            task += f"{item} = {ddiff['values_changed'][item]}; \n"
    if ddiff.get('iterable_item_added') is not None:
        for item in ddiff['iterable_item_added']:
            task += f"{item} = {ddiff['iterable_item_added'][item]}; \n"
    if ddiff.get('iterable_item_removed') is not None:
        for item in ddiff['iterable_item_removed']:
            task += f"{item} = {ddiff['iterable_item_removed'][item]}; \n"

    if task:
        task = task.replace('investments', 'released')
        task = task.replace('trust', 'investment')
        task = task.replace('root', '')
        dictionary = {
            "time_stamp": request['time_stamp'],
            "user": request['user'],
            "page_reference": request['page_reference'],
            "task": task
        }
        print("task", task)
        posted = audit_trail.insert_one(dictionary)
        return f"new trail item inserted: Id = {posted.inserted_id}"

    else:
        print("No Changes made")
        return "No Changes made"


# GET ALL USERS THAT HAVE DONE STUFF ON APP - THIS IS FOR FIRST SELECT ON AUDIT PAGE
@audit.post("/getAuditUsers")
async def get_user_info_from_audit_trail():
    users = audit_trail.distinct('user')
    result = {
        "users": users
    }
    return result


# GET AUDIT TRAIL - DEPENDING ON TIME PERIOD OR USER
@audit.post('/retrieve_audit_trail')
async def get_audit_trail(data: Request):
    request = await data.json()
    print(request)
    time_stamp = request["period"]
    date_object = datetime.strptime(time_stamp, "%Y/%m/%d %H:%M:%S")
    print("date_object", date_object.hour)
    user = request["user"]
    print(user)
    if user != "All Users":
        retrieved_audit_trail = audit_trail.aggregate([
            {
                '$match': {
                    'user': request["user"]
                }
            }, {
                '$set': {
                    'time_stamp_check': '$time_stamp'
                }
            }, {
                '$project': {
                    'time_stamp_check': {
                        '$dateFromString': {
                            'dateString': '$time_stamp'
                        }
                    },
                    'user': 1,
                    'task': 1,
                    'page_reference': 1,
                    'time_stamp': 1,
                    '_id': 0
                }
            }, {
                '$match': {
                    'time_stamp_check': {
                        '$gte': datetime(date_object.year, date_object.month, date_object.day, date_object.hour,
                                         date_object.minute, date_object.second)
                    }
                }
            }, {
                '$project': {
                    'time_stamp': 1,
                    'user': 1,
                    'task': 1,
                    'page_reference': 1
                }
            },
            {
                '$sort': {
                    'time_stamp': -1
                }
            }
        ])
    else:
        retrieved_audit_trail = audit_trail.aggregate([
            {
                '$set': {
                    'time_stamp_check': '$time_stamp'
                }
            }, {
                '$project': {
                    'time_stamp_check': {
                        '$dateFromString': {
                            'dateString': '$time_stamp'
                        }
                    },
                    'user': 1,
                    'task': 1,
                    'page_reference': 1,
                    'time_stamp': 1,
                    '_id': 0
                }
            }, {
                '$match': {
                    'time_stamp_check': {
                        '$gte': datetime(date_object.year, date_object.month, date_object.day, date_object.hour,
                                         date_object.minute, date_object.second)
                    }
                }
            }, {
                '$project': {
                    'time_stamp': 1,
                    'user': 1,
                    'task': 1,
                    'page_reference': 1
                }
            },
            {
                '$sort': {
                    'time_stamp': -1
                }
            }
        ])

    items_to_view = []
    for item in retrieved_audit_trail:
        items_to_view.append(item)
    return {"audit_trail": items_to_view}
