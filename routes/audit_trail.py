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
    return f"new trail item inserted: Id = {posted.inserted_id}"


# UPDATE INVESTOR CHANGES - DELETE TO FOLLOW (CURRENTLY ONLY DEBBIE AND I CAN DELETE AN INVESTOR)
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
            task += f"{item} = {ddiff['values_changed'][item]};"
    if ddiff.get('iterable_item_added') is not None:
        for item in ddiff['iterable_item_added']:
            task += f"{item} = {ddiff['iterable_item_added'][item]};"
    if ddiff.get('iterable_item_removed') is not None:
        for item in ddiff['iterable_item_removed']:
            task += f"{item} = {ddiff['iterable_item_removed'][item]};"

    if task:
        # REPLACE TEXT WITH TEXT MEANINGFUL TO THE USE
        task = task.replace('investments', 'released')
        task = task.replace('trust', 'investment')
        task = task.replace('root', '')

        task_list = task.split(";")
        # REMOVE AUTOMATIC TASKS THAT THE USER DOES NOT DO HIMSELF
        filtered_list = [x for x in task_list if "released" not in x and ("interest" not in x or "exit_value" not in x)]
        filtered_list = [x for x in filtered_list if "investment" not in x and ("interest" not in x or "available_date"
                                                                                not in x)]
        task = ""
        for item in filtered_list:
            if item != '':
                item += ';'
                task += item
        print("filtered_list",filtered_list)
        print()
        print("task", task)
        dictionary = {
            "time_stamp": request['time_stamp'],
            "user": request['user'],
            "page_reference": request['page_reference'],
            "task": task
        }
        posted = audit_trail.insert_one(dictionary)
        return f"new trail item inserted: Id = {posted.inserted_id}"

    else:
        return "No Changes made"


# GET ALL USERS THAT HAVE DONE STUFF ON APP - THIS IS FOR FIRST SELECT ON AUDIT PAGE (THIS IS TO CHOOSE THE USER)
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
    date_object = datetime.strptime(request["period"], "%Y/%m/%d %H:%M:%S")
    if request["user"] != "All Users":
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
