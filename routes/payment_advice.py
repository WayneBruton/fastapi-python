# import os
# import boto3
# from botocore.exceptions import ClientError
from fastapi import APIRouter, Request
# from fastapi.responses import FileResponse
from config.db import db
from bson.objectid import ObjectId

# from decouple import config

advices = APIRouter()

# MONGO COLLECTIONS
constructionValuations = db.constructionValuations
constructionValuationsNonProgress = db.constructionValuationsNonProgress


@advices.post("/updateValuationStatus")
async def update_valuation_status(data: Request):
    request = await data.json()
    # print(request)

    try:
        results = list(constructionValuations.aggregate(
            [
                {
                    "$match": {
                        'development': request['development'],
                        'subcontractor': request['subcontractor'],
                    },
                },
                {'$unwind': {'path': "$tasks"}},
                {
                    '$match': {"tasks.paymentAdviceNumber": request['paymentAdviceNumber']},
                },
                {
                    '$set': {
                        'paymentAdviceNumber': "$tasks.paymentAdviceNumber",
                        'currentProgress': "$tasks.currentProgress",
                        'initialProgress': "$tasks.initialProgress",
                        'approved': "$tasks.approved",
                        'status': "$tasks.status",
                    },
                },
                {
                    '$project': {
                        'paymentAdviceNumber': 1,
                        'currentProgress': 1,
                        'initialProgress': 1,
                        'approved': 1,
                        'status': 1,
                    },
                },
            ]
        ))
    except Exception as e:
        print(e)
        results = []
    # print(results)
    for result in results:
        result['_id'] = str(result['_id'])
        result['id'] = result['_id']
        del result['_id']
        if request['status'] == 'Approved':
            result['approved'] = True
        else:
            result['approved'] = False

        try:
            # DELETE EXISTING TASK
            deleted_tasks = constructionValuations.update_one(
                {"_id": ObjectId(request['id'])}, {"$pop": {"tasks": 1}
                                                   })
            # INSERT TASK WITH LATEST STUFF
            inserted_tasks = constructionValuations.update_one(
                {"_id": ObjectId(request['id'])}, {"$push": {"tasks": result}
                                                   })
        except Exception as e:
            print(e)

        print("results", result)

    # return "Done"
    return {"result": result}
