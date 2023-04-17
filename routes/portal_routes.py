import copy
import os
from fastapi import APIRouter, Request
from fastapi.responses import FileResponse
from portal_statement_files.portal_statement_create import create_pdf

import time

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
