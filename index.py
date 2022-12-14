import os
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from routes.investors import investor
from routes.excel_sales_forecast import excel_sales_forecast
from routes.sales import sales

if not os.path.isdir("sales_documents"):
    os.makedirs("sales_documents")

if not os.path.isdir("loan_agreements"):
    os.makedirs("loan_agreements")

if not os.path.isdir("split_pdf_files"):
    os.makedirs("split_pdf_files")

app = FastAPI()

origins = ["*"]

app.add_middleware(
    CORSMiddleware,
    allow_origins=origins,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

app.include_router(investor)
app.include_router(excel_sales_forecast)
app.include_router(sales)
