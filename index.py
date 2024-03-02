import os
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from routes.investors import investor
from routes.excel_sales_forecast import excel_sales_forecast
from routes.sales import sales
from routes.audit_trail import audit
from routes.payment_advice import advices
from routes.portal_routes import portal_info
from routes.investor_data_analysis import data_analysis
from routes.rentals import rentals
from routes.construction import construction
from routes.xero import xero
from routes.lead_generation import leads
from routes.cashflow_routes import cashflow

if not os.path.isdir("sales_documents"):
    os.makedirs("sales_documents")

if not os.path.isdir("loan_agreements"):
    os.makedirs("loan_agreements")

if not os.path.isdir("split_pdf_files"):
    os.makedirs("split_pdf_files")

if not os.path.isdir("excel_files"):
    os.makedirs("excel_files")

if not os.path.isdir("portal_statements"):
    os.makedirs("portal_statements")

if not os.path.isdir("sales_client_onboarding_docs"):
    os.makedirs("sales_client_onboarding_docs")

if not os.path.isdir("payment_advice"):
    os.makedirs("payment_advice")

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
app.include_router(audit)
app.include_router(advices)
app.include_router(portal_info)
app.include_router(data_analysis)
app.include_router(rentals)
app.include_router(construction)
app.include_router(xero)
app.include_router(leads)
app.include_router(cashflow)

