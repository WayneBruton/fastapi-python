from pymongo import MongoClient
from decouple import config

conn = MongoClient(config("MONGO_CREDENTIALS"))

OMH_DB = config("OMH_DB")

db = conn.OMH_DB

