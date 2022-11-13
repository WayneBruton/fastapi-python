from pymongo import MongoClient
from decouple import config

conn = MongoClient(config("MONGO_CREDENTIALS"))

omh = config("DB_LIVE")

db = conn.omh

