from pymongo import MongoClient
from decouple import config

conn = MongoClient(config("MONGO_CREDENTIALS"))

db = conn.config("DB_LIVE")

