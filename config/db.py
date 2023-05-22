from pymongo import MongoClient
from decouple import config
import motor.motor_asyncio

# conn = motor.motor_asyncio.AsyncIOMotorClient(config("MONGO_CREDENTIALS"))



conn = MongoClient(config("MONGO_CREDENTIALS"))

OMH_DB = config("OMH_DB")

db = conn[OMH_DB]


