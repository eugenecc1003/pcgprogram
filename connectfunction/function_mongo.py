import pymongo
from pymongo.mongo_client import MongoClient
from bson.objectid import ObjectId

from secretstoken import *

uri = MONGODB_API
client = MongoClient(uri)
db = client.member_system
collection = db.user


def check():
    result = ""
    data = collection.find({'status': "check"})
    for doc in data:
        result = result+"["+doc['email']+"]\n"
    amount = data.explain().get("executionStats", {}).get("nReturned")
    return "No user need to be approved !!" if amount == 0 else result + "wait to check !!"


def emaillist():
    emaillist = []
    data = collection.find({})
    for doc in data:
        emaillist.append(doc['email'])
    return emaillist


def approve(email):
    collection.update_one({'email': email}, {'$set': {'status': 'approve'}})
    return "["+email+"]\nhas been approved !!"


def authorization_insert(email, area):
    # find it and authorization
    data = collection.find_one({'$and': [{'email': email},]})
    authorization = data['authorization']
    # insert
    authorization.append(str(area))
    authorization = list(set(authorization))
    authorization.sort()
    collection.update_one(
        {'email': email}, {'$set': {'authorization': authorization}})
    # check again
    data1 = collection.find_one({'$and': [{'email': email},]})
    return "["+email+"]\nauthorization is " + str(data1['authorization'])