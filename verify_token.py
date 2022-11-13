import jwt
from decouple import config

key = config("KEY")





def verify_jwt_token(token):
    try:
        result = jwt.decode(token, key, algorithms="HS256")
        return "Verified"
    except:
        return "Verification Failed"
