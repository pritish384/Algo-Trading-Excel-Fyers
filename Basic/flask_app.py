from flask import Flask, request
import json
import os, signal
app = Flask(__name__)
from fyers_apiv3 import fyersModel
import webbrowser
import threading


def getauthToken(appId, redirect_uri):
    flask_thread = threading.Thread(target=run_flask_app)
    flask_thread.start()
    functionName = "getauthToken"    
    response_type="code"
    grant_type="authorization_code"
    appSession = fyersModel.SessionModel(client_id=appId,redirect_uri=redirect_uri,response_type=response_type, grant_type=grant_type,state="state",scope="",nonce="")

    generateTokenUrl = appSession.generate_authcode()

    webbrowser.open(generateTokenUrl, new=1)
    flask_thread.join()

def generate_access_token(auth_code, appId, secret_key):
    functionName = "generate_access_token"
    appSession = fyersModel.SessionModel(client_id=appId, secret_key=secret_key,grant_type="authorization_code")
    appSession.set_token(auth_code)
    access_token = appSession.generate_token()
    with open("fyers_access_token.json", "w") as file:
        json.dump(access_token, file)
    print("Access Token Generated Successfully")

@app.route('/')
def index():
    query_params = request.args
    # save code in credential file
    with open("fyers_auth_code.json", "w") as file:
        data = {"auth_code": query_params.get("auth_code")}
        json.dump(data, file)

    with open("fyers_credentials.json", "r") as file:
        login_credential = json.load(file)
        appId = login_credential["api_key"]
        secret_key = login_credential["api_secret"]

    # generate access token
    generate_access_token(data["auth_code"] , appId, secret_key)
    
    os.kill(os.getpid(), signal.SIGINT)
    return "Auth Code Received Successfully"



def run_flask_app():
    app.run(host='127.0.0.1', port=5000, debug=False)

if __name__ == '__main__':
    print("Please Run main.py file to get the auth code")