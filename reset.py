import os

def reset():
    print("Warning: This will reset the auth code and access token files")
    try:
        os.remove("fyers_auth_code.json")
    except:
        pass

    try:
       os.remove("fyers_access_token.json")
    except:
        pass
    ask = input("Do you want to reset the credentials file also? (y/n): ")
    if ask.lower() == "y":
        try:
            os.remove("fyers_credentials.json")
        except:
            pass
        print("Credentials file reset successful")
    else:
        print("Credentials file reset cancelled")
    print("Reset successful")

reset()