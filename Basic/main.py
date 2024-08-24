# import essential libraries
import os
from fyers_apiv3 import fyersModel
import time, json, sys
import xlwings as xw
from flask_app import getauthToken

# global variables
def get_login_credentials():
    global login_credential

    def login_credentials():
        # get login credentials
        print("--- Enter your Fyers API credentials ---")
        login_credential = {
            "api_key": str(input("Enter your API key: ")),
            "api_secret": str(input("Enter your API secret: ")),
            "redirect_url": "http://127.0.0.1:5000"
        }

        # save login credentials
        with open("fyers_credentials.json", "w") as file:
            json.dump(login_credential, file)
            print("Credentials saved successfully")

    # check if credentials file exists
    while True:
        try:
            with open("fyers_credentials.json", "r") as file:
                login_credential = json.load(file)
            break
        except:
            print("Credentials file not found. Please enter your credentials.")
            login_credentials()
    return login_credential

# function to get access token
def get_access_token():
    global access_token ,login_credential
    try:
        with open("fyers_access_token.json", "r") as file:
            access_token = json.load(file)
    except:
        print("Access Token not found. Generating new Access Token")
        access_token = getauthToken(login_credential["api_key"], login_credential["redirect_url"])
        with open("fyers_access_token.json", "w") as file:
            json.dump(access_token, file)
        print("Access Token Generated Successfully")


# function to get fyers object
def get_fyers():
    global fyers, login_credential, access_token
    try:
        fyers = fyersModel.FyersModel(client_id=login_credential["api_key"], token=access_token["access_token"])
    except Exception as e:
        try:
            print("Removing Access Token file, Please Generate new Access Token")
            os.remove("fyers_access_token.json")
            sys.exit()
        except:
            print("Access Token file not found, Please Generate new Access Token")
            sys.exit()

# function to get live data of instruments
def get_live_data(instruments):
    global fyers, live_data
    try:
        live_data
    except:
        live_data = {}

    try:
        symbols = []
        for instrument in instruments:
            symbols.append(instrument)
        data = {
            "symbols": ','.join(symbols),
        }
        response = fyers.quotes(data=data)
        new_data = {}
        for i in response["d"]:
            new_data[i["n"]] = i["v"]
        live_data = new_data
    except Exception as e:
        pass

    # return response
    return response

# function to place order
def place_order(data):
    global fyers
    try:
        response = fyers.place_order(data=data)
        return response
    except Exception as e:
        pass

# function to get order book
def get_order_book():
    global orders

    try:
        data = fyers.orderbook()["orderBook"]
    except Exception as e:
        pass

    return data

# function to start excel
def start_excel():
    global fyers, live_data, orders
    print("Excel Starting...")

    # check if excel file exists
    if not os.path.exists("Excel_Algo_Python.xlsx"):
        try:
            wb = xw.Book()
            wb.save("Excel_Algo_Python.xlsx")
            wb.close()
        except Exception as e:
            sys.exit()
    
    # make sheets
    wb = xw.Book("Excel_Algo_Python.xlsx")
    for i in ["Settings","OrderBook","Data"]:
        try:
            wb.sheets(i)
        except:
            wb.sheets.add(i)
            wb.sheets[i].autofit()
        
    try:
        wb.sheets["Sheet1"].delete()
    except:
        pass
    
    # get sheets
    st = wb.sheets("Settings")
    dt = wb.sheets("Data")
    ob = wb.sheets("OrderBook")

    # configure settings sheet
    st.range("a1:b1").value = ["Settings", "Value"]
    st.range("a1:b1").api.Font.Bold = True
    st.range("a1:b1").api.Font.Size = 12
    st.range("a1:b1").api.WrapText = False
    st.range("a1:b1").api.Orientation = 0
    st.range("a1:b1").api.HorizontalAlignment = -4108
    st.range("a1:b1").api.VerticalAlignment = -4108
    st.range("a1:b1").api.ShrinkToFit = False
    st.range("a1:b1").api.ColumnWidth = 15
    st.range("a1:b1").color = (142, 169, 219)
    st.range("a2").value = "Number of Rows"
    st.range("a3").value = "Time Delay"
    if st.range("b2").value is None:
        st.range("b2").value = 10

    if st.range("b3").value is None:    
        st.range("b3").value = 1
    number_of_rows = int(st.range("b2").value)
    time_delay = int(st.range("b3").value)

    # configure data sheet
    dt.range("a1:p1").value = ["Sr No", "Symbol", "LTP", "Qty","Order Type", "Side", "Product Type", "Limit Price", "Stop Price", "Disclosed Qty", "Validity", "Offline Order", "Stop Loss", "Take Profit", "Entry Signal", "Exit Signal"]
    dt.range(f"a1:p{number_of_rows+1}").api.Font.Bold = True
    dt.range(f"a1:p1").api.Font.Size = 12
    dt.range(f"a1:p1").api.WrapText = False
    dt.range(f"a1:p{number_of_rows+1}").api.Orientation = 0
    dt.range(f"a1:p{number_of_rows+1}").api.HorizontalAlignment = -4108
    dt.range(f"a1:p{number_of_rows+1}").api.VerticalAlignment = -4108
    dt.range(f"a1:p{number_of_rows+1}").api.ShrinkToFit = False
    dt.range(f"a1:p{number_of_rows+1}").api.ColumnWidth = 15
    dt.range(f"b1").api.ColumnWidth = 35
    dt.range(f"a1:p1").color = (142, 169, 219)

    # configure order book sheet
    ob.range("a1:n1").value = ["Order Date Time", "Symbol", "Exchange", "Type", "Qty", "Traded Price", "Side", "Product Type", "Limit Price", "Disclosed Qty", "Validity", "Offline Order", "Stop Price", "Order Status"]
    ob.range("a1:n1").api.Font.Bold = True
    ob.range("a1:n1").api.Font.Size = 12
    ob.range("a1:n1").api.WrapText = False
    ob.range("a1:n1").api.Orientation = 0
    ob.range("a1:n1").api.HorizontalAlignment = -4108
    ob.range("a1:n1").api.VerticalAlignment = -4108
    ob.range("a1:n1").api.ShrinkToFit = False
    ob.range("a1:n1").api.ColumnWidth = 15
    ob.range("a1:n1").color = (142, 169, 219)
    ob.range("n1").column_width = 150

    try:
        ob.activate()
        ob.api.Application.ActiveWindow.SplitColumn = 2
        ob.api.Application.ActiveWindow.SplitRow = 0
        ob.api.Application.ActiveWindow.FreezePanes = True

        dt.activate()
        dt.api.Application.ActiveWindow.SplitColumn = 2
        dt.api.Application.ActiveWindow.SplitRow = 0
        dt.api.Application.ActiveWindow.FreezePanes = True
        print("Excel Started Successfully")
    except Exception as e:
        pass
    
    # update data sheet
    dt.range(f"a2:a{number_of_rows+1}").value = [[i] for i in range(1, number_of_rows+1)]
    order_type_values = {
        "LIMIT": 1,
        "MARKET": 2,
        "SL": 3,
        "SL-M": 4
    }
    side_values = {
        "BUY": 1,
        "SELL": -1
    }

    subs_lst = []
    while True:
        try:
            time.sleep(time_delay)
            # update order book
            orders = get_order_book()
            
            # put all orders in order book
            for i in orders:
                try:
                    index_no = orders.index(i)
                    ob.range(f"a{index_no+2}").value = i["orderDateTime"]
                    ob.range(f"b{index_no+2}").value = i["symbol"]
                    ob.range(f"c{index_no+2}").value = i["exchange"]
                    ob.range(f"d{index_no+2}").value = [k for k,v in order_type_values.items() if v == i["type"]][0]
                    ob.range(f"e{index_no+2}").value = i["qty"]
                    ob.range(f"f{index_no+2}").value = i["tradedPrice"]
                    ob.range(f"g{index_no+2}").value = [k for k,v in side_values.items() if v == i["side"]][0]
                    ob.range(f"h{index_no+2}").value = i["productType"]
                    ob.range(f"i{index_no+2}").value = i["limitPrice"]
                    ob.range(f"j{index_no+2}").value = i["disclosedQty"]
                    ob.range(f"k{index_no+2}").value = i["orderValidity"]
                    ob.range(f"l{index_no+2}").value = i["offlineOrder"]
                    ob.range(f"m{index_no+2}").value = i["stopPrice"]
                    ob.range(f"n{index_no+2}").value = i["message"]

                except Exception as e:
                    pass


            if subs_lst:
                get_live_data(subs_lst)
            symbols = [symbol for symbol in dt.range(f"b{2}:b{number_of_rows+1}").value if symbol]
            for i in subs_lst:
                if i not in symbols:
                    subs_lst.remove(i)
                    try:
                        del live_data[i]
                    except Exception as e:
                        pass

            for i in symbols:
                if i not in subs_lst:
                    subs_lst.append(i)
                    try:
                        del live_data[i]
                    except Exception as e:
                        pass
                
                try:
                    idx = symbols.index(i)
                    dt.range(f"c{idx+2}").value = live_data[i]["lp"]
                    # add drop down for order type
                    if dt.range(f"e{idx+2}").value is None:
                        order_type = ["MARKET", "LIMIT", "SL", "SL-M"]
                        side = ["BUY", "SELL"]
                        product_type = ["CNC", "INTRADAY", "MARGIN", "CO", "BO"]
                        validity = ["DAY", "IOC"]
                        offline_order = ["TRUE", "FALSE"]
                        entry_signal = ["TRUE", "FALSE"]
                        exit_signal = ["TRUE", "FALSE"]
                        dt.range(f"e{idx+2}").api.Validation.Add(3, 1, 1, ",".join(order_type))
                        dt.range(f"f{idx+2}").api.Validation.Add(3, 1, 1, ",".join(side))
                        dt.range(f"g{idx+2}").api.Validation.Add(3, 1, 1, ",".join(product_type))
                        dt.range(f"k{idx+2}").api.Validation.Add(3, 1, 1, ",".join(validity))
                        dt.range(f"l{idx+2}").api.Validation.Add(3, 1, 1, ",".join(offline_order))
                        dt.range(f"o{idx+2}").api.Validation.Add(3, 1, 1, ",".join(entry_signal))
                        dt.range(f"p{idx+2}").api.Validation.Add(3, 1, 1, ",".join(exit_signal))
                        # insert default values
                        dt.range(f"e{idx+2}").value = "MARKET"
                        dt.range(f"f{idx+2}").value = "BUY"
                        dt.range(f"g{idx+2}").value = "CNC"
                        dt.range(f"k{idx+2}").value = "DAY"
                        dt.range(f"l{idx+2}").value = "FALSE"
                        dt.range(f"o{idx+2}").value = "FALSE"
                        dt.range(f"p{idx+2}").value = "FALSE"
                        dt.range(f"d{idx+2}").value = 1
                        dt.range(f"h{idx+2}").value = 0
                        dt.range(f"i{idx+2}").value = 0
                        dt.range(f"j{idx+2}").value = 0
                        dt.range(f"m{idx+2}").value = 0
                        dt.range(f"n{idx+2}").value = 0                

                    # PLACE ORDER
                    if dt.range(f"o{idx+2}").value == True:
                        try:
                            data = {
                                "symbol": str(dt.range(f"b{idx+2}").value),
                                "qty": int(dt.range(f"d{idx+2}").value),
                                "type": int(order_type_values[dt.range(f"e{idx+2}").value]),
                                "side": int(side_values[dt.range(f"f{idx+2}").value]),
                                "productType": str(dt.range(f"g{idx+2}").value),
                                "limitPrice": float(dt.range(f"h{idx+2}").value),
                                "stopPrice": float(dt.range(f"i{idx+2}").value),
                                "disclosedQty": int(dt.range(f"j{idx+2}").value),
                                "validity": str(dt.range(f"k{idx+2}").value),
                                "offlineOrder": bool(dt.range(f"l{idx+2}").value),
                                "stopLoss": float(dt.range(f"m{idx+2}").value),
                                "takeProfit": float(dt.range(f"n{idx+2}").value),
                                "orderTag": "ExcelOrder"
                            }
                            place_order(data)
                            dt.range(f"o{idx+2}").value = "FALSE"
                        except Exception as e:
                            pass

                    # check exit signal
                    if dt.range(f"p{idx+2}").value == True:
                        try:
                            data = {
                                "symbol": str(dt.range(f"b{idx+2}").value),
                                "qty": int(dt.range(f"d{idx+2}").value),
                                "type": int(order_type_values[dt.range(f"e{idx+2}").value]),
                                "side": int(side_values[dt.range(f"f{idx+2}").value])*(-1),
                                "productType": str(dt.range(f"g{idx+2}").value),
                                "limitPrice": float(dt.range(f"h{idx+2}").value),
                                "stopPrice": float(dt.range(f"i{idx+2}").value),
                                "disclosedQty": int(dt.range(f"j{idx+2}").value),
                                "validity": str(dt.range(f"k{idx+2}").value),
                                "offlineOrder": bool(dt.range(f"l{idx+2}").value),
                                "stopLoss": float(dt.range(f"m{idx+2}").value),
                                "takeProfit": float(dt.range(f"n{idx+2}").value),
                                "orderTag": "ExcelOrder"
                            }
                            place_order(data)
                            dt.range(f"p{idx+2}").value = "FALSE"
                        except Exception as e:
                            pass
                except Exception as e:
                    pass
        except Exception as e:
            pass
                

            
                        

if __name__ == "__main__":
    get_login_credentials()
    get_access_token()
    get_fyers()
    start_excel()
   