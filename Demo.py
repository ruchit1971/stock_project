# Load Libraries
from bselib.bse import BSE
import numpy as np
import yfinance as yf
import pandas as pd

######################## FUNCTIONS ####################
def getList(dict):
    list = []
    for key in dict.keys():
        list.append(key)

    return list




# # Get BSE Data...
# b = BSE()
#
# # Find Stocks Name...
# print("Enter Stock Name: ")
# x = str(input())
#
# print(x)
# # 500325
# stocks = b.script(x)
# print(stocks)
#
# print("Enter Option Number: ")
#
# num = int(input())
#
# rr = getList(stocks)
#
# data = b.corporate_actions(rr[num - 1])
#
# Bonus = []
# for index in data["bonus"]["data"]:
#     Bonus.append(index[2])
#     #Bonus = Bonus.astype(np.int64)
# print(Bonus)




# Sensex website
Sensex = pd.read_html('https://en.wikipedia.org/wiki/BSE_SENSEX')[1].Symbol.to_list()

# Download the whole data according to wikipedia stocks list
leman = yf.download(Sensex[0])

# Set new stock
stock = yf.Ticker("INFY.BO")
data = yf.download("INFY.BO", period = "10y")
data_1 = []
stock = stock.history(period="10y")
# Initiate index value
i = 0

#divi = stock.history
#bonus = stock.dividends
# for i in data:
#     data_1[i] = data[]


print(stock)