# Load Libraries
from bselib.bse import BSE
import numpy as np
import yfinance as yf
import pandas as pd
import csv
import requests
from openpyxl import load_workbook


#
# # ######################## FUNCTIONS ####################
# def getList(dict):
#     list = []
#     for key in dict.keys():
#         list.append(key)
#
#     return list
#
#
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
# Date = []
# for index in data["bonus"]["data"]:
#     Bonus.append(index[2])
#     Date.append(index[1])
#     # Bonus = Bonus.astype(np.int64)
#
# ee = []
# for i in range(0, len(Bonus)):
#     x = Bonus[i].split(":")
#     ee.append(int(x[1]) / int(x[0]))
#
# dd = pd.DataFrame(list(zip(Date, ee)),
#                   columns=['Date', 'Bonus'])
#
# bonus = pd.DataFrame(Bonus)
#
# read_file = pd.read_csv (r'downloaded.csv')
# read_file.to_excel (r'eew.xlsx', index = None, header=True)
#
#
# dd.to_csv('downloaded.csv', mode='a', header=False)
#
# print(Bonus)
#
#
#
#
# # Sensex website
# Sensex = pd.read_html('https://en.wikipedia.org/wiki/BSE_SENSEX')[1].Symbol.to_list()
#
# # Download the whole data according to wikipedia stocks list
# leman = yf.download(Sensex[0])
#
# # Set new stock
# stock = yf.Ticker("INFY.BO")
# data = yf.download("INFY.BO", period = "10y", interval = "3mo")
# data_1 = []
# stock = stock.history(period="10y")
# # Initiate index value
# i = 0
#
# ages = stock["Dividends"].tolist()
#
# #divi = stock.history
# #bonus = stock.dividends
# # for i in data:
# #     data_1[i] = data[]
#
#
# print(stock)


#
# #
# # # replace the "demo" apikey below with your own key from https://www.alphavantage.co/support/#api-key
# # #url = 'https://www.alphavantage.co/query?function=TIME_SERIES_MONTHLY&symbol=INFY.BSE&apikey=ZV2WA0TR3XRJLDW9'
# # url = 'https://www.alphavantage.co/query?function=EARNINGS&symbol=RELIANCE.BSE&apikey=ZV2WA0TR3XRJLDW9'
# # r = requests.get(url)
# # data = r.json()
# #
# # print(data)
#
# csv_url = 'https://www.alphavantage.co/query?function=TIME_SERIES_MONTHLY&symbol=INFY.BSE&apikey=ZV2WA0TR3XRJLDW9&datatype=csv'
#
# print(csv_url)
#
# req = requests.get(csv_url)
# url_content = req.content
# csv_file = open('downloaded.csv', 'wb')
#
# csv_file.write(url_content)
# csv_file.close()


# msft = yf.Ticker("INFY.BO")
# # hist = msft.history(period="max", interval = "1m" )
# data = yf.download("INFY.BO", period = "10y", interval="1mo")
#

# import os
#
# os.path.isfile('./Infosys.xlsx')
from datetime import date

today = date.today()
print("Today's date:", today)

old_date = 0

a = 0

if a:
    print("Save file")
else:
    print("Dont save file")




