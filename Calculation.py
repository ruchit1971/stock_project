############################# LIBRARIES ################################
from bselib.bse import BSE
import numpy as np
import yfinance as yf
import pandas as pd
from datetime import date

# # BSE Sensex Stocks list
# BSE_SENSEX = pd.read_html('https://en.wikipedia.org/wiki/BSE_SENSEX')[1]
#
# # Remove Unnecessary Columns
# del BSE_SENSEX["Date Added"]
# del BSE_SENSEX["#"]
#
# # Print companies names
# print(BSE_SENSEX['Companies'])
#
# # Enter number based on your choice
# input_number = 13 #int(input("Enter number based on your choice:- "))


BSE_SENSEX = pd.read_excel('Equity.xlsx')

# Remove Unnecessary Columns
del BSE_SENSEX["Group"]
del BSE_SENSEX["Face Value"]
del BSE_SENSEX["ISIN No"]
del BSE_SENSEX["Industry"]
del BSE_SENSEX["Instrument"]

# Print companies names
print(BSE_SENSEX['Security Name'])

# Enter number based on your choice
input_number = 192 #int(input("Enter number based on your choice:- "))

# Consider Stock symbol based on Input
Stock = BSE_SENSEX.loc[input_number].at["Security Id"]

# Selected stock name
print("You have selected " + BSE_SENSEX.loc[input_number].at["Security Name"] + " Stock")

# Create file name based on stock name
Filename = BSE_SENSEX.loc[input_number].at["Security Name"] + ".xlsx"

# Read Excel File according to stock name and create data frame
stock_price_history = pd.read_excel(Filename, sheet_name='Share_price_history')
bonus_history = pd.read_excel(Filename, sheet_name='Bonus_history')

# Enter Year
year = '2011'

# Extract Stock price history based on year
A1 = stock_price_history[(stock_price_history['Date'] > year)]
A2 = A1[['Date', 'Close']]

# Find the index values for initial and end stock price
first_index = A2.index.values[1]
last_index = A2.index.values[-1]
total_rows = len(A2.index)

# Based on Index values, Extract initial and end stock price
initial_stock_price = A2.loc[first_index].at["Close"]
end_stock_price = A2.loc[last_index].at["Close"]

# Extract Bonus history based on year
B1 = bonus_history[(bonus_history['Date'] > year)]

# Extract Dividend info
D1 = stock_price_history[(stock_price_history['Date'] > year)]
D2 = D1[['Date', 'Dividends', 'Stock Splits']]
D3 = D1[['Date', 'Dividends']]
Splits_nonzero = D2[(D2['Stock Splits'] > 0)]

nn = 1
# Loop for Dividends
for index, row in Splits_nonzero.iterrows():
    nn = nn * row['Stock Splits']

# Update Dividends
for index1, row1 in D2.iterrows():
    old = row1['Dividends']
    new = old * nn
    D2.at[index1,'Dividends'] = new
    if int(row1['Stock Splits']) > 0:
        nn = nn/row1['Stock Splits']

# Enter Number of Stock
initial_stock = 1000
num_stock = 1000
div_sum  = 0

# Loop to calculate end number of stock
for index, row in B1.iterrows():
    Datee = row['Date']
    bonus_date = row['Date']
    Div1 = D2[(D2['Date'] < bonus_date)]
    DDD = Div1.to_numpy()
    DDD2 = DDD[:, 1]
    Divident_total = np.sum(DDD2)
    div_sum = num_stock * Divident_total
    Bonus_on_year = row['Bonus']
    num_stock = num_stock + num_stock*int(Bonus_on_year)


####################### Calculation for finding interest ########################

# Initial value of investment
P = initial_stock_price * initial_stock

# Final value of investment
A = end_stock_price * num_stock

# No. of times compounded annually
n = 1

# No. of years for investment
today = date.today()
current_year = today.year
t = (current_year - int(year))

# Interest
r = (n * ((A/P)**(1/(n*t)) - 1)) * 100

print("The interest rate is " + str(r))














############################# Not in Area ###################################
#august_old = stock_price_history[(stock_price_history['timestamp'] == '2014-08-28')]
#close = august_old['close']

tt = stock_price_history.to_numpy()