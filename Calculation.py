############################# LIBRARIES ################################
from bselib.bse import BSE
import numpy as np
import yfinance as yf
import pandas as pd

# BSE Sensex Stocks list
BSE_SENSEX = pd.read_html('https://en.wikipedia.org/wiki/BSE_SENSEX')[1]

# Remove Unnecessary Columns
del BSE_SENSEX["Date Added"]
del BSE_SENSEX["#"]

# Print companies names
print(BSE_SENSEX['Companies'])

# Enter number based on your choice
input_number = int(input("Enter number based on your choice:- "))

# Consider Stock symbol based on Input
Stock = BSE_SENSEX.loc[input_number].at["Symbol"]

# Selected stock name
print("You have selected " + BSE_SENSEX.loc[input_number].at["Companies"] + " Stock")

# Create file name based on stock name
Filename = BSE_SENSEX.loc[input_number].at["Companies"] + ".xlsx"

# Read Excel File according to stock name and create data frame
stock_price_history = pd.read_excel(Filename, sheet_name='Share_price_history')
bonus_history = pd.read_excel(Filename, sheet_name='Bonus_history')

# Enter Year
year = '2014'

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









############################# Not in Area ###################################
#august_old = stock_price_history[(stock_price_history['timestamp'] == '2014-08-28')]
#close = august_old['close']

tt = stock_price_history.to_numpy()