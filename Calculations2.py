############################# LIBRARIES ################################
from bselib.bse import BSE
import numpy as np
import yfinance as yf
import pandas as pd
from datetime import date


def write_excel(filename, sheetname, dataframe):
    with pd.ExcelWriter(filename, engine='openpyxl', mode='a') as writer:
        workBook = writer.book
        try:
            workBook.remove(workBook[sheetname])
        except:
            print("Worksheet does not exist")
        finally:
            dataframe.to_excel(writer, sheet_name=sheetname, index=False)
            writer.save()


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
input_number = int(input("Enter number based on your choice:- "))

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
    A1.at[index1, 'Dividends'] = new
    if int(row1['Stock Splits']) > 0:
        nn = nn / row1['Stock Splits']

############################# For loop to calculate num of shares ####################
no_stock = 1
for index44, row44 in A1.iterrows():
    split = row44['Stock Splits']
    if split > 0:
        no_stock = no_stock + (split - 1) * no_stock
    else:
        no_stock = no_stock
    A1.at[index44, 'Num Stock'] = no_stock

ini_div = 0
for index55, row55 in A1.iterrows():
    div = row55['Dividends']
    num_stock_2 = row55['Num Stock']
    total_div = div * num_stock_2
    A1.at[index55, 'Dividends'] = total_div
    ini_div = ini_div + total_div

####################### Calculation for finding interest ########################

# Initial value of investment
P = initial_stock_price * A1.loc[first_index].at["Num Stock"]

# Final value of investment
A = end_stock_price * A1.loc[last_index].at["Num Stock"]

# No. of times compounded annually
n = 1

# No. of years for investment
today = date.today()
current_year = today.year
t = (current_year - int(year))

# Interest
r = (n * (((A + ini_div) / P) ** (1 / (n * t)) - 1)) * 100

print("The interest rate is " + str(r))

##A1.at[last_index, "Final Interest rate"] = r

final_interest_rate = pd.DataFrame([r], columns=["Final Interest Rate"])

a = 1

if a == 1:
    # Create a Pandas Excel writer using XlsxWriter engine
    writer = pd.ExcelWriter(Filename, engine='openpyxl', mode='a', if_sheet_exists='replace')

    # Convert the bonus history data frame to an XlsxWriter Excel object
    final_interest_rate.to_excel(writer, sheet_name='Final Interest Rate', index=False)

    # Close the Pandas Excel writer and output the Excel file
    writer.save()

############################# Not in Area ###################################
# august_old = stock_price_history[(stock_price_history['timestamp'] == '2014-08-28')]
# close = august_old['close']

##tt = stock_price_history.to_numpy()
