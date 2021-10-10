############################# LIBRARIES ################################
from datetime import date
import numpy as np
import yfinance as yf
import pandas as pd


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

is_file = 1

if is_file:
    # Download selected stock history
    Stock_yf = yf.Ticker(Stock + ".BO")
    Stock_data = Stock_yf.history(period="max")

    # Add Number of Stocks
    strArr = np.empty([len(Stock_data.index)], dtype=int)
    for i in range(0, len(Stock_data.index) - 1):
        tt = 1000
        strArr[i] = tt
    Num_stocks = pd.DataFrame(strArr, columns=["Num Stocks"])
    Stock_data["Num Stock"] = strArr

    # Save data in Excel file
    # Create file name based on stock name
    Filename = BSE_SENSEX.loc[input_number].at["Security Name"] + ".xlsx"

    # Transfer Stock price history to excel file
    Stock_data.to_excel(Filename, sheet_name='Share Price History')

    # Create a Pandas Excel writer using XlsxWriter engine
    writer = pd.ExcelWriter(Filename, engine='openpyxl', mode='a')

    # Close the Pandas Excel writer and output the Excel file
    writer.save()

    print("The data is stored in Excel file")

else:
    print("File is already saved")
