############################# LIBRARIES ################################
from datetime import date
import numpy as np
import yfinance as yf
import pandas as pd


BSE_SENSEX = pd.read_excel('Equity_active.xlsx')

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
    data_his = Stock_yf.history(period="max")

    if data_his.empty:
        print('The file is empty!')
    else:
        AA1 = data_his[['Dividends', 'Stock Splits']]

        down_data = yf.download(Stock + ".BO", period="max")

        frames = [down_data, AA1]
        Stock_data = pd.concat(frames, axis=1, join='inner')

        # Add Number of Stocks
        strArr = np.empty([len(Stock_data.index)], dtype=int)
        for i in range(0, len(Stock_data.index) - 1):
            tt = 1000
            strArr[i] = tt
        Num_stocks = pd.DataFrame(strArr, columns=["Num Stocks"])
        Stock_data["Num Stock"] = strArr

        # Save data in Excel file
        # Create file name based on stock name
        rr = BSE_SENSEX.loc[input_number].at["Security Name"]
        Filename = BSE_SENSEX.loc[input_number].at["Security Name"] + ".xlsx"
        FilePathName = r"/Users/ruchitbhikadiya/Desktop/stock_project/Excel_Files_Stocks/" + Filename

        # Transfer Stock price history to excel file
        Stock_data.to_excel(FilePathName, sheet_name='Share Price History')

        # Create a Pandas Excel writer using XlsxWriter engine
        writer = pd.ExcelWriter(FilePathName, engine='openpyxl', mode='a')

        # Close the Pandas Excel writer and output the Excel file
        writer.save()

        print("The data is stored in Excel file")

else:
    print("File is already saved")
