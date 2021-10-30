import numpy as np
import yfinance as yf
import pandas as pd
from datetime import date
import os.path



BSE_SENSEX = pd.read_excel('Equity_active.xlsx')
df = pd.DataFrame(columns=['Stock Name'])

for i, r in BSE_SENSEX.iterrows():
    # Enter number based on your choice
    input_number = i  # int(input("Enter number based on your choice:- "))

    # Consider Stock symbol based on Input
    Stock = BSE_SENSEX.loc[input_number].at["Security Id"]

    # Selected stock name
    print("You have selected " + BSE_SENSEX.loc[input_number].at["Security Name"] + " Stock")

    # Create file name based on stock name
    Filename = BSE_SENSEX.loc[input_number].at["Security Name"] + ".xlsx"
    FilePathName = r"/Users/ruchitbhikadiya/Desktop/stock_project/Excel_Files_Stocks/" + Filename

    if os.path.isfile(FilePathName):
        try:
            # Read Excel File according to stock name and create data frame
            stock_price_history = pd.read_excel(FilePathName, sheet_name='Share Price History')
            print("The stock file exist")
            df = df.append({'A': BSE_SENSEX.loc[input_number].at["Security Name"]}, ignore_index=True)
        except:
            print("Corrupted")
    else:
        print("The stock file does not exist")




# Transfer Stock price history to excel file
FilePathName2 = r"/Users/ruchitbhikadiya/Desktop/stock_project/Excel_Files_Stocks/0Equity.xlsx"
df.to_excel(FilePathName2, sheet_name='Stock List')

# Create a Pandas Excel writer using XlsxWriter engine
writer = pd.ExcelWriter(FilePathName2, engine='openpyxl', mode='a')

# Close the Pandas Excel writer and output the Excel file
writer.save()
