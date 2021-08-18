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

is_file = 0

if is_file:
    # Download selected stock history
    Stock_yf = yf.Ticker(Stock)
    Stock_data = Stock_yf.history(period="10y")

    # Download Bonus history
    b = BSE()
    corp_action_data = b.corporate_actions(BSE_SENSEX.loc[input_number].at["Exchange ticker"])

    Bonus = []
    Date = []
    for index in corp_action_data["bonus"]["data"]:
        Bonus.append(index[2])
        Date.append(index[1])

    ee = []
    rr = []
    for i in range(0, len(Bonus)):
        x = Bonus[i].split(":")
        e = Date[i].split("-")
        rr.append(e[2])
        ee.append(int(x[1]) / int(x[0]))

    Date = pd.DataFrame(Date, columns=["Date"])
    Date_new = pd.to_datetime(Date.Date)
    Bonus = pd.DataFrame (ee, columns=['Bonus'])

    frames = [Date_new, Bonus]
    Bonus_data = pd.concat(frames, axis=1, join='inner')

    # Save data in Excel file
    # Create file name based on stock name
    Filename = BSE_SENSEX.loc[input_number].at["Companies"] + ".xlsx"

    # Transfer Stock price history to excel file
    Stock_data.to_excel(Filename, sheet_name='Share_price_history')

    # Create a Pandas Excel writer using XlsxWriter engine
    writer = pd.ExcelWriter(Filename, engine='openpyxl', mode='a')

    # Convert the bonus history data frame to an XlsxWriter Excel object
    Bonus_data.to_excel(writer, sheet_name='Bonus_history', index=False)

    # Close the Pandas Excel writer and output the Excel file
    writer.save()

else:
    print("File is already saved")

