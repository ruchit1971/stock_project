# # Import Librabries
# import numpy as np
# import pandas
#
# excel_data_df = pandas.read_excel('Info.xlsx', sheet_name='One',header=None,converters={'Date':str,'Close':str})
#
# tt = excel_data_df.to_numpy()
# # print whole sheet data
# print(tt)


# Import the yfinance. If you get module not found error the run !pip install yfinance from your Jupyter notebook
import yfinance as yf
import pandas as pd

# Sensex website
Sensex = pd.read_html('https://en.wikipedia.org/wiki/BSE_SENSEX')[1].Symbol.to_list()

# Download the whole data according to wikipedia stocks list
leman = yf.download(Sensex[0])

# Set new stock
stock = yf.Ticker("INFY.BO")
data = yf.download("INFY.BO", period = "10y")
data_1 = []

# Initiate index value
i = 0

divi = stock.dividends
bonus = stock.isin
# for i in data:
#     data_1[i] = data[]


print(bonus)




