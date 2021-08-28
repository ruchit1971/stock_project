from bselib.bse import BSE
import numpy as np
import yfinance as yf
import pandas as pd

# # BSE Sensex Stocks list
# BSE_SENSEX = pd.read_html('https://www.bseindia.com/stock-share-price/bajaj-finance-limited/bajfinance/500034/corp-actions/')[1]['Date']
#
# c  = 1

msft = yf.Ticker("HCLTECH.BO")

# get stock info
e = msft.info

# get historical market data
hist = msft.history(period="max")

# show actions (dividends, splits)
dd = msft.actions

# show dividends
ee = msft.dividends

# show splits
rr = msft.splits


