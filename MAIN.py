############################# LIBRARIES ################################
from bselib.bse import BSE
import numpy as np
import yfinance as yf
import pandas as pd



# BSE SENSEX STOCKS
BSE_SENSEX = pd.read_html('https://en.wikipedia.org/wiki/BSE_SENSEX')[1].to_numpy()

x = str(input())