############################# LIBRARIES ################################
import numpy as np
import yfinance as yf
import pandas as pd
from datetime import date


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
stock_price_history = pd.read_excel(Filename, sheet_name='Share Price History')
bonus_history = pd.read_excel(Filename, sheet_name='Bonus History')

# Check if 10 years is possible or not!
today = date.today()
current_month = today.month
year_from_data = stock_price_history.loc[1].at["Date"]
yu = year_from_data.year


# Enter Year
year = str(current_month) + '-'+'2011'
year_5 = str(current_month) + '-'+'2016'

if yu < int('2011'):
    # Calculate for 10 years
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
    no_stock = int(A1.loc[first_index].at["Num Stock"])
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

    # Interest without Inflation
    r = (n * (((A + ini_div) / P) ** (1 / (n * t)) - 1)) * 100

    # Interest with Inflation
    #r_with_inflation =

    print("The interest rate is " + str(r))

    ################# Calculate for 5 years ##################
    # Calculate for 5 years
    # Extract Stock price history based on year
    A1_5 = stock_price_history[(stock_price_history['Date'] > year_5)]
    A2_5 = A1_5[['Date', 'Close']]

    # Find the index values for initial and end stock price
    first_index_5 = A2_5.index.values[1]
    last_index_5 = A2_5.index.values[-1]
    total_rows_5 = len(A2_5.index)

    # Based on Index values, Extract initial and end stock price
    initial_stock_price_5 = A2_5.loc[first_index_5].at["Close"]
    end_stock_price_5 = A2_5.loc[last_index_5].at["Close"]

    # Extract Bonus history based on year
    B1_5 = bonus_history[(bonus_history['Date'] > year_5)]

    # Extract Dividend info
    D1_5 = stock_price_history[(stock_price_history['Date'] > year_5)]
    D2_5 = D1_5[['Date', 'Dividends', 'Stock Splits']]
    D3_5 = D1_5[['Date', 'Dividends']]
    Splits_nonzero_5 = D2_5[(D2_5['Stock Splits'] > 0)]

    nn_5 = 1
    # Loop for Dividends
    for index_5, row_5 in Splits_nonzero_5.iterrows():
        nn_5 = nn_5 * row_5['Stock Splits']

    # Update Dividends
    for index1_5, row1_5 in D2_5.iterrows():
        old_5 = row1_5['Dividends']
        new_5 = old_5 * nn_5
        A1_5.at[index1_5, 'Dividends'] = new_5
        if int(row1_5['Stock Splits']) > 0:
            nn_5 = nn_5 / row1_5['Stock Splits']

    ############################# For loop to calculate num of shares ####################
    no_stock_5 = int(A1_5.loc[first_index_5].at["Num Stock"])
    for index44_5, row44_5 in A1_5.iterrows():
        split_5 = row44_5['Stock Splits']
        if split_5 > 0:
            no_stock_5 = no_stock_5 + (split_5 - 1) * no_stock_5
        else:
            no_stock_5 = no_stock_5
        A1_5.at[index44_5, 'Num Stock'] = no_stock_5

    ini_div_5 = 0
    for index55_5, row55_5 in A1_5.iterrows():
        div_5 = row55_5['Dividends']
        num_stock_2_5 = row55_5['Num Stock']
        total_div_5 = div_5 * num_stock_2_5
        A1_5.at[index55_5, 'Dividends'] = total_div_5
        ini_div_5 = ini_div_5 + total_div_5

    ####################### Calculation for finding interest ########################

    # Initial value of investment
    P_5 = initial_stock_price_5 * A1_5.loc[first_index_5].at["Num Stock"]

    # Final value of investment
    A_5 = end_stock_price_5 * A1_5.loc[last_index_5].at["Num Stock"]

    # No. of times compounded annually
    n_5 = 1

    # No. of years for investment
    today_5 = date.today()
    current_year_5 = today.year
    t_5 = (current_year_5 - int(year_5))

    # Interest
    r_5 = (n_5 * (((A_5 + ini_div_5) / P_5) ** (1 / (n_5 * t_5)) - 1)) * 100

    print("The interest rate is " + str(r_5))

    ####################################

    final_interest_data = np.array([[str(t) + " Years", r, A1.loc[last_index].at["Num Stock"], ini_div],
                                    [str(t_5) + ' Years', r_5, A1_5.loc[last_index_5].at["Num Stock"], ini_div_5]])

    final_interest_rate = pd.DataFrame(final_interest_data, columns=["No. of years", "Interest", "Total Num of Shares",
                                                                     "Total Dividend Value"])
else:
    # Extract Stock price history based on year
    A1 = stock_price_history[(stock_price_history['Date'] > str(yu))]
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
    no_stock = int(A1.loc[first_index].at["Num Stock"])
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

    final_interest_data = np.array([[str(t) + ' Years', r, A1.loc[last_index].at["Num Stock"], ini_div],
                                    [str(t) + ' Years', r, A1.loc[last_index].at["Num Stock"], ini_div]])

    final_interest_rate = pd.DataFrame(final_interest_data, columns=["No. of years", "Interest", "Total Num of Shares",
                                                                     "Total Dividend Value"])

a = 1

if a == 1:
    # Create a Pandas Excel writer using XlsxWriter engine
    writer = pd.ExcelWriter(Filename, engine='openpyxl', mode='a', if_sheet_exists='replace')

    # Convert the bonus history data frame to an XlsxWriter Excel object
    A1.to_excel(writer, sheet_name="Final Calculation and History", index=False)
    final_interest_rate.to_excel(writer, sheet_name='Final Interest Rate', index=False)

    # Close the Pandas Excel writer and output the Excel file
    writer.save()
