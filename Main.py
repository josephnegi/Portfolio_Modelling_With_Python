# -*- coding: utf-8 -*-
"""
Created on Wed Jul 31 02:29:27 2024

@author: joeyn
"""
import yfinance as yf
import pandas as pd
import numpy as np
import xlwings as xw
from datetime import datetime

#main code
'''
Shape definition for tickers:
    [ticker, no_of_stocks, avg_stock_purchase_paid, long_short]
    user_positions_data = [
        ['AAPL', 23, 18.2, 1],
        ['MSFT', 10, 67.8, 1],
        ['GOOG', 13, 34.4, 1],
        ['IBM', 16, 23.7, -1]]
    
Code that gives the tickers:
    for x,y in enumerate(matrix):
        print(matrix[x][0])
        
code for debugging–
Enter data for position no. 1:
FRPT 71 138 1
Enter data for position no. 2:
WMS 124 80 1
Enter data for position no. 3:
AGCO 100 99 1
Enter data for position no. 4:
BDX 40 248 -1
Enter data for position no. 5:
SYY 140 71 -1
'''
no_of_positions = int(input("Enter the number of positions in your portfolio: "))
                   
print("Format for entering data:\nStock ticker_<space>_no of stocks_<space>_average purchase price_<space>_Long(1) or Short(-1)\n")
user_positions_data = []
entered_tickers = set()

for x in range(no_of_positions):
    while True:
        i = input(f"Enter data for position no. {x+1}:\n")
        j = i.split()
        try:
            if len(j) != 4:
                raise ValueError("Incorrect format. Please enter exactly 4 fields.")
            ticker = j[0]
            if ticker in entered_tickers:
                raise ValueError("Ticker already used. Please enter a different ticker.")
            j[1] = int(j[1])
            j[2] = float(j[2])
            j[3] = int(j[3])
            if j[3] not in [-1, 1]:
                raise ValueError("Long/Short field must be 1 for Long or -1 for Short.")
            user_positions_data.append(j)
            entered_tickers.add(ticker)
            break  # Exit the while loop if everything is correct
        except ValueError as e:
            print(f"Error: {e}. Please try again.")

print("User positions data:")
for position in user_positions_data:
    print(position)
#fetching max historical data for all tickers, and assigning all tickers to a new list
stock_data = {}
tickers = []
for x, y in enumerate(user_positions_data):
    ticker = user_positions_data[x][0]
    tickers.append(ticker)
    max_retries = 3
    for attempt in range(max_retries):
        try:
            stock_data[ticker] = yf.download(ticker, period="max", interval="1d")
            if stock_data[ticker].empty:
                raise ValueError(f"No data available for {ticker}")
            print(f"Successfully downloaded data for {ticker}")
            break
        except Exception as e:
            print(f"Error downloading data for {ticker}: {e}")
            if attempt < max_retries - 1:
                print(f"Retrying... (Attempt {attempt + 2} of {max_retries})")
            else:
                print(f"Failed to download data for {ticker} after {max_retries} attempts.")
                while True:
                    user_choice = input(f"Do you want to enter a new ticker to replace {ticker}? (yes/no): ").lower()
                    if user_choice == 'yes':
                        new_ticker = input("Enter the new ticker symbol: ")
                        new_no_of_stocks = int(input("Enter the number of stocks: "))
                        new_avg_purchase_price = float(input("Enter the average purchase price: "))
                        new_long_short = int(input("Enter 1 for long or -1 for short: "))
                        user_positions_data[x] = [new_ticker, new_no_of_stocks, new_avg_purchase_price, new_long_short]
                        tickers[x] = new_ticker
                        try:
                            stock_data[new_ticker] = yf.download(new_ticker, period="max", interval="1wk")
                            if stock_data[new_ticker].empty:
                                raise ValueError(f"No data available for {new_ticker}")
                            print(f"Successfully downloaded data for {new_ticker}")
                            break
                        except Exception as e:
                            print(f"Error downloading data for {new_ticker}: {e}")
                            continue
                    elif user_choice == 'no':
                        print(f"Skipping {ticker}. This may affect your portfolio calculations.")
                        del tickers[x]
                        del user_positions_data[x]
                        no_of_positions -= 1
                        break
                    else:
                        print("Invalid input. Please enter 'yes' or 'no'.")
                break

# Create a list to store the earliest dates
earliest_dates = []

# Populate the list with the earliest dates for each stock
for ticker in tickers:
    earliest_dates.append(stock_data[ticker].index.min())

# Find the maximum date (latest date)
#Ask the user if they want to go with the earliest date or another date
q = input(f"The earliest date of the dataset is {max(earliest_dates)}, do you wish to start the analysis on a later date? (yes/no)\n").lower()
if q == "yes":
    l_date = input("Enter the date in the format yyyy-mm-dd:\n")
    latest_date = l_date
else:
    latest_date = max(earliest_dates)

# Create a new DataFrame with the latest date as the start date
trimmed_data = {}
for ticker in tickers:
    trimmed_data[ticker] = stock_data[ticker][stock_data[ticker].index >= latest_date]['Adj Close']

# Combine the data into a single DataFrame
combined_df = pd.concat(trimmed_data, axis=1)

# Rename columns for better readability
combined_df.columns = tickers

'''
code to test if the new combined dataframe is working:
    stock_data['AAPL'].loc['2004-08-16', 'Adj Close']
'''
#starting the historical CEW portfolio performance analysis
print("Starting the Historical CEW Performance Analysis")

#copying the dataframe as a safety measure
df = combined_df.copy(deep=True)

# Calculate weekly returns
returns = df.pct_change().dropna()

# Define the long/short multiplier vector (long +1 and short -1)
# Make sure it has the same length as the number of assets (i.e., returns.shape[1])
#num_elements = len(user_positions_data)
#multipliers = np.zeros(num_elements, dtype=int)
multipliers = np.zeros(no_of_positions, dtype=int)

for i,data in enumerate(user_positions_data):
    multipliers[i] = data[3]
    
# Element-wise multiplication of returns by multipliers
adjusted_returns = returns * multipliers

# Calculate equally weighted portfolio returns
equal_weights = np.full((returns.shape[1],), 1 / returns.shape[1])
cew_returns = adjusted_returns.dot(equal_weights)

# Initial portfolio index value
initial_index_value = 100

# Calculate the cumulative performance of the portfolio
cew_port_index = (1 + cew_returns).cumprod() * initial_index_value

# Prepare a DataFrame to export
export_df = pd.DataFrame(index=df.index)
for x,y in enumerate(user_positions_data):
    ticker = user_positions_data[x][0]
    export_df[f"{ticker} Returns"] = returns[ticker]
export_df['CEW Returns'] = cew_returns
export_df['CEW Portfolio Index'] = cew_port_index    

#function to print the CEW performance in a sheet
print(export_df.tail())

'''
excelPrint(excel_file_location, sheet_name, new=True/False, dataframe_data)

'''
#Starting the Historical NRB Performance Analysis
print("Starting the Historical NRB Performance Analysis")

#copying the dataframe as a safety measure
df_nrb = combined_df.copy(deep=True)

#creating a vector for the number of shares for the user positions
no_of_shares = np.zeros(no_of_positions, dtype=int)
for i,data in enumerate(user_positions_data):
    no_of_shares[i] = data[1]
    
#calculating the total invested value by multiplying the no of shares and price of shares, and adding them
total_invested_value = 0
for x in user_positions_data:
    total_invested_value += x[1]*x[2]
    
# Creating a Static vector for the no of shares and the long short
weighted_vector = no_of_shares * multipliers

# Custom function to calculate the required result for each row
def calculate_custom_return(row, previous_row):
    current_dot_product = np.dot(row, weighted_vector)
    if previous_row is None:
        return None
    else:
        previous_dot_product = np.dot(previous_row, weighted_vector)
        return current_dot_product - previous_dot_product

# Apply the custom function
nrb_returns = df_nrb.apply(
    lambda row: calculate_custom_return(row, df_nrb.iloc[df_nrb.index.get_loc(row.name) - 1] if df_nrb.index.get_loc(row.name) > 0 else None), axis=1
)
#Calculating the portfolio current value via the cumulative sum of returns
nrb_port_value = nrb_returns.cumsum() + total_invested_value

# Prepare a DataFrame to export
df_nrb['Portfolio Return'] = nrb_returns
df_nrb['Portfolio Value'] = nrb_port_value

print(df_nrb.tail())

# Function to write dataframe and format date column
def write_df_to_sheet(sheet, df, date_col='Date'):
    #df_reset = df.reset_index(names=[date_col])
    sheet.range('A1').options(index=True).value = df
    date_range = sheet.range(f'A2:A{len(df)+1}')
    date_range.number_format = 'dd-mmm-yy'  

# Create a new Excel workbook
wb = xw.Book()

# Write individual stock data to separate sheets
for ticker in tickers:
    sheet = wb.sheets.add(ticker)
    write_df_to_sheet(sheet, stock_data[ticker])

# Write df_nrb to a sheet
sheet2 = wb.sheets.add('NRB Performance')
write_df_to_sheet(sheet2, df_nrb)

# Write export_df to a sheet
sheet1 = wb.sheets.add('CEW Performance')
write_df_to_sheet(sheet1, export_df)
  
# Rename 'Sheet1' to 'Summary' and add user_positions_data
if 'Sheet1' not in [sheet.name for sheet in wb.sheets]:
    summary_sheet = wb.sheets.add('Summary')
else:
    summary_sheet = wb.sheets['Sheet1']
    summary_sheet.name = 'Summary'

# Define headers
headers = ['Ticker', 'No of Shares', 'Purchase Price per Share', 'Long/Short']

# Write headers
summary_sheet.range('A1').value = headers

# Write user_positions_data
for i, position in enumerate(user_positions_data, start=2):
    summary_sheet.range(f'A{i}').value = position[0]  # Ticker
    summary_sheet.range(f'B{i}').value = position[1]  # No of Shares
    summary_sheet.range(f'C{i}').value = position[2]  # Purchase Price per Share
    summary_sheet.range(f'D{i}').value = 'Long' if position[3] == 1 else 'Short'  # Long/Short

# Auto-fit columns
summary_sheet.range('A:D').columns.autofit()



# Format headers
header_range = summary_sheet.range('A1:D1')
header_range.font.bold = True
header_range.color = (200, 200, 200)  # Light gray background

# Add borders
data_range = summary_sheet.range(f'A1:D{len(user_positions_data)+1}')
data_range.api.Borders.Weight = 2  # Medium weight border

print("Summary sheet has been created and populated with user positions data.")

# Save the workbook
file_name = 'Portfolio_Modelling_Py.xlsx'
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
try:
    wb.save(file_name)
except Exception as e:
    print(f"Error saving file: {e}\n")
    
    file_name = f'Portfolio_Modelling_{timestamp}.xlsx'
    print(f"Attempting to save as {file_name}")
    wb.save(file_name)

#asking the user if they want to keep the excel file open
wb.close() if input('Do you wish to close the excel sheet that is open? (yes/no)\n').lower() == 'yes' else None

print(f"Excel file has been created with all the data, and saved as {file_name}")