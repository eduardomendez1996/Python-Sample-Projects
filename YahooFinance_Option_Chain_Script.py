# this project uses Beautiful soup and Yfinance packages to scrape the Yahoo Finance Website and pull the optio chains for a user-inputed ticker symbol. It also saves the file to my personal drive as an excel sheet.
It has an additional step where it checks if the options are in-the-money. If options are in the money, then it highlights th entire row in a light green color.

import requests
import pandas as pd
import yfinance as yf
from bs4 import BeautifulSoup
from datetime import datetime
from io import StringIO
import tkinter as tk
from tkinter import simpledialog

# Function to get ticker symbol using a popup window. user will input the ticker symbol they need to pull. this code will also capitalize in case user inputs in lowercase
def get_ticker():
    root = tk.Tk()
    root.withdraw()  # Hide the root window
    ticker = simpledialog.askstring(title="Input", prompt="Please enter ticker symbol:")
    return ticker.upper()

# Function to get the last close price using yfinance package
def get_last_close_price(ticker):
    stock = yf.Ticker(ticker)
    close_price = stock.history(period="1d")['Close'][0]
    return close_price

# Get the ticker symbol from the user
ticker = get_ticker()

# Get the last close price
close_price = get_last_close_price(ticker)

# Define the URL for options data
url = f'https://finance.yahoo.com/quote/{ticker}/options/'

# Fetch the webpage content for options data
headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3'}
response = requests.get(url, headers=headers)
soup = BeautifulSoup(response.content, 'html.parser')

# Find all tables on the webpage
tables = soup.find_all('table')

# Parse each table into a pandas DataFrame
dataframes = []
for table in tables:
    df = pd.read_html(StringIO(str(table)))[0]
    dataframes.append(df)

# Ensure there are at least two tables, one for calls and one for puts
if len(dataframes) < 2:
    raise ValueError("Less than 2 tables found on the webpage.")

# Separate the dataframes for Calls and Puts
calls_df = dataframes[0]
puts_df = dataframes[1]

# Add 'In the Money' column to calls_df
calls_df['In the Money'] = calls_df['Strike'].apply(lambda x: 'Yes' if x < close_price else 'No')

# Add 'In the Money' column to puts_df
puts_df['In the Money'] = puts_df['Strike'].apply(lambda x: 'Yes' if x > close_price else 'No')

# Get the current date in MMDDYYYY format
current_date = datetime.now().strftime('%m%d%Y')

# list the file path where excel file should be saved. name should be in this format- tickerMMDDYYYY
file_name = f'{ticker}{current_date}.xlsx'
file_path = rf'C:\Users\mende\OneDrive\Documents\Python Projects\Options Chain\{file_name}'

# Save the DataFrames to an Excel file with 2 different tabs/sheets, one for calls and one for puts
with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
    # Save Calls DataFrame to sheet "Calls"
    calls_df.to_excel(writer, index=False, sheet_name='Calls')
    # Save Puts DataFrame to sheet "Puts"
    puts_df.to_excel(writer, index=False, sheet_name='Puts')

    # Get the xlsxwriter workbook and worksheet
    workbook = writer.book
    calls_worksheet = writer.sheets['Calls']
    puts_worksheet = writer.sheets['Puts']

    # Formats header to dark blue fill, white and BOLD text
    header_format = workbook.add_format({
        'bold': True,
        'text_wrap': True,
        'valign': 'top',
        'fg_color': '#00008B',  # Dark blue
        'font_color': 'white',
        'border': 1
    })

    # Format the Calls tab. adjusts column length, changes header font and text color
    for col_num, value in enumerate(calls_df.columns.values):
        calls_worksheet.write(0, col_num, value, header_format)
    for i, col in enumerate(calls_df.columns):
        max_len = calls_df[col].astype(str).map(len).max()
        calls_worksheet.set_column(i, i, max_len + 2)  # Adding extra space

    #  conditional format Calls tab. makes row light green if options are in the money
    calls_worksheet.conditional_format(1, 0, len(calls_df), len(calls_df.columns) - 1, 
                                       {'type': 'formula',
                                        'criteria': '=$L2="Yes"',
                                        'format': workbook.add_format({'bg_color': '#C6EFCE'})})

    # Format the Puts tab. adjusts column length, changes header font and text color
    for col_num, value in enumerate(puts_df.columns.values):
        puts_worksheet.write(0, col_num, value, header_format)
    for i, col in enumerate(puts_df.columns):
        max_len = puts_df[col].astype(str).map(len).max()
        puts_worksheet.set_column(i, i, max_len + 2)  # Adding extra space

    # conditional format Puts tab. makes row light green if options are in the money
    puts_worksheet.conditional_format(1, 0, len(puts_df), len(puts_df.columns) - 1, 
                                      {'type': 'formula',
                                       'criteria': '=$L2="Yes"',
                                       'format': workbook.add_format({'bg_color': '#C6EFCE'})})

print("Data has been saved to the Excel file successfully.")














