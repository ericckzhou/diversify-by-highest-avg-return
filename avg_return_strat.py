from typing import final
import numpy as np
import pandas as pd
import requests
import math
import xlsxwriter

def chunks(lst, n):
    for i in range(0, len(lst), n):
        yield lst[i: i+n]

def get_portfolio_size():
    global size
    size = input('Enter size of portfolio: ')
    while (size.isnumeric() == False):
        print("This is not an integer!")
        size = input('Enter size of portfolio: ')
    return int(size)


IEX_CLOUD_API_TOKEN = 'Tpk_059b97af715d417d9f49f50b51b1c448'
print()
print()
print("==================================================")
print("This program sorts by highest Average Return in a given file.")
print("Note: This uses IEX Cloud's API (Sandbox ver)")
print("This program does not use real data.")
print("==================================================")
print()
print()
file = input("Enter the file-name (.csv file) of the stocks you are interested in: ")
print("File name entered: " + str(file) + '.csv')
print()
print("Loading...")
file_name = str(file) + ".csv"
try:
    stocks = pd.read_csv(file_name)
except FileNotFoundError:
    print("The file specified is not found in the current directory.")
    exit()

symbol_groups = list(chunks(stocks['Ticker'], 100))
symbol_strings = []
for i in range(0, len(symbol_groups)):
    symbol_strings.append(','.join(symbol_groups[i]))

myColumns = ['Ticker', 'Price','1 Month Return', '3 Months Return', '6 Months Return', 'YTD Return', '1 Year Return', 'AVG Return %', 'Num Shares to Buy']
ret_dataframe = pd.DataFrame(columns=myColumns)
for symbol_string in symbol_strings:
    batch_call_url = f'https://sandbox.iexapis.com/stable/stock/market/batch/?types=stats,quote&symbols={symbol_string}&token={IEX_CLOUD_API_TOKEN}'
    data = requests.get(batch_call_url).json()
    for symbol in symbol_string.split(','):
        ret_dataframe=ret_dataframe.append(
            pd.Series(
                [
                symbol,
                data[symbol]['quote']['latestPrice'],
                data[symbol]['stats']['month1ChangePercent'],
                data[symbol]['stats']['month3ChangePercent'],
                data[symbol]['stats']['month6ChangePercent'],
                data[symbol]['stats']['ytdChangePercent'],
                data[symbol]['stats']['year1ChangePercent'],
                'N/A',
                'N/A'
                ],
                index=myColumns
            ),
            ignore_index=True
        )
#modifies original= inplace.
#ret_dataframe.sort_values('YTD Return', ascending=False, inplace=True)
#final_dataframe = ret_dataframe[:25]
#final_dataframe.reset_index(drop=True, inplace=True) 
#resets the index from original (inplace) & drops original index
for i in range(0, len(ret_dataframe['Ticker'])):
    month1 = ret_dataframe['1 Month Return'][i]
    month3 = ret_dataframe['3 Months Return'][i]
    month6 = ret_dataframe['6 Months Return'][i]
    ytd = ret_dataframe['YTD Return'][i]
    year1 = ret_dataframe['1 Year Return'][i]
    try:
        ret_dataframe.loc[i, 'AVG Return %'] = ((float(month1) + float(month3) + float(month6) + float(ytd) + float(year1))/5)#percentage
    except:
        ret_dataframe.loc[i, 'AVG Return %'] = 0

print()
ret_dataframe.sort_values('AVG Return %', ascending = False, inplace = True)
try:
    n = input("Choose top 'n' stocks to select (with highest avg return): ")
except:
    print("N is invalid.")
    n = input("Choose top 'n' stocks to select (with highest avg return): ")
n = int(n)
ret_dataframe=ret_dataframe[:n]
ret_dataframe.reset_index(drop=True, inplace=True)

portfolio_size = math.floor(get_portfolio_size() / n)
for i in range(0, len(ret_dataframe['Ticker'])):
    num_shares = math.floor((portfolio_size / ret_dataframe['Price'][i]))
    ret_dataframe.loc[i, 'Num Shares to Buy'] = num_shares

ret_dataframe = ret_dataframe[['Ticker', 'Price', 'AVG Return %', 'Num Shares to Buy']]
writer = pd.ExcelWriter('output.xlsx', engine='xlsxwriter')
ret_dataframe.to_excel(writer, sheet_name='Average Return Strategy', index=False)
background_color = '#000000'
font_color = '#ffffff'

string_template = writer.book.add_format(
    {
        'font_color': font_color,
        'bg_color': background_color,
        'border': 1
    }
)

dollar_template = writer.book.add_format(
    {
        'num_format':'$0.00',
        'font_color': font_color,
        'bg_color': background_color,
        'border': 1
    }
    )

integer_template = writer.book.add_format(
    {
        'num_format':'0',
        'font_color': font_color,
        'bg_color': background_color,
        'border': 1
    }
)

percent_template = writer.book.add_format(
    {
        'num_format':'0.0%',
        'font_color': font_color,
        'bg_color': background_color,
        'border': 1
    }
)
column_format = {
    'A': ['Ticker', string_template],
    'B': ['Price', dollar_template],
    'C': ['AVG Return %', percent_template],
    'D': ['Num Shares to Buy', integer_template]
}
for column in column_format.keys():
    writer.sheets['Average Return Strategy'].set_column(f'{column}:{column}', 18, column_format[column][1])
    writer.sheets['Average Return Strategy'].write(f'{column}1', column_format[column][0], string_template)

print("The file has been outputed!")
writer.save()