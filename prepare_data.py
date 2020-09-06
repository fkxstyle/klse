import os
from datetime import datetime, date
import re
import pytz
import time
import json
import logging
import requests
import randomcolor
import numpy as np
import pandas as pd
import urllib.parse
import shutil
import glob
from pathlib import Path
from dateutil import tz
from random import randint
from decimal import Decimal
from datetime import datetime
from utils import string_value_converter, best_fit_slope, prepare_numeric_value_format
from statistics import mean
import xlsxwriter

# define constant
million = 1000000
malaysia_government_bond_rate = 3.43

# define year
now = datetime.now()
today_date = date.today()
this_year = str(now.year)
last_year = str(now.year - 1)
last_2_year = str(now.year - 2)
last_3_year = str(now.year -3)

# define columns for each document
balance_sheet_columns = [
    'Total Current Assets',
    'Cash, Cash Equivalents and Short Term Investments',
    'Net Intangible Assets',
    'Total Current Liabilities',
    'Current Debt and Capital Lease Obligation',
    'Long Term Debt and Capital Lease Obligation',
    'Total Equity',
    'Common Shares Outstanding',
    'Total Assets',
    'Total Liabilities',
    'Trade and Other Receivables, Current',
    'Goodwill',
    'Trade/Accounts Receivable, Current',
    'Inventories',
    'Trade/Accounts Payable, Current',
]
cash_flow_columns = [
    'Purchase of Property, Plant and Equipment',
    'Cash Dividends Paid',
    'Cash Flow from Operating Activities, Indirect',
]
income_statement_columns = [
    'Total Revenue',
    'Total Operating Profit/Loss',
    'Diluted Net Income Available to Common Stockholders',
    'Provision for Income Tax',
    'Pretax Income',
    'Net Interest Income/Expense',
    'Depreciation, Amortization and Depletion, Supplemental',
    'Non-Controlling/Minority Interest',
    'Cost of Revenue',
]

# define folder path
sector_root = Path().absolute()
company_folder_root = os.path.join(sector_root,'downloads')

# Delete .DS_Store
ds_store_path = glob.glob(company_folder_root + '/.DS_Store')
for fpath in ds_store_path:
    if os.path.isfile(fpath):
        os.remove(fpath)

# prepare columns and indexes list for data frame
indexes = balance_sheet_columns + cash_flow_columns + income_statement_columns
company_list=os.listdir (company_folder_root)

# Create an new Excel file and add a worksheet.
workbook = xlsxwriter.Workbook('financial_data.xlsx')
worksheet = workbook.add_worksheet()

count = 1
# Write first column in excel
with open('excel_output.json') as json_file:
    excel_output = json.load(json_file)
for category, key_names in excel_output.items():
    worksheet.write(count, 0, category)
    count += 1
    for key_name in key_names:
        worksheet.write(count, 0, key_name)
        count += 1
    count += 1
column_index = 1


# prepare dataframe for data of each company in the folder
for company in company_list:
    company_dir = os.path.join(company_folder_root, company)
    document_list = os.listdir(company_dir)

    # load data.json to stock_data
    with open(company_dir + '/data.json') as json_file:
        stock_data = json.load(json_file)

    print('jimmy')
    print(stock_data['name'])

    # build data frame from excel file
    for document in document_list:
        document_dir =  os.path.join(company_dir,document)
        if 'Balance Sheet' in document:
            key_name = '{}_BalanceSheet_Annual_Restated'.format(stock_data['code'])
            df_bs = pd.read_excel(document_dir,sheet_name=stock_data['code']).fillna(0)
            
            df_bs[key_name] = df_bs[key_name].str.lstrip().str.rstrip()
            df_bs = df_bs.set_index(key_name)
            df_bs_filtered = df_bs.reindex(balance_sheet_columns).loc[balance_sheet_columns].fillna(0)

        elif 'Cash Flow' in document:
            key_name = '{}_CashFlow_Annual_Restated'.format(stock_data['code'])
            df_cf = pd.read_excel(document_dir,sheet_name=stock_data['code']).fillna(0)
            df_cf[key_name] = df_cf[key_name].str.lstrip().str.rstrip()
            df_cf = df_cf.set_index(key_name)
            df_cf_filtered = df_cf.reindex(cash_flow_columns).loc[cash_flow_columns].fillna(0)
                    
        elif 'Income Statement' in document:
            key_name = '{}_IncomeStatement_Annual_Restated'.format(stock_data['code'])
            df_is = pd.read_excel(document_dir,sheet_name=stock_data['code']).fillna(0)
            df_is[key_name] = df_is[key_name].str.lstrip().str.rstrip()
            df_is = df_is.set_index(key_name)
            df_is_filtered = df_is.reindex(income_statement_columns).loc[income_statement_columns].fillna(0)
        
    company_data = pd.concat([df_bs_filtered, df_cf_filtered, df_is_filtered], sort=False)
    company_data = company_data.drop(columns='TTM')
    company_data = company_data.reindex(indexes)

    # clean up data
    company_data = company_data.apply(lambda x: x.str.replace(',','')).fillna(0)
    company_data = company_data.apply(pd.to_numeric)
    last_year_data = company_data[last_year]
    last_2_year_data = company_data[last_2_year]
    last_3_year_data = company_data[last_3_year]

    first_three_year = sorted(company_data.columns)[:3]
    last_three_year = sorted(company_data.columns)[-3:]
    all_years = company_data.columns

    # feed data to json
    stock_data['Year of Study'] = last_year
    stock_data['EBITDA'] = {
        'value': last_year_data['Pretax Income'] + last_year_data['Net Interest Income/Expense'] + last_year_data['Depreciation, Amortization and Depletion, Supplemental']
    }
    stock_data['Book Value Per Share'] = {
        'value': last_year_data['Total Equity'] / last_year_data['Common Shares Outstanding']
    }
    stock_data['Net Tangible Book Value Per Share'] = {
        'value': (last_year_data['Total Equity'] -  last_year_data['Net Intangible Assets']) / last_year_data['Common Shares Outstanding']
    }
    stock_data['Working Capital'] = {
        'value': last_year_data['Total Current Assets'] - last_year_data['Total Current Liabilities']
    }
    stock_data['Total Debt'] = {
        'value': last_year_data['Current Debt and Capital Lease Obligation'] + last_year_data['Long Term Debt and Capital Lease Obligation']
    }
    stock_data['Dividend Per Share'] = {
        'value': last_year_data['Cash Dividends Paid'] / last_year_data['Common Shares Outstanding']
    }
    stock_data['Earning Per Share'] =  {
        'value': last_year_data['Diluted Net Income Available to Common Stockholders'] / last_year_data['Common Shares Outstanding']
    }
    stock_data['Interest Bearing Debt'] = {
        'value': stock_data['Total Debt']['value']
    }
    stock_data['Tax Rate'] = {
        'value': abs(last_year_data['Provision for Income Tax'] / last_year_data['Pretax Income'])
    }
    stock_data['NOPAT'] = {
        'value': last_year_data['Total Operating Profit/Loss'] * (1 - stock_data['Tax Rate']['value'])
    }
    stock_data['Invested Capital'] =  {
        'value': stock_data['Interest Bearing Debt']['value'] + last_year_data['Total Equity'] + last_year_data['Non-Controlling/Minority Interest'] - last_year_data['Cash, Cash Equivalents and Short Term Investments'] - last_year_data['Goodwill']
    }
    stock_data['ROIC'] =  {
        'value': stock_data['NOPAT']['value'] /  stock_data['Invested Capital']['value']
    }
    stock_data['Adequate size of enterprise (Market Cap) (RM mil)'] = {
        'value': string_value_converter(stock_data['market_cap'].replace(",", ""))
    }
    stock_data['2 x Current asset >= Current liabilities'] = {
        'value': last_year_data['Total Current Assets'] / last_year_data['Total Current Liabilities'],
        'condition': last_year_data['Total Current Assets'] / last_year_data['Total Current Liabilities'] > 2
    }
    # stock_data['Net current asset (working capital) > long term debt'] = stock_data['Working Capital'] > last_year_data['Long Term Debt and Capital Lease Obligation']
    stock_data['Net current asset (working capital) > long term debt'] = {
        'value': stock_data['Working Capital']['value'] / last_year_data['Long Term Debt and Capital Lease Obligation'],
        'condition': stock_data['Working Capital']['value'] /last_year_data['Long Term Debt and Capital Lease Obligation'] > 1,
    }

    stock_data['Debt < 2 x Stock Equity (book value)'] = {
        'value': last_year_data['Total Liabilities'] / last_year_data['Total Equity'],
        'condition': last_year_data['Total Liabilities'] / last_year_data['Total Equity'] < 2 
    }
    sr_cash_dividends_paid = company_data.loc['Cash Dividends Paid'].sort_index(ascending=True).reset_index(drop=True)
    sr_cash_dividends_paid = sr_cash_dividends_paid.abs()
    sr_cash_dividends_paid_index = np.array(list(sr_cash_dividends_paid.index.values), dtype=np.float64)
    sr_cash_dividends_paid_values = np.array(sr_cash_dividends_paid.values.tolist(), dtype=np.float64)
    gradient = best_fit_slope(sr_cash_dividends_paid_index, sr_cash_dividends_paid_values)
    stock_data['Earning stability in the past 10 years (EPS)'] = {
        'value': gradient,
        'condition': gradient > 0,
    }

    sr_dividend_year = (company_data.loc['Cash Dividends Paid'] != 0)
    year_with_dividend = sr_dividend_year[sr_dividend_year == True].index.values.tolist()
    stock_data['Dividend uninterupted for past 20 years'] = {
        'value': ', '.join(year_with_dividend),
        'condition': (company_data.loc['Cash Dividends Paid'] != 0 ).all(),
    }
    stock_data['Min. inc of 1/3 using three year averages at beginning & end'] = {
        'value': abs(company_data.loc['Cash Dividends Paid', last_three_year].sum()) / abs(company_data.loc['Cash Dividends Paid', first_three_year].sum()) - 1,
        'condition': abs(company_data.loc['Cash Dividends Paid', last_three_year].sum()) / abs(company_data.loc['Cash Dividends Paid', first_three_year].sum()) - 1 > 1/3,
    }
    stock_data['P/E < 15 average earnings of past three years'] = {
        'value': (float(stock_data['current_pe_ratio'])+float(stock_data['last_year_pe_ratio'])+float(stock_data['last_2_years_pe_ratio'])) / 3,
        'condition': (float(stock_data['current_pe_ratio'])+float(stock_data['last_year_pe_ratio'])+float(stock_data['last_2_years_pe_ratio'])) / 3 < 15,
    }
    stock_data['Price < 1.5 x Book value'] = {
        'value': float(stock_data['price']) / float(stock_data['Book Value Per Share']['value']),
        'condition': float(stock_data['price']) / float(stock_data['Book Value Per Share']['value']) < 1.5
    }
    stock_data['P/E * P/NTA < 22.5'] = {
        'value':  float(stock_data['current_pe_ratio']) * float(stock_data['price']) / stock_data['Net Tangible Book Value Per Share']['value'],
        'condition': float(stock_data['current_pe_ratio']) * float(stock_data['price']) / stock_data['Net Tangible Book Value Per Share']['value'] < 22.5
    }
    # Stockaholics method:stock_data['ROE > 10%'] = last_year_data['Diluted Net Income Available to Common Stockholders'] / last_year_data['Total Equity'] > 0.1
    stock_data['ROE > 10%'] = {
        'value': last_year_data['Diluted Net Income Available to Common Stockholders'] / last_year_data['Total Equity'] * 100,
        'condition': last_year_data['Diluted Net Income Available to Common Stockholders'] / last_year_data['Total Equity'] * 100 > 10
    }
    stock_data['Net cash companies (RM mil)'] = {
        'value': last_year_data['Cash, Cash Equivalents and Short Term Investments']
    }
    stock_data['Companies with good FCF (RM mil)'] = {
        'value': last_year_data['Cash Flow from Operating Activities, Indirect'] - abs(last_year_data['Purchase of Property, Plant and Equipment'])
    }
    stock_data['Dividend Pay-out ratio ≤ 75%'] = {
        'value': abs(stock_data['Dividend Per Share']['value']) / abs(stock_data['Earning Per Share']['value']) * 100,
        'condition': abs(stock_data['Dividend Per Share']['value']) / abs(stock_data['Earning Per Share']['value']) * 100 < 75
    }
    stock_data['Dividend Yield ≥ 10-year Malaysia Government Bond (3.43%)'] = {
        'value': string_value_converter(stock_data['dividend_yield']),
        'condition': string_value_converter(stock_data['dividend_yield']) > malaysia_government_bond_rate
    }
    # adam khoo
    stock_data['3 years x Net income > Long term debt'] = {
        'value': last_year_data['Long Term Debt and Capital Lease Obligation'] / last_year_data['Diluted Net Income Available to Common Stockholders'],
        'condition': last_year_data['Long Term Debt and Capital Lease Obligation'] / last_year_data['Diluted Net Income Available to Common Stockholders'] < 3
    }
    stock_data['ROIC > ROE'] = {
        'value': stock_data['ROIC']['value'] / (last_year_data['Diluted Net Income Available to Common Stockholders'] / last_year_data['Total Equity']),
        'condition': stock_data['ROIC']['value'] / (last_year_data['Diluted Net Income Available to Common Stockholders'] / last_year_data['Total Equity']) > 1
    }

    # # Comparison Method 2: Benjamin Graham
    stock_data["Price"] = {
        'value': float(stock_data['price'])
    }
    stock_data['Number of share of common (mil)'] = {
        'value': last_year_data['Common Shares Outstanding'] / million
    }
    stock_data['Market value of common (RM mil)'] = {
        'value': string_value_converter(stock_data['market_cap'].replace(',','')) / million
    }
    stock_data['Debt (RM mil)'] = {
        'value': (last_year_data['Current Debt and Capital Lease Obligation'] + last_year_data['Long Term Debt and Capital Lease Obligation']) / million
    }
    stock_data['Total capitalization at market (RM mil)'] = {
        'value': stock_data['Market value of common (RM mil)']['value'] + stock_data['Debt (RM mil)']['value']
    }
    stock_data['Book value per share'] = {
        'value': stock_data['Net Tangible Book Value Per Share']['value']
    }
    stock_data['Sales (RM mil)'] = {
        'value': last_year_data['Total Revenue'] / million
    }
    stock_data['Net income'] = {
        'value': last_year_data['Diluted Net Income Available to Common Stockholders']
    }
    stock_data["EPS (Year, n=10)"] = {
        'value': last_year_data['Diluted Net Income Available to Common Stockholders'] / last_year_data['Common Shares Outstanding']
    }
    stock_data["EPS (Year, n=9)"] = {
        'value': last_2_year_data['Diluted Net Income Available to Common Stockholders'] / last_2_year_data['Common Shares Outstanding']
    }
    stock_data["EPS (Year, n=8)"] = {
        'value': last_3_year_data['Diluted Net Income Available to Common Stockholders'] / last_3_year_data['Common Shares Outstanding']
    }
    stock_data['Current dividend rate'] = {
        'value': string_value_converter(stock_data['dividend_yield'])
    }

    # # Ratios:
    stock_data['P/E'] = {
        'value': stock_data['last_year_pe_ratio']
    }
    stock_data['P/book value'] = {
        'value': float(stock_data['price']) / stock_data['Book Value Per Share']['value']
    }
    stock_data['Dividend yield'] = {
        'value': stock_data['dividend_yield']
    }
    stock_data['Net/sales'] = {
        'value': last_year_data['Diluted Net Income Available to Common Stockholders'] / last_year_data['Total Revenue']
    }
    stock_data['Earning/book value'] = {
        'value': last_year_data['Diluted Net Income Available to Common Stockholders'] / (stock_data['Net Tangible Book Value Per Share']['value'] * last_year_data['Common Shares Outstanding'])
    }
    stock_data['Current assets/ liabilities'] = {
        'value': last_year_data['Total Current Assets'] / last_year_data['Total Current Liabilities']
    }
    stock_data['Working capital/ debt'] = {
        'value': stock_data['Working Capital']['value'] / ( last_year_data['Long Term Debt and Capital Lease Obligation'] + last_year_data['Current Debt and Capital Lease Obligation'])
    }
    stock_data['(Year, n=10) versus (Year, n=5)'] = {
        'value': (company_data[all_years[-1]]['Diluted Net Income Available to Common Stockholders'] / company_data[all_years[0]]['Diluted Net Income Available to Common Stockholders']) - 1
    }

    # Financial Health:
    stock_data['Quick Ratio'] = {
        'value': (last_year_data['Cash, Cash Equivalents and Short Term Investments'] + last_year_data['Trade and Other Receivables, Current']) / last_year_data['Total Current Liabilities']
    }
    stock_data['Current Ratio'] = {
        'value': last_year_data['Total Current Assets'] / last_year_data['Total Current Liabilities']
    }
    stock_data['Interest Coverage'] = {
        'value': last_year_data['Total Operating Profit/Loss'] / last_year_data['Net Interest Income/Expense']
    }
    stock_data['Debt/ Equity'] = {
        'value': last_year_data['Total Liabilities'] / last_year_data['Total Equity']
    }

    # Profitability:
    stock_data['Return on Assets'] = {
        'value': last_year_data['Diluted Net Income Available to Common Stockholders'] /  last_year_data['Total Assets']
    }
    stock_data['Return on Equity'] = {
        'value': last_year_data['Diluted Net Income Available to Common Stockholders'] /  last_year_data['Total Equity']
    }
    stock_data['Return on Invested Capital'] = {
        'value': stock_data['ROIC']['value']
    }
    stock_data['Net Margin'] = {
        'value': stock_data['Net/sales']['value']
    }

    
    # -- Print output yoo --
    import pprint
    pp = pprint.PrettyPrinter(indent=4)
    pp.pprint(stock_data)
    pp.pprint(company_data)

    # Write data to excel
    red_cell = workbook.add_format({'bg_color': '#FFC7CE'})
    green_cell = workbook.add_format({'bg_color': '#90EE90'})

    worksheet.write(0, column_index, stock_data['name'])
    row_index = 1
    for category, key_names in excel_output.items():
        row_index += 1
        for key_name in key_names:
            if key_name in stock_data.keys():
                # check if its dict
                if type(stock_data[key_name]) is dict:
                    cell_value = stock_data[key_name]['value']
                    cell_value = prepare_numeric_value_format(cell_value)
                    # print colour cells according to condition
                    if stock_data[key_name].get('condition', None) == True:
                        worksheet.write(row_index, column_index, cell_value, green_cell)
                    elif stock_data[key_name].get('condition', None) == False:
                        worksheet.write(row_index, column_index, cell_value, red_cell)
                    else:
                        worksheet.write(row_index, column_index, cell_value)
                else:
                    cell_value = stock_data[key_name]
                    worksheet.write(row_index, column_index, stock_data[key_name])

            row_index += 1
        row_index += 1
    column_index += 1

workbook.close()



