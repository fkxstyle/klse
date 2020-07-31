import os
from datetime import datetime
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
from utils import string_value_converter 


# define year
now = datetime.now()
this_year = str(now.year)
last_year = str(now.year - 1)
last_2_year = str(now.year - 2)

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

# initialize excel writer engine
writer = pd.ExcelWriter('financial_data.xlsx', engine='xlsxwriter')

# prepare dataframe for data availability
company_data_availability = pd.DataFrame(index=indexes, columns=company_list)
for company in company_list:
    company_dir = os.path.join(company_folder_root, company)
    document_list = os.listdir(company_dir)
    for document in document_list:
        document_dir =  os.path.join(company_dir,document)
        if 'Balance Sheet' in document:
            df_bs = pd.read_excel(document_dir,sheet_name='sheet1').fillna(0)
            df_bs['Name'] = df_bs['Name'].str.lstrip().str.rstrip()
            df_bs = df_bs.set_index('Name')
            df_bs_column = pd.DataFrame(index=balance_sheet_columns)
            df_bs_column[company] = df_bs_column.index.map(lambda x: x in df_bs.index)
        elif 'Cash Flow' in document:
            df_cf = pd.read_excel(document_dir,sheet_name='sheet1').fillna(0)
            df_cf['Name'] = df_cf['Name'].str.lstrip().str.rstrip()
            df_cf = df_cf.set_index('Name')
            df_cf_column = pd.DataFrame(index=cash_flow_columns)
            df_cf_column[company] = df_cf_column.index.map(lambda x: x in df_cf.index) 
        elif 'Income Statement' in document:
            df_is = pd.read_excel(document_dir,sheet_name='sheet1').fillna(0)
            df_is['Name'] = df_is['Name'].str.lstrip().str.rstrip()
            df_is = df_is.set_index('Name')
            df_is_column = pd.DataFrame(index=income_statement_columns)
            df_is_column[company] = df_is_column.index.map(lambda x: x in df_is.index)
    company_data_availability[company] = pd.concat([df_bs_column, df_cf_column, df_is_column])
company_data_availability = company_data_availability.reindex(indexes)

company_data_availability.to_excel(writer, sheet_name='data')

# prepare dataframe for data of each company in the folder
for company in company_list:
    company_dir = os.path.join(company_folder_root, company)
    document_list = os.listdir(company_dir)

    # load data.json to stock_data
    with open(company_dir + '/data.json') as json_file:
        stock_data = json.load(json_file)

    # build data frame from excel file
    for document in document_list:
        document_dir =  os.path.join(company_dir,document)
        if 'Balance Sheet' in document:
            df_bs = pd.read_excel(document_dir,sheet_name='sheet1').fillna(0)
            df_bs['Name'] = df_bs['Name'].str.lstrip().str.rstrip()
            df_bs = df_bs.set_index('Name')
            df_bs_filtered = df_bs.reindex(balance_sheet_columns).loc[balance_sheet_columns].fillna(0)

        elif 'Cash Flow' in document:
            df_cf = pd.read_excel(document_dir,sheet_name='sheet1').fillna(0)
            df_cf['Name'] = df_cf['Name'].str.lstrip().str.rstrip()
            df_cf = df_cf.set_index('Name')
            df_cf_filtered = df_cf.reindex(cash_flow_columns).loc[cash_flow_columns].fillna(0)
                    
        elif 'Income Statement' in document:
            df_is = pd.read_excel(document_dir,sheet_name='sheet1').fillna(0)
            df_is['Name'] = df_is['Name'].str.lstrip().str.rstrip()
            df_is = df_is.set_index('Name')
            df_is_filtered = df_is.reindex(income_statement_columns).loc[income_statement_columns].fillna(0)
        
    company_data = pd.concat([df_bs_filtered, df_cf_filtered, df_is_filtered], sort=False)
    company_data = company_data.drop(columns='TTM')
    company_data = company_data.reindex(indexes)

    # clean up data
    company_data = company_data.apply(lambda x: x.str.replace(',','')).fillna(0)
    company_data = company_data.apply(pd.to_numeric)
    last_year_data = company_data[last_year]
    last_2_year_data = company_data[last_2_year]

    first_three_year = sorted(company_data.columns)[:3]
    last_three_year = sorted(company_data.columns)[-3:]
    all_years = company_data.columns

    # feed data to json
    stock_data['EBITDA'] = last_year_data['Pretax Income'] + last_year_data['Net Interest Income/Expense'] + last_year_data['Depreciation, Amortization and Depletion, Supplemental']
    stock_data['Book Value Per Share'] =  last_year_data['Total Equity'] / last_year_data['Common Shares Outstanding']
    stock_data['Net Tangible Book Value Per Share'] =  (last_year_data['Total Equity'] -  last_year_data['Net Intangible Assets']) / last_year_data['Common Shares Outstanding']
    stock_data['Working Capital'] =  last_year_data['Total Current Assets'] - last_year_data['Total Current Liabilities']
    stock_data['Total Debt'] =  last_year_data['Current Debt and Capital Lease Obligation'] + last_year_data['Long Term Debt and Capital Lease Obligation']
    stock_data['Dividend Per Share'] =  last_year_data['Cash Dividends Paid'] / last_year_data['Common Shares Outstanding']
    stock_data['Earning Per Share'] =  last_year_data['Diluted Net Income Available to Common Stockholders'] / last_year_data['Common Shares Outstanding']
    stock_data['Interest Bearing Debt'] =  stock_data['Total Debt']
    stock_data['Tax Rate'] =  last_year_data['Provision for Income Tax'] / last_year_data['Pretax Income']
    stock_data['NOPAT'] =  last_year_data['Total Operating Profit/Loss'] * (1 - stock_data['Tax Rate'])
    stock_data['Invested Capital'] =  stock_data['Interest Bearing Debt'] + last_year_data['Total Equity'] + last_year_data['Non-Controlling/Minority Interest'] - last_year_data['Cash, Cash Equivalents and Short Term Investments'] - last_year_data['Goodwill']
    stock_data['ROIC'] =  stock_data['NOPAT'] /  stock_data['Invested Capital']
    stock_data['Adequate size of enterprise (Market Cap) (RM mil)'] = string_value_converter(stock_data['market_cap'].replace(",", ""))
    stock_data['2 x Current asset >= Current liabilities'] = last_year_data['Total Current Assets'] / last_year_data['Total Current Liabilities'] > 2
    stock_data['Net current asset (working capital) > long term debt'] = stock_data['Working Capital'] > last_year_data['Long Term Debt and Capital Lease Obligation']
    stock_data['Debt < 2 x Stock Equity (book value)'] = last_year_data['Total Liabilities'] < 2 * last_year_data['Total Equity']
    # stock_data['Earning stability in the past 10 years (EPS)']
    stock_data['Dividend uninterupted for past 20 years'] = (company_data.loc['Cash Dividends Paid'] != 0 ).all()
    stock_data['Min. inc of 1/3 using three year averages at beginning & end'] = abs(company_data.loc['Cash Dividends Paid', last_three_year].sum()) / abs(company_data.loc['Cash Dividends Paid', first_three_year].sum()) - 1 > 1/3
    stock_data['P/E < 15 average earnings of past three years'] = (float(stock_data['current_pe_ratio'])+float(stock_data['last_year_pe_ratio'])+float(stock_data['last_2_years_pe_ratio'])) / 3 < 15
    stock_data['Price < 1.5 x Book value'] = float(stock_data['price']) < 1.5*float(stock_data['Book Value Per Share'])
    stock_data['P/E * P/NTA < 22.5'] = float(stock_data['current_pe_ratio']) * float(stock_data['price']) / stock_data['Net Tangible Book Value Per Share'] < 22.5



    # Ratios:
    stock_data['P/book value'] = float(stock_data['price']) / stock_data['Book Value Per Share']
    stock_data['Net/sales'] = last_year_data['Diluted Net Income Available to Common Stockholders'] / last_year_data['Total Revenue']
    stock_data['Earning/book value'] = last_year_data['Diluted Net Income Available to Common Stockholders'] / stock_data['Net Tangible Book Value Per Share']
    stock_data['Working capital/ debt'] = stock_data['Working Capital'] / ( last_year_data['Long Term Debt and Capital Lease Obligation'] + last_year_data['Current Debt and Capital Lease Obligation'])
    stock_data['(Year, n=10) versus (Year, n=5)'] = (company_data[all_years[-1]]['Diluted Net Income Available to Common Stockholders'] / company_data[all_years[0]]['Diluted Net Income Available to Common Stockholders']) - 1

    # Financial Health:
    stock_data['Quick Ratio'] = (last_year_data['Cash, Cash Equivalents and Short Term Investments'] + last_year_data['Trade and Other Receivables, Current']) / last_year_data['Total Current Liabilities']
    stock_data['Interest Coverage'] = last_year_data['Net Interest Income/Expense'] / stock_data['EBITDA']
    stock_data['Debt/ Equity'] = last_year_data['Total Liabilities'] / last_year_data['Total Equity']

    # Profitability:
    stock_data['Return on Assets'] = last_year_data['Diluted Net Income Available to Common Stockholders'] /  last_year_data['Total Assets']
    stock_data['Return on Equity'] = last_year_data['Diluted Net Income Available to Common Stockholders'] /  last_year_data['Total Equity']
    
    # Operating Performance:
    stock_data['Day Sales Outstanding'] = last_year_data['Trade/Accounts Receivable, Current'] / last_year_data['Total Revenue'] * 365
    # stock_data['Days Inventory'] = ( (last_year_data['Inventories']  + last_2_year_data['Inventories']) / 2) / Cost of Revenues * 365
    # stock_data['Days Payables'] = last_year_data['Trade/Accounts Payable, Current'] * 365 DIVIDED BY Cost of Revenues
    stock_data['Receivable Turnover'] = last_year_data['Total Revenue'] / ((last_year_data['Trade/Accounts Receivable, Current'] + last_2_year_data['Trade/Accounts Receivable, Current']) / 2)
    stock_data['Inventory Turnover'] = last_year_data['Total Revenue'] / ((last_year_data['Inventories'] + last_2_year_data['Inventories']) / 2) 
    # stock_data['Fixed Asset Turnover'] = last_year_data['Diluted Net Income Available to Common Stockholders'] / [(Net Property, Plant and Equipment (nth year) + Net Property, Plant and Equipment (n-1th year)]/2)
    stock_data['Total Asset Turnover'] = last_year_data['Total Revenue'] / ((last_year_data['Total Assets'] + last_2_year_data['Total Assets']) / 2)

    import pprint
    pp = pprint.PrettyPrinter(indent=4)
    pp.pprint(stock_data)
    pp.pprint(company_data)
#     company_data.to_excel(writer, sheet_name=company)

# writer.save()

