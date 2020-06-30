import os
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

# define columns for each document
balance_sheet_columns = [
    'Total Current Assets',
    'Cash, Cash Equivalents and Short Term Investments',
    'Net Intangible Assets',
    'Total Current Liabilities',
    'Current Debt and Capital Lease Obligation',
    'Long Term Debt and Capital Lease Obligation',
    'Total Equity',
    'Common Shares Outstanding'
]
cash_flow_columns = [
    'Purchase of Property, Plant and Equipment',
    'Cash Dividends Paid'
]
income_statement_columns = [
    'Total Revenue',
    'Diluted Net Income Available to Common Stockholders'
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
    company_data.to_excel(writer, sheet_name=company)

writer.save()

