import os
import pandas as pd
pd.set_option("display.max_rows", 5)

# import numpy as np
# import datetime as dt
# from glob import glob
# from openpyxl import workbook,load_workbook
# from openpyxl.utils import get_column_letter
# from openpyxl.styles import Font,PatternFill,Border,Side,Alignment,fills


# This remain same 
exclude_columns = ['Consider As( Credit/ Debit/ None)','User Commission( In%)','Commission Type','Requested For Refund','User Type','Api Transaction Id','Transaction Type','Transaction Date','Commission In','APICommission( In%)','Response','Recharge Refund Date','APICommission Type','APICommission In','Callback Response']
m_admin = pd.read_csv('export_2024-05-17.csv',usecols = lambda column: column not in exclude_columns)
m_admin.rename({'Transaction Unique Id': 'UniqueID'}, inplace=True, axis=1)
m_admin['UniqueID'] = m_admin['UniqueID'].str.replace("'", '')
m_admin['UniqueID'] = pd.to_numeric(m_admin['UniqueID'])

# This i utter for service provider
m_direct = pd.read_excel('total_transaction_agent - 2024-05-17T100247.540.xls', skiprows = (0,1,2))
m_direct2 = (m_direct['TransactionProfile'] == 'PREPAID_TOPUP') | (m_direct['TransactionProfile'] == 'DATA_BUNDLE') #| (tab_9['Trans Type'] == 'Wallet Transfer') | (tab_9['Trans Type'] == 'Wallet Topup') | (tab_9['Trans Type'] == 'PINREDEEM') # To disregards any other top inclusive like 'Wallet Topup'
m_direct2 = m_direct.loc[m_direct2]
m_direct2.rename({'ClientReference': 'UniqueID'},inplace=True,axis=1)
m_direct2['UniqueID'] = pd.to_numeric(m_direct2['UniqueID'])
# m_direct1 = m_direct1.dropna()
m_direct2['Transaction Unique Id'] = pd.to_numeric(m_direct2['UniqueID'])

# Outer Recon takes place
mtn_recon = m_admin.merge(m_direct2, indicator = True, how='outer', on = 'UniqueID')

mtn_recon['Transaction Unique Id'] = mtn_recon['UniqueID']

mtn_recon.rename({'UniqueID': 'ClientReference'},inplace=True,axis=1)
tranID_column = mtn_recon.pop('Transaction Unique Id')  # Remove the 'City' column from the DataFrame
mtn_recon.insert(15, 'Transaction Unique Id', tranID_column)

writer_recon_discripances_seperate = pd.ExcelWriter('MTN_Airtime_Recon.xlsx')
mtn_recon.to_excel(writer_recon_discripances_seperate, sheet_name = 'Sheet1', index = False)
workbook = writer_recon_discripances_seperate.book
worksheet = writer_recon_discripances_seperate.sheets['Sheet1']
writer_recon_discripances_seperate.close()