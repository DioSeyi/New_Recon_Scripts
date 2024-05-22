import numpy as np
import pandas as pd
pd.set_option("display.max_rows", 5)


# This remain same 
exclude_columns = ['Consider As( Credit/ Debit/ None)','User Commission( In%)','Commission Type','Requested For Refund','User Type','Api Transaction Id','Transaction Type','Transaction Date','Commission In','APICommission( In%)','Response','Recharge Refund Date','APICommission Type','APICommission In','Callback Response']
a_admin = pd.read_csv('export_2024-05-20.csv',usecols = lambda column: column not in exclude_columns)
a_admin.rename({'Transaction Unique Id': 'UniqueID'}, inplace=True, axis=1)
a_admin['UniqueID'] = a_admin['UniqueID'].str.replace("'", '')
a_admin['UniqueID'] = pd.to_numeric(a_admin['UniqueID'])

# This is utter for service provider
a_direct = pd.read_excel('c2sTransferChannelUserNew(161).xlsx', skiprows = 9) #load with the first maybe 9 rows in exclusion
a_direct = a_direct.dropna(how='all') #drop empty rows
replacement_row = a_direct.iloc[0] # to pick/select the row of interect 
a_direct.columns = replacement_row # to replace a header by another row
a_direct1 = a_direct.drop(a_direct.index[0])
a_direct1.drop(columns=['Sl. No.'], inplace = True)
a_direct1 = a_direct1.iloc[:-1] # to delete the last row
a_direct1.rename({'Reference ID': 'UniqueID'},inplace=True,axis=1)
a_direct1['UniqueID'] = pd.to_numeric(a_direct1['UniqueID'])


Airtel_recon = a_direct1.merge(a_admin, indicator = True, how='outer', on = 'UniqueID')

Airtel_recon['Reference ID'] = Airtel_recon['UniqueID']

Airtel_recon.rename({'UniqueID': 'Transaction Unique Id'},inplace=True,axis=1)
tranID_column = Airtel_recon.pop('Reference ID')  # Remove the 'City' column from the DataFrame
Airtel_recon.insert(28, 'Reference ID', tranID_column) # and place it in line 15

writer_recon_discripances_seperate = pd.ExcelWriter('Airtel_Airtime_Recon.xlsx')
Airtel_recon.to_excel(writer_recon_discripances_seperate, sheet_name = 'Sheet1', index = False)
workbook = writer_recon_discripances_seperate.book
worksheet = writer_recon_discripances_seperate.sheets['Sheet1']
writer_recon_discripances_seperate.close()