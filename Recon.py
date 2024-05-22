import os
import pandas as pd
import numpy as np
import datetime as dt
from glob import glob
pd.set_option("display.max_rows", 5)

df1 = pd.read_csv('df29.csv')
df1.rename({'Agent Code': 'Trans_Id_Ext_Ref'},inplace=True,axis=1)
df2 = pd.read_csv('df30.csv')
df2.rename({'Agent Code': 'Trans_Id_Ext_Ref'},inplace=True,axis=1)


# INNER with INDICATOR tells that rows occured in BOTH records only (Inner_Both)

dfsp = df1.merge(df2, indicator = True, how='inner', on = 'Trans_Id_Ext_Ref')
dfsp.rename({'Trans_Id_Ext_Ref':'External Reference'},inplace=True,axis=1)
dfsp['Trans Id'] = dfsp['External Reference']
co1_arrange = dfsp.pop('_merge')
dfsp.insert(13,'_merge',co1_arrange)
co2_arrange = dfsp.pop('Trans Id')
dfsp.insert(18,'Trans Id',co2_arrange)

writer_recon_discripances_seperate = pd.ExcelWriter('OneCard11 Discipances Reconcilisation.xlsx')
dfsp.to_excel(writer_recon_discripances_seperate, sheet_name = 'Inner_Both', index = False)
workbook = writer_recon_discripances_seperate.book
worksheet = writer_recon_discripances_seperate.sheets['Inner_Both']
writer_recon_discripances_seperate.close()

print('passss!')

# writer = pd.ExcelWriter('OneCard_1 Discipances Reconcilisation.xlsx')
# dfsp.to_excel(writer, sheet_name = 'Sheet1')#, index = False)
# workbook = writer.book
# worksheet = writer.sheets['Sheet1']
# writer.close()