import os
import pandas as pd
import numpy as np
import datetime as dt
from glob import glob

pd.set_option("display.max_rows", 5)
tab_1 = glob('transactiondetail_*.csv')
tab_3 = pd.concat([pd.read_csv(f) for f in tab_1] ,sort = False)
tab_3.drop((['Transaction Detail','Unnamed: 3','Unnamed: 4']),axis=1,inplace=True)
tab_3.columns = tab_3.iloc[2]
tab_4 = tab_3['Result Description'].str.contains('Transaction Successful',na = False)|tab_3['Result Description'].str.contains('InProgress',na = False)
tab_4 = tab_3.loc[tab_4]
tab_5 = tab_4['Trans Type'].str.contains('Topup',na = False)|tab_4['Trans Type'].str.contains('Fund Transfer',na = False)
tab_5 = tab_4.loc[tab_5]
tab_5.to_csv('tab_5.csv',index = False)
tab_5 = pd.read_csv('tab_5.csv',index_col = 0)
print(tab_5)

# print('Hello World')

# import sys
# import os
# os.path.dirname(sys.executable)

# import csv
# import requests
# import kmlwriter
# import pprint