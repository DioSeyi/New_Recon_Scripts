import numpy as np
import pandas as pd
pd.set_option("display.max_rows", 5)

air = pd.read_excel('dnload.xlsx', skiprows = 9) #load with the first maybe 9 rows in e
glo = pd.read_excel('glotest.xlsx', skiprows = 1)

air = air.dropna(how='all') #drop empty rows


replacement_row = air.iloc[0]


air.columns = replacement_row


air1 = air.drop(air.index[0])


air1.drop(columns=['Sl. No.'], inplace = True)
glo.drop(columns=['Sr. No.'], inplace = True)

air1 = air1.iloc[:-1]


air1.to_excel('airagain.xlsx',index=False)
glo.to_excel('gloagain.xlsx',index=False)



