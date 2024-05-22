#!/usr/bin/env python
# coding: utf-8

# # The Tutorial for GOMYCODE

# In[1]:


import numpy as np
import pandas as pd


# In[2]:


cd Downloads


# In[14]:


pwd!


# In[15]:


file = pd.read_csv('sales_data.csv')


# In[16]:


file


# In[18]:


type(file)


# In[27]:


file.iloc[0:2]


# In[28]:


file.shape


# In[32]:


file[['order_id','product','category']]


# In[33]:


file['price'] == 1000


# In[34]:


file[file['price'] == 1000]


# In[36]:


file['price'].mean()


# In[39]:


nig = pd.read_csv('Nigeria economy kpis.csv')


# In[42]:


nig


# In[41]:


nig.info


# In[43]:


nig['Unemployment'].mean()


# In[48]:


nig.columns


# In[49]:


nig.rename(columns={'Year':'yearCap','Unemployment':'UnemploymentRate'})


# In[105]:


glo = pd.read_excel('glotest.xlsx')


# In[106]:


glo


# In[107]:


glo.rename(columns={} nig.rename(columns={'Year':'yearCap','Unemployment':'UnemploymentRate'})


# In[108]:


replacement_row = glo.iloc[0]


# In[109]:


glo.columns = replacement_row


# In[110]:


glo


# In[111]:


glo = glo.drop(glo.index[0]) # to delete a row


# In[112]:


glo.columns


# In[113]:


glo = glo.drop(columns=['Sr. No.']) # to delete a column 


# In[114]:


glo


# In[115]:


glo[['Date and Time']] = glo['Date and Time'].str.split(' ', expand=True)


# In[76]:


gl = [glo.split('Date and Time') for glo in glo]


# In[80]:


gl.glo


# In[116]:


dates = [split[0] for split in gl]
times = [split[1] for split in gl]


# In[117]:


glo


# In[121]:


glo['Date and Time'] = pd.to_datetime(glo['Date and Time'])


# In[122]:


glo['Date'] = glo['Date and Time'].dt.date
glo['Time'] = glo['Date and Time'].dt.time


# In[123]:


glo


# In[94]:


glo = glo.drop(columns=['Date and Time'])


# In[95]:


glo


# In[ ]:




