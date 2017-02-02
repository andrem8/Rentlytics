
# coding: utf-8

# In[3]:


import xlwt
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np


# In[4]:

Onboarding_data = pd.read_csv('Onboarding Tasks by User - anonymized.csv')


# In[5]:

Onboarding_data.head()


# In[6]:

from xlsxwriter.utility import xl_rowcol_to_cell
User_data = pd.read_excel("User _ Company List - anonymized.xlsx")


# In[7]:

User_data.columns = [ 'Organization Name','First Name','Last Name','End User Guid']
User_data.head()


# In[8]:

all_data = pd.merge(User_data, Onboarding_data, on='End User Guid', how='left')


# In[9]:

pivot_data = pd.pivot_table(all_data,index=["Organization Name","End User Guid"],values=["Task ID"],aggfunc= 'count')


# In[10]:

new_data = pivot_data['Task ID'] / 11 * 100
pivot_data['Percent Completed'] = new_data


# In[11]:

pivot_data.head()


# In[12]:

pivot_data_filtered = pivot_data.ix[pivot_data['Task ID'] >= 1]
pivot_data_filtered.head()


# In[13]:

writer = pd.ExcelWriter('Training Completion Report.xlsx', engine='xlsxwriter')
pivot_data.to_excel(writer, sheet_name='Summary')
writer.save()
writer = pd.ExcelWriter('Training Completion Report for Started Training.xlsx', engine='xlsxwriter')
pivot_data_filtered.to_excel(writer, sheet_name='Summary')
writer.save()


# In[ ]:




# In[ ]:



