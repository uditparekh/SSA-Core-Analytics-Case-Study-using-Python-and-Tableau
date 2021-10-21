import pandas as pd
import pyodbc
from spinner import Spinner
import numpy as np
from datetime import datetime
from time import strftime
from datetime import date
import time


s=Spinner()
s.start()

df2 = pd.read_excel('Core Analytics.xlsx', sheet_name='Opportunities Data', index=False)
pd.set_option('display.max_columns', None)


df1 = pd.read_excel('Core Analytics.xlsx', sheet_name='Rep Data', index=False)
pd.set_option('display.max_columns', None)



'''

                 What’s the total RSE by each region?

'''

df1["Period"] = pd.to_datetime(df1["Period"])

df1['Period']= df1['Period'].dt.strftime('%Y-%m')

table = pd.pivot_table(df1,
                       index=["Period" ],
                       columns=["Region"],
                       values=["RSE"],
                       aggfunc={'RSE': np.sum  },
                       dropna="False",
                       margins='True',
                       margins_name= "Grand Total",
                       fill_value= 0)

table= table.rename(columns={'RSE': ' ' })



'''

                 What’s the total RSE by each Segment?

'''


table1 = pd.pivot_table(df1,
                       index=["Period" ],
                       columns=["Segment"],
                       values=["RSE"],
                       aggfunc={'RSE': np.sum  },
                       dropna="False",
                       margins='True',
                       margins_name= "Grand Total",
                       fill_value= 0)

table1= table1.rename(columns={'RSE': ' ' })



'''

                 #What’s the average life cycle by each region?

'''

df2['Deal_Close_Date'] = pd.to_datetime(df2['Deal_Close_Date'], format = '%Y-%m-%d')
df2['Deal_Start_Date'] = pd.to_datetime(df2['Deal_Start_Date'], format = '%Y-%m-%d')

date1= df2["Deal_Close_Date"]

date2 = df2["Deal_Start_Date"]

df2['Length']= date1.sub(date2, axis=0)
df2['Length']=df2['Length'].dt.days
df2.Length.fillna(0, inplace=True)

table2 = pd.pivot_table(df2,
                       index=["Pipeline_Stage" ],
                       columns=["Region"],
                       values=["Length"],
                       aggfunc={'Length': np.mean  },
                       dropna="False",
                       margins='True',
                       margins_name= "Grand Total")
                       #fill_value= 0)



'''

     #What’s the average deal size by each region? Does classifying them into small/big or any more categories help us find better trends?

'''


pd.to_numeric(df2['eARR'])

table3 = pd.pivot_table(df2,
                       index=["Pipeline_Stage" ],
                       columns=["Region"],
                       values=["eARR"],
                       aggfunc={'eARR': np.mean  },
                       dropna="False",
                       margins='True',
                       margins_name= "Grand Total")
                       #fill_value= 0)

df2['AVG_eARR'] = df2.groupby(['Region']).eARR.transform('mean')

pd.to_numeric(df2['AVG_eARR'])



def func(x):
    if x <=  40000.00:
        return "Deal_Size: Small"
    elif x >=  75000.00:
        return "Deal_Size: Big"
    else:
        return 'Deal_Size: Medium'

def func1(x):
    if x <=  50000.00:
        return "Deal_Size: Small"
    else:
        return 'Deal_Size: Big'

df2['Category'] = df2['AVG_eARR'].apply(func1)

table4 = pd.pivot_table(df2,
                       index=["Region","Pipeline_Stage" ],
                       columns=["Category"],
                       values=["eARR"],
                       aggfunc={'eARR': np.mean  },
                       dropna="False",
                       margins='True',
                       margins_name= "Grand Total")
                       #fill_value= 0)



'''

                 #What’s the conversion rate by each region?

'''

df3 = df2[(df2["Pipeline_Stage"] == "Closed")]

table5 = pd.pivot_table(df3,
                       index=["Pipeline_Stage" ],
                       columns=["Region"],
                       values=["Opportunity_ID"],
                       aggfunc={'Opportunity_ID': len  },
                       dropna="False",
                       margins='True',
                       margins_name= "Grand Total")

table5=table5[:-1]

table5= table5.rename(index={'Closed': 'Grand Total' })

df4 = df2[(df2["Status"] == "SQL Accepted")]

table6 = pd.pivot_table(df4,
                       index=["Status" ],
                       columns=["Region"],
                       values=["Opportunity_ID"],
                       aggfunc={'Opportunity_ID': len  },
                       dropna="False",
                       margins='True',
                       margins_name= "Grand Total")

table6=table6[:-1]

table6= table6.rename(index={'SQL Accepted': 'Grand Total' })

df5 = table5.div(table6)



'''

                 #What’s the conversion rate by each segment?

'''


table7 = pd.pivot_table(df3,
                       index=["Pipeline_Stage" ],
                       columns=["Segment"],
                       values=["Opportunity_ID"],
                       aggfunc={'Opportunity_ID': len  },
                       dropna="False",
                       margins='True',
                       margins_name= "Grand Total")

table7=table7[:-1]

table7= table7.rename(index={'Closed': 'Grand Total' })


table8 = pd.pivot_table(df4,
                       index=["Status" ],
                       columns=["Segment"],
                       values=["Opportunity_ID"],
                       aggfunc={'Opportunity_ID': len  },
                       dropna="False",
                       margins='True',
                       margins_name= "Grand Total")

table8=table8[:-1]

table8= table8.rename(index={'SQL Accepted': 'Grand Total' })

df6 = table7.div(table8)



writer = pd.ExcelWriter('Core_Analytics.xlsx')
df1.to_excel(writer, sheet_name='Sheet1', index=False)
df2.to_excel(writer, sheet_name='Sheet2', index=False)
table.to_excel(writer, sheet_name='RSE by each region', index=True)
table1.to_excel(writer, sheet_name='RSE by each Segment', index=True)
table2.to_excel(writer, sheet_name='Avg life cycle by region', index=True)
table3.to_excel(writer, sheet_name='Avg deal size by region', index=True)
table4.to_excel(writer, sheet_name='Avg deal size by Category', index=True)
df5.to_excel(writer, sheet_name='conversion rate by region', index=True)
df6.to_excel(writer, sheet_name='conversion rate by segment', index=True)
writer.save()

s.stop()
