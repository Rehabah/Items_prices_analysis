# -*- coding: utf-8 -*-
"""
Created on Thu Jul 21 10:53:35 2022

@author: admin
"""


import pandas as pd
import os, glob

path = r"C:\Users\win\Desktop\avg_prices"
df1 = pd.DataFrame()
df2 = pd.DataFrame()

all_files = glob.glob(os.path.join(path, "*.xlsx"))
for f in all_files:
    #print(f)
    data = pd.read_excel(f, skiprows=[0,1],usecols=[1,11,13,18,24,26])
    df=data.dropna(how='all')
    print(f)
    df.columns

    df.rename(columns = {df.columns[0]:'items', df.columns[1]:'units',
                      df.columns[2]:'averg_pr1',df.columns[3]:'sub_cat',
                      df.columns[4]:'unit_ar',
                     df.columns[5]:'items_ar'  }, inplace = True)   
    df.columns
    df.info()
    df['month']=df['averg_pr1'][4]
    df['year']=df['averg_pr1'][5]
    df=df.replace('nan',float('Nan'),regex=True)
    df['sub_cat']=df['sub_cat'].ffill()#.replace('nan',method='ffill')
    df.columns

    df['units1']=df['units'].astype(str)
    df.columns
    df1 = df1.append(df)
    
df2=df1[(df1.units.notnull())]
df2=df2[~df2.units1.str.contains('Unit')]
df2=df2.replace({'month' : { 'May' :'05-May', 'Apr' : '04-Apr', 'Aug' : '08-Aug',
                        'Dec':'12-Dec','Feb':'02-Feb','Jan':'01-Jan','Jul':'07-Jul',
                        'Jun':'06-Jun','Mar':'03-Mar','Nov':'11-Nov',
                        'Oct':'10-Oct','Sep':'09-Sep'}})
    
df2=df2[~df2.sub_cat.str.contains('التبغ | الملابس|المنظفات|الصحة')]

df2.to_excel('final_price_avg_data.xlsx',encoding='utf-8',index=True)

