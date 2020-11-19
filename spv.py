import pandas as pd
import os
address=r'C:\Users\zyzse\Desktop\增速'
today='20201118'
lastmonth='20201031'
lastyear='20191231'



def calspv(date):
    addressin=os.path.join(address,date+'.xlsx')
    df=pd.read_excel(addressin)
    df1=df.groupby('item3')['bal'].agg(sum).reset_index()
    df11=df1[df1['item3'].str.contains('定制型基金')].copy()
    df11.rename(columns={'item3':'资产'},inplace=True)
    df12=df1[df1['item3'].str.contains('资管计划-融通基金')].copy()
    df12.rename(columns={'item3': '资产'},inplace=True)
    df2=df.groupby('item1')['bal'].agg(sum).reset_index()
    df2=df2[df2['item1'].isin(['银登中心','非标资产','同业理财','债券基金','货币基金','信托计划'])].copy()
    df2.rename(columns={'item1': '资产'}, inplace=True)
    result=pd.concat([df11,df12,df2])
    result=result.reset_index()
    result=result.drop('index',axis=1)
    print(date)
    print(result)
    resultsum=result['bal'].sum()
    print(resultsum)
    addressout = os.path.join(address, 'result'+date + '.xlsx')
    result.to_excel(addressout)

    return [result,resultsum]
spvtoday=calspv(today)[1]
spvlasymonth=calspv(lastmonth)[1]
spvlastyear=calspv(lastyear)[1]

spvtoday
spvlasymonth+=132040.34
spvlastyear+=245965.54

growthlastmonth=spvlasymonth/spvlastyear-1
growthtoday=spvtoday/spvlastyear-1
print(growthlastmonth)
print(growthtoday)