import pandas as pd
import os
import datetime

pd.set_option('display.max_columns', None)#显示所有列
pd.set_option('display.max_rows', None)#显示所有行
pd.set_option('display.width',1000)
pd.set_option('display.unicode.ambiguous_as_wide', True)
pd.set_option('display.unicode.east_asian_width', True)


basedate=datetime.datetime(2020, 10, 31)
forcastdate=datetime.datetime(2020, 11, 10)
forcasbasedate=datetime.datetime(2020, 11, 30)
address='C:\\Users\\zyzse\\Desktop\\预测流动性'
filelist=os.listdir(address)
print(filelist)
for i in filelist:
    if '境内汇总数据' in i:
        alllist_file=address+'\\'+i
    elif '流动性报表底稿' in i:
        fmlist_file=address+'\\'+i
    elif 'base' in i:
        baselist_file= address + '\\' + i
    elif 'forcast' in i:
        forcastlist_file=address + '\\' + i

#算财务部的总数据
alllist=pd.read_excel(alllist_file,header=4,index_col=1)
asset_sum=alllist.loc[['1.资产总计','2.表外收入'],['次日','2日至7日','8日至30日','31日至90日']].sum().sum()
debt_sum=alllist.loc[['3.负债合计','4.表外支出'],['次日','2日至7日','8日至30日','31日至90日']].sum().sum()
deposit_def=alllist.loc[['3.2.2活期存放','3.5.2活期存款'],'次日'].sum()-alllist.loc[['7.附注：活期存款','17.附注：活期存放'],['次日','2日至7日','8日至30日','31日至90日']].sum().sum()
ratio_real=(asset_sum-debt_sum+deposit_def)/asset_sum
#算金融市场部提供的数据
fmlist=pd.read_excel(fmlist_file,sheet_name='底稿',header=1)
fmlist=fmlist[['大类','中类','期限','人民币（万元）']]
fmlist['sum']=fmlist.apply(lambda x :'Y' if (x['中类'] =='交易账户' or x['期限'] in ['2日至7日','8日至30日','31日至90日']) and x['期限']!='逾期'  else 'N',axis=1)
fmlist=fmlist[fmlist['sum']=='Y']

#print(fmlist.groupby(['大类','中类'],as_index=False).agg(sum))
fmlist=fmlist.groupby('大类',as_index=False).agg(sum)
print('金融市场部提供数据：')
print(fmlist)

#算金融市场部的日常数据（金融市场部提供数据会有差别，因为提供数据很多基于运营部的）

baselist=pd.read_excel(baselist_file)
baselist=baselist[['item0', 'item1', 'item2', 'item3', 'due_dt', 'bal']]
baselist['sum']=baselist.apply(lambda x: 'Y' if ((x['due_dt'] - basedate).days <= 90 or 'TPL' in str(x['item3']) or x['item1'] == '货币基金') and (not '基金' in str(x['item3'])) else 'N', axis=1)
baselist=baselist[baselist['sum'] == 'Y']
baselist=baselist.groupby(['item1', 'item0'], as_index=False).agg(sum)
print('金融市场部日常数据：')
print(baselist)
baselist.loc[baselist[baselist['item1'] == '债券基金'].index.tolist()[0], 'bal']= baselist[baselist['item1'].str.contains('定制型基金')]['bal'].sum() + baselist[baselist['item1'] == '债券基金']['bal'].sum()
baselist=baselist.append({'item1':'自营投资（非存单）','item0':'资产','bal':sum(baselist[baselist['item1'].isin(['公共事业债','国债','国有产业债','国际机构债','地方政府债','政策性银行','民营企业债','资产支持证券','金融行业债','非标资产'])]['bal'])},ignore_index=True)
baselist=baselist.copy()
baselist=baselist.applymap(lambda x: '发行同业存单' if x == '同业存单发行' else x)
baselist=baselist.applymap(lambda x: '投资同业存单' if x == '同业存单' else x)
baselist=baselist.applymap(lambda x: '资管计划' if x == '银登中心' else x)
df=pd.merge(fmlist, baselist, left_on='大类', right_on='item1', how='outer')
print('合并对比：')
print(df)

#计算新的数据
forcastlist=pd.read_excel(forcastlist_file)
forcastlist=forcastlist[['item0', 'item1', 'item2', 'item3', 'due_dt', 'bal']]
forcastlist['sum']=forcastlist.apply(lambda x: 'Y' if ((x['due_dt'] - forcastdate).days <= 90 or 'TPL' in str(x['item3']) or x['item1'] == '货币基金') and (not '基金' in str(x['item3'])) else 'N', axis=1)
forcastlist=forcastlist[forcastlist['sum'] == 'Y']
forcastlist=forcastlist.groupby(['item1', 'item0'], as_index=False).agg(sum)
print('金融市场部最新数据：')
print(forcastlist)
forcastlist.loc[forcastlist[forcastlist['item1'] == '债券基金'].index.tolist()[0], 'bal']= forcastlist[forcastlist['item1'].str.contains('定制型基金')]['bal'].sum() + forcastlist[forcastlist['item1'] == '债券基金']['bal'].sum()
forcastlist=forcastlist.append({'item1': '自营投资（非存单）', 'item0': '资产', 'bal':sum(forcastlist[forcastlist['item1'].isin(['公共事业债', '国债', '国有产业债', '国际机构债', '地方政府债', '政策性银行', '民营企业债', '资产支持证券', '金融行业债', '非标资产'])]['bal'])}, ignore_index=True)
print(forcastlist[forcastlist['item1'].isin(['公共事业债', '国债', '国有产业债', '国际机构债', '地方政府债', '政策性银行', '民营企业债', '资产支持证券', '金融行业债', '非标资产'])])
forcastlist=forcastlist.copy()
forcastlist=forcastlist.applymap(lambda x: '发行同业存单' if x == '同业存单发行' else x)
forcastlist=forcastlist.applymap(lambda x: '投资同业存单' if x == '同业存单' else x)
forcastlist=forcastlist.applymap(lambda x: '资管计划' if x == '银登中心' else x)
df=pd.merge(df,forcastlist,left_on='大类', right_on='item1', how='outer')









print(df)
df=df[['大类','item0_x','人民币（万元）','bal_x','bal_y']]
df=df[pd.isna(df['大类'])==False]
df=df.groupby('item0_x').agg(sum)
df['result']=-df['bal_x']+df['bal_y']
print(df)
ratio_forcast=((asset_sum+df['result']['资产'])-(debt_sum+(df['result']['负债']))+deposit_def)/(asset_sum+df['result']['资产'])

print('基准日实际：'+str(ratio_real))
print('预测日当日预测：'+str(ratio_forcast))

