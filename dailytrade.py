import pandas as pd
import os
from openpyxl import load_workbook
pd.set_option('display.max_columns', None)#显示所有列
pd.set_option('display.max_rows', None)#显示所有行
pd.set_option('display.width',1000)
pd.set_option('display.unicode.ambiguous_as_wide', True)
pd.set_option('display.unicode.east_asian_width', True)


#day = input('输入交易日')
#lastday = input('输入上个交易日')
day='20201106'
lastday='20201105'
address='E://衍生品日报'
address1 = address + '//' + day
address2 = address + '//' + lastday
outaddress=address1+'//衍生品日报'+day+'.xlsx'
file1 = os.listdir(address1)
file2 = os.listdir(address2)

for i in file1:
    if '现券' in i:
        df_bond = pd.read_excel(address1 + '//' + i, header=1)
    elif 'irs' in i:
        df_irs = pd.read_excel(address1 + '//' + i, header=1)
    elif '即期投组交易查询与维护' in i:
        df_spot = pd.read_excel(address1 + '//' + i, header=1)
    elif '远期投组交易查询与维护' in i:
        df_forward = pd.read_excel(address1 + '//' + i, header=1)
    elif '掉期投组交易查询与维护' in i:
        df_swap = pd.read_excel(address1 + '//' + i, header=1)
    elif '期权投组交易查询与维护' in i:
        df_option = pd.read_excel(address1 + '//' + i, header=1)
    elif '拆借投组交易查询与维护' in i:
        df_borrow = pd.read_excel(address1 + '//' + i, header=1)
    elif '收益风险评估' in i:
        df_dvtoday = pd.read_excel(address1 + '//' + i, header=0)
        df_dvtoday.rename(columns={'Unnamed: 2':'成交编号/债券'},inplace=True)
for i in file2:
    if '收益风险评估' in i:
        df_dvlastday = pd.read_excel(address2 + '//' + i, header=0)
        df_dvlastday.rename(columns={'Unnamed: 2':'成交编号/债券'},inplace=True)

df_irs=pd.merge(df_irs,df_dvtoday,how='left',left_on='成交编号',right_on='成交编号/债券').copy()
df_irs=df_irs[['投资组合','交易方向','名义本金(万)','期限','到期日_x','本方收取_固定利率(%)','本方支付_固定利率(%)','成交编号','PVBP']].copy()
df_irs['交易方向']=df_irs['交易方向'].apply(lambda x:'BUY' if x=='收浮动 付固定'  else 'SELL')
df_irs['期限']=df_irs['期限']/365
df_irs['PVBP']=df_irs['PVBP']/10000
df_irs.insert(1,'交易工具','IRS')
df_irs.insert(5,'利率/汇率',df_irs.apply(lambda x:x['本方收取_固定利率(%)'] if type(x['本方收取_固定利率(%)'])==float else x['本方支付_固定利率(%)'],axis=1))
df_irs.insert(6,'近端日/行权日','-')
df_irs=df_irs.drop(labels=['本方收取_固定利率(%)','本方支付_固定利率(%)','成交编号'],axis=1).copy()
df_irs['交易类型']=df_irs['期限'].apply(lambda x: 'FR007S5Y' if x>4.8 else 'FR007S1Y')
df_irs.rename(columns={'名义本金(万)':'面额/名义本金','期限':'剩余期限','到期日_x':'远端日/交割日/到期日'},inplace=True)

print(df_irs)

df_bond_buy=df_bond[df_bond['交易方向']=='买入']
df_bond_sell=df_bond[df_bond['交易方向']=='卖出']
df_bond_buy=pd.merge(df_bond_buy,df_dvtoday,how='left',left_on=['投资组合','债券'],right_on=['交易投组','成交编号/债券']).copy()
df_bond_sell=pd.merge(df_bond_sell,df_dvlastday,how='left',left_on=['投资组合','债券'],right_on=['交易投组','成交编号/债券']).copy()
df_bond=pd.concat([df_bond_buy,df_bond_sell],sort=False).copy()
df_bond=df_bond[['投资组合','债券名称','交易方向','券面总额(万)','待偿期_y','到期收益率(%)','到期日','PVBP','债券类别_x']].copy()
df_bond.insert(6,'近端日/行权日','-')
df_bond['PVBP']=df_bond['PVBP']/10000
df_bond.rename(columns={'债券名称':'交易工具','券面总额(万)':'面额/名义本金','待偿期_y':'剩余期限','到期日':'远端日/交割日/到期日','到期收益率(%)':'利率/汇率','债券类别_x':'交易类型'},inplace=True)
df_bond['交易方向']=df_bond['交易方向'].apply(lambda x:'BUY' if x=='买入' else 'SELL' )
print(df_bond)

df_account=pd.DataFrame({'交易类型':['FR007S5Y','FR007S5Y','FR007S1Y','FR007S1Y','国债','国债','证金债','证金债','其他债','其他债'],
                         '交易方向':['BUY','SELL','BUY','SELL','BUY','SELL','BUY','SELL','BUY','SELL']})

df_accountA=df_account.copy()
df_irs.groupby(['交易类型'])


print(df_account)




df_trade=pd.concat([df_irs,df_bond],sort=False).iloc[:,:8].copy().reset_index()
df_trade=df_trade.drop(labels='index',axis=1)
print(df_trade)
# book=load_workbook(outaddress)
# writer=pd.ExcelWriter(outaddress,engine='openpyxl')
# writer.book=book
# writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
# df_trade.to_excel(writer,sheet_name='交易记录',startcol=12,startrow=1)
# writer.save()