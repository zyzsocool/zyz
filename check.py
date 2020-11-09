import pandas as pd
import os
from functools import reduce
address='./审批单'
outaddress='./结果.xlsx'

file=os.listdir(address)
df=pd.DataFrame()
file[0]=pd.read_excel(address+'\\'+file[0],header=2)
df=reduce(lambda x,y:pd.concat([x,pd.read_excel(address+'\\'+y,header=2)]),file)


df=df[['交易日','回购方向','交易对手','回购期限(天)','回购利率(%)','交易金额(元)','交易员','新阶段1','新阶段2','新阶段3','新阶段4','新阶段5','备注1']]
df['交易金额(元)']=df['交易金额(元)'].str.replace(',', '').astype(float)
group=df.groupby(['交易日','回购方向','交易对手','回购期限(天)','回购利率(%)'])
def onename(allname):
    namelist=[]
    for i in allname :
        for j in i:
            if j not in namelist :
                namelist.append(j)
    return namelist
def bigname(allname):
    if '莫东一' in allname:
        bigname='莫东一'
    elif '陈焕勋' in allname:
        bigname='陈焕勋'
    elif '洪俊聪' in allname:
        bigname='洪俊聪'
    elif '莫晓婷' in allname or '梁裕彬' in allname:
        bigname='莫晓婷or梁裕彬'
    return bigname
def result(money,bigname):
    okmoney_dict={'莫东一':2000000000,
                 '陈焕勋':2000000000,
                 '洪俊聪':1000000000,
                 '莫晓婷or梁裕彬':0}
    okmoney=okmoney_dict[bigname]
    if okmoney>=money:
        return '正常'
    else:
        return  '化整为零'
def beizhu(bei):
    beizhu='无'
    for i in bei:
        if '匿' in str(i):
            beizhu='存在匿名交易'
    return beizhu
result=group.apply(lambda x:pd.Series({'交易金额':x['交易金额(元)'].sum(),
                                       '交易员':onename([x['交易员']]),
                                       '审批人员':onename([x['新阶段1'],x['新阶段2'],x['新阶段3'],x['新阶段4'],x['新阶段5']]),
                                       '最终审批人员':bigname(onename([x['新阶段1'],x['新阶段2'],x['新阶段3'],x['新阶段4'],x['新阶段5']])),
                                       '是否超权限':result(x['交易金额(元)'].sum(),bigname(onename([x['新阶段1'],x['新阶段2'],x['新阶段3'],x['新阶段4'],x['新阶段5']]))),
                                       '备注':beizhu(x['备注1'])
                                       })).reset_index()


if '化整为零' in list(result['是否超权限']):
    print('存在化整为零')
else:
    print('不存在化整为零')

print('结果已经导出至'+outaddress)
result.to_excel(outaddress)