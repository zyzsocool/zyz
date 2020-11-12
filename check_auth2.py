import pandas as pd
import os
from functools import reduce

pd.set_option('display.max_columns', None)#显示所有列
#pd.set_option('display.max_rows', None)#显示所有行
pd.set_option('display.width',1000)
# pd.set_option('display.unicode.ambiguous_as_wide', True)
# pd.set_option('display.unicode.east_asian_width', True)


address='./comstar表'
outaddress='./结果.xlsx'
filenew=os.listdir(address+'/新版')
fileold=os.listdir(address+'/旧版')

filenew[0]=pd.read_excel(address+'/新版/'+filenew[0],header=1)[['交易日','交易方向','对手方','回购天数','回购利率(%)','交易金额(元)','审批人员','备注']]
dfnew=reduce(lambda x,y:pd.concat([x,pd.read_excel(address+'/新版/'+y,header=1)[['交易日','交易方向','对手方','回购天数','回购利率(%)','交易金额(元)','审批人员','备注']]]),filenew)


fileold[0]=pd.read_excel(address+'/旧版/'+fileold[0],header=1)[['交易日','回购方向','对手方','回购天数','回购利率(%)','交易金额(元)','审批人员','备注']]
dfold=reduce(lambda x,y:pd.concat([x,pd.read_excel(address+'/旧版/'+y,header=1)[['交易日','回购方向','对手方','回购天数','回购利率(%)','交易金额(元)','审批人员','备注']]]),fileold)

dfold.rename(columns={'回购方向':'交易方向'},inplace=True)
df=pd.concat([dfnew,dfold])
group=df.groupby(['交易日','对手方','回购天数','回购利率(%)'])
#print(df)
def bigname(x):
    x=list(x)
    name='无'
    if '莫东一' in x:
        name='莫东一'
    elif '陈焕勋' in x:
        name='陈焕勋'
    elif '洪俊聪' in x:
        name='洪俊聪'
    elif '莫晓婷' in x or '梁裕彬' in x:
        name='莫晓婷or梁裕彬'
    return name
def result(money,bigname):
    okmoney_dict={'莫东一':2000000000,
                 '陈焕勋':2000000000,
                 '洪俊聪':1000000000,
                 '莫晓婷or梁裕彬':0,
                  '无':0}
    okmoney = okmoney_dict[bigname]
    if okmoney >= money:
        return '正常'
    elif bigname!='无':
        return '化整为零或超权限'
    else:
        return '无审批人员'

def beizhu(bei):
    beizhu='无'
    for i in bei:
        if '匿' in str(i):
            beizhu='存在匿名交易'
    return beizhu


result=group.apply(lambda x:pd.Series({'交易金额':sum(x['交易金额(元)']),
                                       '审批人员':bigname(x['审批人员']),
                                       '是否超权限':result(sum(x['交易金额(元)']),bigname(x['审批人员'])),
                                       '备注':beizhu(x['备注'])})).reset_index()
result.to_excel(outaddress)
print('结果已经导出至'+outaddress)