import numpy as np
import pandas as pd
import os
import sys
import datetime
import dateutil
import glob
def 支付周期映射(x):
    if x=='一次性':
        return 0
    elif x=='一个月':
        return 1
    elif x=='二个月':
        return 2
    elif x=='三个月':
        return 3
    elif x=='四个月':
        return 4
    elif x=='五个月':
        return 5
    elif x=='半年':
        return 6
    elif x=='七个月':
        return 7
    elif x=='八个月':
        return 8
    elif x=='九个月':
        return 9
    elif x=='十个月':
        return 10
    elif x=='十一个月':
        return 11
    elif x=='一年':
        return 12
    elif x=='一年半':
        return 18
    elif x=='两年':
        return 24
    elif x=='三年':
        return 36
    elif x=='四年':
        return 48
    elif x=='五年':
        return 60



wbaozhangdian=glob.glob(r'原始数据\报账点缴费台帐-*.csv')
wbaozhangdian.sort(reverse=True)
wbaozhangdian=wbaozhangdian[0]
wbaozhangdianname = os.path.split(wbaozhangdian)[1]
Ht=pd.read_excel('已排除的合同或报账点.xlsx',sheet_name='已排除的合同')
Zd=pd.read_excel('已排除的合同或报账点.xlsx',sheet_name='已排除的报帐点')
try:
    print('正在处理：'+wbaozhangdianname)
    ex=pd.read_csv(wbaozhangdian)
    ex.drop(ex[ex['供电方式']=='直供电'].index,inplace=True)
    ex.drop(ex[ex['供应商名称']=='中国铁塔股份有限公司宜昌市分公司'].index,inplace=True)
    ex.drop(ex[ex['合同编号'].apply(lambda x:x in list(Ht['合同编号']))].index,inplace=True)
    ex.drop(ex[ex['报帐点编码'].apply(lambda x:x in list(Zd['报帐点编码']))].index,inplace=True)
    ex['支付周期']=ex['支付周期'].apply(支付周期映射)
    sheet1=ex.groupby(['报帐点编码','合同编号'])   
    sheet1=sheet1[['缴费期终','合同结束时间']].max()
    sheet2=ex.groupby('合同编号')
    sheet2=sheet2.支付周期.min()
    sheet3=sheet1.join(sheet2,how='left',on='合同编号')
    sheet3.drop(sheet3[sheet3['支付周期']==0].index,inplace=True)
    sheet3['缴费期终']=pd.to_datetime(sheet3['缴费期终'])
    sheet3['距离下次付款时间（天）']=sheet3.apply(lambda x:(pd.tseries.offsets.shift_month(x['缴费期终'],x['支付周期'])-datetime.datetime.now()).days,axis=1)
    sheet3['当前时间']=datetime.datetime.now()
    sheet3['是否已超期']=sheet3['距离下次付款时间（天）'].apply(lambda x:'是' if x<=0 else '否')
    ex['缴费期终']=pd.to_datetime(ex['缴费期终'])
    sheet4=pd.merge(sheet3,ex[['报帐点编码','缴费期终','合同编号','实际报账金额（含税）']],how='left',on=['报帐点编码','缴费期终','合同编号'])
    ew=pd.ExcelWriter(r'结果数据\结果数据'+wbaozhangdianname+'.xlsx')
    sheet4.to_excel(ew,sheet_name='Sheet1',index=False)
    ew.save()
except:
    print('处理失败')
    logFile=open(r'处理历史日志.txt',mode='a',encoding='UTF-8')
    strF=str(datetime.datetime.now())+' :'+wbaozhangdian+'处理失败\n'
    logFile.writelines(strF)
    logFile.flush()
    logFile.close()
else:
    print('处理成功')
    logFile=open(r'处理历史日志.txt',mode='a',encoding='UTF-8')
    strF=str(datetime.datetime.now())+' :'+wbaozhangdian+'处理成功\n'
    logFile.writelines(strF)
    logFile.flush()
    logFile.close()
input("按回车键退出")