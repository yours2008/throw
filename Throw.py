## -*- coding: utf-8 -*-  
#上文是解决中文不出错


#下文 是解决ExcelWriter写入中文的时候不出错
import sys
reload(sys)
sys.setdefaultencoding("utf-8")

import tushare as ts
import pandas as pd
import numpy as np

#init set parama
START_DATE='2012-01-23'
END_DATE='2017-01-01'
amount=10000

# codes=['600036','600150','600446','600519','600570','600887','600999','601668','601989','000002','000423','000651','600048','600104','600547','600683','600820','000333','000568','000858','300003']
# names=['招商银行','中国船舶','金证股份','贵州茅台','恒生电子','伊利股份','招商证券','中国建筑','中国重工','万科A','东阿阿胶','格力电器','保利地产','上汽集团','山东黄金','京投发展','隧道股份','美的集团','泸州老窖','五粮液','乐普医疗']


codes=['600036','600150']
names=['招商银行','中国船舶']

RESULT_EXCEL_FILENAME='Throw_Result.xlsx'

# 数据结构测试数据可以删除
# print ds1  #pandas.DataFrame
# print "============================================"
# print ds1.index
# print ds1.columns
# print ds1.iloc[0,2]
# print ds1.iloc[0,1]
# print ds1.iloc[1,1]

def oneMonth(ds1,i,result):
	data= ds1.iloc[i,0]
	close = ds1.iloc[i,2]
	result[i][0]=data
	result[i][1]=close
	result[i][2]=amount
	result[i][3]=result[i-1][3]+amount

	result[i][4]=amount/close
	result[i][5]=result[i-1][5]+result[i][4]

	result[i][6]=result[i][5]*close
	result[i][7]=result[i][6]/result[i][3]-1
	return result

def firstMonth(ds1,result):
	data= ds1.iloc[0,0]
	close = ds1.iloc[0,2]
	# result[0,0]=data
	result[0][1]=close
	result[0][2]=amount
	result[0][3]=amount
	result[0][4]=amount/close
	result[0][5]=result[0][4]
	result[0][6]=result[0][5]*close
	result[0][7]=result[0][6]/result[0][3]-1
	return result

writer = pd.ExcelWriter(RESULT_EXCEL_FILENAME, engine='xlsxwriter')

# def test():
# 	print START_DATE
# 	print codes

def doit(i):
	print (codes[i])
	ds1=ts.get_k_data(codes[i],start=START_DATE, end=END_DATE, ktype='M', autype='qfq')
	result=[["",0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0] for j in range(ds1.index.size)]
	firstMonth(ds1,result)
	for j in range(ds1.index.size):
		oneMonth(ds1,j,result)	#投资策略这里是修订的地方	
	# print ("---------------------------------------")
	# print result
	datas=np.array(result)
	columns=['日期','收盘价','每月定投金额','投资金额合计','股票当期份额','股票合计份额','股票总市值','投资收益比率','现金','备注']
	df=pd.DataFrame(datas,columns=columns)
	# print ("---------------------------------------")
	# print df
	df.to_excel(writer, sheet_name=names[i])
	return

for i in range(0,len(codes)):
	doit(i)
	

writer.save()	
# test()