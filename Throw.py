## -*- coding: utf-8 -*-  
#
#上文是解决中文不出错

#下文 是解决ExcelWriter写入中文的时候不出错
import sys
reload(sys)
sys.setdefaultencoding("utf-8")

import tushare as ts
import pandas as pd
import numpy as np

#init set parama
START_DATE	=	'2002-01-23'
END_DATE  	=	'2017-01-01'
amount		=	10000
RESULT_EXCEL_FILENAME='Throw_Result.xlsx'

#退出机制参数
STOP_PROFIT_THRESHOLD	=	0.5  	#止盈门槛
STOP_PROFIT_RATIO		=	0.5		#止盈率

#实际执行时选择的数据
codes=['600036','600150','600446','600519','600570','600887','600999','600109','601668','601989','000002','000423','000651','600048','600104','600547','600683','600820','000333','000568','000858','300003']

names=['招商银行','中国船舶','金证股份','贵州茅台','恒生电子','伊利股份','招商证券','国金证券','中国建筑','中国重工','万科A','东阿阿胶','格力电器','保利地产','上汽集团','山东黄金','京投发展','隧道股份','美的集团','泸州老窖','五粮液','乐普医疗']
# codes=['600036','600150','600109']
# names=['招商银行','中国船舶','国金证券']



# 数据结构测试数据可以删除
# print ds1  #pandas.DataFrame
# print "============================================"
# print ds1.index
# print ds1.columns
# print ds1.iloc[0,2]
# print ds1.iloc[0,1]
# print ds1.iloc[1,1]
#1.#######################################2###########################################
def firstSimple(ds1,result):
	data= ds1.iloc[0,0]
	close = ds1.iloc[0,2]
	result[0][0]=data
	result[0][1]=close
	result[0][2]=amount
	result[0][3]=amount
	result[0][4]=amount/close
	result[0][5]=result[0][4]
	result[0][6]=result[0][5]*close
	result[0][7]=result[0][6]/result[0][3]-1
	return result

def monthSimple(ds1,i,result):
	data= ds1.iloc[i,0]           			#取出日期
	close = ds1.iloc[i,2]         			#取出收盘价
	result[i][0]=data             			#日期
	result[i][1]=close						#收盘价
	result[i][2]=amount						#每月定投金额
	result[i][3]=result[i-1][3]+result[i][2]#投资金额合计
	result[i][4]=amount/close				#股票当期份额
	result[i][5]=result[i-1][5]+result[i][4]#股票合计份额
	result[i][6]=result[i][5]*close			#股票总市值
	result[i][7]=result[i][6]/result[i][3]-1#投资收益比率
	return result

def doSimple(i):     #定期定投  没有退出策略
	print (codes[i])
	ds1=ts.get_k_data(codes[i],start=START_DATE, end=END_DATE, ktype='M', autype='qfq')
	result=[["",0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0] for j in range(ds1.index.size)]
	firstSimple(ds1,result)
	for j in range(ds1.index.size):
		monthSimple(ds1,j,result)	#投资策略这里是修订的地方	
	# print ("---------------------------------------")
	# print result
	datas=np.array(result)
	columns=['日期','收盘价','每月定投金额','投资金额合计','股票当期份额','股票合计份额','股票总市值','投资收益比率','现金','备注']
	df=pd.DataFrame(datas,columns=columns)
	# print ("---------------------------------------")
	# print df
	df.to_excel(writer, sheet_name=names[i])
	return

#2###################################################################################
########withdraw 增加了简单的退出机制#########
def firstWithdraw(ds1,result):
    data= ds1.iloc[0,0]
    close = ds1.iloc[0,2]
    result[0][0]=data
    result[0][1]=close
    result[0][2]=amount
    result[0][3]=amount
    result[0][4]=amount/close
    result[0][5]=result[0][4]
    result[0][6]=result[0][5]*close
    result[0][7]=result[0][6]/result[0][3]-1
    if (result[0][7] < STOP_PROFIT_THRESHOLD):
		result[0][8]= result[0][3]			#调整投资金额
		result[0][9]= result[0][5]			#调整投资份额
		result[0][10]= 0		#盈利出金   
    else:
		result[0][8]= result[0][3]*(1-STOP_PROFIT_RATIO)
		result[0][9]= result[0][5]*(1-STOP_PROFIT_RATIO)
		result[0][10]= result[0][6]*STOP_PROFIT_RATIO
    result[0][11]=result[0][10]-result[0][2]	#现金流
    withdrawSum = result[0][10]  #累计分红
  #   for j in range(i-1):
		# withdrawSum=withdrawSum+result[0][j]
    result[0][12]=withdrawSum+result[0][6]		#总收入
    amountSum=result[0][2]
    # for j in range(i):
	   #  amountSum=amountSum+result[0][2]
    result[0][13]=result[0][12]/amountSum-1		#总投资收益率
    return result

def monthWithdraw(ds1,i,result):
	data= ds1.iloc[i,0]           			#取出日期
	close = ds1.iloc[i,2]         			#取出收盘价 
	result[i][0]=data             			#日期
	result[i][1]=close						#收盘价
	result[i][2]=amount						#每月定投金额
	result[i][3]=result[i][2]+result[i-1][8]	#投资金额合计
	result[i][4]=amount/close				#股票当期份额
	result[i][5]=result[i][4]+result[i-1][9]#股票合计份额
	result[i][6]=result[i][5]*close			#股票总市值
	result[i][7]=result[i][6]/result[i][3]-1#投资收益比率
	if result[i][7] < STOP_PROFIT_THRESHOLD:
		result[i][8]= result[i][3]			#调整后投资金额
		result[i][9]= result[i][5]			#调整后投资份额
		result[i][10]= 0					#盈利出金   
	else: 
		result[i][8]= result[i][3]*(1-STOP_PROFIT_RATIO)
		result[i][9]= result[i][5]*(1-STOP_PROFIT_RATIO)
		result[i][10]= result[i][6]*STOP_PROFIT_RATIO
	result[i][11]=result[i][10]-result[i][2]	#现金流
	withdrawSum=0.0  #累计分红
	for m in range(0,i):
		withdrawSum=withdrawSum+result[m][10]
	result[i][12]=withdrawSum+result[i][6]		#总收入
	amountSum=0.0
	for m in range(0,i+1):
		amountSum=amountSum+result[m][2]
	result[i][13]=result[i][12]/amountSum-1		#总投资收益率
	return result


def doWithdraw(i):     #定期定投  没有退出策略
    print (codes[i])
    ds1=ts.get_k_data(codes[i],start=START_DATE, end=END_DATE, ktype='M', autype='qfq')
    result=[["",0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0] for j in range(ds1.index.size)]
    firstWithdraw(ds1,result)
    for j in range(1,ds1.index.size):
        monthWithdraw(ds1,j,result)  #投资策略这里是修订的地方   
    # print ("---------------------------------------")
    print result
    datas=np.array(result)
    columns=['日期','收盘价','每月定投金额','投资金额合计','股票当期份额','股票合计份额','股票总市值','投资收益比率','调整后投资金额','调整后投资份额','盈利出金','现金流','总收入','总投资收益率']
    df=pd.DataFrame(datas,columns=columns)
    # print ("---------------------------------------")
    # print df
    df.to_excel(writer, sheet_name=names[i])
    return




#执行主程序
writer = pd.ExcelWriter(RESULT_EXCEL_FILENAME, engine='xlsxwriter')
for i in range(0,len(codes)):
	# doSimple(i)
	doWithdraw(i)
writer.save()	
# test()