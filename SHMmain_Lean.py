# -*- coding: utf-8 -*-

import xlrd
import xlwt


import numpy as np
import pandas as pd


import matplotlib.pyplot as plt

from matplotlib.ticker import FuncFormatter

secName='I7'

dataExcel='MWLean.xlsx';dataSheet=secName
data=xlrd.open_workbook(dataExcel)
sh=data.sheet_by_name(dataSheet)

measurePointsList=[]    #测点名称
yList=[]    #所有监测数据
labels=[]    #作箱线图用的标签
data=[]    #作箱线图用的数据

alertDic = {
  'LS1': [[-92*2/3,114*2/3,-92,114],[-178*2/3,163*2/3,-178,163]],
  'LS2': [[-92*2/3,114*2/3,-92,114],[-92*2/3,114*2/3,-92,114],[-92*2/3,114*2/3,-92,114],[-92*2/3,114*2/3,-92,114]
          ,[-127*2/3,140*2/3,-127,140],[-127*2/3,140*2/3,-127,140],[-127*2/3,140*2/3,-127,140]],
  'LS6': [[-109*2/3,70*2/3,-109,70],[-178*2/3,163*2/3,-178,163]],
}

limDic={
  'I1': (-0.02, 0.02),
  'I2': (-0.02, 0.02),
  'I3': (-0.025, 0.02),
  'I4': (-0.03, 0.03),
  'I5': (-0.02, 0.02),
  'I6': (-0.02, 0.02),
  'I7': (-0.02, 0.02),
  'I8': (-0.02, 0.02),
}

autoDetectPointNumber=True    #自动判断测点数
autoSetYAxisFormatnum=False    #是否自动设置y轴数字精度
autoYlim=False    #自动设置y轴限值，如果手动，需要设置limDic
plotAlertLine=False    #是否作预警线

reviseFactor=1000    #原始数据修正系数（如m换算成mm）
labelSize=15
tickFontSize=15
dataCounts=14
measurePointsTitleRow=0    #第1列

xLabelName='传感器编号'
yLabelName='倾角（°)'

nPoints=3    #3个测点
if autoDetectPointNumber:
    nPoints=len(sh.row(1))-1

for i in range(0,nPoints):
    measurePointsList.append(sh.cell_value(measurePointsTitleRow,i+1))

yList = [[] for i in range(nPoints)]  # 创建的是多行n列的二维列表
for i in range(0,nPoints):
    startRow=1    #从excel表格第2行开始读取
    temp=[]
    for j in range(0,dataCounts):        
        temp.append(sh.cell_value(startRow+j,i+1))
    yList[i] = np.array(temp)
    
for i in range(0,len(yList)):
    labels.append(measurePointsList[i])
    data.append(yList[i])
#参考等效代码
#data = [yList[0], yList[1], yList[2]]
#labels = [measurePointsList[0],measurePointsList[1],measurePointsList[2]]



#以下两行防止乱码
plt.rcParams['font.sans-serif']=['SimHei'] #用来正常显示中文标签
plt.rcParams['axes.unicode_minus']=False #用来正常显示负号

def formatnum(x, pos):
    return '$%.3f$' % (x)

# Multiple box plots on one Axes
fig, ax = plt.subplots()
ax.boxplot(data,0, '', labels=labels)

if not autoSetYAxisFormatnum:
    ax.yaxis.set_major_formatter(formatnum)


if plotAlertLine:

    alertLineWidth=0.2    #预警线长度
    
    alertColor=['y','y','r','r']    #预警值颜色，'y'表示黄色，'r'表示红色
    alertValue=alertDic[secName]    #各个测点的黄色预警上下界及红色预警上下界
    
    for i in range(0,len(data)):  
        x = np.array([i+1-alertLineWidth/2,i+1+alertLineWidth/2])
        for j in range(0,4):    #遍历黄色预警和红色预警
            y = np.array([alertValue[i][j],alertValue[i][j]])
            #有了 x 和 y 数据之后，我们通过 plt.plot(x, y) 来画出图形，并通过 plt.show() 来显示。
            ax.plot(x,y,c=alertColor[j])

plt.xticks(fontsize=tickFontSize)
plt.yticks(fontsize=tickFontSize) 
ax.set_xlabel('传感器编号',fontsize=labelSize)
ax.set_ylabel('倾角（°)',fontsize=labelSize)

if not autoYlim:
    ax.set_ylim(limDic[secName])

plt.show()
plt.savefig(secName+'.jpg')