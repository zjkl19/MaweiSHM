# -*- coding: utf-8 -*-

import xlrd
import xlwt


import numpy as np
import pandas as pd
 
import matplotlib.pyplot as plt

secName='D4'

dataExcel='Data.xlsx';dataSheet=secName
data=xlrd.open_workbook(dataExcel)
sh=data.sheet_by_name(dataSheet)

measurePointsList=[]    #测点名称
yList=[]    #所有监测数据
labels=[]    #作箱线图用的标签
data=[]    #作箱线图用的数据

alertDic = {
  'LS1': [[-92*2/3,114*2/3,-92,114],[-178*2/3,163*2/3,-178,163]],
  'LS3': [[-114*2/3,158*2/3,-114,158],[-207*2/3,178*2/3,-207,178]],
  'LS6': [[-109*2/3,70*2/3,-109,70],[-109*2/3,70*2/3,-109,70],[-176*2/3,212*2/3,-176,212]],
  'LS8': [[-79*2/3,70*2/3,-79,70],[-79*2/3,70*2/3,-79,70],[-79*2/3,70*2/3,-79,70],[-79*2/3,70*2/3,-79,70]
          ,[-129*2/3,148*2/3,-129,148],[-129*2/3,148*2/3,-129,148],[-129*2/3,148*2/3,-129,148],
          [-129*2/3,148*2/3,-129,148],[-129*2/3,148*2/3,-129,148],[-129*2/3,148*2/3,-129,148]],
  'LSA': [[-183*2/3,237*2/3,-183,237],[-183*2/3,237*2/3,-183,237],
          [-261*2/3,224*2/3,-261,224],[-261*2/3,224*2/3,-261,224],
          [-261*2/3,224*2/3,-261,224]],
  'LS12': [[-79*2/3,70*2/3,-79,70],[-79*2/3,70*2/3,-79,70],
           [-149*2/3,188*2/3,-149,188],[-149*2/3,188*2/3,-149,188]],
  'LSC': [[-171*2/3,243*2/3,-171,243],[-171*2/3,243*2/3,-171,243],[-171*2/3,243*2/3,-171,243]
          ,[-267*2/3,218*2/3,-267,218],[-267*2/3,218*2/3,-267,218]],
  'RSA': [[-180*2/3,237*2/3,-180,237],[-180*2/3,237*2/3,-180,237],[-180*2/3,237*2/3,-180,237],
          [-261*2/3,224*2/3,-261,224],[-261*2/3,224*2/3,-261,224],[-261*2/3,224*2/3,-261,224]],
  #'RSB': [[-109*2/3,70*2/3,-109,70],[-178*2/3,163*2/3,-178,163]],
  'RSC': [[-180*2/3,243*2/3,-180,243],[-180*2/3,243*2/3,-180,243],[-180*2/3,243*2/3,-180,243],
          [-267*2/3,218*2/3,-267,218],[-267*2/3,218*2/3,-267,218],[-267*2/3,218*2/3,-267,218]],
  'RS7': [[-82*2/3,95*2/3,-82,95],
          [-137*2/3,175*2/3,-137,175],[-137*2/3,175*2/3,-137,175],[-137*2/3,175*2/3,-137,175],
          [-137*2/3,175*2/3,-137,175],[-137*2/3,175*2/3,-137,175]],
  'RS10': [[-90*2/3,95*2/3,-90,95],[-90*2/3,95*2/3,-90,95],
           [-90*2/3,95*2/3,-90,95],[-90*2/3,95*2/3,-90,95],],
  'RS14': [[-112*2/3,160*2/3,-112,160],[-112*2/3,160*2/3,-112,160],
           [-211*2/3,178*2/3,-211,178]],
  'RS16': [[-85*2/3,95*2/3,-85,95],[-85*2/3,95*2/3,-85,95]
           ,[-175*2/3,169*2/3,-175,169]],}

limDic={
  'D1': (-0.4, 0.2),
  'D2': (-0.4, 0.2),
  'D3': (-0.4, 0.2),
  'D4': (-0.2, 0.1),
  'D5': (-0.4, 0.2),
  'D6': (-0.3, 0.2),
  'D7': (-0.4, 0.2),
  'D8': (-0.4, 0.2),
  #'RSB': (-150, 150),
  'RSC': (-300, 300),
  'RS7': (-200, 200),
  'RS10': (-120, 120),
  'RS14': (-230, 230),
  'RS16': (-200, 200),
  
}

autoDetectPointNumber=True    #自动判断测点数

labelSize=15
tickFontSize=15
dataCounts=14
measurePointsTitleRow=0    #第1列



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


# Multiple box plots on one Axes
fig, ax = plt.subplots()
ax.boxplot(data,0, '', labels=labels)

alertLineWidth=0.25    #预警线长度

alertColor=['y','y','r','r']    #预警值颜色，'y'表示黄色，'r'表示红色
#alertValue=alertDic[secName]    #各个测点的黄色预警上下界及红色预警上下界

#for i in range(0,len(data)):  
    #x = np.array([i+1-alertLineWidth/2,i+1+alertLineWidth/2])
    #for j in range(0,4):    #遍历黄色预警和红色预警
        #y = np.array([alertValue[i][j],alertValue[i][j]])
        #有了 x 和 y 数据之后，我们通过 plt.plot(x, y) 来画出图形，并通过 plt.show() 来显示。
        #ax.plot(x,y,c=alertColor[j])

plt.xticks(fontsize=tickFontSize)
plt.yticks(fontsize=tickFontSize) 
ax.set_xlabel('传感器编号',fontsize=labelSize)
ax.set_ylabel('相对变化量（mm)',fontsize=labelSize)
ax.set_ylim(limDic[secName])
plt.show()