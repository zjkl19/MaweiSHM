# -*- coding: utf-8 -*-

import xlrd
import xlwt


import numpy as np
import pandas as pd
 
import matplotlib.pyplot as plt
 
dataExcel='TestData.xlsx';dataSheet='Sheet1'
data=xlrd.open_workbook(dataExcel)
sh=data.sheet_by_name(dataSheet)
 
measurePointsList=[]    #测点名称
yList=[]    #监测数据


dataCounts=9410
startRow=2    #从excel表格第3行开始读取
nPoints=15    #15个测点

for i in range(0,nPoints):
    measurePointsList.append(sh.cell_value(startRow-1,i+1))

dates=[]
for i in range(0,dataCounts):
    yList.append(sh.cell_value(startRow+i,1))

yList = [[] for i in range(nPoints)]  # 创建的是多行n列的二维列表
for i in range(0,nPoints):
    startRow=2
    temp=[]
    for j in range(0,dataCounts):        
        temp.append(sh.cell_value(startRow+j,i+1))
    yList[i] = np.array(temp)
    

data = [yList[0], yList[1], yList[2]]
labels = ['A1','B','C']

# Multiple box plots on one Axes
fig, ax = plt.subplots()
ax.boxplot(data,0, '', labels=labels)

alertLineWidth=0.4    #预警线长度

alertColor=['y','y','r','r']
alertValue=[[-178*2/3,163*2/3,-178,163],[-92*2/3,114*2/3,-92,114],[-92*2/3,114*2/3,-92,114]]    #各个测点的黄色预警上下界及红色预警上下界

for i in range(0,len(data)):  
    x = np.array([i+1-alertLineWidth/2,i+1+alertLineWidth/2])
    for j in range(0,4):    #遍历黄色预警和红色预警
        y = np.array([alertValue[i][j],alertValue[i][j]])
        #有了 x 和 y 数据之后，我们通过 plt.plot(x, y) 来画出图形，并通过 plt.show() 来显示。
        ax.plot(x,y,c=alertColor[j])




plt.show()