# -*- coding: utf-8 -*-
"""
Created on Fri Jul 24 10:51:33 2020

@author: eamdf
"""
import matplotlib.pyplot as plt  # 使用import导入模块matplotlib.pyplot，并简写成plt  
import numpy as np  # 使用import导入模块numpy，并简写成np

import xlrd #导入x1rd库


data=xlrd.open_workbook('Data.xlsx') #打开Exce1文件
sh=data.sheet_by_name('传感器监测数据报表') #获得需要的表单

dataCounts=7*24    #7天数据，每小时采集一次数据
scaleFactor=1.0    #控制x轴长度

startRow=2    #第3行
xList=[]
yList=[]
xticksLabelList=[]

#print(sh.cell_value(3,0))
for i in range(0,dataCounts):
    if i % 24 ==0:
        #xticksLabelList.append(i+1)
        xticksLabelList.append(sh.cell_value(startRow+i,0))
    else:
        xticksLabelList.append('')
    xList.append(i*scaleFactor)
    yList.append(sh.cell_value(startRow+i,1))

#xticksLabelList=['2','','3']

xList = np.array(xList)
yList = np.array(yList)
plt.figure(figsize=(15,3))

#plt.plot(xList, yList)
#plt.xticks(xList, xticksLabelList)

import datetime
import matplotlib.dates as mdates

xFigsize=10
yFigsize=6
rotAngle=40
base = datetime.datetime(2020, 7, 17)
dates = [base + datetime.timedelta(hours=(1 * i)) for i in range(dataCounts)]    #1小时1个数据

fig, axs = plt.subplots(1, 1, constrained_layout=True, figsize=(xFigsize, yFigsize))
lims = [(np.datetime64('2020-07-17'), np.datetime64('2020-07-24 00:00')),
        (np.datetime64('2020-07-17'), np.datetime64('2020-07-24 00:00')),
        (np.datetime64('2020-07-17'), np.datetime64('2020-07-24 00:00'))]

yList = [[] for i in range(3)]  # 创建的是多行三列的二维列表
for i in range(0,1):
    startRow=2
    temp=[]    
    for j in range(0,dataCounts):        
        temp.append(sh.cell_value(startRow+j,i+1))
    yList[i] = np.array(temp)

#k=0
#for nn, ax in enumerate(axs):
#          
#    ax.plot(dates, yList[k])
#    ax.set_xlim(lims[nn])
#    # rotate_labels...
#    for label in ax.get_xticklabels():
#        label.set_rotation(rotAngle)
#        label.set_horizontalalignment('right')
#    k=k+1
    
for i in range(0,1):
    axs[0].plot(dates, yList[i])
    axs[0].set_xlim(lims[i])
    # rotate_labels...
    for label in axs[0].get_xticklabels():
        label.set_rotation(rotAngle)
        label.set_horizontalalignment('right')

axs[0].set_title('Default Date Formatter')
plt.show()

#fig, ax = plt.subplots()
#ax.plot(xList, yList)

#ax.xaxis.set_minor_locator(AutoMinorLocator())

#ax.tick_params(which='both', width=2)
#ax.tick_params(which='major', length=7)
#ax.tick_params(which='minor', length=4, color='r')



#以下两行防止乱码
plt.rcParams['font.sans-serif']=['SimHei'] #用来正常显示中文标签
plt.rcParams['axes.unicode_minus']=False #用来正常显示负号

legendList=['E-QD111','E-QD112','E-QD113','E-QD114','E-QD115','E-QD123']
markerList=['o','s','p','x','v','h']
xticksLabelList=['20161105','20161206','20161213','20161225','20170108','20170116','20170205','20170212','20170224']
x = np.array([1,2,3,4,5,6,7,8,9])
yList=[
    np.array([0,0.21,0.02,-1.32,-1.34,-1.11,-1.33,-1.52,-1.65])
    ,np.array([0,-0.18,-0.23,-0.23,-1.93,-1.96,-2.67,-2.78,-2.58])
    ,np.array([0,-0.33,-0.47,-1.02,-1.55,-1.42,-1.63,-1.82,-1.68])
    ,np.array([0,0.31,-0.32,-0.62,-0.72,-0.63,-0.53,-0.35,-0.36])
    ,np.array([0,0.11,0.29,0.25,0.22,0.19,0.12,0.35,0.21])
    ,np.array([0,0.14,0.33,0.29,0.12,0.27,0.27,0.13,0.21])
]

plt.figure()  # 使用plt.figure定义一个图像窗口.

for i in range(0,len(yList)):
    plt.plot(x, yList[i],marker=markerList[i])

plt.legend(legendList)
plt.xticks(x, xticksLabelList)    #自定义标签
plt.xlabel('时间')
plt.ylabel('沉降（mm）')
plt.title('E匝道第一阶段桥墩沉降数据时程曲线图')  # 设置标题
plt.show()