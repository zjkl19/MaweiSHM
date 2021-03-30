# -*- coding: utf-8 -*-
"""
马尾大桥测点箱线图

@author: 林迪南
"""
import os

import matplotlib.pyplot as plt  # 使用import导入模块matplotlib.pyplot，并简写成plt  
import matplotlib.dates as mdates
import numpy as np  # 使用import导入模块numpy，并简写成np
from matplotlib.ticker import FuncFormatter

import datetime
from matplotlib.dates import DateFormatter

import xlrd
import xlwt

monitorYear=2021;monitorMonth=2;monitorDay=1;
totalHours=672

base = datetime.datetime(monitorYear, monitorMonth, monitorDay)
date1=base.strftime("%Y-%m-%d %H:%M:%S");date2=base +datetime.timedelta(hours=totalHours);
date2=date2.strftime("%Y-%m-%d %H:%M:%S");

dataExcel='Data.xlsx';dataSheet='传感器监测数据报表'
data=xlrd.open_workbook(dataExcel)
sh=data.sheet_by_name(dataSheet)

resultExcel='Result.xls';resultSheet='Result'

if os.path.exists(resultExcel):
    os.remove(resultExcel)

if os.path.exists(resultExcel+'x'):
    os.remove(resultExcel+'x')


defaultFontsize=14    #默认字体大小
nPoints=15    #15个测点

dataCounts=totalHours    #例：31天数据，31*24个，每小时采集一次数据
scaleFactor=1.0    #控制x轴长度

startRow=2    #从excel表格第3行开始读取
measurePointsList=[]    #测点名称
yList=[]    #监测数据


for i in range(0,nPoints):
    measurePointsList.append(sh.cell_value(startRow-1,i+1))

dates=[]
for i in range(0,dataCounts):
    dates.append(datetime.datetime.strptime(sh.cell_value(startRow+i,0), "%Y-%m-%d %H:%M:%S"))
    yList.append(sh.cell_value(startRow+i,1))

#dates即x轴数据
#dates = [base + datetime.timedelta(hours=(1 * i)) for i in range(dataCounts)]    #1小时1个数据

yList = np.array(yList)    #列表转换为nparray

xFigsize=10;yFigsize=3
rotAngle=40    #旋转角度





lims = [(np.datetime64(date1), np.datetime64(date2)),
        (np.datetime64(date1), np.datetime64(date2)),
        (np.datetime64(date1), np.datetime64(date2)),
        (np.datetime64(date1), np.datetime64(date2)),
        (np.datetime64(date1), np.datetime64(date2)),
        (np.datetime64(date1), np.datetime64(date2)),
        (np.datetime64(date1), np.datetime64(date2)),
        (np.datetime64(date1), np.datetime64(date2)),
        (np.datetime64(date1), np.datetime64(date2)),
        (np.datetime64(date1), np.datetime64(date2)),
        (np.datetime64(date1), np.datetime64(date2)),
        (np.datetime64(date1), np.datetime64(date2)),
        (np.datetime64(date1), np.datetime64(date2)),
        (np.datetime64(date1), np.datetime64(date2)),
        (np.datetime64(date1), np.datetime64(date2))]

#      1            2                3              4                 5               6                  7                   8                     9                       10             11                12
#F17-D-位移(mm)	G11-D-位移(mm)	F17-L-倾角(°)	G11-L-倾角(°)	D2-2倾角-倾角(°)	ZW24-1倾角-倾角(°)	ZW24-2倾角-倾角(°)	ZE24-1倾角-倾角(°)	ZE24-2倾角-倾角(°)	D2-2-1位移-位移(mm)	D2-2-2位移-位移(mm)	ZW24-1位移-位移(mm)	ZW24-2位移-位移(mm)	ZE24-1位移-位移(mm)	ZE24-2位移-位移(mm)

#      1            2                3              4                 5               
#6                  7                   8                     9                       10
#             11                12
#F17-D-位移(mm)	G11-D-位移(mm)	F17-L-倾角(°)	G11-L-倾角(°)	D2-2倾角-倾角(°)	
#ZW24-1倾角-倾角(°)	ZW24-2倾角-倾角(°)	ZE24-1倾角-倾角(°)	ZE24-2倾角-倾角(°)	D2-2-1位移-位移(mm)	
#D2-2-2位移-位移(mm)	ZW24-1位移-位移(mm)	ZW24-2位移-位移(mm)	ZE24-1位移-位移(mm)	ZE24-2位移-位移(mm)


defaultDispLims=(-1.5,1.5)
defaultLeanLims=(-0.001,0.002)
defaultLeanAlertSection=(-0.1,0.1)

ylabelList=['位移(mm)','位移(mm)','倾角(°)','倾角(°)','倾角(°)','倾角(°)','倾角(°)','倾角(°)','倾角(°)','位移(mm)','位移(mm)','位移(mm)','位移(mm)','位移(mm)','位移(mm)']

ylims = [(-0.08,0.10),(-0.02,0.4),(-0.014,0.006),(-0.005,0.025),(0.00-0.005,0.02+0.005)
        ,(-0.020,0.005),(-0.025,0.000),(-0.01,0.01),(-0.015,0.015),(149.6,150.2)
        ,(169.4,170.4),(249.6,251.4),(222.0,222.6),(148.4,148.8),(129.7,130.3)]

yticklabelFormat=['{:1.2f}','{:1.2f}','{:1.4f}','{:1.4f}','{:1.4f}'
                  ,'{:1.4f}','{:1.4f}','{:1.4f}','{:1.4f}','{:1.2f}'
                  ,'{:1.2f}','{:1.2f}','{:1.2f}','{:1.2f}','{:1.2f}']

alertLine = [(-1.0, 1.0),(-1.0, 1.0),defaultLeanAlertSection,defaultLeanAlertSection,defaultLeanAlertSection
        ,defaultLeanAlertSection,defaultLeanAlertSection,defaultLeanAlertSection,defaultLeanAlertSection,(-1.0, 1.0)
        ,(-1.0, 1.0),(-1.0, 1.0),(-1.0, 1.0),(-1.0, 1.0),(-1.0, 1.0)
        ,(-1.0, 1.0),(-1.0, 1.0),(-1.0, 1.0),(-1.0, 1.0),(-1.0, 1.0)
        ,(-1.0, 1.0),(-1.0, 1.0),(-1.0, 1.0),(-1.0, 1.0),(-1.0, 1.0)]

#参考
#string[] headerTitle = new string[dataColumns] { "F17-D-位移(mm)", "G11-D-位移(mm)", 
#"ZW24-1位移-位移(mm)", "ZW24-2位移-位移(mm)", "ZE24-1位移-位移(mm)", "ZE24-2位移-位移(mm)", "F17-D-位移(mm)"
#, "G11-D-位移(mm)" };

#string[] headerTitle = new string[dataColumns] { "D2-2-1位移-位移(mm)", "D2-2-2位移-位移(mm)", 
#"ZW24-1位移-位移(mm)", "ZW24-2位移-位移(mm)", "ZE24-1位移-位移(mm)", "ZE24-2位移-位移(mm)", "F17-D-位移(mm)"
#, "G11-D-位移(mm)" };
#decimal[] benchmarkData = new decimal[dataColumns] { 170m, 150m, 240m, 210m, 145m, 130m, 338.00m, 250.00m };
baseData      = [0.00,  0.00,0.0,0.0,0.0,0.0,0.0,0.0,0.0,150.0, 170.0, 240.0, 210.0, 145.0, 130.0]
directionCoff = [ -1.00, -1.0   ,1.0,1.0,1.0,1.0,1.0,1.0,1.0,1.0  ,   1.0,   1.0,   1.0,   1.0,   1.0 ];
 
yList = [[] for i in range(nPoints)]  # 创建的是多行n列的二维列表
for i in range(0,nPoints):
    startRow=2
    temp=[]    
    for j in range(0,dataCounts):        
        temp.append(baseData[i]+directionCoff[i]*sh.cell_value(startRow+j,i+1))
    yList[i] = np.array(temp)

plt.rcParams['font.sans-serif']=['SimHei'] #用来正常显示中文标签
plt.rcParams['axes.unicode_minus']=False #用来正常显示负号

def formatnum(x, pos):
    return '$%.2f$' % (x)

formatter = FuncFormatter(formatnum)

xFormatter = DateFormatter('%m-%d')

w=xlwt.Workbook()
ws=w.add_sheet(resultSheet)


for i in range(0,nPoints):
    fig, ax = plt.subplots(constrained_layout=True, figsize=(xFigsize, yFigsize))
    
    ax.plot(dates, yList[i])
    #ax.set(xlabel='时间', ylabel=ylabelList[i])
    
    ax.grid()
    
    ax.xaxis.set_major_formatter(xFormatter)
    
    ax.yaxis.set_major_formatter(formatter)
    ax.set_xlim(lims[i])
    
    ax.set_ylim(ylims[i])
    
    vals = ax.get_yticks()
    ax.set_yticklabels([yticklabelFormat[i].format(x) for x in vals])
    
    
    #作双轴曲线
    #---------
#    ax2 = ax.twinx()  # instantiate a second axes that shares the same x-axis
#
#    color = 'tab:blue'
#    ax2.set_ylabel('sin', color=color)  # we already handled the x-label with ax1
#    ax2.plot(dates, yList[1], color=color)
#    ax2.tick_params(axis='y', labelcolor=color)
    #---------
    
    #-------
    max_indx=np.argmax(yList[i])#max value index
    min_indx=np.argmin(yList[i])#min value index
    #plt.plot(yList[i],'r-o')
    #plt.plot(max_indx,yList[i][max_indx],'ks')
    show_max='['+str(max_indx)+' '+str(yList[i][max_indx])+']'
    print(show_max)
    show_min='['+str(min_indx)+' '+str(yList[i][min_indx])+']'
    print(show_min)
    
    ws.write(0,i,measurePointsList[i])
    ws.write(1,i,yList[i][max_indx])
    ws.write(2,i,yList[i][min_indx])
    #plt.annotate(show_max,xytext=(max_indx,yList[i][max_indx]),xy=(max_indx,yList[i][max_indx]))
    #plt.plot(min_indx,yList[i][min_indx],'gs')
    #-------
    
    #ax.get_yaxis().get_major_formatter().set_scientific(False)
    plt.xticks(fontsize=defaultFontsize)
    plt.yticks(fontsize=defaultFontsize)
    plt.xlabel(xlabel='时间',fontsize=defaultFontsize)
    plt.ylabel(ylabel=ylabelList[i],fontsize=defaultFontsize)
    #作预警线
    #---------
#    plt.axhline(y=alertLine[i][0],c="yellow")#添加水平直线
#    plt.axhline(y=alertLine[i][1],c="yellow")
    #---------
    plt.savefig(measurePointsList[i]+'.jpg')
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
    
#for i in range(0,1):
#    axs[0].plot(dates, yList[i])
#    axs[0].set_xlim(lims[i])
#    # rotate_labels...
#    for label in axs[0].get_xticklabels():
#        label.set_rotation(rotAngle)
#        label.set_horizontalalignment('right')
#
#axs[0].set_title('Default Date Formatter')
#plt.show()

#fig, ax = plt.subplots()
#ax.plot(xList, yList)

#ax.xaxis.set_minor_locator(AutoMinorLocator())

#ax.tick_params(which='both', width=2)
#ax.tick_params(which='major', length=7)
#ax.tick_params(which='minor', length=4, color='r')
    
w.save(resultExcel)

#03excel转07及以上excel
#https://www.cnblogs.com/zifeiy/p/8142853.html
import win32com.client as win32

fname = resultExcel
excel = win32.gencache.EnsureDispatch('Excel.Application')
wb = excel.Workbooks.Open(os.getcwd()+'\\'+resultExcel)

wb.SaveAs(os.getcwd()+'\\'+resultExcel+'x', FileFormat = 51)    #FileFormat = 51 is for .xlsx extension
wb.Close()                               #FileFormat = 56 is for .xls extension
excel.Application.Quit()