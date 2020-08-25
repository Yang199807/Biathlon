from windrose import WindroseAxes
from matplotlib import pyplot as plt
import matplotlib as mpl
import matplotlib.cm as cm
import numpy as np
import seaborn as sns
import xlrd
import datetime
import math

mpl.rcParams['font.sans-serif'] = ['Times New Roman']
mpl.rcParams['font.sans-serif'] = [u'SimHei']
mpl.rcParams['font.size'] = 16
mpl.rcParams['axes.unicode_minus'] = False

#绘图表格列表
#excelList=["2020-08-14.xlsx","2020-08-15.xlsx","2020-08-16.xlsx"]
excelList=["2020-08-19.xlsx"]

# 使用nmupy随机生成风速风向数组
for excel in excelList:
    filePath="D:/2020.8 跟队观测/data/"
    data = xlrd.open_workbook(filePath+excel)
    # 读取数据表
    table = data.sheet_by_index(0)
    # 数据行数
    rows = table.nrows

    V=[[] for i in range(5)]
    D=[]
    for i in range(2,rows):
        #time=xlrd.xldate.xldate_as_datetime(table.cell(i, 1).value, 0).time()
        #startTime=datetime.time(9,0,0,0)
        #endTime = datetime.time(9,15, 0, 0)
        #if(time.__ge__(startTime)):
        for j in range(5):
            v=table.cell(i,2*j+2).value
            d=table.cell(i,2*j+3).value
            dr=d*math.pi/180
            print(dr)
            vx=v*math.sin(dr)
            V[j].append(vx)
        # v=table.cell(i,10).value
        #
        # try:
        #     if(v>4):
        #         V.append(v)
        #         d=table.cell(i,11).value
        #         d+=180
        #         if(d>=360):
        #             d=d-360
        #         D.append(d)
        # except:
        #     pass
        #if (time.__ge__(endTime)):
            #break
    #plt.plot(V,label=excel.split(".")[0])
    #print(V)
    #print(np.average(V),np.average(D))
    #V=[]

#相关系数热力图
cov=np.corrcoef(V)
print(cov)
sns.heatmap(cov,annot=True,cmap="Blues_r",vmin = 0.5,vmax = 1, xticklabels=[1,2,3,4,5], yticklabels=[1,2,3,4,5])
plt.title("风速相关系数图",fontsize=14)
plt.xlabel("风速仪ID")
#plt.xticks([1,2,3,4,5])
plt.ylabel("风速仪ID")
plt.show()

#折线图绘制
# plt.xlabel("时间(上午)")
# plt.xticks([0,300,600,900],["9:00","9:05","9:10","9:15"])
# plt.ylabel("风速(m/s)")
# plt.title("1号风速仪3日风速对比图")
# plt.legend()
# plt.show()
#ws = np.random.random(500) * 6
#wd = np.random.random(500) * 360

#玫瑰图绘制
# ax = WindroseAxes.from_ax()
# plt.title("5号风速仪风玫瑰图(8月14日 高风速)")
# ax.bar(D, V,normed=True, opening=0.8, edgecolor='white')
# #plt.legend(lab, )
# #-0.2，-0.1
# ax.set_legend(loc='lower left', bbox_to_anchor=(-0.3, -0.1),fontsize=20,title="风速（m/s）")
#
# plt.show()