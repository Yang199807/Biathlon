#coding=utf-8

import xlrd
import datetime
import matplotlib as mpl
import matplotlib.pyplot as plt
import pandas as pd
#测试excel读取
def test(filePath):
    pass

if __name__ == '__main__':
    #打开excel
    filePath="D:/2020.8 跟队观测/data/2020-08-19.xlsx"
    #df=pd.read_excel(filePath)
    #print(df[1])
    data=xlrd.open_workbook(filePath)
    #读取数据表
    table=data.sheet_by_index(0)
    #数据行数
    rows=table.nrows
    #表头占两行，数据从第二行开始
    #test=table.cell(2620,1).value
    #print(test+"hhhhh")

    #数据起始时刻
    start_time=xlrd.xldate.xldate_as_datetime(table.cell(2, 1).value, 0)

    #绘图数据
    timeList=[]
    v1=[]
    v2=[]
    v3=[]
    v4=[]
    v5=[]

    #print("总时长为：{0}  数据行数：{1}".format((end_time-start_time).seconds,rows-2))

    #用于记录数据差值
    temp=start_time
    #用于记录数据个数
    count=1
    i=3
    while i<rows:
        try:
            time=xlrd.xldate.xldate_as_datetime(table.cell(i, 1).value, 0)  # 直接转化为datetime对象
            # timeList.append(time.time().strftime('%H:%M:%S'))
            # v1.append(table.cell(i,2).value)
            # v2.append(table.cell(i, 4).value)
            # v3.append(table.cell(i, 6).value)
            # v4.append(table.cell(i, 8).value)
            # v5.append(table.cell(i, 10).value)
            delta = (time - temp).seconds
            if (delta >1):
                print(i)
            temp = time
            count+=1
            i+=1
        except:
            print("数据间断")
            #记录当前时间为结束时间
            end_time = time
            print("开始时间：{0}  结束时间：{1} 总时长为：{2}  数据行数：{3}"
                .format(start_time.time(),end_time.time(),(end_time - start_time).seconds, count))
            #数据间断为3行空白，基于此更新初始时间
            i+=3
            #初始化开始时间与计数
            start_time = xlrd.xldate.xldate_as_datetime(table.cell(i, 1).value, 0)
            count=1

    #记录最后的时间
    end_time  =xlrd.xldate.xldate_as_datetime(table.cell(rows-1, 1).value, 0)
    print("开始时间：{0}  结束时间：{1} 总时长为：{2}  数据行数：{3}"
          .format(start_time.time(), end_time.time(), (end_time - start_time).seconds, count))

    #
    # v1=v1[0:600]
    # v2 = v2[0:600]
    # v3 = v3[0:600]
    # v4 = v4[0:600]
    # v5 = v5[0:600]
    # plt.plot(v1,label="1号风速仪")
    # plt.plot(v2, label="2号风速仪")
    # plt.plot(v3, label="3号风速仪")
    # plt.plot(v4, label="4号风速仪")
    # plt.plot(v5, label="5号风速仪")
    # plt.xticks([i for i in range(0,len(v1),180)],[timeList[i] for i in range(0,len(v1),180)])
    # plt.xticks(rotation=15)
    # # #plt.xticks()
    # plt.legend()
    # plt.show()
    #print(time.time())