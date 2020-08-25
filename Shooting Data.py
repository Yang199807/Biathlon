import xlrd
from xlutils.copy import copy

def getTime():
    timeList=[]
    hour=input()
    #while(hour==""):
    #   hour=input()
    min=input()
    sec=input().split(' ')
    for s in sec:
        time="{0}:{1}:{2}".format(hour,min,s)
        timeList.append(time)
    return timeList
# while True:
#     getTime()
if __name__ == '__main__':
    filePath="运动员射击记录8.21.xls"
    #用于对应运动员表格
    playerDic={"lin":0,"yu":1,"qi":2,"zhi":3,"meng":4,"yan":5,"yuan":6,"han":7,"huan":8,"ming":9,"le":10}
    #每天的日期固定
    date="2020/8/24"
    while True:
        # 读取表格
        workbook = xlrd.open_workbook(filePath)
        # 生成表格拷贝
        new_book = copy(workbook)
        #获取运动员表格
        print("姓名")
        playerName=input()
        sheetNumber=playerDic[playerName]

        #读取对应表格
        rSheet = workbook.sheets()[sheetNumber]
        wSheet = new_book.get_sheet(sheetNumber)
        rows=rSheet.nrows
        #输入信息
        print("靶位")
        loc=input()
        print("姿势")
        pos=input()
        print("时间")
        timeList=getTime()

        #写入信息
        for j in range(len(timeList)):
            wSheet.write(rows+j,0,date)#日期
            wSheet.write(rows+j,1,loc) #靶位
            wSheet.write(rows+j,2,pos) #姿势
            wSheet.write(rows+j,3,timeList[j])#时间

        #处理时间中断情况

        if(len(timeList)<5):
            n=len(timeList)
            #rows=rows+n
            print("补充时间")
            timeList=getTime()
            for j in range(len(timeList)):
                wSheet.write(rows  +n +j, 0, date)  # 日期
                wSheet.write(rows +n+ j, 1, loc)  # 靶位
                wSheet.write(rows +n+ j, 2, pos)  # 姿势
                wSheet.write(rows +n+ j, 3, timeList[j])#时间

        #输入并写入命中信息
        print("命中情况")
        shoot=input()
        for j in range(5):
            wSheet.write(rows+j,4,shoot[j])
        #写完后保存
        new_book.save(filePath)

    print(rSheet.nrows)
    print(rSheet.cell(0,0).value)