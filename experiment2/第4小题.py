import matplotlib.pyplot as plt
import os
import xlrd
import numpy as np
import math
import seaborn as sn
import pandas as pd

#将实验一中清洗处理得到的final_data.xlsx的数据导入data
data= []
filePath = "final_data.xlsx"

if os.path.exists(filePath):
    xlsxFile = xlrd.open_workbook(filePath)
    xlsxSheet = xlsxFile.sheets()[0]
    # 获取行数、列数
    rowNum = int(xlsxSheet.nrows)
    colNum = int(xlsxSheet.ncols)
    for row in range(1,rowNum):
        data.append(xlsxSheet.row_values(row))


#将值为“null”的成绩赋值为0
for i in range(len(data)):
    for j in range(5,14):
        if(data[i][j]=='null'):
            data[i][j]=0

#把9门成绩放入lesson列表中，一个子列表有9门成绩
lessons=[]
for i in range(len(data)):
    lessons.append([])
for i in range(len(data)):
    for j in range(5,14):
        #数据在xlsx中是以string形式存在的，要转化为float
        lessons[i].append(float(data[i][j]))

#第4题：计算出100x100的相关矩阵，并可视化出混淆矩阵
#求标准差
std=[]
for i in range(len(lessons)):
    std.append([])
for i in range(len(lessons)):
        std[i].append(np.std(lessons[i]))

#求均值
sum1=[]
average=[]
for i in range(len(lessons)):
    sum1.append(0)
    average.append(0)
for i in range(len(lessons)):
    for j in range(9):
        sum1[i]+=lessons[i][j]
    average[i]=sum1[i]/len(lessons[i])

#求各行数据之间的协方差
result1=[]
sum2=[]
for i in range(len(lessons)):
    sum2.append([])
    result1.append([])
for i in range(len(lessons)):
    for j in range(len(lessons)):
        sum2[i].append(0)
        result1[i].append(0)

for i in range(len(lessons)):
    for j in range(len(lessons)):
        for k in range(9):
            sum2[i][j]+=(lessons[i][k]-average[i])*(lessons[j][k]-average[j])
        sum2[i][j]=sum2[i][j]/len(lessons[i])
        result1[i][j]='%.6f'%(sum2[i][j]/(std[i][0]*std[j][0]))

#将各行之间的相关系数写入excel文件中，方便查看与验证
#同时将单纯的相关系数存入又一个excel文件中
#后将数据复制存入后缀为xlsx的文件中，方便可视化混淆矩阵时导入、传值
result = open('D://第4题：各行之间相关系数.xls', 'w', encoding='gbk')
title=['姓名','行号']
for i in range(len(result1)):
    title.append(str(i))
for i in range(len(title)):
    result.write(title[i])
    result.write('\t')
result.write('\n')
for m in range(len(result1)):
    result.write(data[m][1])
    result.write('\t')
    result.write(title[m+2])
    result.write('\t')
    for n in range(len(result1[m])):
        result.write(str(result1[m][n]))
        result.write('\t')
    result.write('\n')
result.close()


result = open('D://第4题：混淆矩阵所需导入数据.xls', 'w', encoding='gbk')
for m in range(len(result1)):
    for n in range(len(result1[m])):
        result.write(str(result1[m][n]))
        result.write('\t')
    result.write('\n')
result.close()


#打印相关矩阵
result2=np.matrix(result1)
print('相关矩阵“\n%s'%result2)


#可视化混淆矩阵
b=[]
fileName = "D://第4题：混淆矩阵需导入的数据.xlsx"
if os.path.exists(fileName):
    xls_file = xlrd.open_workbook(fileName)
    xlsxSheet = xls_file.sheets()[0]
    # 获取行数、列数
    nrows = int(xlsxSheet.nrows)
    colNum = int(xlsxSheet.colNum)
    for row in range(nrows):
        b.append(xlsxSheet.row_values(row))

confusion_matrix=b
df_cm=pd.DataFrame(confusion_matrix)
sn.heatmap(df_cm,vmax=1,vmin=-1)
plt.show()
