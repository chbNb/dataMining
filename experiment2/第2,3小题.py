import matplotlib.pyplot as plt
import os
import xlrd
import numpy as npy
import math
from pylab import mpl

mpl.rcParams['font.sans-serif'] = ['SimHei']
mpl.rcParams['axes.unicode_minus'] = False

# 将实验一处理后的数据导入data
data = []
fileName = "final_data.xlsx"
if os.path.exists(fileName):
    xlsxFile = xlrd.open_workbook(fileName)
    xlsxSheet = xlsxFile.sheets()[0]
    # 获取行数、列数
    nrows = int(xlsxSheet.nrows)
    ncols = int(xlsxSheet.ncols)
    for row in range(1, nrows):
        data.append(xlsxSheet.row_values(row))


# 将值为“null”的成绩赋值为0
for i in range(len(data)):
    for j in range(5, 14):
        if (data[i][j] == 'null'):
            data[i][j] = 0


# 将课程1成绩和体能成绩分别放进两个列表中
    #这里设定体能成绩的bad、general、good和excellent对应的分数分别为25、50、75、100
class1 = []
strong = []

for i in range(len(data)):
    class1.append(data[i][5])

for i in range(len(data)):
    if (data[i][14] == 'excellent'):
        strong.append(100)
    elif (data[i][14] == 'good'):
        strong.append(75)
    elif (data[i][14] == 'general'):
        strong.append(50)
    elif (data[i][14] == 'bad'):
        strong.append(25)
    else:
        strong.append(0)


##第2题：绘制课程1的成绩直方图

plt.title('课程1--成绩直方图',fontsize=20,color='blue')
plt.xlabel('课程1成绩',fontsize=15,color='r')
plt.ylabel('在该成绩区间的人数',fontsize=12,color='r')
my_x=npy.arange(0,100,5)
plt.xticks(my_x)

jiange=[]
for i in range(0,100,5):
    jiange.append(i)
n, bins, patches = plt.hist(class1,jiange)
plt.show()


# 第3题：对每门成绩进行z-score归一化，得到归一化的数据矩阵
# 初始化列表，分别是成绩和，均值，平方和，方差

sum1 = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
average = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
squ = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
variance = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0]

# # 求各科成绩总和，然后除以列表data总长度得到10个成绩的均值
for i in range(len(data)):
    for j in range(5, 14):
        sum1[j - 5] += int(data[i][j])
for i in range(len(data)):
    if (data[i][14] == 'excellent'):
        sum1[9] += 100
    elif (data[i][14] == 'good'):
        sum1[9] += 75
    elif (data[i][14] == 'general'):
        sum1[9] += 50
    elif (data[i][14] == 'bad'):
        sum1[9] += 25
    else:
        sum1[9] += 0
for i in range(len(sum1)):
    average[i] = sum1[i] / len(data)

# # 求各科成绩与均值差的平方和，然后除以列表data总长度，得到10个成绩的方差
for i in range(len(data)):
    for j in range(5, 14):
        squ[j - 5] += (int(data[i][j]) - average[j - 5]) ** 2

for i in range(len(data)):
    if (data[i][14] == 'excellent'):
        squ[9] += (95 - average[9]) ** 2
    elif (data[i][14] == 'good'):
        squ[9] += (85 - average[9]) ** 2
    elif (data[i][14] == 'general'):
        squ[9] += (75 - average[9]) ** 2
    elif (data[i][14] == 'bad'):
        squ[9] += (65 - average[9]) ** 2
    else:
        squ[9] += (0 - average[9]) ** 2

for i in range(len(squ)):
    variance[i] = math.sqrt(squ[i] / len(data))


# # 进行Z-score归一化：（初值-均值）/ 方差，保留4位小数
for i in range(len(data)):
    for j in range(5, 14):
        data[i][j] = "%.4f" % ((float(data[i][j]) - average[j - 5]) / variance[j - 5])
for i in range(len(data)):
    if (data[i][14] == 'excellent'):
        data[i][14] = "%.4f" % ((95 - average[9]) / variance[9])
    elif (data[i][14] == 'good'):
        data[i][14] = "%.4f" % ((85 - average[9]) / variance[9])
    elif (data[i][14] == 'general'):
        data[i][14] = "%.4f" % ((75 - average[9]) / variance[9])
    elif (data[i][14] == 'bad'):
        data[i][14] = "%.4f" % ((65 - average[9]) / variance[9])
    else:
        data[i][14] = "%.4f" % ((0 - average[9]) / variance[9])


# # 分别将每一行的9门学习成绩和体能成绩放进嵌套列表这种对应位置的元素中
# # 然后进行列表转矩阵，并打印数据矩阵
# # 将列表a1写入excel文件中，方便查看验证
a1 = []
for i in range(len(data)):
    a1.append([])
for i in range(len(data)):
    for j in range(5, 14):
        a1[i].append(data[i][j])
    a1[i].append(data[i][14])

a2 = npy.matrix(a1)
print("归一化的数据矩阵：\n%s" % a2)

result = open('D://第3题：归一化后列表数据.xls', 'w', encoding='gbk')
title = ['lesson1', 'lesson2', 'lesson3', 'lesson4', 'lesson5', 'lesson6', 'lesson7', 'lesson8', 'lesson9', '体能成绩']
for i in range(len(title)):
    result.write(title[i])
    result.write('\t')
result.write('\n')
for m in range(len(a1)):
    for n in range(len(a1[m])):
        result.write(str(a1[m][n]))
        result.write('\t')
    result.write('\n')
result.close()
