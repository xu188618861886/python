# 筛选出表格2中所包含的表格1中的食谱。
import sys
sys.path.append("d:\\users\\administrator\\appdata\local\\programs\\python\\python36\\lib\\site-packages")
# 将包xlutils所在的位置加到搜索路径中
import xlrd   # 读取excel
import xlwt   # 写excel
import xlrd
import xlutils
from xlutils.copy import copy

path_read1 = 'C:\\Users\\applee\Desktop\\A1.xlsx'
workbook_read1 = xlrd.open_workbook(path_read1)  # 打开读入文件
read_sheet1 = workbook_read1.sheet_by_index(0)    # 读入文件的sheet1

path_read2 = 'C:\\Users\\applee\Desktop\\食谱.xlsx'
workbook_read2 = xlrd.open_workbook(path_read2)  # 打开读入文件
read_sheet2 = workbook_read2.sheet_by_index(0)    # 读入文件的sheet1

path_write1='C:\\Users\\applee\\Desktop\\A2.xlsx'
workbook_write1=xlrd.open_workbook(path_write1)     # 打开文件
workbooknew1 = copy(workbook_write1)
workbook_write1_sheet1 = workbooknew1.get_sheet(0)

row_category1=0     #食谱1的行数
col_category1=0     #食谱1的列数
cell_category1=read_sheet1.cell(row_category1,col_category1).value     #读出食谱1，存在cell中
cell_category1

row_category2=1     #食谱2的行数
col_category2=1     #食谱2的列数
cell_category2=read_sheet2.cell(row_category2,col_category2).value     #读出食谱2，存在cell中
cell_category2

row_write1=0
col_write1_cell_category2_left=0
col_write1_cell_category2=1
col_write1_cell_category1_right=2

for row_category1 in range(0,2376):
    cell_category1 = read_sheet1.cell(row_category1, col_category1).value  # 读出食谱1，存在cell中
    cell_category1_right = read_sheet1.cell(row_category1, col_category1+1).value  # 读出食谱1，存在cell中
    for row_category2 in range(1,13138):
        cell_category2 = read_sheet2.cell(row_category2, col_category2).value  # 读出食谱2，存在cell中
        cell_category2_left = read_sheet2.cell(row_category2, col_category2-1).value  # 读出食谱ID，存在cell中
        match = cell_category1.find(cell_category2);  # 查找，匹配返回开始的索引值，否则返回-1
        if (match != -1):
            print(cell_category2_left,cell_category2,cell_category1_right)
            workbook_write1_sheet1.write(row_write1, col_write1_cell_category2_left, cell_category2_left)
            workbook_write1_sheet1.write(row_write1, col_write1_cell_category2, cell_category2)
            workbook_write1_sheet1.write(row_write1,col_write1_cell_category1_right ,cell_category1_right)
            row_write1 =row_write1 + 1
        # else:
        #     print("****************************************")
workbooknew1.save(path_write1)