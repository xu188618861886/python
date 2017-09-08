import sys
sys.path.append("d:\\users\\administrator\\appdata\local\\programs\\python\\python36\\lib\\site-packages")
# 将包xlutils所在的位置加到搜索路径中
import xlrd   # 读取excel
import xlwt   # 写excel
import xlrd
import xlutils
from xlutils.copy import copy

path_read = 'C:\\Users\\zhixu\Desktop\\4A_foodtag.xlsx'
workbook_read = xlrd.open_workbook(path_read)  # 打开读入文件
read_sheet1 = workbook_read.sheet_by_index(0)    # 读入文件的sheet1
read_sheet2 = workbook_read.sheet_by_index(1)    # 读入文件的sheet2

path_write = 'C:\\Users\\zhixu\Desktop\\4B_foodtag.xlsx'
workbook_write = xlrd.open_workbook(path_write)  # 打开写入文件
workbook_write_new = copy(workbook_write)
write_sheet3 = workbook_write_new.get_sheet(2)    # 写入文件的sheet3
write_sheet4 = workbook_write_new.get_sheet(3)    # 写入文件的sheet4

row_scene = 1   # 场景的行数，从第二行开始
col_scene = 0     # 场景的列数，从第一列开始

row_recipe = 1     # 食谱名称的行数
col_recipe = 1     # 食谱名称的列数

row_match = 1      # 写入的表格中，匹配成功的行
col_match_scene = 0     # 写入表格中，匹配成功的列，场景
col_match_ID = 1     # 写入表格中，匹配成功的列，编号
col_match_recipe = 2     # 写入表格中，匹配成功的列，食谱

row_unmatch = 1     # 写入表格中，匹配失败的行
col_unmatch = 0     # 写入表格中，匹配失败的列，场景

tag = 0  # 标识符，如果场景找到的个数至少一个为1，反之为0

i = 1  # 食材的位置偏置

for row_scene in range(1, 13):   # range(1,216)
    cell_scene = read_sheet2.cell(row_scene, col_scene).value  # 读出场景，存在cell中
    for i in range(1, 3):
        cell_material = read_sheet2.cell(row_scene, col_scene+i).value   # 读出食材，存在cell_material中
        for row_recipe in range(1, 13123):  # range(1,13123)，将每个食材在清单中去寻找
            cell_ID = read_sheet1.cell(row_recipe, col_recipe - 1).value  # 读出食谱ID，存在cell_ID中
            cell_recipe = read_sheet1.cell(row_recipe, col_recipe).value  # 读出食谱名称，存在cell_recipe中
            match = cell_recipe.find(cell_material)      # 在食谱下查找食材，匹配返回开始的索引值，否则返回-1
            if match != -1:
                tag = 1     # 成功标志
                print(row_match, cell_scene, cell_ID, cell_recipe)
                write_sheet3.write(row_match, col_match_scene, cell_scene)
                write_sheet3.write(row_match, col_match_ID, cell_ID)
                write_sheet3.write(row_match, col_match_recipe, cell_recipe)
                row_match = row_match + 1

    if tag == 0:
        print(row_unmatch,col_unmatch,cell_scene)
        write_sheet4.write(row_unmatch,col_unmatch,cell_scene)
        row_unmatch = row_unmatch+1
    tag = 0

workbook_write_new.save(path_write)
