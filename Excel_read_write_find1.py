#

import xlrd   #读取excel
import xlwt   #写excel
import xlrd
from xlutils.copy import copy


path_read='C:\\Users\\applee\\Desktop\\foodtag.xlsx'
workbook=xlrd.open_workbook(path_read)     # 打开文件
path_write='C:\\Users\\applee\\Desktop\\foodtag2.xlsx'
workbook2=xlrd.open_workbook(path_write)     # 打开文件
workbooknew = copy(workbook2)
ws = workbooknew.get_sheet(2)
ws_sheet4= workbooknew.get_sheet(3)

sheet1 = workbook.sheet_by_index(0)
sheet2 = workbook.sheet_by_index(1)
sheet3 = workbook.sheet_by_index(2)

row_scene=1     #场景的行数
col_scene=0     #场景的列数
cell_scene=sheet2.cell(row_scene,col_scene).value     #读出场景，存在cell中



row_effect=1     #食材功效的行数
col_effect=2     #食材功效的列数
cell_ID=sheet1.cell(row_effect,col_effect-2).value       #读出食材ID，存在cell中
cell_food=sheet1.cell(row_effect,col_effect-1).value     #读出食材名称，存在cell中
cell_effect=sheet1.cell(row_effect,col_effect).value     #读出食材功效，存在cell中

row_match=1
col_match_scene=0
col_match_ID=1
col_match_food=2
col_match_effect=3

tag=0   #标识符，如果场景找到的个数至少一个为1，反之为0

row_unmatch=1
col_unmatch=0

for row_scene in  range(1,216):   #range(1,216)
    cell_scene = sheet2.cell(row_scene, col_scene).value  # 读出场景，存在cell中
    for row_effect in range(1,13123):  #range(1,13123)
        cell_ID = sheet1.cell(row_effect, col_effect - 2).value  # 读出食材ID，存在cell中
        cell_food = sheet1.cell(row_effect, col_effect - 1).value  # 读出食材名称，存在cell中
        cell_effect = sheet1.cell(row_effect, col_effect).value  # 读出场景，存在cell中
        match=cell_effect.find(cell_scene);      #在食材功效下查找场景，匹配返回开始的索引值，否则返回-1
        if match!=-1:
            tag = 1
            print(row_match,cell_scene,cell_ID,cell_food,cell_effect)
            ws.write(row_match,col_match_scene,cell_scene)
            ws.write(row_match,col_match_ID,cell_ID)
            ws.write(row_match,col_match_food,cell_food)
            ws.write(row_match,col_match_effect,cell_effect)
            row_match = row_match + 1


    if tag==0:
        print(row_unmatch,col_unmatch,cell_scene)
        ws_sheet4.write(row_unmatch,col_unmatch,cell_scene)
        row_unmatch=row_unmatch+1
    tag=0

workbooknew.save(path)






