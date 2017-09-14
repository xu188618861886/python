#根据用户对食谱的收藏次数，得到食材的收藏次数，作为食材的关注程度
# -*- coding: utf-8 -*-
import sys
sys.path.append("d:\\users\\administrator\\appdata\local\\programs\\python\\python36\\lib\\site-packages")
# 将包xlutils所在的位置加到搜索路径中
import xlrd   # 读取excel
import xlwt   # 写excel
import xlutils  #复制和修改Excel文件
from xlutils.copy import copy

path_read1 = 'C:\\Users\\applee\Desktop\\食材清单.xlsx'
workbook_read1 = xlrd.open_workbook(path_read1)  # 打开读入文件
read1_sheet1 = workbook_read1.sheet_by_index(0)    # 读入文件1的sheet1
read1_sheet1_material = read1_sheet1.col(1)      #文件1中sheet1的第二列，食材
read1_sheet1_material_ID = read1_sheet1.col(0)      #文件1中sheet1的第二列，食材ID

path_read2 = 'C:\\Users\\applee\Desktop\\食谱-食材-评论.xlsx'
workbook_read2 = xlrd.open_workbook(path_read2)  # 打开读入文件
read2_sheet1 = workbook_read2.sheet_by_index(0)    # 读入文件1的sheet1
read2_sheet1_reptile = read2_sheet1.col(1)      #文件1中sheet1的食谱
read2_sheet1_enshrine = read2_sheet1.col(2)      #文件1中sheet1的食谱收藏次数


path_write1 = 'C:\\Users\\applee\Desktop\\食材清单B.xlsx'
workbook_write1 = xlrd.open_workbook(path_write1)  # 打开写入文件
workbook_write1_new = copy(workbook_write1)
#先复制一份Sheet然后再次基础上追加并保存到一份新的Excel文档中去
write_sheet1 = workbook_write1_new.get_sheet(0)    # 写入文件的sheet1
total_enshrine = 0     #某种食材收藏次数的总和
write_total_row = 1


for number_material in range (1,476):

    read1_sheet1_material_element = read1_sheet1_material[number_material].value      #食材单元

    for number_cell_reptile_material in range(1,13138):

        cell_reptile_material_element = read2_sheet1.cell(number_cell_reptile_material, 3).value  #文件1中sheet1的食谱食材
        read2_sheet1_enshrine_element = read2_sheet1_enshrine[number_cell_reptile_material].value     #食谱食材所对应的收藏次数
        match = cell_reptile_material_element.find(read1_sheet1_material_element)      # 在食谱下查找食材，匹配返回开始的索引值，否则返回-1
        if match != -1:
            total_enshrine = total_enshrine + read2_sheet1_enshrine_element

    write_sheet1.write(write_total_row,2,total_enshrine)
    write_total_row = write_total_row +1
    print(number_material,total_enshrine)
    total_enshrine = 0  # 某种食材收藏次数的总和

workbook_write1_new.save(path_write1)