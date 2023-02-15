# FatMonkey Python
# 学python搞钱
# 时间: 2023/1/12 22:01

import os
import xlrd

excle_path = os.path.join(os.path.dirname(__file__),'data/test_data.xlsx')
# print(excle_path)

wb = xlrd.open_workbook(excle_path)  #创建工作簿对象
sheet = wb.sheet_by_index(0)    #创建表格对象
cell_value = sheet.cell_value(2,1)
# print(cell_value)

merged = sheet.merged_cells
def get_merged_cell_value(row_index,col_index):
    cell_value = None
    for(rlow,rhigh,clow,chigh) in merged:
        if (row_index>= rlow and row_index<rhigh):
            if (col_index>= clow and col_index<chigh):
                cell_value =sheet.cell_value(rlow,clow)
                break
            else:
                cell_value = sheet.cell_value(row_index,col_index)
        else:
            cell_value = sheet.cell_value(row_index,col_index)
    return cell_value

# print(get_merged_cell_value(4,0))
for i in range(1,9):
    print(get_merged_cell_value(i,0))