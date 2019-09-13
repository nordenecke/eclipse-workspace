'''
Created on 2019年9月10日

@author: 静火
'''
from openpyxl import load_workbook  # 导入读取excel文件的模块
from openpyxl import Workbook  # 导入新建excel文件的模块

xls_read = load_workbook('pyxl_test.xlsx')  # 打开excel文件名为'pyxl_test.xlsx'
# print(xls_read.sheetnames)  # 查看工作表'pyxl_test.xlsx'中的所有sheet名，以列表形式生成

print(xls_read.active)  # 查看文件pyxl_test的活动中sheet
# xls_read.active.title = 'test'  # 将活动中的sheet名称变更为test
xls_read_sheet = xls_read.active  # 将活动中的sheet赋值给变量

# xls_read_sheet = xls_read.get_sheet_by_name('test') # 获取excel文件的某一个sheet

# print(xls_read_sheet['C'])    # 读取sheet中的C列所有数据，数据是以元组形式呈现
# print(xls_read_sheet['C4'].value) # 读取sheet中'C4'单元格的值
# print(xls_read_sheet.max_column)  # 查看sheet中最大的列合计值，统计依据是只要单元格含有值，就算一列
# print(xls_read_sheet.max_row)     # 查看sheet中最大的行合计值，统计依据是只要单元格含有值，就算一行

# b4 = xls_read_sheet['B4']   # 通过列号+行数 来定位某一个cell
# print(b4.value)   # 使用.value 来获取某一个cell的值

# print(xls_read_sheet.cell(column=2,row=4).value)  # 通过某sheet.cell(column=?,row=?).value 来取的某一个单元格的值

# xls_read_sheet.rows # sheet.rows是一个生成器，把每一行的内容形成一个元组
# for row in xls_read_sheet.rows:
#     print(row)
#     for cell in row:
#         print(cell.value)
#

# for column in xls_read_sheet.columns:   #sheet.columns是一个生成器，遍历每一列 一列的内容形成元组
#     print(list(column))
#     print(list(column)[0].value)
# for i in column:
#     if i.value:
#         print(i.value)
# print(i.value)

for row in range(1, xls_read_sheet.max_row + 1):
    for col in range(1, xls_read_sheet.max_column + 1):
        res = xls_read_sheet.cell(row=row, column=col)
        if res.value:
            print(res.value, end=' ')
            # print(xls_read_sheet.cell(row=row,column=col).value,end=' ')
    print()

xls_read.save('pyxl_test.xlsx')  # 保存该文件

print('=' * 40)

wb = Workbook()  # 新建了一张工作表，并默认创建了一张名叫'Sheet'的sheet，
print(wb)  # <openpyxl.workbook.workbook.Workbook object at 0x10515bc50>
print(wb.get_sheet_names())  # 显示wb工作表中所有的sheet，得到一个列表

wb.create_sheet('Data', index=1)  # 在wb工作表中新建一个名叫'Data'的sheet，该sheet的序号是1
# print(wb.get_sheet_names())
# del wb['Sheet']     #删除wb工作表中名叫'Sheet'的sheet
print(wb.get_sheet_names())
print(wb.active)  # 查看wb工作表中活动中的sheet
print(wb.active.values)  # 将该wb工作表中活动中的sheet的数据形成一个生成器

wb.active.title = 'test_sheet'  # 当前活动中的sheet更名
# print(wb.sheetnames)
wb.active['A1'] = 4
wb.active['B1'] = 2
wb.active['C1'] = '=AVERAGE(A1:B1)'  # 使用excel的公式（但是通过load_workbook data_only=True打开貌似也拿不到值）
wb.active['D1'] = '=A1*B1'

print(wb.active['A1'].value)
print(wb.active['B1'].value)
print(wb.active['C1'].value)
wb.save('pyxl_test1.xlsx')
