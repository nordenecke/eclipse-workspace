'''
Created on 2019��9��10��

@author: ����
'''
import xlrd
import xlwt
from datetime import date, datetime


def read_excel():
    # ���ļ�
    workbook = xlrd.open_workbook(r'file.xls')
    # ��ȡ����sheet
    sheet_name = workbook.sheet_names()[0]
    sheet = workbook.sheet_by_name(sheet_name)
    
    # ��ȡһ�е�����
    for i in range(6, sheet.nrows):
        for j in range(0, sheet.ncols):
            print(sheet.cell(i, j).value.encode('utf-8'))


if __name__ == '__main__':
    read_excel()