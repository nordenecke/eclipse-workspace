'''
Created on 2019��9��10��

@author: ����
'''
import sqlite3
import xlrd


class FileDispose(object):
    """docstring for FileDispose"""

    def __init__(self, file):
        super(FileDispose, self).__init__()
        '''��ʼ�����ݿ�ʵ��'''
        self.conn = sqlite3.connect(file)
        self.cursor = self.conn.cursor()

    def __del__(self):
        '''�ͷ����ݿ�ʵ��'''
        self.cursor.close()
        self.conn.close()

    '''���ݿ�������'''

    def insert(self, id, name, sex, age, score, addr):
        sql = 'insert into student(id,name,sex,age,score,addr) values (%d,\"%s\",\"%s\",\"%s\",\"%s\",\"%s\")' % (int(id), name, sex, age, score, addr)
        print(sql)
        self.cursor.execute(sql)
        self.conn.commit()

    '''��ȡExcel�ļ�'''

    def readFile(self, file):
        data = xlrd.open_workbook(file)
        table = data.sheets()[2]
        for rowId in range(1, 100):
            row = table.row_values(rowId)
            if row:
                self.insert(rowId, row[0], row[1], row[2], row[3], row[4])


fd = FileDispose('F:/test.db')
fd.readFile('F:/excel.xlsx')