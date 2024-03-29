'''
Created on 2019年9月10日

@author: 静火
'''
import sqlite3
import xlrd


class FileDispose(object):
    """docstring for FileDispose"""

    def __init__(self, file):
        super(FileDispose, self).__init__()
        '''初始化数据库实例'''
        self.conn = sqlite3.connect(file)
        self.cursor = self.conn.cursor()

    def __del__(self):
        '''释放数据库实例'''
        self.cursor.close()
        self.conn.close()

    '''数据库插入操作'''

    def insert(self, id, name, sex, age, score, addr):
        sql = 'insert into student(id,name,sex,age,score,addr) values (%d,\"%s\",\"%s\",\"%s\",\"%s\",\"%s\")' % (int(id), name, sex, age, score, addr)
        print(sql)
        self.cursor.execute(sql)
        self.conn.commit()

    '''读取Excel文件'''

    def readFile(self, file):
        data = xlrd.open_workbook(file)
        table = data.sheets()[2]
        for rowId in range(1, 100):
            row = table.row_values(rowId)
            if row:
                self.insert(rowId, row[0], row[1], row[2], row[3], row[4])


fd = FileDispose('F:/test.db')
fd.readFile('F:/excel.xlsx')
