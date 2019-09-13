'''
Created on 2019��9��10��

@author: ����
'''
import sqlite3
from ref import creat_db
from ref import get_file_list
from ref import read_xlsx
import sys


def importData(path):
    # ���ݿ�
    creat_db.createDataBase()
    database = sqlite3.connect("check.db")

    # �ļ�����
    wildcard = ".xlsx"

    list = get_file_list.getFileList(path, wildcard, 1)

    nfiles = len(list[0])
    # �ļ���
    file = list[0]
    # ʱ��
    time = list[1]
    # ʡ��
    province = list[2]
    # #�ļ�����
    FileType = list[3]

    for count in range(0, nfiles):
        filename = file[count]
        check_province = get_file_list.changeCode(province[count])
        check_time = time[count]
        File_type = get_file_list.changeCode(FileType[count])
        read_xlsx.readExcel(filename, database, check_province, check_time, File_type)


if __name__ == '__main__':
    if len(sys.argv) != 2:
        print("Wrong Parameters")
    else:
        path = sys.argv[1]
        importData(path)
