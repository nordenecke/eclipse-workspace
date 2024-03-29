'''
Created on 2019年9月10日

@author: 静火
'''
import sqlite3
import creat_db
import get_file_list
import read_xlsx
import sys


def importData(path):
    # 数据库
    creat_db.createDataBase()
    database = sqlite3.connect("check.db")

    # 文件类型
    wildcard = ".xlsx"

    list = get_file_list.getFileList(path, wildcard, 1)

    nfiles = len(list[0])
    # 文件名
    file = list[0]
    # 时间
    time = list[1]
    # 省份
    province = list[2]
    # #文件类型
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
