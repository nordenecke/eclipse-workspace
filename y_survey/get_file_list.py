'''
Created on 2019年9月10日

@author: 静火
'''

import os


def getFileList(dir, wildcard, recursion):
    os.chdir(dir)

    fileList = []
    check_province = []
    check_time = []
    file_type = []

    exts = wildcard.split(" ")
    files = os.listdir(dir)
    for name in files:
        fullname = os.path.join(dir, name)
        if(os.path.isdir(fullname) & recursion):
            getFileList(fullname, wildcard, recursion)
        else:
            for ext in exts:
                if(name.endswith(ext)):
                    fileList.append(name)
                    check_province.append(name.split('-')[1])
                    check_time.append(name.split('-')[0])
                    file_type.append(name.split('-')[2])
    return fileList, check_time, check_province, file_type


# 转码函数
def changeCode(name):
    name = name.decode('GBK')
    name = name.encode('UTF-8')
    return name
