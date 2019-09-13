'''
Created on 2019��9��10��

@author: ����
'''

import os


def getFileList(directory, wildcard, recursion):
    os.chdir(directory)

    fileList = []
    check_province = []
    check_time = []
    file_type = []

    exts = wildcard.split(" ")
    files = os.listdir(directory)
    for name in files:
        fullname = os.path.join(directory, name)
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


# ת�뺯��
def changeCode(name):
    name = name.decode('GBK')
    name = name.encode('UTF-8')
    return name
