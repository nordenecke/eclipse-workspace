#encoding:utf-8

'''
Created on Sep 12, 2019

@author: eqhuliu
'''
import os

def print_each_list(list_name,count=False,level=0):
    for element in list_name:
        if isinstance(element,list):
            print_each_list(element,count,level+1)
        else:
            if count:
                for tab in range(level):
                    print ("\t",end='')  
                print (element)  
            else:  
                print (element)  
                
def getFileList(directory, wildcard, recursion):
    curdir = os.path.abspath('.')
    os.chdir(directory)

    fileList = []

    exts = wildcard.split(" ")
    files = os.listdir(directory)
    for name in files:
        fullname = os.path.join(directory, name)
        if(os.path.isdir(fullname) & recursion):
            getFileList(fullname, wildcard, recursion)
        else:
            for ext in exts:
                if(name.endswith(ext)):
                    fileList.append(fullname)
    os.chdir(curdir)
    return fileList 

if __name__=='__main__':
    pass
