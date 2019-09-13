#encoding:utf-8
'''
Created on Sep 11, 2019
@author: eqhuliu
'''
import os
import gc
from win32com.client import gencache
import utilities

INPUT_FILES_BASE_PATH = r'data\samples'
EXCEL_FILENAME = r'test_sample_1.xlsx'

if __name__=='__main__':
    
    wildcard = ".xlsx"

    file_list = utilities.getFileList(os.path.abspath('.') +'\\'+ INPUT_FILES_BASE_PATH, wildcard, 1)
    utilities.print_each_list(file_list)
    excel = gencache.EnsureDispatch('Excel.Application')
    wb = excel.Workbooks.Open(os.path.abspath('.') +'\\'+ INPUT_FILES_BASE_PATH + '\\' + EXCEL_FILENAME)
    ws = wb.Worksheets.Item(1)
    print('Shape count: %s' % len(ws.Shapes))
    for shape in ws.Shapes:
        if shape.Type == 8: # form control
            if 'Check Box' in shape.Name:
                print('%s: %s' % (shape.AlternativeText, shape.ControlFormat.Value))
            if 'Option Button' in shape.Name:
                print('%s: %s' % (shape.AlternativeText, shape.ControlFormat.Value))
            if 'Group Box' in shape.Name:
                print('%s:' % (shape.AlternativeText))
    print(ws.Name)
    print(ws.Range('C10'))
    wb.Close(True)
    del wb,ws
    gc.collect()
    
