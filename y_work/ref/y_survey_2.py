#encoding:utf-8
'''
Created on Sep 11, 2019
@author: eqhuliu
'''
import os
import gc
from win32com.client import gencache

INPUT_FILES_BASE_PATH = r'data\samples'
EXCEL_FILENAME = r'test_sample_1.xlsx'


if __name__=='__main__':
    excel = gencache.EnsureDispatch('Excel.Application')
    wb = excel.Workbooks.Open(os.path.abspath('.') +'\\'+ INPUT_FILES_BASE_PATH + '\\' + EXCEL_FILENAME)
    ws = wb.Worksheets.Item(1)
    print('Shape count: %s' % len(ws.Shapes))
    for shape in ws.Shapes:
        if shape.Type == 8: # form control
            print('shape.ID=%d,shape.Name=%s,shape.AlternativeText=%s,shape.ControlFormat.Value=%s'
                  %(shape.ID,shape.Name,shape.AlternativeText,shape.hasattr(win32com, ControlFormat.Value)shape.ControlFormat.Value))
#             if 'Check Box' in shape.Name:
#                 print('%s: %s: %s' % (shape.Name,shape.AlternativeText, shape.ControlFormat.Value))
#             if 'Option Button' in shape.Name:
#                 print('%s: %s' % (shape.AlternativeText, shape.ControlFormat.Value))
#             if 'Group Box' in shape.Name:
#                 print('%s: %s/%d' % (shape.AlternativeText, shape.Name, shape.ID))
    print(ws.Name)
    print(ws.Range('C10'))
    wb.Close(True)
    del wb,ws
    gc.collect()
    
