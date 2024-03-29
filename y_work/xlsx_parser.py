#encoding:utf-8

'''
Created on Sep 12, 2019

@author: eqhuliu
'''

'''
XlFormControl enumeration (Excel)
Specifies the type of the form control.
Name                Value            Description
xlButtonControl      0                Button.
xlCheckBox           1                Check box.
xlDropDown           2                Combo box.
xlEditBox            3                Text box.
xlGroupBox           4                Group box.
xlLabel              5                Label.
xlListBox            6                List box.
xlOptionButton       7                Option button.
xlScrollBar          8                Scroll bar.
xlSpinner            9                Spinner.

print('%s' % (shape.FormControlType)) == 4 means xlGroupBox
'''
import gc
import os
import json
from win32com.client import gencache

class Xlsx_parser(object):
    def __init__(self,xlsx_file_name):
        self.xlsx_file = xlsx_file_name
        self.elements_info = {}
        self.wb = None
        self.ws = None
        
    def open(self):
        excel = gencache.EnsureDispatch('Excel.Application')
        self.wb = excel.Workbooks.Open(self.xlsx_file)
        self.ws = self.wb.Worksheets.Item(1)
        
    def close(self):
        self.wb.Close(True)
        del self.wb,self.ws
        gc.collect()
        
    def get_element_value_by_id(self,ID):
        for shape in self.ws.Shapes:
            if shape.Type == 8: # [MsoShapeType.msoFormControl]: form control
                if ID == shape.ID:
                    print('%s: %s' % (shape.AlternativeText, shape.ControlFormat.Value))
                    return shape.ControlFormat.Value, shape.AlternativeText
        return None
    
    def get_element_value_by_text(self,alternative_text):
        for shape in self.ws.Shapes:
            if shape.Type == 8: # [MsoShapeType.msoFormControl]: form control
                if alternative_text == shape.AlternativeText:
                    print('%s: %s' % (shape.AlternativeText, shape.ControlFormat.Value))
                    return shape.ControlFormat.Value, shape.ID
        return None
    
    def print_all_form_controls(self):
        for shape in self.ws.Shapes:
            if shape.Type == 8: # [MsoShapeType.msoFormControl]: form control
                print('shape.ID=[%d] shape.Name=[%s]'%(shape.ID,shape.Name))
                if 'Check Box' in shape.Name:
                    print('%s %s' % (shape.AlternativeText, shape.ControlFormat.Value))
                if 'Option Button' in shape.Name:
                    print('%s %s' % (shape.AlternativeText, shape.ControlFormat.Value))
                if 'Group Box' in shape.Name:
                    print('%s' % (shape.AlternativeText))
                    print(help(shape))
                    print(dir(shape))
    
    def get_all_elements_to_json(self):                   
        for shape in self.ws.Shapes:
            if shape.Type == 8: # [MsoShapeType.msoFormControl]: form control
                if 'Check Box' in shape.Name:
                    print('%s %s' % (shape.AlternativeText, shape.ControlFormat.Value))
                if 'Option Button' in shape.Name:
                    print('%s %s' % (shape.AlternativeText, shape.ControlFormat.Value))
                if 'Group Box' in shape.Name:
                    print('%s' % (shape.AlternativeText))
                    print(help(shape))
                    print(dir(shape))
                    
        self.elements_info = json.dumps(data, ensure_ascii=False)
        print(self.elements_info)
        return self.elements_info
    
#     def get_question_item_index_and_value_by_group_box(self,shape_group_box):
#         pass
                        
INPUT_FILES_BASE_PATH = r'data\template'
XLSX_TEST_FILENAME = r'test_sample.xlsx'  

if __name__=='__main__':
    xlsx_test_file = os.path.abspath('.') + '\\' + INPUT_FILES_BASE_PATH +'\\'+ XLSX_TEST_FILENAME
    xp = Xlsx_parser(xlsx_test_file)
    xp.open()
    xp.print_all_form_controls()
    xp.close()
        