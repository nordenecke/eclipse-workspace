#encoding:utf-8
'''
Created on Sep 11, 2019
@author: eqhuliu
'''
import os
import gc
from xml.dom.minidom import Document
from win32com.client import gencache
from enum import Enum

XML_TAG_SURVEY_DATA = 'survey_data'
XML_TAG_QUESTION ='question'
XML_TAG_QUESTION_ID = 'question_id'
XML_TAG_QUESTION_TYPE = 'question_type'
XML_TAG_SHAPE_GROUP = 'shape_group'
XML_ATTR_SHAPE_ID = 'shape_id'
XML_ATTR_SHAPE_ITEM_NUM = 'shape_item_num'
XML_TAG_SHAPE_ITEM = 'shape_item'
XML_TAG_SHAPE_ID = 'shape_id'
XML_TAG_SHAPE_TYPE = 'shape_type'

class Question_type(Enum):
    SINGLE_CHOICE = 1
    MULTIPLE_CHOICE = 2
    SINGLE_CHOICE_AND_REST_TEXT =3
    TEXT_CONTENT =3

class Form_control_type(Enum):
    RADIO_BUTTON = 1
    CHECK_BOX = 2
    TEXT_BOX =3

INPUT_FILES_BASE_PATH = r'data\template'
XLSX_TEMPLATE_FILENAME = r'test_sample.xlsx'                
OUTPUT_FILES_BASE_PATH = r'data'
XML_MAP_FILENAME = r'qa-map.xml'                


class Xml_generator:
    def __init__(self,xlsx_file_name,xml_path):
        self.xlsx_template_path = xlsx_file_name
        self.xml_map_path = xml_path
        self.doc = None
        self.element_list = []

    def open(self):
        self.doc= Document()
        
    def close(self):
        f = open(self.xml_map_path, 'w')
        self.doc.writexml(f, indent='\t', newl='\n', addindent='\t', encoding='utf-8')
        f.close()
        
    def add_element(self, parent_obj, element_name, element_value):
        element = self.doc.createElement(element_name)
        if None != element_value:
            element_text = self.doc.createTextNode(element_value)
            element.appendChild(element_text)
        if None != parent_obj:
            parent_obj.appendChild(element)
        else:
            self.doc.appendChild(element)
        return element
    
    def set_attribute(self, element, attribute_name, attribute_value):
        element.setAttribute(attribute_name, attribute_value)
        return element
    
    def set_attribute_by_id(self, element, attribute_name, attribute_value):
        element.setAttribute(attribute_name, attribute_value)
        return element

    def question_get_id(self, content_text):
        text_list=content_text.strip().split('.')
        question_id=text_list[0]
        print(question_id)
        return question_id
    
    def parse_xlsx_template_and_generate_xml(self):
        survey_data = self.add_element(None, XML_TAG_SURVEY_DATA, None)
        excel = gencache.EnsureDispatch('Excel.Application')
        wb = excel.Workbooks.Open(self.xlsx_template_path)
        ws = wb.Worksheets.Item(1)
        question = None
        question_type = None
        shape_group_items_num = 0
        shape_group = None
        for shape in ws.Shapes:
            if shape.Type == 8: # [MsoShapeType.msoFormControl]: form control
                if 'Check Box' in shape.Name:
                    print('%s %s' % (shape.AlternativeText, shape.ControlFormat.Value))
                    if None == question_type:
                        question_type = self.add_element(question, XML_TAG_QUESTION_TYPE, str(Question_type.MULTIPLE_CHOICE.value))
                    shape_group_items_num = shape_group_items_num + 1
#                     shape_group = self.set_attribute(shape_group, XML_ATTR_SHAPE_ITEM_NUM, shape_group_items_num)
                    shape_item = self.add_element(shape_group, XML_TAG_SHAPE_ITEM, None)
                    self.add_element(shape_item, XML_TAG_SHAPE_ID, str(shape.ID))
                    self.add_element(shape_item, XML_TAG_SHAPE_TYPE, str(Form_control_type.CHECK_BOX.value))
                if 'Option Button' in shape.Name:
                    print('%s %s' % (shape.AlternativeText, shape.ControlFormat.Value))
                    if None == question_type:
                        question_type = self.add_element(question, XML_TAG_QUESTION_TYPE, str(Question_type.SINGLE_CHOICE.value))
                    shape_group_items_num = shape_group_items_num + 1
#                     shape_group = self.set_attribute(shape_group, XML_ATTR_SHAPE_ITEM_NUM, shape_group_items_num)
                    shape_item = self.add_element(shape_group, XML_TAG_SHAPE_ITEM, None)
                    self.add_element(shape_item, XML_TAG_SHAPE_ID, str(shape.ID))
                    self.add_element(shape_item, XML_TAG_SHAPE_TYPE, str(Form_control_type.RADIO_BUTTON.value))
                if 'Group Box' in shape.Name:
                    print('%s' % (shape.AlternativeText))
                    #for previous group items
                    if shape_group_items_num != 0:
                        shape_group = self.set_attribute(shape_group, XML_ATTR_SHAPE_ITEM_NUM, str(shape_group_items_num))
                    question = self.add_element(survey_data, XML_TAG_QUESTION, None)
                    question_type = None
                    if None != self.question_get_id(shape.AlternativeText):
                        self.add_element(question, XML_TAG_QUESTION_ID, self.question_get_id(shape.AlternativeText))
                    else:
                        print('Error in question_get_id!')
                    shape_group = self.add_element(question, XML_TAG_SHAPE_GROUP, None)
                    shape_group = self.set_attribute(shape_group, XML_ATTR_SHAPE_ID, str(shape.ID))
                    shape_group_items_num = 0
        #for last one
        if shape_group_items_num != 0:
            shape_group = self.set_attribute(shape_group, XML_ATTR_SHAPE_ITEM_NUM, str(shape_group_items_num))
        print(survey_data)            
        wb.Close(True)
        del wb,ws
        gc.collect()


if __name__ == "__main__":
    xlsx_template_file = os.path.abspath('.') + '\\' + INPUT_FILES_BASE_PATH +'\\'+ XLSX_TEMPLATE_FILENAME
    xml_map_file = os.path.abspath('.') + '\\' + OUTPUT_FILES_BASE_PATH +'\\'+ XML_MAP_FILENAME
    print(xlsx_template_file)
    print(xml_map_file)
    
    xml_gen =Xml_generator(xlsx_template_file,xml_map_file)
    xml_gen.open()
    xml_gen.parse_xlsx_template_and_generate_xml()
    xml_gen.close()