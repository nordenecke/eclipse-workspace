#encoding:utf-8
'''
Created on Sep 11, 2019
@author: eqhuliu
'''
import os
from xml.dom.minidom import Document
# import xml.dom.minidom
import xlsx_parser

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

INPUT_FILES_BASE_PATH = r'data\template'
XLSX_TEMPLATE_FILENAME = r'test_sample.xlsx'                
OUTPUT_FILES_BASE_PATH = r'data'
XML_MAP_FILENAME = r'qa-map.xml'                

class Xml_generator:
    def __init__(self,xlsx_file_name,xml_path):
        self.xlsx_template_path = xlsx_file_name
        self.xml_map_path = xml_path
        self.element_list = []

    def read_xlsx_template(self):
        txtfile = open(self.xlsx_template_path,"r",encoding='gbk',errors='ignore')
        self.element_list = txtfile.readlines()
        for i in self.element_list:
            oneline = i.strip().split(" ")
            if len(oneline) != 5:
                print("TxtError")

    def makexml(self):
        doc = Document()
        orderpack = doc.createElement("OrderPack")
        doc.appendChild(orderpack)
        objecname = "Order"
        for i in self.element_list:
            oneline = i.strip().split(" ")
            objectE = doc.createElement(objecname)
            objectE.setAttribute("number",oneline[0])

            objectcontent = doc.createElement("Content")
            objectcontenttext = doc.createTextNode(oneline[1])
            objectcontent.appendChild(objectcontenttext)
            objectE.appendChild(objectcontent)

            objectresult = doc.createElement("Result")
            objectresulttext = doc.createTextNode(oneline[2])
            objectresult.appendChild(objectresulttext)
            objectE.appendChild(objectresult)

            objectappname = doc.createElement("AppName")
            objectappnametext = doc.createTextNode(oneline[3])
            objectappname.appendChild(objectappnametext)
            objectE.appendChild(objectappname)

            objectdelay = doc.createElement("Delay")
            objectdelaytext = doc.createTextNode(oneline[4])
            objectdelay.appendChild(objectdelaytext)
            objectE.appendChild(objectdelay)

            orderpack.appendChild(objectE)

        f = open(self.xml_map_path, 'w')
        doc.writexml(f, indent='\t', newl='\n', addindent='\t', encoding='gbk')
        f.close()



if __name__ == "__main__":
    xlsx_template_file = os.path.abspath('.') + '\\' + INPUT_FILES_BASE_PATH +'\\'+ XLSX_TEMPLATE_FILENAME
    xml_map_file = os.path.abspath('.') + '\\' + OUTPUT_FILES_BASE_PATH +'\\'+ XML_MAP_FILENAME
    print(xlsx_template_file)
    print(xml_map_file)
    xp = Xlsx_parser(xlsx_template_file)
    xp.open()
#     xp.print_all_form_controls()
    xp.get_all_elements()
    xp.close()
    
    xml_gen =Xml_generator(xlsx_template_file,xml_map_file)
    xml_gen.read_xlsx_template()
    xml_gen.makexml()
    print(xml_gen.xlsx_template_path)
    for i in xml_gen.element_list:
        print(i)
