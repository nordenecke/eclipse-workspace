# -*- coding: utf-8 -*-
"""
Created on Thu Dec 20 15:54:44 2018

@author: eqhuliu
"""
import os
import io
import xml.etree.cElementTree as ET


    
class xml_parser(object):
    def __init__(self, xml_file):
        self.xml_file = xml_file
        
    def ET_parser_iter(self):
        vs_cnt = 0
        str_s = ''
        
        file_io = io.StringIO()
        xm = open(self.xml_file,'rb')
        
        print("Read [%s] finished.\nParsing..." % (os.path.abspath(self.xml_file)))
        d_question = {}
        d_obj = {}
        i = 0
        for event,elem in ET.iterparse(xm,events=('start','end')):
            if i >= 2:
                break        
            elif event == 'start':
                        if elem.tag == 'question':
                                d_eNB = elem.attrib
                        elif elem.tag == 'object':
                                d_obj = elem.attrib
            elif event == 'end' and elem.tag == 'smr':
                        i += 1
            elif event == 'end' and elem.tag == 'v':
                        file_io.write(d_eNB['id']+' '+d_obj['TimeStamp']+' '+d_obj['MmeCode']+' '+d_obj['id']+' '+\
                        d_obj['MmeUeS1apId']+' '+ d_obj['MmeGroupId']+' '+str(elem.text)+'\n')
                        vs_cnt += 1
            elem.clear()
        str_s = file_io.getvalue().replace(' \n','\r\n').replace(' ',',').replace('T',' ').replace('NIL','')    #写入解析后内容
        xm.close()
        file_io.close()
        return (str_s,vs_cnt)
    
    
    
    def ET_parser_iter1(self,gz):
        file_io = io.StringIO()
        xm = open(gz,'rb')
        for event,elem in ET.iterparse(xm,events=('start','end')):
            if event == 'start':
                if elem.tag == 'configuration':
#                     print("--------------------------------------------------")
                    d_configuration=configuration.configuration()
#user_info
                elif elem.tag == 'user_info':
                    d_user_info = configuration.user_info()
                elif elem.tag == 'domain':
                    domain = ''
                elif elem.tag == 'user':
                    user = ''
                elif elem.tag == 'passwords':
                    passwords = ''
#env_info                    
                elif elem.tag == 'env_info':
                    d_env_info = configuration.env_info()
                elif elem.tag == 'system':
                    system = configuration.system()
                    l_host=[]
                elif elem.tag == 'host':
                    host=configuration.host()
                elif elem.tag == 'chrome_location':
                    chrome_location=''
            elif event == 'end': 
                if elem.tag == 'configuration':
#                     print("--------------------------------------------------")
                    print("Configuration analysis done!")
#user_info
                elif elem.tag == 'user_info':
                    d_configuration.d_ui=d_user_info
                elif elem.tag == 'domain':
                    domain = elem.text
                    d_user_info.domain = domain
                elif elem.tag == 'user':
                    user = elem.text
                    d_user_info.user = user
                elif elem.tag == 'passwords':
                    passwords = elem.text
                    d_user_info.passwords = passwords
#env_info                    
                elif elem.tag == 'env_info':
                    d_configuration.d_ei=d_env_info
                elif elem.tag == 'system':
                    system_type= elem.get('os')
#                     print(system_type)
                    system.system_type=system_type
                    system.l_host=l_host
                    d_env_info.l_system.append(system)
                elif elem.tag == 'host':
                    host_name = elem.get('name')
#                     print(host_name)
                    host.host_name = host_name
                    l_host.append(host)
                elif elem.tag == 'chrome_location':
                    chrome_location = elem.text
                    host.chrome_location=chrome_location
                elem.clear()
        xm.close()
        file_io.close()
        return (d_configuration)


