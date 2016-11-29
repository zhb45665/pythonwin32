#!/bin/python
# -*- coding:utf-8 -*-
'''
Created on 2016��10��31��

@author: zhb
'''
import pymysql
import xml.sax
class  xmlHandler(xml.sax.ContentHandler):
    def __init__(self):
        self.Workbook = ""
        self.Worksheet = ""
        self.Table = ""
        self.Row = ""
        self.Cell = ""
        self.Data = ""
        
    def startElement(self, tag, attrs):
        self.CurrentData = tag
        if tag == "Worksheet":
            title = attrs["ss:Name"]
            print(title)

    def endElement(self, name):
#         xml.sax.ContentHandler.endElement(self, name)
        if self.CurrentData == "Cell":
            print(self.Cell)
    
    def characters(self, content):
        if self.CurrentData == "Cell":
            self.Cell = content
#         xml.sax.ContentHandler.characters(self, content)

            
          
        
if ( __name__ == "__main__"):
    parser = xml.sax.make_parser()
    parser.setFeature(xml.sax.handler.feature_namespaces, 0)
    Handler = xmlHandler()
    parser.setContentHandler(Handler)
    parser.parse(r"D:\excel\1111\回访日结果清单-问卷五.xls")
            
