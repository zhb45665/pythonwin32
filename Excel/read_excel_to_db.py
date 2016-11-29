#!/bin/python
# -*- coding:utf-8 -*-
'''
Created on 2016年11月2日

@author: zhb
'''
import os
import os.path
import sys
sys.path.append("C:/Users/zhb/workspace/win32/")
from Excel.read_excel_xld import *
rootdir = "D:/excel/pinganexcel/"
for parent,dirnames,filenames in os.walk(rootdir):
    for dirname in  dirnames:
        print("parent is:%s" % parent)
        print("dirname is: %s" % dirname)

    for filename in filenames:
        print("parent is:%s" % parent)
        print("filename  is:%s" % filename )
        print ("the full name of the file is:%s" % os.path.join(parent,filename))
        file = os.path.join(parent,filename)
        readexcel(file)