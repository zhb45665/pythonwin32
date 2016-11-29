#!/bin/python
# -*- coding:utf-8 -*-
'''
Created on 2016年11月10日

@author: zhb
'''
import os
import os.path
import shutil
rootdir = "E:/平安录音/2013/record_pingan_13/"
path = "E:/平安录音/2013/2013/"
for parent,dirnames,filenames in os.walk(rootdir):
    for dirname in  dirnames:
        print("parent is:%s" % parent)
        print("dirname is: %s" % dirname)

    for filename in filenames:
#         print("parent is:%s" % parent)
#         if filename.endswith(".wav"):
#             print("filename  is:%s" % filename )
#             file = os.path.join(parent,filename)
#             print("file is: %s " % file)
#             shutil.copy2(file, path+filename)
        print("filename  is:%s" % filename )
        file = os.path.join(parent,filename)
        print("file is: %s " % file)
        shutil.copy(file,os.path.join(path,filename))
#         print ("the full name of the file is:%s" % os.path.join(parent,filename))
        