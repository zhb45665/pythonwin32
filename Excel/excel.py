#!/bin/python
# -*- coding:utf-8 -*-
'''
Created on 2016��10��28��

@author: zhb
'''
from tkinter import Tk
from time import sleep
from tkinter.messagebox import showwarning
import win32com.client as win32
from pywin.framework.app import App

warn = lambda app:showwarning(app,'Exit?')
RANGE = range(3,10)

def excel():
    app = 'Excel'
    xl = win32.gencache.EnsureDispatch('%s.Application' % app)
    ss = xl.Workbooks.Add()
    sh = ss.ActiveSheet
    xl.Visible = True
    sleep(1)
    
    sh.Cells(1,1).Value = 'Python to %s Demo' % app
#     sleep(1)
    
    for i in RANGE:
        sh.Cells(i,1).Value = 'Line %d' % i
        sleep(1)
    sh.Cells(i+3,1).Value = "Th-th-th-that's all folks"
        
    warn(app)
    ss.Close(False)
    xl.Application.quit()
        
if __name__ == '__main__':
    Tk().withdraw()
    excel()

    