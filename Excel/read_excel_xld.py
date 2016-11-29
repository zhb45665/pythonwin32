#!/bin/python
# -*- coding:utf-8 -*-
'''
Created on 2016年11月2日
@author: zhb
'''
import xlrd
import pymysql
def readexcel(filename):
    workbook = xlrd.open_workbook(filename)
    worksheets = workbook.sheet_names()
#     cxn = pymysql.connect(host='10.1.250.59',port=3306,user='root',passwd='root',db='pingan_record',charset='utf8')
    cxn = pymysql.connect(host='10.1.250.56',port=3306,user='root',passwd='root',db='zxin_customers_visit_mining_policy',charset='utf8')
    cur = cxn.cursor()
    print('worksheets is %s' %worksheets)
    for worksheet in worksheets:
        print(worksheet)
        sheet = workbook.sheet_by_name(worksheet)
        num_rows = sheet.nrows
        num_cols = sheet.ncols
        print("num_rows is %s" % num_rows)
        print("num_cols is %s" % num_cols)
        list = ["\'\'"]*num_cols
        for row in range(num_rows):
            for col,i in zip(range(num_cols),range(num_cols)):
                cell = sheet.cell_value(row,col)
                print(cell)
                list[i] = "'"+str(cell)+"'"
                print("[%s,%s] value is %s" %(row,col,cell))
                print("list is %s" % list[i])
            if num_cols == 4:
                list[1] = "'"+worksheet+"'"
                sql = "insert into haosenwei (id,quesstion,phone,area) values("+",".join(list)+")"
                print(sql)
                
                cur.execute(sql)
                cur.fetchall()
            elif num_cols == 3:
                quesstions = "'"+worksheet+"'"
                sql = "insert into haosenwei (id,phone,area,quesstion) values("+",".join(list)+","+quesstions+")"
                print(sql)
                cur.execute(sql)
                cur.fetchall()
            elif num_cols == 2:
#                 quesstions = "'"+worksheet+"'"
                area = "'"+worksheet+"'"
#                 sql = "insert into aa_copy (id,phone,area,quesstion) values("+",".join(list)+","+area+","+quesstions+")"
                sql = "insert into aa_copy (phone,quesstion,area) values("+",".join(list)+","+area+")"
                print(sql)
                cur.execute(sql)
                cur.fetchall()
            elif num_cols == 15:
                sql = "insert into 核身未通过件2013 (`批次号`,`序号`,`回访日期`,`客户号`,`客户姓名`,`发呼电话`,`录音文件号`,`发呼开始时间`,`发呼结束时间`\
                ,`通话时长`,`后续处理时间`,`回访结果`,`员工姓名`,`挂机方`,`变更新号码`) values("+",".join(list)+")"
                print(sql)
                cur.execute(sql)
                cur.fetchall()
#             sql = "insert into haosenwei (id,quesstion,phone,area) values("+",".join(list)+")"
#             print(sql)
#             cur.execute(sql)
#             cur.fetchall()
    cur.close()
    cxn.close()            
# readexcel("D:/excel/pinganexcel/核身未通过件.xlsx")