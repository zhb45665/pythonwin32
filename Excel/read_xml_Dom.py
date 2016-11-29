#!/bin/python
# -*- coding:utf-8 -*-
'''
Created on 2016��10��31��

@author: zhb
'''
import sys,os
import pymysql
from xml.dom import minidom, Node
def read_dbconfig_xml(xml_file_path):
    cxn = pymysql.connect(host='10.1.250.56',port=3306,user='root',passwd='root',db='zxin_customers_visit_mining_policy',charset='utf8')
    cur = cxn.cursor()
    content = {}
    dom = minidom.parse(xml_file_path)
    print("hhhh ")
    root = dom.documentElement
    print(root)
    worksheets = root.getElementsByTagName('Worksheet')
    for worksheet in worksheets:
        print(worksheet)
        worksheet_name = worksheet.getAttribute("ss:Name")
        print(worksheet_name)
#     Table = worksheet.getElementsByTagName("Table")[0]
#     table_name = Table.getAttribute("ss:ExpandedColumnCount")
#     print(table_name)
        Rows = worksheet.getElementsByTagName("Row")
        for row in Rows:
            row_name = row.getAttribute("ss:Height")
            print("*********************************************")
            Cells = row.getElementsByTagName("Cell")
            max = 31
            list = ["\'\'"]*max
#             print(len(list))
            for Cell,i in zip(Cells,range(len(list))):
                
                Datas =Cell.getElementsByTagName("Data")
                if Datas:
                    for data in Datas:
#                         list[i] = "\'aa\'"
                        if data.firstChild:
                            datavalue = data.firstChild.data
                            datavalue = datavalue.replace("'","")
                            list[i] = "\'"+datavalue+"\'"
#                             print(datavalue)
#                         else:
#                             list[i] = 'null'
                            
                print("^^^^^^^^^^%s^^^^^^^^^^" % list[i])
                    
        
            print("####################################################")
            
#回访日结果清单_问卷表头            
#list[0] = 批次号
#list[1] = 序号
#list[2] = 回访时间
#list[3] = 机构代码
#list[4] = 机构名称
#list[5] = 服务人员代码
#list[6] = 服务人员姓名
#list[7] = 服务人员类型
#list[8] = 组
#list[9] = 组代码
#list[10] = 客户号
#list[11] = 客户姓名
#list[12] = 客户性别
#list[13] = 证件类型
#list[14] = 证件号码
#list[15] = 出生日期
#list[16] = 邮编
#list[17] = 联系地址
#list[18] = 属相
#list[19] = 交费银行
#list[20] = 受益人姓名
#list[21] = 手机号码
#list[22] = 外呼开始时间
#list[23] = 本次回访开始时间
#list[24] = 本次回访结束时间
#list[25] = 本次回访时长
#list[26] = 本次回访次数
#list[27] = 是否回访结束
#list[28] = 本次回访结果
#list[29] = 问卷结果备注
#list[30] = 录音文件号
#list[31] = 客户意见及建议
#list[32] = 问题4备注
#list[33] = 问题5备注
#list[34] = 员工姓名
#list[35] = 挂机方
#list[36] = 变更新号码
#             sql = "insert into test_copy (`批次号`,`序号`,`回访时间`,`机构代码`,`机构名称`,`服务人员代码`,`服务人员姓名`,`服务人员类型`,`组`,`组代码`,`客户号`,\
#             `客户姓名`,`客户性别`,`证件类型`,`证件号码`,`出生日期`,`邮编`,`联系地址`,`属相`,`交费银行`,`受益人姓名`,`手机号码`,`外呼开始时间`,`本次回访开始时间`,\
#             `本次回访结束时间`,`本次回访时长`,`本次回访次数`,`是否回访结束`,`本次回访结果`,`问卷结果备注`,`录音文件号`,`客户意见及建议`,`问题4备注`,`问题5备注`,\
#             `员工姓名`,`挂机方`,`变更新号码`) \
#             values ("+list[0]+","+list[1]+","+list[2]+","+list[3]+","+list[4]+","+list[5]+","+list[6]+","+list[7]+",\
#             "+list[8]+","+list[9]+","+list[10]+","+list[11]+","+list[12]+","+list[13]+","+list[14]+","+list[15]+",\
#             "+list[16]+","+list[17]+","+list[18]+","+list[19]+","+list[20]+","+list[21]+","+list[22]+","+list[23]+",\
#             "+list[24]+","+list[25]+","+list[26]+","+list[27]+","+list[28]+","+list[29]+","+list[30]+","+list[31]+",\
#             "+list[32]+","+list[33]+","+list[34]+","+list[35]+","+list[36]+")"
#             sql = "insert into 回访日结果清单2013 (`批次号`,`序号`,`回访时间`,`机构代码`,`机构名称`,`服务人员代码`,`服务人员姓名`,`服务人员类型`,`组`,`组代码`,`客户号`,\
#             `客户姓名`,`客户性别`,`证件类型`,`证件号码`,`出生日期`,`邮编`,`联系地址`,`属相`,`交费银行`,`受益人姓名`,`手机号码`,`外呼开始时间`,`本次回访开始时间`,\
#             `本次回访结束时间`,`本次回访时长`,`本次回访次数`,`是否回访结束`,`本次回访结果`,`问卷结果备注`,`录音文件号`,`客户意见及建议`,`问题4备注`,`问题5备注`,\
#             `员工姓名`,`挂机方`,`变更新号码`) \
#             values ("+",".join(list)+")"
#             sql = ','.join(list)
#             sql = "insert into 接通件明细2013 (`批次号`,`序号`,`回访日期`,`客户号`,`客户姓名`,`发呼电话`,`录音文件号`,`发呼开始时间`,`发呼结束时间`\
#                 ,`通话时长`,`后续处理时间`,`回访结果`,`员工姓名`,`挂机方`,`变更新号码`) values("+",".join(list)+")"
#           
#             sql = "insert into 核身未通过件2013 (`批次号`,`序号`,`回访日期`,`客户号`,`客户姓名`,`发呼电话`,`录音文件号`,`发呼开始时间`,`发呼结束时间`\
#                 ,`通话时长`,`后续处理时间`,`回访结果`,`员工姓名`,`挂机方`,`变更新号码`) values("+",".join(list)+")"
#             sql = "insert into 大魔法师  (id,`电话号码`,`地区`,`归属运营商`,`编号`,`当前品牌`,`当前机型`,`备注`,`导入时间`,`导入批次`,`数据类型`\
#             ,`开始时间`,`结束时间`,`通话时长`,`流水号`,`工号`,`员工姓名`,`外呼结果`,`一级营销结果`,`二级营销结果`,`营销备注`) values("+",".join(list)+")"
            sql = "insert into 11月  values("+",".join(list)+")"
            print(sql)    
            cur.execute(sql)
            cur.fetchall()
    cur.close()
    cxn.close()       
#     Cell = worksheet.getElementsByTagName("Data")[1]
#     cell_name = Cell.childNodes[0].nodeValue
#     print(cell_name)
   

    
# read_dbconfig_xml("D:/excel/pinganexcel/核身未通过件.xlsx")