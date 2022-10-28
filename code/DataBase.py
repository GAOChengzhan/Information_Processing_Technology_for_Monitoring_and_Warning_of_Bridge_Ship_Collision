# -*- coding: utf-8 -*-
"""
Created on Sun Aug 19 11:27:35 2018

@author: wt154
"""

import sqlite3
import pprint

class DataBase:
    def __init__(self,address):  #初始化，连接数据库
        self.conn = sqlite3.connect(address)
        print('Opened database successfully')
        
    def createTable(self):  #创建新表
        cur = self.conn.cursor()
        tblname = input('请输入表名：')
        sql = 'CREATE TABLE ' + tblname + '('
        while True:
            name = input('请输入列名：')
            datatype = input('请输入数据类型：')
            key = input('是否作为主键，输入T或S：')
            null = input('是否允许为空，输入T或S：')
            close = input('是否继续添加新列，输入T或S：')
            sql = sql + name + ' ' + datatype + ' '
            if key == 'T':
                sql = sql + 'PRIMARY KEY '
            if null == 'S':
                sql = sql + 'NOT NULL'
            if close == 'S':
                sql = sql + ');'
                break
            else:
                sql = sql + ',' + '\n'
        print(sql)
        cur.execute(sql)
        print('Table created successfully')
        self.conn.commit()
        
    def getColumns(self,tblName):  #获取指定表的所有列名，以元组形式返回
        cur = self.conn.cursor()
        cur.execute("SELECT * FROM {}".format(tblName))
        col_name_list = [tuple[0] for tuple in cur.description]
        pprint.pprint(col_name_list)
        return col_name_list
    
    def get_Col_Information(self,tblName):  #获取指定表的结构信息
        cur = self.conn.cursor()
        cur.execute("PRAGMA table_info({})".format(tblName))
        all_data = cur.fetchall()
        pprint.pprint(all_data)
        return all_data
    
    def getAllData(self,tblName):  #获取指定表的所有数据，以列表形式返还
        cur = self.conn.cursor()
        sql = 'SELECT * FROM ' + tblName
        cursor = cur.execute(sql)
        col_name_list = self.getColumns(tblName)
        l = len(col_name_list)
        result = []
        for row in cursor:
            result_i = []
            for i in range(0,l):
                result_i.append(row[i])
            result.append(result_i)
        return result
    
    def insertData(self,tblName,values):  #向指定表插入数据，参数values为列表，tblName为字符串,比如'COMPANY'
        cur = self.conn.cursor()                         
        print("Opened database successfully")
        col_name_list = self.getColumns(tblName)
        l = len(col_name_list)
        col_name_information = self.get_Col_Information(tblName)
        sql = 'INSERT INTO ' + tblName + ' ('
        for i in range(0,l-1):
            sql = sql + col_name_list[i] + ','
        sql = sql + col_name_list[l-1] + ') VALUES ('
        for i in range(0,l-1):
            if col_name_information[i][2] == 'TEXT':
                sql = sql + '"' + values[i] + '",'
            elif col_name_information[i][2] == 'CHAR(50)':
                sql = sql + '"' + values[i] + '",'
            else:
                sql = sql + values[i] + ','
        if col_name_information[l-1][2] == 'TEXT':
            sql = sql + '"' + values[l-1] + '")'
        elif col_name_information[i][2] == 'CHAR(50)':
            sql = sql + '"' + values[l-1] + '")'
        else:
            sql = sql + values[l-1] + ')'
        print(sql)
        cur.execute(sql)
        self.conn.commit()
        print("Records created successfully")
        
    def selectData(self,col_name_list,tblName,restriction):  #查询数据,返回二维列表，col_name_list为列表
        cur = self.conn.cursor()
        print("Opened database successfully")
        l = len(col_name_list)
        sql = 'SELECT '
        for i in range(0,l-1):
            sql = sql + col_name_list[i] + ','
        sql = sql + col_name_list[l-1] + ' FROM ' + tblName
        sql = sql + ' WHERE ' + restriction  #restriction为字符串，比如'id=4'
#        print(sql)
        cursor = cur.execute(sql)
        result = []
        for row in cursor:
            result_i = []
            for i in range(0,l):
                result_i.append(row[i])
            result.append(result_i)
        return result
    
    def updateData(self,tblName,change,restriction):  #更新数据，change为需要变更的内容，比如'salary=70000'
        cur = self.conn.cursor()
        print("Opened database successfully")
        sql = 'UPDATE ' + tblName + ' SET ' + change
        sql = sql + ' WHERE ' + restriction
        print(sql)
        cur.execute(sql)
        self.conn.commit()
        
    def deleteData(self,tblName,restriction):  #删除数据
        cur = self.conn.cursor()
        print("Opened database successfully")
        sql = 'DELETE FROM ' + tblName + ' WHERE ' + restriction
        cur.execute(sql)
        self.conn.commit()
        
    def closeDataBase(self):  #关闭数据库连接
        self.conn.close()
        print('Close database successfully')
  
    
address = r'C:\Users\DELL\Desktop\自动报告尝试1\docsample\testtest.db'
df = DataBase(address)
tblName = '船只信息'
#values=['2','Mom','30','Guangzhou','30000']
#df.insertData(tblName,values)
result = df.getColumns(tblName)
result=result[0:1]
print("sdfsd")
print(result)
df.closeDataBase()












