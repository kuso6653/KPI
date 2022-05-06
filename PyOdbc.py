# -*-coding:utf-8-*-
import time
import pyodbc
import pandas as pd


class Pyodbc:
    def __init__(self, ip, database, name, password):
        self.conn = pyodbc.connect(
            # "DRIVER={ODBC Driver 17 for SQL Server};SERVER=10.56.164.18;DATABASE=KPI;UID=sa;PWD=Chem123#")
            "DRIVER={ODBC Driver 17 for SQL Server};" + f"SERVER={ip};DATABASE={database};UID={name};PWD={password}")
        # self.conn.setencoding(encoding='utf8')

    def __GetConnect(self):
        self.cursor = self.conn.cursor()
        if not self.cursor:
            raise (NameError, "连接数据库失败")
        else:
            return self.cursor

    def ExecQuery(self, sql):
        cur = self.__GetConnect()
        cur.execute(sql)
        resList = cur.fetchall()

        return resList

    def ExecQuery2(self, sql):
        time.sleep(2)
        cur = self.__GetConnect()
        cur.execute(sql)
        resList = cur.fetchone()

        return resList

    def ExecNonQuery(self, sql, insert_str):
        cur = self.__GetConnect()
        cur.execute(sql, insert_str)
        self.conn.commit()

    def ExecNonQuery2(self, sql):
        time.sleep(1)
        cur = self.__GetConnect()
        cur.execute(sql)
        self.conn.commit()

    def CloseConnect(self):
        self.conn.close()
