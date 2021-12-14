from cx_Oracle import *


class OracleHelper:
    def __init__(self, username, password, host):
        self.username = username
        self.password = password
        self.host = host

    def connect_open(self):
        self.conn = connect(self.username, self.password, self.host)
        # 连链接数据库cx_Oracle.connect("hr", "hrpwd", "IP:端口/数据库")
        self.cursor = self.conn.cursor()
        # 建立一个光标，用于执行sql语句

    def connect_close(self):
        self.cursor.close()
        self.conn.close()
        # 关闭光标和断开数据库

    def find_sql(self, sql):
        try:
            self.connect_open()

            self.cursor.execute(sql)
            # execute输入sql语句
            result = self.cursor.fetchall()
            # fetahone提取数据库表格第一行， fetchall提取数据库表格所有数据
            self.connect_close()

            return result
        except Exception as e:
            print(e.message)
