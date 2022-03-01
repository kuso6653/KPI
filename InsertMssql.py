# -*- coding:utf-8 -*-
from PyOdbc import Pyodbc

if __name__ == '__main__':
    sql = "insert into demo(cinvcode) values (?)"
    ms = Pyodbc("10.56.164.228", "KPI", "sa", "Chem123#")  # 224的数据库链接
    str1 = "123456"
    ms.ExecNonQuery(sql, str1)
