# -*- coding:utf-8 -*-
from PyOdbc import Pyodbc

if __name__ == '__main__':
    sql = "insert into demo2(id, 姓名, 性别) values (?, ?, ?)"
    ms = Pyodbc()
    str1 = (4, "张胜男", "女")
    ms.ExecNonQuery(sql, str1)
