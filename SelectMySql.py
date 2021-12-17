# -*- coding:utf-8 -*-
from PyOdbc import Pyodbc

if __name__ == '__main__':
    year = 2021  # 获取11月工作日和假期 ！！！ 查询日期一定要修改 ！！！
    month = 10
    sql = "select * from demo2"
    ms = Pyodbc()
    List = ms.ExecQuery(sql)  # 从数据库获取的函数
    print(List)
