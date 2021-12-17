
import pandas as pd
from sqlalchemy import create_engine
import pymssql

from PyOdbc import Pyodbc

data1 = [
    [5, "法外狂徒", "男"],
    [6, "姬霓太美", "女"],
    [7, "可视化", "无"]
]
# engine = create_engine('mysql+pymysql://%s:%s@%s:%d/%s' % (sql_user, sql_passwd, sql_host, sql_port, sql_db))
# engine = create_engine('mysql+pymssql://sa:Chem123#@10.56.164.228:3306/demo2?charset=utf8')



USERNAME = 'sa'
PASSWORD = 'Chem123#'
HOST = '10.56.164.228'
PORT = '1433'
DATABASE = 'demo2'

DB_URL = 'mysql+pymysql://{}:{}@{}:{}/{}?charset=utf8'.format(USERNAME, PASSWORD, HOST, PORT, DATABASE)

SQLALCHEMY_DATABASE_URI = DB_URL

# 动态追踪修改设置，如未设置只会提示警告
SQLALCHEMY_TRACK_MODIFICATIONS = False

# 查询时会显示原始sql语句
SQLALCHEMY_ECHO = True

engine = create_engine(DB_URL)
df = pd.DataFrame(data1, columns=["id", "姓名", "性别"])
df.to_sql("demo2", engine, if_exists='append')

