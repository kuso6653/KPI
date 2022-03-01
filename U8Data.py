# -*- coding:utf-8 -*-
from PyOdbc import Pyodbc
import Func
import datetime
import pandas as pd


####################################
####
#### 用于MRP、存货档案、工序派工委外订单
#### 从U8数据库转存到228数据库
####
####################################
class U8Data:
    def __init__(self):
        self.func = Func
        self.U8ms = Pyodbc("10.56.164.234", "UFDATA_006_2019", "sa", "@4miNisTr@t0r")  # U8的数据库链接
        self.Oms = Pyodbc("10.56.164.228", "KPI", "sa", "Chem123#")  # 224的数据库链接

    def InventoryFunc(self):

        InventoryData = self.U8ms.ExecQuery(self.func.ReturnInventorySql())  # 查询获取U8数据库 存货档案 数据

        InsertSql = """insert into Inventory(cinvcode, cinvdefine7, cvenname, cPersonName, cdepname,
        iplandefault,cplanmethod,iinvadvance, ialteradvance, falterbasenum,fminsupply, dsdate, dedate,
        dmodifydate ,cidefine4) values (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?) """  # 插入语句
        # ClearSql = "truncate table Inventory"
        # self.Oms.ExecNonQuery2(ClearSql)
        for data in InventoryData:
            str1 = (data[0], data[1], data[2], data[3], data[4], data[5], data[6], data[7], data[8], data[9], data[10],
                    data[11], data[12], data[13], data[14])
            self.Oms.ExecNonQuery(InsertSql, str1)  # 读取的数据一行行插入数据库

    def MoRoutingDetailFunc(self):
        TodayData = datetime.date.today()
        MoRoutingDetailData = self.U8ms.ExecQuery(self.func.ReturnMoRoutingDetailSql())  # 查询获取U8数据库 委外派工 数据
        InsertSql = """insert into MoRoutingDetail(MoCode, SortSeq, InvCode, OpSeq, DPgFlag, DataTag
                ) values (?, ?, ?, ?, ?, ?) """  # 插入语句
        # ClearSql = "truncate table MoRoutingDetail"
        # self.Oms.ExecNonQuery2(ClearSql)
        for data in MoRoutingDetailData:
            str1 = (data[0], data[1], data[2], data[3], data[4], TodayData)
            self.Oms.ExecNonQuery(InsertSql, str1)  # 读取的数据一行行插入数据库

    def MRPFunc(self):
        CountSql1, CountSql2, MRPDataSql = self.func.ReturnMRPSql()
        MRPDataData = self.U8ms.ExecQuery(MRPDataSql)  # 查询获取U8数据库数据
        print(MRPDataData)

    def run(self):
        # self.InventoryFunc()
        # self.MoRoutingDetailFunc()
        self.MRPFunc()
        self.U8ms.CloseConnect()
        self.Oms.CloseConnect()


if __name__ == '__main__':
    u8 = U8Data()
    u8.run()
