import xlrd
import pandas as pd
import calendar
import datetime
from datetime import timedelta
import openpyxl
import Func
from numpy import datetime64
from openpyxl import load_workbook


class OrderMaintenance:
    def __init__(self):
        self.func = Func
        self.ThisMonthStart, self.ThisMonthEnd, self.LastMonthEnd, self.LastMonthStart = self.func.GetDate()
        self.path = Func.Path()

    def mkdir(self, path):
        self.func.mkdir(path)

    def GetOrderMaintenance(self):
        # 将上月首尾日期切割
        self.ThisMonthStart = str(self.ThisMonthStart).split(" ")[0].replace("-", "")
        self.ThisMonthEnd = str(self.ThisMonthEnd).split(" ")[0].replace("-", "")
        # 材料出库及时率
        ECNData = pd.read_excel(f"{self.path}/DATA/SCM/OM/ECN单列表.XLSX",
                                usecols=['单据编号', '生产订单号', '审核日期'],
                                converters={'审核日期': datetime64, '单据编号': str})
        # '母件编码', '母件名称', '旧子件编码',
        # '旧子件名称', '旧子件基本用量', '新子件编码', '新子件名称', '新子件基本用量',
        ECNDealData = pd.read_excel(f"{self.path}/DATA/SCM/OM/ECN处理单列表.XLSX",
                                    usecols=['Ecn单', 'Ecn处理单', 'Ecn处理单单据日期', '关联单据', '关联单据物料编码',

                                             '关联单据物料名称', '关联单据数量'],
                                    converters={'关联单据数量': str, "Ecn单": str, 'Ecn处理单单据日期': datetime64, 'Ecn处理单': str})

        ECNData = ECNData.dropna(subset=['生产订单号'])  # 去除nan的列
        ECNData = ECNData.drop_duplicates()  # 去重
        ECNDealData = ECNDealData.dropna(subset=['Ecn单'])  # 去除nan的列
        ECNData = ECNData.rename(columns={'单据编号': 'Ecn单'})
        ECNDealData = ECNDealData.drop_duplicates()  # 去重
        del ECNData["生产订单号"]
        OrderMaintenanceData = pd.merge(ECNDealData, ECNData, on='Ecn单')
        OrderMaintenanceData = OrderMaintenanceData.drop_duplicates()  # 去重
        OrderMaintenanceData['审批延时/H'] = (
                (OrderMaintenanceData['Ecn处理单单据日期'] - OrderMaintenanceData['审核日期']) / pd.Timedelta(1, 'H')).astype(
            int)

        OrderMaintenanceData.loc[OrderMaintenanceData["审批延时/H"] > 24, "单据状态"] = "超时"  # 计算出来的审批延时大于1天为超时
        OrderMaintenanceData.loc[OrderMaintenanceData["审批延时/H"] <= 24, "单据状态"] = "正常"  # 小于等于1天为正常

        try:
            OrderMaintenanceCount = OrderMaintenanceData['单据状态'].value_counts()['超时']
        except:
            OrderMaintenanceCount = 0

        OrderMaintenanceCountAll = OrderMaintenanceData.shape[0]
        OrderMaintenanceResult = format(float(1 - OrderMaintenanceCount / OrderMaintenanceCountAll), '.2%')
        dict = {'当月未及时处理工程更改物料数': [OrderMaintenanceCount], '当月工程更改物料总数': [OrderMaintenanceCountAll],
                '生产订单维护及时率': [OrderMaintenanceResult]}
        OrderMaintenanceResult_sheet = pd.DataFrame(dict)

        self.SaveFile(OrderMaintenanceData, OrderMaintenanceResult_sheet)

    def SaveFile(self, OrderMaintenanceData, OrderMaintenanceResult_sheet):
        self.mkdir(self.path+"/RESULT/SCM/OM")
        OrderMaintenanceResult_sheet.to_excel(f'{self.path}/RESULT/SCM/OM/生产订单维护及时率.xlsx', sheet_name="生产订单维护及时率", index=False)
        book = load_workbook(f'{self.path}/RESULT/SCM/OM/生产订单维护及时率.xlsx')
        writer = pd.ExcelWriter(f"{self.path}/RESULT/SCM/OM/生产订单维护及时率.xlsx", engine='openpyxl')
        writer.book = book
        OrderMaintenanceData.to_excel(writer, "当月工程更改物料处理情况", index=False)
        writer.save()

    def run(self):
        self.GetOrderMaintenance()


if __name__ == '__main__':
    OM = OrderMaintenance()
    OM.run()
