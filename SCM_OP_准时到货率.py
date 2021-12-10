import pandas as pd
from numpy import datetime64
from openpyxl import load_workbook

import Func


class ArriveTime:
    def __init__(self):
        self.func = Func
        self.ThisMonthStart, self.ThisMonthEnd, self.LastMonthEnd, self.LastMonthStart = self.func.GetDate()
        self.path = "//10.56.164.228/KPI"
        # 将上月首尾日期切割
        self.ThisMonthStart = str(self.ThisMonthStart).split(" ")[0].replace("-", "")
        self.ThisMonthEnd = str(self.ThisMonthEnd).split(" ")[0].replace("-", "")
        self.LastMonthStart = str(self.LastMonthStart).split(" ")[0].replace("-", "")
        self.LastMonthEnd = str(self.LastMonthEnd).split(" ")[0].replace("-", "")
        self.PurchaseInData = pd.read_excel(
            f"{self.path}/DATA/SCM/OP/采购订单列表-{self.LastMonthStart}-{self.ThisMonthEnd}.XLSX",
            usecols=['订单编号', '行号', '实际到货日期'],
            converters={'行号': int, '实际到货日期': datetime64})
        self.Prescription = pd.read_excel(
            f"{self.path}/DATA/SCM/采购时效性统计表-{self.LastMonthStart}-{self.ThisMonthEnd}.XLSX",
            usecols=[0, 1, 6, 7, 9, 11, 12, 14, 15, 16], header=2,
            names=["行号", "采购订单号", "存货编码", "存货名称", "计划到货日期", "采购订单制单时间", "采购订单审核时间",
                   "到货单号", "到货单行号", "到货单制单时间"],
            converters={'计划到货日期': datetime64, '采购订单制单时间': datetime64, '采购订单审核时间': datetime64, '到货单制单时间': datetime64})

    def mkdir(self, path):
        self.func.mkdir(path)

    def GetThisMonthArriveTime(self):  # 当月准时到货率 和 当月未到货清单
        self.Prescription = self.Prescription.dropna(subset=['行号'])  # 去除nan的列
        self.PurchaseInData = self.PurchaseInData.dropna(subset=['行号'])  # 去除nan的列
        self.PurchaseInData = self.PurchaseInData.rename(columns={'订单编号': '采购订单号', '行号': '采购订单行号'})
        self.Prescription = self.Prescription.rename(columns={'行号': '采购订单行号'})
        ThisMonthArriveData = self.Prescription[self.ThisMonthEnd >= self.Prescription['计划到货日期']]
        ThisMonthArriveData = ThisMonthArriveData[ThisMonthArriveData['计划到货日期'] >= self.ThisMonthStart]
        ThisMonthArriveData = pd.merge(ThisMonthArriveData, self.PurchaseInData, on=['采购订单号', '采购订单行号'])

        # 筛选 实际到货日期 为空的， 用 计划到货日期 补全
        ThisMonthArriveData['实际到货日期'] = ThisMonthArriveData['实际到货日期'].fillna(ThisMonthArriveData['计划到货日期'])
        ThisMonthArriveData['实际到货日期'] = pd.to_datetime(ThisMonthArriveData['实际到货日期'].astype(str)) + pd.to_timedelta(
            '20:00:00')
        ThisMonthNoArriveData = ThisMonthArriveData[ThisMonthArriveData.isnull().any(axis=1)]  #
        ThisMonthArriveData = ThisMonthArriveData[ThisMonthArriveData['到货单制单时间'].notnull()]  #
        ThisMonthArriveData["审批延时/H"] = (
                (ThisMonthArriveData["到货单制单时间"] - ThisMonthArriveData["实际到货日期"]) / pd.Timedelta(1, 'H')).astype(int)
        ThisMonthArriveData.loc[ThisMonthArriveData["审批延时/H"] > 72, "单据状态"] = "逾期"
        ThisMonthArriveData.loc[ThisMonthArriveData["审批延时/H"] <= 72, "单据状态"] = "正常"
        ThisMonthArriveData.loc[ThisMonthArriveData["审批延时/H"] < 0, "单据状态"] = "提前"

        ThisMonthArriveData_Order = ['采购订单号', '采购订单行号', '存货编码', '存货名称', '计划到货日期', '实际到货日期', '采购订单制单时间', '采购订单审核时间',
                                     '到货单号', '到货单行号', '到货单制单时间', '审批延时/H', '单据状态']
        ThisMonthNoArriveData_Order = ['采购订单号', '采购订单行号', '存货编码', '存货名称', '计划到货日期', '实际到货日期', '采购订单制单时间', '采购订单审核时间',
                                       '到货单号', '到货单行号', '到货单制单时间']
        ThisMonthArriveData = ThisMonthArriveData[ThisMonthArriveData_Order]
        ThisMonthNoArriveData = ThisMonthNoArriveData[ThisMonthNoArriveData_Order]

        self.mkdir(self.path + "/RESULT/SCM/OP")
        ThisMonthArriveData.to_excel(f'{self.path}/RESULT/SCM/OP/准时到货率.xlsx', sheet_name="当月准时到货率", index=False)
        book = load_workbook(f'{self.path}/RESULT/SCM/OP/准时到货率.xlsx')
        writer = pd.ExcelWriter(f"{self.path}/RESULT/SCM/OP/准时到货率.xlsx", engine='openpyxl')
        writer.book = book
        ThisMonthNoArriveData.to_excel(writer, "当月未到货清单", index=False)
        writer.save()

    def GetHistoryMonthArriveTime(self):  # 历史未到货清单
        HistoryMonthArriveData = self.Prescription[self.Prescription['计划到货日期'] < self.ThisMonthStart]
        HistoryMonthArriveData = pd.merge(self.PurchaseInData, HistoryMonthArriveData, on=['采购订单号', '采购订单行号'])
        HistoryMonthArriveData["实际到货日期"][HistoryMonthArriveData["实际到货日期"].isnull()] = HistoryMonthArriveData['计划到货日期']
        # 当 采购订单审核时间 或 到货单制单时间 为空值的时候取其数值
        HistoryMonthArriveData = HistoryMonthArriveData[
            (HistoryMonthArriveData["采购订单审核时间"].isnull()) | (HistoryMonthArriveData["到货单制单时间"].isnull())]
        order = ['采购订单号', '采购订单行号', '存货编码', '存货名称', '计划到货日期', '实际到货日期', '采购订单制单时间', '采购订单审核时间', '到货单号', '到货单行号',
                 '到货单制单时间']
        HistoryMonthArriveData = HistoryMonthArriveData[order]

        book = load_workbook(f'{self.path}/RESULT/SCM/OP/准时到货率.xlsx')
        writer = pd.ExcelWriter(f"{self.path}/RESULT/SCM/OP/准时到货率.xlsx", engine='openpyxl')
        writer.book = book
        HistoryMonthArriveData.to_excel(writer, "历史未到货清单", index=False)
        writer.save()

    def run(self):
        self.GetThisMonthArriveTime()
        self.GetHistoryMonthArriveTime()


if __name__ == '__main__':
    AT = ArriveTime()
    AT.run()
