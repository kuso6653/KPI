import xlrd
import pandas as pd
import calendar
import datetime
from datetime import timedelta
import openpyxl
from numpy import datetime64

import Func


class Deliver:
    def __init__(self):
        self.func = Func
        self.ThisMonthStart, self.ThisMonthEnd, self.LastMonthEnd, self.LastMonthStart = self.func.GetDate()
        self.path = Func.Path()

        # 将这个月和上个月首尾日期切割
        self.ThisMonthStart = str(self.ThisMonthStart).split(" ")[0].replace("-", "")
        self.ThisMonthEnd = str(self.ThisMonthEnd).split(" ")[0].replace("-", "")
        self.LastMonthStart = str(self.LastMonthStart).split(" ")[0].replace("-", "")
        self.LastMonthEnd = str(self.LastMonthEnd).split(" ")[0].replace("-", "")

    def mkdir(self, path):
        self.func.mkdir(path)

    def GetDeliver(self):
        SaleOutData = pd.read_excel(f"{self.path}/DATA/SCM/LOGISTIC/销售出库单列表.XLSX",
                                    usecols=['发货单号', '审核时间', '存货编码'],
                                    converters={'发货单号': str, '存货编码': str, '审核时间': datetime64})
        InvoiceData = pd.read_excel(f"{self.path}/DATA/SCM/LOGISTIC/发货单列表.XLSX",
                                    usecols=['发货单号', '审核时间', '存货编码'],
                                    converters={'发货单号': str, '存货编码': str,'审核时间': datetime64})

        SaleOutData = SaleOutData.rename(columns={'审核时间': '销售出库单审核时间'})
        InvoiceData = InvoiceData.rename(columns={'审核时间': '发货单审核时间'})

        DeliverData = pd.merge(SaleOutData, InvoiceData, on=["发货单号", '存货编码'])
        DeliverData = DeliverData.dropna(axis=0, how='any')  # 去除所有nan的列

        DeliverData['审批延时'] = ((DeliverData['销售出库单审核时间'] - DeliverData['发货单审核时间']) / pd.Timedelta(1, 'H')).astype(
            int)
        DeliverData.loc[DeliverData["审批延时"] > 24, "单据状态"] = "超时"
        DeliverData.loc[DeliverData["审批延时"] <= 24, "单据状态"] = "正常"
        self.SaveFile(DeliverData)

    def SaveFile(self, DeliverData):
        self.mkdir(self.path+"/RESULT/SCM/LOGISTIC")
        DeliverData.to_excel(f'{self.path}/RESULT/SCM/LOGISTIC/发货及时率.xlsx', sheet_name="发货及时率", index=False)

    def run(self):
        self.GetDeliver()


if __name__ == '__main__':
    D = Deliver()
    D.run()
