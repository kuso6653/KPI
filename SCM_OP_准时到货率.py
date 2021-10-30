import pandas as pd
from numpy import datetime64

import Func


class ArriveTime:
    def __init__(self):
        self.func = Func
        self.ThisMonthStart, self.ThisMonthEnd, self.LastMonthEnd, self.LastMonthStart = self.func.GetDate()
        self.path = "//10.56.164.127/it&m/KPI"
        # 将上月首尾日期切割
        self.ThisMonthStart = str(self.ThisMonthStart).split(" ")[0].replace("-", "")
        self.ThisMonthEnd = str(self.ThisMonthEnd).split(" ")[0].replace("-", "")
        self.LastMonthStart = str(self.LastMonthStart).split(" ")[0].replace("-", "")
        self.LastMonthEnd = str(self.LastMonthEnd).split(" ")[0].replace("-", "")
        self.GoodsInData = pd.read_excel(f"{self.path}/DATA/SCM/OP/到货单列表-{self.LastMonthStart}-{self.ThisMonthEnd}.XLSX",
                                         usecols=['存货编码', '存货名称', '采购委外订单号', '行号', '制单时间'],
                                         converters={'行号': int, '存货编码': str, '制单时间': datetime64})
        self.PurchaseInData = pd.read_excel(f"{self.path}/DATA/SCM/OP/采购订单列表-{self.LastMonthStart}-{self.ThisMonthEnd}.XLSX",
                                            usecols=['存货编码', '存货名称', '订单编号', '行号', '计划到货日期'],
                                            converters={'存货编码': str, '行号': int, '计划到货日期': datetime64})

    def mkdir(self, path):
        self.func.mkdir(path)

    def GetArriveTime(self):
        self.GoodsInData = self.GoodsInData.dropna(subset=['存货编码'])  # 去除nan的列
        self.PurchaseInData = self.PurchaseInData.dropna(subset=['存货编码'])  # 去除nan的列
        self.PurchaseInData = self.PurchaseInData.rename(columns={'订单编号': '采购委外订单号'})

        ArriveTimeData = pd.merge(self.GoodsInData, self.PurchaseInData, on=['存货编码', '存货名称', '采购委外订单号', '行号'])
        # all_data["out_data"] =all_data["制单时间"]-all_data["计划到货日期"]
        ArriveTimeData['计划到货日期'] = pd.to_datetime(ArriveTimeData['计划到货日期'].astype(str)) + pd.to_timedelta('20:00:00')
        ArriveTimeData["out_data/H"] = (
                (ArriveTimeData["制单时间"] - ArriveTimeData["计划到货日期"]) / pd.Timedelta(1, 'H')).astype(int)
        ArriveTimeData = ArriveTimeData.loc[ArriveTimeData["out_data/H"] > 72]
        self.SaveFile(ArriveTimeData)

    def SaveFile(self, ArriveTimeData):
        self.mkdir(self.path + "/RESULT/SCM/OP")
        ArriveTimeData.to_excel(f'{self.path}/RESULT/SCM/OP/准时到货率.xlsx', sheet_name="准时到货率", index=False)

    def run(self):
        self.GetArriveTime()


if __name__ == '__main__':
    AT = ArriveTime()
    AT.run()
