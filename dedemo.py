import pandas as pd
from numpy import datetime64
from openpyxl import load_workbook

import Func

EarliestTime = datetime64("2000-01-02")  # 设置工艺路线版本日期的最早期限


class WorkHour:
    def __init__(self):
        self.func = Func
        self.ThisMonthStart, self.ThisMonthEnd, self.LastMonthEnd, self.LastMonthStart = self.func.GetDate()
        self.path = "//10.56.164.127/it&m/KPI"
        self.Routing_data = pd.read_excel(f"{self.path}/DATA/WORKHOUR/工艺路线资料表--含资源.xlsx",
                                          header=3, usecols=["物料编码", "工作中心", "版本日期", "资源名称", "工时(分子)"],
                                          converters={'物料编码': str, "工时(分子)": int})

        self.WorkHourData = pd.read_excel(f"{self.path}/DATA/WORKHOUR/报工列表.xlsx",
                                          usecols=["单据日期", "制单人", "生产订单", "行号",
                                                   "物料编码", "物料名称", "生产数量", "资源工时1", "资源名称1",
                                                   "资源工时2", "资源名称2"],
                                          converters={'行号': str, "资源工时1": int, "资源工时2": int})

    def GetWorkHour(self):
        # 重新定义 版本日期格式，再转化为datatime64
        self.Routing_data = self.Routing_data.dropna(subset=['物料编码'])  # 去除nan的列
        self.Routing_data["版本日期"] = self.Routing_data["版本日期"].str.replace("/", "-").astype("datetime64")
        self.Routing_data = self.Routing_data[self.Routing_data["版本日期"] > EarliestTime]
        self.Routing_data = self.Routing_data.drop_duplicates(subset=["物料编码"])  # 去重

    def run(self):
        self.GetWorkHour()


if __name__ == '__main__':
    WH = WorkHour()
    WH.run()
