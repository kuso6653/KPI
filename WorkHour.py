import pandas as pd
from numpy import datetime64
from openpyxl import load_workbook
from xlrd import book

import Func

EarliestTime = datetime64("2000-01-02")  # 设置工艺路线版本日期的最早期限


class WorkHour:
    def __init__(self):
        self.AsNameList = []
        self.func = Func
        self.ThisMonthStart, self.ThisMonthEnd, self.LastMonthEnd, self.LastMonthStart = self.func.GetDate()
        self.path = "//10.56.164.127/it&m/KPI"
        self.work_list = ['打磨', '大、小件划线', '焊接', '冷作', '修毛刺', '人工工时', '水压试验', '攻螺纹、修毛刺、清洁', '清洁']
        self.Routing_data = pd.read_excel(f"{self.path}/DATA/WORKHOUR/工艺路线资料表--含资源.xlsx",
                                          header=3, usecols=["物料编码", "工作中心", "版本日期", "资源名称", "工时(分子)"],
                                          converters={'物料编码': str, "工时(分子)": int})

        self.WorkHourData = pd.read_excel(f"{self.path}/DATA/WORKHOUR/报工列表-20210701-20210731.xlsx",
                                          usecols=["单据日期", "制单人", "生产订单", "行号",
                                                   "物料编码", "物料名称", "生产数量", "资源工时1", "资源名称1",
                                                   "资源工时2", "资源名称2"],
                                          converters={'行号': str, "资源工时1": int, "资源工时2": int, "生产数量": int, '物料编码': str})

    def GetWorkHour(self):
        # 重新定义 版本日期格式，再转化为 datatime64
        self.Routing_data = self.Routing_data.dropna(subset=['物料编码'])  # 去除nan的列
        self.Routing_data["版本日期"] = self.Routing_data["版本日期"].str.replace("/", "-").astype("datetime64")
        self.Routing_data = self.Routing_data[self.Routing_data["版本日期"] > EarliestTime]
        # self.Routing_data = self.Routing_data.drop_duplicates(subset=["物料编码"])  # 去重

        # 筛选 资源名称1，资源名称2 的符合work_list 的 资源名称
        WorkHourData_First = self.WorkHourData[self.WorkHourData['资源名称1'].isin(self.work_list)]
        WorkHourData_Second = self.WorkHourData[self.WorkHourData['资源名称2'].isin(self.work_list)]
        self.WorkHourData = pd.merge(WorkHourData_First, WorkHourData_Second, how="outer")
        self.WorkHourData["资源总工时"] = self.WorkHourData['资源工时1'] + self.WorkHourData['资源工时2']
        self.WorkHourData.to_excel(f'{self.path}/RESULT/WORKHOUR/Data.xlsx', sheet_name="0", index=False)
        WorkHourData_First.to_excel(f'{self.path}/RESULT/WORKHOUR/First.xlsx', sheet_name="0", index=False)
        WorkHourData_Second.to_excel(f'{self.path}/RESULT/WORKHOUR/Second.xlsx', sheet_name="0", index=False)

        for name, group in self.WorkHourData.groupby(["制单人"]):
            group = pd.DataFrame(group)  # 新建pandas
            group = pd.merge(group, self.Routing_data, on=["物料编码"])
            group["总工时"] = group["生产数量"] * group["工时(分子)"]
            self.AsNameList.append(group)

    def mkdir(self, path):
        self.func.mkdir(path)

    def run(self):
        self.GetWorkHour()
        self.func.mkdir(self.path + '/RESULT/WORKHOUR')
        # self.AsNameList[0].to_excel(f'{self.path}/RESULT/WORKHOUR/demo.xlsx', sheet_name="0", index=False)
        flag = 1
        # for data in self.AsNameList[1:]:
        #     writer = pd.ExcelWriter(f"{self.path}/RESULT/WORKHOUR/demo.xlsx", engine='openpyxl')
        #     writer.book = book
        #     data.to_excel(writer, f"{flag}", index=False)
        #     writer.save()
        #     flag = flag + 1


if __name__ == '__main__':
    WH = WorkHour()
    WH.run()
