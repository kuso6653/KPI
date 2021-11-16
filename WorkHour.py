from decimal import Decimal
import pandas as pd
from numpy import datetime64
from openpyxl import load_workbook
import Func


# EarliestTime = datetime64("2000-01-02")  # 设置工艺路线版本日期的最早期限


# 对比 当月的 实际报工工时，大于0 返回 资源工时1 ，否则 返回 资源工时2
def FirsSecondChoice(df):
    if df['实际报工工时'] > 0:
        return df['资源工时1']
    else:
        return df['资源工时2']


# 计划工时分组计算合并值
def GetSumPlanData(Data):
    sum_data_list = []
    for name, group in Data.groupby(["物料编码"]):
        qualified_num = group['工时(分子)'].sum()  # 取合格数量总值保留两位
        qualified_num = Decimal(qualified_num).quantize(Decimal('0.00'))
        group.loc[:, "计划工时"] = qualified_num  # 新建 总合格数量 列
        sum_data_list.append(group.head(1))
    return sum_data_list


# 获取最大版本代号
def GetMaxVersionNumber(Data):
    max_data_list = []
    for name, group in Data.groupby(["物料编码"]):
        group = pd.DataFrame(group)  # 新建pandas
        group = group.sort_values(by='版本代号', ascending=False)  # 降序排序
        MaxNumber = group.iloc[0, 1]  # 取降序排序第一个 版本代号
        group = group[group["版本代号"] == MaxNumber]  # 筛选 版本代号 最大的数据
        max_data_list.append(group)
    return pd.concat(max_data_list)


# 实际工时分组计算合并值
def GetSumActualData(Data):
    sum_data_list = []
    for name, group in Data.groupby(["物料编码", "生产订单", "行号"]):
        qualified_num = group['实际报工工时'].sum()  # 取合格数量总值保留两位
        group.loc[:, "实际报工工时"] = qualified_num  # 新建 总合格数量 列
        sum_data_list.append(group.head(1))
    return sum_data_list


class WorkHour:
    def __init__(self):
        self.userName = []
        self.AsNameList = []
        self.func = Func
        self.ThisMonthStart, self.ThisMonthEnd, self.LastMonthEnd, self.LastMonthStart = self.func.GetDate()
        self.path = "//10.56.164.127/it&m/KPI"
        self.work_list = ['打磨', '大、小件划线', '焊接', '冷作', '修毛刺', '人工工时', '水压试验', '攻螺纹、修毛刺、清洁', '清洁']

        self.PlanHourData = pd.DataFrame
        self.Routing_data = pd.read_excel(f"{self.path}/DATA/WORKHOUR/工艺路线资料表--含资源.xlsx",
                                          header=3, usecols=["物料编码", "工作中心", "版本日期", "版本代号", "资源名称", "工时(分子)"],
                                          converters={'物料编码': str, "工时(分子)": int, "版本代号": int})

        self.WorkHourData = pd.read_excel(f"{self.path}/DATA/WORKHOUR/报工列表-20210901-20210930.xlsx",
                                          usecols=["单据日期", "单据号码", "制单人", "生产订单", "行号",
                                                   "物料编码", "物料名称", "生产数量", "资源工时1", "资源名称1",
                                                   "资源工时2", "资源名称2"],
                                          converters={'行号': str, "资源工时1": int, "资源工时2": int, "生产数量": int, '物料编码': str})

    def GetWorkData(self, group):

        group["标准工时"] = group["生产数量"] * group["计划工时"]  # 计算 标准工时
        self.AsNameList.append(group)

    def GetWorkHour(self):
        # 重新定义 版本日期格式，再转化为 datatime64
        self.Routing_data = self.Routing_data.dropna(subset=['物料编码'])  # 去除nan的列
        self.Routing_data["版本日期"] = self.Routing_data["版本日期"].str.replace("/", "-").astype("datetime64")
        # self.Routing_data = self.Routing_data[self.Routing_data["版本日期"] > EarliestTime]

        # 筛选 资源名称1，资源名称2 的符合work_list 的 资源名称
        WorkHourData_First = self.WorkHourData[self.WorkHourData['资源名称1'].isin(self.work_list)]
        WorkHourData_Second = self.WorkHourData[self.WorkHourData['资源名称2'].isin(self.work_list)]
        self.Routing_data = self.Routing_data[self.Routing_data['资源名称'].isin(self.work_list)]
        self.WorkHourData = pd.merge(WorkHourData_First, WorkHourData_Second, how="outer")
        self.WorkHourData["实际报工工时"] = self.WorkHourData['资源工时1'] - self.WorkHourData['资源工时2']
        # 放入函数 FirsSecondChoice 对比取值
        self.WorkHourData["实际报工工时"] = self.WorkHourData.apply(FirsSecondChoice, axis=1)
        # 使用正则表达式尽心模糊匹配
        PXData = self.Routing_data.loc[self.Routing_data['工作中心'].str.contains('^P')]  # 模糊查询总装PX 车间
        MHData = self.Routing_data.loc[self.Routing_data['工作中心'].str.contains('^M')]  # 模糊查询铆焊 车间
        JJData = self.Routing_data.loc[self.Routing_data['工作中心'].str.contains('^J')]  # 模糊查询机加工 车间
        ZZData = self.Routing_data.loc[self.Routing_data['工作中心'].str.contains('^Z')]  # 模糊查询总装 车间
        DKData = self.Routing_data.loc[self.Routing_data['工作中心'].str.contains('^DK')]  # 模糊查询总装电控 车间

        PXData = GetMaxVersionNumber(PXData)
        MHData = GetMaxVersionNumber(MHData)
        JJData = GetMaxVersionNumber(JJData)
        ZZData = GetMaxVersionNumber(ZZData)
        DKData = GetMaxVersionNumber(DKData)

        for name, group in self.WorkHourData.groupby(["制单人"]):  # 按车间操作人员进行拆分
            group = pd.DataFrame(group)  # 新建pandas
            # group.to_excel(f'{self.path}/RESULT/WORKHOUR/group.xlsx', sheet_name="0", index=False)
            group = pd.concat(GetSumActualData(group), axis=0, ignore_index=True)
            if name == "郭东升":
                # 对各个车间的 以物料编码 为 主键 将 工时(分子) 进行sum合计，返回的值进行合并
                Data = pd.concat(GetSumPlanData(PXData), axis=0, ignore_index=True)
                group = pd.merge(group, Data, on=['物料编码'])
                self.GetWorkData(group)
                self.userName.append([name, "总装PX"])
            elif name == "黄鑫凯":
                Data = pd.concat(GetSumPlanData(DKData), axis=0, ignore_index=True)
                group = pd.merge(group, Data, on=['物料编码'])
                self.GetWorkData(group)
                self.userName.append([name, "总装电控"])
            elif name == "乐美珠":
                Data = pd.concat(GetSumPlanData(JJData), axis=0, ignore_index=True)
                group = pd.merge(group, Data, on=['物料编码'])
                self.GetWorkData(group)
                self.userName.append([name, "机加工"])
            elif name == "吕春华" or name == "夏正棋":
                Data = pd.concat(GetSumPlanData(ZZData), axis=0, ignore_index=True)
                group = pd.merge(group, Data, on=['物料编码'])
                self.GetWorkData(group)
                self.userName.append([name, "总装"])
            elif name == "杨薇1" or name == "林李旭":
                Data = pd.concat(GetSumPlanData(MHData), axis=0, ignore_index=True)
                group = pd.merge(group, Data, on=['物料编码'])
                self.GetWorkData(group)
                self.userName.append([name, "铆焊"])

    def mkdir(self, path):
        self.func.mkdir(path)

    def SaveFile(self):
        self.AsNameList[0].to_excel(f'{self.path}/RESULT/WORKHOUR/Data.xlsx', sheet_name=f"{self.userName[0][0]}",
                                    index=False)
        for index, values in enumerate(zip(self.AsNameList[1:], self.userName[1:])):
            SaveBook = load_workbook(f'{self.path}/RESULT/WORKHOUR/Data.xlsx')
            writer = pd.ExcelWriter(f"{self.path}/RESULT/WORKHOUR/Data.xlsx", engine='openpyxl')
            writer.book = SaveBook
            values[0].to_excel(writer, f"{values[1][0]}", index=False)
            writer.save()

    def run(self):
        self.GetWorkHour()
        self.func.mkdir(self.path + '/RESULT/WORKHOUR')
        self.SaveFile()


if __name__ == '__main__':
    WH = WorkHour()
    WH.run()
