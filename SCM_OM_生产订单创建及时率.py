import pandas as pd
from numpy import datetime64
from openpyxl import load_workbook

import Func

EarliestTime = datetime64("2000-01-02")  # 设置工艺路线版本日期的最早期限


class OrderCreation:
    def __init__(self):
        self.func = Func
        self.ThisMonthStart, self.ThisMonthEnd, self.LastMonthEnd, self.LastMonthStart = self.func.GetDate()
        self.path = Func.Path()
        # 将上月首尾日期切割
        self.OtherMonthEnd = str(self.ThisMonthEnd).split(" ")[0].replace("", "")  # 取当月最后一天
        self.ThisMonthStart = str(self.ThisMonthStart).split(" ")[0].replace("-", "")
        self.ThisMonthEnd = str(self.ThisMonthEnd).split(" ")[0].replace("-", "")

        self.ProductionData = pd.read_excel(
            f"{self.path}/DATA/PROD/生产订单列表-{self.ThisMonthStart}-{self.ThisMonthEnd}.XLSX",
            usecols=['生产订单号', '行号', '物料编码', '物料名称', '生产批号', '制单时间', '类型'],
            converters={'生产订单号': str, '制单时间': datetime64, '物料编码': str})

        self.Material_data = pd.read_excel(f"{self.path}/DATA/SCM/存货档案{self.OtherMonthEnd}.XLSX",
                                           usecols=['存货编码', '计划默认属性', '启用日期'],
                                           converters={'启用日期': datetime64, "存货编码": str})
        self.Routing_data = pd.read_excel(f"{self.path}/DATA/SCM/OM/工艺路线资料表--含资源.xlsx",
                                          usecols=[0, 4, 6], header=3, names=["物料编码", "版本代号", "版本日期"],
                                          converters={'版本日期': str, '物料编码': str})

        self.BOM_data = pd.read_excel(f"{self.path}/DATA/SCM/OM/BOM集成时间表.xlsx", header=1,
                                      usecols=['子件编码', '计划默认属性', '集成时间'],
                                      converters={"子件编码": str})

    def mkdir(self, path):
        self.func.mkdir(path)

    def ModifyFormat(self):  # 修改数据基本格式
        self.ProductionData = self.ProductionData[self.ProductionData["类型"] == "标准"]
        self.Material_data = self.Material_data.rename(columns={'存货编码': '物料编码'})
        self.Routing_data = self.Routing_data.dropna(subset=['版本代号'])  # 去除nan的列
        self.Routing_data["版本日期"] = self.Routing_data["版本日期"].str.replace("/", "-").astype("datetime64")
        self.Routing_data = self.Routing_data[self.Routing_data["版本日期"] > EarliestTime]
        self.Routing_data = self.Routing_data.drop_duplicates(subset=["物料编码"])  # 去重
        self.ProductionData = self.ProductionData[self.ProductionData["制单时间"] > self.ThisMonthStart]
        self.Material_data = self.Material_data[self.Material_data["启用日期"] > self.ThisMonthStart]
        # 合并生成当月数据在导入到U8中进行再查询
        self.ThisMonthData = pd.merge(self.ProductionData, self.Material_data, how="left", on=['物料编码'])
        # self.ThisMonthData.to_excel(f'{self.path}/RESULT/SCM/OM/当月物料查询表.xlsx', sheet_name="当月物料查询表", index=False)
        # 当月 启用日期 为空的物料无旧物料
        self.OldMaterialData = self.ThisMonthData[self.ThisMonthData["启用日期"].isnull()]  # 旧物料
        del self.OldMaterialData['计划默认属性']
        del self.OldMaterialData['启用日期']

        self.BOM_data = self.BOM_data.rename(columns={'子件编码': '物料编码'})
        self.BOM_data = self.BOM_data.dropna(subset=["集成时间"])  # 去nan
        self.BOM_data["集成时间"] = self.BOM_data["集成时间"].astype("datetime64")

    def GetNewMaterial(self):  # 获取新物料
        NewMaterialData = self.ThisMonthData[self.ThisMonthData["启用日期"].notnull()]  # 新物料
        NewProductionData = pd.merge(NewMaterialData, self.Routing_data, on=["物料编码"], how="left")
        NewProductionData = NewProductionData.dropna(subset=["版本日期"])  # 去nan
        NewProductionData['下单延时/H'] = (
                (NewProductionData['制单时间'] - NewProductionData['版本日期']) / pd.Timedelta(1, 'H')).astype(
            int)
        NewProductionData.loc[NewProductionData["下单延时/H"] > 72, "创建及时率"] = "超时"  # 计算出来的审批延时大于3天为超时
        NewProductionData.loc[NewProductionData["下单延时/H"] <= 72, "创建及时率"] = "正常"  # 小于等于3天为正常
        # 输出新物料及时率
        self.mkdir(self.path+"/RESULT/SCM/OM")
        NewProductionData.to_excel(f'{self.path}/RESULT/SCM/OM/生产订单创建及时率.xlsx', sheet_name="新物料", index=False)

    def GetOldMaterial(self):  # 获取旧物料

        OldProductionData = pd.merge(self.OldMaterialData, self.BOM_data, on=["物料编码"], how="left")
        OldProductionData = OldProductionData.dropna(subset=["集成时间"])  # 去nan
        OldProductionData['下单延时/H'] = (
                (OldProductionData['制单时间'] - OldProductionData['集成时间']) / pd.Timedelta(1, 'H')).astype(
            int)
        OldProductionData = OldProductionData.drop_duplicates(subset=["生产订单号", "行号", "物料编码"])  # 去重
        OldProductionData.loc[OldProductionData["下单延时/H"] > 24, "创建及时率"] = "超时"  # 计算出来的审批延时大于1天为超时
        OldProductionData.loc[OldProductionData["下单延时/H"] <= 24, "创建及时率"] = "正常"  # 小于等于1天为正常
        # 输出旧物料及时率
        self.mkdir(self.path + "/RESULT/SCM/OM")
        book = load_workbook(f'{self.path}/RESULT/SCM/OM/生产订单创建及时率.xlsx')
        writer = pd.ExcelWriter(f"{self.path}/RESULT/SCM/OM/生产订单创建及时率.xlsx", engine='openpyxl')
        writer.book = book
        OldProductionData.to_excel(writer, "旧物料", index=False)
        writer.save()

    def run(self):
        self.ModifyFormat()
        self.GetNewMaterial()
        self.GetOldMaterial()


if __name__ == '__main__':
    OC = OrderCreation()
    OC.run()
