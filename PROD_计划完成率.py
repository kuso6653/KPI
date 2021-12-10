import pandas as pd
import Func
from numpy import datetime64


class Plan:
    def __init__(self):
        self.path = "//10.56.164.228/KPI"
        self.func = Func
        self.ThisMonthStart, self.ThisMonthEnd, self.LastMonthEnd, self.LastMonthStart = self.func.GetDate()

        # 将上月首尾日期切割
        self.ThisMonthStart = str(self.ThisMonthStart).split(" ")[0].replace("-", "")
        self.ThisMonthEnd = str(self.ThisMonthEnd).split(" ")[0].replace("-", "")

    def mkdir(self, path):
        self.func.mkdir(path)

    def GetPlan(self):
        GoodsInData = pd.read_excel(
            f"{self.path}/DATA/PROD/产成品入库单列表-{self.ThisMonthStart}-{self.ThisMonthEnd}.XLSX",
            usecols=['表体生产订单号', '生产订单行号', '制单时间', '产品编码'],
            converters={'表体生产订单号': str, '制单时间': datetime64})
        ProductionData = pd.read_excel(f"{self.path}/DATA/PROD/生产订单列表-{self.ThisMonthStart}-{self.ThisMonthEnd}.XLSX",
                                       usecols=['生产订单号', '物料名称', '实际完工日期', '行号'],
                                       converters={'生产订单号': str, '实际完工日期': datetime64})

        GoodsInData = GoodsInData.rename(
            columns={'表体生产订单号': '生产订单号', '生产订单行号': '行号', '制单时间': '产成品入库单制单时间', '产品编码': '物料编码'})
        ProductionData = ProductionData.rename(columns={'实际完工日期': '生产订单完工日期'})

        GoodsInData = GoodsInData.dropna(subset=['生产订单号'])  # 去除nan的列
        ProductionData = ProductionData.dropna(subset=['生产订单号'])  # 去除nan的列

        PlanData = pd.merge(GoodsInData, ProductionData, on=['生产订单号', '行号'])

        PlanData['审批延时/H'] = ((PlanData['产成品入库单制单时间'] - PlanData['生产订单完工日期']) / pd.Timedelta(1, 'H')).astype(int)
        # 将天数转化为小时数
        PlanData.loc[PlanData["审批延时/H"] > 48, "单据状态"] = "超时"  # 计算出来的审批延时大于72为超时
        PlanData.loc[PlanData["审批延时/H"] <= 48, "单据状态"] = "正常"  # 小于等于72为正常

        order = ['生产订单号', '行号', '物料编码', '物料名称', '产成品入库单制单时间', '生产订单完工日期', '审批延时/H', '单据状态']
        PlanData = PlanData[order]
        self.SaveFile(PlanData)

    def SaveFile(self, PlanData):
        self.mkdir(self.path + '/RESULT/PROD')
        PlanData.to_excel(f'{self.path}/RESULT/PROD/计划完成率.xlsx', sheet_name="计划完成率", index=False)

    def run(self):
        self.GetPlan()


if __name__ == '__main__':
    P = Plan()
    P.run()
