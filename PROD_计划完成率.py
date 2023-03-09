import pandas as pd
import Func
from numpy import datetime64
from openpyxl import load_workbook

class Plan:
    def __init__(self):
        self.path = Func.Path()
        self.func = Func
        self.ThisMonthStart, self.ThisMonthEnd, self.LastMonthEnd, self.LastMonthStart = self.func.GetDate()

        # 将上月首尾日期切割
        # self.ThisMonthStart = str(self.ThisMonthStart).split(" ")[0].replace("-", "")
        # self.ThisMonthEnd = str(self.ThisMonthEnd).split(" ")[0].replace("-", "")

    def mkdir(self, path):
        self.func.mkdir(path)

    def GetPlan(self):
        GoodsInData = pd.read_excel(
            f"{self.path}/DATA/PROD/产成品入库单列表.XLSX",
            usecols=['表体生产订单号', '生产订单行号', '制单时间', '产品编码'],
            converters={'表体生产订单号': str, '制单时间': datetime64})
        ProductionData = pd.read_excel(f"{self.path}/DATA/PROD/生产订单列表.XLSX",
                                       usecols=['生产订单号', '物料名称', '实际完工日期', '行号', '部门名称'],
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
        PlanData = PlanData[PlanData['产成品入库单制单时间'] >= datetime64(self.ThisMonthStart)]
        PlanData = PlanData[PlanData['产成品入库单制单时间'] <= datetime64(self.ThisMonthEnd)]

        try:
            PlanCount = PlanData["单据状态"].value_counts()['超时']
        except:
            PlanCount = 0

        PlanCountAll = PlanData.shape[0]
        PlanResult = format(float(1-(PlanCount / PlanCountAll)), '.2%')
        dict = {'未按照计划完成数': [PlanCount], '当月计划完成总数': [PlanCountAll], '计划完成率': [PlanResult]}
        PlanResult_sheet = pd.DataFrame(dict)
        order = ['生产订单号', '行号', '物料编码', '物料名称', '部门名称', '产成品入库单制单时间', '生产订单完工日期', '审批延时/H', '单据状态']
        PlanData = PlanData[order]
        self.SaveFile(PlanData, PlanResult_sheet)

    def SaveFile(self, PlanData, PlanResult_sheet):
        self.mkdir(self.path + '/RESULT/PROD')
        PlanResult_sheet.to_excel(f'{self.path}/RESULT/PROD/计划完成率.xlsx', sheet_name="计划完成率", index=False)
        book = load_workbook(f'{self.path}/RESULT/PROD/计划完成率.xlsx')
        writer = pd.ExcelWriter(f"{self.path}/RESULT/PROD/计划完成率.xlsx", engine='openpyxl')
        writer.book = book
        PlanData.to_excel(writer, "当月计划完成情况清单", index=False)
        writer.save()

    def run(self):
        self.GetPlan()


if __name__ == '__main__':
    P = Plan()
    P.run()
