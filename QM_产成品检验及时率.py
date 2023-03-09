import pandas as pd
from numpy import datetime64
import Func
from openpyxl import load_workbook


class FinishedProduct:
    def __init__(self):
        self.func = Func
        self.ThisMonthStart, self.ThisMonthEnd, self.LastMonthEnd, self.LastMonthStart = self.func.GetDate()
        # 将上月首尾日期切割
        self.LastMonthStart = str(self.LastMonthStart).split(" ")[0]  # .replace("-", "")
        self.ThisMonthStart = str(self.ThisMonthStart).split(" ")[0]
        self.ThisMonthEnd = str(self.ThisMonthEnd).split(" ")[0]
        self.path = Func.Path()

    def mkdir(self, path):
        self.func.mkdir(path)

    def GetFinishedProduct(self):
        # 产成品检验
        # 产成品检验单审核时间-产成品报检单审核时间<24H　
        ProductionData = pd.read_excel(f"{self.path}/DATA/QM/生产时效性统计表.xlsx",
                                       usecols=['检验单号', '报检审核时间', '检验审核时间', '生产订单号码', '行号', '生产批号',
                                                '部门名称', '物料编码', '物料名称', '报检数量'],
                                       header=2,
                                       converters=
                                       {'生产订单号码': str,
                                        '行号': str,
                                        '物料编码': str,
                                        '物料名称': str,
                                        '报检数量': float,
                                        '报检审核时间': datetime64,
                                        '检验审核时间': datetime64})
        ProductionData = ProductionData.dropna(axis=0, how='any')  # 去除所有nan的列
        ProductionData = ProductionData[
            ProductionData['检验审核时间'] >= datetime64(self.ThisMonthStart)]  # 筛选出本月的单据
        # ProductionData = ProductionData[
        #     ProductionData['检验审核时间'] <= datetime64(self.ThisMonthEnd)]  # 筛选出本月的单据
        ProductionData['审批延时'] = (
                (ProductionData['检验审核时间'] - ProductionData['报检审核时间']) / pd.Timedelta(1, 'H')).astype(
            int)
        ProductionData.loc[ProductionData["审批延时"] > 24, "单据状态"] = "超时"
        ProductionData.loc[ProductionData["审批延时"] <= 24, "单据状态"] = "正常"

        try:
            ProductionCount = ProductionData["单据状态"].value_counts()['超时']
        except:
            ProductionCount = 0

        ProductionCountAll = ProductionData.shape[0]
        ProductionResult = format(float(1 - ProductionCount / ProductionCountAll), '.2%')
        dict = {'当月产成品未及时检验数': [ProductionCount], '当月产成品检验总数': [ProductionCountAll], '产成品检验及时率': [ProductionResult]}
        ProductionResult_sheet = pd.DataFrame(dict)

        order = ['检验单号', '生产批号', '部门名称', '生产订单号码', '行号', '物料编码', '物料名称', '报检数量',
                 '报检审核时间', '检验审核时间', '审批延时', '单据状态']
        ProductionData = ProductionData[order]
        self.SaveFile(ProductionData, ProductionResult_sheet)

    def SaveFile(self, ProductionData, ProductionResult_sheet):
        self.mkdir(self.path + '/RESULT/QM')
        ProductionResult_sheet.to_excel(f'{self.path}/RESULT/QM/产成品检验及时率.xlsx', sheet_name="产成品检验及时率", index=False)
        book = load_workbook(f'{self.path}/RESULT/QM/产成品检验及时率.xlsx')
        writer = pd.ExcelWriter(f"{self.path}/RESULT/QM/产成品检验及时率.xlsx", engine='openpyxl')
        writer.book = book
        ProductionData.to_excel(writer, "当月产成品检验情况清单", index=False)
        writer.save()

    def run(self):
        self.GetFinishedProduct()


if __name__ == '__main__':
    FP = FinishedProduct()
    FP.run()
