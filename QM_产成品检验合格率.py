import pandas as pd
import calendar
import datetime
from datetime import timedelta
from numpy import datetime64
from openpyxl import load_workbook

import Func

pd.set_option('display.max_columns', None)


class ProductQualityControl:
    def __init__(self):
        self.func = Func
        self.ThisMonthStart, self.ThisMonthEnd, self.LastMonthEnd, self.LastMonthStart = self.func.GetDate()
        self.path = Func.Path()
        # self.MRPScreenList = []  # 筛选合并的mrp数据
        # self.MRPNewDataList = []  # 本月所有的mrp数据
        # self.PRProcessData = []  # 分组后的mrp数据

        self.ThisMonthStart = str(self.ThisMonthStart).split(" ")[0]
        self.LastMonthStart = str(self.LastMonthStart).split(" ")[0]
        self.ThisMonthEnd = str(self.ThisMonthEnd).split(" ")[0]
        self.this_month_check = self.ThisMonthStart

        OtherThisMonthStart = str(self.ThisMonthStart).split(" ")[0].replace("-", "")
        OtherLastMonthStart = str(self.LastMonthStart).split(" ")[0].replace("-", "")
        OtherThisMonthEnd = str(self.ThisMonthEnd).split(" ")[0].replace("-", "")

        # self.ProductionData = pd.read_excel(f"{self.path}/DATA/PROD/生产订单列表.XLSX",
        #                                     usecols=['生产订单号', '行号', '生产数量'],
        #                                     converters={'行号': str})
        # self.ProductionData = self.ProductionData.rename(columns={'行号': '生产订单行号'})

        self.ProductGRData = pd.read_excel(f"{self.path}/DATA/PROD/产成品入库单列表.xlsx",
                                           usecols=['入库单号', '行号', '表体生产批号', "表体生产订单号", '生产订单行号', '产品编码',
                                                    '产品名称', '规格型号', '数量', '主计量单位', '批号', '表体备注', '制单人', '审核人',
                                                    '制单时间', '审核日期', '表体检验单号'],
                                           converters={'入库单号': str, '行号': str, '生产订单行号': str, '产品编码': str, '批号': str,
                                                       '制单时间': datetime64, '审核日期': datetime64})

        self.ProductGRData = self.ProductGRData.rename(columns={'入库单号': '产成品入库单号', '行号': '产成品入库单行号', '表体生产批号': '生产批号',
                                                                '表体生产订单号': '生产订单号', '产品编码': '存货编码', '产品名称': '存货名称',
                                                                '制单时间': '产成品入库单制单时间', '审核日期': '产成品入库单审核时间',
                                                                '制单人': '产成品入库单制单人', '审核人': '产成品入库单审核人',
                                                                '表体检验单号': '检验单号'})
        # self.ProductData = self.ProductData.dropna(subset=['产成品入库单号'])  # 清除空白行

        self.ProductFailedData = pd.read_excel(f"{self.path}/DATA/QM/产品不良品处理单列表.XLSX",
                                               usecols=['检验单号', '生产订单号', '生产订单行号', '不良品处理单号', '不良品处理日期', '不良品处理时间',
                                                        '不良品数量', '不良品原因', '不良品原因备注', '不良品责任部门', '不良品处理方式',
                                                        '不良品处理流程', '制单人', '审核人', '制单时间', '审核时间'],
                                               converters={'生产订单行号': str, '制单时间': datetime64, '审核时间': datetime64})

        self.ProductFailedData = self.ProductFailedData.rename(columns={'制单时间': '产成品不良品单制单时间', '审核时间': '产成品不良品单审核时间',
                                                                        '制单人': '产成品不良品单制单人', '审核人': '产成品不良品单审核人'
                                                                        })

    def mkdir(self, path):
        self.func.mkdir(path)

    def ThisMonthPQM(self):
        #  self.QMData        已报检的数据
        #  self.NotQMData     已报检未检验的数据
        #  self.NotGRData     已检验未入库的数据
        #  self.QMFailedData  已检验不合格的数据
        #  self.GRData        已检验不合格做让步接收的数据
        # self.ProductGRData = pd.merge(self.ProductGRData, self.ProductionData, how="left", on=['生产订单号', '生产订单行号'])
        self.QMFailedData = pd.merge(self.ProductFailedData, self.ProductGRData, how="left", on=['生产订单号', '生产订单行号', '检验单号'])  # 合并两张报表
        self.QMFailedData = self.QMFailedData[
            self.QMFailedData['产成品不良品单制单时间'] >= datetime64(self.ThisMonthStart)]  # 筛选出本月的单据
        self.QMFailedData = self.QMFailedData[
            self.QMFailedData['产成品不良品单制单时间'] <= datetime64(self.ThisMonthEnd)]  # 筛选出本月的单据
        self.QMFailedData = self.QMFailedData.dropna(subset=['不良品处理单号'])

        self.NotGRData = self.QMFailedData.loc[self.QMFailedData['产成品入库单号'].isnull()]  # 筛选出未入库的单据
        self.GRData = self.QMFailedData.loc[self.QMFailedData['产成品入库单号'].notnull()]  # 筛选出已入库的单据

        self.ProductGRData = self.ProductGRData[
            self.ProductGRData['产成品入库单审核时间'] >= datetime64(self.ThisMonthStart)]  # 筛选出本月的单据
        self.ProductGRData = self.ProductGRData[
            self.ProductGRData['产成品入库单审核时间'] <= datetime64(self.ThisMonthEnd)]  # 筛选出本月的单据
        # self.ProductData = self.ProductData.loc[self.ProductData["不良品处理单号"].notnull()]  # 筛选出本月的单据
        try:
            ProductionCount = self.QMFailedData.shape[0]
        except:
            ProductionCount = 0

        ProductionCountAll = self.ProductGRData.shape[0]
        ProductionResult = format(float(1 - ProductionCount / ProductionCountAll), '.2%')
        dict = {'当月产成品不合格数': [ProductionCount], '当月产成品入库总数': [ProductionCountAll], '产成品检验合格率': [ProductionResult]}
        ProductionResult_sheet = pd.DataFrame(dict)
        self.SaveFile(ProductionResult_sheet)


    def SaveFile(self, ProductionResult_sheet):
        self.mkdir(self.path + '/RESULT/QM')
        ProductionResult_sheet.to_excel(f'{self.path}/RESULT/QM/产成品检验合格率.xlsx', sheet_name="产成品检验合格率", index=False)

        book = load_workbook(f'{self.path}/RESULT/QM/产成品检验合格率.xlsx')
        writer = pd.ExcelWriter(f"{self.path}/RESULT/QM/产成品检验合格率.xlsx", engine='openpyxl')
        writer.book = book
        self.QMFailedData.to_excel(writer, "本月不合格的物料清单", index=False)
        self.ProductGRData.to_excel(writer, "本月已入库的物料清单", index=False)
        # self.NotGRData.to_excel(writer, "本月已质检未入库的物料清单", index=False)
        # self.GRData.to_excel(writer, "本月让步接收的物料清单", index=False)
        writer.save()

    def run(self):
        self.ThisMonthPQM()


if __name__ == '__main__':
    PQC = ProductQualityControl()
    PQC.run()
