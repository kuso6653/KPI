import pandas as pd
import calendar
import datetime
from datetime import timedelta
from numpy import datetime64
from openpyxl import load_workbook

import Func

pd.set_option('display.max_columns', None)


class QualityControl:
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

    def mkdir(self, path):
        self.func.mkdir(path)

    def ThisMonthQM(self):
        #  self.QMData_temp2  已报检的数据
        #  self.NotQMData     已报检未检验的数据
        #  self.NotGRData     已检验未入库的数据
        #  self.DefectivePrd  已检验不合格的数据
        #  self.GRData        已检验且已入库的数据
        #  self.CGRData       已检验不合格做让步接收的数据

        self.DefectivePrd_List = pd.read_excel(f"{self.path}/DATA/QM/来料不良品处理单列表.xlsx",
                                               usecols=['采购/委外订单号', '不良品处理单号', '批号', "存货编码", '不良品数量', '不良品原因',
                                                        '不良品原因备注', '不良品责任部门',
                                                        '不良品处理方式', '不良品处理流程', '降级后存货编码', '降级后存货名称', '制单人', '审核人'],
                                               converters={'采购/委外订单号': str, '存货编码': str, '降级后存货编码': str})

        self.DefectivePrd_List = self.DefectivePrd_List.rename(columns={'采购/委外订单号': '采购订单号'})
        self.DefectivePrd_List = self.DefectivePrd_List.dropna(subset=['采购订单号'])  # 清除空白行

        self.PO_List = pd.read_excel(f"{self.path}/DATA/SCM/OP/采购订单列表.xlsx",
                                     usecols=['订单编号', '行号', '行备注'],
                                     converters={'订单编号': str, '行号': str, '行备注': str})

        self.PO_List = self.PO_List.rename(columns={'订单编号': '采购订单号', '行号': '采购订单行号'})
        self.PO_List = self.PO_List.dropna(subset=['采购订单号'])  # 清除空白行

        self.PRProcessData = pd.read_excel(f"{self.path}/DATA/SCM/OP/请购执行进度表.XLSX",
                                           usecols=[1, 3, 6, 7, 8, 9, 10, 14, 15, 16, 19, 20, 23, 24, 25, 27, 29, 31,
                                                    33, 34, 35, 37, 38, 39, 40, 42, 43], header=4,
                                           names=["请购单号", "请购单审核日期", "请购单行号", "存货编码", "存货名称", "规格型号", "数量",
                                                  "采购订单号", "采购订单行号", "采购订单下单日期", "供应商简称", "计划到货日期", "采购订单制单人",
                                                  "到货单号", "到货单行号", "到货单审核日期",
                                                  "来料报检单号", "来料报检单审核日期", "来料检验单号", "来料检验单制单日期", "来料检验单审核日期",
                                                  "检验合格数量", "检验不合格数量", "入库单号", "入库单行号", "入库单审核日期", "入库数量"],
                                           converters={'请购单号': str, '采购订单号': str, '采购订单行号': str, '存货编码': str,
                                                       '入库单号': str, '来料检验单制单日期': datetime64})

        self.QMData = pd.merge(self.PRProcessData, self.DefectivePrd_List, how="left", on=['采购订单号', '存货编码'])  # 合并两张报表
        self.QMData_temp1 = pd.merge(self.QMData, self.PO_List, how="left", on=['采购订单号', '采购订单行号'])  # 抓取行备注字段
        self.QMData_temp1 = self.QMData_temp1[self.QMData_temp1['来料检验单制单日期'] >= datetime64(self.ThisMonthStart)]  # 筛选出本月的单据
        self.QMData_temp1 = self.QMData_temp1[self.QMData_temp1['来料检验单制单日期'] <= datetime64(self.ThisMonthEnd)]  # 筛选出本月的单据
        self.QMData_temp2 = self.QMData_temp1.loc[self.QMData_temp1["来料报检单号"].notnull()]  # 筛出已报检的行
        self.NotQMData = self.QMData_temp2.loc[self.QMData_temp2["来料检验单号"].isnull()]  # 筛出未质检的行
        self.NotGRData = self.NotQMData.loc[self.NotQMData["入库单号"].isnull()]  # 筛出已质检未入库的行
        self.GRData = self.QMData_temp1.loc[self.QMData_temp1["入库单号"].notnull()]  # 筛出已完成质检并且入库的行
        # self.QMData = self.QMData.dropna(subset=['请购单号'])  # 去除nan的列
        self.DefectivePrd_temp1 = self.QMData_temp1.loc[self.QMData_temp1["检验不合格数量"] > 0]  # 筛出质检不合格的临时表1
        self.DefectivePrd_temp2 = self.QMData_temp1.loc[self.QMData_temp1["不良品数量"] > 0]  # 筛出质检不合格的临时表2
        self.DefectivePrd = pd.concat([self.DefectivePrd_temp1, self.DefectivePrd_temp2],
                                      ignore_index=True)  # 筛出质检不合格的表
        self.CGRData = self.GRData.loc[self.GRData["检验不合格数量"] > 0]  # 筛出让步接收的行
        self.QMData_temp1 = self.QMData_temp1.loc[self.QMData_temp1['行备注'].notnull()]

        self.CPPOData = self.QMData_temp1[self.QMData_temp1['行备注'].str.contains('橡')]  # 筛出硫化机项目的行

        self.ExPOData = self.QMData_temp1[self.QMData_temp1['行备注'].str.contains('挤')]  # 筛出挤出机项目的行

        self.PlasticPOData = self.QMData_temp1[~self.QMData_temp1['行备注'].str.contains('橡|挤')]  # 筛出注塑机项目的行
        # self.PlasticPOData = self.QMData_temp1[~self.QMData_temp1['行备注'].str.contains('挤')]  # 筛出挤出机项目的行
        # self.PlasticPOData = pd.concat([PlasticPOData_temp1, PlasticPOData_temp2], ignore_index=True)  # 筛出挤出机项目的行
        try:
            QMCount = self.DefectivePrd.shape[0]
        except:
            QMCount = 0

        QMCountAll = self.GRData.shape[0]
        QMResult = format(float(1 - QMCount / QMCountAll), '.2%')
        dict = {'当月来料不合格数': [QMCount], '当月已入库物料总数': [QMCountAll], '来料检验合格率': [QMResult]}
        QMResult_sheet = pd.DataFrame(dict)
        self.SaveFile(QMResult_sheet)
        
        

    def SaveFile(self, QMResult_sheet):
        self.mkdir(self.path + '/RESULT/QM')
        QMResult_sheet.to_excel(f'{self.path}/RESULT/QM/来料检验合格率.xlsx', sheet_name="来料检验合格率", index=False)
        book = load_workbook(f'{self.path}/RESULT/QM/来料检验合格率.xlsx')
        writer = pd.ExcelWriter(f"{self.path}/RESULT/QM/来料检验合格率.xlsx", engine='openpyxl')
        writer.book = book
        self.DefectivePrd.to_excel(writer, "本月不合格的物料清单", index=False)
        self.GRData.to_excel(writer, "本月已入库的物料清单", index=False)
        self.NotQMData.to_excel(writer, "本月未质检的物料清单", index=False)
        self.CPPOData.to_excel(writer, "硫化机项目物料", index=False)
        self.ExPOData.to_excel(writer, "挤出机项目物料", index=False)
        self.PlasticPOData.to_excel(writer, "塑机项目物料", index=False)
        writer.save()

    def run(self):
        self.ThisMonthQM()
        # self.SaveFile()

if __name__ == '__main__':
    QC = QualityControl()
    QC.run()
