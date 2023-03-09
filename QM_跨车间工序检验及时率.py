import pandas as pd
import calendar
import datetime
from datetime import timedelta
from numpy import datetime64
import Func
from openpyxl import load_workbook

class CrossWorkshop:
    def __init__(self):
        self.func = Func
        self.ThisMonthStart, self.ThisMonthEnd, self.LastMonthEnd, self.LastMonthStart = self.func.GetDate()
        self.path = Func.Path()
        # 将上月首尾日期切割
        self.ThisMonthStart = str(self.ThisMonthStart).split(" ")[0].replace("-", "")
        self.ThisMonthEnd = str(self.ThisMonthEnd).split(" ")[0].replace("-", "")

    def mkdir(self, path):
        self.func.mkdir(path)

    def GetCrossWorkshopData(self):
        # 跨车间工序检验
        # 工序检验单审核时间-工序报检单审核时间<24H
        Process_Inspection_data = pd.read_excel(
            f"{self.path}/DATA/QM/工序检验单列表.XLSX",
            usecols=['工序检验单号', '生产批号', '生产订单号', '生产订单行号', '报检单号', '存货编码', '物料描述', '报检数量', '审核时间', '生产部门名称'],
            converters={'报检单号': str, '生产订单号': str, '存货编码': str, '审核时间': datetime64})
        Process_Inspection_Application_data = pd.read_excel(
            f"{self.path}/DATA/QM/工序报检单列表.XLSX",
            usecols=['审核时间', '工序报检单号'], converters={'工序报检单号': str, '审核时间': datetime64})

        # 重命名报检单号为工序报检单号，别分命名报检和检验审核时间
        Process_Inspection_data = Process_Inspection_data.rename(columns={'报检单号': '工序报检单号', '审核时间': '检验审核时间'})
        Process_Inspection_Application_data = Process_Inspection_Application_data.rename(columns={'审核时间': '报检审核时间'})

        Process_Inspection_data = Process_Inspection_data.dropna(axis=0, how='any')  # 去除所有nan的列
        Process_Inspection_Application_data = Process_Inspection_Application_data.dropna(axis=0, how='any')  # 去除所有nan的列

        # 合并两个表
        Process_Inspection_all = pd.merge(Process_Inspection_data, Process_Inspection_Application_data, on="工序报检单号")
        del Process_Inspection_all['工序报检单号']
        Process_Inspection_all['审批延时'] = (
                (Process_Inspection_all['检验审核时间'] - Process_Inspection_all['报检审核时间']) / pd.Timedelta(1, 'H')).astype(
            int)
        Process_Inspection_all.loc[Process_Inspection_all["审批延时"] > 24, "单据状态"] = "超时"
        Process_Inspection_all.loc[Process_Inspection_all["审批延时"] <= 24, "单据状态"] = "正常"

        try:
            Process_InspectionCount = Process_Inspection_all['单据状态'].value_counts()['超时']
        except:
            Process_InspectionCount = 0

        Process_InspectionCountAll = Process_Inspection_all.shape[0]
        Process_InspectionResult = format(float(1 - Process_InspectionCount / Process_InspectionCountAll), '.2%')
        dict = {'当月工序流转不及时物料数': [Process_InspectionCount], '当月工序流转物料总数': [Process_InspectionCountAll], '跨车间工序检验及时率': [Process_InspectionResult]}
        Process_InspectionResult_sheet = pd.DataFrame(dict)

        order = ['工序检验单号', '生产批号', '生产部门名称', '生产订单号', '生产订单行号', '存货编码', '物料描述', '报检数量',
                 '报检审核时间', '检验审核时间', '审批延时', '单据状态']
        Process_Inspection_all = Process_Inspection_all[order]
        self.SaveFile(Process_Inspection_all, Process_InspectionResult_sheet)

    def SaveFile(self, Process_Inspection_all, Process_InspectionResult_sheet):
        self.mkdir(self.path+"/RESULT/QM")
        Process_InspectionResult_sheet.to_excel(f'{self.path}/RESULT/QM/跨车间工序检验及时率.xlsx', sheet_name="跨车间工序检验及时率", index=False)
        book = load_workbook(f'{self.path}/RESULT/QM/跨车间工序检验及时率.xlsx')
        writer = pd.ExcelWriter(f"{self.path}/RESULT/QM/跨车间工序检验及时率.xlsx", engine='openpyxl')
        writer.book = book
        Process_Inspection_all.to_excel(writer, "当月工序流转情况清单", index=False)
        writer.save()

    def run(self):
        self.GetCrossWorkshopData()


if __name__ == '__main__':
    CW = CrossWorkshop()
    CW.run()
