from decimal import Decimal
import pandas as pd
import Func
from numpy import datetime64
from openpyxl import load_workbook


class WorkReport:
    def __init__(self):
        self.WorkReportList = []
        # self.Workshop1 = []  # 生产运营部-机加工生产
        # self.Workshop2 = []  # 生产运营部-电控柜车间生产
        # self.Workshop3 = []  # 生产运营部-铆焊生产
        # self.Workshop4 = []  # 生产运营部-装配生产-PX
        # self.Workshop5 = []  # 销售部-售后服务
        # self.Workshop6 = []  # 生产运营部-装配生产
        self.func = Func
        self.ThisMonthStart, self.ThisMonthEnd, self.LastMonthEnd, self.LastMonthStart = self.func.GetDate()
        self.path = Func.Path()

        # 将上月首尾日期切割
        # self.ThisMonthStart = str(self.ThisMonthStart).split(" ")[0].replace("-", "")
        # self.ThisMonthEnd = str(self.ThisMonthEnd).split(" ")[0].replace("-", "")
        self.WorkReportDate = pd.read_excel(
            f"{self.path}/DATA/PROD/报工列表.XLSX",
            usecols=['单据日期', '单据号码', '生产订单', '行号', '物料编码', '物料名称', '移入标准工序', '移入工作中心', '合格数量', '审核时间', '生产数量'],
            converters={'单据日期': datetime64, '单据号码': str, '生产订单': str, '合格数量': float})
        self.ProductionData = pd.read_excel(
            f"{self.path}/DATA/PROD/生产订单列表.XLSX",
            usecols=['生产订单号', '实际完工日期', '行号'],
            converters={'生产订单号': str})
        self.ProductionData = self.ProductionData.rename(columns={'生产订单号': '生产订单'})
        self.WorkCenterData = pd.read_excel(
            f"{self.path}/DATA/PROD/工作中心维护.XLSX",
            usecols=['工作中心代号', '部门名称'],
            converters={'工作中心代号': str})
        self.WorkCenterData = self.WorkCenterData.rename(columns={'工作中心代号': '移入工作中心', '部门名称': '移入生产部门'})

    def mkdir(self, path):
        self.func.mkdir(path)

    def GetWorkReport(self):
        self.WorkReportDate = self.WorkReportDate.rename(columns={'审核时间': '报工单审核时间'})
        self.WorkReportDate = self.WorkReportDate.dropna(subset=['单据号码'])  # 去除nan的列
        self.WorkReportDate = pd.merge(self.WorkReportDate, self.WorkCenterData, how="left", on=['移入工作中心'])  # 匹配移入生产部门

        # 将Work_data的数据分组保存，
        # 然后取 合格数量 的总值要等于 生产数量
        # 取分组后最后的时间作为
        for name1, group in self.WorkReportDate.groupby(["生产订单", "行号", '物料编码', '移入标准工序']):
            group = pd.DataFrame(group)  # 新建pandas
            group = group.sort_values(by='报工单审核时间', ascending=False)  # 降序排序
            qualified_num = group['合格数量'].sum()  # 取合格数量总值保留两位
            qualified_num = Decimal(qualified_num).quantize(Decimal('0.00'))
            group.loc[:, "总合格数量"] = qualified_num  # 新建 总合格数量 列
            # a[0:0:0] [start:end:step]
            max_data_list = group.head(1)  # 取最后生产日期，就是排序后的第一列
            self.WorkReportList.append(max_data_list)  # 加入list

        WorkReportDateMerge = pd.concat(self.WorkReportList, axis=0, ignore_index=True)
        # print(end_merge_data)
        del WorkReportDateMerge["合格数量"]

        # 计算出来的 报工完整状态
        # 生产数量 等于 总合格数量 为 完整
        # 生产数量 大于 总合格数量 为 不完整
        WorkReportDateMerge.loc[WorkReportDateMerge["生产数量"] == WorkReportDateMerge["总合格数量"], "完整状态"] = "完整"
        WorkReportDateMerge.loc[WorkReportDateMerge["生产数量"] > WorkReportDateMerge["总合格数量"], "完整状态"] = "不完整"
        WorkReportDateMerge.loc[WorkReportDateMerge["生产数量"] < WorkReportDateMerge["总合格数量"], "完整状态"] = "超定额生产"

        EndWorkReport = pd.merge(WorkReportDateMerge, self.ProductionData, on=["生产订单", "行号"])
        EndWorkReport['实际完工日期'] = pd.to_datetime(EndWorkReport['实际完工日期'].astype(str)) + pd.to_timedelta(
            '20:00:00')
        EndWorkReport['实际比较日期'] = pd.to_datetime(EndWorkReport['实际完工日期'].astype(str)) + pd.to_timedelta(
            '92:00:00')  # 超出实际完工日期内3天均算正常
        EndWorkReport.loc[(EndWorkReport["报工单审核时间"] <= EndWorkReport["实际比较日期"]) & (EndWorkReport["完整状态"] != "完整"),
                          "报工状态"] = "不及时"
        EndWorkReport.loc[(EndWorkReport["报工单审核时间"] <= EndWorkReport["实际比较日期"]) & (EndWorkReport["完整状态"] == "完整"),
                          "报工状态"] = "及时"
        EndWorkReport.loc[EndWorkReport["报工单审核时间"] > EndWorkReport["实际比较日期"], "报工状态"] = "不及时"
        EndWorkReport = EndWorkReport[EndWorkReport['单据日期'] >= datetime64(self.ThisMonthStart)]
        EndWorkReport = EndWorkReport[EndWorkReport['单据日期'] <= datetime64(self.ThisMonthEnd)]
        EndWorkReport['状态合并'] = EndWorkReport['完整状态'] + EndWorkReport['报工状态']

        try:
            EndWorkCount = EndWorkReport["状态合并"].value_counts()['完整及时']
        except:
            EndWorkCount = 0

        EndWorkCountAll = EndWorkReport.shape[0]
        EndWorkResult = format(float(EndWorkCount / EndWorkCountAll), '.2%')
        dict = {'合格报工总数': [EndWorkCount], '当月总报工数': [EndWorkCountAll], '报工及时率和完整率': [EndWorkResult]}
        EndWorkResult_sheet = pd.DataFrame(dict)
        order = ['单据日期', '单据号码', '生产订单', '行号', '物料编码', '物料名称', '移入标准工序', '移入生产部门', '生产数量', '总合格数量', '完整状态', '实际完工日期',
                 '报工单审核时间', '报工状态']
        EndWorkReport = EndWorkReport[order]
        self.SaveFile(EndWorkReport, EndWorkResult_sheet)

    def SaveFile(self, EndWorkReport, EndWorkResult_sheet):
        # self.Workshop1 = EndWorkReport.loc[EndWorkReport['移入工作中心'].str.contains('生产运营部-机加工生产')]
        # self.Workshop2 = EndWorkReport.loc[EndWorkReport['移入工作中心'].str.contains('生产运营部-电控柜车间生产')]
        # self.Workshop3 = EndWorkReport.loc[EndWorkReport['移入工作中心'].str.contains('生产运营部-铆焊生产')]
        # self.Workshop4 = EndWorkReport.loc[EndWorkReport['移入工作中心'].str.contains('生产运营部-装配生产-PX')]
        # self.Workshop5 = EndWorkReport.loc[EndWorkReport['移入工作中心'].str.contains('销售部-售后服务')]
        # self.Workshop6 = EndWorkReport.loc[EndWorkReport['移入工作中心'].str.contains('生产运营部-装配生产')]

        self.mkdir(self.path + '/RESULT/PROD')
        EndWorkResult_sheet.to_excel(f"{self.path}/RESULT/PROD/报工及时率和完整率.xlsx", sheet_name="报工及时率和完整率", index=False)
        book = load_workbook(f'{self.path}/RESULT/PROD/报工及时率和完整率.xlsx')
        writer = pd.ExcelWriter(f"{self.path}/RESULT/PROD/报工及时率和完整率.xlsx", engine='openpyxl')
        writer.book = book
        EndWorkReport.to_excel(writer, "当月报工清单", index=False)
        writer.save()

    def run(self):
        self.GetWorkReport()


if __name__ == '__main__':
    WR = WorkReport()
    WR.run()
