from decimal import Decimal
import pandas as pd
import Func


class WorkReport:
    def __init__(self):
        self.WorkReportList = []
        self.func = Func
        self.ThisMonthStart, self.ThisMonthEnd, self.LastMonthEnd, self.LastMonthStart = self.func.GetDate()
        self.path = "//10.56.164.127/it&m/KPI"
        # 将上月首尾日期切割
        self.ThisMonthStart = str(self.ThisMonthStart).split(" ")[0].replace("-", "")
        self.ThisMonthEnd = str(self.ThisMonthEnd).split(" ")[0].replace("-", "")
        self.WorkReportDate = pd.read_excel(
            f"{self.path}/DATA/PROD/报工列表-{self.ThisMonthStart}-{self.ThisMonthEnd}.XLSX",
            usecols=['单据号码', '生产订单', '行号', '物料编码', '物料名称', '移入标准工序', '合格数量', '审核时间', '生产数量'],
            converters={'单据号码': str, '生产订单': str, '合格数量': float})
        self.ProductionData = pd.read_excel(
            f"{self.path}/DATA/PROD/生产订单列表-{self.ThisMonthStart}-{self.ThisMonthEnd}.XLSX",
            usecols=['生产订单号', '完工日期', '行号'],
            converters={'生产订单号': str})
        self.ProductionData = self.ProductionData.rename(columns={'生产订单号': '生产订单', '完工日期': '生产订单完工日期'})

    def mkdir(self, path):
        self.func.mkdir(path)

    def GetWorkReport(self):
        self.WorkReportDate = self.WorkReportDate.rename(columns={'审核时间': '报工单审核时间'})
        self.WorkReportDate = self.WorkReportDate.dropna(subset=['单据号码'])  # 去除nan的列

        # 将Work_data的数据分组保存，
        # 然后取 合格数量 的总值要等于 生产数量
        # 取分组后最后的时间作为
        for name1, group in self.WorkReportDate.groupby(["生产订单", "行号", '物料编码', '移入标准工序']):
            group = pd.DataFrame(group)  # 新建pandas
            group = group.sort_values(by='报工单审核时间', ascending=False)  # 降序排序
            qualified_num = group['合格数量'].sum()  # 取合格数量总值保留两位
            qualified_num = Decimal(qualified_num).quantize(Decimal('0.00'))
            group.loc[:, "总合格数量"] = qualified_num  # 新建 总合格数量 列
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
        EndWorkReport.loc[EndWorkReport["报工单审核时间"] <= EndWorkReport["生产订单完工日期"], "报工状态"] = "及时"
        EndWorkReport.loc[EndWorkReport["报工单审核时间"] > EndWorkReport["生产订单完工日期"], "报工状态"] = "不及时"
        order = ['单据号码', '生产订单', '行号', '物料编码', '物料名称', '移入标准工序', '生产数量', '总合格数量', '完整状态', '生产订单完工日期', '报工单审核时间', '报工状态']
        EndWorkReport = EndWorkReport[order]
        self.SaveFile(EndWorkReport)

    def SaveFile(self, EndWorkReport):
        self.mkdir(self.path+'/RESULT/PROD')
        EndWorkReport.to_excel(f"{self.path}/RESULT/PROD/报工及时率和完整率.xlsx", sheet_name="报工及时率和完整率", index=False)

    def run(self):
        self.GetWorkReport()


if __name__ == '__main__':
    WR = WorkReport()
    WR.run()
