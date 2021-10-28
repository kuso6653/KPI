from decimal import Decimal

import xlrd
import pandas as pd
import calendar
import datetime
from datetime import timedelta

all_list = []
now = datetime.date.today()

# 获取当月首尾日期
this_month_start = datetime.datetime(now.year, now.month, 1)
this_month_end = datetime.datetime(now.year, now.month, calendar.monthrange(now.year, now.month)[1])

# 获取上月首尾日期
last_month_end = this_month_start - timedelta(days=1)
last_month_start = datetime.datetime(last_month_end.year, last_month_end.month, 1)

# 将上月首尾日期切割
last_month_start = str(last_month_start).split(" ")[0].replace("-", "")
last_month_end = str(last_month_end).split(" ")[0].replace("-", "")

Work_data = pd.read_excel(f"./KPI/PROD/报工列表-{last_month_start}-{last_month_end}.XLSX",
                          usecols=['单据号码', '生产订单', '行号', '物料编码', '物料名称', '移入标准工序', '合格数量', '审核时间', '生产数量'],
                          converters={'单据号码': str, '生产订单': str, '合格数量': float})
Production_data = pd.read_excel(f"./KPI/PROD/生产订单列表-{last_month_start}-{last_month_end}.XLSX",
                                usecols=['生产订单号', '完工日期', '行号'],
                                converters={'生产订单号': str})
Production_data = Production_data.rename(columns={'生产订单号': '生产订单', '完工日期': '生产订单完工日期'})
Work_data = Work_data.rename(columns={'审核时间': '报工单审核时间'})
Work_data = Work_data.dropna(subset=['单据号码'])  # 去除nan的列

# 将Work_data的数据分组保存，
# 然后取 合格数量 的总值要等于 生产数量
# 取分组后最后的时间作为
for name1, group in Work_data.groupby(["生产订单", "行号", '物料编码', '移入标准工序']):
    group = pd.DataFrame(group)  # 新建pandas
    group = group.sort_values(by='报工单审核时间', ascending=False)  # 降序排序
    qualified_num = group['合格数量'].sum()  # 取合格数量总值保留两位
    qualified_num = Decimal(qualified_num).quantize(Decimal('0.00'))
    group.loc[:, "总合格数量"] = qualified_num  # 新建 总合格数量 列
    max_data_list = group.head(1)  # 取最后生产日期，就是排序后的第一列
    all_list.append(max_data_list)  # 加入list

complete_merge_data = pd.concat(all_list, axis=0, ignore_index=True)
# print(end_merge_data)
del complete_merge_data["合格数量"]

# 计算出来的 报工完整状态
# 生产数量 等于 总合格数量 为 完整
# 生产数量 大于 总合格数量 为 不完整
complete_merge_data.loc[complete_merge_data["生产数量"] == complete_merge_data["总合格数量"], "完整状态"] = "完整"
complete_merge_data.loc[complete_merge_data["生产数量"] > complete_merge_data["总合格数量"], "完整状态"] = "不完整"
complete_merge_data.loc[complete_merge_data["生产数量"] < complete_merge_data["总合格数量"], "完整状态"] = "超定额生产"

Work_report_data = pd.merge(complete_merge_data, Production_data, on=["生产订单", "行号"])
Work_report_data.loc[Work_report_data["报工单审核时间"] <= Work_report_data["生产订单完工日期"], "报工状态"] = "及时"
Work_report_data.loc[Work_report_data["报工单审核时间"] > Work_report_data["生产订单完工日期"], "报工状态"] = "不及时"
order = ['单据号码', '生产订单', '行号', '物料编码', '物料名称', '移入标准工序', '生产数量', '总合格数量', '完整状态', '生产订单完工日期', '报工单审核时间', '报工状态']
Work_report_data = Work_report_data[order]

Work_report_data.to_excel("./KPI/PROD/报工及时率和完整率.xlsx", sheet_name="报工及时率和完整率", index=False)
