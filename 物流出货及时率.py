import xlrd
import pandas as pd
import calendar
import datetime
from datetime import timedelta
import openpyxl

# 职责要求：审核LOG组数据有效性，依托LOG专员和LOG助理KPI完成情况；
# 数据范围：
# 1）成品出货数据；
# 指标目标值：
# 1）出货及时率（50%）=100%；
# 当月销售出库单审核时间-销售发货单审核时间＞24H/发货单总条数

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

Sale_out_data = pd.read_excel(f"./KPI/SCM/LOGISTIC/销售出库单列表-{last_month_start}-{last_month_end}.XLSX",
                              usecols=['发货单号', '审核时间', '存货编码'],
                              converters={'发货单号': str, '存货编码': str})
Invoice_data = pd.read_excel(f"./KPI/SCM/LOGISTIC/发货单列表-{last_month_start}-{last_month_end}.XLSX",
                             usecols=['发货单号', '审核时间', '存货编码'],
                             converters={'发货单号': str, '存货编码': str})
# usecols=['发货单号', '审核时间'] 为读取指定列名
# usecols='B:E' 为B至E的列
# usecols=[0, 2] 为0至2的列
# converters={'发货单号': str}  解决读取过长数字导致显示为科学计数法

Sale_out_data = Sale_out_data.rename(columns={'审核时间': '销售出库单审核时间'})
Invoice_data = Invoice_data.rename(columns={'审核时间': '发货单审核时间'})

merge_data = pd.merge(Sale_out_data, Invoice_data, on=["发货单号", '存货编码'])
merge_data = merge_data.dropna(axis=0, how='any')  # 去除所有nan的列

t1 = pd.to_datetime(merge_data["销售出库单审核时间"])
t2 = pd.to_datetime(merge_data["发货单审核时间"])
t3 = t1 - t2

merge_data['审批延时'] = ((merge_data['销售出库单审核时间'] - merge_data['发货单审核时间']) / pd.Timedelta(1, 'H')).astype(
    int)
merge_data.loc[merge_data["审批延时"] > 24, "单据状态"] = "超时"
merge_data.loc[merge_data["审批延时"] <= 24, "单据状态"] = "正常"

# all = merge_data["out_time"].count()
# qualify = merge_data["out_time"].loc[merge_data["diff_date_hour"] <= 24].count()

# print("%.2f" % float(qualify / all))
# merge_data["word_time"] = pd.Timedelta(t1 - t2).seconds/3600.00
# print(merge_data["word_time"])
merge_data.to_excel('./KPI/SCM/LOGISTIC/物流出货及时率.xlsx', sheet_name="物流出货及时率")

