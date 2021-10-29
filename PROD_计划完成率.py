import xlrd
import pandas as pd
import calendar
import datetime
from datetime import timedelta
import openpyxl



now = datetime.date.today()

# 获取当月首尾日期
this_month_start = datetime.datetime(now.year, now.month, 1)
this_month_end = datetime.datetime(now.year, now.month, calendar.monthrange(now.year, now.month)[1])

# 获取上月首尾日期
last_month_end = this_month_start - timedelta(days=1)
last_month_start = datetime.datetime(last_month_end.year, last_month_end.month, 1)

# 将上月首尾日期切割
this_month_start = str(this_month_start).split(" ")[0].replace("-", "")
this_month_end = str(this_month_end).split(" ")[0].replace("-", "")

Finished_goods_in_data = pd.read_excel(f"./DATA/PROD/产成品入库单列表-{this_month_start}-{this_month_end}.XLSX",
                                       usecols=['表体生产订单号', '生产订单行号', '制单时间', '产品编码'],
                                       converters={'表体生产订单号': str})
Production_data = pd.read_excel(f"./DATA/PROD/生产订单列表-{this_month_start}-{this_month_end}.XLSX",
                                usecols=['生产订单号', '物料名称', '完工日期', '行号'],
                                converters={'生产订单号': str})


Finished_goods_in_data = Finished_goods_in_data.rename(columns={'表体生产订单号': '生产订单号', '生产订单行号': '行号', '制单时间': '产成品入库单制单时间', '产品编码': '物料编码'})
Production_data = Production_data.rename(columns={'完工日期': '生产订单完工日期'})

Finished_goods_in_data = Finished_goods_in_data.dropna(subset=['生产订单号'])  # 去除nan的列
Production_data = Production_data.dropna(subset=['生产订单号'])  # 去除nan的列

all_data = pd.merge(Finished_goods_in_data, Production_data, on=['生产订单号', '行号'])

all_data['审批延时/H'] = ((all_data['产成品入库单制单时间'] - all_data['生产订单完工日期']) / pd.Timedelta(1, 'H')).astype(
    int)
# 将天数转化为小时数
all_data.loc[all_data["审批延时/H"] > 48, "单据状态"] = "超时"  # 计算出来的审批延时大于72为超时
all_data.loc[all_data["审批延时/H"] <= 48, "单据状态"] = "正常"  # 小于等于72为正常

# all = merge_data["out_time"].count()
# qualify = merge_data["out_time"].loc[merge_data["diff_date_hour"] <= 24].count()

# print("%.2f" % float(qualify / all))
order = ['生产订单号', '行号', '物料编码', '物料名称', '产成品入库单制单时间', '生产订单完工日期', '审批延时/H', '单据状态']
all_data = all_data[order]

all_data.to_excel('./RESULT/PROD/计划完成率.xlsx', sheet_name="计划完成率", index=False)
