import xlrd
import pandas as pd
import calendar
import datetime
from datetime import timedelta
import openpyxl

# 职责要求：录入出库/入库数据至SAP，确保其有效性；
# 数据范围：
# 1）出库/入库数据；
# 指标目标值：
# 1）出库/入库准确率（50%）=100%；
# 2）出库/入库及时性（50%）<72H；
# 当月材料出库单审核时间-当月材料出库单创建时间＞72H / 当月材料出库单总条数当月
# （采购入库单时间-采购到货单时间）＞ 72H/当月采购到货单总条


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

Material_out_data = pd.read_excel(f"./KPI/SCM/WM/材料出库单列表-{last_month_start}-{last_month_end}.XLS",
                                  usecols=['材料编码', '物料描述', '审核时间', '制单时间'],
                                  converters={'材料编码': str})

# usecols=['发货单号', '审核时间'] 为读取指定列名
# usecols='B:E' 为B至E的列
# usecols=[0, 2] 为0至2的列
# converters={'发货单号': str}  解决读取过长数字导致显示为科学计数法


Material_out_data = Material_out_data.dropna(axis=0, how='any')  # 去除所有nan的列
# df.dropna(subset=['name', 'born'])
#
# #删除在'name' 'born'列含有缺失值的行
Material_out_data = Material_out_data.drop_duplicates()  # 去重

Material_out_data['审批延时'] = ((Material_out_data['审核时间'] - Material_out_data['制单时间']) / pd.Timedelta(1, 'H')).astype(
    int)
# 将天数转化为小时数
Material_out_data.loc[Material_out_data["审批延时"] > 73, "单据状态"] = "超时"  # 计算出来的审批延时大于72为超时
Material_out_data.loc[Material_out_data["审批延时"] <= 72, "单据状态"] = "正常"  # 小于等于72为正常

# all = merge_data["out_time"].count()
# qualify = merge_data["out_time"].loc[merge_data["diff_date_hour"] <= 24].count()

# print("%.2f" % float(qualify / all))
Material_out_data.to_excel('./KPI/SCM/WM/仓库出库及时率.xlsx', sheet_name="仓库出库及时率", index=False)
