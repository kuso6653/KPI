import xlrd
import pandas as pd
import calendar
import datetime
from datetime import timedelta
import openpyxl

from numpy import datetime64
from openpyxl import load_workbook

now = datetime.date.today()

# 获取当月首尾日期
this_month_start = datetime.datetime(now.year, now.month, 1)
this_month_end = datetime.datetime(now.year, now.month, calendar.monthrange(now.year, now.month)[1])

# 获取上月首尾日期
last_month_end = this_month_start - timedelta(days=1)
last_month_start = datetime.datetime(last_month_end.year, last_month_end.month, 1)

# 将上月首尾日期切割
_last_month_end = str(last_month_end).split(" ")[0].replace("", "")
last_month_start = str(last_month_start).split(" ")[0].replace("-", "")
last_month_end = str(last_month_end).split(" ")[0].replace("-", "")
# 材料出库及时率
Production_data = pd.read_excel(f"./KPI/PROD/生产订单列表-{last_month_start}-{last_month_end}.XLSX",
                                usecols=['生产订单号', '行号', '物料编码', '物料名称', '生产批号', '制单时间'],
                                converters={'生产订单号': str, '制单时间': datetime64})
# '母件编码', '母件名称', '旧子件编码',
# '旧子件名称', '旧子件基本用量', '新子件编码', '新子件名称', '新子件基本用量',

Material_data = pd.read_excel(f"./KPI/SCM/OM/存货档案{_last_month_end}.XLSX",
                              usecols=['存货编码', '存货名称', '计划默认属性', '启用日期'],
                              converters={'启用日期': datetime64})
Production_data = Production_data[Production_data["制单时间"] > this_month_start]

# ECN_data = ECN_data.dropna(subset=['生产订单号'])  # 去除nan的列
# ECN_data = ECN_data.drop_duplicates()  # 去重
# ECN_deal_data = ECN_deal_data.dropna(subset=['Ecn单'])  # 去除nan的列
# ECN_data = ECN_data.rename(columns={'单据编号': 'Ecn单'})
# ECN_deal_data = ECN_deal_data.drop_duplicates()  # 去重
# del ECN_data["生产订单号"]
# all_data = pd.merge(ECN_deal_data, ECN_data, on='Ecn单')
# all_data = all_data.drop_duplicates()  # 去重
# all_data['审批延时/H'] = ((all_data['Ecn处理单单据日期'] - all_data['审核日期']) / pd.Timedelta(1, 'H')).astype(
#     int)
#
# all_data.loc[all_data["审批延时/H"] > 24, "单据状态"] = "超时"  # 计算出来的审批延时大于1天为超时
# all_data.loc[all_data["审批延时/H"] <= 24, "单据状态"] = "正常"  # 小于等于1天为正常
#
# all_data.to_excel('./KPI/SCM/OM/生产订单维护及时率.xlsx', sheet_name="生产订单维护及时率", index=False)
