import xlrd
import pandas as pd
import calendar
import datetime
from datetime import timedelta
import openpyxl

from numpy import datetime64

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
this_month_start = str(this_month_start).split(" ")[0].replace("-", "")
this_month_end = str(this_month_end).split(" ")[0].replace("-", "")
# 职责要求：录入出库/入库数据至SAP，确保其有效性；
# 数据范围：
# 1）出库/入库数据；
# 指标目标值：
# 1）出库/入库准确率（50%）=100%；
# 2）出库/入库及时性（50%）<72H；
# 当月材料出库单审核时间-当月材料出库单创建时间＞72H / 当月材料出库单总条数当月
# （采购入库单时间-采购到货单时间）＞ 72H/当月采购到货单总条
# 导入采购时效性表格  到货单列表-20210901-20211031
goods_in_data = pd.read_excel(f"./KPI/SCM/OP/到货单列表-{last_month_start}-{this_month_end}.XLSX",
                              usecols=['存货编码', '存货名称', '采购委外订单号', '行号', '制单时间'],
                              converters={'行号': int, '存货编码': str, '制单时间': datetime64}
                              )
Purchase_in_data = pd.read_excel(f"./KPI/SCM/OP/采购订单列表-{last_month_start}-{last_month_end}.XLSX",
                                 usecols=['存货编码', '存货名称', '订单编号', '行号', '计划到货日期'],
                                 converters={'存货编码': str, '行号': int, '计划到货日期': datetime64}
                                 )
goods_in_data = goods_in_data.dropna(subset=['存货编码'])  # 去除nan的列
Purchase_in_data = Purchase_in_data.dropna(subset=['存货编码'])  # 去除nan的列
Purchase_in_data = Purchase_in_data.rename(columns={'订单编号': '采购委外订单号'})

all_data = pd.merge(goods_in_data, Purchase_in_data, on=['存货编码', '存货名称', '采购委外订单号', '行号'])
# all_data["out_data"] =all_data["制单时间"]-all_data["计划到货日期"]
all_data["out_data/H"] = ((all_data["制单时间"]-all_data["计划到货日期"]) / pd.Timedelta(1, 'H')).astype(int)
all_data = all_data.loc[all_data["out_data/H"] > 72]

all_data.to_excel('./KPI/SCM/OP/到货单及时率.xlsx', sheet_name="到货单及时率", index=False)
