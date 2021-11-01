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
# 导入采购时效性表格
Purchase_in_data = pd.read_excel(f"./DATA/SCM/采购时效性统计表-{this_month_start}-{this_month_end}.XLSX",
                                 usecols=[1, 6, 7, 12, 19, 22, 26, 30], header=2,
                                 names=["订单号", "存货编码", "存货名称", "订单制单时间", "报检单号", "报检审核时间", "检验审核时间", "入库制单时间"],
                                 converters={'订单制单时间': datetime64, '报检审核时间': datetime64, '检验审核时间': datetime64,
                                             '入库制单时间': datetime64, '存货编码': float, "订单号": str})
# 来料检验
# 来料检验单审核时间-来料报检单审核时间<48H
# 跨车间工序检验
# 建立质检空表
Purchase_in_data = Purchase_in_data.dropna(subset=['报检审核时间'])  # 去除nan的列

# 将质检时间导入形成一个新表

Purchase_in_data['审批延时'] = ((Purchase_in_data['检验审核时间'] - Purchase_in_data['报检审核时间']) / pd.Timedelta(1, 'H')).astype(int)
Purchase_in_data.loc[Purchase_in_data["审批延时"] > 24, "单据状态"] = "超时"  # 计算出来的质检的审批延时大于24为超时
Purchase_in_data.loc[Purchase_in_data["审批延时"] <= 24, "单据状态"] = "正常"  # 小于等于24为正常

Purchase_in_data.to_excel('./RESULT/QM/来料检验及时率.xlsx', sheet_name="来料检验及时率", index=False)
