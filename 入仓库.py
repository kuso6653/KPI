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
# 职责要求：录入出库/入库数据至SAP，确保其有效性；
# 数据范围：
# 1）出库/入库数据；
# 指标目标值：
# 1）出库/入库准确率（50%）=100%；
# 2）出库/入库及时性（50%）<72H；
# 当月材料出库单审核时间-当月材料出库单创建时间＞72H / 当月材料出库单总条数当月
# （采购入库单时间-采购到货单时间）＞ 72H/当月采购到货单总条
# 导入采购时效性表格
Purchase_in_data = pd.read_excel(f"./KPI/SCM/WM/采购时效性统计表-{last_month_start}-{last_month_end}.XLSX",
                                 usecols=[6, 7, 12, 22, 26, 30], header=3, names=["存货编码", "存货名称", "订单制单时间", "报检审核时间", "检验审核时间", "入库制单时间"],
                                 converters={'订单制单时间': datetime64, '报检审核时间': datetime64 ,'检验审核时间': datetime64 ,'入库制单时间': datetime64, '存货编码': float})


Purchase_in_data = Purchase_in_data.dropna(axis=0, how='any')  # 去除所有nan的列
# Purchase_in_data = Purchase_in_data.drop_duplicates()  # 去重

Purchase_in_data['审批延时'] = ((Purchase_in_data['入库制单时间'] - Purchase_in_data['订单制单时间']
                             - (Purchase_in_data['检验审核时间'] - Purchase_in_data['报检审核时间']))
                            / pd.Timedelta(1, 'H')).astype(int)  # 制单时间相减，然后减去 质检的审核时间
# 将天数转化为小时数
Purchase_in_data.loc[Purchase_in_data["审批延时"] > 73, "单据状态"] = "超时"  # 计算出来的审批延时大于72为超时
Purchase_in_data.loc[Purchase_in_data["审批延时"] <= 72, "单据状态"] = "正常"  # 小于等于72为正常
Purchase_in_data.to_excel('./KPI/SCM/WM/仓库入库及时率.xlsx', sheet_name="仓库入库及时率", index=False)


# 来料检验
# 来料检验单审核时间-来料报检单审核时间<48H
# 跨车间工序检验
# 建立质检空表
QM_data = pd.DataFrame(dtype=datetime64)

QM_data['报检审核时间'] = Purchase_in_data['报检审核时间']
QM_data['检验审核时间'] = Purchase_in_data['检验审核时间']
# 将质检时间导入形成一个新表

QM_data['审批延时'] = ((QM_data['检验审核时间'] - QM_data['报检审核时间']) / pd.Timedelta(1, 'H')).astype(int)
QM_data.loc[QM_data["审批延时"] > 25, "单据状态"] = "超时"  # 计算出来的质检的审批延时大于24为超时
QM_data.loc[QM_data["审批延时"] <= 24, "单据状态"] = "正常"  # 小于等于24为正常


QM_data.to_excel('./KPI/QM/来料检验及时率.xlsx', sheet_name="来料检验及时率", index=False)
