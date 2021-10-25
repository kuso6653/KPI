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
# 材料出库及时率
Purchase_in_data = pd.read_excel(f"./KPI/SCM/WM/采购时效性统计表-{last_month_start}-{last_month_end}.XLSX",
                                 usecols=[1, 6, 7, 12, 22, 26, 30], header=3,
                                 names=["订单号", "存货编码", "存货名称", "订单制单时间", "报检审核时间", "检验审核时间", "入库制单时间"],
                                 converters={'订单制单时间': datetime64, '报检审核时间': datetime64, '检验审核时间': datetime64,
                                             '入库制单时间': datetime64, '存货编码': float})

Purchase_in_data = Purchase_in_data.dropna(axis=0, how='any')  # 去除所有nan的列
Purchase_in_data['审批延时'] = ((Purchase_in_data['入库制单时间'] - Purchase_in_data['订单制单时间']
                             - (Purchase_in_data['检验审核时间'] - Purchase_in_data['报检审核时间']))
                            / pd.Timedelta(1, 'H')).astype(int)  # 制单时间相减，然后减去 质检的审核时间
# 将天数转化为小时数
Purchase_in_data.loc[Purchase_in_data["审批延时"] > 72, "单据状态"] = "超时"  # 计算出来的审批延时大于72为超时
Purchase_in_data.loc[Purchase_in_data["审批延时"] <= 72, "单据状态"] = "正常"  # 小于等于72为正常

# 材料出库及时率
Material_out_data = pd.read_excel(f"./KPI/SCM/WM/材料出库单列表-{last_month_start}-{last_month_end}.XLSX",
                                  usecols=['出库单号', '材料编码', '物料描述', '审核时间', '制单时间'],
                                  converters={'材料编码': str})

Material_out_data = Material_out_data.dropna(axis=0, how='any')  # 去除所有nan的列
Material_out_data = Material_out_data.drop_duplicates()  # 去重
Material_out_data['审批延时'] = ((Material_out_data['审核时间'] - Material_out_data['制单时间']) / pd.Timedelta(1, 'H')).astype(
    int)
# 将天数转化为小时数
Material_out_data.loc[Material_out_data["审批延时"] > 72, "单据状态"] = "超时"  # 计算出来的审批延时大于72为超时
Material_out_data.loc[Material_out_data["审批延时"] <= 72, "单据状态"] = "正常"  # 小于等于72为正常
# 保存
Purchase_in_data.to_excel('./KPI/SCM/WM/仓库出入库及时率.xlsx', sheet_name="仓库入库及时率", index=False)
Material_out_data.to_excel('./KPI/SCM/WM/仓库出入库及时率.xlsx', sheet_name="仓库出库及时率", index=False)
