import pandas as pd
import calendar
import datetime
from datetime import timedelta
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

# 跨车间工序检验
# 工序检验单审核时间-工序报检单审核时间<24H
Process_Inspection_data = pd.read_excel(f"./DATA/QM/工序检验单列表-{this_month_start}-{this_month_end}.XLSX",
                                        usecols=['工序检验单号', '生产批号', '生产订单号', '生产订单行号', '报检单号', '存货编码', '物料描述',  '报检数量', '审核时间'],
                                        converters={'报检单号': str, '生产订单号': str, '存货编码': str})
Process_Inspection_Application_data = pd.read_excel(f"./DATA/QM/工序报检单列表-{this_month_start}-{this_month_end}.XLSX",
                                                    usecols=['审核时间', '工序报检单号'], converters={'工序报检单号': str})

# 重命名报检单号为工序报检单号，别分命名报检和检验审核时间
Process_Inspection_data = Process_Inspection_data.rename(columns={'报检单号': '工序报检单号', '审核时间': '检验审核时间'})
Process_Inspection_Application_data = Process_Inspection_Application_data.rename(columns={'审核时间': '报检审核时间'})

Process_Inspection_data = Process_Inspection_data.dropna(axis=0, how='any')  # 去除所有nan的列
Process_Inspection_Application_data = Process_Inspection_Application_data.dropna(axis=0, how='any')  # 去除所有nan的列

# 合并两个表
Process_Inspection_all = pd.merge(Process_Inspection_data, Process_Inspection_Application_data, on="工序报检单号")
del Process_Inspection_all['工序报检单号']
Process_Inspection_all['审批延时'] = (
        (Process_Inspection_all['检验审核时间'] - Process_Inspection_all['报检审核时间']) / pd.Timedelta(1, 'H')).astype(
    int)
Process_Inspection_all.loc[Process_Inspection_all["审批延时"] > 24, "单据状态"] = "超时"
Process_Inspection_all.loc[Process_Inspection_all["审批延时"] <= 24, "单据状态"] = "正常"
# print(Process_Inspection_all)
Process_Inspection_all.to_excel('./RESULT/QM/跨车间工序检验及时率.xlsx', sheet_name="跨车间工序检验及时率",index=False)
