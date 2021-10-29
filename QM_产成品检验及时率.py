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

# 产成品检验
# 产成品检验单审核时间-产成品报检单审核时间<24H　
Production_Aging_data = pd.read_excel(f"./DATA/QM/生产时效性统计表-{this_month_start}-{this_month_end}.xlsx",
                                      usecols=['报检审核时间', '检验审核时间', '生产订单号码', '物料编码', '物料名称'],
                                      header=2,
                                      converters=
                                      {'生产订单号码': str,
                                       '物料编码': str,
                                       '物料名称': str,
                                       '报检审核时间': datetime64,
                                       '检验审核时间': datetime64})
Production_Aging_data = Production_Aging_data.dropna(axis=0, how='any')  # 去除所有nan的列
Production_Aging_data['审批延时'] = (
        (Production_Aging_data['检验审核时间'] - Production_Aging_data['报检审核时间']) / pd.Timedelta(1, 'H')).astype(
    int)
Production_Aging_data.loc[Production_Aging_data["审批延时"] > 24, "单据状态"] = "超时"
Production_Aging_data.loc[Production_Aging_data["审批延时"] <= 24, "单据状态"] = "正常"
Production_Aging_data.to_excel('./RESULT/QM/产成品检验及时率.xlsx', sheet_name="产成品检验及时率")
