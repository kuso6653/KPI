import pandas as pd
import calendar
import datetime
from datetime import timedelta
from numpy import datetime64
from openpyxl import load_workbook


def ReformDays(Days):
    now_work_days = []
    for i in Days:
        if i < 10:
            now_work_days.append("0" + str(i))
        else:
            now_work_days.append(str(i))
    return now_work_days


def CheckData(first, two):
    # print(first)
    # print(two)
    del two["物料名称"]
    del two["物料属性"]  # 删除不要的列，以免合并时候出现两列
    out_data = pd.merge(first, two, on=['物料编码', '需求跟踪号', '需求跟踪行号'])
    # out_data = out_data.drop_duplicates()  # 去重
    out_data = out_data.dropna(axis=0, how='any')  # 去除所有nan的列
    print(out_data)


# 获取当月工作日函数
def WorkDays(year, month):
    # 利用日历函数，创建截取工作日日期
    cal = calendar.Calendar()
    work_days = []  # 创建工作日数组
    for week in cal.monthdayscalendar(int(year), int(month)):
        for i, day in enumerate(week):
            # 为0或者大于等于5的为休息日
            if day == 0 or i >= 5:
                continue
            # 否则加入数组
            work_days.append(day)
    return work_days


if __name__ == '__main__':
    now = datetime.date.today()
    # 获取当月首尾日期
    this_month_start = datetime.datetime(now.year, now.month, 1)
    this_month_end = datetime.datetime(now.year, now.month, calendar.monthrange(now.year, now.month)[1])
    # 获取上月首尾日期
    last_month_end = this_month_start - timedelta(days=1)
    last_month_start = datetime.datetime(last_month_end.year, last_month_end.month, 1)
    # 获取截取这个月份、年、上个月
    this_month_start = str(this_month_start).split(" ")[0]
    this_month = this_month_start.split("-")[1]

    year = this_month_start.split("-")[0]

    last_month_start = str(last_month_start).split(" ")[0]
    this_month_end = str(this_month_end).split(" ")[0]
    last_month = last_month_start.split("-")[1]

    last_work_days = WorkDays(year, last_month)  # 获取上个月工作日
    this_work_days = WorkDays(year, this_month)  # 获取这个月工作日

    work_days = []
    work_days.extend(last_work_days[-7:])
    work_days.extend(this_work_days)  # 将上个月最后三天和这个月工作日相合并
    # print(work_days)
    work_days = ReformDays(work_days)  # 改造
    print(work_days)
