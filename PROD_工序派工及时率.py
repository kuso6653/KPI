import pandas as pd
import calendar
import datetime
from datetime import timedelta
from numpy import datetime64

all_data_work = []
pd.set_option('display.max_columns', None)


def ReformDays(Days):
    now_work_days = []
    for day in Days:
        if day < 10:
            now_work_days.append("0" + str(day))
        else:
            now_work_days.append(str(day))
    return now_work_days


def CheckDataWork(first, two):
    global all_data_work
    first = first.dropna(subset=['物料编码'])  # 去除nan的列
    two = two.dropna(subset=['物料编码'])  # 去除nan的列
    out_data = pd.merge(first.drop(labels=['派工标识'], axis=1), two, on=['物料编码', '生产订单', '工序行号', '行号'])
    out_data = out_data.loc[out_data['派工标识'] != "*"]
    all_data_work.append(out_data)


# 获取当月工作日函数
def WorkDays(year, month):
    # 利用日历函数，创建截取工作日日期
    cal = calendar.Calendar()
    Work_Days = []  # 创建工作日数组
    for week in cal.monthdayscalendar(int(year), int(month)):
        for i, day in enumerate(week):
            # 为0或者大于等于5的为休息日
            if day == 0 or i >= 5:
                continue
            # 否则加入数组
            Work_Days.append(day)
    return Work_Days


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
    last_month = last_month_start.split("-")[1]

    last_work_days = WorkDays(year, last_month)  # 获取上个月工作日
    this_work_days = WorkDays(year, this_month)  # 获取这个月工作日
    work_days = []  # 设置到上月3天
    work_days.extend(last_work_days[-3:])
    work_days.extend(this_work_days)  # 将上个月最后三天和这个月工作日相合并

    work_days = ReformDays(work_days)  # 改造
    base_data_work = []
    flag = 0
    for x in work_days:
        if flag < 3:
            try:
                base_data = pd.read_excel(f"./DATA/PROD/工序派工资料维护{year}-{last_month}-{x}.XLSX",
                                          usecols=['物料编码', '生产订单', '工序行号', '派工标识', '行号'],
                                          converters={'物料编码': int, '工序行号': int}
                                          )
                base_data_work.append(base_data)
                flag = flag + 1
                continue
            except:
                flag = flag + 1
                continue
        else:
            try:
                now_data = pd.read_excel(f"./DATA/PROD/工序派工资料维护{year}-{this_month}-{x}.XLSX",
                                         usecols=['物料编码', '生产订单', '工序行号', '派工标识', '行号'],
                                         converters={'物料编码': int, '工序行号': int}
                                         )
            except:
                continue
            base_data_work.append(now_data)  # 新添加新的base
            CheckDataWork(base_data_work[0], now_data)  # 合并检查是否存在一样的
            del (base_data_work[0])  # 删除第一个base

    res = pd.concat(all_data_work, axis=0, ignore_index=True)
    res = res.drop_duplicates()
    res.to_excel(f'./RESULT/PROD/工序派工及时率.xlsx', sheet_name="工序派工及时率", index=False)
