import pandas as pd
import calendar
import datetime
from datetime import timedelta
from numpy import datetime64

all_data_mdm = []
pd.set_option('display.max_columns', None)


def ReformDays(Days):
    now_work_days = []
    for day in Days:
        if day < 10:
            now_work_days.append("0" + str(day))
        else:
            now_work_days.append(str(day))
    return now_work_days


def CheckDataStock(first, two):
    global all_data_mdm
    first = first.dropna(subset=['存货编码'])  # 去除nan的列
    two = two.dropna(subset=['存货编码'])  # 去除nan的列
    out_data = pd.merge(first.drop(labels=['主要供货单位名称', '最低供应量', '采购员名称', '固定提前期', '计划默认属性'], axis=1), two,
                        on=['存货编码', '存货名称'])
    out_data = out_data[out_data.isnull().any(axis=1)]
    out_data = out_data.loc[out_data["固定提前期"] == 0]

    all_data_mdm.append(out_data)


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
    work_days7 = []  # 设设置到上月7天

    work_days7.extend(last_work_days)
    work_days7.extend(this_work_days)  # 将上个月最后七天和这个月工作日相合并
    work_days7 = ReformDays(work_days7)  # 改造

    base_data_mrp = []

    base_data_stock = []

    flag = 0
    for x in work_days7:
        if flag < 7:
            try:  # 存货档案-20211001
                base_data = pd.read_excel(f"./KPI/SCM/SP/存货档案-{year}{last_month}{x}.XLSX",
                                          usecols=['存货编码', '存货名称', '主要供货单位名称', '采购员名称', '最低供应量', '固定提前期', '计划默认属性'],
                                          converters={'最低供应量': int, '固定提前期': int}
                                          )
                base_data = base_data.loc[base_data["计划默认属性"] == "采购"]
                base_data_stock.append(base_data)
                flag = flag + 1
                continue
            except:
                continue
        else:
            try:
                now_data = pd.read_excel(f"./KPI/SCM/SP/存货档案-{year}{this_month}{x}.XLSX",
                                         usecols=['存货编码', '存货名称', '主要供货单位名称', '采购员名称', '最低供应量', '固定提前期', '计划默认属性'],
                                         converters={'最低供应量': int, '固定提前期': int}
                                         )
                now_data = now_data.loc[now_data["计划默认属性"] == "采购"]
            except:
                continue
            base_data_stock.append(now_data)  # 新添加新的base
            CheckDataStock(base_data_stock[0], now_data)  # 合并检查是否存在一样的
            del (base_data_stock[0])  # 删除第一个base

    res = pd.concat(all_data_mdm, axis=0, ignore_index=True)
    res = res.drop_duplicates()
    res.to_excel('./KPI/SCM/SP/存货档案维护及时率.xlsx', sheet_name="存货档案维护及时率", index=False)
