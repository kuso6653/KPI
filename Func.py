import os
import calendar
from datetime import timedelta
import datetime


# 创建文件夹路径
def mkdir(path):
    folder = os.path.exists(path)
    if not folder:  # 判断是否存在文件夹如果不存在则创建为文件夹
        os.makedirs(path)  # makedirs 创建文件时如果路径不存在会创建这个路径
    else:
        pass


# "--------------------------------------------------------------------------------------------"
# OA 截取字段匹配判断逻辑
def GeneralOffice(approval):
    if approval.find("提交") != -1 and approval.find("综合管理部盖章") != -1:
        return True
    else:
        return False


def OrderRegistration(approval):
    if approval.find("提交") != -1 and approval.find("U8销售订单号登记") != -1:
        return True
    else:
        return False


def RelevantPersonnel(approval):
    if approval.find("提交") != -1 and approval.find("相关人员办理") != -1:
        return True
    else:
        return False


# "--------------------------------------------------------------------------------------------"
# 读取指定文件
def ReadTxT():
    with open("//10.56.164.228/KPI/员工手册.txt") as f:
        txt = f.read()
        f.close()
    return txt


# "--------------------------------------------------------------------------------------------"
# 添加日期小于10时不全01、 02 等
def ReformDays(Days):
    now_work_days = []
    for day in Days:
        if day < 10:
            now_work_days.append("0" + str(day))
        else:
            now_work_days.append(str(day))
    return now_work_days


# 获取当月工作日函数
def WorkDays(year, month):
    # 利用日历函数，创建截取工作日日期
    cal = calendar.Calendar()
    WorkDay = []  # 创建工作日数组
    for week in cal.monthdayscalendar(int(year), int(month)):
        for i, day in enumerate(week):
            # 为0或者大于等于5的为休息日
            if day == 0 or i >= 5:
                continue
            # 否则加入数组
            WorkDay.append(day)
    return WorkDay


# 获取当月每天日期函数
def EveryDays(year, month):
    # 利用日历函数，创建截取工作日日期
    cal = calendar.Calendar()
    WorkDay = []  # 创建工作日数组
    for week in cal.monthdayscalendar(int(year), int(month)):
        for i, day in enumerate(week):
            WorkDay.append(day)
    return WorkDay


# 获取当月、上月日期
def GetDate():
    now = datetime.date.today()
    # 获取当月首尾日期
    ThisMonthStart = datetime.datetime(now.year, now.month, 1)
    ThisMonthEnd = datetime.datetime(now.year, now.month, calendar.monthrange(now.year, now.month)[1])
    # 获取上月首尾日期
    LastMonthEnd = ThisMonthStart - timedelta(days=1)
    LastMonthStart = datetime.datetime(LastMonthEnd.year, LastMonthEnd.month, 1)
    return ThisMonthStart, ThisMonthEnd, LastMonthEnd, LastMonthStart


def str2sec(x):
    # 字符串时分秒转换
    h, m, s = x.strip().split(':')  # .split()函数将其通过':'分隔开，.strip()函数用来除去空格
    return int(h) + int(m) / 60  # int()函数转换成整数运算


# "--------------------------------------------------------------------------------------------"
def Path():
    return "//10.56.164.228/KPI"
