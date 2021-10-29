import xlrd
import pandas as pd
import calendar
import datetime
from datetime import timedelta
import openpyxl

from numpy import datetime64
from openpyxl import load_workbook

now = datetime.date.today()

# 获取当月首尾日期
this_month_start = datetime.datetime(now.year, now.month, 1)
this_month_end = datetime.datetime(now.year, now.month, calendar.monthrange(now.year, now.month)[1])

# 获取上月首尾日期
last_month_end = this_month_start - timedelta(days=1)
last_month_start = datetime.datetime(last_month_end.year, last_month_end.month, 1)

# 将上月首尾日期切割
last_month_start = str(last_month_start).split(" ")
last_month_end = str(last_month_end).split(" ")

this_month_start = str(this_month_start).split(" ")[0].replace("-", "")
this_month_end = str(this_month_end).split(" ")[0].replace("-", "")
_this_month_end = str(this_month_end).split(" ")[0].replace("", "")
data_time = datetime64("2000-01-02")
print(_this_month_end)