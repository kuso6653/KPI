import calendar
from datetime import timedelta
import datetime
now = datetime.date.today()
# 获取当月首尾日期
ThisMonthStart = datetime.datetime(now.year, now.month, 1)
ThisMonthEnd = datetime.datetime(now.year, now.month, calendar.monthrange(now.year, now.month)[1])
# 获取上月首尾日期
LastMonthEnd = ThisMonthStart - timedelta(days=1)
LastMonthStart = datetime.datetime(LastMonthEnd.year, LastMonthEnd.month, 1)
ThisMonthStart = str(ThisMonthStart).split(" ")[0]
print(ThisMonthStart, ThisMonthEnd, LastMonthEnd, LastMonthStart)